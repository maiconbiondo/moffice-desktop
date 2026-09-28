using System.Text.Json;

namespace mOffice.Nucleo.Recuperacao;

/// <summary>
/// O que e' preciso para reabrir uma pasta de recuperacao depois de uma queda (Task 2 do Plano 3), gravado em
/// <c>&lt;pasta&gt;\sessao.json</c> ao abrir o documento: o arquivo de origem, o formato, o titulo, se era um
/// documento NOVO (sem arquivo do usuario), o app dono (mtexto/mplanilha/mslides) e o <see cref="IdUsuario"/>
/// desta abertura.
///
/// <see cref="IdUsuario"/> e' o que <c>AscDesktopEditor.CheckUserId()</c> responde ao editor (o
/// <c>DesktopOfflineAppDocumentEndLoad</c> do sdkjs grava o valor em <c>AscCommon.Hj.FJ</c>, o prefixo dos ids
/// de elementos criados na sessao). Tem de AUMENTAR a cada reabertura da mesma pasta: os changes recuperados
/// foram gerados com o id da sessao que caiu, e a sessao nova reusando o mesmo id criaria elementos com ids que
/// ja' existem no documento.
///
/// <see cref="Protegido"/> (Task 7 do Plano 3, parametro opcional NO FIM: um <c>sessao.json</c> antigo, sem o
/// campo, continua legivel e vale <c>false</c>): o documento desta sessao tinha senha. A SENHA em si nunca vem
/// para ca' (so' existe em memoria, <c>SessaoDocumento.Senha</c>); por isso uma sessao recuperada de um documento
/// protegido reabre SEM senha, e o host avisa o usuario para defini-la de novo antes de salvar.
///
/// <see cref="OrigemEscritaUtc"/>/<see cref="OrigemTamanho"/> (revisao final do Plano 3, I2; opcionais NO FIM,
/// como <see cref="Protegido"/>: um <c>sessao.json</c> antigo continua legivel e vale <c>null</c>): o carimbo do
/// arquivo de origem (<c>LastWriteTimeUtc</c> + tamanho) na abertura e depois de cada salvamento bem-sucedido. Na
/// recuperacao, um original com carimbo DIFERENTE (ou sumido, ou sem carimbo gravado) foi mexido depois da queda
/// — o usuario reabriu e salvou o original, por exemplo — e a sessao recuperada nao pode grava-lo por cima.
/// </summary>
public sealed record MetadadosSessao(string Origem, int Formato, string Titulo, bool Novo, string App, int IdUsuario, bool Protegido = false,
    DateTime? OrigemEscritaUtc = null, long? OrigemTamanho = null)
{
    public const string NomeArquivo = "sessao.json";

    /// <summary>Escrita atomica e no disco antes da troca (<see cref="GravacaoAtomica"/>; o flush e' o M7 da revisao
    /// final do Plano 3). Uma queda NO MEIO da gravacao deixa o arquivo antigo inteiro (ou o <c>.tmp</c> orfao),
    /// nunca um <c>sessao.json</c> pela metade — que e' justamente o arquivo lido depois de uma queda.</summary>
    public void Gravar(string pasta)
    {
        Directory.CreateDirectory(pasta);
        GravacaoAtomica.Gravar(Path.Combine(pasta, NomeArquivo), System.Text.Encoding.UTF8.GetBytes(JsonSerializer.Serialize(this)));
    }

    /// <summary>Carimbo de <paramref name="arquivo"/> (<c>LastWriteTimeUtc</c>, tamanho), ou <c>(null, null)</c>
    /// quando ele nao existe ou nao da' para ler os atributos.</summary>
    public static (DateTime? EscritaUtc, long? Tamanho) Carimbo(string arquivo)
    {
        try
        {
            var f = new FileInfo(arquivo);
            return f.Exists ? (f.LastWriteTimeUtc, f.Length) : (null, null);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return (null, null); }
    }

    /// <summary>Este metadado com o carimbo ATUAL de <paramref name="arquivo"/>.</summary>
    public MetadadosSessao ComCarimboDe(string arquivo)
    {
        var (escrita, tamanho) = Carimbo(arquivo);
        return this with { OrigemEscritaUtc = escrita, OrigemTamanho = tamanho };
    }

    /// <summary>O arquivo de origem MUDOU desde o ultimo carimbo gravado (I2 da revisao final do Plano 3)? Falha
    /// FECHADA: sem carimbo gravado (<c>sessao.json</c> de antes desta versao, ou carimbo que nao pode ser lido na
    /// abertura), original sumido ou ilegivel contam como "mudou". Documento NOVO nao tem original: nunca muda.</summary>
    public bool OrigemMudou()
    {
        if (Novo) return false;
        if (OrigemEscritaUtc is not { } escrita || OrigemTamanho is not { } tamanho) return true;
        var (agora, tamanhoAgora) = Carimbo(Origem);
        return agora is null || agora.Value.ToUniversalTime().Ticks != escrita.ToUniversalTime().Ticks || tamanhoAgora != tamanho;
    }

    /// <summary>Metadados da pasta, ou <c>null</c> quando o arquivo nao existe, nao e' legivel ou esta'
    /// corrompido (uma pasta assim nao e' recuperavel — quem chama decide o que fazer com ela).</summary>
    public static MetadadosSessao? Ler(string pasta)
    {
        try
        {
            var m = JsonSerializer.Deserialize<MetadadosSessao>(GravacaoAtomica.LerCompartilhado(Path.Combine(pasta, NomeArquivo)));
            // JSON valido mas sem os campos (ex.: "{}" ou "null") tambem nao serve para reabrir nada.
            return m is null || string.IsNullOrEmpty(m.Origem) || m.App is null || m.Titulo is null ? null : m;
        }
        catch (Exception e) when (e is JsonException or IOException or UnauthorizedAccessException) { return null; }
    }
}
