using System.Text.Json;
using mOffice.Nucleo;

namespace mOffice.Comum;

/// <summary>config.json por usuario. Le a cada get e grava a cada set (como AppConfig do mPDF): sem cache, uma
/// mudanca feita com o app aberto vale na hora. Campo ausente = padrao do record (sem migracao).</summary>
public sealed class Configuracoes
{
    // Padrao Claro (o "Clássico claro" do editor), escolha do usuario em 2026-09-25: a aparencia do
    // ONLYOFFICE que o escritorio ja' conhece, independente do tema do Windows. "Seguir o Windows" continua
    // disponivel em Configuracoes.
    // Campos novos SEMPRE no fim, com padrao: um config.json antigo continua lendo (Plano 6, Task 4).
    private sealed record Dados(ModoTema Tema = ModoTema.Claro, bool Grampo = false, string? NomeUsuario = null,
        string? PastaPadraoSalvar = null, bool FormatoOdf = false, string? PastaModelos = null, string? TemaEditorId = null);
    private readonly string _arquivo;

    public Configuracoes(string arquivo) { _arquivo = arquivo; Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!); }
    public static Configuracoes Padrao() => new(Pastas.Config);

    public ModoTema Tema { get => Ler().Tema; set => Alterar(d => d with { Tema = value }); }
    public bool Grampo { get => Ler().Grampo; set => Alterar(d => d with { Grampo = value }); }

    /// <summary>Nome do autor de comentarios e revisoes; <c>null</c> = o do Office/Windows (<see cref="Autor.DoWindows"/>).</summary>
    public string? NomeUsuario { get => Ler().NomeUsuario; set => Alterar(d => d with { NomeUsuario = Vazio(value) }); }
    /// <summary>Pasta onde o "Salvar como" de documento novo abre; <c>null</c> = Documentos.</summary>
    public string? PastaPadraoSalvar { get => Ler().PastaPadraoSalvar; set => Alterar(d => d with { PastaPadraoSalvar = Vazio(value) }); }
    /// <summary>Formato do documento NOVO: <c>false</c> = Microsoft Office (docx/xlsx/pptx), <c>true</c> = OpenDocument.</summary>
    public bool FormatoOdf { get => Ler().FormatoOdf; set => Alterar(d => d with { FormatoOdf = value }); }
    /// <summary>Pasta de modelos pessoais; <c>null</c> = <see cref="PastaModelosEfetiva"/> (a do Office).</summary>
    public string? PastaModelos { get => Ler().PastaModelos; set => Alterar(d => d with { PastaModelos = Vazio(value) }); }
    /// <summary>Tema escolhido DENTRO do editor (<c>theme-white</c>…), lembrado; <c>null</c> = o do mOffice.</summary>
    public string? TemaEditorId { get => Ler().TemaEditorId; set => Alterar(d => d with { TemaEditorId = Vazio(value) }); }

    /// <summary>Ultima posicao da janela de documento do app (<c>mtexto</c>/<c>mplanilha</c>/<c>mslides</c>); <c>null</c> = nunca salva ou ilegivel (a janela nasce maximizada).
    /// Vive num arquivo POR APP (<c>posicao-&lt;app&gt;.json</c>, ao lado do config.json), gravado em todo fechamento de janela: separado do config.json
    /// compartilhado para que um logoff (os tres apps fecham juntos) nunca dispute com as configuracoes do usuario. O campo <c>Posicoes</c> de config.json
    /// antigos e' ignorado.</summary>
    public PosicaoSalva? Posicao(string app)
    {
        var arq = ArquivoDaPosicao(app);
        try { return File.Exists(arq) ? JsonSerializer.Deserialize<PosicaoSalva>(ArquivoAtomico.Ler(arq)) : null; }
        catch (JsonException) { return null; }
        catch (IOException) { return null; }
    }

    public void DefinirPosicao(string app, PosicaoSalva posicao)
    {
        var arq = ArquivoDaPosicao(app);
        // Gravada em todo fechamento de janela: uma falha (somente leitura, disco cheio) so' perde a posicao, nunca o fechamento (I2).
        try { ArquivoAtomico.ComExclusao(arq, () => { ArquivoAtomico.Escrever(arq, JsonSerializer.Serialize(posicao)); return 0; }); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { Logar("CONFIG posicao nao gravada: " + e.Message); }
    }

    private string ArquivoDaPosicao(string app) =>
        Path.Combine(Path.GetDirectoryName(Path.GetFullPath(_arquivo))!, $"posicao-{app.ToLowerInvariant()}.json");

    public string PastaModelosEfetiva => PastaModelos ?? PastaModelosDoOffice;

    /// <summary>A pasta de modelos pessoais do Office (<c>%APPDATA%MicrosoftTemplates</c>).</summary>
    public static string PastaModelosDoOffice =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates");

    private static string? Vazio(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    /// <summary>Valor guardado so' na memoria deste processo quando o config.json nao pode ser LIDO por E/S (preso por outro processo) ou nao
    /// pode ser GRAVADO (somente leitura, sem permissao — I2 da revisao final do Plano 6): o valor vale ate' uma gravacao dar certo.</summary>
    private Dados? _memoria;

    /// <summary>A ultima gravacao falhou: <see cref="_memoria"/> prevalece sobre o arquivo (que ainda tem o valor antigo) ate' uma gravacao dar
    /// certo. Sem isto o valor escolhido pelo usuario "voltaria" na leitura seguinte.</summary>
    private bool _memoriaPrevalece;

    /// <summary>Linha de log opcional (a casca liga ao log do processo).</summary>
    public static Action<string>? Log { get; set; }

    private static void Logar(string linha) { try { Log?.Invoke(linha); } catch (Exception) { } }

    private enum Leitura { Ok, ConteudoIlegivel, Inacessivel }

    /// <summary>(dados, leitura): arquivo ausente = padroes e Ok; JSON invalido/vazio = padroes e <see cref="Leitura.ConteudoIlegivel"/> (corrupcao
    /// persistente: as escritas sao atomicas); E/S ou permissao = padroes e <see cref="Leitura.Inacessivel"/> (pode ser passageiro).</summary>
    private (Dados Dados, Leitura Leitura) Tentar()
    {
        try
        {
            if (!File.Exists(_arquivo)) return (new Dados(), Leitura.Ok);
            var d = JsonSerializer.Deserialize<Dados>(ArquivoAtomico.Ler(_arquivo));
            return d is null ? (new Dados(), Leitura.ConteudoIlegivel) : (d, Leitura.Ok);
        }
        catch (JsonException) { return (new Dados(), Leitura.ConteudoIlegivel); }
        catch (IOException) { return (new Dados(), Leitura.Inacessivel); }
        catch (UnauthorizedAccessException) { return (new Dados(), Leitura.Inacessivel); }
    }

    private Dados Ler()
    {
        var (d, leitura) = Tentar();
        if (_memoria is not null && (_memoriaPrevalece || leitura != Leitura.Ok)) return _memoria;
        if (leitura == Leitura.Ok) _memoria = null;
        return d;
    }

    /// <summary>
    /// Leitura-modificacao-escrita sob o mutex entre processos, com escrita atomica. Nunca lanca (I2 da revisao final do Plano 6: um setter
    /// chamado de um evento da interface derrubava o app quando o arquivo estava somente leitura).
    /// <list type="bullet">
    /// <item>Conteudo ilegivel (M2): o arquivo vai para <c>config.json.ilegivel-&lt;yyyyMMddHHmmss&gt;</c> (bytes preservados), e a gravacao recomeca
    /// dos padroes — antes, um config.json corrompido nunca mais era gravado.</item>
    /// <item>Inacessivel (preso por outro processo): nao grava; o valor fica na memoria.</item>
    /// <item>Gravacao que falha: log, e o valor fica na memoria e prevalece ate' uma gravacao dar certo.</item>
    /// </list>
    /// </summary>
    private void Alterar(Func<Dados, Dados> mudanca)
    {
        try
        {
            ArquivoAtomico.ComExclusao(_arquivo, () =>
            {
                var (d, leitura) = Tentar();
                if (leitura == Leitura.Inacessivel)
                {
                    _memoria = mudanca(_memoria ?? d);
                    Logar("CONFIG config.json inacessivel: a mudanca vale so' neste processo e o arquivo nao foi alterado");
                    return 0;
                }
                if (leitura == Leitura.ConteudoIlegivel)
                {
                    var guardado = PorDeLado(_arquivo);
                    Logar("CONFIG config.json ilegivel guardado como " + Path.GetFileName(guardado) + "; recomecando dos padroes");
                    d = new Dados();
                }
                var novo = mudanca(_memoriaPrevalece && _memoria is not null ? _memoria : leitura == Leitura.ConteudoIlegivel ? _memoria ?? d : d);
                try
                {
                    ArquivoAtomico.Escrever(_arquivo, JsonSerializer.Serialize(novo));
                    _memoria = null; _memoriaPrevalece = false;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException)
                {
                    _memoria = novo; _memoriaPrevalece = true;
                    Logar("CONFIG config.json nao gravado (" + e.Message + "): a mudanca vale so' neste processo");
                }
                return 0;
            });
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            // O proprio PorDeLado falhou (arquivo preso/sem permissao para renomear): o valor fica na memoria, o arquivo como estava.
            _memoria = mudanca(_memoria ?? new Dados()); _memoriaPrevalece = true;
            Logar("CONFIG mudanca nao gravada (" + e.Message + "): vale so' neste processo");
        }
    }

    /// <summary>Renomeia <paramref name="arquivo"/> para <c>&lt;nome&gt;.ilegivel-&lt;yyyyMMddHHmmss&gt;</c> (com <c>-N</c> se ja' existir) e devolve o
    /// novo caminho. Lanca se nao conseguir (o chamador mantem tudo como estava).</summary>
    internal static string PorDeLado(string arquivo)
    {
        var baseNome = arquivo + ".ilegivel-" + DateTime.Now.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture);
        var destino = baseNome;
        for (var n = 2; File.Exists(destino); n++) destino = baseNome + "-" + n;
        File.Move(arquivo, destino);
        return destino;
    }
}
