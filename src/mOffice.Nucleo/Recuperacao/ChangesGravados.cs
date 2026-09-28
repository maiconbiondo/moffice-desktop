using System.Text;
using System.Text.Json;

namespace mOffice.Nucleo.Recuperacao;

/// <summary>Resultado de <see cref="ChangesGravados.Ler"/>. <see cref="ArrayJs"/> e <see cref="Quantidade"/> sao
/// o MAIOR PREFIXO VALIDO do arquivo; <see cref="Descartou"/> diz se havia bytes depois dele (cauda rasgada por
/// uma queda no meio da gravacao, zeros, lixo) e <see cref="BytesDescartados"/> quantos.</summary>
public readonly record struct LeituraChanges(string ArrayJs, int Quantidade, bool Descartou, long BytesDescartados);

/// <summary>O que <see cref="ChangesGravados.Reparar"/> fez: nada (<see cref="Reparou"/> falso) ou regravou o
/// arquivo com <see cref="Quantidade"/> changes, descartando <see cref="BytesDescartados"/> bytes, com o
/// original guardado em <see cref="Copia"/>.</summary>
public readonly record struct ReparoChanges(bool Reparou, int Quantidade, long BytesDescartados, string? Copia);

/// <summary>
/// Leitura do <c>changes/changes0.json</c> para devolver ao editor numa reabertura (Task 2 do Plano 3). O
/// arquivo e' gravado no formato do app oficial (<see cref="PonteDesktop.SalvarChanges"/>): cada change entre
/// aspas seguido de virgula — <c>"c1","c2",</c> — ou seja, um array JS sem os colchetes e com a virgula final.
/// Para o editor (<c>window.DesktopOfflineAppDocumentApplyChanges([...])</c>) basta tirar a ultima virgula e
/// envolver em colchetes.
///
/// Arquivo danificado (fix round 1 da revisao da Task 2 — decisao do controlador, substitui o "JSON invalido =
/// vazio" do brief): uma queda no MEIO de uma gravacao deixa a cauda rasgada (string pela metade, bloco de
/// zeros do sistema de arquivos). Jogar o arquivo inteiro fora por causa da cauda descartaria em silencio
/// justamente o trabalho que a recuperacao existe para salvar. Por isso a leitura devolve o MAIOR PREFIXO de
/// changes completos (cada um uma string JSON seguida de virgula) e marca o que sobrou como descartado.
/// </summary>
public static class ChangesGravados
{
    private static readonly byte[] Bom = [0xEF, 0xBB, 0xBF];

    /// <summary>PURA (so' le'; a Task 3 chama isto em pastas que nao sao dela): o maior prefixo valido do
    /// arquivo como array JS e quantos changes ele tem. Arquivo ausente, vazio ou ilegivel = <c>("[]", 0)</c>
    /// sem descarte. A contagem e a validacao vem de um leitor JSON de verdade, nao de contar aspas: um change
    /// pode ter aspas escapadas.
    ///
    /// A validacao tambem e' a defesa do script: o texto do arquivo vem da PAGINA (o <c>param</c> de
    /// <c>LocalFileSaveChanges</c> e' gravado como veio) e vai embutido num <c>ExecuteScriptAsync</c>. So'
    /// strings JSON completas separadas por virgula entram no array — dado puro em JS, sem como "fechar o
    /// colchete" e executar codigo.</summary>
    public static LeituraChanges Ler(string arquivoChanges)
    {
        byte[] bytes;
        try
        {
            if (!File.Exists(arquivoChanges)) return new("[]", 0, false, 0);
            // Compartilhando escrita/delecao (fix round 2 da revisao final do Plano 3): a boas-vindas e outras
            // instancias leem pastas alheias, e uma leitura com FileShare.Read fazia o FileMode.Append da sessao viva
            // falhar no meio da edicao.
            bytes = GravacaoAtomica.LerCompartilhado(arquivoChanges);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return new("[]", 0, false, 0); }
        return Montar(bytes, Analisar(bytes));
    }

    /// <summary>
    /// Regrava <paramref name="arquivoChanges"/> com EXATAMENTE o prefixo valido quando <see cref="Ler"/>
    /// descartaria bytes — so' <c>JanelaEditor.RecuperarAsync</c> chama, depois de adquirir a trava da pasta
    /// (e' o unico dono dela nesse instante). Antes de regravar, copia o original para
    /// <c>changes0.json.corrompido</c> (ou <c>.corrompido.1</c>, <c>.2</c>... — uma copia existente nunca e'
    /// sobrescrita). A regravacao e' atomica (<c>.tmp</c> + <c>File.Move(..., true)</c>).
    ///
    /// Por que regravar em vez de so' aplicar o prefixo: a sessao recuperada vai APENDAR changes novos ao
    /// arquivo; apendados depois da cauda rasgada, eles ficariam atras de lixo e a PROXIMA recuperacao (ou o
    /// x2t ao salvar) perderia tudo o que veio depois do rasgo.
    ///
    /// Retencao: a copia vive DENTRO da pasta de recuperacao, entao um fechamento limpo apaga-a junto com a
    /// pasta. E' intencional — o usuario seguiu trabalhando sobre a versao salva do rasgo e, ao fechar, decidiu
    /// o destino do documento; a copia so' serve enquanto a pasta existe (diagnostico, ou uma nova queda).
    /// </summary>
    public static ReparoChanges Reparar(string arquivoChanges) => RepararELer(arquivoChanges).Reparo;

    /// <summary>
    /// <see cref="Reparar"/> + a leitura do que ficou no arquivo, com UMA leitura do disco so' (revisao final do
    /// Plano 3, M1). Antes, <c>JanelaEditor.RecuperarAsync</c> reparava e depois chamava <see cref="Ler"/> — que e'
    /// PURA e por isso engole erro de E/S devolvendo <c>("[]", 0)</c>: um erro transitorio entre as duas leituras
    /// (antivirus, backup segurando o arquivo) abria a sessao recuperada com ZERO changes, e o prefixo recuperado
    /// ficava truncavel pelo primeiro <c>deleteIndex</c> da sessao nova. Aqui um arquivo que nao da' para ler
    /// LANCA (<see cref="IOException"/>/<see cref="UnauthorizedAccessException"/>) — quem chama transforma isso na
    /// recusa pt-BR de sempre. Arquivo ausente continua sendo "sem changes".
    /// </summary>
    public static (ReparoChanges Reparo, LeituraChanges Leitura) RepararELer(string arquivoChanges)
    {
        if (!File.Exists(arquivoChanges)) return (new(false, 0, 0, null), new("[]", 0, false, 0));
        var bytes = GravacaoAtomica.LerCompartilhado(arquivoChanges);
        var a = Analisar(bytes);
        var leitura = Montar(bytes, a);
        if (a.Descartados == 0) return (new(false, a.Quantidade, 0, null), leitura);
        var copia = arquivoChanges + ".corrompido";
        for (var i = 1; File.Exists(copia); i++) copia = arquivoChanges + ".corrompido." + i;
        File.Copy(arquivoChanges, copia, overwrite: false);
        // Atomico e no disco ANTES da troca (fix round 2; GravacaoAtomica): sem o flush uma queda logo depois do
        // Move podia deixar o nome novo apontando para dados ainda no cache — o mesmo rasgo que isto conserta.
        GravacaoAtomica.Gravar(arquivoChanges, bytes.AsSpan(a.Inicio, a.Fim - a.Inicio));
        // O arquivo agora e' exatamente o prefixo: a leitura devolvida descreve o que esta' no disco (sem descarte).
        return (new(true, a.Quantidade, a.Descartados, copia), leitura with { Descartou = false, BytesDescartados = 0 });
    }

    private static LeituraChanges Montar(byte[] bytes, Analise a)
    {
        var corpo = a.Fim > a.Inicio + 1 ? Encoding.UTF8.GetString(bytes, a.Inicio, a.Fim - 1 - a.Inicio) : "";
        return new("[" + corpo + "]", a.Quantidade, a.Descartados > 0, a.Descartados);
    }

    /// <summary>Onde comecam os changes (depois de um BOM, se houver), onde termina o prefixo valido (depois da
    /// virgula do ultimo change completo), quantos changes ha' nele e quantos bytes nao-brancos sobraram.</summary>
    private readonly record struct Analise(int Inicio, int Fim, int Quantidade, long Descartados);

    private static Analise Analisar(byte[] bytes)
    {
        var inicio = bytes.AsSpan().StartsWith(Bom) ? Bom.Length : 0;
        // O leitor ve' "[" + arquivo: cada string no nivel 1 e' um change. Ele lanca no primeiro byte invalido
        // (zero, string pela metade, falta de virgula) — o que foi lido ate' ali e' o prefixo valido.
        var dados = new byte[bytes.Length - inicio + 1];
        dados[0] = (byte)'[';
        Array.Copy(bytes, inicio, dados, 1, bytes.Length - inicio);
        var fim = inicio;
        var n = 0;
        var leitor = new Utf8JsonReader(dados, isFinalBlock: true, state: default);
        try
        {
            if (leitor.Read() && leitor.TokenType == JsonTokenType.StartArray)
            {
                while (leitor.Read() && leitor.TokenType == JsonTokenType.String && leitor.CurrentDepth == 1)
                {
                    // BytesConsumed conta o "[" artificial: menos 1 = posicao no arquivo logo depois da aspa final.
                    var depoisDaAspa = inicio + (int)leitor.BytesConsumed - 1;
                    // Um change so' esta' completo com a virgula que o segue (formato "c",): sem ela, a gravacao
                    // foi interrompida entre a aspa e a virgula e o proximo change nunca comecou.
                    if (depoisDaAspa >= bytes.Length || bytes[depoisDaAspa] != (byte)',') break;
                    fim = depoisDaAspa + 1;
                    n++;
                }
            }
        }
        catch (JsonException) { }
        long descartados = 0;
        for (var i = fim; i < bytes.Length; i++)
            if (bytes[i] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')) { descartados = bytes.Length - fim; break; }
        return new(inicio, fim, n, descartados);
    }
}
