using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>
/// Lado C# da ponte desktop, exposto ao JS como <c>chrome.webview.hostObjects.sync.ponte</c> (no documento de
/// cima E em cada quadro — o editor vive num iframe e o <c>CoreWebView2.AddHostObjectToScript</c> nao alcanca
/// quadros; ver <see cref="JanelaEditor"/>).
///
/// Canal unico JS -> C#: <see cref="Avisar"/>. O WebView2 entrega mensagens de <c>chrome.webview.postMessage</c>
/// de um iframe no evento do QUADRO, nao no da janela; e o JS ja precisa do objeto de host de qualquer jeito,
/// entao um metodo so evita manter dois caminhos (divergencia registrada em relacao ao brief, que usava
/// postMessage).
///
/// Gravacao de changes no formato do app oficial (client_renderer_wrapper.cpp ~2119: aspas + change + aspas +
/// virgula, appendado). Nao exercitado ainda — a sonda so cobriu a abertura; a Task 6 fecha isso.
/// </summary>
[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]
public sealed class PonteDesktop
{
    private readonly SessaoDocumento _s;
    private readonly Action<string, string>? _aviso;
    private readonly string? _pastaFontes;
    private readonly Action<string>? _log;
    private readonly Action<bool>? _aoMudarModificado;
    private readonly Func<string, string?>? _ortografia;
    private readonly Action<string>? _aoFalharGravacao;
    private readonly object _trava = new();
    // Arquivos que o USUARIO escolheu num dialogo do host (ver JanelaEditor.ResponderDialogoArquivoAsync) e
    // que, por isso, a pagina pode mandar copiar para media/. Mesma defesa do app oficial, onde
    // `g_pLocalResolver->AddFile(sPath)` registra o caminho escolhido e `GetLocalImageUrl` chama
    // `g_pLocalResolver->Check(sUrlFile)` antes de ler qualquer arquivo do disco.
    // Consentimento de UM USO SO' (minor 2 da revisao 1, ver Resolver): a entrada e' REMOVIDA assim que o
    // conteudo e' copiado para media/ com sucesso — um segundo pedido do MESMO caminho sem passar de novo
    // pelo dialogo do host e' recusado, exatamente como escolher o arquivo autoriza aquela colagem/insercao,
    // nao um acesso permanente da pagina ao arquivo.
    private readonly HashSet<string> _escolhidosPeloUsuario = new(StringComparer.OrdinalIgnoreCase);
    // Conteudo ja gravado -> nome em media/ (o app oficial guarda o mesmo mapa por CRC32): colar duas vezes a
    // MESMA imagem nao duplica o arquivo dentro do documento.
    private readonly Dictionary<string, string> _midiaPorConteudo = [];

    /// <summary>Teto do que a ponte aceita gravar em media/ de uma vez. Nao existe no app oficial; aqui existe
    /// porque o dado pode vir de uma `data:` URL montada pela PAGINA (conteudo web), e uma imagem de documento
    /// de escritorio nao chega perto disso.</summary>
    private const int MaximoBytesImagem = 64 * 1024 * 1024;

    /// <summary><paramref name="aoMudarModificado"/> (Task 2 do Marco 1 / Plano 2) e' um parametro OPCIONAL no
    /// fim de proposito: os testes que constroem a ponte sem janela continuam compilando sem mudanca. Idem
    /// <paramref name="ortografia"/> (Task 1 do Plano 3): pedido do SpellCheck -&gt; JSON da resposta, ou
    /// <c>null</c> para nao responder (ver <see cref="Ortografia.PedidoOrtografia"/>).</summary>
    public PonteDesktop(SessaoDocumento s, Action<string, string>? aviso = null, string? pastaFontes = null,
        Action<string>? log = null, Action<bool>? aoMudarModificado = null, Func<string, string?>? ortografia = null,
        Action<string>? aoFalharGravacao = null)
    {
        _s = s; _aviso = aviso; _pastaFontes = pastaFontes; _log = log; _aoMudarModificado = aoMudarModificado;
        _ortografia = ortografia; _aoFalharGravacao = aoFalharGravacao;
    }

    private bool _congelada;

    /// <summary>Daqui para a frente nenhum <see cref="SalvarChanges"/> grava nada (fix round 2 da revisao da Task
    /// 2): a reaplicacao de uma sessao recuperada falhou e o <c>changes0.json</c> em disco e' a fonte da verdade —
    /// tem de ficar byte a byte como estava para uma nova tentativa. `internal`: so' o host congela (todo metodo
    /// PUBLICO desta classe e' chamavel pelo JS).</summary>
    internal void Congelar()
    {
        bool mudou;
        lock (_trava) { mudou = !_congelada; _congelada = true; }
        if (mudou) AvisarCongelou();
    }

    /// <summary>A ponte acabou de congelar (pelo host ou por uma gravacao de changes que falhou; fix round 3 da revisao
    /// final do Plano 3). Quem espera um <see cref="ChangesSalvos"/> (<c>JanelaEditor.DescarregarChangesAsync</c>)
    /// desiste NA HORA em vez de esperar os 10 s. Fora da trava. `internal`: membro publico viraria superficie COM.</summary>
    internal event Action? Congelou;

    private void AvisarCongelou()
    {
        try { Congelou?.Invoke(); }
        catch (Exception ex) { _log?.Invoke("PONTE erro no aviso de congelamento: " + ex.Message); }
    }

    private long _geracaoDeGravacao;

    /// <summary>Quantas vezes <see cref="SalvarChanges"/> gravou com sucesso (fix round 4 da revisao final do Plano 3). O
    /// salvamento guarda o valor logo antes do x2t ler a pasta: se mudou depois, o arquivo gravado nao tem tudo o que
    /// esta' no <c>changes0.json</c> — o tamanho do arquivo nao basta (desfazer um change e digitar outro do mesmo
    /// tamanho o mantem). `internal`: membro publico viraria superficie COM.</summary>
    internal long GeracaoDeGravacao { get { lock (_trava) return _geracaoDeGravacao; } }

    /// <summary>Costura de teste (fix round 3): <see cref="SalvarChanges"/> LANCA para o JS (como uma falha do proprio
    /// objeto de host), antes de qualquer gravacao. `internal`.</summary>
    internal bool LancarEmSalvarChangesParaTeste { get; set; }

    /// <summary>A ponte foi congelada (<see cref="Congelar"/>): nenhum lote de changes chega mais ao disco. Quem
    /// espera uma descarga (<c>JanelaEditor.DescarregarChangesAsync</c>) falha na hora em vez de esperar um
    /// <see cref="ChangesSalvos"/> que nunca vira'.</summary>
    internal bool Congelada { get { lock (_trava) return _congelada; } }

    /// <summary>
    /// Um <c>LocalFileSaveChanges</c> do editor acabou de ser tratado (Task 4 do Plano 3): o lote ja' esta' no
    /// <c>changes0.json</c> (ou nao havia o que gravar, <c>count == 0</c>). Dispara DEPOIS de gravar e fora da
    /// trava, na thread que chamou (a da interface, dentro da chamada sincrona do JS). Nao dispara com a ponte
    /// congelada — nada foi gravado. E' o sinal que <c>JanelaEditor.DescarregarChangesAsync</c> espera.
    ///
    /// `internal` (o brief pedia `public`): esta classe e' objeto de host COM-visivel com
    /// <c>ClassInterfaceType.AutoDual</c>, e todo membro PUBLICO vira superficie chamavel pelo JS do editor —
    /// um evento publico viraria <c>add_ChangesSalvos</c>/<c>remove_ChangesSalvos</c> expostos a' pagina. Mesmo
    /// criterio de <see cref="Congelar"/> e <see cref="PermitirArquivoEscolhido"/>.
    /// </summary>
    internal event Action? ChangesSalvos;

    /// <summary>Autoriza a pagina a mandar copiar ESTE arquivo para media/ (so' o host chama: `internal` de
    /// proposito — esta classe e' um objeto de host COM-visivel, entao todo metodo PUBLICO daqui e' chamavel
    /// pelo JS do editor, e um "autorize este caminho" publico anularia a defesa).</summary>
    internal void PermitirArquivoEscolhido(string caminho)
    {
        if (string.IsNullOrWhiteSpace(caminho)) return;
        lock (_trava) _escolhidosPeloUsuario.Add(Path.GetFullPath(caminho));
    }

    // -----------------------------------------------------------------------------------------------------
    // Conversao e leitura de arquivo pedidas pelo editor (Plano 6, Task 8). Mesma defesa do PermitirArquivoEscolhido:
    // o caminho so' e' convertido/lido se o USUARIO o escolheu num dialogo do host com o filtro que da' esse uso
    // (FiltrosDoEditor: "word"/"cell"/"(*.xml)" -> conversao, "csv/txt" -> leitura), e o consentimento vale UM uso.
    // -----------------------------------------------------------------------------------------------------
    private readonly HashSet<string> _conversaoPermitida = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _leituraPermitida = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Usos de <see cref="PedirConversao"/>: "converter" (<c>convertFile</c>), "comparar" e "combinar" (<c>CompareDocumentFile</c>/<c>MergeDocumentFile</c>).</summary>
    internal static readonly IReadOnlySet<string> UsosDeConversao = new HashSet<string>(StringComparer.Ordinal) { "converter", "comparar", "combinar" };

    /// <summary>Canvas de planilha (AVS_OFFICESTUDIO_FILE_CANVAS_SPREADSHEET): o que o sdkjs pede para atualizar um vinculo externo.</summary>
    internal const int CanvasPlanilha = 8194;

    /// <summary>Autoriza UMA conversao deste caminho (so' o host chama — `internal` pelo mesmo motivo de <see cref="PermitirArquivoEscolhido"/>;
    /// o brief pedia `public`, o que daria a' pagina um "autorize este caminho").</summary>
    internal void PermitirConversao(string caminho)
    {
        if (string.IsNullOrWhiteSpace(caminho)) return;
        lock (_trava) _conversaoPermitida.Add(Path.GetFullPath(caminho));
    }

    /// <summary>Autoriza UMA leitura deste caminho (ver <see cref="PermitirConversao"/>).</summary>
    internal void PermitirLeitura(string caminho)
    {
        if (string.IsNullOrWhiteSpace(caminho)) return;
        lock (_trava) _leituraPermitida.Add(Path.GetFullPath(caminho));
    }

    /// <summary>
    /// <c>convertFile</c>/<c>CompareDocumentFile</c>/<c>MergeDocumentFile</c> do editor: devolve o id do pedido (o host converte FORA desta chamada
    /// sincrona e entrega por <c>window.__moffice.entregarConversao</c> ou <c>onDocumentCompare</c>/<c>onDocumentMerge</c>), ou <c>""</c> se recusado.
    /// Aceita o caminho consentido por <see cref="PermitirConversao"/> (dialogo do host; consumido aqui) e o VINCULO EXTERNO de planilha
    /// (<paramref name="uso"/> "converter", <paramref name="formato"/> 8194): este e' um caminho que o DOCUMENTO cita (fix 1 da revisao da Task 8,
    /// seguranca) e por isso so' vira pedido com <c>confirmar</c> — o host pergunta ao usuario (caminho completo, padrao Nao) ANTES de qualquer
    /// E/S no caminho, e a decisao fica guardada por caminho na sessao (<see cref="RegistrarDecisaoDeVinculo"/>). O "Atualizar" do proprio editor
    /// NAO conta como consentimento. Caminho de dispositivo/namespace ou nao totalmente qualificado: recusado antes de qualquer GetFullPath/
    /// File.Exists, com faixa (<see cref="CaminhoBruto"/>).
    /// </summary>
    public string PedirConversao(string caminho, int formato, string uso)
    {
        uso ??= "";
        if (!UsosDeConversao.Contains(uso)) { _log?.Invoke("CONVERSAO recusada (uso desconhecido: " + Cortar(uso) + ")"); return ""; }
        var vinculo = uso == "converter" && formato == CanvasPlanilha;
        // Fix 1b: vinculo relativo (o Excel grava "custos.xlsx", "..\pasta\custos.xlsx") resolvido pela pasta do original salvo — so' texto.
        if (vinculo) caminho = ResolverRelativo(caminho, PastaDoDocumento) ?? "";
        var bruto = CaminhoBruto(caminho);
        if (bruto is null)
        {
            _log?.Invoke("CONVERSAO recusada (caminho de dispositivo, relativo ou ilegivel; nenhuma E/S): " + Cortar(caminho ?? ""));
            if (vinculo) _aviso?.Invoke("vinculoRecusado", caminho ?? "");
            return "";
        }
        var completo = Consumir(_conversaoPermitida, bruto);
        var confirmar = false;
        if (completo is null && vinculo)
        {
            completo = VinculoPermitido(caminho, out confirmar);
            if (completo is null) return "";
        }
        if (completo is null) { _log?.Invoke("CONVERSAO recusada (nao veio do dialogo do host): " + Cortar(caminho ?? "")); return ""; }
        var id = Guid.NewGuid().ToString("N");
        EmitirPedido("converterArquivo", id, JsonSerializer.Serialize(new { id, caminho = completo, formato, uso, confirmar }));
        return id;
    }

    /// <summary><c>loadLocalFile</c> (Dados &gt; De texto/CSV): id do pedido (os bytes chegam por <c>window.__moffice.entregarLeitura</c>) ou <c>""</c>
    /// se o caminho nao foi consentido por <see cref="PermitirLeitura"/>. Consome o consentimento.</summary>
    public string PedirLeitura(string caminho)
    {
        var bruto = CaminhoBruto(caminho);
        var completo = bruto is null ? null : Consumir(_leituraPermitida, bruto);
        if (completo is null) { _log?.Invoke("LEITURA recusada (nao veio do dialogo do host): " + Cortar(caminho ?? "")); return ""; }
        var id = Guid.NewGuid().ToString("N");
        EmitirPedido("lerArquivo", id, JsonSerializer.Serialize(new { id, caminho = completo }));
        return id;
    }

    /// <summary><c>close()</c> do objeto de conversao: o host apaga a pasta da conversao.</summary>
    public void LiberarConversao(string id) => _aviso?.Invoke("liberarConversao", id ?? "");

    /// <summary>
    /// <c>openExternalReference(ref, cb)</c> (vinculo externo -&gt; "Abrir origem"). MEDIDO em cell/sdk-all-min.js 9.4.0.129: <c>ref</c> e' o caminho do
    /// vinculo sem o "file:///" (<c>AscCommon.o$j</c>) e o callback trata 0 como "nao abriu" (o aviso do proprio sdkjs) e qualquer outro valor como
    /// sucesso. Devolve o id do pedido (o host pergunta se preciso, confere a existencia DEPOIS do sim, pede a abertura em janela nova a' casca e
    /// responde por <c>window.__moffice.entregarReferencia(id, 0|1)</c>) ou <c>""</c> (recusado: o JS responde 0). Mesmas regras do vinculo de
    /// <see cref="PedirConversao"/>.
    /// </summary>
    public string AbrirReferenciaExterna(string referencia)
    {
        var original = referencia;
        referencia = ResolverRelativo(referencia, PastaDoDocumento) ?? "";
        if (CaminhoBruto(referencia) is null)
        {
            _log?.Invoke("VINCULO recusado (caminho de dispositivo, relativo ou ilegivel; nenhuma E/S): " + Cortar(original ?? ""));
            _aviso?.Invoke("vinculoRecusado", original ?? "");
            return "";
        }
        var completo = VinculoPermitido(referencia, out var confirmar);
        if (completo is null) return "";
        var id = Guid.NewGuid().ToString("N");
        EmitirPedido("abrirReferenciaExterna", id, JsonSerializer.Serialize(new { id, caminho = completo, confirmar }));
        return id;
    }

    // Decisoes do usuario sobre caminhos de vinculo, por caminho EXATO normalizado, valendo para esta sessao (esta ponte e' por sessao).
    private readonly Dictionary<string, bool> _decisoesDeVinculo = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Guarda a resposta do usuario para este caminho de vinculo (so' o host chama: `internal` — publico seria a pagina se autorizando).</summary>
    internal void RegistrarDecisaoDeVinculo(string caminho, bool permitir)
    {
        var chave = CaminhoDeVinculo(caminho) ?? caminho;
        lock (_trava) _decisoesDeVinculo[chave] = permitir;
    }

    /// <summary>Resposta ja' dada pelo usuario para este caminho nesta sessao, ou <c>null</c>.</summary>
    internal bool? DecisaoDeVinculo(string caminho)
    {
        var chave = CaminhoDeVinculo(caminho) ?? caminho;
        lock (_trava) return _decisoesDeVinculo.TryGetValue(chave, out var d) ? d : null;
    }

    /// <summary>Caminho do vinculo se ele PODE virar pedido (planilha, nao negada nesta sessao), com <paramref name="confirmar"/> = o usuario ainda
    /// nao respondeu para ele; <c>null</c> = recusado (so' log). Nenhuma E/S.</summary>
    private string? VinculoPermitido(string? referencia, out bool confirmar)
    {
        confirmar = false;
        var completo = CaminhoDeVinculo(referencia);
        if (completo is null) { _log?.Invoke("VINCULO recusado (nao e' planilha que o mOffice abre): " + Cortar(referencia ?? "")); return null; }
        var decisao = DecisaoDeVinculo(completo);
        if (decisao == false) { _log?.Invoke("VINCULO recusado (o usuario ja' negou nesta sessao): " + completo); return null; }
        confirmar = decisao is null;
        return completo;
    }

    /// <summary>
    /// O caminho LOCAL de <paramref name="referencia"/> (a forma <c>file:</c> vira caminho), ou <c>null</c> quando ele nao pode nem ser considerado:
    /// vazio, URI ilegivel, URL que nao e' <c>file:</c>, caminho de DISPOSITIVO/namespace do Windows (<c>\\?\</c>, <c>\\.\</c>, <c>\??\</c> e as
    /// grafias com <c>/</c>) ou nao totalmente qualificado (relativo, <c>C:x</c>, <c>\x</c>). So' manipulacao de texto: nenhuma chamada ao sistema de
    /// arquivos (fix 1 da revisao da Task 8). `internal static`: testavel e fora da superficie COM.
    /// </summary>
    internal static string? CaminhoBruto(string? referencia)
    {
        if (string.IsNullOrWhiteSpace(referencia)) return null;
        var bruto = referencia;
        // '?' e '*' nunca fazem parte de um caminho de arquivo do Windows (nem de um file: sem consulta): fora ja' na grafia original, antes de
        // qualquer interpretacao como URI (o Uri do .NET reescreve "file:////?/C:/..." de formas que nao devem decidir nada aqui).
        if (EhDispositivo(bruto) || bruto.IndexOfAny(['?', '*']) >= 0) return null;
        if (bruto.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
        {
            if (!Uri.TryCreate(bruto, UriKind.Absolute, out var uri) || !uri.IsFile) return null;
            try { bruto = uri.LocalPath; } catch (InvalidOperationException) { return null; }
            if (uri.IsUnc && (uri.Host is "." or "?" || uri.Host.Length == 0)) return null;
        }
        else if (bruto.Contains("://", StringComparison.Ordinal)) return null;
        // '?' e '*' nunca fazem parte de um nome de arquivo do Windows: qualquer grafia com eles (ex. a forma file: de um caminho \\?\) sai aqui.
        if (EhDispositivo(bruto) || bruto.IndexOfAny(Path.GetInvalidPathChars()) >= 0 || bruto.IndexOfAny(['?', '*']) >= 0) return null;
        return Path.IsPathFullyQualified(bruto) ? bruto : null;
    }

    /// <summary>Pasta do arquivo ORIGINAL salvo desta sessao (base dos vinculos relativos), ou <c>null</c> em documento novo / sem pasta qualificada.</summary>
    private string? PastaDoDocumento
    {
        get
        {
            if (_s.Novo) return null;
            try { var p = Path.GetDirectoryName(_s.ArquivoOrigem); return !string.IsNullOrEmpty(p) && Path.IsPathFullyQualified(p) ? p : null; }
            catch (Exception e) when (e is ArgumentException or PathTooLongException) { return null; }
        }
    }

    /// <summary>
    /// Fix 1b da revisao da Task 8: vinculo RELATIVO (o Excel guarda os vinculos relativos a' pasta de trabalho: <c>custos.xlsx</c>,
    /// <c>..\pasta\custos.xlsx</c>) vira caminho completo pela pasta do documento (<see cref="Path.Combine(string, string)"/> +
    /// <see cref="Path.GetFullPath(string)"/>: so' texto, nenhuma E/S). <c>..</c> que sobe acima do compartilhamento so' resolve — o usuario ve' o caminho
    /// real na pergunta do host. Devolve a referencia como veio quando ela nao e' relativa (URI, dispositivo, <c>?</c>/<c>*</c>, qualificada ou
    /// enraizada pela metade como <c>\x</c>/<c>C:x</c>: quem decide e' <see cref="CaminhoBruto"/>) ou quando nao ha' pasta (documento novo: o relativo
    /// continua relativo e e' recusado). <c>null</c> = nao resolveu. O resultado passa pelas mesmas regras (dispositivo inclusive).
    /// </summary>
    internal static string? ResolverRelativo(string? referencia, string? pastaDoDocumento)
    {
        if (string.IsNullOrWhiteSpace(referencia)) return referencia;
        if (referencia.StartsWith("file:", StringComparison.OrdinalIgnoreCase) || referencia.Contains("://", StringComparison.Ordinal)
            || EhDispositivo(referencia) || referencia.IndexOfAny(['?', '*']) >= 0 || referencia.IndexOfAny(Path.GetInvalidPathChars()) >= 0
            || Path.IsPathRooted(referencia))
            return referencia;
        if (pastaDoDocumento is null || !Path.IsPathFullyQualified(pastaDoDocumento)) return referencia;
        try { return Path.GetFullPath(Path.Combine(pastaDoDocumento, referencia)); }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return null; }
    }

    private static bool EhDispositivo(string s)
    {
        var t = s.TrimStart().Replace('/', '\\');
        return t.StartsWith(@"\\?\", StringComparison.Ordinal) || t.StartsWith(@"\\.\", StringComparison.Ordinal) || t.StartsWith(@"\??\", StringComparison.Ordinal)
            || t.Equals(@"\\?", StringComparison.Ordinal) || t.Equals(@"\\.", StringComparison.Ordinal);
    }

    /// <summary>Caminho NORMALIZADO (completo, como aparece na pergunta ao usuario) de um vinculo de planilha que o mOffice abre, ou <c>null</c>.
    /// <see cref="Path.GetFullPath(string)"/> aqui so' resolve <c>..</c>/barras de um caminho JA' totalmente qualificado (sem E/S); a existencia do
    /// arquivo so' e' vista pelo host DEPOIS de o usuario permitir.</summary>
    internal static string? CaminhoDeVinculo(string? referencia)
    {
        var bruto = CaminhoBruto(referencia);
        if (bruto is null) return null;
        try
        {
            var completo = Path.GetFullPath(bruto);
            if (EhDispositivo(completo)) return null;
            return TabelaFormatos.DaExtensao(completo) is { Abre: true, DocType: "cell" } ? completo : null;
        }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return null; }
    }

    private string? Consumir(HashSet<string> conjunto, string? caminho)
    {
        if (string.IsNullOrWhiteSpace(caminho)) return null;
        string completo;
        try { completo = Path.GetFullPath(caminho); }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return null; }
        lock (_trava)
        {
            if (!conjunto.Contains(completo) || !File.Exists(completo)) return null;
            conjunto.Remove(completo);
            return completo;
        }
    }

    /// <summary>Pasta temporaria do Windows, onde o Word grava as figuras de uma copia (<c>%TEMP%\msohtmlclip1\01\
    /// clip_image002.png</c>). `internal` + `init`: so' os testes apontam outra pasta (ver <see cref="EhClipeDoWord"/>);
    /// todo membro PUBLICO desta classe e' chamavel pelo JS do editor.</summary>
    internal string PastaTempDoSistema { get; init; } = Path.GetTempPath();

    /// <summary>Extensoes de figura que o Word deixa na pasta do clipe e que <see cref="Midia.ExtensaoDaImagem"/>
    /// reconhece pelo conteudo. SVG fica FORA de proposito (divergencia registrada no relatorio da Task 6 do
    /// Plano 3): no caminho de arquivo, <see cref="Midia.ExtensaoDaImagem"/> aceita SVG pela extensao declarada
    /// — um SVG com &lt;script&gt; largado por qualquer processo em <c>%TEMP%\msohtmlclip*\</c> entraria no
    /// documento sem dialogo nenhum. O Word sempre grava tambem um PNG/JPG equivalente, e e' ele que o HTML cita.</summary>
    private static readonly HashSet<string> ExtensoesDoClipe = new(StringComparer.OrdinalIgnoreCase)
        { ".png", ".jpg", ".jpeg", ".gif", ".bmp" };

    /// <summary>
    /// A figura de uma colagem do Word (Task 6 do Plano 3). O HTML que o Word poe na area de transferencia cita
    /// as figuras por <c>&lt;img src="file:///C:/Users/…/AppData/Local/Temp/msohtmlclip1/01/clip_image002.png"&gt;</c>;
    /// o sdkjs pede cada uma a <c>LocalFileGetImageUrl</c> (<c>NativeCorrectImageUrlOnPaste</c>). Nao e' arquivo
    /// escolhido num dialogo do host, entao o <see cref="Resolver"/> recusava — e as figuras sumiam da colagem.
    ///
    /// <c>true</c> SO' quando o caminho, DEPOIS de resolvido (<see cref="Path.GetFullPath(string)"/>: <c>..</c>
    /// e barras repetidas somem aqui), fica DENTRO de <c>&lt;pastaTemp&gt;\msohtmlclip*\</c> — o primeiro
    /// segmento abaixo de <paramref name="pastaTemp"/> comeca com <c>msohtmlclip</c> (sem diferenca de caixa) e
    /// ha' pelo menos um segmento depois dele —, a extensao e' de figura (<see cref="ExtensoesDoClipe"/>) e
    /// nenhum segmento entre <paramref name="pastaTemp"/> e o arquivo (inclusive) e' ponto de reanalise
    /// (junction/symlink: um link ali dentro apontaria a leitura para fora da pasta). Contencao por CAMINHO
    /// RESOLVIDO (<see cref="Path.GetRelativePath(string, string)"/>), nunca por prefixo de texto cru — ver a
    /// mutacao no relatorio. Aceita a forma <c>file:///</c> (com <c>%20</c> etc.) que o HTML do Word traz.
    /// `internal static`: a regra e' pura e testavel sem janela; a pagina nao a alcanca.
    /// </summary>
    internal static bool EhClipeDoWord(string caminhoOuUrl, string pastaTemp)
    {
        if (string.IsNullOrWhiteSpace(caminhoOuUrl) || string.IsNullOrWhiteSpace(pastaTemp)) return false;
        var bruto = caminhoOuUrl;
        if (bruto.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
        {
            if (!Uri.TryCreate(bruto, UriKind.Absolute, out var uri) || !uri.IsFile || uri.IsUnc) return false;
            bruto = uri.LocalPath;
        }
        else if (bruto.Contains("://", StringComparison.Ordinal)) return false;
        string cheio, raiz;
        try
        {
            if (!Path.IsPathFullyQualified(bruto)) return false;
            // Forma LONGA dos dois lados (fix round 1): %TEMP% pode vir em 8.3 (C:\Users\FULANO~1\...) e o HTML do
            // Word na forma longa (ou o contrario) — a contencao por caminho relativo falharia so' pela grafia. Nomes
            // curtos no proprio candidato (MSOHTM~1\CLIP_I~1.PNG) tambem viram longos aqui, antes das regras.
            // MEDIDO: o .NET 10 ja' expande segmentos com '~' dentro de Path.GetFullPath/GetRelativePath; a chamada
            // explicita deixa a regra independente desse detalhe de implementacao do runtime.
            cheio = FormaLonga(Path.GetFullPath(bruto));
            raiz = Path.TrimEndingDirectorySeparator(FormaLonga(Path.GetFullPath(pastaTemp)));
        }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return false; }

        var relativo = Path.GetRelativePath(raiz, cheio);
        if (Path.IsPathRooted(relativo) || relativo == "." || relativo == ".." ||
            relativo.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)) return false;
        var segmentos = relativo.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        if (segmentos.Length < 2) return false;                   // a figura mora DENTRO de msohtmlclip*\, nunca no %TEMP% raiz
        if (!segmentos[0].StartsWith("msohtmlclip", StringComparison.OrdinalIgnoreCase)) return false;
        if (!ExtensoesDoClipe.Contains(Path.GetExtension(cheio))) return false;

        // Nenhum segmento abaixo da pasta temporaria e' link (o proprio arquivo inclusive). Segmento que nao
        // existe nao e' link (e a leitura seguinte falha sozinha); FileSystemInfo.Attributes de algo inexistente
        // devolve -1 (todas as flags), por isso o Exists antes.
        var atual = raiz;
        foreach (var s in segmentos)
        {
            atual = Path.Combine(atual, s);
            try
            {
                if ((File.Exists(atual) || Directory.Exists(atual)) &&
                    File.GetAttributes(atual).HasFlag(FileAttributes.ReparsePoint)) return false;
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return false; }
        }
        return true;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetLongPathNameW(string curto, char[]? longo, uint tamanho);

    /// <summary>Forma longa de um caminho absoluto (nomes 8.3 como <c>FULANO~1</c> expandidos), pela API do
    /// Windows (<c>GetLongPathNameW</c>). So' o que EXISTE tem nome curto: o maior ancestral existente e'
    /// expandido e o resto vai como veio. Qualquer falha -> o caminho como veio (a contencao, entao, so' aceita
    /// se a grafia ja' bater). `internal static`: testavel e fora da superficie COM.</summary>
    internal static string FormaLonga(string caminho)
    {
        var existente = caminho;
        var resto = new Stack<string>();
        while (!(File.Exists(existente) || Directory.Exists(existente)))
        {
            var pai = Path.GetDirectoryName(existente);
            if (string.IsNullOrEmpty(pai)) return caminho;
            resto.Push(Path.GetFileName(existente));
            existente = pai;
        }
        var n = GetLongPathNameW(existente, null, 0);
        if (n == 0) return caminho;
        var buf = new char[n];
        var m = GetLongPathNameW(existente, buf, n);
        if (m == 0 || m >= n) return caminho;
        var longo = new string(buf, 0, (int)m);
        while (resto.Count > 0) longo = Path.Combine(longo, resto.Pop());
        return longo;
    }

    /// <summary>O caminho que o editor le' em <c>LocalFileGetSourcePath</c>. Plano 6, Task 2 — MEDIDO ao vivo: o web-apps reconhece a
    /// apresentacao de slides pela extensao deste caminho e INICIA a apresentacao sozinho ao abrir um <c>.ppsx</c>/<c>.ppsm</c> (o
    /// Editor.bin de um .ppsx e' identico ao do mesmo conteudo em .pptx). Quem abre para EDICAO (Abrir comum) nao pode ver isso; quem
    /// quer apresentar pede pela ponte (<c>iniciarApresentacao</c>). So' o editor ve' o irmao (.pptx/.pptm): o salvar do host usa
    /// sempre <see cref="SessaoDocumento.ArquivoOrigem"/>, que continua o arquivo real.</summary>
    ///
    /// <para>Documento NOVO (I1 da revisao final do Plano 6, A10): a origem da sessao e' a copia do modelo dentro da recovery
    /// (<c>new.docx</c>, ou o nome do modelo), e o sdkjs tira DESTE caminho o nome interno do documento (word <c>Z$</c>: nome do arquivo COM
    /// extensao -> <c>asc_getDocumentName</c>), que o web-apps poe no cabecalho a cada fim de acao longa — o cabecalho virava "new.docx" depois da
    /// 1a edicao. O editor ve' <c>&lt;pasta da copia&gt;\&lt;Titulo&gt;</c> SEM extensao: MEDIDO ao vivo, com "Documento1.docx" o cabecalho virava
    /// "Documento1.docx"; o Word mostra documento nao salvo sem extensao ("Documento1"). Esse caminho nao existe e nunca e' usado para E/S: o host
    /// so' le'/grava pela <see cref="SessaoDocumento.ArquivoOrigem"/>, a ponte esconde "Abrir local do arquivo" em documento novo
    /// (<c>&amp;arquivo=0</c>) e o vinculo relativo nao resolve em documento novo (<c>PastaDoDocumento</c> = null).</para></summary>
    internal static string OrigemParaOEditor(SessaoDocumento s)
    {
        if (s.Novo)
        {
            try { return Path.Combine(Path.GetDirectoryName(s.ArquivoOrigem) ?? "", s.Titulo); }
            catch (ArgumentException) { }
        }
        return s.Formato switch
        {
            FormatoDocumento.Ppsx => Path.ChangeExtension(s.ArquivoOrigem, ".pptx"),
            FormatoDocumento.Ppsm => Path.ChangeExtension(s.ArquivoOrigem, ".pptm"),
            _ => s.ArquivoOrigem,
        };
    }

    /// <summary>Estado que o JS le' uma vez, na carga do documento (as funcoes sincronas da ponte respondem dele).</summary>
    public string SessaoJson() => JsonSerializer.Serialize(new
    {
        origem = OrigemParaOEditor(_s),
        formato = (int)_s.Formato,
        recovery = _s.PastaRecovery.Replace('\\', '/'),
        titulo = _s.Titulo,
        fontes = (_pastaFontes ?? "").Replace('\\', '/'),
        // Recuperacao (Task 2 do Plano 3): CheckUserId responde idUsuario; recuperada/changesAbertos ficam
        // visiveis para o JS e para quem inspeciona a sessao de fora.
        idUsuario = _s.IdUsuario,
        recuperada = _s.Recuperada,
        changesAbertos = _s.ChangesAbertos,
    });

    /// <summary>
    /// Tipos de aviso que SO' a propria ponte emite (pelo <c>_aviso</c> interno), depois de conferir o consentimento: <see cref="PedirConversao"/>,
    /// <see cref="PedirLeitura"/>, <see cref="AbrirReferenciaExterna"/> e <see cref="LiberarConversao"/>. Revisao final do Plano 6, C1 (seguranca): o
    /// host trata esses pedidos confiando no parametro (caminho, <c>confirmar</c>), entao <see cref="Avisar"/> — publico, chamavel por qualquer
    /// script da pagina — os RECUSA. Todos os outros tipos sao os que o <c>ponte-desktop.js</c> manda (a lista dele e' conferida contra esta em
    /// <c>PonteAvisosReservadosTests</c>).
    /// </summary>
    internal static readonly IReadOnlySet<string> TiposSoDoHost = new HashSet<string>(StringComparer.Ordinal)
        { "converterArquivo", "lerArquivo", "abrirReferenciaExterna", "vinculoRecusado", "liberarConversao" };

    /// <summary>Unico caminho JS -> C#. Os tipos de <see cref="TiposSoDoHost"/> sao recusados aqui (so' a ponte os emite, pelo canal interno).</summary>
    public void Avisar(string tipo, string param)
    {
        tipo ??= "";
        if (TiposSoDoHost.Contains(tipo)) { _log?.Invoke("PONTE aviso reservado ao host recusado: " + tipo); return; }
        _aviso?.Invoke(tipo, param ?? "");
    }

    // Pedidos emitidos por esta ponte (id -> tipo + parametro exato entregue ao host). Segunda camada do C1: o host so' executa um pedido cujo id
    // esta' aqui, e usa o parametro GUARDADO (caminho normalizado, formato, uso, confirmar), nunca o que chegou junto com o id.
    private readonly Dictionary<string, (string Tipo, string Param)> _pedidosEmitidos = new(StringComparer.Ordinal);

    /// <summary>Entrega um pedido ao host registrando o id (ver <see cref="ConsumirPedidoEmitido"/>).</summary>
    private void EmitirPedido(string tipo, string id, string param)
    {
        lock (_trava) _pedidosEmitidos[id] = (tipo, param);
        _aviso?.Invoke(tipo, param);
    }

    /// <summary>O parametro que esta ponte emitiu para (<paramref name="tipo"/>, <paramref name="id"/>), ou <c>null</c> (id forjado, de outro tipo
    /// ou ja' usado). Um uso so'. `internal`: so' o host consome.</summary>
    internal string? ConsumirPedidoEmitido(string tipo, string? id)
    {
        if (string.IsNullOrEmpty(id)) return null;
        lock (_trava)
        {
            if (!_pedidosEmitidos.TryGetValue(id, out var p) || p.Tipo != tipo) return null;
            _pedidosEmitidos.Remove(id);
            return p.Param;
        }
    }

    /// <summary>
    /// Semantica (achado 2 da revisao 2026-09-17, espelha o desktop-sdk oficial):
    /// <c>deleteIndex &lt; 0</c> = "sem delecao, so apenda a partir do indice atual" — NAO trunca, so soma
    /// <c>count</c> a <see cref="SessaoDocumento.IndiceChanges"/>. <c>0 &lt;= deleteIndex &lt; IndiceChanges</c>
    /// trunca para os `deleteIndex` primeiros changes (0 e' legitimo: "desfazer tudo" esvazia o arquivo) e
    /// depois apenda os `count` novos, com <c>IndiceChanges = deleteIndex + count</c>.
    /// <c>deleteIndex &gt;= IndiceChanges</c> nao tem nada para truncar; so avanca o indice se `count` &gt; 0.
    /// ANTES: qualquer `deleteIndex &lt; IndiceChanges` (inclusive negativo) truncava — `RemoverChangesAPartirDe`
    /// com indice negativo vira `Take(negativo)` = sequencia vazia = arquivo apagado, destruindo changes que
    /// deveriam so ser apendados. Ver mutacao/teste em PonteDesktopTests.
    ///
    /// <para><b>Invariante (fix round 2 da revisao final do Plano 3):</b> <see cref="SessaoDocumento.IndiceChanges"/>
    /// (e o <see cref="ChangesSalvos"/>, e o <c>estado.indice</c> do JS, que so' avanca com o retorno "ok") so' avancam
    /// DEPOIS de o <c>changes0.json</c> no disco bater com o que o editor acredita ter mandado. Se o arquivo nao puder
    /// ser posto nesse estado (truncamento e apendice com novas tentativas e, no truncamento, o corte no lugar), a
    /// sessao FALHA FECHADA: a ponte congela aqui (C#) e no JS (retorno diferente de "ok"), os contadores ficam onde
    /// estavam, e o host e' avisado (<c>aoFalharGravacao</c>) para fechar a janela sem perguntar, mantendo a pasta.
    /// Nunca contador e arquivo divergindo em silencio: um salvamento depois disso gravaria algo diferente da tela.</para>
    /// </summary>
    /// <returns><c>"ok"</c> (arquivo e indice em dia), <c>"congelada"</c> (ponte ja' congelada: nada gravado) ou
    /// <c>"falhou"</c> (esta chamada congelou a sessao).</returns>
    public string SalvarChanges(string param, int deleteIndex, int count)
    {
        if (LancarEmSalvarChangesParaTeste) throw new InvalidOperationException("falha simulada do objeto de host (teste)");
        string? falha = null;
        lock (_trava)
        {
            if (_congelada) { _log?.Invoke($"RECUPERACAO LocalFileSaveChanges ignorado (ponte congelada): deleteIndex {deleteIndex}, {count} changes"); return "congelada"; }
            try
            {
                GravarChanges(param, deleteIndex, count);
                _geracaoDeGravacao++;   // fix round 4: toda gravacao que deu certo (truncamentos inclusive) muda a geracao
            }
            catch (Exception ex)
            {
                _congelada = true;
                falha = ex.Message;
                _log?.Invoke($"ALTERACOES changes0.json NAO gravado (deleteIndex {deleteIndex}, {count} changes): {ex.Message} — a ponte congela e o indice fica em {_s.IndiceChanges}");
            }
        }
        // Fora da trava: quem assina pode ler o estado da sessao (ou chamar de volta a ponte) sem impasse.
        if (falha is not null)
        {
            AvisarCongelou();
            try { _aoFalharGravacao?.Invoke(falha); }
            catch (Exception ex) { _log?.Invoke("PONTE erro no aviso de falha ao gravar alteracoes: " + ex.Message); }
            return "falhou";
        }
        try { ChangesSalvos?.Invoke(); }
        catch (Exception ex) { _log?.Invoke("PONTE erro no aviso de changes salvos: " + ex.Message); }
        return "ok";
    }

    /// <summary>Corpo de <see cref="SalvarChanges"/>, SEMPRE chamado com <see cref="_trava"/> na mao. Primeiro o
    /// arquivo, DEPOIS o indice (invariante acima): uma excecao em qualquer passo sai daqui sem tocar no indice.</summary>
    private void GravarChanges(string param, int deleteIndex, int count)
    {
        int novoIndice;
        var truncar = false;
        if (deleteIndex < 0) novoIndice = _s.IndiceChanges + count;
        else if (deleteIndex < _s.IndiceChanges)
        {
            // Truncamento RELATIVO A ESTA SESSAO (protocolo oficial, Task 2 do Plano 3): o editor numera
            // os changes a partir do que ELE gravou nesta abertura; numa pasta recuperada o arquivo ja'
            // comeca com `ChangesAbertos` changes da sessao que caiu (reaplicados antes do EndLoad), e esse
            // prefixo nunca e' desfeito por um deleteIndex desta sessao. Numa abertura normal
            // ChangesAbertos = 0 e nada muda.
            truncar = true;
            novoIndice = deleteIndex + count;
        }
        else if (count > 0) novoIndice = deleteIndex + count;
        else novoIndice = _s.IndiceChanges;   // deleteIndex >= IndiceChanges e count == 0: nada muda

        if (truncar) RemoverChangesAPartirDe(_s.ChangesAbertos + deleteIndex);
        if (count > 0)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_s.ArquivoChanges)!);
            // O JS manda os changes ja unidos por "\",\"" (arrayChanges.join('","')) quando count <= 100000.
            Apendar(Encoding.UTF8.GetBytes("\"" + param + "\","));
        }
        _s.IndiceChanges = novoIndice;
    }

    /// <summary>Apende o lote ao <c>changes0.json</c>. So' a ABERTURA tenta de novo (arquivo preso por um instante:
    /// antivirus, outra instancia lendo) — uma escrita que falha depois de aberta desfaz o que escreveu
    /// (<c>SetLength</c> de volta) e sobe, para nunca repetir meio lote no arquivo.</summary>
    private void Apendar(byte[] lote)
    {
        FileStream? fs = null;
        Recuperacao.GravacaoAtomica.ComNovasTentativas(() =>
            fs = new FileStream(_s.ArquivoChanges, FileMode.Append, FileAccess.Write, FileShare.Read | FileShare.Delete));
        using (fs)
        {
            var antes = fs!.Length;
            try { fs.Write(lote); fs.Flush(); }
            catch
            {
                try { fs.SetLength(antes); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
                throw;
            }
        }
    }

    /// <summary><c>AscDesktopEditor.SpellCheck</c> (JSON) -&gt; JSON da resposta, ou <c>""</c> quando nao ha'
    /// resposta. Chamado pelo proxy ASSINCRONO do objeto de host (<c>chrome.webview.hostObjects.ponte</c>, sem
    /// <c>.sync</c>): o renderer nao fica bloqueado esperando. O metodo em si roda na thread da interface (e'
    /// assim que o WebView2 despacha objeto de host), entao a carga do dicionario e' feita ANTES, em segundo
    /// plano (<see cref="Ortografia.Corretor.PreCarregarAsync"/> na inicializacao da janela); se um pedido
    /// chegar antes de ela terminar, ele espera a carga em andamento (nao carrega de novo).</summary>
    public string VerificarOrtografia(string pedido)
    {
        try { return _ortografia?.Invoke(pedido ?? "") ?? ""; }
        catch (Exception ex) { _log?.Invoke("ORTOGRAFIA " + ex.Message); return ""; }
    }

    /// <summary><c>onDocumentModifiedChanged</c> do sdkjs. Avisa o host so' quando o valor MUDA (Task 2 do
    /// Marco 1 / Plano 2): o editor chama isto varias vezes com o mesmo valor, e a casca redesenharia o
    /// asterisco a cada tecla. O callback sai DEPOIS de gravar, para quem o receber ja ler o estado novo em
    /// <see cref="SessaoDocumento.Modificado"/>.</summary>
    public void MarcarModificado(bool modificado)
    {
        var mudou = _s.Modificado != modificado;
        _s.Modificado = modificado;
        if (mudou) _aoMudarModificado?.Invoke(modificado);
    }

    /// <summary>
    /// <c>LocalFileGetImageUrl(dados)</c> — o pedido do sdkjs para transformar uma imagem em midia DO
    /// DOCUMENTO. Devolve o NOME do arquivo dentro de <c>media/</c> (nunca a URL: o sdkjs monta
    /// <c>AscCommon.Ss.tia + "/media/" + &lt;isto&gt;</c>, <c>fe.RS</c>/<c>fe.vdc</c> em sdk-all.js), ou
    /// <c>"error"</c>, exatamente como o <c>GetLocalImageUrl</c> do app oficial
    /// (desktop-sdk v9.4.0.129, client_renderer_wrapper.cpp:5269).
    ///
    /// Quem chama, no sdkjs desta versao:
    /// <list type="bullet">
    /// <item><c>window.NativeCorrectImageUrlOnPaste(url)</c> — COLAR (sdk-all.js:9910597);</item>
    /// <item><c>Asc.asc_docs_api.prototype.AddImageUrl</c> (<c>nQc</c>, sdk-all.js:9918718) e
    /// <c>asc_addImage</c> (<c>qUa</c>) no callback do <c>OpenFilenameDialog</c> — INSERIR do arquivo;</item>
    /// <item><c>AscCommon.jT</c> e o <c>ondrop</c> (sdk-all.js:9909389/9912107) — arrastar e soltar.</item>
    /// </list>
    ///
    /// Divergencias DELIBERADAS em relacao ao app oficial, ambas por causa do egresso zero:
    /// <list type="number">
    /// <item>URL <c>http(s)/ftp/www.</c> NAO e' baixada (o oficial chama <c>CFileDownloaderWrapper</c>):
    /// devolve <c>"error"</c> e registra no log. Este host nao sai para a rede.</item>
    /// <item>Caminho de arquivo so' e' lido se o USUARIO o escolheu num dialogo do host
    /// (<see cref="PermitirArquivoEscolhido"/>), espelhando o <c>g_pLocalResolver->Check</c> do oficial:
    /// sem isso, conteudo web poderia mandar copiar QUALQUER arquivo da maquina para dentro de <c>media/</c>,
    /// que e' servido para a propria pagina por <c>https://moffice-doc.local</c>. Excecao (Task 6 do Plano 3):
    /// a figura de uma colagem do Word em <c>%TEMP%\msohtmlclip*\</c> (<see cref="EhClipeDoWord"/>) — sem ela
    /// as figuras sumiam da colagem (medido: 0 arquivo em media/ sem a regra, 1 com ela).</item>
    /// </list>
    /// </summary>
    public string ImagemUrl(string dados)
    {
        try { return Resolver(dados ?? ""); }
        catch (Exception ex)
        {
            _log?.Invoke("MIDIA erro ao resolver imagem: " + ex.Message);
            return "error";
        }
    }

    private string Resolver(string dados)
    {
        if (dados.Length == 0) return "error";

        // Ja e' midia deste documento (o oficial: `if (sUrl.find(L"image") == 0 || sUrl.find(L"display") == 0)
        // return sUrl;`). Nome, nunca caminho: barra ou ".." aqui viraria pedido fora de media/.
        if ((dados.StartsWith("image", StringComparison.Ordinal) || dados.StartsWith("display", StringComparison.Ordinal))
            && dados.IndexOfAny(['/', '\\']) < 0 && !dados.Contains(".."))
            return dados;

        if (dados.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            // SVG e' recusado DE PROPOSITO, nao "conteudo nao reconhecido" (minor 3 da revisao 1): e' XML e
            // pode trazer um <script> valido dentro — o editor exibiria isso como midia do documento sem
            // nenhuma sandbox adicional por cima. Erro PROPRIO (nao o "error" generico) para a decisao ficar
            // visivel para quem olha de fora, em vez de se confundir com "base64 ilegivel"/"nao e' imagem".
            // So' se aplica a `data:` (colar/arrastar, conteudo web): um .svg ESCOLHIDO PELO USUARIO num
            // dialogo do host ainda passa por <see cref="Midia.ExtensaoDaImagem"/> normalmente — ver o
            // comentario la'.
            if (dados.StartsWith("data:image/svg+xml", StringComparison.OrdinalIgnoreCase))
            {
                _log?.Invoke("MIDIA svg recusado de proposito (formato com script): " + Cortar(dados));
                return "formato-nao-suportado";
            }
            var marca = dados.IndexOf("base64,", StringComparison.OrdinalIgnoreCase);
            if (marca < 0) { _log?.Invoke("MIDIA data: sem base64 recusada"); return "error"; }
            var base64 = dados[(marca + 7)..];
            // 4 caracteres de base64 = 3 bytes: corta ANTES de alocar o vetor decodificado.
            if ((long)base64.Length / 4 * 3 > MaximoBytesImagem) { _log?.Invoke("MIDIA data: grande demais"); return "error"; }
            var bytes = Convert.FromBase64String(base64);
            // Farejo do CONTEUDO (fix 2 da revisao da Task 5, item 3): o MIME declarado e' da PAGINA, entao
            // `data:image/png;base64,<texto svg>` escapava da recusa propria acima e caia no "error" generico
            // so' por acaso (ExtensaoDaImagem nao reconhece o texto). Qualquer `data:` cujos bytes decodificados
            // sao texto XML/HTML/SVG (comecam com '<', pulando BOM e espacos) recebe a MESMA recusa proposital.
            // Poliglota (magico PNG + SVG anexado): passa por aqui como png de verdade e e' gravado como
            // .png — o host virtual serve como image/png, o navegador decodifica como bitmap e nao executa nada
            // (um <script> dentro de um bitmap e' so' lixo apos o IEND). Nao precisa de defesa extra.
            if (Midia.PareceXml(bytes))
            {
                _log?.Invoke("MIDIA svg/xml recusado de proposito (conteudo XML sob MIME declarado): " + Cortar(dados));
                return "formato-nao-suportado";
            }
            return Gravar(bytes, declarada: null);
        }

        if (EhUrlDeRede(dados))
        {
            // O app oficial BAIXARIA. Aqui nao: egresso zero e' propriedade do produto, nao detalhe.
            _log?.Invoke("MIDIA recusada (URL de rede, este host nao baixa): " + Cortar(dados));
            return "error";
        }

        // Figura de uma colagem do Word (Task 6 do Plano 3): aceita como se fosse arquivo escolhido, SEM consumir
        // consentimento nenhum — o gesto do usuario aqui foi o Ctrl+V/Colar, e a figura so' existe dentro de
        // %TEMP%\msohtmlclip*\ (regra inteira em EhClipeDoWord). Mesmo teto de tamanho do arquivo escolhido.
        if (EhClipeDoWord(dados, PastaTempDoSistema))
        {
            var clipe = CaminhoLocal(dados);
            var info = new FileInfo(clipe);
            if (!info.Exists) { _log?.Invoke("MIDIA clipe do Word nao existe mais: " + Cortar(dados)); return "error"; }
            if (info.Length > MaximoBytesImagem) { _log?.Invoke("MIDIA clipe do Word grande demais: " + info.Length + " bytes"); return "error"; }
            _log?.Invoke("MIDIA clipe do Word aceito: " + Path.GetFileName(clipe));
            return Gravar(File.ReadAllBytes(clipe), Path.GetExtension(clipe));
        }

        // Consentimento de UM USO SO' (minor 2 da revisao 1; atomico desde o fix 2 da revisao da Task 5): a
        // entrada e' CONSUMIDA dentro do mesmo lock da checagem, em CaminhoDeArquivo, ANTES da leitura — antes
        // disso o Contains e o Remove ficavam em dois locks separados com a leitura + Gravar no meio, e dois
        // pedidos simultaneos do mesmo caminho podiam passar os dois pela checagem. Um SEGUNDO pedido do MESMO
        // caminho sem um novo dialogo do host tem de ser recusado: o dialogo autoriza aquela escolha, nao um
        // acesso permanente ao arquivo (o app oficial tambem nao reusa `g_pLocalResolver->AddFile` para sempre
        // — ver o comentario de <see cref="PermitirArquivoEscolhido"/>).
        var caminho = CaminhoDeArquivo(dados);
        if (caminho is null) { _log?.Invoke("MIDIA recusada (nao e' arquivo escolhido pelo usuario): " + Cortar(dados)); return "error"; }
        try
        {
            var info = new FileInfo(caminho);
            if (info.Length > MaximoBytesImagem)
            {
                _log?.Invoke("MIDIA arquivo grande demais: " + info.Length + " bytes");
                DevolverConsentimento(caminho);
                return "error";
            }
            return Gravar(File.ReadAllBytes(caminho), Path.GetExtension(caminho));
        }
        catch
        {
            // "Falha nao consome": se a leitura ou o Gravar lancar (conteudo escolhido nao e' imagem), nao
            // houve "uso" nenhum — a entrada volta ao conjunto e o dialogo continua valendo para uma nova
            // tentativa no mesmo caminho. Ver ImagemUrl_ArquivoEscolhidoQueNaoEImagem_NaoConsomeAEntrada.
            DevolverConsentimento(caminho);
            throw;
        }
    }

    /// <summary>Desfaz um <see cref="PermitirArquivoEscolhido"/> que nao foi (ou nao pode mais ser) usado — a
    /// insercao de uma imagem solta que terminou (Task 6 do Plano 3, fix round 1). Devolve <c>true</c> se ainda
    /// havia a entrada. `internal` pelo mesmo motivo de <see cref="PermitirArquivoEscolhido"/>.</summary>
    internal bool RevogarArquivoEscolhido(string caminho)
    {
        if (string.IsNullOrWhiteSpace(caminho)) return false;
        string completo;
        try { completo = Path.GetFullPath(caminho); }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return false; }
        lock (_trava) return _escolhidosPeloUsuario.Remove(completo);
    }

    private void DevolverConsentimento(string caminho)
    {
        lock (_trava) _escolhidosPeloUsuario.Add(caminho);
    }

    /// <summary>Caminho local que o usuario escolheu num dialogo do host — ou <c>null</c>. Aceita tambem a
    /// forma <c>file:///…</c> (o oficial trata a mesma, porque o Word do Windows/mac cola imagem com ela).
    /// CONSOME o consentimento (remove a entrada) no mesmo lock da checagem: quem chama e' responsavel por
    /// devolve-lo (<see cref="DevolverConsentimento"/>) se a leitura falhar antes de haver "uso".</summary>
    private string? CaminhoDeArquivo(string dados)
    {
        var bruto = dados;
        if (bruto.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
        {
            if (!Uri.TryCreate(bruto, UriKind.Absolute, out var uri) || !uri.IsFile) return null;
            bruto = uri.LocalPath;
        }
        string completo;
        try { completo = Path.GetFullPath(bruto); }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return null; }
        lock (_trava)
        {
            if (!_escolhidosPeloUsuario.Contains(completo)) return null;
            // Arquivo sumiu entre o dialogo e o pedido: nao ha' o que consumir (nem o que ler).
            if (!File.Exists(completo)) return null;
            _escolhidosPeloUsuario.Remove(completo);
            return completo;
        }
    }

    /// <summary>Caminho local de um <c>file:///…</c> (desfazendo <c>%20</c> etc.) ou do proprio texto, ja' resolvido.
    /// So' chamado depois de <see cref="EhClipeDoWord"/> aceitar o mesmo valor.</summary>
    private static string CaminhoLocal(string dados) =>
        Path.GetFullPath(dados.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? new Uri(dados).LocalPath : dados);

    private static bool EhUrlDeRede(string s) =>
        s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        s.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
        s.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
        s.StartsWith("www.", StringComparison.OrdinalIgnoreCase);

    private string Gravar(byte[] bytes, string? declarada)
    {
        var chave = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));
        lock (_trava)
        {
            if (_midiaPorConteudo.TryGetValue(chave, out var jaGravado)
                && File.Exists(Path.Combine(Midia.PastaDeMidia(_s), jaGravado)))
                return jaGravado;
        }
        // Chamada SINCRONA vinda do JS (thread da interface, renderer bloqueado esperando o retorno): bloquear
        // aqui so' e' seguro porque Midia.GravarImagemAsync usa ConfigureAwait(false) de ponta a ponta — sem
        // isso a continuacao voltaria para a thread da interface, que esta parada neste GetResult(), e o
        // impasse seria certo.
        var nome = Midia.GravarImagemAsync(_s, bytes, declarada ?? "").GetAwaiter().GetResult();
        lock (_trava) _midiaPorConteudo[chave] = nome;
        _log?.Invoke($"MIDIA gravada {nome} ({bytes.Length} bytes)");
        return nome;
    }

    private static string Cortar(string s) => s.Length <= 120 ? s : s[..120] + "…";

    private void RemoverChangesAPartirDe(int indice)
    {
        // HIPOTESE (o RemoveChanges do app oficial nao foi lido): o arquivo e' a sequencia "c1","c2",... —
        // manter os `indice` primeiros. A Task 6 confirma com changes reais.
        if (!File.Exists(_s.ArquivoChanges)) return;
        var partes = SepararChanges(Encoding.UTF8.GetString(Recuperacao.GravacaoAtomica.LerCompartilhado(_s.ArquivoChanges))).Take(indice).ToList();
        if (partes.Count == 0)
        {
            // "Desfazer tudo": o arquivo sai. Preso por um leitor que nao compartilha delecao, fica VAZIO (corte no
            // lugar) — o mesmo "zero changes" para a recuperacao e para o x2t.
            try { Recuperacao.GravacaoAtomica.ComNovasTentativas(() => File.Delete(_s.ArquivoChanges)); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { Recuperacao.GravacaoAtomica.Truncar(_s.ArquivoChanges, []); }
            return;
        }
        // Atomico (revisao final do Plano 3, M2): o changes0.json e' a UNICA copia do trabalho nao salvo. O
        // File.WriteAllText direto truncava o arquivo antes de escrever o prefixo — uma queda no meio deixava o
        // arquivo vazio ou rasgado. Agora: .tmp + flush + troca; se a troca nao for possivel (leitor segurando o
        // arquivo, fix round 2), o corte NO LUGAR do fim do arquivo (GravacaoAtomica.Truncar). Se nem isso: lanca, e
        // SalvarChanges falha fechado.
        Recuperacao.GravacaoAtomica.Truncar(_s.ArquivoChanges, new UTF8Encoding(false).GetBytes(string.Concat(partes.Select(p => "\"" + p + "\","))));
    }

    private static IEnumerable<string> SepararChanges(string s)
    {
        var i = 0;
        while (i < s.Length)
        {
            var ini = s.IndexOf('"', i); if (ini < 0) yield break;
            var fim = s.IndexOf("\",", ini + 1, StringComparison.Ordinal); if (fim < 0) yield break;
            yield return s.Substring(ini + 1, fim - ini - 1);
            i = fim + 2;
        }
    }
}
