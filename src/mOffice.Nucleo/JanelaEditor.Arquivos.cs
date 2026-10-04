using System.Text.Json;
using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>
/// Plano 6, Task 8 — recursos do motor que pedem ARQUIVO ao host: dialogo por filtro (<see cref="FiltrosDoEditor"/>), conversao para o binario do
/// editor (<c>convertFile</c>: Inserir texto do arquivo, vinculo externo, importar XML), Comparar/Combinar (<c>onDocumentCompare</c>/
/// <c>onDocumentMerge</c>), leitura de CSV/TXT (<c>loadLocalFile</c>), "Abrir origem" de vinculo e "Abrir recente" (<c>onupdaterecents</c>).
/// Tudo chega por avisos da ponte DENTRO da chamada sincrona do JS e roda fora dela (<c>Task.Yield</c>/<c>Dispatcher</c>); o editor e' sempre
/// respondido (mesmo em erro), senao o recurso fica preso. Os arquivos escolhidos sao entradas SO' DE LEITURA: nada e' escrito neles.
/// </summary>
public partial class JanelaEditor
{
    /// <summary>Costura de teste do dialogo de arquivo (todos os filtros): recebe o filtro do motor e o <c>multi</c> e devolve os caminhos
    /// escolhidos, ou <c>null</c> = cancelou — no lugar do <c>OpenFileDialog</c> do Windows (modal). Para "images",
    /// <see cref="EscolherImagemParaTeste"/>, se presente, continua valendo. `internal`: nao e' API do host.</summary>
    internal Func<string, bool, string[]?>? EscolherArquivoParaTeste { get; set; }

    /// <summary>Recentes do app (caminhos, do mais recente para o mais antigo) para o "Abrir recente" do menu Arquivo do editor. Fornecido pela
    /// casca (o nucleo nao conhece o arquivo de recentes); sem ele o menu fica vazio, como antes.</summary>
    public Func<IReadOnlyList<string>>? FonteDeRecentes { get; init; }

    /// <summary>Host virtual dos dicionarios de hifenizacao (<c>motor\dictionaries</c>): o que <c>getDictionariesPath</c> devolve na ponte.</summary>
    internal const string HostDicionarios = "moffice-dic.local";

    /// <summary>Faixa quando o editor pede um dialogo de audio/video (so' com <c>IsSupportMedia</c>, que o mOffice responde false).</summary>
    public const string MensagemMidiaSemSuporte = "Áudio e vídeo não são suportados no mOffice.";

    /// <summary>Teto do que a leitura de CSV/TXT entrega ao editor (o conteudo vai em base64 por um script).</summary>
    internal const long MaximoBytesLeitura = 64L * 1024 * 1024;

    private ConversaoParaEditor? _conversaoParaEditor;
    private ConversaoParaEditor Conversao => _conversaoParaEditor ??= new ConversaoParaEditor(_conv);

    /// <summary>Conversoes entregues/terminadas (so' para os testes esperarem de forma deterministica). `internal`.</summary>
    internal int ConversoesTerminadas => Volatile.Read(ref _conversoesTerminadas);
    private int _conversoesTerminadas;

    /// <summary>Ultimo filtro de dialogo pedido pelo editor (para os testes). `internal`.</summary>
    internal string? UltimoFiltroDeDialogo { get; private set; }

    /// <summary>Avisos desta parte da ponte. <c>true</c> = tratado. Roda DENTRO da chamada sincrona do JS: nada de WebView aqui.</summary>
    private bool TratarAvisoDeArquivos(SessaoDocumento origem, string tipo, string param)
    {
        switch (tipo)
        {
            case "abrirArquivo":
                var (filtro, multi) = LerPedidoDeDialogo(param);
                _ = ResponderDialogoArquivoAsync(filtro, multi);
                return true;
            // C1 da revisao final do Plano 6: so' pedido que a ponte DESTA sessao emitiu (id conhecido), com o parametro que ELA guardou.
            case "converterArquivo":
                if (ReferenceEquals(origem, _sessao) && PedidoEmitido(_ponte, tipo, param, _arquivos.Log) is { } conv) _ = ConverterEEntregarAsync(origem, conv);
                return true;
            case "lerArquivo":
                if (ReferenceEquals(origem, _sessao) && PedidoEmitido(_ponte, tipo, param, _arquivos.Log) is { } leit) _ = LerEEntregarAsync(origem, leit);
                return true;
            case "liberarConversao":
                ConversaoParaEditor.Liberar(origem.PastaRecovery, param);
                return true;
            case "abrirReferenciaExterna":
                if (ReferenceEquals(origem, _sessao) && PedidoEmitido(_ponte, tipo, param, _arquivos.Log) is { } refe) _ = AbrirReferenciaEEntregarAsync(origem, refe);
                return true;
            case "vinculoRecusado":
                if (ReferenceEquals(origem, _sessao)) Dispatcher.BeginInvoke(() => { if (!_fechada) AvisarUsuario(MensagemVinculoRecusado(param)); });
                return true;
            case "recentes":
                if (ReferenceEquals(origem, _sessao)) _ = EntregarRecentesAsync();
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Segunda camada do C1 (revisao final do Plano 6): o parametro GUARDADO pela ponte para o id deste pedido, ou <c>null</c> (id que a ponte nao
    /// emitiu, de outro tipo ou ja' usado — recusado e registrado). O caminho/<c>confirmar</c> que chegam junto com o id sao ignorados.
    /// `internal static`: testavel sem janela.
    /// </summary>
    internal static string? PedidoEmitido(PonteDesktop? ponte, string tipo, string param, Action<string> log)
    {
        string? id = null;
        try
        {
            using var d = JsonDocument.Parse(param);
            if (d.RootElement.ValueKind == JsonValueKind.Object && d.RootElement.TryGetProperty("id", out var pid) && pid.ValueKind == JsonValueKind.String)
                id = pid.GetString();
        }
        catch (JsonException) { }
        var guardado = ponte?.ConsumirPedidoEmitido(tipo, id);
        if (guardado is null) log("PONTE pedido nao emitido pela ponte recusado: " + tipo);
        return guardado;
    }

    /// <summary><c>{filtro, multi}</c> do aviso "abrirArquivo"; texto que nao e' JSON = so' o filtro (forma antiga).</summary>
    internal static (string Filtro, bool Multi) LerPedidoDeDialogo(string param)
    {
        try
        {
            using var d = JsonDocument.Parse(param);
            var f = d.RootElement.TryGetProperty("filtro", out var pf) && pf.ValueKind == JsonValueKind.String ? pf.GetString() ?? "" : "";
            var m = d.RootElement.TryGetProperty("multi", out var pm) && pm.ValueKind == JsonValueKind.True;
            return (f, m);
        }
        catch (JsonException) { return (param, false); }
    }

    /// <summary>
    /// <c>OpenFilenameDialog(filtro, multi, cb)</c> — o <c>ResponderDialogoImagemAsync</c> de antes generalizado (Plano 6, Task 8). O dialogo e' aberto
    /// FORA da chamada sincrona do JS (<c>await Task.Yield()</c>) e o callback pendente e' SEMPRE respondido (inclusive com erro e com filtro sem
    /// suporte): o sdkjs guarda um callback pendente por vez e deixa-lo pendurado mataria o recurso ate' recarregar o editor. Cada caminho escolhido
    /// recebe o consentimento de UM uso que o filtro autoriza (<see cref="FiltrosDoEditor"/>). Resposta: string (um) ou array (multi).
    /// </summary>
    private async Task ResponderDialogoArquivoAsync(string filtro, bool multi)
    {
        await Task.Yield();
        UltimoFiltroDeDialogo = filtro;
        string[]? escolhidos = null;
        var pedido = FiltrosDoEditor.Para(filtro);
        try
        {
            if (pedido is null)
            {
                _arquivos.Log("PONTE dialogo sem suporte (filtro " + filtro + ")");
                AvisarUsuario(MensagemMidiaSemSuporte);
            }
            else
            {
                if (filtro == "images" && EscolherImagemParaTeste is not null)
                    escolhidos = EscolherImagemParaTeste() is { } img ? [img] : null;
                else if (EscolherArquivoParaTeste is not null)
                    escolhidos = EscolherArquivoParaTeste(filtro, multi);
                else
                    escolhidos = EscolherArquivoNoWindows(pedido, multi);
                escolhidos = escolhidos?.Where(c => !string.IsNullOrWhiteSpace(c)).ToArray();
                if (escolhidos is { Length: 0 }) escolhidos = null;
                if (escolhidos is not null && !multi && escolhidos.Length > 1) escolhidos = [escolhidos[0]];
                foreach (var c in escolhidos ?? [])
                {
                    switch (pedido.Consentimento)
                    {
                        case Consentimento.Imagem: _ponte?.PermitirArquivoEscolhido(c); break;
                        case Consentimento.Conversao: _ponte?.PermitirConversao(c); break;
                        case Consentimento.Leitura: _ponte?.PermitirLeitura(c); break;
                    }
                }
                _arquivos.Log("PONTE dialogo de arquivo (" + filtro + "): " + (escolhidos is null ? "cancelado" : string.Join(", ", escolhidos.Select(Path.GetFileName))));
            }
        }
        catch (Exception ex)
        {
            escolhidos = null;
            _arquivos.Log("PONTE dialogo de arquivo falhou (" + filtro + "): " + ex.Message);
        }
        var resposta = escolhidos is null ? "\"\""
            : multi ? JsonSerializer.Serialize(escolhidos)
            : JsonSerializer.Serialize(escolhidos[0]);
        var js = """
            (function () {
              try {
                var w = (window.__moffice && window.__moffice.dialogoPendente) ? window : window.frames[0];
                if (!w || !w.__moffice || typeof w.__moffice.responderDialogo !== "function") return "sem ponte no quadro";
                return "ok " + w.__moffice.responderDialogo(RESPOSTA);
              } catch (e) { return "erro: " + e; }
            })()
            """.Replace("RESPOSTA", resposta);
        try
        {
            var r = (await ExecutarJsAsync(js)).Trim('"');
            if (!r.StartsWith("ok", StringComparison.Ordinal)) _arquivos.Log("PONTE resposta do dialogo de arquivo: " + r);
        }
        catch (Exception ex) { _arquivos.Log("PONTE nao conseguiu responder o dialogo de arquivo: " + ex.Message); }
    }

    private static string[]? EscolherArquivoNoWindows(PedidoDeArquivo pedido, bool multi)
    {
        var dialogo = new Microsoft.Win32.OpenFileDialog
        {
            Title = pedido.Titulo,
            Filter = pedido.Filtro,
            CheckFileExists = true,
            Multiselect = multi,
        };
        return dialogo.ShowDialog() == true ? dialogo.FileNames : null;
    }

    private sealed record PedidoConversao(string id, string caminho, int formato, string uso, bool confirmar = false);

    // -----------------------------------------------------------------------------------------------------
    // Vinculo externo (fix 1 da revisao da Task 8, seguranca): o caminho vem do DOCUMENTO, nao de um dialogo do host. Antes de QUALQUER E/S nele
    // (File.Exists, x2t, abrir), o usuario confirma vendo o caminho completo — padrao Nao —, e a resposta vale por caminho nesta sessao
    // (PonteDesktop.RegistrarDecisaoDeVinculo). O "Atualizar" do proprio editor (que nem mostra os caminhos) nao conta como consentimento.
    // -----------------------------------------------------------------------------------------------------

    /// <summary>Confirmacao do vinculo externo (<see cref="MensagemConfirmarVinculo"/>): <c>true</c> = permitir. A casca injeta o dialogo do design
    /// system; sem ela, um <c>MessageBox</c> com Nao como padrao.</summary>
    public Func<string, bool>? ConfirmarVinculo { get; init; }

    /// <summary>Costura de teste da confirmacao do vinculo (recebe a mensagem). `internal`.</summary>
    internal Func<string, bool>? ConfirmarVinculoParaTeste { get; set; }

    public static string MensagemConfirmarVinculo(string documento, string caminho) =>
        $"Esta planilha (\"{documento}\") quer ler dados de:\n\n{caminho}\n\nPermitir?";

    public static string MensagemVinculoRecusado(string caminho) =>
        $"O mOffice não lê vínculos para este caminho (dispositivo do Windows ou caminho incompleto): {(caminho.Length > 200 ? caminho[..200] + "…" : caminho)}";

    /// <summary>Faixa quando o arquivo convertido para o editor passa de <see cref="MaximoBytesLeitura"/>.</summary>
    public const string MensagemConversaoGrandeDemais = "O arquivo é grande demais para ser inserido ou comparado pelo mOffice.";

    private readonly SemaphoreSlim _confirmacaoDeVinculo = new(1, 1);

    /// <summary>Uma pergunta por vez (o motor pode pedir varios vinculos de uma vez); a decisao ja' dada para o mesmo caminho e' relida depois de
    /// esperar a vez, para nao perguntar duas vezes. Falha na pergunta = nao permite (e nao guarda).</summary>
    private async Task<bool> ConfirmarVinculoAsync(SessaoDocumento origem, string caminho)
    {
        var ponte = _ponte;
        if (ponte is null || !ReferenceEquals(origem, _sessao)) return false;
        await _confirmacaoDeVinculo.WaitAsync();
        try
        {
            if (ponte.DecisaoDeVinculo(caminho) is { } ja) return ja;
            if (!ReferenceEquals(origem, _sessao) || _fechada) return false;
            bool sim;
            try { sim = (ConfirmarVinculoParaTeste ?? ConfirmarVinculo ?? ConfirmarVinculoPadrao)(MensagemConfirmarVinculo(origem.Titulo, caminho)); }
            catch (Exception ex) { _arquivos.Log("VINCULO erro na confirmacao (nao permite): " + ex.Message); return false; }
            ponte.RegistrarDecisaoDeVinculo(caminho, sim);
            _arquivos.Log("VINCULO " + (sim ? "permitido" : "negado") + " pelo usuario: " + caminho);
            return sim;
        }
        finally { _confirmacaoDeVinculo.Release(); }
    }

    private bool ConfirmarVinculoPadrao(string mensagem) =>
        System.Windows.MessageBox.Show(this, mensagem, NomeApp, System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question,
            System.Windows.MessageBoxResult.No) == System.Windows.MessageBoxResult.Yes; // Enter = Nao

    private sealed record PedidoReferencia(string id, string caminho, bool confirmar);

    /// <summary>"Abrir origem": confirma (se preciso), so' DEPOIS confere a existencia, pede a abertura a' casca (janela nova) e responde ao motor
    /// por <c>__moffice.entregarReferencia(id, 0|1)</c>.</summary>
    private async Task AbrirReferenciaEEntregarAsync(SessaoDocumento origem, string param)
    {
        await Task.Yield();
        PedidoReferencia? p = null;
        try { p = JsonSerializer.Deserialize<PedidoReferencia>(param); } catch (JsonException) { }
        if (p is null || string.IsNullOrEmpty(p.id)) { _arquivos.Log("VINCULO pedido ilegivel"); return; }
        var abriu = false;
        if (!p.confirmar || await ConfirmarVinculoAsync(origem, p.caminho))
        {
            if (File.Exists(p.caminho))
            {
                try { AbrirPedido?.Invoke(p.caminho); abriu = true; }
                catch (Exception ex) { _arquivos.Log("VINCULO erro no pedido de abertura: " + ex.Message); }
            }
            else _arquivos.Log("VINCULO origem nao encontrada: " + p.caminho);
        }
        if (!ReferenceEquals(origem, _sessao) || _fechada) return;
        var js = $$"""
            (function () {
              try {
                var w = window.frames[0];
                if (!w || !w.__moffice || typeof w.__moffice.entregarReferencia !== "function") return "sem ponte no quadro";
                return "ok " + w.__moffice.entregarReferencia({{JsonSerializer.Serialize(p.id)}}, {{(abriu ? 1 : 0)}});
              } catch (e) { return "erro: " + e; }
            })()
            """;
        try
        {
            var res = (await ExecutarJsAsync(js)).Trim('"');
            if (!res.StartsWith("ok", StringComparison.Ordinal)) _arquivos.Log("VINCULO entrega: " + res);
        }
        catch (Exception ex) { _arquivos.Log("VINCULO nao conseguiu entregar: " + ex.Message); }
    }

    /// <summary>Converte o arquivo do pedido e entrega: "converter" -&gt; <c>__moffice.entregarConversao(id, json)</c>; "comparar"/"combinar" -&gt;
    /// <c>onDocumentCompare</c>/<c>onDocumentMerge</c> no quadro (em erro com base64 "": o proprio editor mostra o aviso dele).</summary>
    private async Task ConverterEEntregarAsync(SessaoDocumento origem, string param)
    {
        await Task.Yield();
        PedidoConversao? p = null;
        try { p = JsonSerializer.Deserialize<PedidoConversao>(param); } catch (JsonException) { }
        if (p is null || string.IsNullOrEmpty(p.id)) { _arquivos.Log("CONVERSAO pedido ilegivel"); return; }
        ResultadoConversao? r = null;
        IReadOnlyDictionary<string, string> imagens = new Dictionary<string, string>();
        // Vinculo externo ainda nao respondido: pergunta ANTES de qualquer E/S no caminho. Nao -> o motor recebe "nao converteu", sem x2t.
        var negado = p.confirmar && !await ConfirmarVinculoAsync(origem, p.caminho);
        if (negado) _arquivos.Log("CONVERSAO nao feita: o usuario nao permitiu o vinculo " + p.caminho);
        else try
        {
            r = await Conversao.ConverterAsync(p.caminho, p.formato, origem.PastaRecovery, p.id);
            if (r.Bin.Length > MaximoBytesLeitura)
            {
                _arquivos.Log($"CONVERSAO {p.uso} {Path.GetFileName(p.caminho)}: {r.Bin.Length} bytes, acima do teto");
                r = null;
                _ = Dispatcher.BeginInvoke(() => { if (!_fechada) AvisarUsuario(MensagemConversaoGrandeDemais); });
                throw new OperationCanceledException("grande demais");
            }
            // Midia do convertido -> media/ do documento aberto, com nome novo (ver Midia.ImportarMidiaConvertidaAsync). O vinculo/XML de planilha
            // (8194) so' le' o binario (get()): nao copia nada.
            if (p.formato != PonteDesktop.CanvasPlanilha && r.Imagens.Count > 0)
                imagens = await Midia.ImportarMidiaConvertidaAsync(origem, Path.Combine(ConversaoParaEditor.PastaDaConversao(origem.PastaRecovery, p.id), Midia.PastaMedia));
            _arquivos.Log($"CONVERSAO {p.uso} {Path.GetFileName(p.caminho)} -> {r.Bin.Length} bytes, {imagens.Count} midia(s)");
        }
        catch (OperationCanceledException) { r = null; }
        catch (Exception ex)
        {
            r = null;
            _arquivos.Log($"CONVERSAO {p.uso} {Path.GetFileName(p.caminho)} falhou: {ex.Message}");
        }
        if (!ReferenceEquals(origem, _sessao) || _fechada) { ConversaoParaEditor.Liberar(origem.PastaRecovery, p.id); return; }
        string js;
        if (p.uso == "converter")
        {
            var json = r is null ? """{"ok":false}"""
                : JsonSerializer.Serialize(new { ok = true, base64 = Convert.ToBase64String(r.Bin), tamanho = r.Bin.Length, pasta = r.Pasta, imagens });
            if (r is null) ConversaoParaEditor.Liberar(origem.PastaRecovery, p.id);
            js = $$"""
                (function () {
                  try {
                    var w = window.frames[0];
                    if (!w || !w.__moffice || typeof w.__moffice.entregarConversao !== "function") return "sem ponte no quadro";
                    return "ok " + w.__moffice.entregarConversao({{JsonSerializer.Serialize(p.id)}}, {{JsonSerializer.Serialize(json)}});
                  } catch (e) { return "erro: " + e; }
                })()
                """;
        }
        else
        {
            var funcao = p.uso == "combinar" ? "onDocumentMerge" : "onDocumentCompare";
            var argumentos = r is null
                ? "\"\", \"\", 0, {}"
                : string.Join(", ", JsonSerializer.Serialize(r.Pasta), JsonSerializer.Serialize(Convert.ToBase64String(r.Bin)), r.Bin.Length, JsonSerializer.Serialize(imagens));
            // A pasta da conversao nao e' mais lida depois daqui (o binario vai no script e a midia ja' esta' em media/).
            ConversaoParaEditor.Liberar(origem.PastaRecovery, p.id);
            js = $$"""
                (function () {
                  try {
                    var w = window.frames[0];
                    if (!w || typeof w.{{funcao}} !== "function") return "sem {{funcao}}";
                    w.{{funcao}}({{argumentos}});
                    return "ok";
                  } catch (e) { return "erro: " + e; }
                })()
                """;
        }
        var entregue = false;
        try
        {
            var res = (await ExecutarJsAsync(js)).Trim('"');
            // "converter": "ok ok" = o callback do motor recebeu o objeto (a pasta sai no close() dele). Qualquer outra resposta: ninguem vai chamar
            // close() — a pasta sai aqui.
            entregue = p.uso != "converter" || (r is not null && res == "ok ok");
            if (!res.StartsWith("ok", StringComparison.Ordinal)) _arquivos.Log("CONVERSAO entrega: " + res);
        }
        catch (Exception ex) { _arquivos.Log("CONVERSAO nao conseguiu entregar: " + ex.Message); }
        finally
        {
            if (!entregue) ConversaoParaEditor.Liberar(origem.PastaRecovery, p.id);
            Interlocked.Increment(ref _conversoesTerminadas);
        }
    }

    private sealed record PedidoLeitura(string id, string caminho);

    /// <summary>Le' o arquivo do pedido (so' leitura, compartilhada) e entrega o base64 por <c>__moffice.entregarLeitura</c> ("" em erro).</summary>
    private async Task LerEEntregarAsync(SessaoDocumento origem, string param)
    {
        await Task.Yield();
        PedidoLeitura? p = null;
        try { p = JsonSerializer.Deserialize<PedidoLeitura>(param); } catch (JsonException) { }
        if (p is null || string.IsNullOrEmpty(p.id)) { _arquivos.Log("LEITURA pedido ilegivel"); return; }
        var b64 = "";
        try
        {
            var info = new FileInfo(p.caminho);
            if (info.Length > MaximoBytesLeitura) _arquivos.Log($"LEITURA {info.Name} grande demais ({info.Length} bytes)");
            else
            {
                await using var fs = new FileStream(p.caminho, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                using var ms = new MemoryStream();
                await fs.CopyToAsync(ms);
                b64 = Convert.ToBase64String(ms.ToArray());
                _arquivos.Log($"LEITURA {info.Name}: {ms.Length} bytes");
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { _arquivos.Log("LEITURA falhou: " + ex.Message); }
        if (!ReferenceEquals(origem, _sessao) || _fechada) return;
        var js = $$"""
            (function () {
              try {
                var w = window.frames[0];
                if (!w || !w.__moffice || typeof w.__moffice.entregarLeitura !== "function") return "sem ponte no quadro";
                return "ok " + w.__moffice.entregarLeitura({{JsonSerializer.Serialize(p.id)}}, {{JsonSerializer.Serialize(b64)}});
              } catch (e) { return "erro: " + e; }
            })()
            """;
        try
        {
            var res = (await ExecutarJsAsync(js)).Trim('"');
            if (!res.StartsWith("ok", StringComparison.Ordinal)) _arquivos.Log("LEITURA entrega: " + res);
        }
        catch (Exception ex) { _arquivos.Log("LEITURA nao conseguiu entregar: " + ex.Message); }
    }

    /// <summary><c>LocalFileRecents</c> -&gt; <c>window.onupdaterecents(lista)</c> no quadro do editor (Desktop.js do web-apps).</summary>
    private async Task EntregarRecentesAsync()
    {
        await Task.Yield();
        string json;
        try { json = RecentesParaOEditor.Json(FonteDeRecentes?.Invoke() ?? []); }
        catch (Exception ex) { _arquivos.Log("RECENTES fonte falhou: " + ex.Message); json = "[]"; }
        var js = $$"""
            (function () {
              try {
                var w = window.frames[0];
                if (!w || typeof w.onupdaterecents !== "function") return "sem onupdaterecents";
                w.onupdaterecents({{json}});
                return "ok";
              } catch (e) { return "erro: " + e; }
            })()
            """;
        try
        {
            var r = (await ExecutarJsAsync(js)).Trim('"');
            if (!r.StartsWith("ok", StringComparison.Ordinal)) _arquivos.Log("RECENTES entrega: " + r);
        }
        catch (Exception ex) { _arquivos.Log("RECENTES nao conseguiu entregar: " + ex.Message); }
    }
}
