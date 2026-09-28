using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Revisao final do Plano 4 (re-revisao) — fluxos do app oficial que o mOffice nao tem e que o editor de TEXTO alcanca:
/// "Comparar"/"Combinar" da URL (<c>asc_CompareDocumentUrl</c> -> <c>asc_CompareDocumentUrl_local</c> ->
/// <c>AscDesktopEditor.DownloadFiles</c>, que abre uma "acao longa" e so' a fecha no callback: sem resposta a mascara de
/// carregamento ficava para sempre; MEDIDO: "Imagem da URL" NAO passa por DownloadFiles neste build — la' ele so' e'
/// usado no modo de criptografia) e "Comparar"/"Combinar" com documento de arquivo
/// (<c>asc_CompareDocumentFile</c>/<c>asc_MergeDocumentFile</c> -> dialogo "word"). Prova: faixa pt-BR, nenhuma acao longa
/// sobrando, nenhum erro de JS nem asc_onError critico, e a edicao continua.
/// </summary>
[Collection(JanelaReal.Nome)]
public class TextoSemSuporteTests
{
    private readonly ITestOutputHelper _saida;
    public TextoSemSuporteTests(ITestOutputHelper saida) { _saida = saida; }

    private static readonly string EscutarErros = EditorReal.NoEditor("""
            w.__errosApi = [];
            api.asc_registerCallback("asc_onError", function (id, nivel) { w.__errosApi.push(id + ":" + nivel); });
            return "ok";
        """);

    /// <summary>Acoes longas ainda abertas no web-apps (Main.stackLongActions) e se a mascara de carregamento esta' visivel.</summary>
    private static readonly string EstadoDeCarga = EditorReal.NoEditor("""
            var app = w.DE, main = app && app.getController("Main"), pilha = main && main.stackLongActions;
            var n = -1;
            if (pilha && typeof pilha.length === "function") n = pilha.length();   // Common.IrregularStack
            var mascara = [].some.call(w.document.querySelectorAll(".asc-loadmask"), function (e) { return e.offsetParent !== null; });
            return "acoes=" + n + " mascara=" + mascara + " erros=" + JSON.stringify(w.__errosApi || []) + " js=" + JSON.stringify((w.__moffice && w.__moffice.errosJs) || []);
        """);

    [SkippableFact]
    public void CompararECombinar_DaUrlEDeArquivo_AvisamENaoTravam()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-texto-semsuporte-" + Guid.NewGuid().ToString("N"))).FullName;
        var docx = FabricaDocumentos.Docx(pasta, "Base");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var avisos = new List<string>();
                j.AvisoAoUsuario += m => avisos.Add(m);
                // Comparar/combinar NAO pode abrir o dialogo de imagem do host (ele nao deixaria escolher um .docx).
                j.EscolherImagemParaTeste = () => throw new InvalidOperationException("nao devia abrir o dialogo de imagem");
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EscutarErros)).Trim('"'));

                // 1) Comparar e Combinar "da URL" (web-apps: api.asc_CompareDocumentUrl/asc_MergeDocumentUrl(url, opcoes)): no
                //    desktop local vao a asc_*DocumentUrl_local, que abre a acao longa e chama DownloadFiles.
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var o = w.AscCommonWord && w.AscCommonWord.ComparisonOptions ? new w.AscCommonWord.ComparisonOptions() : undefined;
                        api.asc_CompareDocumentUrl("https://exemplo.invalid/outro.docx", o);
                        api.asc_MergeDocumentUrl("https://exemplo.invalid/outro.docx", o);
                        return "ok";
                    """))).Trim('"'));
                for (var i = 0; i < 50 && avisos.Count < 2; i++) await Task.Delay(100);
                Assert.Equal([JanelaEditor.MensagemBaixarArquivos, JanelaEditor.MensagemBaixarArquivos], avisos);
                var carga = "";
                for (var i = 0; i < 30 && !carga.StartsWith("acoes=0 mascara=false", StringComparison.Ordinal); i++)
                {
                    await Task.Delay(100);
                    carga = (await j.ExecutarJsAsync(EstadoDeCarga)).Trim('"').Replace("\\\"", "\"");
                }
                _saida.WriteLine("depois de comparar/combinar por URL: " + carga);
                Assert.StartsWith("acoes=0 mascara=false", carga);

                // 2) Comparar e Combinar "de arquivo" (web-apps: api.asc_CompareDocumentFile/asc_MergeDocumentFile(opcoes)).
                avisos.Clear();
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var o = w.AscCommonWord && w.AscCommonWord.ComparisonOptions ? new w.AscCommonWord.ComparisonOptions() : undefined;
                        api.asc_CompareDocumentFile(o);
                        api.asc_MergeDocumentFile(o);
                        return "ok";
                    """))).Trim('"'));
                for (var i = 0; i < 50 && avisos.Count < 2; i++) await Task.Delay(100);
                Assert.Equal([JanelaEditor.MensagemCompararDocumento, JanelaEditor.MensagemCompararDocumento], avisos);

                await Task.Delay(300);
                carga = (await j.ExecutarJsAsync(EstadoDeCarga)).Trim('"').Replace("\\\"", "\"");
                _saida.WriteLine("final: " + carga);
                Assert.StartsWith("acoes=0 mascara=false", carga);
                Assert.Contains("js=[]", carga);
                Assert.DoesNotContain(":-1", carga);   // nenhum asc_onError critico

                // 3) A edicao continua.
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" depois-semsuporte"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                Assert.True((await j.SalvarAsync()).Gravou);
                Assert.Null(j.UltimoErro);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Contains("depois-semsuporte", FabricaDocumentos.TextoDocx(docx));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // nivel JS: cada funcao sem suporte avisa a casca (faixa certa) e responde o callback que o chamador espera
    public void FuncoesSemSuporte_AvisamACasca_ERespondemOCallback()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-texto-semsuporte-js-" + Guid.NewGuid().ToString("N"))).FullName;
        var docx = FabricaDocumentos.Docx(pasta, "Base");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var avisos = new List<string>();
                j.AvisoAoUsuario += m => avisos.Add(m);
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                async Task<string> Chamar(string js)
                {
                    avisos.Clear();
                    var r = (await j.ExecutarJsAsync(EditorReal.NoEditor(js))).Trim('"');
                    for (var i = 0; i < 30 && avisos.Count < 1; i++) await Task.Delay(100);
                    return r;
                }
                const string Callback = """
                        w.__cb = "nao chamado";
                        function cb(x) { w.__cb = "chamado " + JSON.stringify(x); }
                    """;
                const string LerCallback = "return String(w.__cb);";

                await Chamar(Callback + "w.AscDesktopEditor.DownloadFiles([\"https://exemplo.invalid/a.png\"], [], cb); return 'ok';");
                Assert.Equal([JanelaEditor.MensagemBaixarArquivos], avisos);
                Assert.Equal("chamado {}", (await j.ExecutarJsAsync(EditorReal.NoEditor(LerCallback))).Trim('"'));

                await Chamar(Callback + "w.AscDesktopEditor.loadLocalFile(\"C:\\\\x.csv\", cb); return 'ok';");
                Assert.Equal([JanelaEditor.MensagemImportarTexto], avisos);
                Assert.Equal("chamado null", (await j.ExecutarJsAsync(EditorReal.NoEditor(LerCallback))).Trim('"'));

                await Chamar("w.AscDesktopEditor.CompareDocumentFile(\"C:\\\\x.docx\", {}); return 'ok';");
                Assert.Equal([JanelaEditor.MensagemCompararDocumento], avisos);
                await Chamar("w.AscDesktopEditor.MergeDocumentFile(\"C:\\\\x.docx\", {}); return 'ok';");
                Assert.Equal([JanelaEditor.MensagemCompararDocumento], avisos);

                await Chamar(Callback + "w.AscDesktopEditor.OpenFilenameDialog(\"csv/txt\", false, cb); return 'ok';");
                Assert.Equal([JanelaEditor.MensagemImportarTexto], avisos);
                Assert.Equal("true", (await j.ExecutarJsAsync(EditorReal.NoEditor("return String(w.__cb === 'chamado \\\"\\\"');"))).Trim('"'));

                // Controle: uma chave desconhecida do aviso "semSuporte" nao vira faixa (so' log).
                await Chamar("w.chrome.webview.hostObjects.sync.ponte.Avisar('semSuporte', 'chave-inexistente'); return 'ok';");
                Assert.Empty(avisos);
                j.Close();
            }, TimeSpan.FromSeconds(120));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
