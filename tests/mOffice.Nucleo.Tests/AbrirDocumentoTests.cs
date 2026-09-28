using mOffice.Nucleo;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class AbrirDocumentoTests
{
    private readonly ITestOutputHelper _saida;
    public AbrirDocumentoTests(ITestOutputHelper saida) { _saida = saida; }

    [SkippableFact] // janela real: abre um DOCX sintetico e o editor sinaliza documento pronto (asc_onDocumentContentReady) com o texto na tela
    public void AbreDocx_EditorFicaPronto_ComTexto()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-abrir-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto de abertura 2026");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(docx);

            // 0) SmartScreen desligado pela CONFIGURACAO oficial da API, nao so' pela flag de linha de comando
            //    (fix da revisao 1, achado 2) — a base dupla do aviso dispensado na Secao 8 da licenca do
            //    runtime (docs/licencas/LICENCA-WebView2-Runtime.txt). Lido da JANELA REAL, e nao so' do valor
            //    que o C# atribuiu: prova que o WebView2 aceitou.
            Assert.False(j.Web.CoreWebView2.Settings.IsReputationCheckingRequired);

            // 1) o editor saiu do "Carregando documento": asc_onDocumentContentReady disparou no quadro do editor
            //    (a ponte replica a marca em window.top para o teste poder ler do documento de cima).
            var pronto = await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300);
            Assert.Equal("true", pronto);

            // 2) o texto do documento veio pelo x2t e esta DENTRO do editor (nao no DOM: a pagina e' canvas).
            //    API confirmada no sdkjs/word/sdk-all-min.js desta versao: asc_EditSelectAll + asc_GetSelectedText
            //    (GetAllText/Get_AllText, a hipotese do brief, nao existem neste build).
            var texto = await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor);
            Assert.Contains("Texto de abertura 2026", texto);

            // 3) as fontes chegam pelo esquema nativo ascdesktop:// (e nao por file:// nem por patch no
            //    sdk-all.js). Duas provas: o proprio editor pediu fontes por esse esquema durante a abertura...
            Assert.Contains(j.Guarda.Permitidas, u => u.StartsWith("ascdesktop://fonts/", StringComparison.OrdinalIgnoreCase)
                                                   && u.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase));
            //    ...e uma requisicao nova recebe 200 com o conteudo do arquivo (e nao ERR_FAILED, como na sonda).
            await j.ExecutarJsAsync(SondarFonte);
            var fonte = await EditorReal.Esperar(j, "String(window.__fonte || '')", "200", 100);
            Assert.Equal("200", fonte);

            // 4) midia: o sdkjs prefixa file:// na pasta do EndLoad (sdk-all.js: `0 != tia.indexOf("file:")`),
            //    e o Chromium recusa subrecurso file:// vindo de uma pagina https — MEDIDO nesta janela:
            //    fetch -> "Failed to fetch" e <img> -> onerror, sem chegar ao disco. O caminho de saida ja
            //    esta montado (a pasta de recovery e' servida por https://moffice-doc.local); o que falta e'
            //    so trocar AscCommon.Ss.tia depois do EndLoad, na Task 6, quando houver documento com imagem.
            //    Aqui provamos a metade que e' desta task: o host virtual do documento responde.
            var doc = await EditorReal.Esperar(j, RecoveryPorHttps, "200", 50);
            Assert.Equal("200", doc);

            // 5) fronteira da ponte (achado 1 da revisao 2026-09-17, e achado da revisao 2026-09-17 fix review 2):
            //    o quadro do editor roda documento de terceiros e NAO pode enxergar "arquivos"
            //    (LerBase64/GravarBase64/dialogos = I/O de arquivo arbitrario) — nem diretamente (w.chrome...),
            //    nem indiretamente via window.parent/window.top (o quadro e' same-origin com o documento de
            //    cima, entao essa rota estava aberta enquanto "arquivos" ainda ia para a janela por padrao).
            //    "log" e "ponte" (sem primitiva nenhuma de arquivo) continuam la.
            var fronteira = await j.ExecutarJsAsync(FronteiraDoQuadro);
            using (var doc2 = System.Text.Json.JsonDocument.Parse(fronteira))
            {
                var arquivosTipo = doc2.RootElement.GetProperty("arquivos").GetString();
                var parentArquivosTipo = doc2.RootElement.GetProperty("parentArquivos").GetString();
                var topArquivosTipo = doc2.RootElement.GetProperty("topArquivos").GetString();
                _saida.WriteLine($"fronteira do quadro: arquivos={arquivosTipo} parentArquivos={parentArquivosTipo} topArquivos={topArquivosTipo}");
                Assert.True(arquivosTipo == "undefined" || arquivosTipo == "throw",
                    $"'arquivos' deveria ser inacessivel no quadro do editor, mas: {arquivosTipo}");
                Assert.True(parentArquivosTipo == "undefined" || parentArquivosTipo!.StartsWith("erro:"),
                    $"'arquivos' nao deveria ser alcancavel via window.parent, mas: {parentArquivosTipo}");
                Assert.True(topArquivosTipo == "undefined" || topArquivosTipo!.StartsWith("erro:"),
                    $"'arquivos' nao deveria ser alcancavel via window.top, mas: {topArquivosTipo}");
                // "function": e' o que o proxy sincrono do WebView2 devolve para um metodo real de um objeto
                // de host COM (medido nesta janela) — prova que log.Log e ponte.SessaoJson existem e sao chamaveis.
                Assert.Equal("function", doc2.RootElement.GetProperty("log").GetString());
                Assert.Equal("function", doc2.RootElement.GetProperty("ponte").GetString());
            }

            // 6) nada saiu para fora das origens permitidas durante a abertura inteira
            Assert.DoesNotContain(j.Guarda.Bloqueadas, u => !u.Contains("localhost"));

            // 7) retencao (achado da revisao final 2026-09-17): a pasta de recovery guarda o documento
            //    INTEIRO em claro (Editor.bin + media/ + changes/) e tem de sumir quando a janela fecha — o
            //    PoC nao tem recuperacao apos queda, entao nao ha o que preservar. Controle positivo primeiro:
            //    sem afirmar que ela EXISTIA (e que o Editor.bin estava la'), "sumiu" nao provaria nada.
            var recovery = j.SessaoAtual!.PastaRecovery;
            Assert.True(Directory.Exists(recovery), "a pasta de recovery deveria existir com o documento aberto");
            Assert.True(File.Exists(j.SessaoAtual!.Bin), "o Editor.bin deveria estar na pasta de recovery");
            j.Close();
            Assert.False(Directory.Exists(recovery), $"a pasta de recovery ficou para tras: {recovery}");
        }, TimeSpan.FromMinutes(3));
    }

    // Le typeof direto no quadro (window.frames[0]): host object ausente e' "undefined"; se o proxy do
    // WebView2 lancar em vez de devolver undefined, o try converte para "throw" (o achado permite os dois —
    // o que importa e' que "arquivos" NUNCA e' um objeto usavel ali). Alem do quadro em si (w.chrome...),
    // tambem checa window.parent e window.top: o quadro e' same-origin com o documento de nivel superior
    // (window.frames[0] alcanca DesktopOfflineAppDocumentEndLoad nele — ver ChamarEndLoadAsync), entao um
    // "arquivos" exposto SOMENTE na janela (nao no quadro) ainda seria alcancavel dali por essas duas rotas
    // (achado da revisao 2026-09-17, fix review 2).
    private const string FronteiraDoQuadro = """
        (function () {
          var w = window.frames[0];
          var arquivos;
          try { arquivos = typeof w.chrome.webview.hostObjects.sync.arquivos; }
          catch (e) { arquivos = "throw"; }
          var parentArquivos;
          try { parentArquivos = typeof w.parent.chrome.webview.hostObjects.sync.arquivos; }
          catch (e) { parentArquivos = "erro:" + e.name; }
          var topArquivos;
          try { topArquivos = typeof w.top.chrome.webview.hostObjects.sync.arquivos; }
          catch (e) { topArquivos = "erro:" + e.name; }
          // Nao basta typeof !== "undefined": prova que da' pra CHAMAR um metodo de verdade em cada um.
          var logUsavel = "nao";
          try { logUsavel = typeof w.chrome.webview.hostObjects.sync.log.Log; } catch (e) { logUsavel = "erro:" + e; }
          var ponteUsavel = "nao";
          try { ponteUsavel = typeof w.chrome.webview.hostObjects.sync.ponte.SessaoJson; } catch (e) { ponteUsavel = "erro:" + e; }
          return { arquivos: arquivos, parentArquivos: parentArquivos, topArquivos: topArquivos, log: logUsavel, ponte: ponteUsavel };
        })()
        """;

    /// <summary>A pasta de recovery e' servida como https://moffice-doc.local (mapeamento feito no AbrirAsync).</summary>
    private const string RecoveryPorHttps = """
        (function () {
          if (!window.__doc) {
            window.__doc = "pedindo";
            fetch("https://moffice-doc.local/Editor.bin")
              .then(function (r) { window.__doc = String(r.status); })
              .catch(function (e) { window.__doc = "erro: " + e; });
          }
          return String(window.__doc);
        })()
        """;

    // A pasta de fontes vem do proprio objeto de host (o __moffice.sessao do documento de cima e' preenchido
    // sob demanda e pode ainda estar vazio: quem usa a ponte de verdade e' o quadro do editor).
    private const string SondarFonte = """
        (function () {
          window.__fonte = "";
          try {
            var s = JSON.parse(chrome.webview.hostObjects.sync.ponte.SessaoJson());
            fetch("ascdesktop://fonts/" + s.fontes + "/fonts_thumbnail.png")
              .then(function (r) { window.__fonte = String(r.status); })
              .catch(function (e) { window.__fonte = "erro: " + e; });
          } catch (e) { window.__fonte = "erro: " + e; }
        })()
        """;
    [SkippableFact] // janela real: girar a roda do mouse sobre o documento nao pode gerar erro de JS no editor
                    // (2026-09-25: CheckNeedWheel ausente na ponte -> TypeError -> "Ocorreu um erro" e documento travado)
    public void RodaDoMouse_NoDocumento_SemErroDeJs()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-roda-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto para rolar");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));

            // Roda do mouse de verdade para o editor: WheelEvent no elemento sob o centro da area de edicao do
            // quadro — e' la' que o sdkjs pendura o handler que chamava AscDesktopEditor.CheckNeedWheel().
            var alvo = await j.ExecutarJsAsync("""
                (function () {
                  var f = window.frames[0]; if (!f) return "sem quadro";
                  var el = f.document.elementFromPoint(f.innerWidth / 2, f.innerHeight / 2);
                  if (!el) return "sem elemento";
                  for (var i = 0; i < 3; i++)
                    el.dispatchEvent(new f.WheelEvent("wheel", { deltaY: 120, bubbles: true, cancelable: true }));
                  return el.tagName + "#" + el.id;
                })()
                """);
            _saida.WriteLine("roda disparada em " + alvo);
            Assert.DoesNotContain("sem ", alvo);
            await Task.Delay(1500);

            var erros = await j.ExecutarJsAsync("JSON.stringify((window.frames[0].__moffice && window.frames[0].__moffice.errosJs) || [])");
            Assert.Equal("\"[]\"", erros);
            j.Close();
        });
    }
}
