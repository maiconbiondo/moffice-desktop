using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 4, Task 2b — o menu Arquivo mostra "Salvar como" (e nao "Baixar em", o painel de download da versao web),
/// igual ao app oficial. O web-apps 9.4 decide por <c>mode.isDesktopApp &amp;&amp; mode.isOffline</c> (FileMenu:
/// <c>miDownload</c> escondido, <c>miSaveAs</c> = <c>#fm-btn-save-desktop</c>, acao "save-desktop" ->
/// <c>api.asc_DownloadAs()</c>); o <c>isOffline</c> vem de <c>api.asc_isOffline()</c>, que o ponte-desktop.js
/// sobrepoe na instancia (CreateEditorApi). O item e' clicado pelo DOM do menu, entao o caminho inteiro do web-apps
/// ate' o dialogo do host (<see cref="JanelaEditor.EscolherDestinoParaTeste"/>) e' o do produto.
/// Regressao: o Ctrl+S / botao Salvar (Toolbar.tryToSave, canSaveToFile -> asc_Save) de um arquivo existente
/// continua gravando no lugar SEM dialogo; documento NOVO continua perguntando o destino.
/// </summary>
[Collection(JanelaReal.Nome)]
public class SalvarComoDesktopTests
{
    private readonly ITestOutputHelper _saida;
    public SalvarComoDesktopTests(ITestOutputHelper saida) { _saida = saida; }

    private static string NovaPasta(string prefixo) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, prefixo + Guid.NewGuid().ToString("N"))).FullName;

    /// <summary>O objeto da aplicacao web-apps no quadro: DE (texto), SSE (planilha), PE (apresentacao).</summary>
    private const string App = "var app = w.DE || w.SSE || w.PE; if (!app || !app.getController) return 'sem app';";

    /// <summary>Abre o menu Arquivo (o FileMenu so' e' desenhado na primeira abertura) e le' o que o usuario ve'.
    /// "visivel" = no layout (offsetParent) e sem display:none — medido no DOM renderizado, nao no markup.</summary>
    private static readonly string LerMenuArquivo = EditorReal.NoEditor(App + """
            var lm = app.getController("LeftMenu");
            lm.leftMenu.showMenu("file");
            var d = w.document;
            function vis(id) {
              var e = d.getElementById(id);
              if (!e) return "ausente";
              return (e.offsetParent !== null && w.getComputedStyle(e).display !== "none") ? "visivel" : "oculto";
            }
            var sa = d.getElementById("fm-btn-save-desktop");
            var ativo = d.querySelector(".fm-btn.active");
            var modo = app.getController("Main").appOptions;
            return JSON.stringify({
              offline: api.asc_isOffline(),
              modoOffline: !!modo.isOffline,
              desktop: !!modo.isDesktopApp,
              salvarComo: vis("fm-btn-save-desktop"),
              baixarEm: vis("fm-btn-download"),
              salvarCopia: vis("fm-btn-save-copy"),
              historico: vis("fm-btn-history"),
              direitos: vis("fm-btn-rights"),
              textoSalvarComo: sa ? sa.textContent.replace(/\s+/g, " ").trim() : "",
              painelAtivo: ativo ? ativo.id : "",
              baixarCabecalho: vis("slot-hbtn-download") === "visivel" && !!d.querySelector("#slot-hbtn-download button, #slot-hbtn-download .btn-header"),
              chat: !!modo.canChat,
              historicoModo: !!modo.canUseHistory,
              forcesave: !!modo.canForcesave,
              canSaveToFile: !!modo.canSaveToFile
            });
        """);

    private static readonly string FecharMenuArquivo = EditorReal.NoEditor(App + """
            var lm = app.getController("LeftMenu");
            lm.leftMenu.menuFile.hide();
            return "ok";
        """);

    /// <summary>Clica "Salvar como" no menu Arquivo pelo DOM (o &lt;a&gt; do item, que e' onde o MenuItem liga o clique).</summary>
    private static readonly string ClicarSalvarComo = EditorReal.NoEditor(App + """
            app.getController("LeftMenu").leftMenu.showMenu("file");
            var li = w.document.getElementById("fm-btn-save-desktop");
            if (!li) return "sem item";
            var alvo = li.querySelector("a") || li;
            alvo.click();
            return "ok";
        """);

    /// <summary>O botao Salvar do web-apps (o mesmo tratador do Ctrl+S do cabecalho): <c>Toolbar.tryToSave</c> em texto e
    /// apresentacao (com canSaveToFile vira asc_Save; sem, iria para o asc_DownloadAs = "Salvar como") e
    /// <c>Toolbar.onSave</c> na planilha (sempre asc_Save).</summary>
    private static readonly string TentarSalvar = EditorReal.NoEditor(App + """
            var tb = app.getController("Toolbar");
            var salvar = tb.tryToSave || tb.onSave;   // texto/apresentacao: tryToSave; planilha: onSave (o clique do btnSave)
            if (!salvar) return "sem salvar no Toolbar";
            salvar.call(tb);
            return "ok";
        """);

    private static string Criar(string pasta, string tipo) => tipo switch
    {
        "xlsx" => FabricaDocumentos.Xlsx(pasta, "Base"),
        "pptx" => FabricaDocumentos.Pptx(pasta, "Base"),
        _ => FabricaDocumentos.Docx(pasta, "Base"),
    };

    private static async Task<bool> Aguardar(Func<bool> condicao, int decimos = 300)
    {
        for (var i = 0; i < decimos && !condicao(); i++) await Task.Delay(100);
        return condicao();
    }

    [SkippableTheory] // menu Arquivo: "Salvar como" visivel e "Baixar em" oculto; clicar chega ao dialogo do host com o
                      // formato do documento pre-selecionado (cancelar nao grava nada); o Salvar (Ctrl+S) de arquivo
                      // existente grava no lugar sem dialogo
    [InlineData("mtexto", "docx")]
    [InlineData("mplanilha", "xlsx")]
    [InlineData("mslides", "pptx")]
    public void MenuArquivo_SalvarComo_ComoNoAppOficial(string app, string tipo)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-salvarcomo-");
        var arquivo = Criar(pasta, tipo);
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela(app);
                var salvos = new List<string>();
                var exportados = new List<string>();
                var pedidos = new List<(string filtro, FormatoDocumento sugerido)>();
                j.Salvo += c => salvos.Add(c);
                j.Exportado += c => exportados.Add(c);
                j.EscolherDestinoParaTeste = (filtro, sugerido) => { pedidos.Add((filtro, sugerido)); return null; };   // cancela
                j.Show();
                await j.AbrirAsync(arquivo);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var formato = j.SessaoAtual!.Formato;

                // 1) O que o menu mostra.
                // O FileMenu e' desenhado de forma assincrona na primeira abertura (depois que os scripts do app terminam de
                // carregar): le' ate' os itens existirem.
                var menu = "";
                for (var i = 0; i < 100 && (menu == "" || menu.Contains("\"salvarComo\":\"ausente\"")); i++)
                {
                    if (i > 0) await Task.Delay(100);
                    menu = (await j.ExecutarJsAsync(LerMenuArquivo)).Trim('"').Replace("\\\"", "\"");
                }
                _saida.WriteLine($"{tipo}: {menu}");
                Assert.Contains("\"salvarComo\":\"visivel\"", menu);
                Assert.Contains("\"baixarEm\":\"oculto\"", menu);
                Assert.Contains("\"offline\":true", menu);
                Assert.Contains("\"modoOffline\":true", menu);
                Assert.Contains("\"desktop\":true", menu);
                Assert.Contains("\"salvarCopia\":\"oculto\"", menu);
                Assert.Contains("\"textosalvarcomo\":\"salvar como\"", menu.ToLowerInvariant());   // locale pt: "Salvar Como" (texto/planilha), "Salvar como" (apresentacao)
                Assert.Contains("\"canSaveToFile\":true", menu);
                Assert.Equal("ok", (await j.ExecutarJsAsync(FecharMenuArquivo)).Trim('"'));

                // 2) Editar e clicar "Salvar como": o dialogo do host e' pedido com o formato do documento pre-selecionado.
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(formato.DocType(), " salvarcomo-" + tipo))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                var antes = File.GetLastWriteTimeUtc(arquivo);
                Assert.Equal("ok", (await j.ExecutarJsAsync(ClicarSalvarComo)).Trim('"'));
                Assert.True(await Aguardar(() => pedidos.Count > 0), "o clique em \"Salvar como\" nao chegou ao dialogo do host");
                var pedido = Assert.Single(pedidos);
                Assert.Equal(formato, pedido.sugerido);
                Assert.Equal(JanelaEditor.FiltroSalvar(formato), pedido.filtro);
                await Task.Delay(1000);   // o cancelamento responde EndSave(1): nada pode ser gravado depois
                Assert.Null(j.UltimoErro);
                Assert.Empty(salvos);
                Assert.Empty(exportados);
                Assert.Equal(antes, File.GetLastWriteTimeUtc(arquivo));
                Assert.True(j.Modificado);

                // 3) Regressao: o Salvar do web-apps (Ctrl+S / botao) grava NO LUGAR, sem dialogo.
                Assert.Equal("ok", (await j.ExecutarJsAsync(TentarSalvar)).Trim('"'));
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Salvar do editor nao gravou");
                Assert.Equal([arquivo], salvos);
                Assert.Single(pedidos);   // nenhum dialogo novo
                Assert.Null(j.UltimoErro);
                Assert.NotEqual(antes, File.GetLastWriteTimeUtc(arquivo));
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            var texto = tipo switch
            {
                "xlsx" => FabricaDocumentos.TextoXlsx(arquivo),
                "pptx" => FabricaDocumentos.TextoPptx(arquivo),
                _ => FabricaDocumentos.TextoDocx(arquivo),
            };
            Assert.Contains("salvarcomo-" + tipo, texto);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Ctrl+S REAL (tecla sintetica pelo DevTools Protocol, como o <c>ColarAsync</c> do host faz com o Ctrl+V): e' o
    /// sdkjs que trata a tecla no quadro com foco, nao o teste.</summary>
    private static async Task CtrlS(JanelaEditor j)
    {
        j.Web.Focus();
        string Evento(string tipo) => System.Text.Json.JsonSerializer.Serialize(new
        {
            type = tipo, modifiers = 2, windowsVirtualKeyCode = 0x53, nativeVirtualKeyCode = 0x53, key = "s", code = "KeyS",
        });
        await j.Web.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", Evento("rawKeyDown"));
        await j.Web.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", Evento("keyUp"));
    }

    [SkippableTheory] // Ctrl+S de TECLADO num arquivo existente editado: grava no lugar, sem dialogo (regressao do isOffline),
                      // nos tres editores (texto desde a Task 2b; planilha e apresentacao na revisao final do Plano 4)
    [InlineData("mtexto", "docx")]
    [InlineData("mplanilha", "xlsx")]
    [InlineData("mslides", "pptx")]
    public void CtrlS_ArquivoExistente_GravaNoLugarSemDialogo(string app, string tipo)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-salvarcomo-ctrls-");
        var arquivo = Criar(pasta, tipo);
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela(app);
                var salvos = new List<string>();
                j.Salvo += c => salvos.Add(c);
                j.EscolherDestinoParaTeste = (_, _) => throw new InvalidOperationException("o Ctrl+S nao devia perguntar o destino");
                j.Show();
                j.Activate();
                await j.AbrirAsync(arquivo);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("true", (await j.ExecutarJsAsync(EditorReal.NoEditor("return String(api.asc_isOffline());"))).Trim('"'));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(j.SessaoAtual!.Formato.DocType(), " ctrls-" + tipo))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                await CtrlS(j);
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Ctrl+S de teclado nao gravou");
                Assert.Equal([arquivo], salvos);
                Assert.Null(j.UltimoErro);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Contains("ctrls-" + tipo, Texto(arquivo, tipo));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    private static string Texto(string arquivo, string tipo) => tipo switch
    {
        "xlsx" => FabricaDocumentos.TextoXlsx(arquivo),
        "pptx" => FabricaDocumentos.TextoPptx(arquivo),
        _ => FabricaDocumentos.TextoDocx(arquivo),
    };

    [SkippableFact] // linha "PDF assinado com mPDF" com extensao de documento digitada: grava o DOCUMENTO, nao pede assinatura e
                    // avisa na faixa; controle: a mesma linha com .pdf pede a assinatura e nao avisa (revisao final do Plano 4)
    public void LinhaAssinada_ComExtensaoDeDocumento_AvisaQueNadaFoiAssinado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-salvarcomo-assinar-");
        var docx = FabricaDocumentos.Docx(pasta, "Base");
        var odt = Path.Combine(pasta, "copia.odt");
        var pdf = Path.Combine(pasta, "copia");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var avisos = new List<string>();
                var assinar = new List<string>();
                j.AvisoAoUsuario += m => avisos.Add(m);
                j.AssinarPedido += (_, p) => assinar.Add(p);
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                // O que o dialogo real devolveria: linha 5 ("PDF assinado com mPDF") e o nome digitado com .odt.
                j.EscolherDestinoParaTeste = (_, sugerido) => JanelaEditor.DestinoDoDialogo(odt, 5, sugerido, true);
                Assert.True((await j.SalvarComoComDialogoAsync()).Gravou);
                Assert.True(File.Exists(odt));
                Assert.Empty(assinar);
                Assert.Equal([JanelaEditor.MensagemAssinarSalvouDocumento(FormatoDocumento.Odt)], avisos);
                // Controle: a mesma linha com nome sem extensao vira .pdf, pede a assinatura e nao avisa nada.
                avisos.Clear();
                j.EscolherDestinoParaTeste = (_, sugerido) => JanelaEditor.DestinoDoDialogo(pdf, 5, sugerido, true);
                Assert.True((await j.SalvarComoComDialogoAsync()).Exportou);
                Assert.Equal([pdf + ".pdf"], assinar);
                Assert.Empty(avisos);
                Assert.Null(j.UltimoErro);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(180));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // documento NOVO: o mesmo Salvar do web-apps (Ctrl+S) continua perguntando o destino ao host
    public void DocumentoNovo_Salvar_PerguntaODestino()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-salvarcomo-novo-");
        var destino = Path.Combine(pasta, "novo.docx");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var salvos = new List<string>();
                var pedidos = new List<FormatoDocumento>();
                j.Salvo += c => salvos.Add(c);
                j.EscolherDestinoParaTeste = (_, sugerido) => { pedidos.Add(sugerido); return new DestinoSalvar(destino, FormatoDocumento.Docx); };
                j.Show();
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("true", (await j.ExecutarJsAsync(EditorReal.NoEditor("return String(api.asc_isOffline());"))).Trim('"'));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("novo-salvarcomo"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(TentarSalvar)).Trim('"'));
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Salvar do documento novo nao gravou");
                Assert.Equal([FormatoDocumento.Docx], pedidos);
                Assert.Equal([destino], salvos);
                Assert.Null(j.UltimoErro);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Contains("novo-salvarcomo", FabricaDocumentos.TextoDocx(destino));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
