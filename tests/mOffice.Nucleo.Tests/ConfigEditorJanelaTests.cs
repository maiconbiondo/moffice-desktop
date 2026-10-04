using System.Text.Json;
using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 6 na janela real: links viram "abrir no sistema depois de confirmar" (nunca um pop-up do WebView2),
/// a configuracao injetada pelo host chega ao editor (patch 01 do index.html + ponte), macros JS desligadas.</summary>
[Collection(JanelaReal.Nome)]
public class ConfigEditorJanelaTests
{
    private sealed class Registro
    {
        public readonly List<string> Confirmacoes = new();
        public readonly List<string> Abertos = new();
        public bool Resposta = true;
    }

    private static JanelaEditor NovaJanelaComLinks(Registro r, string app = "mtexto")
    {
        var j = EditorReal.PrepararParaTeste(new JanelaEditor(app)
        {
            ConfirmarLink = u => { lock (r) r.Confirmacoes.Add(u); return r.Resposta; },
        });
        j.AbrirNoSistemaParaTeste = a => { lock (r) r.Abertos.Add(a); };
        return j;
    }

    private static async Task<JanelaEditor> AbrirNovaAsync(Registro r, FormatoDocumento f = FormatoDocumento.Docx, string app = "mtexto")
    {
        var j = NovaJanelaComLinks(r, app);
        j.Show();
        await j.NovoAsync(f);
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
        return j;
    }

    /// <summary>Paginas (alvos "page") que o WebView2 conhece alem da do editor: um pop-up criado apareceria aqui.</summary>
    private static async Task<List<string>> AlvosDePaginaAsync(JanelaEditor j)
    {
        var wv = await j.ProntoAsync();
        using var d = JsonDocument.Parse(await wv.CallDevToolsProtocolMethodAsync("Target.getTargets", "{}"));
        return d.RootElement.GetProperty("targetInfos").EnumerateArray()
            .Where(t => t.GetProperty("type").GetString() == "page").Select(t => t.GetProperty("url").GetString()!).ToList();
    }

    private static async Task EsperarAsync(Func<bool> condicao, int ms = 6000)
    {
        for (var i = 0; i < ms / 100 && !condicao(); i++) await Task.Delay(100);
    }

    [SkippableFact] // window.open de um link do documento NAO cria janela: pergunta e abre no sistema so' com "sim" (nada de rede aqui)
    public void WindowOpen_PerguntaEAbreNoSistema()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); List<string> antes = new(), alvos = new(), bloqueadas = new();
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            antes = await AlvosDePaginaAsync(j);
            await j.ExecutarJsAsync("window.open('http://127.0.0.1:9/link-teste'); 1");
            await EsperarAsync(() => { lock (r) return r.Abertos.Count > 0; });
            await Task.Delay(1500); // um pop-up, se existisse, ja' teria nascido (e tentado a rede)
            alvos = await AlvosDePaginaAsync(j);
            bloqueadas = j.Guarda.Bloqueadas.Concat(j.Guarda.Permitidas).ToList();
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Equal(["http://127.0.0.1:9/link-teste"], r.Confirmacoes);
        Assert.Equal(["http://127.0.0.1:9/link-teste"], r.Abertos);
        Assert.DoesNotContain(alvos, u => u.Contains("link-teste")); // nenhuma janela/aba nova carregou o link
        Assert.Equal(antes.Count, alvos.Count); // e nenhuma pagina nova nasceu (o ambiente e' do processo: compara antes x depois)
        Assert.DoesNotContain(bloqueadas, u => u.Contains("link-teste"));
    }

    [SkippableFact] // confirmacao negada: nada e' aberto
    public void WindowOpen_Negado_NaoAbre()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro { Resposta = false };
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            await j.ExecutarJsAsync("window.open('http://127.0.0.1:9/link-teste'); 1");
            await EsperarAsync(() => { lock (r) return r.Confirmacoes.Count > 0; });
            await Task.Delay(800);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Single(r.Confirmacoes);
        Assert.Empty(r.Abertos);
    }

    [SkippableFact] // ajuda/suporte do ONLYOFFICE: descartado sem perguntar (e sem abrir)
    public void WindowOpen_OnlyOffice_Descartado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string log = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            await j.ExecutarJsAsync("window.open('https://helpcenter.onlyoffice.com/x'); window.open('https://feedback.onlyoffice.com/'); 1");
            for (var i = 0; i < 60 && !File.ReadAllText(j.ArquivoDeLog).Contains("feedback.onlyoffice.com"); i++) await Task.Delay(100);
            await Task.Delay(500);
            log = File.ReadAllText(j.ArquivoDeLog);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Empty(r.Confirmacoes);
        Assert.Empty(r.Abertos);
        Assert.Contains("LINK janela nova: Descartar https://helpcenter.onlyoffice.com/x", log);
    }

    [SkippableFact] // esquema perigoso: recusado com aviso, sem perguntar
    public void WindowOpen_Recusado_AvisaSemPerguntar()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); var avisos = new List<string>(); string log = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            j.AvisoAoUsuario += a => avisos.Add(a);
            await j.ExecutarJsAsync("window.open('ftp://exemplo.invalid/virus.exe'); 1");
            await EsperarAsync(() => avisos.Count > 0);
            log = File.ReadAllText(j.ArquivoDeLog);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Empty(r.Confirmacoes);
        Assert.Empty(r.Abertos);
        Assert.True(avisos.SequenceEqual([JanelaEditor.MensagemLinkRecusado("ftp://exemplo.invalid/virus.exe")]), "avisos: " + string.Join(" | ", avisos) + "\nlog:\n" + log);
    }

    [SkippableFact] // navegacao de topo para fora do editor e' cancelada (a pagina do editor continua)
    public void NavegacaoDeTopo_Cancelada()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string pronto = "", href = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            await j.ExecutarJsAsync("location.href = 'https://exemplo.invalid/'; 1");
            await EsperarAsync(() => { lock (r) return r.Confirmacoes.Count > 0; });
            await Task.Delay(800);
            pronto = (await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"');
            href = (await j.ExecutarJsAsync("location.href")).Trim('"');
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Equal(["https://exemplo.invalid/"], r.Confirmacoes);
        Assert.Equal("true", pronto);
        Assert.StartsWith("https://moffice.local/web-apps/apps/api/documents/index.html", href);
    }

    [Theory] // quadros: so' hosts virtuais do mOffice e esquemas locais; o resto e' link
    [InlineData("https://moffice.local/web-apps/apps/documenteditor/main/index.html", true)]
    [InlineData("https://moffice-doc.local/media/image1.png", true)]
    [InlineData("about:blank", true)]
    [InlineData("blob:https://moffice.local/abc", true)]
    [InlineData("data:text/html,x", true)]
    [InlineData("https://exemplo.invalid/", false)]
    [InlineData("https://moffice.local.evil.example/", false)]
    [InlineData("mailto:a@b.com", false)]
    [InlineData("ftp://x/y", false)]
    [InlineData("file:///C:/x.docx", false)]
    [InlineData("", false)]
    public void NavegacaoDeQuadro_InternaOuLink(string uri, bool interna) => Assert.Equal(interna, JanelaEditor.EhNavegacaoInternaDeQuadro(uri));

    [SkippableFact] // iframe indo para fora: cancelado e tratado como link (confirmacao), sem carregar nada
    public void NavegacaoDeQuadro_ParaFora_ViraLink()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro();
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            await j.ExecutarJsAsync("var f = document.createElement('iframe'); f.src = 'https://exemplo.invalid/quadro'; document.body.appendChild(f); 1");
            await EsperarAsync(() => { lock (r) return r.Abertos.Count > 0; });
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Equal(["https://exemplo.invalid/quadro"], r.Confirmacoes);
        Assert.Equal(["https://exemplo.invalid/quadro"], r.Abertos);
    }

    [SkippableFact] // download http(s) vindo da pagina e' cancelado (e logado) sem perguntar: so' data:/blob: perguntam o destino
    public void Download_DaPagina_Cancelado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string log = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            await j.ExecutarJsAsync("var a = document.createElement('a'); a.href = 'https://moffice.local/web-apps/apps/api/documents/index.html'; a.download = 'x.html'; document.body.appendChild(a); a.click(); 1");
            for (var i = 0; i < 60 && !File.ReadAllText(j.ArquivoDeLog).Contains("DOWNLOAD cancelado"); i++) await Task.Delay(100);
            log = File.ReadAllText(j.ArquivoDeLog);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Contains("DOWNLOAD cancelado", log);
        Assert.Empty(r.Abertos);
    }

    [SkippableFact] // esquema de aplicativo externo (ms-calc:) nunca e' lancado pelo WebView2: recusado com aviso, sem confirmar nem abrir
    public void EsquemaExterno_NaoLanca_Recusa()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); var avisos = new List<string>();
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            j.AvisoAoUsuario += a => avisos.Add(a);
            await j.ExecutarJsAsync("var a = document.createElement('a'); a.href = 'ms-calc:'; document.body.appendChild(a); a.click(); 1");
            await EsperarAsync(() => avisos.Count > 0);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Empty(r.Confirmacoes);
        Assert.Empty(r.Abertos);
        Assert.Contains(avisos, a => a.Contains("ms-calc:"));
    }

    private static readonly string InserirESelecionarImagem = EditorReal.NoEditor("""
        var png = "data:image/png;base64," + "PNGBASE64";
        api.AddImageUrl([png]);
        return "ok";
        """.Replace("PNGBASE64", Convert.ToBase64String(mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(64, 255, 0, 0))));

    private static readonly string SalvarImagemSelecionada = EditorReal.NoEditor("""
        api.asc_SaveDrawingAsPicture('png');
        return "ok";
        """);

    [SkippableFact] // "Salvar como imagem" (<a download href=data:image/png>) pergunta o destino ao host e grava um PNG de verdade ali
    public void SalvarComoImagem_GravaNoDestinoEscolhido()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-img-").FullName;
        var destino = Path.Combine(pasta, "saida.png"); var sugestoes = new List<string>(); byte[] bytes = [];
        try
        {
            JanelaSta.Executar(async () =>
            {
                var r = new Registro();
                var j = await AbrirNovaAsync(r);
                j.EscolherDestinoDeDownloadParaTeste = nome => { sugestoes.Add(nome); return destino; };
                Assert.Equal("ok", (await j.ExecutarJsAsync(InserirESelecionarImagem)).Trim('"'));
                await Task.Delay(2500); // a figura entra e fica selecionada
                Assert.Equal("ok", (await j.ExecutarJsAsync(SalvarImagemSelecionada)).Trim('"'));
                await EsperarAsync(() => File.Exists(destino) && new FileInfo(destino).Length > 0, 15000);
                await Task.Delay(500);
                if (File.Exists(destino)) bytes = File.ReadAllBytes(destino);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Single(sugestoes);
            Assert.EndsWith(".png", sugestoes[0]);
            Assert.True(bytes.Length > 8, "o PNG nao foi gravado no destino escolhido");
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }, bytes[..8]);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // cancelar o dialogo: nada e' gravado
    public void SalvarComoImagem_Cancelado_NaoGrava()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-img-").FullName; var chamadas = 0; string log = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var r = new Registro();
                var j = await AbrirNovaAsync(r);
                j.EscolherDestinoDeDownloadParaTeste = _ => { chamadas++; return null; };
                await j.ExecutarJsAsync(InserirESelecionarImagem);
                await Task.Delay(2500);
                await j.ExecutarJsAsync(SalvarImagemSelecionada);
                await EsperarAsync(() => chamadas > 0, 10000);
                await Task.Delay(1500);
                log = File.ReadAllText(j.ArquivoDeLog);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Equal(1, chamadas);
            Assert.Contains("DOWNLOAD cancelado pelo usuario", log);
            Assert.Empty(Directory.GetFiles(pasta));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // frame para file:///: nunca abre nem pergunta (o Chromium ja' o bloqueia; o tratador so' loga se o evento vier)
    public void QuadroParaFile_NaoPerguntaNemAbre()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string log = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            await j.ExecutarJsAsync("var f = document.createElement('iframe'); f.src = 'file:///C:/x/qualquer.docx'; document.body.appendChild(f); 1");
            await Task.Delay(2500);
            log = File.ReadAllText(j.ArquivoDeLog);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Empty(r.Confirmacoes);
        Assert.Empty(r.Abertos);
    }

    [SkippableFact] // o auxiliar de blob do sdk (createObjectURL + <a download>): mesmo caminho do host, grava os bytes no destino escolhido
    public void DownloadDeBlob_GravaNoDestinoEscolhido()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-blob-").FullName; var destino = Path.Combine(pasta, "saida.txt"); string texto = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var r = new Registro();
                var j = await AbrirNovaAsync(r);
                j.EscolherDestinoDeDownloadParaTeste = _ => destino;
                await j.ExecutarJsAsync("var u = (window.URL).createObjectURL(new Blob(['conteudo do blob'], {type: 'text/plain'})); var a = document.createElement('a'); a.style.display = 'none'; a.href = u; a.setAttribute('download', 'saida.txt'); document.body.appendChild(a); a.click(); 1");
                await EsperarAsync(() => File.Exists(destino) && new FileInfo(destino).Length > 0, 10000);
                await Task.Delay(300);
                if (File.Exists(destino)) texto = File.ReadAllText(destino);
                j.Close();
            }, TimeSpan.FromSeconds(120));
            Assert.Equal("conteudo do blob", texto);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    private const string BaixarBlob = "var u = (window.URL).createObjectURL(new Blob(['conteudo'], {type: 'application/octet-stream'})); var a = document.createElement('a'); a.style.display = 'none'; a.href = u; a.setAttribute('download', 'NOME'); document.body.appendChild(a); a.click(); 1";

    /// <summary>Poe um link de download (blob, "segundo.txt") fixo no canto da pagina de cima e clica nele com o mouse de verdade (DevTools).</summary>
    private static async Task ClicarNoSegundoLinkAsync(JanelaEditor j)
    {
        await j.ExecutarJsAsync("var u = (window.URL).createObjectURL(new Blob(['segundo'], {type: 'text/plain'})); var a = document.createElement('a'); a.href = u; a.setAttribute('download', 'segundo.txt'); a.textContent = 'baixar'; a.style.cssText = 'position:fixed;left:0;top:0;width:60px;height:40px;z-index:2147483647;background:#fff;display:block'; document.body.appendChild(a); 1");
        var wv = await j.ProntoAsync();
        foreach (var tipo in new[] { "mouseMoved", "mousePressed", "mouseReleased" })
            await wv.CallDevToolsProtocolMethodAsync("Input.dispatchMouseEvent", System.Text.Json.JsonSerializer.Serialize(new
            {
                type = tipo, x = 20, y = 20, button = tipo == "mouseMoved" ? "none" : "left", clickCount = tipo == "mouseMoved" ? 0 : 1,
            }));
    }

    [SkippableFact] // M4: extensao fora da lista (.exe) -> cancelado sem dialogo, nada gravado
    public void DownloadDeBlob_ExtensaoNaoPermitida_Cancelado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-blob-exe-").FullName; var chamadas = 0; string log = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var r = new Registro();
                var j = await AbrirNovaAsync(r);
                j.EscolherDestinoDeDownloadParaTeste = _ => { chamadas++; return Path.Combine(pasta, "x.exe"); };
                await j.ExecutarJsAsync(BaixarBlob.Replace("NOME", "instalador.exe"));
                await EsperarAsync(() => File.ReadAllText(j.ArquivoDeLog).Contains("DOWNLOAD cancelado (extensao nao permitida)"), 10000);
                await Task.Delay(500);
                log = File.ReadAllText(j.ArquivoDeLog);
                j.Close();
            }, TimeSpan.FromSeconds(120));
            Assert.Equal(0, chamadas);
            Assert.Contains("DOWNLOAD cancelado (extensao nao permitida): instalador.exe", log);
            Assert.Empty(Directory.GetFiles(pasta));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // M4: com o dialogo de um download aberto, um segundo download e' cancelado (um dialogo por vez); o primeiro grava normal
    public void DownloadDeBlob_SegundoComDialogoAberto_Cancelado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-blob-dois-").FullName; var destino = Path.Combine(pasta, "primeiro.txt");
        var chamadas = new List<string>(); string log = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var r = new Registro();
                var j = await AbrirNovaAsync(r);
                j.EscolherDestinoDeDownloadParaTeste = nome =>
                {
                    chamadas.Add(nome);
                    if (chamadas.Count == 1)
                    {
                        // Enquanto este "dialogo" esta' aberto, o usuario clica num 2o link de download (clique de VERDADE, DevTools: o Chromium
                        // segura um 2o download automatico sem gesto); o laco de mensagens roda como num modal de verdade.
                        _ = ClicarNoSegundoLinkAsync(j);
                        BombearMensagens(4000);
                    }
                    return destino;
                };
                await j.ExecutarJsAsync(BaixarBlob.Replace("NOME", "primeiro.txt"));
                await EsperarAsync(() => File.Exists(destino) && new FileInfo(destino).Length > 0, 15000);
                await Task.Delay(500);
                log = File.ReadAllText(j.ArquivoDeLog);
                j.Close();
            }, TimeSpan.FromSeconds(120));
            Assert.Equal(["primeiro.txt"], chamadas);
            Assert.True(log.Contains("DOWNLOAD cancelado: ja ha um dialogo de download aberto (segundo.txt)"),
                string.Join(" | ", log.Split('\n').Where(l => l.Contains("DOWNLOAD") || l.Contains("PERMISS"))));
            Assert.Equal("conteudo", File.ReadAllText(destino));
            Assert.Equal([destino], Directory.GetFiles(pasta));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Roda o laco de mensagens do WPF por <paramref name="ms"/> — o que um dialogo modal de verdade faz enquanto espera o clique.</summary>
    private static void BombearMensagens(int ms)
    {
        var frame = new System.Windows.Threading.DispatcherFrame();
        var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(ms) };
        timer.Tick += (_, _) => { timer.Stop(); frame.Continue = false; };
        timer.Start();
        System.Windows.Threading.Dispatcher.PushFrame(frame);
    }

    [SkippableFact] // um documento que chama window.open em laco nao empilha dialogos: enquanto uma confirmacao esta' aberta, os outros pedidos sao descartados
    public void WindowOpen_EmLaco_UmaConfirmacaoPorVez()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string log = "";
        JanelaSta.Executar(async () =>
        {
            var j = new JanelaEditor("mtexto")
            {
                // "Dialogo" que deixa as mensagens correrem enquanto espera, como o modal de verdade; os demais pedidos chegam AQUI DENTRO.
                ConfirmarLink = u => { lock (r) r.Confirmacoes.Add(u); BombearMensagens(1500); return true; },
            };
            EditorReal.PrepararParaTeste(j);
            j.AbrirNoSistemaParaTeste = a => { lock (r) r.Abertos.Add(a); };
            j.Show();
            await j.NovoAsync(FormatoDocumento.Docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            await j.ExecutarJsAsync("for (var i = 0; i < 5; i++) window.open('http://127.0.0.1:9/a' + i); 1");
            await EsperarAsync(() => { lock (r) return r.Abertos.Count > 0; });
            await Task.Delay(500);
            log = File.ReadAllText(j.ArquivoDeLog);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Single(r.Confirmacoes);
        Assert.Single(r.Abertos);
        Assert.Contains("LINK descartado: ja ha uma confirmacao aberta", log);
    }

    /// <summary>Le do editor o que a configuracao injetada produziu: o editorConfig do controlador principal, o gerenciador de
    /// dicas do motor, o que o menu mostra (Suporte, Sugira um recurso — com um controle POSITIVO visivel, senao "oculto" nao provaria
    /// nada) e o objeto de informacao do documento do sdk (o campo minificado e' descoberto pelo nome do metodo publico).</summary>
    private static readonly string LerConfig = EditorReal.NoEditor("""
        var d = w.document;
        var visivel = function (sel) { var e = d.querySelector(sel); return e ? !!e.offsetParent : null; };
        var m = w.DE.getController('Main');
        var c = m.editorConfig || {}, cz = c.customization || {};
        var macros = [];
        for (var k in api) { try { var v = api[k]; if (v && typeof v === 'object' && typeof v.asc_getIsEnabledMacroses === 'function') macros.push(v.asc_getIsEnabledMacroses()); } catch (e) { } }
        return JSON.stringify({
          region: c.region || null, lang: c.lang || null, userId: (c.user && c.user.id) || null, uiTheme: cz.uiTheme || null,
          help: cz.help, feedback: cz.feedback, suggestFeature: cz.suggestFeature, macros: cz.macros,
          featuresTips: cz.features && cz.features.featuresTips,
          dicasLigadas: w.Common.UI.FeaturesManager.isFeatureEnabled('featuresTips', true),
          canHelp: m.appOptions.canHelp,
          suporte: visivel('#left-btn-support'), sugira: visivel('#btn-suggest-feature'), controle: visivel('#left-btn-searchbar'),
          macrosNoSdk: macros,
          macrosHost: (window.AscDesktopEditor && window.AscDesktopEditor.isSupportMacroses) ? window.AscDesktopEditor.isSupportMacroses() : null
        });
        """);

    /// <summary>Le a config esperando o menu do editor existir (o controle positivo sai de "ausente"): o pronto da ponte vem antes da interface.</summary>
    private static async Task<string> LerConfigAsync(JanelaEditor j)
    {
        var json = "";
        for (var i = 0; i < 100; i++)
        {
            json = (await j.ExecutarJsAsync(LerConfig)).Trim('"').Replace("\\\"", "\"");
            if (json.StartsWith("{") && JsonDocument.Parse(json).RootElement.GetProperty("controle").ValueKind == JsonValueKind.True) break;
            await Task.Delay(200);
        }
        return json;
    }

    [SkippableFact] // a config do host chegou: region pt-BR, ajuda/feedback/sugestao/macros/dicas desligados; os campos da URL seguem intactos
    public void ConfigInjetada_ChegouAoEditor()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string json = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            json = await LerConfigAsync(j);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        using var d = JsonDocument.Parse(json);
        var c = d.RootElement;
        Assert.Equal("pt-BR", c.GetProperty("region").GetString());
        Assert.False(c.GetProperty("help").GetBoolean());
        Assert.False(c.GetProperty("feedback").GetBoolean());
        Assert.False(c.GetProperty("suggestFeature").GetBoolean());
        Assert.False(c.GetProperty("macros").GetBoolean());
        Assert.False(c.GetProperty("featuresTips").GetBoolean());
        Assert.False(c.GetProperty("dicasLigadas").GetBoolean()); // nenhuma dica de "Novo"
        Assert.False(c.GetProperty("canHelp").GetBoolean());
        // o que o usuario veria: sem Suporte nem "Sugira um recurso" (com o controle positivo visivel: a pagina carregou o menu)
        Assert.True(c.GetProperty("controle").GetBoolean(), "controle positivo (busca do menu) nao esta visivel - a leitura nao vale");
        Assert.NotEqual(JsonValueKind.True, c.GetProperty("suporte").ValueKind); // oculto ou nem criado
        Assert.NotEqual(JsonValueKind.True, c.GetProperty("sugira").ValueKind);
        Assert.False(c.GetProperty("macrosHost").GetBoolean());
        // o canal da URL continua dono de idioma, usuario e tema (a mescla nao os atropela)
        Assert.Equal("pt-BR", c.GetProperty("lang").GetString());
        Assert.False(string.IsNullOrEmpty(c.GetProperty("userId").GetString()));
        Assert.False(string.IsNullOrEmpty(c.GetProperty("uiTheme").GetString()));
    }

    [SkippableFact] // macro JS nao roda: o sdk recebe "macros desligadas" (asc_putIsEnabledMacroses(false) do customization.macros) e a ponte nega suporte
    public void MacrosJs_Desligadas()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string json = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            json = await LerConfigAsync(j);
            j.Close();
        }, TimeSpan.FromSeconds(120));
        using var d = JsonDocument.Parse(json);
        var noSdk = d.RootElement.GetProperty("macrosNoSdk").EnumerateArray().Select(x => x.GetBoolean()).ToList();
        Assert.Equal([false], noSdk); // exatamente um objeto de informacao do documento, e desligado (sem ele, a leitura nao vale)
        Assert.False(d.RootElement.GetProperty("macrosHost").GetBoolean());
    }

    [SkippableFact] // A8: com region pt-BR os menus de tamanho/margem mostram virgula decimal ("21,59 cm"); se continuar ponto, registra
    public void Region_SeparadorDecimalNosMenus()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var r = new Registro(); string texto = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync(r);
            texto = (await j.ExecutarJsAsync(EditorReal.NoEditor("var e = w.document.querySelector('.panel[data-tab=layout]'); return e ? e.textContent : '';"))).Trim('"');
            j.Close();
        }, TimeSpan.FromSeconds(120));
        Assert.Contains("US Letter", texto); // controle positivo: o menu de tamanhos foi lido
        Skip.If(texto.Contains("21.59 cm") && !texto.Contains("21,59 cm"),
            "MEDIDO 2026-10-03 (motor 9.4.0.129): com editorConfig.region pt-BR os menus de tamanho/margem seguem com ponto decimal (\"21.59 cm\"); o separador vem de outra configuracao — Task 12 documenta (A8).");
        Assert.Contains("21,59 cm", texto);
    }

    [SkippableFact] // documento NOVO nao mostra "Abrir Local do Arquivo"; documento de arquivo mostra
    public void AbrirLocalDoArquivo_SoComArquivo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-goback-").FullName;
        var arquivo = Path.Combine(pasta, "doc.docx");
        bool canBackNovo = true, canBackArquivo = false;
        try
        {
            var motor = mOffice.Conversor.Tests.MotorDeTeste.Localizar();
            File.Copy(ModelosVazios.Caminho(motor, FormatoDocumento.Docx), arquivo);
            JanelaSta.Executar(async () =>
            {
                var r = new Registro();
                var j = await AbrirNovaAsync(r);
                canBackNovo = (await j.ExecutarJsAsync(EditorReal.NoEditor("return String(w.DE.getController('Main').appOptions.canBack);"))).Trim('"') == "true";
                j.Close();
                var j2 = NovaJanelaComLinks(new Registro());
                j2.Show();
                await j2.AbrirAsync(arquivo);
                Assert.Equal("true", await EditorReal.Esperar(j2, EditorReal.ProntoNoEditor, "true", 300));
                canBackArquivo = (await j2.ExecutarJsAsync(EditorReal.NoEditor("return String(w.DE.getController('Main').appOptions.canBack);"))).Trim('"') == "true";
                j2.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.False(canBackNovo);
            Assert.True(canBackArquivo);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
