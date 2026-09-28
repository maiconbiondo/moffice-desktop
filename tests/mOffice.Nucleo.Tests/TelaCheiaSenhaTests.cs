using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using mOffice.Conversor;
using mOffice.Conversor.Tests;
using mOffice.Nucleo;
using mOffice.Nucleo.Recuperacao;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>Task 7 do Plano 3 — as pecas puras da senha: leitura do XML do editor, metadados da sessao e o que
/// NUNCA pode carregar a senha (sessao.json, SessaoJson da ponte, log).</summary>
public class TelaCheiaSenhaTests
{
    [Theory] // a senha sai do XML do editor, decodificada; sem a tag ou XML quebrado = null
    [InlineData("<m_sPassword>abc</m_sPassword>", "abc")]
    [InlineData("<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding><m_sPassword>a&amp;b &lt;c&gt;</m_sPassword>", "a&b <c>")]
    [InlineData("<m_sPassword>&apos;&quot;x&quot;&apos;</m_sPassword>", "'\"x\"'")]   // os 5 escapes do AscCommon.lyd
    [InlineData("<m_sPassword></m_sPassword>", "")]
    [InlineData("<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>", null)]
    [InlineData("<m_sPassword>abc", null)]
    [InlineData("<!DOCTYPE x [<!ENTITY e \"y\">]><m_sPassword>&e;</m_sPassword>", null)]   // DTD recusado
    [InlineData("", null)]
    [InlineData(null, null)]
    public void Senha_DoXml(string? xml, string? esperado) => Assert.Equal(esperado, OpcoesAvancadas.Senha(xml));

    private static string Pasta() { var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }

    [Fact] // sessao.json de antes da Task 7 (sem "Protegido") continua legivel e vale false; o novo vai e volta
    public void Metadados_Protegido_CompativelComArquivoAntigo()
    {
        var p = Pasta();
        File.WriteAllText(Path.Combine(p, MetadadosSessao.NomeArquivo),
            """{"Origem":"C:\\x\\a.docx","Formato":65,"Titulo":"a.docx","Novo":false,"App":"mtexto","IdUsuario":3}""");
        var antigo = MetadadosSessao.Ler(p);
        Assert.NotNull(antigo);
        Assert.False(antigo!.Protegido);
        Assert.Equal(3, antigo.IdUsuario);

        (antigo with { Protegido = true }).Gravar(p);
        Assert.True(MetadadosSessao.Ler(p)!.Protegido);
        Assert.Contains("\"Protegido\":true", File.ReadAllText(Path.Combine(p, MetadadosSessao.NomeArquivo)));
    }

    [Fact] // a senha da sessao nunca volta para a pagina: SessaoJson (lido pelo JS) nao a contem
    public void SessaoJson_NaoLevaASenha()
    {
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\a.docx", Formato = FormatoDocumento.Docx, PastaRecovery = Pasta() };
        s.Senha = "s3gr3d0-que-nao-vaza";
        var json = new PonteDesktop(s).SessaoJson();
        Assert.Contains("a.docx", json);   // controle positivo: e' o JSON da sessao mesmo
        Assert.DoesNotContain("s3gr3d0", json);
        // E a ponte (objeto COM exposto a' pagina) nao tem membro publico nenhum que fale de senha.
        Assert.DoesNotContain(typeof(PonteDesktop).GetMembers(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance),
            m => m.Name.Contains("Senha", StringComparison.OrdinalIgnoreCase) || m.Name.Contains("Password", StringComparison.OrdinalIgnoreCase));
    }

    [Theory] // aviso da faixa ao salvar documento protegido num formato sem senha (fix round 1)
    [InlineData(FormatoDocumento.Pdf, "O arquivo foi salvo sem senha: o formato PDF não guarda proteção.")]
    [InlineData(FormatoDocumento.PdfA, "O arquivo foi salvo sem senha: o formato PDF não guarda proteção.")]
    public void MensagemSalvoSemSenha_Texto(FormatoDocumento f, string esperado)
    {
        Assert.False(X2tConversor.FormatoGuardaSenha(f));
        Assert.Equal(esperado, JanelaEditor.MensagemSalvoSemSenha(f));
    }

    [Fact] // os dois avisos que levam a senha estao na lista do que o log omite
    public void AvisosSigilosos_CobremOsAvisosComSenha()
    {
        Assert.Contains("opcoesAvancadas", JanelaEditor.AvisosSigilosos);
        Assert.Contains("senhaSalvar", JanelaEditor.AvisosSigilosos);
        // E a ponte manda a senha SO' por esses dois (varredura do JS: todo avisar(...) que recebe `senha`/`x` do
        // SetAdvancedOptions tem de usar um tipo sigiloso).
        var js = File.ReadAllText(Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo", "Ponte", "ponte-desktop.js"));
        Assert.Contains("avisar(\"opcoesAvancadas\", String(x || \"\"))", js);
        Assert.Contains("avisar(\"senhaSalvar\", typeof senha === \"string\" ? senha : \"\")", js);
        Assert.DoesNotContain("JSON.stringify({ param: param || \"\", fileType: fileType || 0, jsonOptions: jsonOptions || \"\", senha", js);
    }

    [SkippableFact] // o fim de TODA demonstracao chama AscDesktopEditor.endReporter() sem testar (sdkjs/slide/sdk-all-min.js,
                    // `DAf`): as funcoes do modo apresentador existem na ponte. Guarda propria porque o PonteCompletaTests
                    // nao varre os sdk-all-min.js (foi assim que a falta passou: Esc travava a janela em tela cheia).
    public void Ponte_TemAsFuncoesDoModoApresentador()
    {
        var raiz = FronteiraTests.RaizRepo();
        var min = Path.Combine(raiz, "third_party", "onlyoffice", "editors", "sdkjs", "slide", "sdk-all-min.js");
        Skip.If(!File.Exists(min), "motor ONLYOFFICE nao obtido");
        var bundle = File.ReadAllText(min);
        var js = File.ReadAllText(Path.Combine(raiz, "src", "mOffice.Nucleo", "Ponte", "ponte-desktop.js"));
        foreach (var nome in new[] { "endReporter", "startReporter", "sendToReporter", "sendFromReporter" })
        {
            Assert.Contains("AscDesktopEditor." + nome + "(", bundle);   // controle positivo: o motor chama
            Assert.Matches(@"(?m)^\s+" + nome + @"\s*:\s*function", js);
        }
    }

    [Fact] // o grampo (opt-in) mascara os argumentos de senha antes de logar
    public void Grampo_MascaraArgumentosDeSenha()
    {
        var js = File.ReadAllText(Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo", "Ponte", "grampo.js"));
        Assert.Contains("SetAdvancedOptions: [0], LocalFileSave: [1, 5]", js);
        // As duas linhas de log do grampo passam pelo mascaramento (nenhuma usa descrever(args) cru).
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(js, @"descrever\(mascarar\(String\(prop\), args\)\)").Count);
        Assert.DoesNotContain("descrever(args))", js);
    }
}

/// <summary>Task 7 do Plano 3 na janela real: tela cheia da apresentacao e abrir/salvar arquivo protegido por
/// senha pelo dialogo do PROPRIO editor.</summary>
[Collection(JanelaReal.Nome)]
public class TelaCheiaSenhaJanelaTests
{
    private readonly ITestOutputHelper _saida;
    public TelaCheiaSenhaJanelaTests(ITestOutputHelper saida) { _saida = saida; }

    // Senhas UNICAS por teste (fix round 1): o xunit cria uma instancia da classe por teste, e as buscas "a senha nao
    // sobrou" varrem o %TEMP% compartilhado com os outros projetos de teste rodando em paralelo.
    private readonly string Senha = $"certa-{Guid.NewGuid():N}";
    private readonly string SenhaErrada = $"errada-{Guid.NewGuid():N}";

    // ------------------------------------------------------------------------------------------------------
    // Tela cheia
    // ------------------------------------------------------------------------------------------------------

    [StructLayout(LayoutKind.Sequential)] private struct Ret { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] private struct InfoMonitor { public int Tam; public Ret Mon; public Ret Trab; public uint F; }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out Ret r);
    [DllImport("user32.dll")] private static extern IntPtr MonitorFromWindow(IntPtr h, uint f);
    [DllImport("user32.dll", EntryPoint = "GetMonitorInfoW")] private static extern bool GetMonitorInfo(IntPtr m, ref InfoMonitor i);

    /// <summary>Retangulo da janela e do monitor dela, em PIXELS, lidos direto do Windows (independente da conta
    /// em DIPs que a janela faz).</summary>
    private static (Ret Janela, Ret Monitor) RetangulosWin32(Window w)
    {
        var h = new WindowInteropHelper(w).Handle;
        GetWindowRect(h, out var jr);
        var info = new InfoMonitor { Tam = Marshal.SizeOf<InfoMonitor>() };
        Assert.True(GetMonitorInfo(MonitorFromWindow(h, 2), ref info));
        return (jr, info.Mon);
    }

    private const string IniciarApresentacao = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w.Common || !w.Common.NotificationCenter) return "sem NotificationCenter";
            w.Common.NotificationCenter.trigger("preview:start", 0);
            return "ok";
          } catch (e) { return "erro: " + e; }
        })()
        """;

    [SkippableFact] // o editor inicia a apresentacao -> SetFullscreen(true): sem borda, do tamanho do monitor, por cima;
                    // Esc -> SetFullscreen(false): tudo de volta; perder a ativacao tira o "por cima"
    public void TelaCheia_EntraESai()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mslides");
            Window? outra = null;
            try
            {
                j.Left = 120; j.Top = 90; j.Width = 1100; j.Height = 760;
                j.Show();
                j.Activate();
                await j.NovoAsync(FormatoDocumento.Pptx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var antes = (j.Left, j.Top, j.Width, j.Height, j.WindowState, j.WindowStyle, j.ResizeMode, j.Topmost);
                Assert.False(j.EmTelaCheia);

                // Entrar pelo caminho REAL do editor: o web-apps inicia a demonstracao e o sdkjs chama SetFullscreen(true).
                Assert.Equal("ok", (await j.ExecutarJsAsync(IniciarApresentacao)).Trim('"'));
                for (var i = 0; i < 100 && !j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.True(j.EmTelaCheia, "o editor nao pediu tela cheia ao iniciar a apresentacao");
                Assert.Equal("true", (await j.ExecutarJsAsync("String(window.__moffice.telaCheia === true)")).Trim('"'));
                Assert.Equal(WindowStyle.None, j.WindowStyle);
                Assert.Equal(ResizeMode.NoResize, j.ResizeMode);
                Assert.Equal(WindowState.Normal, j.WindowState);
                Assert.True(j.Topmost);
                var (jan, mon) = RetangulosWin32(j);
                _saida.WriteLine($"janela {jan.L},{jan.T},{jan.R},{jan.B} monitor {mon.L},{mon.T},{mon.R},{mon.B}");
                // O retangulo da janela e' o do MONITOR inteiro (barra de tarefas inclusive), medido pelo Windows.
                Assert.Equal((mon.L, mon.T, mon.R, mon.B), (jan.L, jan.T, jan.R, jan.B));

                // Alt+Tab: outra janela ativa -> a apresentacao deixa de ficar por cima; voltar -> por cima de novo.
                outra = new Window { Width = 200, Height = 100, Left = 10, Top = 10, ShowActivated = true, Title = "outra (teste)" };
                outra.Show(); outra.Activate();
                for (var i = 0; i < 50 && j.Topmost; i++) await Task.Delay(100);
                Assert.False(j.Topmost, "em tela cheia sem ativacao a janela nao pode ficar por cima de tudo");
                outra.Close(); outra = null;
                j.Activate();
                for (var i = 0; i < 50 && !j.Topmost; i++) await Task.Delay(100);
                Assert.True(j.Topmost);

                // Sair pelo caminho do usuario: Esc no editor (o sdkjs encerra a demonstracao e chama SetFullscreen(false)).
                // Tecla REAL (CDP Input.dispatchKeyEvent, o mesmo caminho do teclado) no editor com foco: a area de texto
                // do editor recebe o Esc e o gerente da demonstracao encerra. MEDIDO: antes de a ponte ter endReporter,
                // o encerramento lancava TypeError em `DAf` e a janela ficava PRESA em tela cheia — errosJs abaixo.
                await ApertarEscAsync(j);
                for (var i = 0; i < 100 && j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.False(j.EmTelaCheia, "Esc nao encerrou a tela cheia");
                // Quem saiu foi o EDITOR (SetFullscreen(false)), nao a rede de seguranca do host.
                await Task.Delay(JanelaEditor.EsperaSaidaDaTelaCheia + TimeSpan.FromMilliseconds(300));
                Assert.DoesNotContain("saida forcada", File.ReadAllText(j.ArquivoDeLog));
                Assert.Equal("[]", (await j.ExecutarJsAsync("JSON.stringify(window.frames[0].__moffice.errosJs || [])")).Trim('"'));
                Assert.Equal("none", (await j.ExecutarJsAsync("getComputedStyle(window.frames[0].document.getElementById('pe-preview')).display")).Trim('"'));
                Assert.Equal(antes, (j.Left, j.Top, j.Width, j.Height, j.WindowState, j.WindowStyle, j.ResizeMode, j.Topmost));
            }
            finally
            {
                outra?.Close();
                if (j.EmTelaCheia) j.DefinirTelaCheia(false);   // nunca deixar a tela presa, mesmo com o teste falhando
                j.Close();
            }
        }, TimeSpan.FromSeconds(180));
    }

    /// <summary>Esc REAL (CDP Input.dispatchKeyEvent, o mesmo caminho do teclado) no editor com foco.</summary>
    private static async Task ApertarEscAsync(JanelaEditor j)
    {
        var wv = await j.ProntoAsync();
        j.Web.Focus();
        await j.ExecutarJsAsync("window.frames[0].focus()");
        foreach (var tipo in new[] { "rawKeyDown", "keyUp" })
            await wv.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", System.Text.Json.JsonSerializer.Serialize(new
            { type = tipo, key = "Escape", code = "Escape", windowsVirtualKeyCode = 27, nativeVirtualKeyCode = 27 }));
    }

    [SkippableFact] // rede de seguranca (fix round 1): o editor NAO pede SetFullscreen(false) depois do Esc (costura) ->
                    // o host restaura a janela sozinho em ~1 s e registra no log
    public void TelaCheia_EditorNaoSai_EscDoHostRestaura()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mslides");
            try
            {
                j.Left = 130; j.Top = 100; j.Width = 1000; j.Height = 700;
                j.Show(); j.Activate();
                await j.NovoAsync(FormatoDocumento.Pptx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var antes = (j.Left, j.Top, j.Width, j.Height, j.WindowState, j.WindowStyle, j.ResizeMode, j.Topmost);
                j.IgnorarSaidaDaTelaCheiaParaTeste = true;
                Assert.Equal("ok", (await j.ExecutarJsAsync(IniciarApresentacao)).Trim('"'));
                for (var i = 0; i < 100 && !j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.True(j.EmTelaCheia);
                var relogio = System.Diagnostics.Stopwatch.StartNew();
                await ApertarEscAsync(j);
                // O editor encerrou a demonstracao e pediu a saida, mas a costura a descartou: a janela continua em tela
                // cheia ate' a rede de seguranca agir (nao antes de ~1 s).
                Assert.Contains("SetFullscreen(false) ignorado (teste)", await EsperarNoLog(j, "SetFullscreen(false) ignorado (teste)"));
                Assert.True(j.EmTelaCheia);
                for (var i = 0; i < 50 && j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.False(j.EmTelaCheia, "a rede de seguranca nao tirou a janela da tela cheia");
                Assert.True(relogio.Elapsed >= JanelaEditor.EsperaSaidaDaTelaCheia - TimeSpan.FromMilliseconds(100), $"saiu cedo demais: {relogio.Elapsed}");
                Assert.Equal(antes, (j.Left, j.Top, j.Width, j.Height, j.WindowState, j.WindowStyle, j.ResizeMode, j.Topmost));
                Assert.Contains("TELA CHEIA saida forcada pelo host: Esc (editor)", File.ReadAllText(j.ArquivoDeLog));
            }
            finally
            {
                if (j.EmTelaCheia) j.DefinirTelaCheia(false);
                j.Close();
            }
        }, TimeSpan.FromSeconds(180));
    }

    private static async Task<string> EsperarNoLog(JanelaEditor j, string trecho)
    {
        var texto = "";
        for (var i = 0; i < 50 && !texto.Contains(trecho); i++)
        {
            await Task.Delay(50);
            try { texto = File.ReadAllText(j.ArquivoDeLog); } catch (IOException) { }
        }
        return texto;
    }

    [SkippableFact] // fechar (Alt+F4) em tela cheia: sai da tela cheia antes; com o fechamento CANCELADO a janela fica normal
    public void TelaCheia_FecharCancelado_NaoFicaPreso()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mslides");
            try
            {
                j.Show(); j.Activate();
                await j.NovoAsync(FormatoDocumento.Pptx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var estiloAntes = j.WindowStyle;
                // Documento modificado + pergunta respondida com Cancelar: o fechamento nao acontece.
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoSlide("mudou"))).Trim('"'));
                for (var i = 0; i < 100 && !j.Modificado; i++) await Task.Delay(100);
                Assert.True(j.Modificado);
                var perguntou = false;
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.Cancelar); };
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("w.AscDesktopEditor.SetFullscreen(true); return \"ok\";"))).Trim('"'));
                for (var i = 0; i < 50 && !j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.True(j.EmTelaCheia);
                j.Close();
                for (var i = 0; i < 50 && !perguntou; i++) await Task.Delay(100);
                Assert.True(perguntou);
                Assert.False(j.EmTelaCheia, "fechar em tela cheia tem de sair dela antes da pergunta");
                Assert.Equal(estiloAntes, j.WindowStyle);
                Assert.False(j.Topmost);
            }
            finally
            {
                if (j.EmTelaCheia) j.DefinirTelaCheia(false);
                j.PerguntarAoFechar = _ => Task.FromResult(DecisaoFechar.NaoSalvar);
                j.Close();
            }
        }, TimeSpan.FromSeconds(180));
    }

    [SkippableFact] // "Mostrar modo de apresentador" (fix round 1): sem janela de apresentador, vira apresentacao normal —
                    // sem ficar preso no "Loading" — e o Esc sai pelo editor
    public void TelaCheia_ModoApresentador_ViraApresentacaoNormal()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mslides");
            try
            {
                j.Show(); j.Activate();
                await j.NovoAsync(FormatoDocumento.Pptx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                // O mesmo evento do item "Mostrar modo de apresentador" do web-apps (onPreviewItemClick case 2 ->
                // onPreview(0, true) -> trigger("preview:start", 0, true)).
                Assert.Equal("ok", (await j.ExecutarJsAsync("(function(){ window.frames[0].Common.NotificationCenter.trigger('preview:start', 0, true); return 'ok'; })()")).Trim('"'));
                for (var i = 0; i < 100 && !j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.True(j.EmTelaCheia);
                // A espera pelo apresentador ("Loading", div dem_id_wait_reporter) some: a apresentacao comecou.
                Assert.Equal("false", await EditorReal.Esperar(j, "String(!!window.frames[0].document.getElementById('dem_id_wait_reporter'))", "false", 50));
                Assert.Contains("SEM SUPORTE startReporter", await EsperarNoLog(j, "SEM SUPORTE startReporter"));
                await ApertarEscAsync(j);
                for (var i = 0; i < 100 && j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.False(j.EmTelaCheia, "Esc nao saiu da apresentacao iniciada pelo modo apresentador");
                await Task.Delay(JanelaEditor.EsperaSaidaDaTelaCheia + TimeSpan.FromMilliseconds(300));
                Assert.DoesNotContain("saida forcada", File.ReadAllText(j.ArquivoDeLog));   // saiu pelo editor
                Assert.Equal("[]", (await j.ExecutarJsAsync("JSON.stringify(window.frames[0].__moffice.errosJs || [])")).Trim('"'));
            }
            finally
            {
                if (j.EmTelaCheia) j.DefinirTelaCheia(false);
                j.Close();
            }
        }, TimeSpan.FromSeconds(180));
    }

    [SkippableFact] // janela MAXIMIZADA: SetFullscreen(true/false) pela ponte volta a maximizada, com o tamanho normal de antes
    public void TelaCheia_JanelaMaximizada_VoltaMaximizada()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mslides");
            try
            {
                j.Left = 150; j.Top = 110; j.Width = 900; j.Height = 650;
                j.Show();
                j.WindowState = WindowState.Maximized;
                await j.NovoAsync(FormatoDocumento.Pptx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var normalAntes = j.RestoreBounds;
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("w.AscDesktopEditor.SetFullscreen(true); return \"ok\";"))).Trim('"'));
                for (var i = 0; i < 50 && !j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.True(j.EmTelaCheia);
                Assert.Equal(WindowState.Normal, j.WindowState);
                var (jan, mon) = RetangulosWin32(j);
                Assert.Equal((mon.L, mon.T, mon.R, mon.B), (jan.L, jan.T, jan.R, jan.B));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("w.AscDesktopEditor.SetFullscreen(false); return \"ok\";"))).Trim('"'));
                for (var i = 0; i < 50 && j.EmTelaCheia; i++) await Task.Delay(100);
                Assert.False(j.EmTelaCheia);
                Assert.Equal(WindowState.Maximized, j.WindowState);
                Assert.Equal(WindowStyle.SingleBorderWindow, j.WindowStyle);
                Assert.Equal(ResizeMode.CanResize, j.ResizeMode);
                Assert.Equal(normalAntes, j.RestoreBounds);
            }
            finally
            {
                if (j.EmTelaCheia) j.DefinirTelaCheia(false);
                j.Close();
            }
        }, TimeSpan.FromSeconds(180));
    }

    // ------------------------------------------------------------------------------------------------------
    // Senha
    // ------------------------------------------------------------------------------------------------------

    /// <summary>Digita a senha NO DIALOGO do editor (Common.Views.OpenDialog do web-apps: campo em
    /// <c>#id-password-txt</c>, botao OK com <c>result="ok"</c>) e clica OK — o mesmo caminho do usuario: o
    /// dialogo chama <c>api.asc_setAdvancedOptions(DRM, …)</c>, que escapa a senha e chama SetAdvancedOptions.</summary>
    private static string DigitarSenhaNoDialogo(string senha) => """
        (function () {
          try {
            var d = window.frames[0].document;
            var dlg = d.querySelector(".open-dlg");
            if (!dlg || dlg.offsetParent === null && getComputedStyle(dlg).display === "none") return "sem dialogo";
            var campo = dlg.querySelector("#id-password-txt input");
            if (!campo) return "sem campo de senha";
            // O InputField do web-apps guarda o valor PROPRIO (getValue devolve this.value) e so' o atualiza no "blur"
            // do campo ou no Enter (MEDIDO: com o valor so' no <input> e o clique direto no OK, o editor mandou
            // "<m_sPassword></m_sPassword>"). O clique do usuario no OK tira o foco do campo antes: o blur e' esse.
            campo.focus();
            campo.value = SENHA;
            campo.dispatchEvent(new Event("input", { bubbles: true }));
            campo.dispatchEvent(new Event("blur"));
            var ok = dlg.querySelector("[result=ok]");
            if (!ok) return "sem botao ok";
            ok.click();
            return "ok";
          } catch (e) { return "erro: " + e; }
        })()
        """.Replace("SENHA", System.Text.Json.JsonSerializer.Serialize(senha));

    private const string UltimaOpcao = "String(window.__moffice && window.__moffice.ultimaOpcaoAvancada)";

    /// <summary>Abre <paramref name="arquivo"/> protegido: espera o pedido 90, digita a senha errada (se houver),
    /// espera o 91, digita a certa e espera o documento ficar pronto.</summary>
    private async Task AbrirComSenhaAsync(JanelaEditor j, string arquivo, string? errada = null)
    {
        await j.AbrirAsync(arquivo);
        Assert.Equal("90", await EditorReal.Esperar(j, UltimaOpcao, "90", 300));
        Assert.Equal(X2tConversor.CodigoProtegido, j.OpcaoAvancadaPendente);
        Assert.Equal("false", (await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"'));
        if (errada is not null)
        {
            Assert.Equal("ok", await EsperarDialogoEDigitar(j, errada));
            Assert.Equal("91", await EditorReal.Esperar(j, UltimaOpcao, "91", 300));
            Assert.Equal(X2tConversor.CodigoSenhaErrada, j.OpcaoAvancadaPendente);
            Assert.Null(j.SessaoAtual!.Senha);
        }
        Assert.Equal("ok", await EsperarDialogoEDigitar(j, Senha));
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
        Assert.Null(j.UltimoErro);
        Assert.Null(j.OpcaoAvancadaPendente);
    }

    internal static async Task<string> EsperarDialogoEDigitar(JanelaEditor j, string senha)   // internal: SalvarPdfATests (Plano 4)
    {
        var r = "";
        for (var i = 0; i < 100 && r != "ok"; i++)
        {
            r = (await j.ExecutarJsAsync(DigitarSenhaNoDialogo(senha))).Trim('"');
            if (r != "ok") await Task.Delay(100);
        }
        return r;
    }

    [SkippableFact] // abrir DOCX com senha: o editor pede (asc_initAdvancedOptions 90), a errada volta com 91, a certa
                    // abre e o texto aparece; a senha nunca vai para o log (nem com o grampo), o sessao.json nem a pagina
    public void Senha_AbreComASenhaCerta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-senha-abrir-" + Guid.NewGuid().ToString("N"))).FullName;
        var cifrado = FabricaDocumentos.DocxComSenha(pasta, "conteudo-sigiloso", Senha).GetAwaiter().GetResult();
        string log = "", recovery = "";
        var grampoAntes = JanelaEditor.GrampoLigado;
        JanelaEditor.GrampoLigado = true;   // o grampo loga TODA chamada da ponte: e' o pior caso para vazar a senha
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.Show();
                log = j.ArquivoDeLog;
                await AbrirComSenhaAsync(j, cifrado, errada: SenhaErrada);
                Assert.Contains("conteudo-sigiloso", await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor));
                Assert.Equal(Senha, j.SessaoAtual!.Senha);
                recovery = j.SessaoAtual.PastaRecovery;
                // sessao.json: marcado protegido, sem a senha.
                Assert.True(MetadadosSessao.Ler(recovery)!.Protegido);
                Assert.DoesNotContain(Senha, File.ReadAllText(Path.Combine(recovery, MetadadosSessao.NomeArquivo)));
                // A pagina nao recebe a senha de volta pela sessao (SessaoJson, cacheado em __moffice.sessao).
                var sessaoNaPagina = await j.ExecutarJsAsync("chrome.webview.hostObjects.sync.ponte.SessaoJson()")
                    + await j.ExecutarJsAsync("JSON.stringify(window.frames[0].__moffice.sessao || {})");
                Assert.Contains("protegido.docx", sessaoNaPagina);   // controle positivo: e' a sessao deste documento
                Assert.DoesNotContain(Senha, sessaoNaPagina);
                Assert.False(j.RecuperadoSemSenha);
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { JanelaEditor.GrampoLigado = grampoAntes; }
        var texto = File.ReadAllText(log);
        // A senha (certa e errada) nunca aparece no log da janela...
        Assert.DoesNotContain(Senha, texto);
        Assert.DoesNotContain(SenhaErrada, texto);
        // ...e nao porque o log ficou mudo: o fluxo passou por ele (os avisos existem, so' sem conteudo; o grampo viu
        // as chamadas, com o argumento mascarado).
        Assert.Contains("PONTE aviso opcoesAvancadas (conteudo omitido)", texto);
        Assert.Contains("SENHA documento protegido (x2t 90)", texto);
        Assert.Contains("SENHA documento protegido (x2t 91)", texto);
        Assert.Contains("ponte SetAdvancedOptions [\"<omitido>\"]", texto);
        // O XML temporario do x2t (que levava a senha em claro) nao sobrou no %TEMP%.
        Assert.DoesNotContain(Directory.GetFiles(Path.GetTempPath(), "moffice-x2t-*.xml"), f =>
        {
            try { var t = File.ReadAllText(f); return t.Contains(Senha) || t.Contains(SenhaErrada); } catch (IOException) { return false; }
        });
        Assert.False(Directory.Exists(recovery));
    }

    [SkippableFact] // salvar o documento aberto com senha mantem a protecao (reabrir sem senha da' 90); o PDF exportado
                    // e o de IMPRESSAO saem SEM senha (a faixa avisa do exportado); trocar a senha em Arquivo -> Proteger
                    // (asc_setCurrentPassword) grava com a nova
    public void Senha_SalvarMantemProtecao()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-senha-salvar-" + Guid.NewGuid().ToString("N"))).FullName;
        var cifrado = FabricaDocumentos.DocxComSenha(pasta, "conteudo-sigiloso", Senha).GetAwaiter().GetResult();
        var novo = Path.Combine(pasta, "novo.docx");
        var pdf = Path.Combine(pasta, "exportado.pdf");
        string log = "";
        var avisos = new List<string>();
        var grampoAntes = JanelaEditor.GrampoLigado;
        JanelaEditor.GrampoLigado = true;   // LocalFileSave(param, SENHA, ...) passa pelo grampo a cada salvamento
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var salvos = new List<string>();
                j.Salvo += c => salvos.Add(c);
                j.AvisoAoUsuario += a => avisos.Add(a);
                j.Show();
                log = j.ArquivoDeLog;
                await AbrirComSenhaAsync(j, cifrado);
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" editado-777"))).Trim('"'));
                Assert.True((await j.SalvarComoAsync(novo, FormatoDocumento.Docx)).Gravou);
                Assert.Null(j.UltimoErro);
                Assert.Empty(avisos);   // DOCX guarda a senha: nada a avisar
                Assert.True((await j.SalvarComoAsync(pdf, FormatoDocumento.Pdf)).Exportou);
                // PDF nao guarda senha (fix round 1): a faixa avisa, uma vez, com o texto exato.
                Assert.Equal(["O arquivo foi salvo sem senha: o formato PDF não guarda proteção."], avisos);
                // PDF de IMPRESSAO (Task 4) de documento protegido: nunca cifrado (a impressora nao tem como pedir senha).
                var pdfImpressao = await j.GerarPdfDeImpressaoAsync(mOffice.Nucleo.Impressao.OpcoesImpressao.Padrao);
                var bytesImpressao = File.ReadAllBytes(pdfImpressao);
                Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(bytesImpressao, 0, 4));
                Assert.DoesNotContain("/Encrypt", System.Text.Encoding.Latin1.GetString(bytesImpressao));

                // Arquivo -> Proteger -> trocar a senha: o editor chama asc_Save sozinho e o LocalFileSave traz a NOVA.
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_setCurrentPassword(\"nova-senha-1\"); return \"ok\";"))).Trim('"'));
                for (var i = 0; i < 100 && j.SessaoAtual!.Senha != "nova-senha-1"; i++) await Task.Delay(100);
                Assert.Equal("nova-senha-1", j.SessaoAtual!.Senha);
                // O salvamento que o editor disparou grava no arquivo atual (novo.docx, depois do "salvar como").
                for (var i = 0; i < 150 && salvos.Count < 2; i++) await Task.Delay(100);
                Assert.Equal([novo, novo], salvos);
                Assert.Null(j.UltimoErro);
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { JanelaEditor.GrampoLigado = grampoAntes; }

        var conv = new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar());
        // Reabrir sem senha: protegido (90). Com a senha antiga: errada (91). Com a nova: abre, com a edicao.
        var e90 = Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(novo, Path.Combine(pasta, "r1"))).GetAwaiter().GetResult();
        Assert.Equal(X2tConversor.CodigoProtegido, e90.CodigoSaida);
        var e91 = Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(novo, Path.Combine(pasta, "r2"), senha: Senha)).GetAwaiter().GetResult();
        Assert.Equal(X2tConversor.CodigoSenhaErrada, e91.CodigoSaida);
        conv.ParaBinAsync(novo, Path.Combine(pasta, "r3"), senha: "nova-senha-1").GetAwaiter().GetResult();
        var claro = Path.Combine(pasta, "claro.docx");
        conv.DeBinAsync(Path.Combine(pasta, "r3"), FormatoDocumento.Docx, claro).GetAwaiter().GetResult();
        Assert.Contains("editado-777", FabricaDocumentos.TextoDocx(claro));
        Assert.Contains("conteudo-sigiloso", FabricaDocumentos.TextoDocx(claro));
        // PDF exportado de documento protegido: SEM senha (fix round 1: PJe/e-SAJ recusam PDF com senha; o Word exporta
        // aberto). Ver Senha_ExportarPdf_SoCifraSeReceberSenha no conversor: o x2t cifraria se recebesse a senha.
        var bytesPdf = File.ReadAllBytes(pdf);
        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(bytesPdf, 0, 4));
        Assert.DoesNotContain("/Encrypt", System.Text.Encoding.Latin1.GetString(bytesPdf));
        Assert.Single(avisos);   // o Proteger (DOCX) depois do PDF nao avisou de novo
        var texto = File.ReadAllText(log);
        Assert.Contains("protegido por senha", texto);
        Assert.DoesNotContain(Senha, texto);
        Assert.DoesNotContain("nova-senha-1", texto);
        Assert.Contains("SENHA definida pelo editor", texto);
        Assert.Contains("PONTE aviso senhaSalvar (conteudo omitido)", texto);
        Assert.Contains("ponte LocalFileSave [\"\",\"<omitido>\"", texto);
    }

    [SkippableFact] // queda com documento protegido: a recuperacao reabre o Editor.bin JA' decifrado (sem x2t, sem pedir
                    // senha) e sinaliza RecuperadoSemSenha; definir a senha de novo em Proteger desliga o aviso
    public void Senha_RecuperacaoDeProtegido_AvisaESemPedirSenha()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-senha-queda-" + Guid.NewGuid().ToString("N"))).FullName;
        var cifrado = FabricaDocumentos.DocxComSenha(pasta, "conteudo-sigiloso", Senha).GetAwaiter().GetResult();
        var recovery = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
            await AbrirComSenhaAsync(j, cifrado);
            recovery = j.SessaoAtual!.PastaRecovery;
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" antes-da-queda"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
            await j.DescarregarChangesAsync();
            j.Close();
        }, TimeSpan.FromSeconds(240));
        var meta = MetadadosSessao.Ler(recovery);
        Assert.True(meta!.Protegido);
        foreach (var f in Directory.GetFiles(recovery, "*.json", SearchOption.AllDirectories))
            Assert.DoesNotContain(Senha, File.ReadAllText(f));

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto");
            var salvos = new List<string>();
            j.Salvo += c => salvos.Add(c);
            j.Show();
            await j.RecuperarAsync(recovery);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            // Sem pedido de senha: o Editor.bin da pasta ja' e' o documento decifrado.
            Assert.Null(j.OpcaoAvancadaPendente);
            Assert.Equal("undefined", (await j.ExecutarJsAsync(UltimaOpcao)).Trim('"'));
            Assert.Contains("antes-da-queda", await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor));
            Assert.True(j.RecuperadoSemSenha);
            Assert.Contains("Arquivo → Proteger", JanelaEditor.AvisoProtegidoRecuperado);
            // Definir a senha de novo (Arquivo -> Proteger): o editor salva e o aviso deixa de valer.
            Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_setCurrentPassword(\"de-novo-2\"); return \"ok\";"))).Trim('"'));
            for (var i = 0; i < 100 && j.RecuperadoSemSenha; i++) await Task.Delay(100);
            Assert.False(j.RecuperadoSemSenha);
            for (var i = 0; i < 150 && salvos.Count < 1; i++) await Task.Delay(100);
            Assert.Equal([cifrado], salvos);
            Assert.Null(j.UltimoErro);
            j.Close();
        }, TimeSpan.FromSeconds(240));
        // O salvamento disparado pelo Proteger gravou o arquivo de origem com a senha nova.
        var conv = new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar());
        var e = Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r1"))).GetAwaiter().GetResult();
        Assert.Equal(X2tConversor.CodigoProtegido, e.CodigoSaida);
        conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r2"), senha: "de-novo-2").GetAwaiter().GetResult();
    }

    [SkippableFact] // revisao final do Plano 3 (I1, FALHA FECHADA): sessao recuperada de documento protegido, sem senha
                    // definida de novo — "Salvar" (API do host e Ctrl+S do editor) e "Salvar como" POR CIMA do original
                    // cifrado sao RECUSADOS (arquivo byte a byte igual, EndSave de erro, faixa, documento continua
                    // modificado); "Salvar como" num arquivo NOVO grava (sem senha, com aviso); depois de Arquivo ->
                    // Proteger, o original volta a poder ser gravado — cifrado com a senha nova
    public void Senha_RecuperadoSemSenha_RecusaGravarOOriginalSemProtecao()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-senha-i1-" + Guid.NewGuid().ToString("N"))).FullName;
        var cifrado = FabricaDocumentos.DocxComSenha(pasta, "conteudo-sigiloso", Senha).GetAwaiter().GetResult();
        var copia = Path.Combine(pasta, "copia.docx");
        var copiaSemSenha = Path.Combine(pasta, "copia-como-foi-salva.docx");
        const string novaSenha = "de-novo-i1-3";
        var recovery = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
            await AbrirComSenhaAsync(j, cifrado);
            recovery = j.SessaoAtual!.PastaRecovery;
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" antes-da-queda"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
            await j.DescarregarChangesAsync();
            j.Close();
        }, TimeSpan.FromSeconds(240));
        var original = File.ReadAllBytes(cifrado);
        var avisos = new List<string>();
        var salvos = new List<string>();
        var log = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto");
            j.AvisoAoUsuario += a => avisos.Add(a);
            j.Salvo += c => salvos.Add(c);
            j.Show();
            log = j.ArquivoDeLog;
            await j.RecuperarAsync(recovery);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            Assert.True(j.RecuperadoSemSenha);
            Assert.False(j.OriginalAlteradoDepoisDaQueda);   // o original nao mudou: so' a regra da senha vale aqui
            const string recusa = JanelaEditor.MensagemProtegidoSemSenhaRecusado;

            // (1) "Salvar" pela API do host: recusado, antes de qualquer E/S no arquivo do usuario.
            var e1 = await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarAsync());
            Assert.Equal(recusa, e1.Message);
            Assert.Equal([recusa], avisos);
            Assert.True(j.Modificado);
            Assert.Equal(original, File.ReadAllBytes(cifrado));

            // (2) Ctrl+S do editor (asc_Save forcado, o mesmo do botao Salvar): mesma recusa, pelo caminho do editor.
            Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"))).Trim('"'));
            for (var i = 0; i < 100 && avisos.Count < 2; i++) await Task.Delay(100);
            Assert.Equal([recusa, recusa], avisos);
            Assert.True(j.Modificado);
            Assert.Empty(salvos);
            Assert.Equal(recusa, j.UltimoErro);
            Assert.Equal(original, File.ReadAllBytes(cifrado));

            // (2b) fix round 2: EXPORTAR PDF nao avisa "salvo sem senha" (PDF nunca leva senha; a regra dele e' outra).
            var avisosAntesDoPdf = avisos.Count;
            Assert.True((await j.SalvarComoAsync(Path.Combine(pasta, "exportado.pdf"), FormatoDocumento.Pdf)).Exportou);
            Assert.Equal(avisosAntesDoPdf, avisos.Count);
            // (3) "Salvar como" num arquivo NOVO: grava, sem senha, e a faixa avisa.
            Assert.True((await j.SalvarComoAsync(copia, FormatoDocumento.Docx)).Gravou);
            File.Copy(copia, copiaSemSenha);
            Assert.Equal(JanelaEditor.MensagemRecuperadoSalvoSemSenha, avisos[^1]);
            Assert.Equal([copia], salvos);
            // (3b) fix round 2: o aviso e' UMA vez — o Ctrl+S seguinte da copia nao repete.
            var avisosDepoisDaCopia = avisos.Count;
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" na-copia"));
            Assert.True((await j.SalvarAsync()).Gravou);
            Assert.Equal(avisosDepoisDaCopia, avisos.Count);
            Assert.Equal([copia, copia], salvos);

            // (4) "Salvar como" escolhendo o proprio original cifrado: continua recusado.
            var e4 = await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarComoAsync(cifrado, FormatoDocumento.Docx));
            Assert.Equal(recusa, e4.Message);
            Assert.Equal(original, File.ReadAllBytes(cifrado));

            // (5) Arquivo -> Proteger: a senha de novo (o editor salva sozinho no arquivo atual, a copia)...
            Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor($"api.asc_setCurrentPassword(\"{novaSenha}\"); return \"ok\";"))).Trim('"'));
            for (var i = 0; i < 100 && j.SessaoAtual!.Senha != novaSenha; i++) await Task.Delay(100);
            for (var i = 0; i < 150 && salvos.Count < 3; i++) await Task.Delay(100);
            Assert.Equal([copia, copia, copia], salvos);
            // (6) ...e o original volta a poder ser gravado, cifrado com ela.
            Assert.True((await j.SalvarComoAsync(cifrado, FormatoDocumento.Docx)).Gravou);
            Assert.Null(j.UltimoErro);
            j.Close();
        }, TimeSpan.FromSeconds(300));
        Assert.DoesNotContain(novaSenha, File.ReadAllText(log));
        Assert.Contains("PONTE salvamento RECUSADO", File.ReadAllText(log));
        var conv = new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar());
        // A copia do passo (3) saiu SEM senha (abre sem pedir) e com a edicao recuperada.
        conv.ParaBinAsync(copiaSemSenha, Path.Combine(pasta, "r-copia")).GetAwaiter().GetResult();
        var claro = Path.Combine(pasta, "claro.docx");
        conv.DeBinAsync(Path.Combine(pasta, "r-copia"), FormatoDocumento.Docx, claro).GetAwaiter().GetResult();
        Assert.Contains("antes-da-queda", FabricaDocumentos.TextoDocx(claro));
        // O original, gravado so' no passo (6): protegido (90 sem senha) e aberto pela senha NOVA, com a edicao.
        var e = Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r1"))).GetAwaiter().GetResult();
        Assert.Equal(X2tConversor.CodigoProtegido, e.CodigoSaida);
        conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r2"), senha: novaSenha).GetAwaiter().GetResult();
        var claro2 = Path.Combine(pasta, "claro2.docx");
        conv.DeBinAsync(Path.Combine(pasta, "r2"), FormatoDocumento.Docx, claro2).GetAwaiter().GetResult();
        Assert.Contains("antes-da-queda", FabricaDocumentos.TextoDocx(claro2));
        Assert.False(Directory.Exists(recovery));   // fechamento limpo depois de salvar
    }
}
