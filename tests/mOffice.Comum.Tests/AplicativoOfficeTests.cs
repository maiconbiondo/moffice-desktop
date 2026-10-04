using mOffice.Comum;
using mOffice.Comum.Telas;
using mOffice.Nucleo.Assinatura;

namespace mOffice.Comum.Tests;

public class AplicativoOfficeTests
{
    [Fact] // Task 4 do Plano 3: o PDF de impressao nao saiu -> a faixa diz o que aconteceu, com a causa do nucleo
    public void MensagemErroAoImprimir_PrefixoPtBr()
        => Assert.Equal("Não foi possível preparar a impressão: O editor não gravou as alterações pendentes em 10s.",
            AplicativoOffice.MensagemErroAoImprimir("O editor não gravou as alterações pendentes em 10s."));

    [Fact] // Task 5: o que a faixa diz depois do envio — sucesso com o nome da impressora; erro com a causa do nucleo
    public void MensagensDoEnvioParaAImpressora()
    {
        Assert.Equal("Enviado para Sala 2 (HP).", AplicativoOffice.MensagemEnviadoParaImpressora("Sala 2 (HP)"));
        // fix round 1: offline nao e' erro — vai, e a faixa avisa
        Assert.Equal("Enviado para Brother. O Windows indica que ela está desligada ou offline: o trabalho fica na fila até ela voltar.",
            AplicativoOffice.MensagemEnviadoParaImpressora("Brother", offline: true));
        Assert.Equal("Não foi possível imprimir: A impressora \"HP\" não foi encontrada no Windows.",
            AplicativoOffice.MensagemErroAoEnviar("A impressora \"HP\" não foi encontrada no Windows."));
    }

    [Fact] // primeiro argumento que e' arquivo aceito pelo perfil; flags e outros formatos ignorados; caminho absoluto
    public void PrimeiroArquivoAceito_FiltraEAbsolutiza()
    {
        Assert.Equal(Path.GetFullPath("a.docx"), AplicativoOffice.PrimeiroArquivoAceito(["--x", "/y", "a.docx", "b.docx"], PerfilApp.Texto));
        Assert.Null(AplicativoOffice.PrimeiroArquivoAceito(["a.xlsx"], PerfilApp.Texto));
        Assert.Null(AplicativoOffice.PrimeiroArquivoAceito([], PerfilApp.Texto));
    }

    [Fact] // so' entra nos recentes o que ESTE app reabre: PDF (exportacao) e formato de outro app ficam de fora
    public void DeveRegistrarRecente_SoOsFormatosDoApp()
    {
        Assert.True(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Texto, @"C:\p\a.docx"));
        Assert.True(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Texto, @"C:\p\a.ODT"));
        Assert.False(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Texto, @"C:\p\a.pdf"));   // exportacao
        Assert.False(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Texto, @"C:\p\a.xlsx"));  // outro app da suite
        Assert.False(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Texto, "   "));
        Assert.True(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Planilha, @"C:\p\a.ods"));
        Assert.False(AplicativoOffice.DeveRegistrarRecente(PerfilApp.Planilha, @"C:\p\a.pdf"));
    }

    [Theory] // ruido conhecido do web-apps: handshake, estado de UI, barra de titulo e dicas de Alt — nao vao para o log
    [InlineData("webapps:entry")]
    [InlineData("webapps:features")]
    [InlineData("editor:config")]
    [InlineData("editor:onready")]
    [InlineData("doc:onready")]
    [InlineData("quickaccess:changed")]
    [InlineData("title:mousemove")]
    [InlineData("title:button")]
    [InlineData("althints:show")]
    public void EhRuidoConhecido_SilenciaOQueChegaSempre(string comando)
        => Assert.True(AplicativoOffice.EhRuidoConhecido(comando));

    [Theory] // um comando NOVO (ou um que a casca deveria tratar) continua aparecendo no log
    [InlineData("create:new")]
    [InlineData("open:recent")]
    [InlineData("editor:event")]
    [InlineData("go:folder")]
    [InlineData("recent:forget")]
    [InlineData("comando:inventado")]
    [InlineData("")]
    public void EhRuidoConhecido_NaoEngoleComandoDesconhecido(string comando)
        => Assert.False(AplicativoOffice.EhRuidoConhecido(comando));

    [Theory] // Plano 6, Task 4: o tema escolhido dentro do editor vira id do motor; theme-system e inventado voltam ao do mOffice
    [InlineData("{\"name\":\"theme-white\",\"type\":\"light\"}", "theme-white")]
    [InlineData("{\"name\":\"theme-night\",\"type\":\"dark\"}", "theme-night")]
    [InlineData("{\"name\":\"theme-system\"}", null)]
    [InlineData("{\"name\":\"inventado\"}", null)]
    [InlineData("nao e json", null)]
    [InlineData("", null)]
    public void TemaDoEditor_DeUithemeChanged(string json, string? esperado)
        => Assert.Equal(esperado, AplicativoOffice.TemaDoComando(json));

    [Theory] // minor: so' tema conhecido do motor vai para a janela (config.json editado a mao nao injeta nada na URL)
    [InlineData("theme-night", "theme-night")]
    [InlineData("theme-system", null)]
    [InlineData("x&userid=1", null)]
    [InlineData(null, null)]
    public void TemaEditorValido_SoConhecidos(string? id, string? esperado)
        => Assert.Equal(esperado, AplicativoOffice.TemaEditorValido(id));

    [Fact] // uitheme:changed agora e' tratado: nao e' mais ruido silenciado
    public void UithemeChanged_NaoEhMaisSilencioso() => Assert.False(AplicativoOffice.EhRuidoConhecido("uitheme:changed"));

    [Theory] // formato do documento novo: OOXML por padrao, o OpenDocument irmao quando as Configuracoes pedem
    [InlineData("mtexto", false, mOffice.Conversor.FormatoDocumento.Docx)]
    [InlineData("mtexto", true, mOffice.Conversor.FormatoDocumento.Odt)]
    [InlineData("mplanilha", false, mOffice.Conversor.FormatoDocumento.Xlsx)]
    [InlineData("mplanilha", true, mOffice.Conversor.FormatoDocumento.Ods)]
    [InlineData("mslides", false, mOffice.Conversor.FormatoDocumento.Pptx)]
    [InlineData("mslides", true, mOffice.Conversor.FormatoDocumento.Odp)]
    public void FormatoDoNovo_PorConfiguracao(string app, bool odf, mOffice.Conversor.FormatoDocumento esperado)
    {
        var perfil = app == "mtexto" ? PerfilApp.Texto : app == "mplanilha" ? PerfilApp.Planilha : PerfilApp.Slides;
        Assert.Equal(esperado, AplicativoOffice.FormatoDoNovo(perfil, odf));
    }

    [Fact] // achado da revisao final do Plano 2: a raiz de recuperacao e' COMPARTILHADA pelos tres apps, entao
           // a varredura de subida de um apagaria a pasta de um documento ocioso aberto em outro. A spec §5
           // diz "> 30 dias"; era 1 dia, copiado da Sonda.
    public void IdadeMinimaRecovery_PeloMenos30Dias()
        => Assert.True(AplicativoOffice.IdadeMinimaRecovery >= TimeSpan.FromDays(30),
            $"varredura de recuperacao com {AplicativoOffice.IdadeMinimaRecovery.TotalDays:0.#} dia(s): a spec §5 exige 30");

    [Fact] // recuperacao que falhou: a janela sai da politica NA HORA (antes de fechar sozinha) e a mensagem so'
           // vai para a boas-vindas no Fechado (a trava ja' solta: o item volta a ser listado). A mensagem
           // termina com a saida do laco (Descartar / abrir o original). Fechar sem falha nao mostra nada.
    public void AvisoFalhaRecuperacao_SoltaNaFalhaEMostraNoFechado()
    {
        var eventos = new List<string>();
        var aviso = new AplicativoOffice.AvisoFalhaRecuperacao(() => eventos.Add("soltar"), m => eventos.Add("mostrar:" + m));
        aviso.AoFalhar("Não foi possível reaplicar as alterações recuperadas de \"a.docx\". O documento continua guardado para recuperação.");
        Assert.Equal(["soltar"], eventos);
        aviso.AoFalhar("segunda falha"); // idempotente, como o FalharRecuperacao do nucleo
        Assert.Equal(["soltar"], eventos);
        aviso.AoFechar();
        Assert.Equal(2, eventos.Count);
        Assert.StartsWith("mostrar:Não foi possível reaplicar as alterações recuperadas de \"a.docx\".", eventos[1]);
        Assert.EndsWith("Se o problema continuar, use Descartar ou abra o arquivo original.", eventos[1]);

        var semFalha = new List<string>();
        var normal = new AplicativoOffice.AvisoFalhaRecuperacao(() => semFalha.Add("soltar"), m => semFalha.Add(m));
        normal.AoFechar();
        Assert.Empty(semFalha);
    }

    [Fact] // fix round 3: a falha ao GRAVAR as alteracoes de uma sessao viva vai para a boas-vindas SEM o "use Descartar"
           // da recuperacao (a mensagem do nucleo diz "use Recuperar"); a fiacao da casca liga o evento proprio
    public void AvisoFalhaRecuperacao_FalhaDeGravacao_SemSufixoDeDescartar()
    {
        var mostrado = new List<string>();
        var aviso = new AplicativoOffice.AvisoFalhaRecuperacao(() => { }, mostrado.Add);
        var m = mOffice.Nucleo.JanelaEditor.MensagemFalhaAoGravarAlteracoes("a.docx", "mTexto", recuperavel: true);
        aviso.AoFalhar(m, sufixoDeRecuperacao: false);
        aviso.AoFechar();
        Assert.Equal([m], mostrado);
        Assert.Equal("Não foi possível gravar as alterações de \"a.docx\" no disco. O documento foi guardado em Documentos recuperados, na tela inicial do mTexto: use Recuperar.", m);
        Assert.DoesNotContain("Descartar", m);
        Assert.DoesNotContain("de novo", m);   // fix round 4: o app esta' aberto
        // Sem nada para recuperar (fix round 4): nao manda procurar em Documentos recuperados.
        var perdidas = mOffice.Nucleo.JanelaEditor.MensagemFalhaAoGravarAlteracoes("a.docx", "mTexto", recuperavel: false);
        Assert.Equal("Não foi possível gravar as últimas alterações de \"a.docx\" no disco, e elas não puderam ser guardadas para recuperação. O arquivo continua como foi salvo pela última vez.", perdidas);
        Assert.DoesNotContain("Recuperar", perdidas);
        var fonte = File.ReadAllText(Path.Combine(RaizRepo(), "src", "mOffice.Comum", "AplicativoOffice.cs"));
        Assert.Contains("j.AlteracoesNaoGravadas += (_, m) => { Logar(\"ALTERACOES \" + m); falha.AoFalhar(m, sufixoDeRecuperacao: false); };", fonte);
    }

    [Fact] // textos da recuperacao: sufixo que tira do laco; aviso da faixa com a contagem e o nome do app
    public void TextosDaRecuperacao()
    {
        Assert.Equal("Falhou. Se o problema continuar, use Descartar ou abra o arquivo original.",
            AplicativoOffice.MensagemFalhaRecuperacao("Falhou. "));
        Assert.Equal("Há 2 documento(s) não salvo(s) de uma sessão anterior. Para recuperá-los, feche o mPlanilha e abra-o de novo pelo menu Iniciar.",
            AplicativoOffice.TextoAvisoRecuperaveis(PerfilApp.Planilha, 2));
    }

    [Fact] // faixa ao ficar pronto: os avisos DESTE documento vao juntos (midias + recuperado protegido sem senha +
           // original alterado depois da queda); sem nenhum, o aviso de recuperaveis (consumido)
    public void AvisoAoFicarPronto_Prioridades()
    {
        const string senha = mOffice.Nucleo.JanelaEditor.AvisoProtegidoRecuperado;
        const string alterado = mOffice.Nucleo.JanelaEditor.AvisoOriginalAlterado;
        // Task 7 do Plano 3: sessao recuperada de documento protegido, sem senha -> aviso de definir a senha de novo,
        // sem consumir o aviso de recuperaveis (ele fica para a proxima janela que abrir).
        Assert.Equal((senha, false), AplicativoOffice.AvisoAoFicarPronto(false, null, true, false, "ha recuperaveis", []));
        Assert.Contains("Arquivo → Proteger", senha);
        // Revisao final do Plano 3: midias indisponiveis NAO escondem mais o aviso da senha (antes o primeiro vencia e
        // o usuario so' via "sem midia" — e salvava o documento sem protecao). Os dois aparecem, sem consumir.
        Assert.Equal(("sem midia. " + senha, false), AplicativoOffice.AvisoAoFicarPronto(true, "sem midia", true, false, "ha recuperaveis", []));
        // I2: original alterado depois da queda, sozinho e junto com os outros.
        Assert.Equal((alterado, false), AplicativoOffice.AvisoAoFicarPronto(false, null, false, true, "ha recuperaveis", []));
        Assert.Equal(("sem midia. " + senha + " " + alterado, false), AplicativoOffice.AvisoAoFicarPronto(true, "sem midia", true, true, null, []));
        Assert.Contains("nome novo", alterado);
        // Midias sozinhas: o erro do nucleo, sem consumir (como antes).
        Assert.Equal(("sem midia.", false), AplicativoOffice.AvisoAoFicarPronto(true, "sem midia", false, false, "ha recuperaveis", []));
        // Sem nada disso: o aviso de recuperaveis aparece e e' consumido (como antes da Task 7).
        Assert.Equal(("ha recuperaveis", true), AplicativoOffice.AvisoAoFicarPronto(false, null, false, false, "ha recuperaveis", []));
        Assert.Equal(((string?)null, true), AplicativoOffice.AvisoAoFicarPronto(false, null, false, false, null, []));
    }

    [Fact] // Plano 6, Task 2: o aviso de abertura (compatibilidade/macros) entra depois dos demais e, como eles, nao consome o de recuperaveis
    public void AvisoAoFicarPronto_UsaOAvisoDeAbertura()
    {
        var compat = mOffice.Nucleo.JanelaEditor.MensagemCompatibilidade(mOffice.Conversor.TabelaFormatos.DaExtensao(".doc")!);
        Assert.Equal((compat, false), AplicativoOffice.AvisoAoFicarPronto(false, null, false, false, "ha recuperaveis", [compat]));
        // TODAS as mensagens de abertura, na ordem de prioridade (revisao final do Plano 6: assinado E somente leitura mostrava so' a 1a)
        Assert.Equal((compat + " segundo aviso.", false), AplicativoOffice.AvisoAoFicarPronto(false, null, false, false, null, [compat, "segundo aviso"]));
        var ro = mOffice.Nucleo.JanelaEditor.MensagemSomenteLeitura(mOffice.Nucleo.Travas.EstadoAcesso.SomenteLeituraAtributo);
        Assert.Equal((ro + " " + mOffice.Nucleo.JanelaEditor.AvisoAssinaturaDigital, false),
            AplicativoOffice.AvisoAoFicarPronto(false, null, false, false, null, [ro, mOffice.Nucleo.JanelaEditor.AvisoAssinaturaDigital, " "]));
        // junto dos avisos do nucleo, depois deles
        Assert.Equal(("sem midia. " + compat, false), AplicativoOffice.AvisoAoFicarPronto(true, "sem midia", false, false, null, [compat]));
        Assert.Equal("Este documento está no formato antigo do Word (.doc). Ao salvar, o mOffice grava uma cópia em .docx; o arquivo original não é alterado.", compat);
        Assert.Equal("Este documento está no formato antigo do Excel (.xlsb). Ao salvar, o mOffice grava uma cópia em .xlsx; o arquivo original não é alterado.",
            mOffice.Nucleo.JanelaEditor.MensagemCompatibilidade(mOffice.Conversor.TabelaFormatos.DaExtensao(".xlsb")!));
        Assert.Equal("Este documento está no formato antigo do PowerPoint (.pps). Ao salvar, o mOffice grava uma cópia em .pptx; o arquivo original não é alterado.",
            mOffice.Nucleo.JanelaEditor.MensagemCompatibilidade(mOffice.Conversor.TabelaFormatos.DaExtensao(".pps")!));
    }

    [Fact] // FMT-05: "Abrir com" de um arquivo que o app nao abre nao some em silencio — o rejeitado sai com a mensagem do perfil
    public void PrimeiroArquivoAceito_FachadaDoPedido()
    {
        Assert.Equal(@"C:\x\a.doc", AplicativoOffice.PrimeiroArquivoAceito([@"C:\x\a.doc"], PerfilApp.Texto));
        Assert.Null(AplicativoOffice.PrimeiroArquivoAceito([@"C:\x\a.pdf"], PerfilApp.Texto));
    }

    [Fact] // recusa do RecuperarAsync (antes de o editor abrir) vai SEM o sufixo "use Descartar": diante de "já
           // está aberto em outra janela", mandar descartar seria mandar apagar um documento aberto e bem
    public void MensagemErroAoRecuperar_SemSufixoDeDescartar()
    {
        Assert.Equal("Este documento já está aberto em outra janela.",
            AplicativoOffice.MensagemErroAoRecuperar(new InvalidOperationException("Este documento já está aberto em outra janela.")));
        Assert.Equal("Não foi possível recuperar o documento: disco cheio",
            AplicativoOffice.MensagemErroAoRecuperar(new IOException("disco cheio")));
        Assert.DoesNotContain("Descartar", AplicativoOffice.MensagemErroAoRecuperar(new InvalidOperationException("x")));
    }

    [Fact] // revisao final do Plano 3: a subida varre os XMLs de parametros do x2t deixados no %TEMP% (M3, 1 h) e o
           // encerramento do Windows fica registrado no log da casca (C1). Nao da' para subir um Application num teste:
           // guarda estrutural sobre a fiacao (a regra de cada um tem teste proprio no Conversor/Nucleo).
    public void Subida_VarreXmlDoX2t_EEncerramentoDoWindowsELogado()
    {
        Assert.Equal(TimeSpan.FromHours(1), AplicativoOffice.IdadeMinimaParametrosX2t);
        var fonte = File.ReadAllText(Path.Combine(RaizRepo(), "src", "mOffice.Comum", "AplicativoOffice.cs"));
        var subida = fonte[fonte.IndexOf("protected override void OnStartup(", StringComparison.Ordinal)..];
        subida = subida[..subida.IndexOf("_config = Configuracoes.Padrao();", StringComparison.Ordinal)];
        Assert.Contains("X2tConversor.VarrerParametrosOrfaos(Path.GetTempPath(), IdadeMinimaParametrosX2t);", subida);
        Assert.Contains("X2tConversor.VarrerPastasTemporarias(ConfiguracaoMotor.PastaTempPadrao, IdadeMinimaParametrosX2t);", subida);
        var metodo = typeof(AplicativoOffice).GetMethod("OnSessionEnding",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        Assert.Equal(typeof(AplicativoOffice), metodo?.DeclaringType);
    }

    [Fact] // versao unica: os tres exes e o Comum carregam a mesma <Version> (src/Directory.Build.props)
    public void Versao_UnicaEmSrcDirectoryBuildProps()
    {
        var raiz = RaizRepo();
        var props = File.ReadAllText(Path.Combine(raiz, "src", "Directory.Build.props"));
        Assert.Matches(@"<Version>\d+\.\d+\.\d+</Version>", props);
        foreach (var p in new[] { "src/mTexto/mTexto.csproj", "src/mPlanilha/mPlanilha.csproj", "src/mSlides/mSlides.csproj", "src/mOffice.Comum/mOffice.Comum.csproj" })
            Assert.DoesNotContain("<Version>", File.ReadAllText(Path.Combine(raiz, p)));
        Assert.Matches(@"^\d+\.\d+\.\d+$", AplicativoOffice.Versao);
    }

    [Fact] // os tres apps: WinExe, icone proprio, referencia so' ao Comum, sem WebView2
    public void TresApps_CsprojCoerentes()
    {
        var raiz = RaizRepo();
        foreach (var (pasta, ico) in new[] { ("mTexto", "mtexto.ico"), ("mPlanilha", "mplanilha.ico"), ("mSlides", "mslides.ico") })
        {
            var csproj = File.ReadAllText(Path.Combine(raiz, "src", pasta, pasta + ".csproj"));
            Assert.Contains("<OutputType>WinExe</OutputType>", csproj);
            Assert.Contains($"<ApplicationIcon>Assets\\{ico}</ApplicationIcon>", csproj);
            Assert.Contains("mOffice.Comum.csproj", csproj);
            Assert.DoesNotContain("WebView2", csproj); Assert.DoesNotContain("mOffice.Nucleo.csproj", csproj);
            Assert.True(File.Exists(Path.Combine(raiz, "src", pasta, "Assets", ico)));
        }
    }

    [SkippableFact] // segunda invocacao do exe com um arquivo entrega o caminho ao processo vivo (criterio 4, duplo clique)
    public async Task SegundoProcesso_EntregaCaminhoAoPrimeiro()
    {
        Skip.If(mOffice.Nucleo.Tests.JanelaSta.SessaoRemota, "abre janelas reais; so' em sessao local");
        var raiz = RaizRepo();
        var exe = Path.Combine(raiz, "src", "mTexto", "bin", "Debug", "net10.0-windows", "mTexto.exe");
        Skip.If(!File.Exists(exe), "mTexto nao compilado (dotnet build src/mTexto)");
        mOffice.Conversor.Tests.MotorDeTeste.Localizar(); // o mTexto abre o editor: sem motor nesta arvore, pula
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        Skip.If(System.Diagnostics.Process.GetProcessesByName("mTexto").Length > 0, "ja' ha' um mTexto aberto nesta sessao (o mutex e' por usuario, nao por pasta de dados)");
        var pastaB = Path.Combine(pasta, "b"); Directory.CreateDirectory(pastaB);
        var a = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "um"); var b = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pastaB, "dois");
        var env = new Dictionary<string, string?> { ["MOFFICE_DADOS_DIR"] = Path.Combine(pasta, "dados") };
        using var p1 = Iniciar(exe, a, env);
        // try/finally: QUALQUER asercao que falhe no meio ainda tem de matar o mTexto e a arvore de
        // processos do WebView2 dele — senao a falha de um teste deixa um app com janelas abertas na
        // maquina, e o proximo Skip.If("ja' ha' um mTexto aberto") ignora a suite inteira para sempre.
        try
        {
            await Task.Delay(8000); // editor pronto
            using var p2 = Iniciar(exe, b, env);
            Assert.True(p2.WaitForExit(10000)); // a segunda instancia encaminha e sai
            await Task.Delay(8000);
            Assert.False(p1.HasExited);
            Assert.Single(System.Diagnostics.Process.GetProcessesByName("mTexto")); // um so' mTexto vivo: o segundo encaminhou e saiu
        }
        // catch largo de PROPOSITO: este finally existe para a maquina nao ficar com um mTexto vivo. Matar
        // uma arvore de processos falha de varias formas (o processo ja' morreu, um filho do WebView2
        // sumiu no meio, Win32Exception de acesso) e NENHUMA delas pode mascarar a asercao que falhou.
        finally { try { p1.Kill(entireProcessTree: true); } catch (Exception) { } }
        var recentes = new Recentes(Path.Combine(pasta, "dados", "recentes-mtexto.json")).Listar().Select(i => i.Caminho).ToList();
        Assert.Contains(a, recentes, StringComparer.OrdinalIgnoreCase); Assert.Contains(b, recentes, StringComparer.OrdinalIgnoreCase);
        static System.Diagnostics.Process Iniciar(string exe, string arquivo, Dictionary<string, string?> env)
        {
            var psi = new System.Diagnostics.ProcessStartInfo(exe, $"\"{arquivo}\"") { UseShellExecute = false };
            foreach (var (k, v) in env) psi.Environment[k] = v;
            return System.Diagnostics.Process.Start(psi)!;
        }
    }

    // ---------------------------------------------------------------------------------------------------
    // "PDF assinado com mPDF" (Plano 4, Task 5): o PedidoDeAssinatura com detector e lancador falsos. A faixa e' o
    // FaixaAvisoViewModel de verdade (o que AvisarNaBarra escreve).
    // ---------------------------------------------------------------------------------------------------

    private sealed class LancadorFalso(Exception? falha = null) : ILancadorMpdf
    {
        public List<(MpdfInstalado Mpdf, string Pdf)> Chamadas { get; } = [];
        public void Assinar(MpdfInstalado mpdf, string pdf)
        {
            Chamadas.Add((mpdf, pdf));
            if (falha is not null) throw falha;
        }
    }

    private const string PdfDeTeste = @"C:\Processos\2026\Petição inicial.pdf";

    private static (FaixaAvisoViewModel Faixa, List<string> Log) Pedir(MpdfInstalado? mpdf, LancadorFalso lancador)
    {
        var faixa = new FaixaAvisoViewModel();
        var log = new List<string>();
        var pedido = new AplicativoOffice.PedidoDeAssinatura(() => mpdf, lancador, log.Add);
        // Pelo caminho assincrono (o que a casca usa: fora da thread da interface).
        faixa.Aviso = pedido.ProcessarAsync(PdfDeTeste).GetAwaiter().GetResult();
        return (faixa, log);
    }

    [Fact]
    public void AssinarPedido_ComMpdfNovo_LancaComOPdf()
    {
        var mpdf = new MpdfInstalado(@"C:\Program Files\mPDF\mPdf.App.exe", new Version(2, 13, 0));
        var lancador = new LancadorFalso();
        var (faixa, _) = Pedir(mpdf, lancador);
        Assert.Equal([(mpdf, PdfDeTeste)], lancador.Chamadas);
        Assert.Contains("O mPDF foi aberto para assinar", faixa.Aviso);
        Assert.Equal(@"PDF salvo em C:\Processos\2026. O mPDF foi aberto para assinar.", faixa.Aviso);
    }

    [Fact]
    public void AssinarPedido_SemMpdf_SoAvisa()
    {
        var lancador = new LancadorFalso();
        var (faixa, _) = Pedir(null, lancador);
        Assert.Empty(lancador.Chamadas);
        Assert.Contains("instale o mPDF 2.13", faixa.Aviso);
        Assert.Equal($"O PDF foi salvo em {PdfDeTeste}. Para assinar, instale o mPDF 2.13 ou mais novo e abra este PDF nele.", faixa.Aviso);
    }

    [Fact]
    public void AssinarPedido_MpdfAntigo_SoAvisa()
    {
        var lancador = new LancadorFalso();
        var (faixa, _) = Pedir(new MpdfInstalado(@"C:\Program Files\mPDF\mPdf.App.exe", new Version(2, 12, 2)), lancador);
        Assert.Empty(lancador.Chamadas);
        Assert.Contains("2.12.2", faixa.Aviso);
        Assert.Equal($"O PDF foi salvo em {PdfDeTeste}. O mPDF instalado (2.12.2) é anterior à 2.13 e não abre a assinatura automaticamente: atualize o mPDF ou abra o PDF nele e use Assinar.", faixa.Aviso);
    }

    [Fact]
    public void AssinarPedido_LancadorFalha_AvisaSemDetalheTecnico()
    {
        var mpdf = new MpdfInstalado(@"C:\Program Files\mPDF\mPdf.App.exe", new Version(2, 14, 1));
        var lancador = new LancadorFalso(new System.ComponentModel.Win32Exception(5, "Access is denied"));
        var (faixa, log) = Pedir(mpdf, lancador);
        Assert.Single(lancador.Chamadas);
        Assert.Equal("O PDF foi salvo, mas não foi possível abrir o mPDF: abra o PDF no mPDF e use Assinar.", faixa.Aviso);
        Assert.DoesNotContain("Access is denied", faixa.Aviso);
        Assert.Contains(log, l => l.Contains("Access is denied"));   // o detalhe tecnico fica no log
    }

    [Fact] // o detector que lanca (registro inacessivel) conta como "sem mPDF" — nada e' lancado, nada estoura
    public void AssinarPedido_DetectorLanca_ComoSemMpdf()
    {
        var lancador = new LancadorFalso();
        var log = new List<string>();
        var pedido = new AplicativoOffice.PedidoDeAssinatura(() => throw new System.Security.SecurityException("Requested registry access is not allowed."), lancador, log.Add);
        var aviso = pedido.Processar(PdfDeTeste);
        Assert.Empty(lancador.Chamadas);
        Assert.Contains("instale o mPDF 2.13", aviso);
        Assert.DoesNotContain("registry", aviso);
    }

    [Fact] // a fiacao da casca: producao = registro + processo de verdade; o evento vai por AssinarAsync (fora da thread
           // da interface, resultado na faixa DESTA janela). Nao da' para subir um Application num teste: guarda pela fonte.
    public void AssinarPedido_FiacaoDaCasca()
    {
        var fonte = File.ReadAllText(Path.Combine(RaizRepo(), "src", "mOffice.Comum", "AplicativoOffice.cs"));
        Assert.Contains("internal Func<MpdfInstalado?> DetectarMpdf { get; init; } = DetectorMpdf.DetectarNoWindows;", fonte);
        Assert.Contains("internal ILancadorMpdf LancadorDoMpdf { get; init; } = new LancadorMpdf();", fonte);
        Assert.Contains("var assinatura = new PedidoDeAssinatura(DetectarMpdf, LancadorDoMpdf, Logar);", fonte);
        Assert.Contains("j.AssinarPedido += (origem, pdf) => _ = AssinarAsync(assinatura, origem, pdf);", fonte);
        Assert.Contains("try { AvisarNaBarra(origem, await pedido.ProcessarAsync(pdf)); }", fonte);
    }

    internal static string RaizRepo()
    {
        var d = AppContext.BaseDirectory;
        while (!File.Exists(Path.Combine(d, "mOffice.slnx"))) d = Path.GetDirectoryName(d)!;
        return d;
    }

    [Fact] // Fix 1 (seguranca): "open:recent" so' abre o que ESTA nos recentes do app (o caminho vem da pagina)
    public void RecenteDaLista_SoCaminhosDaLista()
    {
        string[] lista = [@"C:\Processos\peticao inicial.docx", @"\\srv\juridico\contestacao.docx"];
        Assert.Equal(@"C:\Processos\peticao inicial.docx", AplicativoOffice.RecenteDaLista(@"c:\processos\PETICAO INICIAL.docx", lista));
        Assert.Equal(@"\\srv\juridico\contestacao.docx", AplicativoOffice.RecenteDaLista(@"\\srv\juridico\x\..\contestacao.docx", lista));
        Assert.Null(AplicativoOffice.RecenteDaLista(@"C:\Users\fulano\Documents\salarios.xlsx", lista));
        Assert.Null(AplicativoOffice.RecenteDaLista(@"\\atacante\share\x.docx", lista));
        Assert.Null(AplicativoOffice.RecenteDaLista("peticao inicial.docx", lista));
        Assert.Null(AplicativoOffice.RecenteDaLista("", lista));
    }

    [Fact] // "Abrir Local do Arquivo" abre o Explorer com o arquivo selecionado
    public void ComandoAbrirPasta_ExplorerSelect()
    {
        var p = AplicativoOffice.ComandoAbrirPasta(@"C:\Processos\peticao inicial.docx");
        Assert.Equal(("explorer.exe", "/select,\"C:\\Processos\\peticao inicial.docx\""), (p.FileName, p.Arguments));
    }

    [Fact] // revisao final do Plano 6: o aviso do /imprimir (sem janela) vai na frente (DefaultDesktopOnly), e e' o unico MessageBox sem dono da casca
    public void Imprimir_AvisoSemJanela_NaFrente()
    {
        Assert.Equal(System.Windows.MessageBoxOptions.DefaultDesktopOnly, AplicativoOffice.OpcoesDoAvisoSemJanela);
        var fonte = File.ReadAllText(Path.Combine(FonteSimbolosRaiz(), "src", "mOffice.Comum", "AplicativoOffice.cs"));
        Assert.Contains("avisar: m => AvisarSemJanela(m, _perfil.NomeExibicao)", fonte);
    }

    private static string FonteSimbolosRaiz()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d is not null && !File.Exists(Path.Combine(d.FullName, "mOffice.slnx"))) d = d.Parent;
        return d?.FullName ?? throw new InvalidOperationException("raiz do repositorio nao encontrada");
    }
}
