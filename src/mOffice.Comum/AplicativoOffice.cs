using System.ComponentModel;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using mOffice.Comum.Atualizacao;
using mOffice.Comum.Dialogos;
using mOffice.Comum.Telas;
using mOffice.Conversor;
using mOffice.Nucleo;
using mOffice.Nucleo.Assinatura;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Comum;

/// <summary>
/// A classe <c>Application</c> dos TRES aplicativos da suite. Cada exe (mTexto/mPlanilha/mSlides) e' so'
/// um <c>App.xaml</c> + tres linhas de <c>App.xaml.cs</c> que passam o <see cref="PerfilApp"/> — todo o
/// comportamento (subida, tema, instancia unica, politica de janelas, barra, dialogos, comandos do editor)
/// mora aqui, uma vez so'.
///
/// Divisao de responsabilidade: a POLITICA de janelas (quem cria, quem ativa, quando encerra) esta' em
/// <see cref="GerenteDeJanelas"/>, que nao conhece WPF e por isso e' testavel; esta classe e' a FIACAO —
/// tudo que so' existe com uma janela de verdade na tela.
/// </summary>
public abstract class AplicativoOffice : Application
{
    private readonly PerfilApp _perfil;

    private Configuracoes _config = null!;
    private TemaService _tema = null!;
    private Recentes _recentes = null!;
    private GerenteDeJanelas _janelas = null!;
    private ImpressoesEmCurso _impressoes = null!;
    private InstanciaUnica? _instancia;
    private JanelaBoasVindas? _boasVindas;
    private TelaBoasVindasViewModel? _vmBoasVindas;
    private BitmapImage? _icone;
    private bool _iconeTentado;

    // Aviso de "ha' documentos recuperaveis" para a PRIMEIRA janela de documento de uma subida que abriu um
    // arquivo direto (associacao/linha de comando, sem boas-vindas). Consumido uma vez; a boas-vindas o zera
    // (ela mesma mostra a lista).
    private string? _avisoRecuperaveis;

    // O app ja' pediu Shutdown (a ultima janela fechou, ou a boas-vindas foi fechada sem documento): daqui
    // para a frente nao se abre janela nenhuma — mostrar a boas-vindas com uma mensagem de erro depois do
    // Shutdown lanca ("nao e' possivel ... apos o encerramento").
    private bool _encerrando;

    // UM arquivo de log por PROCESSO (o nucleo ja' abre um por JANELA): aqui entram as falhas da casca —
    // instancia unica, abertura que falhou, comando do editor que ninguem trata. Criado no PRIMEIRO uso
    // (achado da revisao 1): uma SEGUNDA instancia, que so' encaminha o caminho pelo pipe e morre, nao
    // pode deixar nome de log reservado nem pasta tocada por uma execucao que nao teve nada a dizer.
    private string? _log;

    // Mutex `Global\` de PRESENCA para o instalador (Plano 5, Task 2; porte do `_updateAppMutex` do mPDF): segurado
    // pela vida do processo e solto em OnExit. Independente da instancia unica por app. Melhor esforco.
    private Mutex? _mutexDeAtualizacao;

    // "Instalar agora" das Configuracoes: o instalador verificado fica armado enquanto as janelas fecham pelo fluxo
    // normal, e e' iniciado (com reabertura) no encerramento — ver IniciarAtualizacao e EncerrarAgora.
    private readonly AtualizacaoAoEncerrar _atualizacao;

    protected AplicativoOffice(PerfilApp perfil)
    {
        _perfil = perfil;
        _atualizacao = new AtualizacaoAoEncerrar(
            // /REABRIR=<este app>: o instalador (um so para a suite) reabre ESTE app no fim (DeveReabrir do .iss).
            a => System.Diagnostics.Process.Start(InstaladorSilencioso.MontarStartInfo(a, app: _perfil.Nome)),
            ex => Logar("ATUALIZACAO nao foi possivel iniciar o instalador: " + ex),
            // A presenca sai ANTES do Process.Start (o AppMutex= do instalador nunca ve' este processo) e volta se o
            // lancamento falhar (o app continua aberto).
            soltarPresenca: () => { _mutexDeAtualizacao?.Dispose(); _mutexDeAtualizacao = null; },
            retomarPresenca: AdquirirPresenca,
            // Conferido de novo no encerramento, ja' com as janelas deste app fechadas (fix round 2, M-3).
            outrosAppsAbertos: () => OutrosAppsDaSuite.Abertos(_perfil));
    }

    /// <summary>Mutex de presenca para o instalador (AppMutex=). Melhor esforco: uma falha aqui (sem permissao para
    /// Global\) so' significa que o instalador nao espera este processo — nunca impede o app de abrir.</summary>
    private void AdquirirPresenca()
    {
        if (_mutexDeAtualizacao is not null) return;
        try { _mutexDeAtualizacao = InstaladorSilencioso.AdquirirMutexDoApp(); }
        catch (Exception ex) { Logar("ATUALIZACAO mutex de presenca indisponivel: " + ex.Message); }
    }

    /// <summary>Como achar o mPDF instalado ("PDF assinado com mPDF", Plano 4 Task 5). Padrao: o registro do Windows.
    /// Trocavel (<c>init</c>) como as demais costuras: os testes usam <see cref="PedidoDeAssinatura"/> com um falso.</summary>
    internal Func<MpdfInstalado?> DetectarMpdf { get; init; } = DetectorMpdf.DetectarNoWindows;

    /// <summary>Quem abre o mPDF com <c>/assinar</c>. Padrao: <see cref="Nucleo.Assinatura.LancadorMpdf"/> (processo de verdade).</summary>
    internal ILancadorMpdf LancadorDoMpdf { get; init; } = new LancadorMpdf();

    /// <summary>Perfil do app desta instancia (mTexto/mPlanilha/mSlides).</summary>
    protected PerfilApp Perfil => _perfil;

    /// <summary>
    /// Versao do produto, <c>Major.Minor.Build</c>. Sai do assembly do COMUM (nao de
    /// <c>GetEntryAssembly</c>, que sob <c>dotnet test</c> e' o testhost e daria a versao errada) — e os
    /// quatro assemblies carregam a mesma porque a <c>&lt;Version&gt;</c> e' declarada uma unica vez, em
    /// <c>src/Directory.Build.props</c>. Mesma normalizacao do <c>UpdateService.CurrentVersion()</c> do mPDF.
    /// </summary>
    public static string Versao
    {
        get
        {
            var v = typeof(AplicativoOffice).Assembly.GetName().Version;
            return v is null ? "0.0.0" : $"{v.Major}.{v.Minor}.{Math.Max(0, v.Build)}";
        }
    }

    /// <summary>
    /// Primeiro argumento da linha de comando que e' um ARQUIVO que este app abre: nao comeca com
    /// <c>-</c>/<c>/</c> (flag) e tem extensao aceita pelo perfil. Devolve o caminho ABSOLUTO — o Explorer
    /// manda absoluto, mas um lancamento por linha de comando pode mandar relativo, e o diretorio de
    /// trabalho do processo nao e' um lugar confiavel para resolver isto mais tarde (o mesmo motivo pelo
    /// qual <see cref="InstanciaUnica"/> recusa caminho relativo vindo do pipe). Os demais argumentos sao
    /// ignorados neste plano: abrir N arquivos de uma vez fica para depois.
    /// </summary>
    public static string? PrimeiroArquivoAceito(string[] args, PerfilApp perfil)
    {
        foreach (var a in args)
        {
            if (string.IsNullOrWhiteSpace(a)) continue;
            if (a.StartsWith('-') || a.StartsWith('/')) continue;
            if (!perfil.Aceita(a)) continue;
            try { return Path.GetFullPath(a); }
            catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException or System.Security.SecurityException) { return null; }
        }
        return null;
    }

    /// <summary>
    /// Idade minima de uma pasta de recuperacao para a varredura de subida apaga-la: <b>30 dias</b>, como a
    /// spec §5 sempre disse ("varredura de orfas > 30 dias").
    ///
    /// Era 1 dia, copiado da Sonda, e isso era um RISCO REAL de perda de documento (achado da revisao final
    /// do Plano 2): os tres apps da suite COMPARTILHAM
    /// <c>%LOCALAPPDATA%\mOffice\recuperacao</c>, entao a subida do mPlanilha varre as pastas do mTexto que
    /// ja' esta' aberto. Um documento aberto e ocioso por mais de 24 h — o fim de semana de um texto deixado
    /// na tela — teria a pasta de trabalho apagada por baixo dele, sem nenhuma janela ser avisada. Com 30
    /// dias isso deixa de ser um cenario plausivel de escritorio. (A Sonda continua em 1 dia DE PROPOSITO:
    /// e' bancada, as sessoes duram minutos e a higiene importa mais — ver <c>mOffice.Sonda/App.xaml.cs</c>.)
    ///
    /// Risco residual, registrado no gate: um documento aberto e ocioso por MAIS de 30 dias ainda perderia a
    /// pasta. O conserto definitivo e' a pasta em uso se anunciar (lock/heartbeat), o que fica para o Plano 3,
    /// junto com a recuperacao apos queda de verdade.
    /// </summary>
    internal static readonly TimeSpan IdadeMinimaRecovery = TimeSpan.FromDays(30);

    /// <summary>Idade minima de um <c>moffice-x2t-*.xml</c> orfao no <c>%TEMP%</c> para a subida apaga-lo (M3): 1 h
    /// protege uma conversao em andamento de uma instancia antiga (o limite do x2t e' 3 min).</summary>
    internal static readonly TimeSpan IdadeMinimaParametrosX2t = TimeSpan.FromHours(1);

    /// <summary>
    /// Logoff/reinicio/desligamento do Windows (WM_QUERYENDSESSION; revisao final do Plano 3, C1). O WPF encerra o app
    /// fechando as janelas SEM respeitar o <c>e.Cancel</c> do Closing, entao a pergunta "Salvar?" nunca aparece; quem
    /// garante que o trabalho nao salvo nao se perde e' o <c>Closed</c> da <see cref="JanelaEditor"/>, que mantem a
    /// pasta de recuperacao de toda sessao com alteracao pendente fechada sem a decisao do usuario. Aqui so' fica o
    /// registro (o log da casca e' o unico lugar onde se ve' depois por que o app fechou sozinho).
    /// </summary>
    protected override void OnSessionEnding(SessionEndingCancelEventArgs e)
    {
        Logar($"SESSAO do Windows encerrando ({e.ReasonSessionEnding}): {_janelas?.Abertas ?? 0} janela(s) de documento aberta(s); "
            + "as com alteracoes nao salvas ficam para recuperacao na proxima abertura");
        base.OnSessionEnding(e);
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        // (0) Verbos do RMM (/checkupdate, /update; Plano 5 Task 2, porte do mPDF): ANTES de tudo — da instancia
        // unica, das varreduras e de qualquer janela. Processo efemero: escreve UMA linha ASCII, sai com 0/10/20.
        // Task.Run OBRIGATORIO: no OnStartup o WPF ja' instalou o DispatcherSynchronizationContext, e esperar a
        // tarefa assincrona direto nele travaria (deadlock). Vale para os TRES executaveis: todos sobem por aqui.
        if (ExecutorDeAtualizacao.VerboDe(e.Args) is { } verbo)
        {
            var executor = new ExecutorDeAtualizacao();
            var status = Task.Run(() => executor.ExecutarAsync(verbo, CancellationToken.None)).GetAwaiter().GetResult();
            var linha = ExecutorDeAtualizacao.FormatarStatus(verbo, status);
            SaidaDeStatus.Escrever(linha);
            Logar("ATUALIZACAO " + linha); // trilha no log de diagnostico
            _encerrando = true;
            Shutdown(ExecutorDeAtualizacao.MapearExitCode(status.Resultado));
            return;
        }

        // (0b) /gerarcatalogo (Plano 5 Task 3): o [Run] do instalador gera o catalogo de fontes por maquina, como
        // administrador, sem janela, pelo mesmo caminho da janela do editor. Mesmo padrao dos verbos acima: antes de
        // tudo, Task.Run (deadlock do DispatcherSynchronizationContext), UMA linha ASCII e o codigo de saida.
        if (VerboCatalogo.EhVerbo(e.Args))
        {
            var (linha, codigo) = Task.Run(() => VerboCatalogo.ExecutarAsync(() => CatalogoFontes.Padrao(ConfiguracaoMotor.Localizar()))).GetAwaiter().GetResult();
            SaidaDeStatus.Escrever(linha);
            _encerrando = true;
            Shutdown(codigo);
            return;
        }

        // (1) Pastas por usuario + retencao. Igual a' subida da Sonda: recovery orfao de uma queda anterior
        // guarda o documento inteiro em claro, e os .log por janela se acumulam. Nenhuma das duas lanca.
        Pastas.Garantir(Pastas.Raiz);
        SessaoDocumento.VarrerRecoveryAntigo(SessaoDocumento.RaizRecovery, IdadeMinimaRecovery);
        Logs.Varrer(Pastas.Logs, Logs.RetencaoPadrao);
        // PDFs de impressao que uma queda deixou (Task 4 do Plano 3): cada janela apaga os seus ao fechar.
        mOffice.Nucleo.Impressao.PastaImpressao.Varrer(Pastas.Impressao, mOffice.Nucleo.Impressao.PastaImpressao.IdadeMinima);
        // XML de parametros do x2t que versoes anteriores deixavam na raiz do %TEMP% quando o processo morria no meio
        // de uma conversao (revisao final do Plano 3, M3): pode levar a senha de um documento em claro.
        X2tConversor.VarrerParametrosOrfaos(Path.GetTempPath(), IdadeMinimaParametrosX2t);
        // ...e as pastas temporarias por chamada que um processo morto deixou (fix round 2: la' dentro fica o XML de hoje).
        X2tConversor.VarrerPastasTemporarias(ConfiguracaoMotor.PastaTempPadrao, IdadeMinimaParametrosX2t);
        // Downloads de atualizacao abandonados (%TEMP%\mOffice\update-*, > 1 dia; ~1,4 GB cada). Nunca segue link.
        ServicoAtualizacao.VarrerDownloadsAntigos();

        // (2) Configuracao e tema ANTES de qualquer janela: a primeira ja' nasce no tema certo (o tema do
        // EDITOR e' `init` na JanelaEditor — nao da' para corrigir depois).
        _config = Configuracoes.Padrao();
        JanelaEditor.GrampoLigado = _config.Grampo;
        _tema = TemaService.DoApp();
        _tema.Aplicar(_config.Tema);
        _recentes = Recentes.DoApp(_perfil.Nome);

        // (3) Instancia unica ANTES de base.OnStartup e de qualquer janela: a segunda invocacao (duplo
        // clique num segundo arquivo) tem de encaminhar o caminho e MORRER sem nunca piscar uma tela.
        var caminho = PrimeiroArquivoAceito(e.Args, _perfil);
        _instancia = new InstanciaUnica(NomesDeInstancia.Mutex(_perfil.Nome), NomesDeInstancia.Pipe(_perfil.Nome));
        // O loop do pipe roda numa thread de worker: BeginInvoke traz para a thread de UI. Nao ha' corrida
        // com o `_janelas` criado no passo (4) logo abaixo — estamos NA thread do Dispatcher, entao
        // qualquer item enfileirado aqui so' roda depois que este OnStartup inteiro terminar.
        _instancia.CaminhoRecebido += c => Dispatcher.BeginInvoke(() => _ = AbrirSeguroAsync(c));
        // O app aberto de novo SEM arquivo (menu Iniciar) enquanto este processo esta' vivo (fix round 1 da Task 5).
        _instancia.AtivacaoRecebida += () => Dispatcher.BeginInvoke(AoPedirAtivacao);
        if (!InstanciaUnica.DeveContinuar(() => _instancia.TentarAssumir(caminho), LogarErro))
        {
            Shutdown();
            return;
        }

        // Presenca para o instalador (AppMutex=), melhor esforco.
        AdquirirPresenca();

        base.OnStartup(e);

        // (4) Quem encerra o processo e' a POLITICA (ultima janela fechada), nunca o WPF contando janelas:
        // com OnLastWindowClose, esconder a boas-vindas para abrir a primeira janela de documento ja'
        // derrubaria o app.
        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        // A ultima janela fechada nao encerra NA HORA se ha' impressao em curso (fix round 1 da Task 5): o app
        // espera o envio terminar (sem janela nenhuma), ate' ImpressoesEmCurso.LimiteDeEspera.
        _impressoes = new ImpressoesEmCurso(EncerrarAgora, Logar);
        _janelas = new GerenteDeJanelas(CriarJanela, MostrarBoasVindas, EsconderBoasVindas, Encerrar);

        // (5) Sem arquivo: boas-vindas. Com arquivo: abre direto (nunca pisca a boas-vindas antes).
        if (caminho is null) _janelas.MostrarBoasVindasSeVazio();
        else
        {
            // Aberto direto: a boas-vindas (onde mora "Documentos recuperados") nao aparece. Se uma sessao
            // anterior caiu com trabalho nao salvo deste app, a janela do documento avisa (Task 3 do Plano 3).
            var n = ListarRecuperaveis().Count;
            if (n > 0) _avisoRecuperaveis = TextoAvisoRecuperaveis(_perfil, n);
            _ = AbrirSeguroAsync(caminho);
        }
    }

    // ---------------------------------------------------------------------------------------------------
    // Documentos recuperados (Task 3 do Plano 3)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Os documentos recuperaveis DESTE app (a raiz de recuperacao e' compartilhada pelos tres; o
    /// filtro e' o <c>App</c> do <c>sessao.json</c>, que a <see cref="JanelaEditor"/> grava com o mesmo
    /// <see cref="PerfilApp.Nome"/> passado ao construtor).</summary>
    private IReadOnlyList<DocumentoRecuperavel> ListarRecuperaveis() => Recuperaveis.Listar(SessaoDocumento.RaizRecovery, _perfil.Nome);

    /// <summary>
    /// Aviso na faixa da janela aberta direto quando ha' recuperaveis. Diverge do texto do brief ("Abra o X pelo
    /// menu Iniciar"): com o app ja' rodando, a instancia unica faz uma segunda invocacao SEM arquivo sair calada
    /// (<see cref="InstanciaUnica.TentarAssumir"/> so' encaminha caminho), entao abrir pelo menu Iniciar nao
    /// mostraria nada. O caminho que funciona e' fechar o app (a ultima janela encerra) e abri-lo de novo.
    /// </summary>
    internal static string TextoAvisoRecuperaveis(PerfilApp perfil, int n) =>
        $"Há {n} documento(s) não salvo(s) de uma sessão anterior. Para recuperá-los, feche o {perfil.NomeExibicao} e abra-o de novo pelo menu Iniciar.";

    /// <summary>
    /// O que a faixa mostra quando o documento fica pronto. Os avisos DESTE documento vao JUNTOS, nesta ordem (revisao
    /// final do Plano 3: antes era "o primeiro vence", e o aviso de midias indisponiveis ESCONDIA o da sessao recuperada
    /// sem senha — justamente o que evita gravar o documento sem protecao): midias indisponiveis (o erro do nucleo);
    /// sessao RECUPERADA de documento protegido, ainda sem senha (<see cref="JanelaEditor.AvisoProtegidoRecuperado"/>);
    /// sessao recuperada cujo original mudou depois da queda (<see cref="JanelaEditor.AvisoOriginalAlterado"/>). Sem
    /// nenhum deles, o aviso de recuperaveis pendente, que so' entao e' CONSUMIDO (o segundo valor). Pura, para teste.
    /// </summary>
    internal static (string? Aviso, bool ConsumiuRecuperaveis) AvisoAoFicarPronto(bool midiasIndisponiveis, string? ultimoErro,
        bool recuperadoSemSenha, bool originalAlterado, string? avisoRecuperaveis)
    {
        var avisos = new List<string>();
        if (midiasIndisponiveis && !string.IsNullOrWhiteSpace(ultimoErro)) avisos.Add(ultimoErro.Trim());
        if (recuperadoSemSenha) avisos.Add(JanelaEditor.AvisoProtegidoRecuperado);
        if (originalAlterado) avisos.Add(JanelaEditor.AvisoOriginalAlterado);
        if (avisos.Count > 0) return (string.Join(" ", avisos.Select(a => char.IsLetterOrDigit(a[^1]) ? a + "." : a)), false);
        return (avisoRecuperaveis, true);
    }

    /// <summary>Mensagem na boas-vindas quando a REAPLICACAO falha (<c>RecuperacaoFalhou</c>; so' ela — ver
    /// <see cref="MensagemErroAoRecuperar"/>). O final e' o que tira o usuario de um laco:
    /// uma falha deterministica (o mesmo changes0.json quebra toda vez) deixaria o item na lista para sempre, e
    /// "Recuperar" daria o mesmo erro — Descartar (que continua oferecido) ou abrir o original saem dele.</summary>
    internal static string MensagemFalhaRecuperacao(string mensagem) =>
        $"{mensagem.TrimEnd()} Se o problema continuar, use Descartar ou abra o arquivo original.";

    /// <summary>"Recuperar" da boas-vindas: reabre a pasta numa janela nova pela politica de janelas (mesmo
    /// tratamento de erro de <see cref="AbrirSeguroAsync"/>).</summary>
    private async Task RecuperarSeguroAsync(DocumentoRecuperavel d)
    {
        if (_encerrando) return;
        _impressoes.Cancelar();
        DesarmarAtualizacao("o usuario recuperou um documento");
        // O original ja' esta' aberto (o usuario o reabriu depois da queda): recuperar criaria uma SEGUNDA janela
        // do mesmo arquivo, e as duas gravariam por cima uma da outra.
        if (!d.Novo && _janelas.EstaAberto(d.Origem))
        {
            MostrarErroNaBoasVindas($"\"{d.Titulo}\" já está aberto em outra janela. Feche-o antes de recuperar as alterações não salvas.");
            return;
        }
        try { await _janelas.AbrirPorAsync(() => CriarJanela(j => j.RecuperarAsync(d.Pasta))); }
        catch (Exception ex)
        {
            LogarErro(ex);
            MostrarErroNaBoasVindas(MensagemErroAoRecuperar(ex));
        }
    }

    /// <summary>Mensagem de uma recusa/falha de <c>JanelaEditor.RecuperarAsync</c> (antes de o editor abrir). As
    /// recusas do nucleo ja' vem em pt-BR e completas ("Este documento já está aberto em outra janela.", "Não
    /// foi possível recuperar o documento: ..."); o resto ganha o prefixo. SEM o sufixo de
    /// <see cref="MensagemFalhaRecuperacao"/> (fix round 1 da revisao): ele e' so' da falha da reaplicacao
    /// (<c>RecuperacaoFalhou</c>) — "use Descartar" diante de "já está aberto em outra janela" mandaria o
    /// usuario apagar um documento que esta' aberto e bem.</summary>
    internal static string MensagemErroAoRecuperar(Exception ex) =>
        ex is InvalidOperationException ? ex.Message : $"Não foi possível recuperar o documento: {ex.Message}";

    /// <summary>"Descartar" da boas-vindas (depois do "Sim" na pergunta, que e' do VM).</summary>
    private static void DescartarRecuperavel(DocumentoRecuperavel d) => Recuperaveis.Descartar(d.Pasta);

    protected override void OnExit(ExitEventArgs e)
    {
        _instancia?.Dispose();
        _mutexDeAtualizacao?.Dispose();
        base.OnExit(e);
    }

    // ---------------------------------------------------------------------------------------------------
    // Janela de documento
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Adaptador do nucleo para a politica de janelas (ver <see cref="IJanelaDocumento"/>).</summary>
    private sealed class JanelaAdaptada : IJanelaDocumento
    {
        private readonly JanelaEditor _j;
        private readonly Func<Task> _carregar;

        public JanelaAdaptada(JanelaEditor j, Func<Task> carregar)
        {
            _j = j; _carregar = carregar;
            _j.Fechado += () => Fechado?.Invoke();
        }

        /// <summary>So' documento com ARQUIVO DE ORIGEM tem caminho: um documento novo (ainda nao salvo) tem
        /// como origem uma copia do modelo dentro da pasta de recuperacao — comparar por ele faria dois
        /// "Novo documento" parecerem a mesma janela. Sessao recuperada DESVINCULADA (I2 da revisao final do Plano 3)
        /// tambem nao: o original mudou depois da queda e esta janela nao o mostra (nem pode grava-lo) — abrir o
        /// original tem de abrir a versao NOVA numa janela propria, e nao ativar esta.</summary>
        public string? Caminho => _j.SessaoAtual is { Novo: false, Desvinculada: false } s ? s.ArquivoOrigem : null;

        public void Ativar()
        {
            if (_j.WindowState == WindowState.Minimized) _j.WindowState = WindowState.Normal;
            _j.Activate();
        }

        /// <summary>Abre o arquivo ou cria o documento novo. Leva segundos; o gerente so' chama depois de a
        /// janela estar registrada, entao um fechamento no meio disto e' visto pela politica.</summary>
        public Task CarregarAsync() => _carregar();

        public void Fechar() => _j.Close();

        public event Action? Fechado;
    }

    /// <summary>
    /// Cria e MOSTRA a janela do documento, ja' com toda a fiacao (barra, eventos, comandos), e devolve o
    /// adaptador na hora — sem esperar o documento carregar. O carregamento fica em
    /// <see cref="JanelaAdaptada.CarregarAsync"/>, que o gerente chama depois de registrar a janela: e' o
    /// que garante que fechar a janela durante o carregamento seja visto pela politica (revisao 1).
    /// </summary>
    private IJanelaDocumento CriarJanela(string? caminho) =>
        CriarJanela(j => caminho is null ? j.NovoAsync(_perfil.FormatoNovo) : j.AbrirAsync(caminho));

    /// <param name="carregar">O que a janela carrega: abrir um arquivo, criar um documento novo ou reabrir uma
    /// pasta de recuperacao.</param>
    private IJanelaDocumento CriarJanela(Func<JanelaEditor, Task> carregar)
    {
        var j = new JanelaEditor(_perfil.Nome)
        {
            NomeApp = _perfil.NomeExibicao,
            Tema = _tema.TemaDoEditor,
            PerguntarAoFechar = DialogoFechar.PerguntarAsync,
        };
        if (IconeDoApp() is { } icone) j.Icon = icone;

        // Acima do editor, SO' a faixa de aviso (invisivel sem aviso). A barra de comandos saiu a pedido do
        // usuario (2026-09-25): Novo/Abrir/Salvar/Salvar como ja' estao no menu Arquivo do editor e chegam
        // aqui por ComandoRecebido/LocalFileSave; Configuracoes e Sobre foram para o menu da janela.
        var avisoVm = new FaixaAvisoViewModel();
        j.BarraDoHost = new FaixaAviso { DataContext = avisoVm };
        AdicionarAoMenuDaJanela(j);

        j.DocumentoPronto += () => Dispatcher.BeginInvoke(() =>
        {
            // DocumentoPronto chega DENTRO da chamada sincrona do JS (doc do nucleo): gravar recentes e
            // mexer na faixa sai pela fila do Dispatcher, nunca ali dentro.
            if (j.SessaoAtual is { Novo: false } s) _recentes.Registrar(s.ArquivoOrigem);
            var (aviso, consumiu) = AvisoAoFicarPronto(j.MidiasIndisponiveis, j.UltimoErro, j.RecuperadoSemSenha, j.OriginalAlteradoDepoisDaQueda, _avisoRecuperaveis);
            avisoVm.Aviso = aviso;
            if (consumiu) _avisoRecuperaveis = null;
        });
        // TODO salvamento passa por aqui — o da barra, o Ctrl+S do editor e o "Salvar" do dialogo de
        // fechamento (achado da revisao 1): antes so' o "Salvar como" da barra registrava nos recentes,
        // entao um documento novo gravado por Ctrl+S nunca aparecia na lista. O nucleo ja' separa
        // exportacao de salvamento (PDF dispara Exportado, nao Salvo); o filtro aqui e' a segunda camada,
        // pelo formato que ESTE app abre — ver DeveRegistrarRecente.
        j.Salvo += c => { if (DeveRegistrarRecente(_perfil, c)) _recentes.Registrar(c); };
        // "Salvar" no dialogo de fechar que NAO deu certo (achado da revisao final do Plano 2): sem esta
        // linha o usuario clicava em Salvar, a janela continuava aberta e NADA explicava por que — a razao
        // so' existia no UltimoErro e no .log. A mensagem ja' vem em pt-BR do nucleo; a casca so' a mostra
        // na mesma faixa de aviso usada pelas falhas de salvar da barra.
        j.ErroAoFechar += m => { Logar("FECHAR " + m); AvisarNaBarra(j, m); };
        // Fechamento abandonado (Cancelar, "Salvar como" cancelado, falha ao salvar): o usuario ficou com o documento,
        // entao o "instalar ao encerrar" do botao de atualizacao e' desarmado (Plano 5, Task 2, fix round 1).
        j.FechamentoAbortado += () => DesarmarAtualizacao("um documento nao foi fechado", j);
        // Task 7 do Plano 3 (fix round 1): documento protegido salvo num formato que nao guarda senha (PDF).
        j.AvisoAoUsuario += m => AvisarNaBarra(j, m);
        // Impressao (Tasks 4 e 5 do Plano 3): o PDF de impressao nao pode ser gerado -> faixa de aviso; o PDF
        // pronto (ImprimirPedido) vai para a impressora pela WebView2 oculta (Impressora) e o resultado vai para a
        // mesma faixa ("Enviado para …" ou o erro). Nenhum dialogo nosso: o painel do editor escolhe tudo, e o
        // "diálogo do sistema" do painel e' o dialogo Imprimir do proprio Windows.
        j.ErroAoImprimir += (_, m) => { Logar("IMPRESSAO " + m); AvisarNaBarra(j, MensagemErroAoImprimir(m)); };
        // O PDF passa a ser de _impressoes no disparo (a janela nao o apaga mais): ele e' apagado no fim do envio,
        // mesmo com a janela ja' fechada — e fechar a ultima janela no meio nao encerra o app antes do fim.
        // Salvamentos em curso (fix round 4): fechar a ultima janela no meio de um Ctrl+S nao encerra o app antes de o
        // arquivo do usuario ser gravado — o mesmo dreno das impressoes, com o mesmo limite.
        j.SalvamentoIniciado += fim => _ = _impressoes.AcompanharAsync(fim, "salvamento");
        // "PDF assinado com mPDF" (Plano 4, Task 5): o PDF ja' foi exportado. O evento chega na thread da interface e,
        // pelo painel do editor, ANTES de o salvamento ser dado por terminado: detectar e lancar (registro + Process.Start,
        // que pode demorar) rodam fora dela, e so' a mensagem volta para a faixa desta janela.
        var assinatura = new PedidoDeAssinatura(DetectarMpdf, LancadorDoMpdf, Logar);
        j.AssinarPedido += (origem, pdf) => _ = AssinarAsync(assinatura, origem, pdf);
        j.ImprimirPedido += (origem, o, pdf) => _ = _impressoes.ImprimirAsync(pdf,
            () => new mOffice.Nucleo.Impressao.Impressora(_perfil.Nome, origem).ImprimirAsync(pdf, o),
            m => AvisarNaBarra(origem, m));
        j.ComandoRecebido += c => Dispatcher.BeginInvoke(() => AoComando(j, c));
        // Documento solto no editor (Task 6 do Plano 3): abre em janela nova pelo MESMO caminho da boas-vindas
        // (AbrirSeguroAsync, filtrado pelo perfil como o Tela_Drop dela). Formato de outro app da suite nao abre
        // aqui: a faixa diz por que, em vez de nada acontecer.
        j.AbrirPedido += c => Dispatcher.BeginInvoke(() =>
        {
            if (_perfil.Aceita(c)) _ = AbrirSeguroAsync(c);
            else AvisarNaBarra(j, $"O {_perfil.NomeExibicao} não abre arquivos {Path.GetExtension(c).TrimStart('.').ToUpperInvariant()}: abra-o pelo aplicativo do mOffice correspondente.");
        });

        // Reabertura recuperada que NAO conseguiu reaplicar as alteracoes (Task 2/3 do Plano 3): a janela se
        // fecha sozinha, sem a pergunta de salvar, e MANTEM a pasta. O evento chega ANTES do fechamento: a
        // janela sai do registro ja' (o fechamento dela nao pode encerrar o app se era a unica) e a mensagem
        // espera o Fechado — so' depois dele a trava esta' solta e a pasta volta a ser listada, entao a
        // boas-vindas recarregada ja' mostra o item de novo, com Descartar.
        JanelaAdaptada? adaptada = null;
        var falha = new AvisoFalhaRecuperacao(
            soltar: () => { if (adaptada is not null) _janelas.Soltar(adaptada); },
            mostrar: m => Dispatcher.BeginInvoke(() => MostrarErroNaBoasVindas(m)));
        j.RecuperacaoFalhou += (_, m) => { Logar("RECUPERACAO " + m); falha.AoFalhar(m); };
        // Falha ao gravar as alteracoes de uma sessao VIVA (fix round 3): mesma fiacao, SEM o "use Descartar" da
        // recuperacao — a mensagem do nucleo ja' diz para usar Recuperar.
        j.AlteracoesNaoGravadas += (_, m) => { Logar("ALTERACOES " + m); falha.AoFalhar(m, sufixoDeRecuperacao: false); };
        j.Fechado += falha.AoFechar;
        j.Show();

        adaptada = new JanelaAdaptada(j, () => carregar(j));
        return adaptada;
    }

    /// <summary>Desarma o "instalar ao encerrar" (se armado): um fechamento foi abandonado (<paramref name="j"/> recebe o
    /// aviso na faixa) ou o usuario voltou a trabalhar no app (abrir, novo, recuperar, ativar) — nos dois casos o app
    /// nao vai encerrar agora, e instalar mais tarde, num fechamento qualquer, seria uma surpresa.</summary>
    private void DesarmarAtualizacao(string motivo, JanelaEditor? j = null)
    {
        var caminho = _atualizacao.Caminho;
        if (caminho is null || !_atualizacao.Desarmar()) return;
        Logar("ATUALIZACAO adiada: " + motivo);
        if (j is not null) AvisarNaBarra(j, AtualizacaoAoEncerrar.MensagemAdiada(caminho));
    }

    /// <summary>
    /// A fiacao de <c>JanelaEditor.RecuperacaoFalhou</c>, separada da janela para ser testavel (Task 3 do Plano
    /// 3). A ORDEM e' o contrato: <see cref="AoFalhar"/> chega ANTES de a janela fechar sozinha e a solta da
    /// politica ja' (<paramref name="soltar"/>: o fechamento dela nao pode encerrar o app se era a unica);
    /// <see cref="AoFechar"/> chega depois que a janela soltou a trava da pasta (o <c>Closed</c> da
    /// <c>JanelaEditor</c> solta a trava e so' entao dispara <c>Fechado</c>) — so' ai' a mensagem vai para a
    /// boas-vindas (<paramref name="mostrar"/>), que ao reaparecer relista o item, ainda com Descartar.
    /// Fechamento sem falha nao mostra nada.
    /// </summary>
    internal sealed class AvisoFalhaRecuperacao(Action soltar, Action<string> mostrar)
    {
        private string? _mensagem;
        private bool _sufixo;

        /// <param name="sufixoDeRecuperacao">Acrescenta o "Se o problema continuar, use Descartar…" de
        /// <see cref="MensagemFalhaRecuperacao"/> — so' na falha da REAPLICACAO; a falha de gravacao de uma sessao viva
        /// (fix round 3) vai sem ele.</param>
        public void AoFalhar(string mensagem, bool sufixoDeRecuperacao = true)
        {
            if (_mensagem is not null) return;
            _mensagem = mensagem;
            _sufixo = sufixoDeRecuperacao;
            soltar();
        }

        public void AoFechar()
        {
            if (_mensagem is { } m) mostrar(_sufixo ? MensagemFalhaRecuperacao(m) : m);
        }
    }

    /// <summary>Faz o pedido de assinatura fora da thread da interface e mostra o resultado na faixa da janela (a
    /// continuacao do <c>await</c> volta ao Dispatcher). Nunca lanca: e' disparado sem ninguem esperar.</summary>
    private async Task AssinarAsync(PedidoDeAssinatura pedido, JanelaEditor origem, string pdf)
    {
        try { AvisarNaBarra(origem, await pedido.ProcessarAsync(pdf)); }
        catch (Exception ex) { LogarErro(ex); }
    }

    /// <summary>
    /// "PDF assinado com mPDF" depois da exportacao (Plano 4, Task 5), separado da janela para ser testavel (o mesmo
    /// padrao de <see cref="AvisoFalhaRecuperacao"/>): acha o mPDF (<paramref name="detectar"/>), abre-o com
    /// <c>/assinar</c> se ele for 2.13 ou mais novo (<paramref name="lancador"/>) e devolve a mensagem da faixa, em
    /// pt-BR. Sem mPDF, ou mPDF antigo, NAO lanca nada — so' explica. Falha do lancamento: a faixa diz o que fazer e o
    /// detalhe tecnico (em ingles, do Windows) vai so' para o log (<paramref name="logar"/>).
    /// </summary>
    internal sealed class PedidoDeAssinatura(Func<MpdfInstalado?> detectar, ILancadorMpdf lancador, Action<string> logar)
    {
        /// <summary><see cref="Processar"/> numa thread do pool: o registro e o <c>Process.Start</c> nao seguram a interface.</summary>
        public Task<string> ProcessarAsync(string pdf) => Task.Run(() => Processar(pdf));

        public string Processar(string pdf)
        {
            MpdfInstalado? mpdf;
            try { mpdf = detectar(); }
            catch (Exception ex) { logar("ASSINAR deteccao do mPDF falhou: " + ex); mpdf = null; }
            if (mpdf is null)
            {
                logar("ASSINAR mPDF nao encontrado; PDF em " + pdf);
                return MensagemAssinarSemMpdf(pdf);
            }
            if (!DetectorMpdf.AbreAssinatura(mpdf))
            {
                logar($"ASSINAR mPDF {mpdf.Versao} anterior a {DetectorMpdf.Minima}; PDF em {pdf}");
                return MensagemAssinarMpdfAntigo(pdf, mpdf.Versao);
            }
            try
            {
                lancador.Assinar(mpdf, pdf);
                logar($"ASSINAR mPDF {mpdf.Versao} aberto com /assinar ({mpdf.Exe})");
                return MensagemAssinarAberto(pdf);
            }
            catch (Exception ex)
            {
                logar($"ASSINAR nao foi possivel abrir o mPDF ({mpdf.Exe}): {ex}");
                return MensagemAssinarFalhou;
            }
        }
    }

    /// <summary>Faixa: o mPDF foi aberto com o PDF para assinar.</summary>
    internal static string MensagemAssinarAberto(string pdf) =>
        $"PDF salvo em {Path.GetDirectoryName(pdf)}. O mPDF foi aberto para assinar.";

    /// <summary>Faixa: nao ha' mPDF instalado.</summary>
    internal static string MensagemAssinarSemMpdf(string pdf) =>
        $"O PDF foi salvo em {pdf}. Para assinar, instale o mPDF 2.13 ou mais novo e abra este PDF nele.";

    /// <summary>Faixa: o mPDF instalado e' anterior a' 2.13 (sem o verbo <c>/assinar</c>).</summary>
    internal static string MensagemAssinarMpdfAntigo(string pdf, Version versao) =>
        $"O PDF foi salvo em {pdf}. O mPDF instalado ({versao}) é anterior à 2.13 e não abre a assinatura automaticamente: atualize o mPDF ou abra o PDF nele e use Assinar.";

    /// <summary>Faixa: o mPDF existe e e' novo, mas o processo nao pode ser iniciado (o detalhe fica so' no log).</summary>
    internal const string MensagemAssinarFalhou =
        "O PDF foi salvo, mas não foi possível abrir o mPDF: abra o PDF no mPDF e use Assinar.";

    /// <summary>
    /// Este arquivo recem-gravado entra na lista de RECENTES deste app? So' entra o que o app sabe REABRIR:
    /// um recente que o mTexto nao abre e' um item que so' pode dar erro quando clicado. Na pratica corta
    /// a exportacao para PDF (o nucleo ja' a separa em <c>Exportado</c>, mas a lista de recentes e' do app,
    /// e quem decide o que ele abre e' o <see cref="PerfilApp"/>) e qualquer formato de outro app da suite.
    /// <c>internal</c> por ser a costura de teste desta regra — nao ha' como instanciar um
    /// <see cref="Application"/> no teste.
    /// </summary>
    internal static bool DeveRegistrarRecente(PerfilApp perfil, string caminho) =>
        !string.IsNullOrWhiteSpace(caminho) && perfil.Aceita(caminho);

    /// <summary>Abre pelo gerente convertendo QUALQUER falha em mensagem na boas-vindas. E' o unico caminho
    /// de abertura do app (linha de comando, pipe da segunda instancia, boas-vindas, barra, comando do
    /// editor): uma excecao solta em qualquer um deles vira excecao nao tratada do Dispatcher — app abaixo
    /// por um arquivo corrompido.</summary>
    private async Task AbrirSeguroAsync(string caminho)
    {
        // O app ja' esta' encerrando (a ultima janela fechou): abrir agora criaria uma janela que o WPF
        // recusa mostrar. Acontece de verdade — um caminho chega pelo pipe da segunda instancia no mesmo
        // instante em que a primeira fecha a ultima janela.
        if (_encerrando) return;
        _impressoes.Cancelar(); // aberto pelo usuario durante a espera das impressoes: o app segue
        DesarmarAtualizacao("o usuario abriu um arquivo");
        try { await _janelas.AbrirAsync(caminho); }
        catch (Exception ex) { LogarErro(ex); MostrarErroNaBoasVindas($"Não foi possível abrir o arquivo: {ex.Message}"); }
    }

    private async Task NovoSeguroAsync()
    {
        if (_encerrando) return;
        _impressoes.Cancelar();
        DesarmarAtualizacao("o usuario criou um documento");
        try { await _janelas.NovoAsync(); }
        catch (Exception ex) { LogarErro(ex); MostrarErroNaBoasVindas($"Não foi possível criar o documento: {ex.Message}"); }
    }

    // ---------------------------------------------------------------------------------------------------
    // Boas-vindas (uma so' por processo, criada sob demanda)
    // ---------------------------------------------------------------------------------------------------

    private void MostrarBoasVindas()
    {
        if (_boasVindas is null)
        {
            _vmBoasVindas = new TelaBoasVindasViewModel(_perfil, _recentes,
                novo: NovoSeguroAsync,
                abrir: AbrirSeguroAsync,
                escolherArquivo: () => Task.FromResult(EscolherArquivoParaAbrir(_boasVindas)),
                configuracoes: () => { if (_boasVindas is { } b) MostrarConfiguracoes(b); },
                sobre: () => { if (_boasVindas is { } b) MostrarSobre(b); },
                recuperaveis: ListarRecuperaveis,
                recuperar: RecuperarSeguroAsync,
                descartar: DescartarRecuperavel,
                confirmarDescarte: d => TelaBoasVindasViewModel.ConfirmarDescarteNoSistema(_boasVindas, d));
            _boasVindas = new JanelaBoasVindas(_vmBoasVindas, _perfil.ChaveLogo);
            if (IconeDoApp() is { } icone) _boasVindas.Icon = icone;
            AdicionarAoMenuDaJanela(_boasVindas);
            _boasVindas.Closing += BoasVindas_Closing;
            _boasVindas.Closed += (_, _) => { _boasVindas = null; _vmBoasVindas = null; };
        }
        _vmBoasVindas!.Recarregar(); // recentes e recuperados podem ter mudado enquanto ela estava escondida
        _avisoRecuperaveis = null;   // a propria boas-vindas mostra a lista: o aviso da faixa ficou redundante
        _boasVindas.Show();
        _boasVindas.Activate();
    }

    private void EsconderBoasVindas() => _boasVindas?.Hide();

    /// <summary>O que o gerente chama quando a ultima janela fecha (e a boas-vindas, quando fechada sem documento):
    /// encerra, mas so' depois das impressoes em curso (<see cref="ImpressoesEmCurso.PedirEncerramento"/>).</summary>
    private void Encerrar() => _impressoes.PedirEncerramento();

    /// <summary>Encerra de verdade. A marca serve a <see cref="MostrarErroNaBoasVindas"/>: depois do Shutdown nenhuma
    /// janela pode mais ser mostrada. Defensivo: se uma janela de documento reabriu no meio, nao encerra.</summary>
    private void EncerrarAgora()
    {
        if (_janelas.Abertas > 0) { Logar("ENCERRAMENTO ignorado: ha' janela de documento aberta"); return; }
        // "Instalar agora" das Configuracoes: todas as janelas fecharam, entao o instalador armado sobe AGORA (antes
        // do Shutdown, como no mPDF). Se nao subir (outro app da suite aberto, instalador alterado, UAC recusado, falha
        // do Windows), o app NAO encerra: a boas-vindas explica.
        if (_atualizacao.IniciarSeArmada() is { } falha)
        {
            Dispatcher.BeginInvoke(() => MostrarErroNaBoasVindas(falha));
            return;
        }
        _encerrando = true;
        Shutdown();
    }

    /// <summary>
    /// O app foi aberto de novo SEM arquivo (menu Iniciar) com este processo vivo (fix round 1 da Task 5). Com
    /// documento aberto nada muda (a janela ja' esta' na tela; comportamento de antes). Sem documento — o app
    /// esperando impressoes para encerrar, ou so' a boas-vindas — cancela a espera e mostra a boas-vindas: sem isto o
    /// usuario abriria o app e nada apareceria (o processo vivo, invisivel, engolia o pedido).
    /// </summary>
    private void AoPedirAtivacao()
    {
        if (_encerrando) return;
        DesarmarAtualizacao("o app foi aberto de novo");
        if (_janelas.Abertas > 0) { Logar("ATIVACAO pedida pela segunda instancia com documento aberto (ignorada)"); return; }
        _impressoes.Cancelar();
        _janelas.MostrarBoasVindasSeVazio();
    }

    private void MostrarErroNaBoasVindas(string mensagem)
    {
        // Esperando impressoes para encerrar: nada aparece sozinho (o usuario fechou tudo).
        if (_encerrando || _impressoes.Drenando) { Logar("AVISO nao mostrado (encerrando): " + mensagem); return; }
        // MOSTRA sempre (nao "se vazio"): a boas-vindas e' a unica superficie de erro do app, e uma falha de
        // abertura com outra janela aberta ficaria invisivel se so' aparecesse com o app vazio.
        MostrarBoasVindas();
        _vmBoasVindas!.MensagemErro = mensagem;
    }

    private void BoasVindas_Closing(object? sender, CancelEventArgs e)
    {
        // Com documento aberto ela e' so' um painel: fechar ESCONDE. Sem documento ela e' a janela do app:
        // fechar encerra (spec §4, "a ultima janela fechada encerra").
        if (_janelas.Abertas > 0) { e.Cancel = true; _boasVindas?.Hide(); return; }
        Encerrar();
    }

    // ---------------------------------------------------------------------------------------------------
    // Abrir / salvar pela barra
    // ---------------------------------------------------------------------------------------------------

    /// <param name="dono">Janela sobre a qual o dialogo abre. Sem dono, o dialogo do Windows pode ficar
    /// ATRAS da janela que o pediu — o usuario ve' o app travado sem entender por que.</param>
    private string? EscolherArquivoParaAbrir(Window? dono)
    {
        var d = new OpenFileDialog { Filter = _perfil.FiltroAbrir, CheckFileExists = true, Multiselect = false };
        var ok = dono is null ? d.ShowDialog() : d.ShowDialog(dono);
        return ok == true ? d.FileName : null;
    }

    private async Task EscolherEAbrirAsync(Window? dono)
    {
        var caminho = EscolherArquivoParaAbrir(dono);
        if (caminho is not null) await AbrirSeguroAsync(caminho);
    }

    /// <summary>Texto da faixa quando o PDF de impressao nao pode ser gerado (separado para teste).</summary>
    internal static string MensagemErroAoImprimir(string mensagem) => "Não foi possível preparar a impressão: " + mensagem;

    /// <summary>Texto da faixa quando o trabalho chegou a' fila da impressora (separado para teste). A faixa fica
    /// ate' o usuario dispensar, como os demais avisos (sem temporizador). Fila que o Windows da' como offline NAO e'
    /// erro (fix round 1, decisao do controlador: impressora de rede dormindo aceita o trabalho) — so' e' avisada.</summary>
    internal static string MensagemEnviadoParaImpressora(string impressora, bool offline = false) =>
        offline
            ? $"Enviado para {impressora}. O Windows indica que ela está desligada ou offline: o trabalho fica na fila até ela voltar."
            : $"Enviado para {impressora}.";

    /// <summary>Texto da faixa quando o envio para a impressora falhou (separado para teste). As mensagens do
    /// nucleo ja' vem em pt-BR e com o nome da impressora.</summary>
    internal static string MensagemErroAoEnviar(string mensagem) => "Não foi possível imprimir: " + mensagem;

    private static void AvisarNaBarra(JanelaEditor j, string mensagem)
    {
        if (j.BarraDoHost is FaixaAviso f && f.DataContext is FaixaAvisoViewModel vm) vm.Aviso = mensagem;
    }

    // ---------------------------------------------------------------------------------------------------
    // Configuracoes e Sobre (menu da janela + rodape da boas-vindas)
    // ---------------------------------------------------------------------------------------------------

    private void MostrarConfiguracoes(Window dono)
    {
        var vm = new DialogoConfiguracoesViewModel(_config, _tema,
            outrosAppsAbertos: () => OutrosAppsDaSuite.Abertos(_perfil), nomeApp: _perfil.NomeExibicao);
        DialogoConfiguracoes.Mostrar(dono, vm);
        // "Verificar atualizações" -> "Instalar agora?" -> Sim -> baixado e verificado: o dialogo ja' fechou.
        // Pela fila do Dispatcher: o dialogo pode ter vindo do menu do sistema (dentro do WndProc da propria janela
        // que vai ser fechada) — fechar janelas so' depois que essa pilha terminar.
        if (vm.InstaladorPronto is { } arquivo) Dispatcher.BeginInvoke(() => IniciarAtualizacao(arquivo));
    }

    /// <summary>
    /// Fecha o app para instalar a atualizacao: arma o instalador e pede o fechamento de TODAS as janelas de documento
    /// pelo fluxo normal (cada documento modificado pergunta se salva). Quando a ultima fecha, o gerente encerra e
    /// <see cref="EncerrarAgora"/> inicia o instalador (com reabertura). Sem documento aberto (so' a boas-vindas),
    /// encerra direto. Um fechamento abandonado em qualquer documento desarma (<see cref="JanelaEditor.FechamentoAbortado"/>).
    /// Com OUTRO app da suite aberto nao arma (ver <see cref="OutrosAppsDaSuite"/>).
    /// </summary>
    private void IniciarAtualizacao(ServicoAtualizacao.ArquivoVerificado arquivo)
    {
        // Outro app da suite aberto (conferido de novo: o download demora e o usuario pode te-lo aberto no meio): o
        // instalador o fecharia sem perguntar se salva. Nao arma; o arquivo verificado fica onde esta'.
        if (OutrosAppsDaSuite.Abertos(_perfil) is { Count: > 0 } outros)
        {
            Logar("ATUALIZACAO nao armada: outro app da suite aberto (" + string.Join(", ", outros) + ")");
            MostrarErroNaBoasVindas(OutrosAppsDaSuite.MensagemFecharOutros(outros));
            return;
        }
        _atualizacao.Armar(arquivo);
        Logar($"ATUALIZACAO instalar agora: {_janelas.Abertas} janela(s) de documento para fechar; instalador em {arquivo.Caminho}");
        if (_janelas.Abertas == 0) Encerrar();
        else _janelas.FecharTodas();
    }

    private void MostrarSobre(Window dono) => DialogoSobre.Mostrar(dono, _perfil, Versao, Motor.Versao);

    /// <summary>Os textos do menu da janela, separados para teste.</summary>
    internal static IReadOnlyList<string> TextosDoMenuDaJanela(PerfilApp perfil) =>
        ["Configurações do mOffice…", $"Sobre o {perfil.NomeExibicao}…"];

    private void AdicionarAoMenuDaJanela(Window janela)
    {
        var textos = TextosDoMenuDaJanela(_perfil);
        MenuDoSistema.Adicionar(janela, [
            (textos[0], () => MostrarConfiguracoes(janela)),
            (textos[1], () => MostrarSobre(janela)),
        ]);
    }

    // ---------------------------------------------------------------------------------------------------
    // Comandos do editor (execCommand do web-apps)
    // ---------------------------------------------------------------------------------------------------

    private void AoComando(JanelaEditor j, ComandoDoEditor c)
    {
        switch (c.Comando)
        {
            case "create:new":
                _ = NovoSeguroAsync();
                return;

            case "open:recent":
                // O web-apps manda {title, path, id, type} (documenteditor/main/app.js, tag 9.4.0.129) — o
                // campo e' `path`; `url` fica como reserva para as telas que usam o nome do modelo.
                var recente = CampoDeTexto(c.Json, "path") ?? CampoDeTexto(c.Json, "url");
                if (recente is null) { Logar($"COMANDO open:recent sem caminho: {c.Json}"); return; }
                _ = AbrirSeguroAsync(recente);
                return;

            case "recent:forget":
                var alvo = CampoDeTexto(c.Json, "path") ?? CampoDeTexto(c.Json, "url");
                // O web-apps chama `execCommand("recent:forget")` SEM parametro. Sem saber QUAL item
                // esquecer, nao se apaga nada: limpar a lista inteira por um comando ambiguo seria perder
                // dado do usuario. So' registra (e, quando um dia vier com caminho, funciona).
                if (alvo is null) { Logar("COMANDO recent:forget sem caminho (ignorado)"); return; }
                _recentes.Remover(alvo);
                _vmBoasVindas?.Recarregar();
                return;

            case "editor:event":
                switch (c.AcaoDoEvento())
                {
                    case "file:open": _ = EscolherEAbrirAsync(j); return;
                    case "file:close": j.Close(); return;
                    default: Logar($"COMANDO editor:event nao tratado: {c.Json}"); return;
                }

            default:
                // So' o que NAO e' ruido conhecido (achado da revisao final do Plano 2): antes, todo
                // execCommand sem tratamento virava uma linha no log do processo — inclusive os de mouse da
                // barra de titulo, ~150 por segundo. O log da casca existe para achar um comando novo que
                // ninguem tratou; um rio de telemetria o torna inutil exatamente para isso.
                if (!EhRuidoConhecido(c.Comando)) Logar($"COMANDO nao tratado: {c.Comando} {c.Json}");
                return;
        }
    }

    /// <summary>Comandos de UI do web-apps que chegam SEMPRE e que a casca nao trata de proposito: handshake
    /// de inicializacao (<c>webapps:entry</c>, <c>webapps:features</c>, <c>editor:config</c>,
    /// <c>editor:onready</c>, <c>doc:onready</c>) e estado de interface (<c>quickaccess:changed</c>,
    /// <c>uitheme:changed</c>). Exatos, nao por prefixo — um comando novo qualquer tem de continuar
    /// aparecendo no log.</summary>
    internal static readonly IReadOnlySet<string> ComandosSilenciosos = new HashSet<string>(StringComparer.Ordinal)
    {
        "webapps:entry", "webapps:features", "editor:config", "editor:onready", "doc:onready",
        "quickaccess:changed", "uitheme:changed",
    };

    /// <summary>Familias inteiras de comando que sao ruido: a barra de titulo customizada
    /// (<c>title:mousemove</c>, <c>title:button</c>, …) e as dicas de Alt (<c>althints:show</c>, …). Aqui o
    /// prefixo se justifica porque o ruido e' a familia toda, e nao um comando ou outro dela — o nucleo
    /// silencia so' os tres de mouse (<see cref="mOffice.Nucleo.JanelaEditor.ComandosSilenciosos"/>) porque
    /// la' o log e' o de EGRESSO daquela janela, onde ver `title:button` ainda ajuda.</summary>
    internal static readonly string[] PrefixosSilenciosos = ["title:", "althints:"];

    /// <summary>Este comando e' ruido conhecido (nao vai para o log)? <c>internal</c> por ser a costura de
    /// teste da regra — nao ha' como instanciar um <see cref="Application"/> no teste.</summary>
    internal static bool EhRuidoConhecido(string comando) =>
        ComandosSilenciosos.Contains(comando)
        || PrefixosSilenciosos.Any(p => comando.StartsWith(p, StringComparison.Ordinal));

    /// <summary>Le um campo string de um JSON de comando; null se o parametro nao for objeto JSON ou o campo
    /// nao existir/nao for string. Nunca lanca: o que chega aqui veio do editor, nao do nosso codigo.</summary>
    private static string? CampoDeTexto(string json, string campo)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind != JsonValueKind.Object) return null;
            return doc.RootElement.TryGetProperty(campo, out var v) && v.ValueKind == JsonValueKind.String
                ? v.GetString() is { Length: > 0 } s ? s : null
                : null;
        }
        catch (JsonException) { return null; }
    }

    // ---------------------------------------------------------------------------------------------------
    // Icone e log
    // ---------------------------------------------------------------------------------------------------

    /// <summary>O <c>.ico</c> commitado do app (<c>Assets/&lt;app&gt;.ico</c> dentro do EXE), usado na
    /// JanelaEditor e na boas-vindas. Uma so' instancia por processo; se a leitura falhar, a janela fica com
    /// o icone padrao em vez de o app nao subir.</summary>
    private ImageSource? IconeDoApp()
    {
        if (_iconeTentado) return _icone;
        _iconeTentado = true;
        try { _icone = new BitmapImage(new Uri($"pack://application:,,,/{_perfil.NomeExibicao};component/Assets/{_perfil.Nome}.ico")); }
        catch (Exception ex) when (ex is IOException or UriFormatException or NotSupportedException or FileFormatException) { LogarErro(ex); }
        return _icone;
    }

    private void LogarErro(Exception ex) => Logar("ERRO " + ex);

    private void Logar(string linha)
    {
        try { File.AppendAllText(_log ??= Logs.NovoArquivo(_perfil.Nome), $"{DateTime.Now:yyyy-MM-ddTHH:mm:ss} {linha}{Environment.NewLine}"); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* log e' diagnostico: nunca derruba o app */ }
    }
}
