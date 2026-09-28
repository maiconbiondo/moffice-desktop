using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using mOffice.Conversor;
using mOffice.Nucleo.Impressao;
using mOffice.Nucleo.Ortografia;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Nucleo;

public partial class JanelaEditor : Window
{
    /// <summary>Pasta de dados do WebView2 do app "sonda" (o host do PoC). Mantida para os testes que
    /// precisam falar dela por nome; quem decide o caminho e' <see cref="AmbienteWebView2"/>.</summary>
    public static string PastaDadosWebView => AmbienteWebView2.PastaDadosUsuario(AppPadrao);

    /// <summary>App da suite a que esta janela pertence quando ninguem diz outro (o host do PoC).</summary>
    public const string AppPadrao = "sonda";

    /// <summary>Liga a injecao do grampo.js (loga toda chamada da ponte JS -&gt; C#) — desligado por padrao no
    /// produto (backlog 15: o grampo e' instrumentacao de diagnostico, nao pode logar cada chamada por
    /// padrao). Lido em <see cref="InicializarAsync"/>, nao no construtor: a Sonda liga por <c>--grampo</c>
    /// antes de criar a janela; o produto liga pela configuracao (Task 4).</summary>
    public static bool GrampoLigado { get; set; }

    /// <summary>Origens que podem usar os objetos de host da ponte — as MESMAS que o ambiente autoriza no
    /// esquema ascdesktop:// (uma lista so', em <see cref="AmbienteWebView2.OrigensPonte"/>).</summary>
    private static string[] OrigensPonte => AmbienteWebView2.OrigensPonte;

    public GuardaEgresso Guarda { get; } = new();
    /// <summary>Último erro de abertura ou de salvamento. O PoC não abre MessageBox: em teste de janela real ela
    /// travaria a thread STA.</summary>
    public string? UltimoErro { get; private set; }
    /// <summary>Documento aberto no momento (nulo antes do primeiro <see cref="AbrirAsync"/>).</summary>
    public SessaoDocumento? SessaoAtual => _sessao;

    /// <summary>A troca da base das mídias (<see cref="ChamarEndLoadAsync"/>) falhou no último documento
    /// aberto: o texto continua editável, mas nenhuma figura vai aparecer (achado maior da revisão 1,
    /// 2026-09-18). Antes disso a falha só existia dentro de <see cref="TrocaDaBaseDeMidia"/> — uma string
    /// de log que nada olhava — e o PoC seguia em frente como se as imagens estivessem lá. Este bool é o que
    /// os apps do Plano 2 leem para mostrar uma barra de aviso; ver também <see cref="UltimoErro"/>, setado
    /// junto.</summary>
    public bool MidiasIndisponiveis { get; private set; }

    /// <summary>Costura de teste (achado maior da revisão 1): corrompe de propósito o valor que
    /// <c>trocarBaseMidia</c> procura, simulando uma troca que NÃO acha a propriedade (motor com um bundle
    /// diferente do esperado) sem precisar de um motor desatualizado de verdade. `internal`: não é API do
    /// host, só o controle negativo do teste de janela real de mídias liga isto.</summary>
    internal bool DesligarTrocaDeBaseParaTeste { get; set; }

    /// <summary>
    /// Liga a contagem das midias do documento REALMENTE servidas (<see cref="ContarMidiasServidasAsync"/>).
    /// Tem de ser definida ANTES de <see cref="Window.Show"/>: a instrumentacao e' instalada na inicializacao
    /// da janela, antes de qualquer navegacao — depois disso os pedidos que ja passaram nao voltam.
    ///
    /// Por que nao e' sempre ligada: o unico jeito de VER o que o mapeamento de host virtual serve e' o
    /// DevTools Protocol (<c>Network</c>) — a <see cref="GuardaEgresso"/> nao enxerga esse trafego (ver o
    /// comentario de classe dela e o de <see cref="SondaInstrumentacao"/>), e habilitar o dominio
    /// <c>Network</c> faz o navegador emitir um evento por requisicao em TODA abertura de documento. E'
    /// instrumentacao de teste/sonda, entao fica desligada no produto.
    /// </summary>
    public bool InstrumentarRede { get; set; }

    /// <summary>Disparado UMA vez por documento, quando o editor sinaliza `asc_onDocumentContentReady`
    /// (ponte-desktop.js -> Avisar("pronto")). E' o instante que a medida de abertura da Task 7 cronometra: o
    /// host nao tem como saber sozinho quando o documento ficou pronto — quem sabe e' o editor. Roda na thread
    /// da interface, DENTRO da chamada sincrona do JS: quem assina deve fazer pouca coisa.</summary>
    public event Action? DocumentoPronto;

    // -------------------------------------------------------------------------------------------------------
    // API de documento para a CASCA (Task 2 do Marco 1 / Plano 2). Tudo o que o mTexto/mPlanilha/mSlides
    // precisam da janela e que o PoC nao tinha: nome do app no titulo, tema do editor, slot para a barra do
    // host, eventos de estado e a pergunta de "salvar ao fechar".
    // -------------------------------------------------------------------------------------------------------

    /// <summary>Texto depois do travessao na barra de titulo ("peticao.docx — mTexto"). Cada app da suite passa
    /// o proprio nome; o PoC fica com "mOffice".</summary>
    public string NomeApp { get; init; } = "mOffice";

    /// <summary>Tema do EDITOR desta janela. Lido UMA vez na inicializacao (vai na URL de entrada como
    /// <c>uitheme=</c> e no <c>window.RendererProcessVariable.theme</c>): `init` de proposito porque trocar o
    /// tema de uma janela JA ABERTA nao entra neste plano — uma janela nova nasce com o tema atual.</summary>
    public TemaEditor Tema { get; init; } = TemaEditor.Claro;

    /// <summary>Conteudo WPF do host ACIMA do editor (a FaixaAviso do mOffice.Comum; a barra de comandos saiu em 2026-09-25). Propriedade
    /// de dependencia porque o <c>ContentPresenter</c> do XAML se liga a ela — com um CLR simples a barra
    /// definida depois do <c>Show()</c> nunca apareceria.</summary>
    public static readonly DependencyProperty BarraDoHostProperty =
        DependencyProperty.Register(nameof(BarraDoHost), typeof(object), typeof(JanelaEditor));

    public object? BarraDoHost { get => GetValue(BarraDoHostProperty); set => SetValue(BarraDoHostProperty, value); }

    /// <summary>Ha alteracao nao salva no documento aberto.</summary>
    public bool Modificado => _sessao?.Modificado == true;

    /// <summary>Dispara na thread da interface quando <see cref="SessaoDocumento.Modificado"/> MUDA de valor
    /// (nao a cada aviso do editor). A casca usa para o asterisco da barra e para habilitar "Salvar".</summary>
    public event Action<bool>? ModificadoAlterado;

    /// <summary>Todo <c>execCommand</c> do web-apps (create:new, open:recent, editor:event, …), na thread da
    /// interface. O nucleo NAO decide nada com eles: quem trata e' a casca.</summary>
    public event Action<ComandoDoEditor>? ComandoRecebido;

    /// <summary>Depois do <see cref="Window.Closed"/> e do apagamento da pasta de recuperacao.</summary>
    public event Action? Fechado;

    /// <summary>
    /// O documento acabou de ser gravado com sucesso em <c>caminho</c> — TODO caminho de salvamento passa
    /// por aqui: o <see cref="SalvarAsync"/>/<see cref="SalvarComoAsync"/> do host, o Ctrl+S e o "Salvar
    /// como" do proprio editor (<c>SalvarPedidoDoEditorAsync</c>) e o "Salvar" do dialogo de fechamento.
    ///
    /// Existe porque a casca precisa registrar o arquivo nos RECENTES, e antes disto ela so' conseguia
    /// faze-lo no unico salvamento que passava por ela (o "Salvar como" da barra): um Ctrl+S num documento
    /// novo, ou o "Salvar" do dialogo de fechamento, gravavam um arquivo que nunca aparecia na lista de
    /// recentes. Dispara na thread da interface, DEPOIS de a sessao ja' apontar para o arquivo novo.
    /// </summary>
    public event Action<string>? Salvo;

    /// <summary>
    /// O documento foi EXPORTADO para <c>caminho</c> num formato que este editor nao reabre (PDF/PDF/A —
    /// ver <see cref="FormatoDocumentoExtensoes.EhExportacao"/>). E' evento SEPARADO do
    /// <see cref="Salvo"/> de proposito: exportar nao e' salvar. A sessao continua no documento de origem
    /// (o titulo nao muda, o asterisco de "modificado" continua onde estava, o proximo Ctrl+S grava no
    /// documento e nao no PDF) e a casca NAO poe o arquivo nos recentes — o mTexto nao abre PDF.
    /// </summary>
    public event Action<string>? Exportado;

    /// <summary>
    /// O usuario escolheu a linha "PDF assinado com mPDF" no "Salvar como" do host (Plano 4, Task 2) e o PDF ja'
    /// foi EXPORTADO com sucesso (depois do <see cref="Exportado"/>): argumentos = a janela e o caminho do PDF.
    /// Quem abre o mPDF para assinar e' quem assina (a casca, Task 5). Nunca dispara se a exportacao falhou ou foi
    /// cancelada, nem numa gravacao de documento. Roda na thread da interface, com try/catch como os demais
    /// eventos publicos (uma excecao do assinante nao vira erro de salvamento).
    /// </summary>
    public event Action<JanelaEditor, string>? AssinarPedido;

    /// <summary>
    /// O fechamento com alteracoes pendentes NAO conseguiu ir adiante e a janela CONTINUA ABERTA: a gravacao
    /// do "Salvar" falhou, a pergunta de fechar estourou (e a reserva tambem), ou o destino escolhido era uma
    /// exportacao (ver <see cref="ResultadoSalvar"/>). A mensagem ja' vem pronta em pt-BR, para a casca so'
    /// mostrar.
    ///
    /// Existe porque ate' a revisao final do Plano 2 essa falha era INVISIVEL: o <see cref="DecidirFechamentoAsync"/>
    /// setava <see cref="UltimoErro"/>, escrevia no log e voltava — do lado do usuario, clicar em "Salvar" no
    /// dialogo de fechar simplesmente nao fazia nada, sem explicacao nenhuma. Dispara na thread da interface,
    /// com try/catch como os demais eventos publicos.
    /// </summary>
    public event Action<string>? ErroAoFechar;

    /// <summary>
    /// Um pedido de fechar que PERGUNTOU e terminou com a janela ainda aberta: "Cancelar" na pergunta, "Salvar" com o
    /// dialogo de destino cancelado, ou qualquer falha que chama <see cref="ErroAoFechar"/>. Plano 5, Task 2 (fix round
    /// 1): a casca desarma o "instalar a atualizacao ao encerrar" — o usuario decidiu manter o documento aberto. Nao
    /// dispara quando a janela fecha (inclusive por fora, no encerramento do Windows). Thread da interface, try/catch.
    /// </summary>
    public event Action? FechamentoAbortado;

    /// <summary>Pergunta "salvar as alteracoes?" no fechamento com documento modificado. O padrao e' um
    /// <c>MessageBox</c> Sim/Nao/Cancelar em pt-BR (<see cref="PerguntaPadrao"/>); o mOffice.Comum troca pelo
    /// dialogo com tema (Task 5), e os testes trocam por uma resposta fixa (um modal de verdade travaria a
    /// thread STA).</summary>
    public Func<JanelaEditor, Task<DecisaoFechar>> PerguntarAoFechar { get; set; } = PerguntaPadrao;

    /// <summary>Pergunta de RESERVA, usada UMA vez quando <see cref="PerguntarAoFechar"/> estoura (achado da
    /// revisao final do Plano 2): sem ela, um dialogo da casca que falhe SEMPRE — recurso de tema ausente,
    /// excecao no construtor do dialogo — deixa a janela impossivel de fechar, porque todo <c>Close()</c> cai
    /// no mesmo caminho e volta. `internal` (e nao `public`): nao e' API do host: e' o padrao do WPF, e o
    /// teste que prova o caminho "ate a reserva falhou" precisa faze-la falhar tambem (um
    /// <c>MessageBox</c> de verdade travaria a thread STA do teste).</summary>
    internal Func<JanelaEditor, Task<DecisaoFechar>> PerguntaDeReserva { get; set; } = PerguntaPadrao;

    private static Task<DecisaoFechar> PerguntaPadrao(JanelaEditor j)
    {
        var r = MessageBox.Show(j, $"Salvar as alterações em \"{j.SessaoAtual?.Titulo}\"?", j.NomeApp,
            MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        return Task.FromResult(r switch
        {
            MessageBoxResult.Yes => DecisaoFechar.Salvar,
            MessageBoxResult.No => DecisaoFechar.NaoSalvar,
            _ => DecisaoFechar.Cancelar,
        });
    }

    /// <summary>O fechamento ja passou pela pergunta (ou nao precisava): o <see cref="AoFechar"/> deixa passar.</summary>
    private bool _fechamentoConfirmado;

    /// <summary>Uma decisao de fechamento esta em curso (pergunta aberta ou gravacao em andamento). Enquanto
    /// for true, um novo <see cref="AoFechar"/> so' cancela o fechamento — ver o comentario la'.</summary>
    private bool _decidindoFechamento;

    /// <summary>Quanto esperar o editor devolver <c>LocalFileSave</c> depois de pedirmos <c>asc_Save</c>. Nao
    /// mutavel (achado 8/9 da revisao 2026-09-17, fix review 1): nada no projeto precisava trocar o valor em
    /// runtime, e um `set` publico era superficie para um teste mudar o timeout de uma janela e esquecer de
    /// devolver, vazando para os testes seguintes.</summary>
    private static readonly TimeSpan EsperaSalvar = TimeSpan.FromSeconds(30);

    private readonly TaskCompletionSource<CoreWebView2> _pronto = new(TaskCreationOptions.RunContinuationsAsynchronously);
    // Task 1 do Marco 1 / Plano 2 (backlog 5/15): morava numa pasta de artefato de bancada do PoC (fora dos
    // dados por usuario); agora e' um arquivo por janela em Pastas.Logs (Logs.NovoArquivo, que ja' carrega
    // app+data+pid — o sufixo unico que este campo precisava contra duas janelas criadas no MESMO segundo,
    // ex.: dois testes de janela real em paralelo, vem do pid mais o timestamp com segundos).
    private readonly string _log;
    /// <summary>Arquivo de log desta janela (so' para os testes conferirem o que foi — e o que NUNCA pode ser — registrado).</summary>
    internal string ArquivoDeLog => _log;
    private readonly ConfiguracaoMotor _motor;
    private readonly PonteArquivos _arquivos;
    private readonly PonteLog _ponteLog;
    // NAO e' readonly: o servidor le o AllFonts.js UMA vez, na construcao (a whitelist por arquivo sai de
    // la'). Num PC onde o catalogo ainda nao existe, o construtor da janela roda ANTES de
    // CatalogoFontes.GarantirAtualizadoAsync e leria zero fontes — exatamente o bug da Task 7 (PPTX com fonte
    // do Microsoft 365 nao abria). Por isso ele e' recriado em InicializarAsync, depois da geracao.
    private ServidorAscDesktop _fontes;
    private readonly X2tConversor _conv;
    private SessaoDocumento? _sessao;
    private PonteDesktop? _ponte;
    private Task<string>? _binPronto;
    // Changes da sessao recuperada, lidos UMA vez em RecuperarAsync (depois do reparo) e entregues no EndLoad
    // (fix round 1 da revisao da Task 2). Nulo em sessoes normais e depois do EndLoad (o array pode ser grande).
    private (string ArrayJs, int Quantidade)? _changesRecuperados;
    // Salvamento pedido pelo C# (SalvarComoAsync) e ainda sem o LocalFileSave de volta. Quando esta' nulo, um
    // LocalFileSave que chega veio do proprio editor (Ctrl+S / menu) e o host e' quem decide o destino.
    private TaskCompletionSource<string>? _pedidoSalvar;
    // Midias do documento servidas com 200 (so' com InstrumentarRede ligado). O receptor do evento do
    // DevTools Protocol precisa continuar referenciado: sem isso o GC o recolhe e os eventos param de chegar
    // (mesma licao de SondaInstrumentacao).
    private int _midiasServidas;
    private CoreWebView2DevToolsProtocolEventReceiver? _receptorRede;

    /// <summary>Costura de teste para "inserir imagem do arquivo": substitui o <c>OpenFileDialog</c> do
    /// Windows (modal — nao da para dirigir num teste automatizado) por um caminho fixo. Todo o RESTO do
    /// caminho continua o do produto. `internal`: nao e' API do host.</summary>
    internal Func<string?>? EscolherImagemParaTeste { get; set; }

    /// <summary>Costura de teste da recuperacao apos queda (Task 2 do Plano 3): ao fechar, so' solta a trava da
    /// pasta de recuperacao em vez de apaga-la — o estado em disco que uma queda de verdade deixaria.
    /// `internal`: nao e' API do host. E' o MESMO mecanismo que a falha da reaplicacao liga
    /// (<see cref="FalharRecuperacao"/>).</summary>
    internal bool ManterRecoveryAoFecharParaTeste { get => _manterRecoveryAoFechar; set => _manterRecoveryAoFechar = value; }

    /// <summary>Ao fechar: solta a trava e DEIXA a pasta de recuperacao (em vez de apaga-la).</summary>
    private bool _manterRecoveryAoFechar;

    /// <summary>
    /// A sessao RECUPERADA ja' mostrou o conteudo recuperado (fix rounds 1 e 2 da revisao da Task 3): chegaram os
    /// DOIS sinais, da MESMA sessao, em qualquer ordem — <see cref="_endLoadRecuperadoOk"/> (o EndLoad com os
    /// changes voltou "ok", sem erro da reaplicacao) e <see cref="_prontoRecuperado"/> (o editor avisou
    /// "pronto": <c>asc_onDocumentContentReady</c>, via <c>marcarPronto</c> da ponte). So' o EndLoad nao basta:
    /// no editor de texto o <c>DesktopOfflineAppDocumentApplyChanges</c> so' ENFILEIRA os changes (medido no
    /// bundle 9.4: <c>qNd</c> -&gt; <c>NHe</c> -&gt; <c>AscCommon.Vc.Crh</c>, que e' <c>this.c4.push</c>), e eles
    /// sao aplicados depois, durante a abertura — uma falha ali acontece DEPOIS de o EndLoad voltar "ok".
    /// Politica de RETENCAO: ate' isto ficar
    /// verdadeiro, a pasta de recuperacao e' a UNICA copia do trabalho nao salvo do usuario — o editor ainda nao
    /// o mostrou, entao "Nao salvar" nao e' uma escolha informada. Enquanto <see cref="RecuperacaoPendente"/>,
    /// QUALQUER fechamento (o X, o <c>Fechar</c> do gerente de janelas depois de uma falha de carregamento, o
    /// encerramento do app) fecha sem perguntar e so' solta a trava, mantendo a pasta. Zerado a cada sessao
    /// nova na janela.
    /// </summary>
    private bool _recuperacaoConfirmada;

    /// <summary>Sinal 1 da confirmacao: o EndLoad da sessao recuperada voltou "ok" e sem erro da reaplicacao.</summary>
    private bool _endLoadRecuperadoOk;

    /// <summary>Sinal 2 da confirmacao: o editor desta sessao avisou "pronto" (conteudo na tela).</summary>
    private bool _prontoRecuperado;

    /// <summary>Sinal 1 ja' chegou (so' leitura, para os testes).</summary>
    internal bool EndLoadRecuperadoOk => _endLoadRecuperadoOk;

    /// <summary>Costura de teste (fix round 2 da revisao da Task 3): o aviso "pronto" do editor nao conta para a
    /// confirmacao da recuperacao — como um editor que aplicou os changes enfileirados, falhou e nunca ficou
    /// pronto, depois de um EndLoad "ok". `internal`.</summary>
    internal bool IgnorarProntoParaTeste { get; set; }

    /// <summary>Confirma a recuperacao quando os dois sinais (EndLoad ok + pronto) ja' chegaram para a sessao
    /// atual e nenhuma falha foi registrada. Chamado por cada um dos dois sinais — a ordem nao importa.</summary>
    private void ConfirmarRecuperacaoSePronta()
    {
        if (_recuperacaoConfirmada || _sessao is not { Recuperada: true } || FalhaDaRecuperacao is not null) return;
        if (!_endLoadRecuperadoOk || !_prontoRecuperado) return;
        _recuperacaoConfirmada = true;
        _arquivos.Log("RECUPERACAO confirmada: EndLoad ok e editor pronto (conteudo recuperado na tela)");
    }

    /// <summary>Sessao recuperada cujo conteudo ainda nao apareceu no editor (ver <see cref="_recuperacaoConfirmada"/>).</summary>
    private bool RecuperacaoPendente => _sessao is { Recuperada: true } && !_recuperacaoConfirmada;

    /// <summary>O conteudo recuperado ja' apareceu no editor (so' leitura, para os testes).</summary>
    internal bool RecuperacaoConfirmada => _recuperacaoConfirmada;

    /// <summary>A janela ja' fechou (Closed): um segundo <see cref="Window.Close"/> enfileirado vira no-op.</summary>
    private bool _fechada;

    /// <summary>Costura de teste (fix round 1 da revisao da Task 3): a navegacao para o editor lanca, como se o
    /// WebView tivesse falhado — a abertura falha DEPOIS de a sessao ja' estar na janela. `internal`.</summary>
    internal bool FalharNavegacaoParaTeste { get; set; }

    /// <summary>Costura de teste (fix round 1 da revisao da Task 3): o host nunca responde ao
    /// <c>LocalStartOpen</c> — o editor fica em "Carregando documento", como num travamento. `internal`.</summary>
    internal bool SuprimirEndLoadParaTeste { get; set; }

    /// <summary>Costura de teste (fix round 2 da revisao da Task 2): faz o bloco de reaplicacao do script do
    /// EndLoad lancar, como se <c>DesktopOfflineAppDocumentApplyChanges</c> tivesse falhado. `internal`.</summary>
    internal bool FalharReaplicacaoParaTeste { get; set; }

    /// <summary>Mensagem (pt-BR) da falha ao reaplicar as alteracoes de uma sessao RECUPERADA, ou <c>null</c>.
    /// Nada a apaga: a janela desta sessao fecha em seguida e a pasta de recuperacao fica intacta (ver
    /// <see cref="RecuperacaoFalhou"/>). Nome diferente do evento porque o C# nao deixa propriedade e evento
    /// com o mesmo nome.</summary>
    public string? FalhaDaRecuperacao { get; private set; }

    /// <summary>
    /// A reabertura de uma pasta de recuperacao (<see cref="RecuperarAsync"/>) NAO conseguiu reaplicar as
    /// alteracoes gravadas (fix round 2 da revisao da Task 2 — "falhar fechado"). O editor mostraria o
    /// documento sem elas (ou com parte delas: o de planilha aplica uma a uma) enquanto o arquivo de changes
    /// continua com todas — salvar gravaria algo diferente da tela, e "nao salvar" apagaria a unica copia.
    /// Por isso, nesta ordem: a falha fica em <see cref="FalhaDaRecuperacao"/>, a janela e' marcada para
    /// MANTER a pasta ao fechar, este evento dispara (argumentos: a janela e a mensagem pt-BR para o usuario)
    /// e a janela fecha sem perguntar se deseja salvar. A pasta continua recuperavel. Roda na thread da
    /// interface. Quem mostra a mensagem e' a casca (AplicativoOffice, Task 3).
    /// </summary>
    public event Action<JanelaEditor, string>? RecuperacaoFalhou;

    /// <summary>
    /// O editor pediu para imprimir (painel <b>Arquivo -&gt; Imprimir</b> ou impressao rapida, Task 4 do Plano 3) e
    /// o PDF de impressao ja' esta' pronto: argumentos = a janela, as opcoes do painel e o caminho do PDF (em
    /// <see cref="Pastas.Impressao"/>). Quem manda o PDF para a impressora e' quem
    /// assina (Task 5). Roda na thread da interface, com try/catch como os demais eventos publicos.
    ///
    /// <para><b>Dono do PDF</b> (Task 5): com assinante, o PDF passa a ser DELE no disparo — a janela o tira da
    /// lista dela e NAO o apaga ao fechar (fechar o documento no meio da impressao nao pode apagar o arquivo por
    /// baixo da impressora); quem assina apaga depois de imprimir, com sucesso ou nao. Sem assinante o PDF continua
    /// da janela. A varredura de 1 dia da subida (<see cref="PastaImpressao.Varrer"/>) e' a rede de seguranca.</para>
    /// </summary>
    public event Action<JanelaEditor, OpcoesImpressao, string>? ImprimirPedido;

    /// <summary>O PDF de impressao NAO pode ser gerado (descarga dos changes ou x2t falhou). A mensagem vem em
    /// pt-BR, sem prefixo — a casca mostra "Não foi possível preparar a impressão: {mensagem}".</summary>
    public event Action<JanelaEditor, string>? ErroAoImprimir;

    /// <summary>PDFs de impressao que ESTA janela gerou e ainda sao DELA: apagados no fechamento (melhor esforco;
    /// ver <see cref="PastaImpressao"/>). Os entregues a quem imprime (<see cref="ImprimirPedido"/>) saem daqui no
    /// disparo. Nunca a pasta inteira — ela e' compartilhada com as outras janelas.</summary>
    private readonly List<string> _pdfsDeImpressao = [];

    /// <summary>Trava de <see cref="_pdfsDeImpressao"/> e da passagem de <see cref="_fechada"/> para true (fix round
    /// 1 da revisao da Task 4): um PDF que fica pronto DEPOIS do fechamento nunca entra numa lista que o Closed ja'
    /// esvaziou — ou entra antes (e o Closed apaga), ou ve' a janela fechada (e e' apagado na hora).</summary>
    private readonly object _travaImpressao = new();

    /// <summary>A impressao pedida pelo editor mais recente (so' para os testes esperarem ela acabar). `internal`.</summary>
    internal Task? ImpressaoEmCursoParaTeste { get; private set; }

    /// <summary>Costura de teste (fix round 1): roda depois de o x2t gerar o PDF de impressao e ANTES de ele ser
    /// registrado na janela — o teste fecha a janela aqui para provar o caminho "terminou depois de fechar". `internal`.</summary>
    internal Func<Task>? AntesDeRegistrarPdfParaTeste { get; set; }

    /// <summary>App da suite dono desta janela (mTexto/mPlanilha/mSlides; "sonda" no host do PoC). Decide a
    /// pasta de dados do WebView2 na PRIMEIRA janela do processo — as seguintes herdam o ambiente ja' criado.</summary>
    private readonly string _app;

    public JanelaEditor() : this(AppPadrao) { }

    public JanelaEditor(string app) : this(app, ConfiguracaoMotor.Localizar()) { }

    /// <summary>Costura de teste (Task 1 do Marco 1 / fonte de simbolos livre): abre a janela sobre OUTRA pasta
    /// de motor. O controle negativo do teste de janela real precisa de um motor cujo unico bundle diferente e' o
    /// `word/sdk-all.js` com uma fonte de simbolos VAZIA — e isso nao pode ser feito no motor compartilhado
    /// (a suite roda outros testes de janela real sobre ele). O teste monta um motor temporario — `editors` e'
    /// COPIA de verdade (o mapeamento de pasta do WebView2 nao segue juncoes); `converter`, `fonts` e
    /// `dictionaries` sao juncoes (mklink /J) — e passa a pasta por aqui. `internal`: nao e' API do host; o app
    /// sempre usa <see cref="ConfiguracaoMotor.Localizar"/>.</summary>
    internal JanelaEditor(string app, ConfiguracaoMotor motor)
    {
        _app = app;
        _motor = motor;
        _log = Logs.NovoArquivo(app);
        InitializeComponent();
        _arquivos = new PonteArquivos(_log);
        _ponteLog = new PonteLog(_log);
        _conv = new X2tConversor(_motor);
        _fontes = CriarServidorDeFontes();
        // AtualizarTitulo aqui (e nao no construtor): NomeApp e' `init`, entao o inicializador de objeto do
        // chamador (`new JanelaEditor("mtexto") { NomeApp = "mTexto" }`) so' roda DEPOIS do construtor.
        Loaded += async (_, _) => { AtualizarTitulo(); try { await InicializarAsync(); } catch (Exception ex) { _pronto.TrySetException(ex); } };
        // Retencao (achado da revisao final 2026-09-17): a pasta de recovery guarda o documento inteiro em
        // claro (Editor.bin + media/ + changes/) e nada a apagava — 353 pastas orfas medidas em %TEMP%. O PoC
        // nao tem recuperacao apos queda, entao nao ha o que preservar depois que a janela fecha. Melhor
        // esforco (ver SessaoDocumento.ApagarRecovery); o que escapar cai na varredura da proxima subida.
        Closed += (_, _) =>
        {
            // Fechamento limpo apaga a pasta (e antes solta a trava — SessaoDocumento.ApagarRecovery). So' a
            // costura de teste da queda (ManterRecoveryAoFecharParaTeste, Task 2 do Plano 3) e a falha da
            // reaplicacao de uma sessao recuperada (FalharRecuperacao) deixam a pasta la', soltando a trava
            // como o SO faria se o processo tivesse morrido.
            // Fix round 1 da revisao da Task 3: sessao recuperada que ainda nao mostrou o conteudo tambem mantem
            // a pasta (ver _recuperacaoConfirmada) — inclusive no encerramento do app e no Fechar do gerente.
            //
            // Revisao final do Plano 3 (C1, critico): fechamento SEM a decisao do usuario com alteracao pendente
            // tambem mantem a pasta. E' o encerramento do Windows (logoff/reinicio: WM_QUERYENDSESSION) e o
            // Application.Shutdown — o WPF fecha as janelas IGNORANDO o e.Cancel do Closing, entao a pergunta
            // "Salvar?" enfileirada pelo AoFechar nunca roda, e antes disto este Closed apagava a UNICA copia do
            // trabalho nao salvo. So' um fechamento CONFIRMADO (_fechamentoConfirmado: "Salvar" deu certo ou o
            // usuario escolheu "Nao salvar") apaga uma sessao com alteracao pendente.
            lock (_travaImpressao) _fechada = true;
            // Fix round 2: o segundo termo e' "ha' um salvamento em curso" (o asc_Save ja' zerou o Modificado, e o
            // arquivo ainda nao foi gravado), NAO a anotacao _edicaoNaoGravada — ela fica verdadeira depois de editar,
            // salvar, editar e desfazer ate' o estado salvo, e a pasta de um documento inteiramente salvo sobrava e
            // reaparecia em "Documentos recuperados" (com o Editor.bin em claro de um documento protegido).
            var pendenteSemDecisao = !_fechamentoConfirmado && (_sessao?.Modificado == true || SalvamentoEmCurso);
            if (_manterRecoveryAoFechar || RecuperacaoPendente || pendenteSemDecisao)
            {
                if (RecuperacaoPendente && !_manterRecoveryAoFechar)
                    _arquivos.Log("RECUPERACAO fechada antes de o conteudo aparecer: a pasta fica para outra tentativa");
                else if (pendenteSemDecisao && !_manterRecoveryAoFechar)
                    _arquivos.Log("FECHAR sem a decisao do usuario com alteracao nao salva (encerramento do Windows ou do app): a pasta fica para recuperacao");
                _sessao?.LiberarTrava();
            }
            else _sessao?.ApagarRecovery();
            string[] pdfs;
            lock (_travaImpressao) { pdfs = [.. _pdfsDeImpressao]; _pdfsDeImpressao.Clear(); }
            foreach (var pdf in pdfs) PastaImpressao.Apagar(pdf);
            // try/catch como nos demais eventos publicos: um assinante que estoure aqui nao tem para onde
            // propagar (a janela ja fechou) e derrubaria o processo pelo manipulador de evento.
            try { Fechado?.Invoke(); } catch (Exception ex) { _arquivos.Log("FECHAR erro no aviso de fechado: " + ex.Message); }
        };
        // Task 2 do Marco 1 / Plano 2: "fechar com alteracoes pendentes" pergunta antes (ver AoFechar).
        Closing += AoFechar;
        // Tela cheia (Task 7 do Plano 3): por cima de tudo SO' enquanto ativa — um Alt+Tab durante a apresentacao
        // nao pode deixar a janela presa na frente das outras.
        Activated += (_, _) => { if (EmTelaCheia) Topmost = true; };
        Deactivated += (_, _) => { if (EmTelaCheia) Topmost = false; };
        // Rede de seguranca (fix round 1): um Esc que chegue ao WPF (tecla aceleradora que o WebView2 repassa) tambem
        // arma a saida garantida — o caminho principal e' o aviso "escTelaCheia" da ponte (ver EscNaTelaCheia).
        PreviewKeyDown += (_, e) => { if (e.Key == System.Windows.Input.Key.Escape && EmTelaCheia) EscNaTelaCheia("teclado do host"); };
    }

    /// <summary>"peticao.docx* — mTexto": titulo do documento + asterisco quando ha alteracao nao salva + nome
    /// do app. Sem documento aberto, so' o nome do app. Chamado ao abrir, ao salvar e a cada mudanca de
    /// <see cref="SessaoDocumento.Modificado"/>.</summary>
    private void AtualizarTitulo() =>
        Title = _sessao is null ? NomeApp : $"{_sessao.Titulo}{(_sessao.Modificado ? "*" : "")} — {NomeApp}";

    /// <summary>
    /// Pergunta antes de descartar alteracoes. O WPF NAO espera uma continuacao assincrona dentro do
    /// <see cref="Window.Closing"/> — quando o `await` da pergunta devolvesse, o fechamento ja teria
    /// acontecido. Por isso: cancela ESTE fechamento (<c>e.Cancel = true</c>), decide FORA do evento e, com a
    /// decisao tomada, chama <see cref="Window.Close"/> de novo com <see cref="_fechamentoConfirmado"/> ligado.
    ///
    /// O <c>BeginInvoke</c> nao e' enfeite: o <see cref="PerguntarAoFechar"/> padrao (MessageBox) e o dos
    /// testes respondem de forma SINCRONA, entao um `await` direto continuaria DENTRO do proprio manipulador
    /// de <c>Closing</c> — e um <see cref="Window.Close"/> reentrante ali e' invalido no WPF. Enfileirar a
    /// decisao garante que o fechamento em curso ja terminou (cancelado) quando ela roda.
    /// </summary>
    private void AoFechar(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // Fechar (Alt+F4, X da barra de tarefas) com a apresentacao em tela cheia (fix round 1 da Task 7): sai da tela
        // cheia ANTES de tudo — a pergunta de salvar aparece sobre a janela normal e, se o fechamento for cancelado, o
        // usuario nao fica preso numa janela sem borda por cima de tudo.
        if (EmTelaCheia) DefinirTelaCheia(false);
        if (_fechamentoConfirmado || _sessao is null || !_sessao.Modificado) return;
        // Sessao recuperada que ainda nao mostrou o conteudo (fix round 1 da revisao da Task 3): nao ha' o que
        // perguntar — o usuario nao viu o documento, e "Nao salvar" apagaria a unica copia. Fecha direto; o
        // Closed so' solta a trava e mantem a pasta. (Sem o e.Cancel, o fechamento segue: vale tambem para o
        // Fechar do gerente e para o encerramento do app, que nao podem ficar presos numa pergunta.)
        if (RecuperacaoPendente) return;
        e.Cancel = true;
        // Ja ha uma decisao em curso (achado 3 da revisao 1): o usuario clicou no X de novo enquanto a
        // pergunta estava aberta, ou enquanto o "Salvar" dela gravava. Enfileirar uma SEGUNDA decisao seria
        // pior do que nao fazer nada: o segundo SalvarAsync troca o _pedidoSalvar por Interlocked.Exchange e
        // CANCELA o primeiro ("O pedido de salvamento foi substituído por outro antes de o editor
        // responder"), entao o fechamento que estava salvando morre na excecao e a janela nao fecha — com o
        // documento possivelmente ja gravado. So' cancela este fechamento e deixa a decisao em curso terminar.
        if (_decidindoFechamento) return;
        Dispatcher.BeginInvoke(() => _ = DecidirFechamentoAsync());
    }

    private async Task DecidirFechamentoAsync()
    {
        var sessao = _sessao;
        // A janela ja' fechou sem esperar a decisao (C1 da revisao final do Plano 3: o encerramento do Windows fecha
        // ignorando o e.Cancel do AoFechar, e esta decisao enfileirada chega depois — ou nunca). Nao ha' mais o que
        // perguntar: o Closed ja' manteve a pasta de recuperacao.
        if (sessao is null || _fechada) return;
        _decidindoFechamento = true;
        try
        {
            DecisaoFechar decisao;
            try { decisao = await PerguntarAoFechar(this); }
            catch (Exception ex)
            {
                _arquivos.Log("FECHAR pergunta falhou: " + ex.Message);
                // Reserva UMA vez (achado da revisao final do Plano 2): uma pergunta da casca que estoure
                // SEMPRE tornava a janela impossivel de fechar — cada Close() vinha parar aqui e voltava sem
                // nada na tela. Cair no MessageBox padrao devolve o controle ao usuario.
                try { decisao = await PerguntaDeReserva(this); }
                catch (Exception ex2)
                {
                    // Nem a reserva respondeu: trata como CANCELAR (nunca como "nao salvar" — descartar
                    // alteracao do usuario por causa de um dialogo quebrado seria o pior desfecho possivel),
                    // mas agora dizendo POR QUE a janela nao fechou.
                    _arquivos.Log("FECHAR pergunta de reserva tambem falhou: " + ex2.Message);
                    Falhar("Não foi possível perguntar se deseja salvar as alterações: " + ex2.Message);
                    return;
                }
            }
            // A janela fechou por fora enquanto a pergunta estava aberta (encerramento do Windows): nada a gravar
            // numa janela fechada, e a pasta ja' ficou (Closed).
            if (_fechada) return;
            switch (decisao)
            {
                case DecisaoFechar.Cancelar:
                    Abortado();
                    return;
                case DecisaoFechar.Salvar:
                    try
                    {
                        // O destino sai SEMPRE da `sessao` capturada, nos dois ramos (achado 14 da revisao
                        // 1): antes, o ramo "novo" lia a sessao capturada e o outro chamava SalvarAsync(),
                        // que le' `_sessao` — dois estados diferentes no mesmo metodo. (A gravacao em si
                        // continua sendo do documento corrente da janela: SalvarComoAsync e
                        // ConverterESinalizarAsync operam sobre `_sessao`.)
                        //
                        // `permitirExportacao: false` e' o achado CRITICO da revisao final do Plano 2: o
                        // dialogo oferecia PDF tambem aqui, e escolher PDF caia no ramo de EXPORTACAO de
                        // ConverterESinalizarAsync — grava o PDF, dispara Exportado, NAO reponta a sessao. O
                        // fechamento seguia em frente do mesmo jeito, confirmava, fechava, e o Closed apagava
                        // a pasta de recuperacao: o documento EDITAVEL do usuario sumia, sobrando so' o PDF.
                        // Neste caminho o dialogo so' oferece formatos de DOCUMENTO.
                        // Sessao recuperada DESVINCULADA (I2 da revisao final do Plano 3): o original mudou
                        // depois da queda, entao o "Salvar" do fechamento tambem vira "Salvar como".
                        DestinoSalvar? alvo = sessao.Novo || sessao.Desvinculada
                            ? EscolherDestino(sessao.Formato, permitirExportacao: false)
                            : new DestinoSalvar(sessao.ArquivoOrigem, sessao.Formato);
                        if (alvo is null) { Abortado(); return; }  // cancelou o dialogo = cancelou o fechamento
                        var r = await SalvarComoAsync(alvo.Caminho, alvo.Formato);
                        // Defesa em profundidade: com o filtro acima isto e' inalcancavel pelo dialogo, mas se
                        // um dia outro caminho trouxer uma exportacao para ca, o fechamento NAO pode seguir —
                        // o documento editavel nao foi gravado em lugar nenhum.
                        if (r.Exportou)
                        {
                            Falhar("O arquivo foi exportado, não salvo: o documento editável não foi gravado, "
                                + "então a janela continua aberta. Salve em um formato de documento antes de fechar.");
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        // Nao fecha com a gravacao falhada: seria perder o documento do usuario em silencio.
                        _arquivos.Log("FECHAR salvar falhou: " + ex.Message);
                        Falhar("Não foi possível salvar: " + ex.Message);
                        return;
                    }
                    break;
            }
            _fechamentoConfirmado = true;
            Close();
        }
        finally { _decidindoFechamento = false; }

        // Fechamento abortado: guarda a razao em pt-BR (UltimoErro, para quem consultar depois) e avisa a
        // casca AGORA (ErroAoFechar, para o usuario ver uma faixa em vez de um clique que nao faz nada).
        // Ja' estamos na thread da interface; o try/catch e' o dos demais eventos publicos.
        void Falhar(string mensagem)
        {
            UltimoErro = mensagem;
            try { ErroAoFechar?.Invoke(mensagem); }
            catch (Exception ex) { _arquivos.Log("FECHAR erro no aviso de erro ao fechar: " + ex.Message); }
            Abortado();
        }

        // Todo fechamento que perguntou e ficou aberto passa por aqui (Cancelar, destino cancelado, Falhar).
        void Abortado()
        {
            try { FechamentoAbortado?.Invoke(); }
            catch (Exception ex) { _arquivos.Log("FECHAR erro no aviso de fechamento abortado: " + ex.Message); }
        }
    }

    private ServidorAscDesktop CriarServidorDeFontes()
    {
        // O logger vai junto: um AllFonts.js malformado faz o parser falhar FECHADO (zero fontes extras), e
        // isso tem de aparecer no log de egresso — senao volta a ser o mesmo bug de PPTX/ODP que nega em silencio.
        var s = ServidorAscDesktop.Padrao(_motor, _arquivos.Log);
        s.RaizExtra = () => _sessao?.PastaRecovery;
        return s;
    }

    /// <summary>Dicionario do usuario ("Adicionar ao dicionario") — um por PROCESSO, como o corretor: duas
    /// janelas com instancias proprias perderiam as palavras uma da outra ate' reabrir o app.</summary>
    private static readonly Lazy<DicionarioUsuario> _dicionarioUsuario = new(() => new DicionarioUsuario(Pastas.DicionarioUsuario));

    private async Task InicializarAsync()
    {
        // Corretor ortografico (Task 1 do Plano 3): os dicionarios (pt_BR ~0,7 s) carregam em segundo plano
        // AGORA, antes do catalogo de fontes e da navegacao, para o primeiro SpellCheck do editor nao pagar a
        // carga na thread da interface. Uma carga por processo (Corretor.Compartilhado); nunca lanca.
        _ = Corretor.Compartilhado(_motor, _arquivos.Log).PreCarregarAsync();
        // PRIMEIRA coisa da inicializacao: o catalogo de fontes do mOffice (AllFonts.js + font_selection.bin +
        // fonts_thumbnail*.png), gerado por nos com o allfontsgen embarcado. Antes da Task 2 o produto
        // EMPRESTAVA o catalogo do ONLYOFFICE instalado — ou seja, nao rodava num PC sem ele. Vem antes de
        // criar o WebView2 porque tudo depois disto depende do catalogo: o mapeamento de pasta (a copia do
        // AllFonts.js para dentro do motor) e a whitelist do ServidorAscDesktop.
        var catalogo = await CatalogoFontes.Padrao(_motor, _arquivos.Log).GarantirAtualizadoAsync();
        if (catalogo.Erro is not null)
        {
            // Sem catalogo o editor nao carrega (o requirejs morre no modulo "allfonts"). O modo de falha
            // certo e' dizer o que aconteceu, em pt-BR, e NAO navegar — nunca "Carregando documento" para
            // sempre. Quem chamar AbrirAsync/ProntoAsync recebe a mesma mensagem, em vez de esperar eterno.
            UltimoErro = catalogo.Erro;
            _arquivos.Log("CATALOGO DE FONTES: " + catalogo.Erro);
            _pronto.TrySetException(new InvalidOperationException(catalogo.Erro));
            return;
        }
        // Recriado AQUI (e nao so' no construtor): e' agora que o AllFonts.js existe, na primeira execucao.
        _fontes = CriarServidorDeFontes();

        // O ambiente do WebView2 e' UNICO por processo e vive em AmbienteWebView2 (Task 3/S2): runtime FIXO
        // embarcado, argumentos de linha de comando documentados e o registro do esquema nativo ascdesktop://
        // (que e' por AMBIENTE, nao por janela — foi medido na sonda da Task 4: sem ele o editor leva
        // ERR_FAILED ao pedir ascdesktop://fonts/<caminho>). Antes da Task 3 cada janela criava o seu, o que
        // dava uma arvore de processos do navegador por documento aberto.
        var env = await AmbienteWebView2.ObterAsync(_app);
        // O aviso e' registrado no log DESTA janela (e nao so' no da primeira do processo, que e' a unica que
        // chega a criar o ambiente): cada log de egresso tem de dizer sobre qual runtime aquela sessao rodou.
        _arquivos.Log(AmbienteWebView2.Aviso is not null
            ? "AMBIENTE WEBVIEW2: " + AmbienteWebView2.Aviso
            : $"AMBIENTE WEBVIEW2: runtime fixo {AmbienteWebView2.VersaoRuntimeFixo} (motor {env.BrowserVersionString})");
        await Web.EnsureCoreWebView2Async(env);
        var wv = Web.CoreWebView2;
        wv.Settings.IsGeneralAutofillEnabled = false;
        wv.Settings.IsPasswordAutosaveEnabled = false;
        wv.Settings.AreDefaultContextMenusEnabled = true;
        wv.Settings.AreDevToolsEnabled = true;
        // SmartScreen desligado pela CONFIGURACAO oficial da API, alem da flag de linha de comando
        // --disable-features=msSmartScreenProtection (AmbienteWebView2.Flags): fix da revisao 1, achado 2. As
        // DUAS formas juntas sao a base do aviso dispensado na Secao 8 da licenca do runtime (ver
        // docs/licencas/LICENCA-WebView2-Runtime.txt) — se uma mudanca futura do Chromium deixar de respeitar
        // a flag, a configuracao continua cortando. Le-se de volta no log (nao so' se atribui): e' a prova de
        // que o WebView2 aceitou o valor, e nao so' de que o C# tentou.
        wv.Settings.IsReputationCheckingRequired = false;
        _arquivos.Log($"WEBVIEW2 SmartScreen: IsReputationCheckingRequired={wv.Settings.IsReputationCheckingRequired}");
        Guarda.ServidorLocal = ServirAscDesktop;
        Guarda.Instalar(wv, _log);
        // AllFonts.js: o host o serve DIRETO do catalogo (CatalogoNoEditor — fix round 1 da Task 1 do Plano 5);
        // nada e' copiado para dentro do motor, que instalado fica em Program Files, so' leitura. Conferido
        // ANTES da navegacao e falha FECHADO (item 2 da revisao 1): sem o arquivo o modulo "allfonts" do
        // requirejs morreria e o usuario veria "Carregando documento" para sempre.
        var catalogoNoEditor = CatalogoFontes.ConferirParaOEditor(_motor);
        _arquivos.Log("motor " + catalogoNoEditor.Mensagem);
        if (!catalogoNoEditor.Ok)
        {
            UltimoErro = catalogoNoEditor.Mensagem;
            _pronto.TrySetException(new InvalidOperationException(catalogoNoEditor.Mensagem));
            return;
        }
        // Quantos arquivos o catalogo autorizou alem das raizes fixas (Task 7: sem eles uma PPTX com fonte do
        // Microsoft 365 nao abre). Zero aqui e' sinal de catalogo ausente/ilegivel, nao de maquina sem fontes.
        _arquivos.Log($"motor catalogo de fontes: {_fontes.ArquivosDoCatalogo} arquivos autorizados"
            + $" ({(catalogo.Regenerou ? $"regenerado em {catalogo.Duracao.TotalSeconds:0.0}s" : "reaproveitado")}, {catalogo.Fontes} fontes)");
        wv.SetVirtualHostNameToFolderMapping("moffice.local", Path.Combine(_motor.PastaMotor, "editors"), CoreWebView2HostResourceAccessKind.Allow);
        // Host da pagina de teste do PoC (Ponte/ponte-teste.html), usado pelo teste de fumaca; esta na lista
        // de origens permitidas da guarda (GuardaEgresso.OrigensPermitidas), senao ela mesma seria bloqueada.
        wv.SetVirtualHostNameToFolderMapping("moffice-poc.local", Path.Combine(AppContext.BaseDirectory, "Ponte"), CoreWebView2HostResourceAccessKind.Allow);
        // "arquivos" (LerBase64/GravarBase64/dialogos = I/O de arquivo arbitrario) NAO vai para a janela por
        // padrao (achado 1 da revisao 2026-09-17, fix review 2): o quadro do editor e' same-origin com este
        // documento de nivel superior (window.frames[0] alcanca DesktopOfflineAppDocumentEndLoad nele — ver
        // ChamarEndLoadAsync), entao um objeto de host adicionado AQUI seria alcancavel de dentro do quadro
        // via window.parent/window.top mesmo sem estar exposto diretamente ao quadro (a defesa anterior, que
        // so tirava "arquivos" do FrameCreated, nao fechava essa rota). Por isso o host nao entrega I/O de
        // arquivo a conteudo web por padrao: quem precisa de "arquivos" (so a pagina de teste do PoC,
        // Ponte/ponte-teste.html) chama <see cref="ExporArquivosParaTeste"/> explicitamente. "log" continua
        // saindo sempre — e' o unico objeto de host que tambem vai para o quadro (ver ExporObjetosAoQuadro).
        wv.AddHostObjectToScript("log", _ponteLog);
        // O editor vive num iframe e o AddHostObjectToScript da JANELA nao chega la (0x80070490, medido na
        // sonda da Task 4): sem isto a ponte simplesmente nao existe para o editor.
        wv.FrameCreated += (_, e) => ExporObjetosAoQuadro(e.Frame);
        // Arquivos soltos do Explorer (Task 6 do Plano 3): o documento de cima tambem tem o ouvinte de drop
        // (ponte-desktop.js roda nos dois), e a mensagem dele chega AQUI, nao no evento do quadro.
        wv.WebMessageReceived += (_, e) => AoMensagemDaPagina(e);
        // Tema do editor (Task 2 do Marco 1 / Plano 2). Registrado ANTES do ponte-desktop.js de proposito: os
        // scripts de DocumentCreated rodam na ordem em que foram adicionados, e e' o ponte-desktop.js que monta
        // o window.RendererProcessVariable.theme a partir deste window.__moffice.tema — se viesse depois, o
        // objeto ainda nao existiria e o tema cairia no padrao claro. `Tema` e' `init`, entao um registro so'
        // vale para a vida inteira da janela.
        await wv.AddScriptToExecuteOnDocumentCreatedAsync(
            "window.__moffice = window.__moffice || {}; window.__moffice.tema = "
            + System.Text.Json.JsonSerializer.Serialize(new { id = Tema.Id(), type = Tema.Tipo() }) + ";");
        // Reescrita do <script src> do AllFonts.js para o host do catalogo (CatalogoNoEditor): registrada antes
        // de qualquer script da pagina rodar, em todo documento (janela de cima e quadro do editor).
        await wv.AddScriptToExecuteOnDocumentCreatedAsync(CatalogoNoEditor.ScriptReescrita);
        // Ordem importa: ponte-desktop.js define o objeto real; grampo.js embrulha e loga o que for chamado.
        await wv.AddScriptToExecuteOnDocumentCreatedAsync(await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "Ponte", "ponte-desktop.js")));
        // Grampo opt-in (Task 1 do Marco 1 / Plano 2, backlog 15): logar TODA chamada da ponte e' diagnostico,
        // nao comportamento padrao do produto. Lido aqui (nao no construtor) porque quem liga o grampo — a
        // Sonda por --grampo — faz isso ANTES de Show(), mas depois de a janela ja existir.
        if (GrampoLigado)
            await wv.AddScriptToExecuteOnDocumentCreatedAsync(await File.ReadAllTextAsync(Path.Combine(AppContext.BaseDirectory, "Ponte", "grampo.js")));
        if (InstrumentarRede) await InstalarContagemDeMidiaAsync(wv);
        _pronto.TrySetResult(wv);
    }

    /// <summary>Conta as respostas 200 das midias do documento (<c>https://moffice-doc.local/media/…</c>) pelo
    /// DevTools Protocol. E' a unica observacao possivel desse trafego: o que o mapeamento de host virtual
    /// serve nao passa pelo <c>WebResourceRequested</c> (ver <see cref="GuardaEgresso"/>).</summary>
    private async Task InstalarContagemDeMidiaAsync(CoreWebView2 wv)
    {
        await wv.CallDevToolsProtocolMethodAsync("Network.enable", "{}");
        _receptorRede = wv.GetDevToolsProtocolEventReceiver("Network.responseReceived");
        _receptorRede.DevToolsProtocolEventReceived += (_, e) =>
        {
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(e.ParameterObjectAsJson);
                if (!doc.RootElement.TryGetProperty("response", out var resposta)) return;
                var url = resposta.TryGetProperty("url", out var u) ? u.GetString() : null;
                var status = resposta.TryGetProperty("status", out var s) ? s.GetInt32() : 0;
                if (status == 200 && Midia.EhMidiaDoDocumento(url))
                {
                    Interlocked.Increment(ref _midiasServidas);
                    _arquivos.Log($"MIDIA servida {status} {url}");
                }
            }
            catch (System.Text.Json.JsonException) { /* instrumentacao: um evento ilegivel nao derruba a janela */ }
        };
    }

    /// <summary>Quantas midias do documento o editor pediu e receberam 200 (ver <see cref="InstrumentarRede"/>).
    /// Zero com um documento que TEM figura significa que a base das midias continua sendo <c>file://</c>.</summary>
    public async Task<int> ContarMidiasServidasAsync()
    {
        await ProntoAsync();
        if (_receptorRede is null)
            throw new InvalidOperationException("ContarMidiasServidasAsync exige InstrumentarRede = true ANTES de Show()");
        return Volatile.Read(ref _midiasServidas);
    }

    /// <summary>Captura PNG da janela (o mesmo <c>CapturePreviewAsync</c> que a sonda usava, promovido ao
    /// nucleo). E' o que prova que uma FIGURA aparece: nenhuma API do editor diz que os bytes da imagem
    /// chegaram — na Task 5A o desenho existia no modelo e a tela mostrava o quadro vazio.</summary>
    public async Task<byte[]> CapturarPngAsync()
    {
        var wv = await ProntoAsync();
        using var memoria = new MemoryStream();
        await wv.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, memoria);
        return memoria.ToArray();
    }

    /// <summary>Somente para o teste de fumaça da ponte de arquivos; nunca chamar no app. `internal` (e não
    /// `public`): expõe I/O de arquivo arbitrário à página, então a superfície fica restrita ao assembly e ao
    /// projeto de teste (InternalsVisibleTo), em vez de ser parte da API do host.</summary>
    internal void ExporArquivosParaTeste() => Web.CoreWebView2.AddHostObjectToScript("arquivos", _arquivos);

    /// <summary>
    /// O quadro do editor roda documento de terceiros (parsing de DOCX/ODT/etc arbitrario vindo de fora) e por
    /// isso NAO recebe <see cref="PonteArquivos"/> (achado 1 da revisao 2026-09-17: LerBase64/GravarBase64/
    /// dialogos sao I/O de arquivo arbitrario — expor isso ao quadro seria leitura/escrita arbitraria de
    /// arquivo a partir de conteudo nao confiavel). Só "log" (<see cref="PonteLog"/>) e "ponte"
    /// (<see cref="PonteDesktop"/>, que nao tem primitiva nenhuma de arquivo: SessaoJson/SalvarChanges/
    /// MarcarModificado/ImagemUrl/Avisar so leem/escrevem dentro da pasta de recovery) vao para o quadro.
    /// </summary>
    private void ExporObjetosAoQuadro(CoreWebView2Frame quadro)
    {
        try
        {
            quadro.AddHostObjectToScript("log", _ponteLog, OrigensPonte);
            if (_ponte is not null) quadro.AddHostObjectToScript("ponte", _ponte, OrigensPonte);
            // O postMessage de um iframe chega no evento do QUADRO (nao no da janela): e' daqui que vem o
            // "moffice:soltar" do ouvinte de drop do editor (Task 6 do Plano 3).
            quadro.WebMessageReceived += (_, e) => AoMensagemDaPagina(e);
        }
        catch (Exception ex) { _arquivos.Log("PONTE falha ao expor objetos ao quadro: " + ex.Message); }
    }

    // -----------------------------------------------------------------------------------------------------
    // Area de transferencia pelos botoes e arquivos soltos do Explorer (Task 6 do Plano 3)
    // -----------------------------------------------------------------------------------------------------

    /// <summary>Um documento dos tres apps foi solto no editor (<see cref="ArquivosSoltos"/>): a casca abre em
    /// janela nova, como a boas-vindas faz com o que e' solto nela. Chega na thread da interface.</summary>
    public event Action<string>? AbrirPedido;

    /// <summary>Imagens aceitas ao soltar: o mesmo conjunto do filtro do dialogo "Imagem do arquivo"
    /// (<see cref="EscolherImagemNoWindows"/>) — soltar e' escolher o arquivo com o mouse.</summary>
    internal static readonly IReadOnlySet<string> ExtensoesImagemSolta = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg" };

    /// <summary>Documentos que o nucleo sabe abrir (<see cref="FormatoDocumentoExtensoes.DaExtensao"/>, sem o PDF,
    /// que e' so' exportacao). <c>.csv</c>/<c>.txt</c> do brief ficam de fora: o conversor nao os abre (divergencia
    /// registrada no relatorio da Task 6).</summary>
    internal static readonly IReadOnlySet<string> ExtensoesDocumentoSolto = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { ".docx", ".odt", ".xlsx", ".ods", ".pptx", ".odp" };

    /// <summary>Mensagem da pagina (documento de cima ou quadro). So' "moffice:soltar" interessa, e so' vinda da
    /// origem do editor. Os caminhos vem de <see cref="CoreWebView2File.Path"/>, que o NAVEGADOR preenche a
    /// partir do arquivo realmente solto — a pagina nao tem como forjar um (File montado em JS nao tem caminho).</summary>
    private void AoMensagemDaPagina(CoreWebView2WebMessageReceivedEventArgs e)
    {
        string? mensagem;
        try { mensagem = e.TryGetWebMessageAsString(); }
        catch (ArgumentException) { return; }
        if (mensagem != "moffice:soltar") return;
        if (!Uri.TryCreate(e.Source, UriKind.Absolute, out var origem) ||
            !OrigensPonte.Contains(origem.GetLeftPart(UriPartial.Authority), StringComparer.OrdinalIgnoreCase))
        {
            _arquivos.Log("SOLTAR mensagem recusada (origem " + e.Source + ")");
            return;
        }
        var caminhos = (e.AdditionalObjects?.OfType<CoreWebView2File>().Select(f => f.Path) ?? [])
            .Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        _ = ArquivosSoltos(caminhos);
    }

    /// <summary>
    /// O que fazer com arquivos soltos no editor (o <c>postMessageWithAdditionalObjects</c> do ouvinte de drop de
    /// <c>ponte-desktop.js</c> entrega os caminhos). Imagem: consentimento de UM uso para aquele caminho
    /// (<see cref="PonteDesktop.PermitirArquivoEscolhido"/>, o mesmo do dialogo — soltar e' uma escolha do
    /// usuario) e insercao pelo MESMO caminho do "Imagem do arquivo" (<c>asc_addImage</c> -&gt;
    /// <c>OpenFilenameDialog</c> respondido na hora com o caminho -&gt; <c>LocalFileGetImageUrl</c>). So' a
    /// PRIMEIRA imagem entra, como no ondrop do app oficial (sdk-all.js: <c>…LocalFileGetImageUrl(f[e])…;break</c>).
    /// Documento dos tres apps: <see cref="AbrirPedido"/>. Qualquer outra coisa: ignorada, com log.
    /// `internal`: os testes chamam direto (nao ha' como fazer um arrastar do sistema num teste automatizado).
    /// </summary>
    internal async Task ArquivosSoltos(IReadOnlyList<string> caminhos)
    {
        await Task.Yield();   // nunca dentro do evento do WebView2 (a insercao chama o WebView de volta)
        var imagemInserida = false;
        foreach (var caminho in caminhos)
        {
            var ext = Path.GetExtension(caminho);
            if (ExtensoesImagemSolta.Contains(ext))
            {
                if (imagemInserida) { _arquivos.Log("SOLTAR imagem extra ignorada (so' a primeira entra): " + Path.GetFileName(caminho)); continue; }
                imagemInserida = true;
                await InserirImagemSoltaAsync(caminho);
            }
            else if (ExtensoesDocumentoSolto.Contains(ext))
            {
                _arquivos.Log("SOLTAR documento -> abrir em janela nova: " + Path.GetFileName(caminho));
                try { AbrirPedido?.Invoke(caminho); }
                catch (Exception ex) { _arquivos.Log("SOLTAR erro no pedido de abertura: " + ex.Message); }
            }
            else _arquivos.Log("SOLTAR ignorado (formato nao tratado): " + Path.GetFileName(caminho));
        }
    }

    private async Task InserirImagemSoltaAsync(string caminho)
    {
        if (_ponte is null || _sessao is null) { _arquivos.Log("SOLTAR imagem sem documento aberto: ignorada"); return; }
        _ponte.PermitirArquivoEscolhido(caminho);
        var js = """
            (function () {
              try {
                var w = window.frames[0] || window;   // o editor vive no quadro (DesktopOfflineAppDocumentEndLoad idem)
                if (!w || !w.__moffice || typeof w.__moffice.inserirImagemSolta !== "function") return "sem ponte no quadro";
                return w.__moffice.inserirImagemSolta(CAMINHO);
              } catch (e) { return "erro: " + e; }
            })()
            """.Replace("CAMINHO", System.Text.Json.JsonSerializer.Serialize(caminho));
        var r = "";
        try
        {
            r = (await ExecutarJsAsync(js)).Trim('"');
            _arquivos.Log("SOLTAR imagem " + Path.GetFileName(caminho) + ": " + r);
        }
        catch (Exception ex) { _arquivos.Log("SOLTAR imagem falhou: " + ex.Message); }
        finally
        {
            // O consentimento vale para ESTA insercao e nada mais (fix round 1 da revisao da Task 6). A insercao e'
            // sincrona dentro do script (asc_addImage -> OpenFilenameDialog -> callback -> LocalFileGetImageUrl),
            // entao aqui ela ja' terminou: com "ok" a entrada ja' foi consumida (revogar e' no-op) — a menos que a
            // leitura tenha falhado e o Resolver a tenha devolvido ("falha nao consome"); sem "ok" (editor nao
            // pediu o dialogo, erro, script falhou) ela nunca foi usada. Nos dois casos ela nao pode sobrar para
            // a pagina pedir o mesmo arquivo depois, sem um novo soltar.
            if (_ponte.RevogarArquivoEscolhido(caminho))
                _arquivos.Log("SOLTAR consentimento revogado sem uso (" + (r.Length == 0 ? "falha" : r) + "): " + Path.GetFileName(caminho));
        }
    }

    /// <summary>Botao "Colar" do editor (e o item do menu de contexto): um Ctrl+V REAL ao quadro com foco, pelo
    /// DevTools Protocol (<c>Input.dispatchKeyEvent</c> com o comando de edicao "paste"). E' o mesmo caminho do
    /// teclado — o navegador le' a area de transferencia do Windows e entrega o evento <c>paste</c> que o sdkjs
    /// ja' trata (texto, HTML do Word, imagem). <c>execCommand("paste")</c> e' bloqueado em conteudo web. O CDP
    /// e' local ao processo do navegador: nao e' rede.</summary>
    public Task ColarAsync() => EnviarAtalhoAsync("v", "KeyV", 0x56, "paste");

    /// <summary>Reserva de Copiar/Recortar quando o <c>execCommand</c> da pagina nao pode (sem ativacao do
    /// usuario): o mesmo atalho real que <see cref="ColarAsync"/> usa.</summary>
    internal Task CopiarAsync() => EnviarAtalhoAsync("c", "KeyC", 0x43, "copy");

    internal Task RecortarAsync() => EnviarAtalhoAsync("x", "KeyX", 0x58, "cut");

    private async Task EnviarAtalhoAsync(string tecla, string codigo, int vk, string comando)
    {
        var wv = await ProntoAsync();
        // Foco de volta ao WebView2 (o clique pode ter sido num menu do web-apps, mas o WPF pode ter o foco em
        // outro controle da janela, como a faixa de aviso).
        Web.Focus();
        string Evento(string tipo) => System.Text.Json.JsonSerializer.Serialize(new
        {
            type = tipo,
            modifiers = 2,                   // Ctrl
            windowsVirtualKeyCode = vk,
            nativeVirtualKeyCode = vk,
            key = tecla,
            code = codigo,
            commands = tipo == "rawKeyDown" ? new[] { comando } : [],
        });
        await wv.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", Evento("rawKeyDown"));
        await wv.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", Evento("keyUp"));
    }

    /// <summary>Pedido do editor (tipos "colar"/"copiar"/"recortar" do <see cref="Avisar"/>): fora da chamada
    /// sincrona do JS, e nenhuma falha escapa (e' fogo-e-esquece).</summary>
    private async Task AtalhoPedidoDoEditorAsync(Func<Task> atalho, string nome)
    {
        await Task.Yield();
        // So' com a janela ATIVA (fix round 1 da revisao da Task 6): o pedido vem de um clique do usuario nos
        // botoes do editor, que ativa a janela. Janela em segundo plano pedindo Ctrl+V/C/X e' pagina agindo
        // sozinha: o atalho sintetico nao sai (e o Copy/Cut da pagina, sem foco, tambem nao apagou nada).
        if (!IsActive) { _arquivos.Log("AREA DE TRANSFERENCIA " + nome + " ignorado: a janela nao esta' ativa"); return; }
        try { await atalho(); }
        catch (Exception ex) { _arquivos.Log("AREA DE TRANSFERENCIA " + nome + " falhou: " + ex.Message); }
    }

    /// <summary>Serve ascdesktop://fonts/&lt;caminho&gt; do disco, com contencao real (ver <see cref="ServidorAscDesktop"/>).
    /// Qualquer outro caminho ascdesktop:// (a guarda so conhece o esquema inteiro, nao so "fonts/") devolve 403
    /// explicito em vez de cair sem resposta — sem isto o log registrava "PERMITIDA" e a requisicao ficava sem
    /// tratamento nenhum, uma leitura desonesta do que realmente aconteceu (achado 4 da revisao 2026-09-17).</summary>
    private CoreWebView2WebResourceResponse? ServirAscDesktop(CoreWebView2 wv, string uri)
    {
        if (CatalogoNoEditor.EhDoHost(uri)) return ServirCatalogo(wv, uri);
        if (!uri.StartsWith("ascdesktop://", StringComparison.OrdinalIgnoreCase)) return null;
        if (!uri.StartsWith(ServidorAscDesktop.PrefixoFontes, StringComparison.OrdinalIgnoreCase))
        {
            _arquivos.Log("ASCDESKTOP DESCONHECIDO " + uri);
            return wv.Environment.CreateWebResourceResponse(null, 403, "Caminho ascdesktop:// nao reconhecido", "Access-Control-Allow-Origin: *");
        }
        var arquivo = _fontes.ResolverArquivo(uri);
        if (arquivo is null)
        {
            _arquivos.Log("FONTE RECUSADA " + uri);
            return wv.Environment.CreateWebResourceResponse(null, 403, "Fora das pastas de fontes permitidas", "Access-Control-Allow-Origin: *");
        }
        // A leitura do disco vai dentro de try/catch (achado da revisao final 2026-09-17): isto roda no
        // WebResourceRequested, ou seja NA THREAD DA INTERFACE. Entre o File.Exists de ResolverArquivo e o
        // ReadAllBytes daqui o arquivo pode ter sido apagado, travado por outro processo ou perdido a
        // permissao — e uma excecao aqui nao vira 500, vira excecao nao tratada no manipulador de evento e
        // derruba a janela. O modo de falha certo e' o mesmo de "fonte recusada": 403 + log.
        byte[] bytes;
        try { bytes = File.ReadAllBytes(arquivo); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _arquivos.Log("FONTE ERRO DE LEITURA " + Path.GetFileName(arquivo) + ": " + ex.Message);
            return wv.Environment.CreateWebResourceResponse(null, 403, "Nao foi possivel ler o arquivo de fonte", "Access-Control-Allow-Origin: *");
        }
        _arquivos.Log("FONTE 200 " + Path.GetFileName(arquivo));
        var conteudo = new MemoryStream(bytes);
        var tipo = Path.GetExtension(arquivo).ToLowerInvariant() switch { ".png" => "image/png", ".js" => "text/javascript", _ => "application/octet-stream" };
        return wv.Environment.CreateWebResourceResponse(conteudo, 200, "OK", $"Content-Type: {tipo}\nAccess-Control-Allow-Origin: *\nCache-Control: no-store");
    }

    /// <summary>Host do catalogo (<see cref="CatalogoNoEditor"/>): so' a URL exata do AllFonts.js recebe 200, com o
    /// arquivo do catalogo desta maquina; qualquer outro caminho do host recebe 404 (nunca cai na rede). Roda
    /// na thread da interface: leitura em try/catch, falha vira 403 + log (mesmo modo de falha das fontes).</summary>
    private CoreWebView2WebResourceResponse ServirCatalogo(CoreWebView2 wv, string uri)
    {
        var arquivo = CatalogoNoEditor.Resolver(uri, _motor.PastaFontes);
        if (arquivo is null)
        {
            _arquivos.Log("CATALOGO RECUSADO " + uri);
            return wv.Environment.CreateWebResourceResponse(null, 404, "Nao encontrado", "Access-Control-Allow-Origin: *");
        }
        byte[] bytes;
        try { bytes = File.ReadAllBytes(arquivo); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _arquivos.Log("CATALOGO ERRO DE LEITURA " + arquivo + ": " + ex.Message);
            return wv.Environment.CreateWebResourceResponse(null, 403, "Nao foi possivel ler o catalogo de fontes", "Access-Control-Allow-Origin: *");
        }
        _arquivos.Log($"CATALOGO 200 AllFonts.js ({bytes.Length} bytes)");
        return wv.Environment.CreateWebResourceResponse(new MemoryStream(bytes), 200, "OK",
            "Content-Type: text/javascript; charset=utf-8\nAccess-Control-Allow-Origin: *\nCache-Control: no-store");
    }

    public Task<CoreWebView2> ProntoAsync() => _pronto.Task;

    public async Task NavegarAsync(Uri uri)
    {
        var wv = await ProntoAsync();
        var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        void Fim(object? s, CoreWebView2NavigationCompletedEventArgs e) { wv.NavigationCompleted -= Fim; tcs.TrySetResult(); }
        wv.NavigationCompleted += Fim;
        wv.Navigate(uri.ToString());
        await tcs.Task;
    }

    public async Task<string> ExecutarJsAsync(string js) => await (await ProntoAsync()).ExecuteScriptAsync(js);

    /// <summary>
    /// Abre um documento: converte com o x2t (Editor.bin na pasta de recovery) EM PARALELO com a navegacao,
    /// como o app oficial, e responde ao LocalStartOpen() do editor com
    /// DesktopOfflineAppDocumentEndLoad(pasta, base64, tamanho).
    /// </summary>
    /// <summary><c>async</c> (e nao uma expressao que devolve a Task de <see cref="AbrirSessaoAsync"/>): com a
    /// forma de expressao, a <see cref="ErroConversao"/> de uma extensao desconhecida era lancada de forma
    /// SINCRONA, antes de qualquer Task existir — quem chamasse `_ = j.AbrirAsync(x)` ou guardasse a Task para
    /// tratar depois levava a excecao na chamada, nao no `await` (achado 11 da revisao 1).</summary>
    public async Task AbrirAsync(string caminho) =>
        await AbrirSessaoAsync(caminho, FormatoDocumentoExtensoes.DaExtensao(caminho), novo: false);

    /// <summary>
    /// "Novo documento" (Task 2 do Marco 1 / Plano 2): copia o modelo vazio que o proprio motor traz
    /// (<see cref="ModelosVazios"/>, <c>converter/empty/pt-BR</c> com reserva em <c>default</c>) para a pasta
    /// de recuperacao desta sessao e abre a copia como documento NOVO — sem arquivo de origem do usuario,
    /// titulo generico ("Documento"/"Planilha"/"Apresentação") e "Salvar" que vira "Salvar como".
    ///
    /// Os formatos ODF nao tem modelo proprio no motor: usam o irmao OOXML (<c>Odt</c> -&gt; <c>new.docx</c>) e
    /// a sessao GUARDA <c>Formato = Odt</c>, para que o primeiro "Salvar como" ja sugira o filtro ODF e o x2t
    /// converta na gravacao.
    /// </summary>
    public async Task NovoAsync(FormatoDocumento formato)
    {
        var modelo = ModelosVazios.Caminho(_motor, formato);
        await AbrirSessaoAsync(modelo, formato, novo: true);
    }

    /// <summary>URL de entrada do web-apps no modo desktop. Funcao PURA (testavel sem janela): o
    /// <c>uitheme=</c> e' lido por <c>web-apps/apps/api/documents/index.html</c> (linha 112 da tag 9.4.0.129),
    /// que o transforma em <c>customization.uiTheme</c> da configuracao do editor.</summary>
    public static string MontarUrlDeEntrada(FormatoDocumento formato, string titulo, TemaEditor tema) =>
        $"https://moffice.local/web-apps/apps/api/documents/index.html?placement=desktop&doctype={formato.DocType()}"
        + $"&filetype={formato.Extensao()}&title={Uri.EscapeDataString(titulo)}&lang=pt-BR&mode=edit&userid=1&username=mOffice"
        + $"&uitheme={tema.Id()}";

    private async Task AbrirSessaoAsync(string arquivoDeEntrada, FormatoDocumento formato, bool novo)
    {
        await ProntoAsync();
        var recovery = Path.Combine(SessaoDocumento.RaizRecovery, Guid.NewGuid().ToString("N"));
        // A pasta nasce AQUI (e nao so' dentro do x2t): a trava e o sessao.json vao para ela antes da conversao.
        Directory.CreateDirectory(recovery);
        // Documento NOVO: a origem e' uma COPIA do modelo vazio dentro da propria pasta de recuperacao — nunca
        // o modelo do motor (que e' arquivo instalado, compartilhado entre todas as janelas e todos os
        // usuarios da maquina: uma gravacao por cima dele estragaria o "novo documento" de todo mundo). Como a
        // copia vive na recovery, ela some junto com a sessao (ApagarRecovery) sem deixar lixo.
        var caminho = arquivoDeEntrada;
        if (novo)
        {
            caminho = Path.Combine(recovery, Path.GetFileName(arquivoDeEntrada));
            File.Copy(arquivoDeEntrada, caminho, true);
        }
        var sessao = new SessaoDocumento { ArquivoOrigem = caminho, Formato = formato, PastaRecovery = recovery, Novo = novo };
        // Recuperacao apos queda (Task 2 do Plano 3): a trava diz "sessao viva" enquanto a janela existir (o SO
        // a solta sozinho se o processo morrer), e os metadados dizem como reabrir a pasta depois de uma queda.
        // Melhor esforco nos dois: sem eles o documento abre do mesmo jeito, so' nao e' recuperavel.
        // Carimbo do original (I2 da revisao final do Plano 3): e' o que diz, numa recuperacao, se o arquivo foi
        // mexido depois da queda. Documento novo nao tem original.
        var metadados = new MetadadosSessao(caminho, (int)formato, sessao.Titulo, novo, _app, sessao.IdUsuario);
        MarcarRecuperavel(sessao, novo ? metadados : metadados.ComCarimboDe(caminho), _arquivos.Log);
        await IniciarSessaoAsync(sessao, _conv.ParaBinAsync(caminho, recovery));
    }

    /// <summary>
    /// Trava a pasta da sessao e grava o <c>sessao.json</c> — NESTA ordem e so' com a trava na mao (fix round 1
    /// da revisao da Task 3): uma pasta viva com <c>sessao.json</c> e SEM trava pareceria "caida" para outra
    /// instancia (<c>Recuperaveis</c>), que a listaria e poderia descarta-la por baixo desta janela. Sem trava,
    /// a sessao simplesmente nao e' recuperavel (e fica registrado no log). Melhor esforco: o documento abre do
    /// mesmo jeito. `internal static` por ser a costura de teste da regra.
    /// </summary>
    internal static void MarcarRecuperavel(SessaoDocumento sessao, MetadadosSessao metadados, Action<string> log)
    {
        sessao.Trava = TravaSessao.Adquirir(sessao.PastaRecovery);
        if (sessao.Trava is null)
        {
            log("RECUPERACAO sem trava (rec.lock) em " + sessao.PastaRecovery + ": sessao.json nao gravado, este documento nao sera' recuperavel");
            return;
        }
        sessao.Metadados = metadados;
        try { metadados.Gravar(sessao.PastaRecovery); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { log("RECUPERACAO nao gravou sessao.json: " + e.Message); }
    }

    /// <summary>
    /// Reabre uma pasta de recuperacao que sobrou de uma queda (Task 2 do Plano 3), no desenho do app oficial:
    /// SEM rodar o <c>x2t</c> (o <c>Editor.bin</c> da pasta e' o documento como foi aberto) e com os changes
    /// gravados reaplicados pelo proprio editor ANTES do <c>EndLoad</c> (ver <see cref="ChamarEndLoadAsync"/>).
    /// O documento abre modificado. O <see cref="SessaoDocumento.IdUsuario"/> sobe 1 em relacao ao da sessao
    /// que caiu (ver <see cref="MetadadosSessao"/>).
    /// </summary>
    /// <exception cref="InvalidOperationException">A pasta nao tem <c>sessao.json</c> legivel ou
    /// <c>Editor.bin</c>, esta' aberta em outra janela, ou o <c>changes0.json</c> nao pode ser lido/consertado
    /// (causa em <see cref="Exception.InnerException"/>). Em todos os casos a trava fica livre.</exception>
    public async Task RecuperarAsync(string pasta)
    {
        await ProntoAsync();
        var m = MetadadosSessao.Ler(pasta)
            ?? throw new InvalidOperationException("Não foi possível recuperar o documento: os dados da sessão (sessao.json) estão ausentes ou danificados.");
        var ficticia = new SessaoDocumento { ArquivoOrigem = m.Origem, Formato = (FormatoDocumento)m.Formato, PastaRecovery = pasta };
        if (!File.Exists(ficticia.Bin))
            throw new InvalidOperationException("Não foi possível recuperar o documento: a cópia de trabalho (Editor.bin) não existe mais.");
        var trava = TravaSessao.Adquirir(pasta)
            ?? throw new InvalidOperationException("Este documento já está aberto em outra janela.");
        // Cauda rasgada por uma queda no meio da gravacao (fix round 1 da revisao): so' AGORA, com a trava na mao
        // (esta janela e' a dona da pasta), o arquivo e' consertado — o original vai para changes0.json.corrompido
        // e fica o maior prefixo valido. Sem isto os changes novos desta sessao seriam apendados atras do lixo.
        LeituraChanges changes;
        try
        {
            // Reparo e leitura numa leitura SO' do disco, que LANCA se o arquivo nao da' para ler (M1 da revisao final
            // do Plano 3): antes a leitura depois do reparo era a ChangesGravados.Ler, que engole erro de E/S e devolve
            // zero changes — a sessao abria com ChangesAbertos = 0 e o prefixo recuperado ficava truncavel.
            var (reparo, lidos) = ChangesGravados.RepararELer(ficticia.ArquivoChanges);
            if (reparo.Reparou)
                _arquivos.Log($"RECUPERACAO changes0.json danificado: mantidos {reparo.Quantidade} changes, descartados {reparo.BytesDescartados} bytes (original em {Path.GetFileName(reparo.Copia)})");
            // Lido UMA vez (fix round 1): este array e' o que vai para o editor no EndLoad.
            changes = lidos;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            // Mesma excecao documentada das outras recusas (fix round 2 da revisao), com a causa dentro — e a
            // trava solta, para a pasta continuar recuperavel numa proxima tentativa.
            trava.Dispose();
            throw new InvalidOperationException("Não foi possível recuperar o documento: o arquivo de alterações (changes0.json) não pôde ser lido ou consertado.", e);
        }
        catch
        {
            trava.Dispose();
            throw;
        }
        var n = changes.Quantidade;
        var sessao = new SessaoDocumento
        {
            ArquivoOrigem = m.Origem,
            Formato = (FormatoDocumento)m.Formato,
            PastaRecovery = pasta,
            Novo = m.Novo,
            Recuperada = true,
            ChangesAbertos = n,
            IdUsuario = m.IdUsuario + 1,
            Modificado = true,
            Trava = trava,
            // I1 da revisao final do Plano 3: o original cifrado nao pode ser regravado sem senha por esta sessao.
            OriginalProtegidoSemSenha = m.Protegido && !m.Novo ? m.Origem : null,
            // I2: o original mudou (ou sumiu) depois da queda — salvar vira "Salvar como", nunca por cima dele.
            Desvinculada = m.OrigemMudou(),
        };
        if (sessao.Desvinculada)
            _arquivos.Log("RECUPERACAO o arquivo original mudou (ou sumiu) depois da queda: a sessao salva so' com outro nome");
        // Regravado JA' com o id novo: se ESTA sessao tambem cair, a proxima recuperacao sobe mais um. O carimbo do
        // original fica o que era (o da ultima gravacao conhecida): uma nova queda continua desvinculada.
        sessao.Metadados = m with { IdUsuario = sessao.IdUsuario };
        try { sessao.Metadados.Gravar(pasta); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { _arquivos.Log("RECUPERACAO nao regravou sessao.json: " + e.Message); }
        _arquivos.Log($"RECUPERACAO reabrindo {pasta} ({n} changes, idUsuario {sessao.IdUsuario})");
        await IniciarSessaoAsync(sessao, Task.FromResult(sessao.Bin), (changes.ArrayJs, n));
    }

    /// <summary>Trecho comum de <see cref="AbrirSessaoAsync"/> e <see cref="RecuperarAsync"/>: troca a sessao
    /// da janela, prepara o mapeamento da pasta e a ponte e navega para o editor. <paramref name="binPronto"/> e'
    /// o <c>Editor.bin</c> que <see cref="ResponderAberturaAsync"/> entrega no <c>EndLoad</c> — a conversao do
    /// <c>x2t</c> em andamento, ou o arquivo que ja' esta' na pasta recuperada.</summary>
    private async Task IniciarSessaoAsync(SessaoDocumento sessao, Task<string> binPronto, (string ArrayJs, int Quantidade)? changesRecuperados = null)
    {
        var wv = await ProntoAsync();
        var recovery = sessao.PastaRecovery;
        var formato = sessao.Formato;
        var caminho = sessao.ArquivoOrigem;
        var novo = sessao.Novo;
        // Reabrir na MESMA janela (fix 2 da revisao da Task 5, critico): os tres estados de midia so' eram
        // escritos no caminho de FALHA de ChamarEndLoadAsync, entao um documento B saudavel aberto depois de um
        // A cuja troca da base falhou continuava reportando MidiasIndisponiveis == true e o UltimoErro velho
        // — os apps do Plano 2 mostrariam a barra de aviso num documento sem problema nenhum. Zerar AQUI, no
        // comeco de cada abertura, e' o mesmo espirito do `UltimoErro = null` de ConverterESinalizarAsync
        // (SalvarAsync): o estado descreve o documento ATUAL. Depois do ProntoAsync de proposito: se a
        // inicializacao da janela falhou, o UltimoErro dela (catalogo/motor) tem de sobreviver.
        MidiasIndisponiveis = false;
        TrocaDaBaseDeMidia = null;
        UltimoErro = null;
        // Outro documento na janela: a apresentacao do anterior acabou (Task 7) e nenhum pedido de senha dele vale.
        if (EmTelaCheia) DefinirTelaCheia(false);
        _opcaoAvancadaPendente = null;
        // Documento novo na janela: nada pendente do anterior (o campo e' por DOCUMENTO, nao por janela). Numa
        // sessao RECUPERADA ha', sim, edicao nao gravada desde o inicio: os changes reaplicados nunca foram
        // salvos no arquivo (ver RestaurarModificadoSeHouveEdicao).
        _edicaoNaoGravada = sessao.Recuperada;
        // Abrir um SEGUNDO documento na mesma janela (achado 4 da revisao final, fix review 2) descartava a
        // sessao anterior sem apagar a pasta de recovery dela — so' o Closed da janela chamava ApagarRecovery,
        // entao a pasta do documento anterior virava mais uma orfa, so' encontrada na proxima varredura na
        // subida do app. A pasta e' apagada DEPOIS de NavegarAsync (abaixo), nao aqui: ver o comentario la'.
        var sessaoAnterior = _sessao;
        // A anterior era uma recuperacao que nunca apareceu? Entao a pasta dela fica (so' solta a trava).
        var anteriorPendente = RecuperacaoPendente;
        _sessao = sessao;
        _recuperacaoConfirmada = false;
        _endLoadRecuperadoOk = false;
        _prontoRecuperado = false;
        // Documento novo na janela = nada modificado (achado 4 da revisao 1). Se o ANTERIOR estava modificado,
        // a casca guardou "modificado = true" no ultimo evento e nunca mais receberia um false — o asterisco
        // ficaria preso num documento recem-aberto. Ja estamos na thread da interface (AbrirSessaoAsync roda
        // no Dispatcher); o try/catch e' o mesmo dos demais eventos publicos. Generalizado na Task 2 do
        // Plano 3: a sessao recuperada ja' nasce modificada, e a casca tem de receber o `true`.
        if ((sessaoAnterior?.Modificado == true) != sessao.Modificado)
        {
            try { ModificadoAlterado?.Invoke(sessao.Modificado); }
            catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de modificado (troca de sessao): " + ex.Message); }
        }
        _binPronto = binPronto;
        _changesRecuperados = changesRecuperados;
        wv.SetVirtualHostNameToFolderMapping("moffice-doc.local", recovery, CoreWebView2HostResourceAccessKind.Allow);
        var corretor = Corretor.Compartilhado(_motor, _arquivos.Log);
        var dicionarioUsuario = _dicionarioUsuario.Value;
        // O aviso leva a sessao DONA desta ponte (fix round 2 da revisao da Task 3): um "pronto" atrasado da
        // pagina anterior (a pagina velha vive ate' o NavigationCompleted da nova) nao pode confirmar a sessao nova.
        _ponte = new PonteDesktop(_sessao, (tipo, param) => Avisar(sessao, tipo, param), _motor.PastaFontes, _arquivos.Log, AoMudarModificado,
            p => PedidoOrtografia.Responder(p, corretor.Correta, corretor.Sugestoes, dicionarioUsuario, _arquivos.Log),
            // O changes0.json nao pode ser posto em dia (fix round 2 da revisao final do Plano 3; invariante em
            // PonteDesktop.SalvarChanges): a sessao falha FECHADA. Chega DENTRO da chamada sincrona do JS: so' agenda.
            detalhe => Dispatcher.BeginInvoke(() => { if (ReferenceEquals(sessao, _sessao) && !_fechada) FalharGravacaoDeAlteracoes(detalhe); }));
        wv.AddHostObjectToScript("ponte", _ponte);
        var url = MontarUrlDeEntrada(formato, _sessao.Titulo, Tema);
        _arquivos.Log($"PONTE abrindo {Path.GetFileName(caminho)} ({formato}{(novo ? ", novo" : "")}{(sessao.Recuperada ? ", recuperado" : "")}) -> {url}");
        AtualizarTitulo();
        try
        {
            if (FalharNavegacaoParaTeste) throw new InvalidOperationException("falha simulada da navegação (teste)");
            await NavegarAsync(new Uri(url));
        }
        finally
        {
            // So' DEPOIS de a navegacao completar (fix final da revisao do branch): a pagina ANTIGA continua viva
            // ate' o NavigationCompleted da nova e ainda segura o proxy do _ponte antigo — um SalvarChanges
            // tardio dela (Directory.CreateDirectory em PonteDesktop) recriaria a pasta logo depois de apagada,
            // deixando uma orfa pequena. Melhor esforco, como o Closed: nao ha' para onde propagar uma falha.
            if (anteriorPendente) sessaoAnterior?.LiberarTrava();
            else sessaoAnterior?.ApagarRecovery();
        }
    }

    /// <summary>A ponte da sessao atual (so' para os testes guardarem a de um documento anterior). `internal`.</summary>
    internal PonteDesktop? PonteParaTeste => _ponte;

    /// <summary>Único canal JS -&gt; C# da ponte (PonteDesktop.Avisar). Roda na thread da interface, DENTRO da
    /// chamada síncrona do JS: não pode chamar o WebView de volta aqui (ver <see cref="ResponderAberturaAsync"/>).</summary>
    private void Avisar(SessaoDocumento origem, string tipo, string param)
    {
        if (tipo == "pronto" && ReferenceEquals(origem, _sessao) && origem.Recuperada)
        {
            if (IgnorarProntoParaTeste) _arquivos.Log("RECUPERACAO pronto ignorado (teste)");
            else { _prontoRecuperado = true; ConfirmarRecuperacaoSePronta(); }
        }
        // execCommand sai por um caminho proprio porque e' o unico `tipo` com volume: o web-apps manda
        // `title:mousemove` a cada movimento do mouse sobre a barra de titulo (~150/s, medido na revisao
        // final do Plano 2), e cada um virava uma linha no .log da janela — megabytes de ruido por sessao,
        // afogando as linhas que importam (FONTE RECUSADA, PONTE erro de abertura, PONTE erro ao salvar).
        if (tipo == "execCommand") { AvisarComando(param); return; }
        // Avisos que levam a SENHA do documento (Task 7 do Plano 3): o parametro NUNCA vai para o log.
        if (AvisosSigilosos.Contains(tipo)) _arquivos.Log("PONTE aviso " + tipo + " (conteudo omitido)");
        else _arquivos.Log("PONTE aviso " + tipo + " " + (param.Length > 300 ? param[..300] : param));
        if (tipo == "opcoesAvancadas") { if (ReferenceEquals(origem, _sessao)) ReceberOpcoesAvancadas(origem, param); return; }
        if (tipo == "senhaSalvar") { if (ReferenceEquals(origem, _sessao)) AtualizarSenha(origem, param); return; }
        if (tipo == "telaCheia")
        {
            if (param != "1" && IgnorarSaidaDaTelaCheiaParaTeste) { _arquivos.Log("TELA CHEIA SetFullscreen(false) ignorado (teste)"); return; }
            // Fora da chamada sincrona do JS (mexe no tamanho da janela e no foco do WebView2).
            if (ReferenceEquals(origem, _sessao)) Dispatcher.BeginInvoke(() => { if (!_fechada) DefinirTelaCheia(param == "1"); });
            return;
        }
        if (tipo == "escTelaCheia") { if (ReferenceEquals(origem, _sessao)) EscNaTelaCheia("editor"); return; }
        // A chamada SalvarChanges LANCOU no JS (fix round 3): o lote se perdeu sem o host saber. Mesma falha fechada do
        // retorno "falhou" (a ponte congela ja'; o resto sai da chamada sincrona).
        if (tipo == "falhaGravacao")
        {
            if (!ReferenceEquals(origem, _sessao)) return;
            _ponte?.Congelar();
            Dispatcher.BeginInvoke(() => { if (ReferenceEquals(origem, _sessao) && !_fechada) FalharGravacaoDeAlteracoes("a chamada ao host lancou: " + param); });
            return;
        }
        // Acoes do host pedidas pela PAGINA (M5 da revisao final do Plano 3): so' da sessao ATUAL, como os tipos
        // acima. A pagina anterior vive ate' o NavigationCompleted da nova e ainda segura a ponte antiga — um
        // Print/Colar/Copiar/Recortar dela nao pode imprimir nem mexer na area de transferencia do documento novo.
        if (tipo is "imprimir" or "colar" or "copiar" or "recortar" && !ReferenceEquals(origem, _sessao))
        {
            _arquivos.Log("PONTE aviso " + tipo + " ignorado: veio da pagina de um documento anterior");
            return;
        }
        if (tipo == "pronto") { try { DocumentoPronto?.Invoke(); } catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de pronto: " + ex.Message); } }
        else if (tipo == "localStartOpen") _ = ResponderAberturaAsync();
        else if (tipo == "abrirImagem") _ = ResponderDialogoImagemAsync();
        else if (tipo == "semSuporte" && ReferenceEquals(origem, _sessao) && MensagensSemSuporte.TryGetValue(param, out var semSuporte))
            Dispatcher.BeginInvoke(() => { if (!_fechada) AvisarUsuario(semSuporte); });
        else if (tipo == "colar") _ = AtalhoPedidoDoEditorAsync(ColarAsync, tipo);
        else if (tipo == "copiar") _ = AtalhoPedidoDoEditorAsync(CopiarAsync, tipo);
        else if (tipo == "recortar") _ = AtalhoPedidoDoEditorAsync(RecortarAsync, tipo);
        else if (tipo == "imprimir") ImpressaoEmCursoParaTeste = ImprimirPedidoDoEditorAsync(param);
        else if (tipo == "localFileSave" && FalhaDaRecuperacao is not null)
        {
            // Sessao recuperada que falhou (FalharRecuperacao): a janela esta' fechando; salvar rodaria o x2t sobre
            // Editor.bin + TODOS os changes, que nao e' o que o editor mostra. Ignorado de proposito.
            _arquivos.Log("RECUPERACAO pedido de salvar ignorado (a reaplicacao falhou; a janela esta' fechando)");
        }
        else if (tipo == "localFileSave")
        {
            // Se o C# pediu o salvamento, ele ja sabe o destino e o formato: so libera quem esta' esperando (o
            // SalvarComoAsync ja' conta no _salvamentosEmCurso). Senao, o pedido veio do editor (Ctrl+S / menu Salvar /
            // Salvar como) e quem decide o destino e' o host: conta DESDE JA' (o asc_Save acabou de zerar o Modificado)
            // e o finally de SalvarPedidoDoEditorAsync desconta uma vez.
            var pendente = Interlocked.Exchange(ref _pedidoSalvar, null);
            if (pendente is not null) pendente.TrySetResult(param);
            else
            {
                var salvamento = IniciarSalvamento();
                _ = SalvarPedidoDoEditorAsync(param, salvamento);
            }
        }
    }

    /// <summary>
    /// Comandos do editor que o nucleo NAO registra no log, nunca. Sao os de mouse da barra de titulo
    /// customizada do web-apps: `title:mousemove` chega a cada pixel de movimento (~150 linhas por segundo
    /// com o ponteiro passeando na barra) e `title:mousedown`/`title:mouseup` vem em par a cada clique. O
    /// log da janela e' diagnostico de EGRESSO e de erro; enche-lo com telemetria de mouse torna inutil o
    /// unico artefato que existe para entender uma sessao que deu errado.
    ///
    /// Lista pequena e FECHADA de proposito: o nucleo nao sabe o que a casca trata, entao o silencio vale so'
    /// para o que e' comprovadamente ruido de alta frequencia — qualquer comando novo continua aparecendo no
    /// log ate' alguem decidir o contrario. (A casca tem a lista dela, maior, em <c>AplicativoOffice</c>.)
    /// </summary>
    internal static readonly IReadOnlySet<string> ComandosSilenciosos =
        new HashSet<string>(StringComparer.Ordinal) { "title:mousemove", "title:mousedown", "title:mouseup" };

    /// <summary>Um <c>execCommand</c> do web-apps (ponte-desktop.js manda <c>{cmd, json}</c>). O nucleo nao
    /// interpreta nenhum deles: reemite na thread da interface, FORA da chamada sincrona do JS (o renderer
    /// esta' bloqueado esperando o retorno; um assinante que abrisse um dialogo aqui travaria os dois lados).
    /// O JSON e' lido UMA vez (e nao duas, como quando o log vinha antes): e' o mesmo objeto que decide se a
    /// linha vai para o log e que e' entregue a' casca.</summary>
    private void AvisarComando(string param)
    {
        var comando = ComandoDoEditor.Ler(param);
        if (comando is null) { _arquivos.Log("PONTE execCommand ilegivel: " + param); return; }
        if (!ComandosSilenciosos.Contains(comando.Comando))
            _arquivos.Log("PONTE aviso execCommand " + (param.Length > 300 ? param[..300] : param));
        Dispatcher.BeginInvoke(() =>
        {
            try { ComandoRecebido?.Invoke(comando); }
            catch (Exception ex) { _arquivos.Log("PONTE erro no tratamento de " + comando.Comando + ": " + ex.Message); }
        });
    }

    /// <summary>
    /// Houve edicao do usuario que AINDA NAO foi gravada em nenhum arquivo de DOCUMENTO. E' diferente de
    /// <see cref="SessaoDocumento.Modificado"/>, que e' o espelho do estado do EDITOR — e o editor zera o
    /// dele cedo demais para o host confiar so' nele.
    ///
    /// O porque (achado Critico de uma revisao, medido na janela real ao salvar e ao exportar para PDF): o <c>asc_Save</c> dispara
    /// <c>onDocumentModifiedChanged(false)</c> — que chega ao host por <see cref="PonteDesktop.MarcarModificado"/>
    /// e zera <c>Modificado</c> — ANTES de mandar o <c>LocalFileSave</c> que faz o host gravar de fato. Ou
    /// seja: quando a gravacao FALHA, ou quando ela e' uma EXPORTACAO (PDF, que nao guarda o documento
    /// editavel), o editor ja' apagou o sinal de "tem alteracao pendente" — o titulo perde o asterisco, o
    /// botao Salvar da barra desliga e, o pior, o proximo <c>Close()</c> nao pergunta nada e apaga a pasta de
    /// recuperacao. As alteracoes do usuario somem sem uma palavra.
    ///
    /// Este campo e' o sinal do HOST, e so' um salvamento de DOCUMENTO bem-sucedido o zera. Sobre ele,
    /// <see cref="RestaurarModificadoSeHouveEdicao"/> devolve o <c>Modificado</c> que o editor zerou cedo.
    ///
    /// Lado escolhido do erro: a anotacao pode SUPER-relatar (editar, desfazer ate' o estado salvo — o
    /// editor manda <c>false</c> — e depois exportar: volta um asterisco fantasma), nunca SUB-relatar. Um
    /// asterisco a mais custa uma pergunta ao fechar; um a menos custa o documento.
    /// </summary>
    private bool _edicaoNaoGravada;

    /// <summary>
    /// Um salvamento esta' em curso (fix round 2 da revisao final do Plano 3, C1): ligado quando o <c>LocalFileSave</c>
    /// chega e quando <see cref="SalvarComoAsync"/> comeca; desligado no fim de <see cref="ConverterESinalizarAsync"/>
    /// e de quem o chama. E' a janela em que o editor ja' zerou <see cref="SessaoDocumento.Modificado"/> e o arquivo do
    /// usuario ainda nao foi gravado: um encerramento do Windows NESSA hora tem de manter a pasta de recuperacao.
    /// </summary>
    private int _salvamentosEmCurso;

    /// <summary>
    /// Um salvamento comecou (fix round 4 da revisao final do Plano 3): o argumento e' a Task que termina quando ele
    /// terminar (com sucesso ou nao — nunca falha). A casca a registra no dreno do encerramento
    /// (<c>ImpressoesEmCurso.AcompanharAsync</c>): fechar a ULTIMA janela no meio de um Ctrl+S nao pode encerrar o app
    /// antes de o arquivo ser gravado. Dispara na thread da interface, SINCRONO no inicio do salvamento.
    /// </summary>
    public event Action<Task>? SalvamentoIniciado;

    /// <summary>Conta o salvamento (<see cref="_salvamentosEmCurso"/>) e o anuncia (<see cref="SalvamentoIniciado"/>).
    /// Quem chama TEM de chamar <see cref="TerminarSalvamento"/> uma vez, num finally.</summary>
    private Salvamento IniciarSalvamento()
    {
        Interlocked.Increment(ref _salvamentosEmCurso);
        var salvamento = new Salvamento(new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously),
            Interlocked.Read(ref _avisosDeModificado));
        try { SalvamentoIniciado?.Invoke(salvamento.Fim.Task); }
        catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de salvamento iniciado: " + ex.Message); }
        return salvamento;
    }

    private void TerminarSalvamento(Salvamento salvamento)
    {
        Interlocked.Decrement(ref _salvamentosEmCurso);
        salvamento.Fim.TrySetResult();
    }

    /// <summary>Um salvamento em curso: o sinal de fim (<see cref="SalvamentoIniciado"/>) e quantos "modificado" o editor
    /// ja' tinha avisado quando ele comecou (fix round 5: um aviso depois disso e' edicao DURANTE o salvamento).</summary>
    private sealed record Salvamento(TaskCompletionSource Fim, long AvisosDeModificadoNoInicio);

    /// <summary>Quantas vezes o editor avisou "modificado = true" (fix round 5). O asc_Save avisa "false" ANTES do
    /// LocalFileSave: um "true" depois do inicio do salvamento e' o usuario digitando enquanto ele corre.</summary>
    private long _avisosDeModificado;

    /// <summary>Ha' pelo menos um salvamento em curso (fix round 3: contador, nao bool — dois salvamentos ao mesmo
    /// tempo nao se desligam um ao outro).</summary>
    private bool SalvamentoEmCurso => Volatile.Read(ref _salvamentosEmCurso) > 0;

    /// <summary>Callback que a <see cref="PonteDesktop"/> chama quando <see cref="SessaoDocumento.Modificado"/>
    /// MUDA (ja depois de gravar o novo valor). Vem da thread do JS; o evento publico sai na thread da
    /// interface, que e' onde a casca desenha o asterisco.
    ///
    /// O <c>true</c> e' anotado em <see cref="_edicaoNaoGravada"/> de forma SINCRONA (fora do
    /// <c>BeginInvoke</c>): o sinal do editor tem de ficar registrado no instante em que chega, nao quando a
    /// fila do Dispatcher o alcancar. O <c>false</c> NAO apaga a anotacao — quem apaga e' o salvamento de
    /// documento que deu certo, em <see cref="ConverterESinalizarAsync"/>.</summary>
    private void AoMudarModificado(bool modificado)
    {
        if (modificado) { _edicaoNaoGravada = true; Interlocked.Increment(ref _avisosDeModificado); }
        Dispatcher.BeginInvoke(() =>
        {
            AtualizarTitulo();
            try { ModificadoAlterado?.Invoke(modificado); }
            catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de modificado: " + ex.Message); }
        });
    }

    /// <summary>
    /// Devolve o <c>Modificado</c> que o <c>asc_Save</c> zerou cedo demais, quando a gravacao NAO guardou o
    /// documento editavel (exportacao para PDF, ou falha). Sem isto, os dois caminhos perdem o documento:
    ///
    /// <list type="bullet">
    /// <item>documento novo -&gt; "Salvar" na barra -&gt; usuario escolhe PDF -&gt; fecha: o fechamento nao
    /// pergunta nada (acha que nao ha alteracao) e o <c>Closed</c> apaga a recuperacao;</item>
    /// <item>"Salvar" ao fechar falha -&gt; a janela fica aberta (correto) -&gt; o usuario fecha de novo:
    /// desta vez nao pergunta nada e descarta.</item>
    /// </list>
    ///
    /// Roda na thread da interface (todos os chamadores de <see cref="ConverterESinalizarAsync"/> ja' estao
    /// nela — e' de la' que <see cref="AtualizarTitulo"/> e <see cref="ModificadoAlterado"/> ja' saiam). O
    /// try/catch do evento e' o dos demais eventos publicos, e pelo mesmo motivo: este metodo e' chamado
    /// DEPOIS de <c>EndSave</c> ja' ter sido respondido ao editor, entao uma excecao de assinante subindo
    /// daqui viraria um <c>EndSave(2)</c> indevido no <c>catch</c> de <see cref="SalvarPedidoDoEditorAsync"/>.
    /// </summary>
    private void RestaurarModificadoSeHouveEdicao()
    {
        var s = _sessao;
        if (s is null || !_edicaoNaoGravada || s.Modificado) return;
        s.Modificado = true;
        AtualizarTitulo();
        _arquivos.Log("PONTE alteracao pendente restaurada (o salvamento nao gravou o documento editavel)");
        try { ModificadoAlterado?.Invoke(true); }
        catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de modificado (restauracao): " + ex.Message); }
    }

    private async Task ResponderAberturaAsync()
    {
        // Devolve o controle ao JS ANTES de tocar no WebView: o LocalStartOpen chega dentro de uma chamada
        // sincrona de objeto de host, com o renderer bloqueado esperando o retorno.
        await Task.Yield();
        if (_sessao is null || _binPronto is null) return;
        if (SuprimirEndLoadParaTeste) { _arquivos.Log("PONTE EndLoad suprimido (teste)"); return; }
        // O try cobre TUDO ate o EndLoad de sucesso (achado 3 da revisao 2026-09-17): antes, só a conversao
        // (x2t) tinha rede de seguranca; um erro em ReadAllBytesAsync ou no proprio ChamarEndLoadAsync
        // (ex.: JSON malformado, WebView fechado) virava excecao nao observada numa Task fire-and-forget
        // (Avisar chama isto com `_ = ResponderAberturaAsync()`) e o editor ficava preso em "Carregando
        // documento" para sempre, sem log nenhum explicando por que.
        var sessaoAberta = _sessao;
        try
        {
            string bin;
            try { bin = await _binPronto; }
            catch (ErroConversao ec) when (PedeSenha(ec.CodigoSaida) && !sessaoAberta.Recuperada)
            {
                // Arquivo protegido por senha (Task 7 do Plano 3): nao e' erro de abertura. O editor mostra o proprio
                // dialogo de senha e responde por SetAdvancedOptions (ReceberOpcoesAvancadas), que reconverte e volta
                // aqui. Uma sessao recuperada nunca passa por aqui: o Editor.bin dela ja' esta' decifrado (sem x2t).
                if (!ReferenceEquals(sessaoAberta, _sessao)) return;
                sessaoAberta.Senha = null;   // 91: a tentativa anterior estava errada
                await PedirSenhaAoEditorAsync(ec.CodigoSaida);
                return;
            }
            // A senha digitada abriu o arquivo: a sessao passa a ser "protegida" no sessao.json (a senha, nunca).
            if (sessaoAberta.Senha is not null) MarcarProtegido(sessaoAberta, true);
            var bytes = await File.ReadAllBytesAsync(bin);
            // O JS prefixa file:// nesta pasta (sdk-all.js: `0 != tia.indexOf("file:")`), e a midia sai como
            // file://<pasta>/media/... — ver a pergunta em aberto no relatorio da Task 5A.
            // Sessao recuperada (Task 2 do Plano 3): os changes gravados vao junto, aplicados pelo editor antes do
            // EndLoad — os MESMOS que RecuperarAsync leu (uma leitura so', fix round 1 da revisao): a contagem
            // do array e ChangesAbertos sao a mesma leitura, e o truncamento relativo depende disso.
            var changes = _sessao.Recuperada ? _changesRecuperados : null;
            if (changes is { } c) _arquivos.Log($"RECUPERACAO aplicando {c.Quantidade} changes ({c.ArrayJs.Length} caracteres) antes do EndLoad");
            await ChamarEndLoadAsync(_sessao.PastaRecovery.Replace('\\', '/'), Convert.ToBase64String(bytes), bytes.Length, changes);
            _arquivos.Log($"PONTE EndLoad enviado ({bytes.Length} bytes)");
            // Impressoras do Windows no painel de impressao (Task 5): depois do EndLoad, fora da thread de interface.
            if (_sessao is { } aberta) ImpressorasEnviadasParaTeste = EnviarImpressorasAoEditorAsync(aberta);
        }
        catch (Exception ex)
        {
            UltimoErro = ex.Message;
            _arquivos.Log("PONTE erro de abertura: " + ex.Message);
            // "" faz o editor mostrar o proprio asc_onError em vez de ficar em "Carregando documento" para
            // sempre. Um segundo try: se ATE ISSO falhar, so loga — nao ha mais nada que possamos tentar aqui.
            try { await ChamarEndLoadAsync("", "", 0); }
            catch (Exception ex2) { _arquivos.Log("PONTE EndLoad de erro tambem falhou: " + ex2.Message); }
            // Sessao RECUPERADA que nao abriu: mesma regra da falha da reaplicacao (fix round 2) — nao ha' como
            // saber o que o editor chegou a aplicar, entao a janela fecha e a pasta fica para outra tentativa.
            if (_sessao is { Recuperada: true }) FalharRecuperacao(ex.Message);
        }
        finally
        {
            // Nao segura o array (pode ser grande) alem do EndLoad, com sucesso ou nao (fix round 2).
            _changesRecuperados = null;
        }
    }

    /// <summary>
    /// "Falhar fechado" da reabertura de uma pasta recuperada (fix round 2 da revisao da Task 2; decisao do
    /// controlador): nunca manter aberta uma sessao cujo editor nao mostra o que o <c>changes0.json</c> diz.
    /// Nesta ordem: congela a ponte (nenhum <c>LocalFileSaveChanges</c>/<c>LocalFileSave</c> que chegue daqui
    /// para a frente grava nada — o arquivo em disco e' a fonte da verdade e tem de ficar byte a byte como o
    /// <see cref="ChangesGravados.Reparar"/>/<see cref="ChangesGravados.Ler"/> deixaram); (a) registra a falha em
    /// <see cref="FalhaDaRecuperacao"/> e no log; (b) marca a janela para manter a pasta ao fechar (so' solta a
    /// trava: a pasta volta a ser recuperavel); (c) dispara <see cref="RecuperacaoFalhou"/>; (d) fecha a janela
    /// SEM a pergunta de salvar (o estado do editor nao e' confiavel, salvar nao pode ser oferecido).
    /// Idempotente: uma segunda falha na mesma janela nao repete nada.
    /// </summary>
    private void FalharRecuperacao(string detalhe) => FalharFechado(
        $"Não foi possível reaplicar as alterações recuperadas de \"{_sessao?.Titulo}\". O documento continua guardado para recuperação.",
        "RECUPERACAO FALHOU (a janela fecha e a pasta fica): " + detalhe, RecuperacaoFalhou);

    /// <summary>
    /// As alteracoes de uma sessao VIVA nao puderam ser gravadas na pasta de recuperacao (fix rounds 2 e 3; ver
    /// <see cref="FalharGravacaoDeAlteracoes"/>): a janela fecha sozinha, sem perguntar, e a pasta fica — o documento
    /// aparece em "Documentos recuperados". Evento PROPRIO (e nao <see cref="RecuperacaoFalhou"/>) porque a orientacao
    /// ao usuario e' outra: aqui e' "use Recuperar", nunca "use Descartar". Mensagem pt-BR pronta
    /// (<see cref="MensagemFalhaAoGravarAlteracoes"/>). Mesma ordem de <see cref="RecuperacaoFalhou"/>: chega ANTES do
    /// fechamento.
    /// </summary>
    public event Action<JanelaEditor, string>? AlteracoesNaoGravadas;

    /// <summary>
    /// O <c>changes0.json</c> nao pode ser posto em dia com o que o editor mandou (fix round 2 da revisao final do Plano
    /// 3; invariante em <see cref="PonteDesktop.SalvarChanges"/>): o arquivo esta' preso por outro programa e nem a
    /// troca atomica, nem o corte no lugar, nem as novas tentativas deram certo. O arquivo em disco continua num estado
    /// COERENTE anterior (o lote que falhou nao entrou; os contadores nao andaram), mas o editor ja' mostra o lote — a
    /// sessao nao pode seguir, senao um salvamento gravaria algo diferente da tela. Mesmo desfecho da falha da
    /// reaplicacao: ponte congelada (JS e C#), pasta mantida, janela fechada sem perguntar, mensagem pt-BR pelo
    /// <see cref="RecuperacaoFalhou"/> (a casca mostra na boas-vindas, onde o documento aparece para recuperar).
    /// </summary>
    private void FalharGravacaoDeAlteracoes(string detalhe)
    {
        // O quadro ja' congelou sozinho (retorno "falhou" do SalvarChanges); isto e' a rede para o documento de cima.
        _ = ExecutarJsAsync("(function(){try{var w=window.frames[0];if(w&&w.__moffice)w.__moffice.congelado=true;if(window.__moffice)window.__moffice.congelado=true;}catch(e){}})()")
            .ContinueWith(t => _arquivos.Log("ALTERACOES congelar o JS falhou: " + t.Exception?.GetBaseException().Message), TaskContinuationOptions.OnlyOnFaulted);
        // A pasta tem o que recuperar? (fix round 4) Sem nenhum change gravado ela nem aparece na lista: a mensagem nao
        // pode mandar o usuario procurar la'.
        var recuperavel = _sessao is { } sessaoAtual && Recuperaveis.TemConteudo(sessaoAtual.PastaRecovery);
        FalharFechado(MensagemFalhaAoGravarAlteracoes(_sessao?.Titulo, NomeApp, recuperavel),
            "ALTERACOES nao gravadas no changes0.json (a sessao congela, a janela fecha e a pasta fica): " + detalhe, AlteracoesNaoGravadas);
    }

    /// <summary>Texto da falha ao gravar as alteracoes de uma sessao viva (fix rounds 2 e 3), pronto para a boas-vindas.</summary>
    /// <param name="recuperavel">A pasta tem changes para recuperar (aparece em "Documentos recuperados"). Sem eles, a
    /// mensagem diz que as ultimas alteracoes se perderam e NAO aponta para Recuperar (fix round 4).</param>
    public static string MensagemFalhaAoGravarAlteracoes(string? titulo, string app, bool recuperavel) => recuperavel
        ? $"Não foi possível gravar as alterações de \"{titulo}\" no disco. O documento foi guardado em Documentos recuperados, "
          + $"na tela inicial do {app}: use Recuperar."
        : $"Não foi possível gravar as últimas alterações de \"{titulo}\" no disco, e elas não puderam ser guardadas para "
          + "recuperação. O arquivo continua como foi salvo pela última vez.";

    /// <summary>Tronco comum das falhas FECHADAS da sessao (reaplicacao recuperada, gravacao das alteracoes): ver
    /// <see cref="FalharRecuperacao"/>. <see cref="FalhaDaRecuperacao"/> guarda a mensagem nos dois casos (bloqueia
    /// salvar/descarregar/imprimir). Idempotente.</summary>
    private void FalharFechado(string mensagem, string linhaDeLog, Action<JanelaEditor, string>? aviso)
    {
        if (FalhaDaRecuperacao is not null) return;
        _ponte?.Congelar();
        FalhaDaRecuperacao = mensagem;
        UltimoErro = mensagem;
        _arquivos.Log(linhaDeLog);
        _manterRecoveryAoFechar = true;
        try { aviso?.Invoke(this, mensagem); }
        catch (Exception ex) { _arquivos.Log("RECUPERACAO erro no aviso de falha: " + ex.Message); }
        _fechamentoConfirmado = true;
        // Enfileirado: podemos estar dentro da continuacao de um ExecuteScriptAsync ou de um aviso do JS.
        Dispatcher.BeginInvoke(() => { if (!_fechada) Close(); });
    }

    /// <summary>
    /// "Inserir imagem do arquivo". Caminho REAL do sdkjs 9.4.0.129 (medido em sdkjs/word/sdk-all.js:9919046,
    /// <c>Asc.asc_docs_api.prototype.qUa = asc_addImage</c>): o editor chama
    /// <c>AscDesktopEditor.OpenFilenameDialog("images", false, cb)</c> e, no callback, chama sozinho
    /// <c>LocalFileGetImageUrl(&lt;caminho&gt;)</c> + <c>editor.RHa([AscCommon.Ss.RS(nome)])</c>. Ou seja: o
    /// host devolve o CAMINHO escolhido, nao o nome da midia — quem copia para <c>media/</c> e' a chamada
    /// seguinte, em <see cref="PonteDesktop.ImagemUrl"/>.
    /// (<c>LocalFileGetImageUrlFromOpenFileDialog</c>, a hipotese do brief, existe no desktop-sdk mas NAO e'
    /// chamada por nenhum arquivo do motor desta versao — ver o relatorio da Task 5.)
    ///
    /// O dialogo e' aberto FORA da chamada sincrona do JS (<c>await Task.Yield()</c>, como o
    /// <see cref="ResponderAberturaAsync"/>): um modal do Windows com o renderer bloqueado esperando o retorno
    /// de um objeto de host trava os dois lados.
    /// </summary>
    private async Task ResponderDialogoImagemAsync()
    {
        await Task.Yield();
        string? escolhido = null;
        try
        {
            escolhido = EscolherImagemParaTeste is not null ? EscolherImagemParaTeste() : EscolherImagemNoWindows();
            if (escolhido is not null) _ponte?.PermitirArquivoEscolhido(escolhido);
            _arquivos.Log("PONTE dialogo de imagem: " + (escolhido is null ? "cancelado" : Path.GetFileName(escolhido)));
        }
        catch (Exception ex)
        {
            // O callback TEM de ser respondido mesmo assim: o sdkjs guarda um unico callback pendente por
            // vez (`if (window.on_native_open_filename_dialog) return;` no app oficial), e deixa-lo pendurado
            // deixaria "inserir imagem" morto ate recarregar o editor.
            _arquivos.Log("PONTE dialogo de imagem falhou: " + ex.Message);
        }
        var js = """
            (function () {
              try {
                var w = (window.__moffice && window.__moffice.dialogoPendente) ? window : window.frames[0];
                if (!w || !w.__moffice || typeof w.__moffice.responderDialogo !== "function") return "sem ponte no quadro";
                return "ok " + w.__moffice.responderDialogo(ESCOLHIDO);
              } catch (e) { return "erro: " + e; }
            })()
            """.Replace("ESCOLHIDO", System.Text.Json.JsonSerializer.Serialize(escolhido ?? ""));
        try
        {
            var r = (await ExecutarJsAsync(js)).Trim('"');
            if (!r.StartsWith("ok", StringComparison.Ordinal)) _arquivos.Log("PONTE resposta do dialogo de imagem: " + r);
        }
        catch (Exception ex) { _arquivos.Log("PONTE nao conseguiu responder o dialogo de imagem: " + ex.Message); }
    }

    private static string? EscolherImagemNoWindows()
    {
        var dialogo = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Inserir imagem",
            Filter = "Imagens|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.svg|Todos os arquivos|*.*",
            CheckFileExists = true,
            Multiselect = false,
        };
        return dialogo.ShowDialog() == true ? dialogo.FileName : null;
    }

    /// <summary>O que a troca da base das midias fez no ultimo <c>EndLoad</c> (<c>"trocado &lt;objeto&gt;.&lt;prop&gt;
    /// -&gt; https://moffice-doc.local"</c>, ou a explicacao de por que nao trocou). E' a unica leitura possivel
    /// desse passo de fora: a propriedade vive no bundle minificado do editor e o nome dela e' diferente em
    /// cada um dos tres (ver o comentario em <c>ponte-desktop.js</c>).</summary>
    public string? TrocaDaBaseDeMidia { get; private set; }

    /// <summary>A MESMA URL que o <c>DesktopOfflineAppDocumentEndLoad</c> acabou de montar a partir da pasta:
    /// prefixa <c>/</c> quando o caminho nao comeca com barra (o caso Windows, <c>C:/…</c>) e depois
    /// <c>file://</c> — copiado da regra do sdkjs. E' por esse valor que a troca da base encontra a
    /// propriedade, ja' que o NOME dela muda a cada bundle.</summary>
    private static string UrlDeArquivoDaPasta(string pasta) =>
        pasta.Length == 0 ? "" : "file://" + (pasta.StartsWith('/') ? pasta : "/" + pasta);

    /// <summary><paramref name="changesRecuperados"/> (Task 2 do Plano 3): numa sessao recuperada, o array JS
    /// dos changes gravados (<see cref="ChangesGravados.Ler"/>) e quantos sao. Vao para o editor ANTES do
    /// EndLoad, na ordem do app oficial: <c>LocalFileSetSaved(false)</c>, <c>LocalFileRecoverFolder</c>,
    /// <c>LocalFileSetSourcePath</c>, <c>DesktopOfflineAppDocumentApplyChanges([...])</c>,
    /// <c>LocalFileSetOpenChangesCount(n)</c> — o proprio editor aplica os changes ao abrir o documento.</summary>
    private async Task ChamarEndLoadAsync(string pasta, string base64, int tamanho, (string ArrayJs, int Quantidade)? changesRecuperados = null)
    {
        // A sessao para a qual este EndLoad foi pedido (antes de qualquer await): o sinal 1 da confirmacao da
        // recuperacao so' vale para ela.
        var sessaoDoEndLoad = _sessao;
        // DesktopOfflineAppDocumentEndLoad vive no quadro do editor (sdkjs/word/sdk-all.js), nao no documento
        // de cima; mesma origem, entao window.frames[0] alcanca.
        var argumentos = string.Join(", ",
            System.Text.Json.JsonSerializer.Serialize(pasta),
            System.Text.Json.JsonSerializer.Serialize(base64),
            tamanho.ToString());
        // Montagem por interpolacao (UMA passada), nao por .Replace encadeado (Task 2 do Plano 3): com
        // Replace, o texto ja' inserido (o base64 do Editor.bin, o array de changes) era varrido de novo pelas
        // trocas seguintes — um "BASE" dentro do base64 do documento viraria a URL das midias e corromperia o
        // documento antes de chegar ao editor.
        // try PROPRIO em volta da reaplicacao (fix round 1 da revisao): se ApplyChanges/SetOpenChangesCount
        // lancarem, o EndLoad TEM de rodar mesmo assim — sem ele o editor fica em "Carregando documento" para
        // sempre. A falha fica guardada no quadro e e' lida logo depois (ver abaixo) -> FalharRecuperacao.
        // Fix round 2: o catch CONGELA a ponte do quadro na mesma hora (`__moffice.congelado`, lido por
        // LocalFileSaveChanges/LocalFileSave em ponte-desktop.js). O JS e' de uma thread so': entre a falha e o
        // host ficar sabendo, nenhum lote de changes do editor chega a gravar no changes0.json.
        var falhaDeTeste = FalharReaplicacaoParaTeste ? "throw new Error(\"falha simulada da reaplicacao (teste)\"); " : "";
        var recuperar = changesRecuperados is { } c && pasta.Length > 0 && _sessao is { } s
            ? "w.__mofficeErroRecuperacao = \"\"; try { " + falhaDeTeste + "var d = w.AscDesktopEditor; d.LocalFileSetSaved(false); d.LocalFileRecoverFolder(" + System.Text.Json.JsonSerializer.Serialize(pasta)
              + "); d.LocalFileSetSourcePath(" + System.Text.Json.JsonSerializer.Serialize(s.ArquivoOrigem)
              + "); w.DesktopOfflineAppDocumentApplyChanges(" + c.ArrayJs + "); d.LocalFileSetOpenChangesCount(" + c.Quantidade + ");"
              + " } catch (e) { if (w.__moffice) w.__moffice.congelado = true; w.__mofficeErroRecuperacao = String((e && (e.stack || e.message)) || e) || \"erro\"; }"
            : "";
        var baseMidia = System.Text.Json.JsonSerializer.Serialize(pasta.Length == 0 ? "" : Midia.BaseUrlDocumento);
        var antigo = System.Text.Json.JsonSerializer.Serialize(AntigoParaTrocaDeBase(pasta));
        var js = $$"""
            (function () {
              try {
                var w = (typeof DesktopOfflineAppDocumentEndLoad === "function") ? window : window.frames[0];
                if (!w || typeof w.DesktopOfflineAppDocumentEndLoad !== "function") return "sem EndLoad";
                {{recuperar}}
                w.DesktopOfflineAppDocumentEndLoad({{argumentos}});
                // Troca da base das midias, LOGO depois do EndLoad (Task 5/S4, pergunta aberta da Task 5A):
                // o proprio EndLoad acabou de fazer AscCommon.Ss.tia = "file://" + <pasta>, e o Chromium
                // recusa subrecurso file:// numa pagina https. A pasta de recovery ja' e' servida por
                // https://moffice-doc.local; e' so' repontar a propriedade (nenhum patch no motor).
                // A base so' e' trocada no EndLoad de SUCESSO: no EndLoad de erro ("") nao ha documento nem
                // pasta, e apontar a midia para o host virtual mapeado a um documento anterior seria mentira.
                if (!w.__moffice || typeof w.__moffice.trocarBaseMidia !== "function") return "sem ponte no quadro";
                return "ok " + w.__moffice.trocarBaseMidia({{baseMidia}}, {{antigo}});
              } catch (e) { return "erro: " + e; }
            })()
            """;
        var r = await ExecutarJsAsync(js);
        var limpo = r.Trim('"');
        if (!limpo.StartsWith("ok", StringComparison.Ordinal))
        {
            _arquivos.Log("PONTE EndLoad falhou: " + r);
            // So' com pasta REAL (fix 2 da revisao da Task 5, item 4): no EndLoad de ERRO (pasta == "", chamado
            // pelo catch de ResponderAberturaAsync) o UltimoErro JA' tem a causa verdadeira da falha de
            // abertura (x2t, leitura do .bin...), e sobrescreve-la com "sem EndLoad"/"erro: ..." esconderia o
            // motivo — mesma regra do ramo de sucesso abaixo. O log sai sempre.
            if (pasta.Length > 0)
            {
                UltimoErro = r;
                // "erro: ..." = o JS LANCOU dentro do bloco (por exemplo dentro de trocarBaseMidia, DEPOIS de
                // o proprio DesktopOfflineAppDocumentEndLoad ja' ter rodado): o documento pode ter carregado e
                // a troca da base nao e' confiavel — as figuras podem nao aparecer. Marcar MidiasIndisponiveis
                // como no ramo de sucesso-sem-troca, para o estado ficar observavel pelos apps do Plano 2.
                // Nao ha' teste real para isto: nao existe costura barata que faca o JS lancar aqui sem
                // alterar o bundle, e ChamarEndLoadAsync e' privado; fica documentado em vez de fabricado.
                if (limpo.StartsWith("erro:", StringComparison.Ordinal)) MidiasIndisponiveis = true;
            }
        }
        else
        {
            // Guarda contra resultado curto (achado maior da revisao 1): "ok" sozinho (sem o "ok " + algo que
            // o JS sempre devolve) nunca deveria acontecer, mas um `[3..]` sem isto estouraria
            // ArgumentOutOfRangeException dentro do handler do JS e derrubaria a janela por um detalhe de
            // string, exatamente o tipo de falha muda que esta revisao esta' fechando.
            TrocaDaBaseDeMidia = limpo.Length > 3 ? limpo[3..] : "";
            _arquivos.Log("PONTE base das midias: " + TrocaDaBaseDeMidia);
            // A troca so' e' "midia indisponivel" quando HAVIA pasta para trocar (documento carregado de
            // verdade). No EndLoad de ERRO (pasta == "", chamado pelo catch de ResponderAberturaAsync)
            // trocarBaseMidia devolve de proposito "sem troca (EndLoad de erro)" — sem este `pasta.Length > 0`,
            // esse caminho sobrescreveria o UltimoErro que JA' explica a causa real da falha de abertura com
            // uma mensagem sobre midia, escondendo o motivo verdadeiro.
            if (pasta.Length > 0 && !TrocaDaBaseDeMidia.StartsWith("trocado ", StringComparison.Ordinal))
            {
                MidiasIndisponiveis = true;
                // Loud on purpose (achado maior da revisao 1): antes, "nao achei a base"/"sem AscCommon" so'
                // existiam no log de egresso — o documento abria, o texto era editavel, e NADA na API do host
                // dizia que as figuras nunca vao aparecer. O texto continua editavel de proposito (nao ha' por
                // que bloquear o documento so' porque a midia falhou), mas agora o estado fica OBSERVAVEL: os
                // apps do Plano 2 leem MidiasIndisponiveis para mostrar uma barra de aviso.
                UltimoErro = "As imagens do documento não vão aparecer: a ponte não encontrou a base de mídias "
                    + "do editor (versão do motor diferente da esperada?)";
                _arquivos.Log("PONTE " + UltimoErro + " (" + TrocaDaBaseDeMidia + ")");
            }
        }
        // Falha na reaplicacao dos changes recuperados (fix round 1 da revisao): o EndLoad rodou mesmo assim (try
        // proprio no script), mas o documento pode ter aberto SEM as alteracoes recuperadas (ou com parte delas).
        // Fix round 2: isso nao e' so' "observavel" — a sessao e' inconsistente e FALHA FECHADA
        // (FalharRecuperacao: a janela fecha sem oferecer salvar e a pasta fica). Idem se o script inteiro nao
        // deu "ok" numa sessao recuperada: nao ha' como saber o que o editor chegou a aplicar.
        if (changesRecuperados is not null && pasta.Length > 0)
        {
            if (!limpo.StartsWith("ok", StringComparison.Ordinal)) { FalharRecuperacao("EndLoad: " + r); return; }
            var erro = System.Text.Json.JsonSerializer.Deserialize<string>(await ExecutarJsAsync("""
                (function () {
                  try {
                    var w = (typeof DesktopOfflineAppDocumentEndLoad === "function") ? window : window.frames[0];
                    return String((w && w.__mofficeErroRecuperacao) || "");
                  } catch (e) { return "erro ao ler a falha: " + e; }
                })()
                """)) ?? "";
            if (erro.Length > 0) { FalharRecuperacao("reaplicacao: " + erro); return; }
            // Sinal 1 da confirmacao (fix round 2 da revisao da Task 3): o EndLoad da sessao recuperada voltou "ok"
            // e a reaplicacao nao deixou erro. NAO basta sozinho — os changes podem so' ter sido enfileirados (ver
            // _recuperacaoConfirmada); a confirmacao espera tambem o "pronto" do editor DESTA sessao.
            if (FalhaDaRecuperacao is null && ReferenceEquals(sessaoDoEndLoad, _sessao) && _sessao is { Recuperada: true })
            {
                _endLoadRecuperadoOk = true;
                _arquivos.Log("RECUPERACAO EndLoad ok (esperando o editor ficar pronto para confirmar)");
                ConfirmarRecuperacaoSePronta();
            }
        }
    }

    /// <summary>Valor que <c>trocarBaseMidia</c> procura em <c>AscCommon</c> — o mesmo que
    /// <see cref="UrlDeArquivoDaPasta"/> monta, corrompido quando <see cref="DesligarTrocaDeBaseParaTeste"/>
    /// esta' ligado (so' quando ha' pasta real: com pasta vazia o EndLoad de erro nem chega a procurar nada,
    /// ver o comentario em <see cref="ChamarEndLoadAsync"/>).</summary>
    private string AntigoParaTrocaDeBase(string pasta)
    {
        var antigo = UrlDeArquivoDaPasta(pasta);
        return DesligarTrocaDeBaseParaTeste && antigo.Length > 0 ? antigo + "-teste-nao-existe" : antigo;
    }

    // ---------------------------------------------------------------------------------------------------
    // Arquivos protegidos por senha (Task 7 do Plano 3). Protocolo MEDIDO nos bundles 9.4.0.129 e no desktop-sdk
    // (relatorio da Task 7): o x2t sai com 90 (protegido) ou 91 (senha errada) na conversao para Editor.bin; o
    // host chama window.asc_initAdvancedOptions(codigo) no quadro do editor (o oficial faz o mesmo em
    // "onlocaldocument_additionalparams"), o editor mostra o proprio dialogo de senha e responde
    // AscDesktopEditor.SetAdvancedOptions("<m_sPassword>senha</m_sPassword>"); o host reconverte com m_sPassword e
    // segue para o EndLoad normal. Ao salvar, a senha vai em m_sSavePassword. A senha so' vive em memoria
    // (SessaoDocumento.Senha): nunca em disco, no log ou de volta para a pagina.
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Avisos da ponte cujo parametro leva a SENHA do documento: o <see cref="Avisar"/> nunca registra o
    /// conteudo deles no log. `internal` para o teste da regra.</summary>
    internal static readonly IReadOnlySet<string> AvisosSigilosos =
        new HashSet<string>(StringComparer.Ordinal) { "opcoesAvancadas", "senhaSalvar" };

    /// <summary>Aviso na faixa quando uma sessao RECUPERADA era de um documento protegido: o Editor.bin recuperado
    /// ja' esta' decifrado e a senha se perdeu com a sessao que caiu — salvar sem defini-la de novo gravaria o
    /// arquivo sem protecao. "Arquivo → Proteger" e' o rotulo do menu do editor em pt-BR (web-apps
    /// <c>DE.Views.FileMenu.btnProtectCaption</c>), disponivel porque a ponte responde IsProtectionSupport = true.</summary>
    public const string AvisoProtegidoRecuperado =
        "Este documento era protegido por senha: antes de salvar, defina a senha de novo em Arquivo → Proteger.";

    /// <summary>A sessao atual e' uma RECUPERACAO de documento protegido e ainda nao tem senha (ver
    /// <see cref="AvisoProtegidoRecuperado"/>). A casca le' isto quando o documento fica pronto.</summary>
    public bool RecuperadoSemSenha => _sessao is { Recuperada: true, Senha: null, Metadados.Protegido: true };

    /// <summary>Recusa de "Salvar" por cima do original cifrado numa sessao recuperada sem senha (I1 da revisao final
    /// do Plano 3). Vai para o editor como falha de salvamento (<c>EndSave</c> 2) e para a faixa
    /// (<see cref="AvisoAoUsuario"/>); o documento continua modificado.</summary>
    public const string MensagemProtegidoSemSenhaRecusado =
        "Este documento era protegido por senha. Defina a senha de novo em Arquivo → Proteger antes de salvar, ou use Salvar como para gravar uma cópia.";

    /// <summary>Faixa depois de um "Salvar como" bem-sucedido de uma sessao recuperada de documento protegido, ainda sem
    /// senha (I1): a copia saiu SEM protecao.</summary>
    public const string MensagemRecuperadoSalvoSemSenha =
        "O arquivo foi salvo sem senha: a senha do documento original não é recuperada depois de uma queda. Para protegê-lo, use Arquivo → Proteger.";

    /// <summary>Faixa de "Baixar arquivos" (imagem por URL, comparar/combinar por URL): <c>AscDesktopEditor.DownloadFiles</c>.</summary>
    public const string MensagemBaixarArquivos = "Baixar arquivos da internet não é suportado no mOffice (sem acesso à rede).";

    /// <summary>Faixa de "Importar texto/CSV de arquivo" da planilha (<c>loadLocalFile</c> e o dialogo "csv/txt").</summary>
    public const string MensagemImportarTexto = "Importar texto ou CSV de um arquivo não é suportado no mOffice.";

    /// <summary>Faixa de "Comparar"/"Combinar" com documento de arquivo (<c>CompareDocumentFile</c>/<c>MergeDocumentFile</c>
    /// e o dialogo "word").</summary>
    public const string MensagemCompararDocumento = "Comparar ou combinar com outro documento não é suportado no mOffice.";

    /// <summary>Chave do aviso "semSuporte" da ponte (ponte-desktop.js) -> faixa pt-BR. Chave desconhecida so' vai ao log.</summary>
    internal static readonly IReadOnlyDictionary<string, string> MensagensSemSuporte = new Dictionary<string, string>
    {
        ["novaPastaDeTrabalho"] = MensagemNovaPastaDeTrabalho,
        ["baixarArquivos"] = MensagemBaixarArquivos,
        ["importarTexto"] = MensagemImportarTexto,
        ["compararDocumento"] = MensagemCompararDocumento,
    };

    /// <summary>Faixa quando a linha "PDF assinado com mPDF" foi escolhida mas o nome digitado tinha extensao de DOCUMENTO
    /// (<see cref="DestinoDoDialogo"/>: a extensao vence e o documento e' gravado).</summary>
    public static string MensagemAssinarSalvouDocumento(FormatoDocumento f) =>
        $"O arquivo foi salvo como documento ({f.Extensao().ToUpperInvariant()}), não como PDF: nada foi enviado para assinar.";

    /// <summary>Faixa quando a planilha pede "Mover ou copiar" para "(Criar nova planilha)" (revisao final do Plano 4): o
    /// app oficial abre outra janela com a pasta de trabalho nova (<c>AscDesktopEditor.OpenWorkbook</c>); o mOffice nao.</summary>
    public const string MensagemNovaPastaDeTrabalho =
        "Mover ou copiar para uma nova pasta de trabalho não é suportado no mOffice: copie a planilha dentro desta pasta de trabalho.";

    /// <summary>A sessao atual e' uma recuperacao cujo arquivo original mudou (ou sumiu) depois da queda (I2 da revisao
    /// final do Plano 3, <see cref="SessaoDocumento.Desvinculada"/>). A casca le' isto quando o documento fica pronto.</summary>
    public bool OriginalAlteradoDepoisDaQueda => _sessao is { Recuperada: true, Desvinculada: true };

    /// <summary>Faixa da sessao recuperada desvinculada (ver <see cref="OriginalAlteradoDepoisDaQueda"/>).</summary>
    public const string AvisoOriginalAlterado =
        "O arquivo original foi alterado (ou removido) depois da queda: para não sobrescrever a versão mais nova, ao salvar o mOffice vai pedir um nome novo.";

    /// <summary>Recusa de gravar por cima do original numa sessao recuperada desvinculada (I2), para quem escolher o
    /// proprio original no "Salvar como".</summary>
    public const string MensagemOriginalAlteradoRecusado =
        "O arquivo original foi alterado (ou removido) depois da queda: para não sobrescrever a versão mais nova, salve estas alterações com outro nome.";

    /// <summary>Por que este salvamento NAO pode gravar em <paramref name="destino"/> (texto pt-BR), ou <c>null</c>.
    /// Falha FECHADA das duas regras da revisao final do Plano 3: I1 (original cifrado sem senha definida) e I2
    /// (original alterado depois da queda). Exportacao (PDF) nunca cai aqui: o destino e' sempre outro arquivo.</summary>
    private static string? MotivoParaRecusarGravacao(SessaoDocumento s, string destino, FormatoDocumento fmt)
    {
        if (s.OriginalProtegidoSemSenha is { } cifrado && s.Senha is null && X2tConversor.FormatoGuardaSenha(fmt)
            && MesmoArquivo(destino, cifrado))
            return MensagemProtegidoSemSenhaRecusado;
        if (s.Desvinculada && MesmoArquivo(destino, s.ArquivoOrigem)) return MensagemOriginalAlteradoRecusado;
        return null;
    }

    private static bool MesmoArquivo(string a, string b) =>
        string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);

    /// <summary>Texto da faixa quando um documento protegido e' gravado num formato que nao guarda senha
    /// (<see cref="X2tConversor.FormatoGuardaSenha"/>).</summary>
    public static string MensagemSalvoSemSenha(FormatoDocumento f) =>
        $"O arquivo foi salvo sem senha: o formato {f.Extensao().ToUpperInvariant()} não guarda proteção.";

    /// <summary>Aviso pt-BR, pronto para a faixa da casca, sobre algo que deu certo mas o usuario precisa saber (hoje:
    /// <see cref="MensagemSalvoSemSenha"/>). Thread da interface; try/catch como os demais eventos publicos.</summary>
    public event Action<string>? AvisoAoUsuario;

    /// <summary>Codigo do x2t que o editor esta' tratando com o dialogo de senha (90/91), ou <c>null</c> sem pedido
    /// em curso. So' com um pedido pendente um SetAdvancedOptions e' aceito.</summary>
    private int? _opcaoAvancadaPendente;

    /// <summary>Ha' um pedido de senha em curso no editor (so' leitura, para os testes).</summary>
    internal int? OpcaoAvancadaPendente => _opcaoAvancadaPendente;

    private static bool PedeSenha(int codigo) => codigo is X2tConversor.CodigoProtegido or X2tConversor.CodigoSenhaErrada;

    /// <summary>Chama <c>window.asc_initAdvancedOptions(codigo)</c> no quadro do editor. Assinatura MEDIDA em
    /// <c>sdkjs/*/sdk-all-min.js</c>: <c>(codigo, hash, docinfo, arquivoAberto)</c> — os tres ultimos so' servem a'
    /// criptografia de plugin/visualizador nativo e vao ausentes, como o oficial manda <c>undefined</c> sem eles. Com
    /// <c>AscDesktopEditor.CryptoMode === 0</c> (a ponte define) o editor dispara <c>asc_onAdvancedOptions(DRM)</c> e o
    /// web-apps abre o dialogo de senha (91 = com o aviso de senha incorreta). Espelho para os testes:
    /// <c>__moffice.ultimaOpcaoAvancada</c> (documento de cima e quadro).</summary>
    /// <exception cref="InvalidOperationException">O quadro nao tem o editor (a abertura falha como qualquer outra).</exception>
    private async Task PedirSenhaAoEditorAsync(int codigo)
    {
        _opcaoAvancadaPendente = codigo;
        _arquivos.Log($"SENHA documento protegido (x2t {codigo}): pedindo a senha ao editor");
        var js = """
            (function () {
              try {
                var w = (typeof asc_initAdvancedOptions === "function") ? window : window.frames[0];
                if (!w || typeof w.asc_initAdvancedOptions !== "function") return "sem asc_initAdvancedOptions";
                if (window.__moffice) window.__moffice.ultimaOpcaoAvancada = CODIGO;
                if (w.__moffice) w.__moffice.ultimaOpcaoAvancada = CODIGO;
                w.asc_initAdvancedOptions(CODIGO);
                return "ok";
              } catch (e) { return "erro: " + e; }
            })()
            """.Replace("CODIGO", codigo.ToString());
        var r = (await ExecutarJsAsync(js)).Trim('"');
        if (r != "ok")
        {
            _opcaoAvancadaPendente = null;
            throw new InvalidOperationException("O documento é protegido por senha e o editor não conseguiu pedi-la: " + r);
        }
    }

    /// <summary>Resposta do dialogo de senha (<c>SetAdvancedOptions</c>). Roda DENTRO da chamada sincrona do JS: so'
    /// agenda. Sem pedido pendente, ignora (a pagina nao dispara conversao sozinha). Sem <c>&lt;m_sPassword&gt;</c>
    /// (XML quebrado, opcoes de CSV) = abertura cancelada: a janela fecha, sem mensagem de erro. Com senha:
    /// reconverte com ela e volta para <see cref="ResponderAberturaAsync"/> (91 de novo = pede de novo).</summary>
    private void ReceberOpcoesAvancadas(SessaoDocumento sessao, string xml)
    {
        if (_opcaoAvancadaPendente is null)
        {
            _arquivos.Log("SENHA opcoes avancadas sem pedido pendente: ignoradas");
            return;
        }
        _opcaoAvancadaPendente = null;
        var senha = OpcoesAvancadas.Senha(xml);
        if (senha is null)
        {
            _arquivos.Log("SENHA abertura cancelada (resposta sem senha): a janela fecha");
            Dispatcher.BeginInvoke(() => { if (!_fechada) Close(); });
            return;
        }
        sessao.Senha = senha;
        _binPronto = _conv.ParaBinAsync(sessao.ArquivoOrigem, sessao.PastaRecovery, senha: senha);
        _ = ResponderAberturaAsync();
    }

    /// <summary>A senha que o editor manda em cada <c>LocalFileSave</c> (a atual dele: a da abertura ou a definida/
    /// removida em Arquivo -&gt; Proteger; "" = sem senha) passa a ser a da sessao — e' ela que o salvamento que vem
    /// logo atras grava em <c>m_sSavePassword</c>. O <c>sessao.json</c> acompanha (<see cref="MetadadosSessao.Protegido"/>).</summary>
    private void AtualizarSenha(SessaoDocumento sessao, string senha)
    {
        var nova = senha.Length == 0 ? null : senha;
        // Recuperada de um protegido (I1): com a senha definida de novo, o original cifrado volta a poder ser
        // regravado — cifrado com ela.
        if (nova is not null && sessao.OriginalProtegidoSemSenha is not null)
        {
            sessao.OriginalProtegidoSemSenha = null;
            _arquivos.Log("SENHA definida de novo numa sessao recuperada de documento protegido: o original volta a poder ser salvo");
        }
        if (nova == sessao.Senha) return;
        sessao.Senha = nova;
        _arquivos.Log("SENHA " + (nova is null ? "removida" : "definida") + " pelo editor (vale a partir deste salvamento)");
        MarcarProtegido(sessao, nova is not null);
    }

    /// <summary>Regrava o <c>sessao.json</c> com <see cref="MetadadosSessao.Protegido"/> (melhor esforco, como o
    /// resto dos metadados: sem trava nao ha' metadados, e uma falha de E/S so' vai para o log).</summary>
    private void MarcarProtegido(SessaoDocumento sessao, bool protegido)
    {
        if (sessao.Metadados is not { } m || m.Protegido == protegido) return;
        sessao.Metadados = m with { Protegido = protegido };
        try { sessao.Metadados.Gravar(sessao.PastaRecovery); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { _arquivos.Log("RECUPERACAO nao regravou sessao.json: " + e.Message); }
    }

    // ---------------------------------------------------------------------------------------------------
    // Tela cheia da apresentacao (Task 7 do Plano 3): o editor de slides chama SetFullscreen(true) ao iniciar a
    // demonstracao e SetFullscreen(false) ao sair (Esc/fim). O host deixa a janela sem borda, do tamanho do
    // MONITOR onde ela esta' (por cima da barra de tarefas) e depois devolve tudo como estava.
    // ---------------------------------------------------------------------------------------------------

    /// <summary>A janela esta' em tela cheia (apresentacao de slides).</summary>
    public bool EmTelaCheia { get; private set; }

    /// <summary>O que a tela cheia muda e a saida devolve. <c>Retangulo</c> e' o da janela NORMAL (numa janela
    /// maximizada, o <see cref="Window.RestoreBounds"/>), para a volta a "maximizada" restaurar tambem o tamanho
    /// que ela tem quando o usuario desmaximiza.</summary>
    private (WindowStyle Estilo, ResizeMode Redimensionar, WindowState Estado, Rect Retangulo, bool Topmost, Visibility Barra)? _antesDaTelaCheia;

    /// <summary>Entra/sai da tela cheia. Idempotente. Roda na thread da interface. `internal`: o caminho do produto e'
    /// o aviso "telaCheia" da ponte.</summary>
    internal void DefinirTelaCheia(bool ligar)
    {
        if (ligar == EmTelaCheia) return;
        _geracaoTelaCheia++;   // qualquer transicao cancela uma saida garantida pendente (EscNaTelaCheia)
        if (ligar)
        {
            var retangulo = WindowState == WindowState.Normal ? new Rect(Left, Top, Width, Height) : RestoreBounds;
            _antesDaTelaCheia = (WindowStyle, ResizeMode, WindowState, retangulo, Topmost, SlotBarra.Visibility);
            // O monitor e' medido ANTES de mexer no estado: e' o monitor onde a janela esta' agora (maximizada ou nao).
            var monitor = RetanguloDoMonitorEmDips();
            WindowState = WindowState.Normal;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            SlotBarra.Visibility = Visibility.Collapsed;   // a faixa de aviso nao fica por cima dos slides
            Left = monitor.Left;
            Top = monitor.Top;
            Width = monitor.Width;
            Height = monitor.Height;
            EmTelaCheia = true;
            Topmost = IsActive;   // e segue a ativacao (Activated/Deactivated no construtor)
            _arquivos.Log($"TELA CHEIA ligada: {monitor.Width:0.##}x{monitor.Height:0.##} em ({monitor.Left:0.##},{monitor.Top:0.##})");
            Web.Focus();
        }
        else
        {
            EmTelaCheia = false;
            if (_antesDaTelaCheia is { } a)
            {
                // Ordem inversa da entrada: o estilo (com a borda) volta ANTES do retangulo e do estado, para uma
                // janela que era maximizada voltar a maximizar com a borda certa.
                Topmost = a.Topmost;
                SlotBarra.Visibility = a.Barra;
                WindowStyle = a.Estilo;
                ResizeMode = a.Redimensionar;
                Left = a.Retangulo.Left;
                Top = a.Retangulo.Top;
                Width = a.Retangulo.Width;
                Height = a.Retangulo.Height;
                WindowState = a.Estado;
            }
            _antesDaTelaCheia = null;
            _arquivos.Log("TELA CHEIA desligada: janela restaurada");
            Web.Focus();
        }
    }

    /// <summary>Quanto o host espera, depois de um Esc em tela cheia, o editor pedir <c>SetFullscreen(false)</c> antes de
    /// restaurar a janela sozinho (fix round 1 da Task 7).</summary>
    internal static readonly TimeSpan EsperaSaidaDaTelaCheia = TimeSpan.FromSeconds(1);

    /// <summary>Contador de transicoes da tela cheia: a saida garantida so' age se nada mudou desde o Esc.</summary>
    private int _geracaoTelaCheia;

    /// <summary>Costura de teste (fix round 1): o <c>SetFullscreen(false)</c> do editor e' ignorado, como um editor que
    /// falhe no encerramento da apresentacao. `internal`.</summary>
    internal bool IgnorarSaidaDaTelaCheiaParaTeste { get; set; }

    /// <summary>Rede de seguranca da tela cheia (fix round 1 da Task 7): um Esc com a janela em tela cheia SEMPRE tira o
    /// usuario dela. O normal e' o editor encerrar a apresentacao e pedir <c>SetFullscreen(false)</c>; se isso nao chegar
    /// em <see cref="EsperaSaidaDaTelaCheia"/>, o host restaura a janela e registra no log. Roda na thread da interface.</summary>
    private void EscNaTelaCheia(string origem)
    {
        if (!EmTelaCheia) return;
        var geracao = _geracaoTelaCheia;
        _ = GarantirSaidaDaTelaCheiaAsync(geracao, origem);
    }

    private async Task GarantirSaidaDaTelaCheiaAsync(int geracao, string origem)
    {
        await Task.Delay(EsperaSaidaDaTelaCheia);   // continua na thread da interface (contexto do Dispatcher)
        if (_fechada || !EmTelaCheia || geracao != _geracaoTelaCheia) return;
        _arquivos.Log($"TELA CHEIA saida forcada pelo host: Esc ({origem}) sem SetFullscreen(false) do editor em {EsperaSaidaDaTelaCheia.TotalSeconds:0} s");
        DefinirTelaCheia(false);
    }

    /// <summary>Retangulo INTEIRO do monitor da janela (<c>rcMonitor</c>, que inclui a barra de tarefas — e' o que
    /// uma apresentacao cobre), convertido de pixels para DIPs pela escala da propria janela
    /// (<see cref="VisualTreeHelper.GetDpi(Visual)"/>: licao do mPDF sobre escala de RDP/monitor — nunca supor 96
    /// DPI). Sem handle ou se a API falhar, o monitor principal do WPF.</summary>
    internal Rect RetanguloDoMonitorEmDips()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        var info = new MonitorInfo { Tamanho = Marshal.SizeOf<MonitorInfo>() };
        if (hwnd == IntPtr.Zero || !GetMonitorInfo(MonitorFromWindow(hwnd, MonitorMaisProximo), ref info))
            return new Rect(0, 0, SystemParameters.PrimaryScreenWidth, SystemParameters.PrimaryScreenHeight);
        var dpi = VisualTreeHelper.GetDpi(this);
        var r = info.Monitor;
        return new Rect(r.Esquerda / dpi.DpiScaleX, r.Topo / dpi.DpiScaleY,
            (r.Direita - r.Esquerda) / dpi.DpiScaleX, (r.Base - r.Topo) / dpi.DpiScaleY);
    }

    private const uint MonitorMaisProximo = 2; // MONITOR_DEFAULTTONEAREST

    [StructLayout(LayoutKind.Sequential)]
    internal struct RetanguloWin32 { public int Esquerda, Topo, Direita, Base; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MonitorInfo
    {
        public int Tamanho;
        public RetanguloWin32 Monitor;
        public RetanguloWin32 AreaDeTrabalho;
        public uint Flags;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);

    // ---------------------------------------------------------------------------------------------------
    // Impressao (Task 4 do Plano 3). Decisao do usuario: imprimir pelo proprio mOffice com o que o ONLYOFFICE
    // ja' da' — o painel Arquivo -> Imprimir do editor (pre-visualizacao, impressora, paginas, copias...). O
    // editor chama AscDesktopEditor.Print(json); o renderizador nativo do app oficial (Print_Start/Page/End com
    // comandos de desenho) nao e' reproduzivel, entao o caminho e': descarregar os changes pendentes, gerar o
    // PDF pelo x2t (o mesmo DeBinAsync do salvar, alvo PDF, com o JSON do editor em m_sJsonParams) e entregar o
    // PDF a quem manda para a impressora (ImprimirPedido; a casca liga o envio pela WebView2 oculta —
    // Impressao/Impressora.cs, decisao A da sonda S5). A lista de impressoras do painel vem do host
    // (EnviarImpressorasAoEditorAsync, Task 5).
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Limite de <see cref="DescarregarChangesAsync"/>. O autosave do proprio editor grava sozinho em
    /// ~2 s depois de uma edicao (medido nas Tasks 2 e 4); 10 s e' folga, nao expectativa.</summary>
    private static readonly TimeSpan EsperaDescarga = TimeSpan.FromSeconds(10);

    /// <summary>
    /// Forca o editor a gravar em <c>changes0.json</c> os changes locais ainda nao enviados, SEM salvar o arquivo
    /// do usuario e SEM mexer em <see cref="Modificado"/>. Volta na hora quando nao ha' nada pendente.
    ///
    /// Desenho (medido no sdkjs 9.4.0.129, ver o relatorio da Task 4):
    /// <list type="number">
    /// <item><b>Ha' algo pendente?</b> O historico do editor (<c>AscCommon.History</c>) tem o metodo que o autosave
    /// consulta — <c>Have_Changes</c>, minificado com um nome DIFERENTE em cada bundle (<c>l5</c> no texto,
    /// <c>r8</c> na planilha, <c>Via</c> na apresentacao). Com o 1o argumento <c>true</c> ele compara o ponto
    /// atual do historico com o ultimo ENVIADO (e nao com o ultimo salvo): e' exatamente "ha' change local nao
    /// gravado". Como o nome muda, o metodo e' achado pelo CONTEUDO — e' o unico do historico que chama
    /// <c>LocalFileGetOpenChangesCount</c> (uma ocorrencia por bundle) —, no mesmo espirito da troca da base de
    /// midias. Na planilha, com o historico vazio, esse metodo cai no ramo "documento recuperado"
    /// (<c>LocalFileGetOpenChangesCount() != 0 || !LocalFileGetSaved()</c>) mesmo com o 1o argumento true; por
    /// isso, so' durante a pergunta, essas duas funcoes da ponte respondem "0"/"salvo" (JS e' de uma thread
    /// so': nada mais roda no meio).</item>
    /// <item><b>Descarregar.</b> <c>asc_Save(true, false, true)</c> = o caminho do AUTOSAVE (<c>n1(!0)</c> no
    /// texto/apresentacao, <c>Uxa(!0)</c> na planilha): com o 1o argumento true o <c>StartSave</c>
    /// (<c>LocalFileSave</c>) NAO roda — so' os changes vao para <c>LocalFileSaveChanges</c> (Task 2). O 3o
    /// argumento forca a entrada no texto/apresentacao mesmo quando o editor acha que o documento nao esta'
    /// modificado (editar e desfazer ate' o estado salvo ainda deixa changes a enviar).</item>
    /// <item><b>Esperar.</b> O <see cref="PonteDesktop.ChangesSalvos"/> do lote; depois pergunta de novo (o usuario
    /// pode ter digitado no meio). Tudo dentro de <see cref="EsperaDescarga"/>.</item>
    /// </list>
    /// </summary>
    /// <exception cref="InvalidOperationException">Sem documento; sessao congelada (recuperacao que falhou —
    /// falha na hora, sem esperar); o editor recusou; ou nao gravou em 10 s.</exception>
    public async Task DescarregarChangesAsync()
    {
        var ponte = _ponte;
        if (_sessao is null || ponte is null) throw new InvalidOperationException("Nenhum documento aberto.");
        if (FalhaDaRecuperacao is not null || ponte.Congelada) throw new InvalidOperationException(MensagemSessaoCongelada);
        var relogio = System.Diagnostics.Stopwatch.StartNew();
        var lote = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        void AoGravar() => lote.TrySetResult();
        // A sessao congelou no meio da espera (fix round 3: a gravacao do lote falhou): falha NA HORA, sem os 10 s.
        void AoCongelar() => lote.TrySetException(new InvalidOperationException(MensagemSessaoCongelada));
        ponte.ChangesSalvos += AoGravar;
        ponte.Congelou += AoCongelar;
        try
        {
            while (true)
            {
                var r = (await ExecutarJsAsync(PedirDescargaAoEditor)).Trim('"');
                if (r == "nada") return;
                if (r == "congelada") throw new InvalidOperationException(MensagemSessaoCongelada);
                if (r != "pediu") throw new InvalidOperationException("O editor não aceitou o pedido de gravar as alterações pendentes: " + r);
                var resta = EsperaDescarga - relogio.Elapsed;
                try
                {
                    if (resta <= TimeSpan.Zero) throw new TimeoutException();
                    await lote.Task.WaitAsync(resta);
                }
                catch (TimeoutException)
                {
                    throw new InvalidOperationException($"O editor não gravou as alterações pendentes em {EsperaDescarga.TotalSeconds:0}s.");
                }
                _arquivos.Log($"DESCARGA changes gravados pelo editor em {relogio.ElapsedMilliseconds} ms");
                // Um lote chegou: pergunta de novo (o usuario pode ter digitado no meio). O `lote` novo e' lido
                // pelo AoGravar, que captura a VARIAVEL.
                lote = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            }
        }
        finally { ponte.ChangesSalvos -= AoGravar; ponte.Congelou -= AoCongelar; }
    }

    /// <summary>Neutra quanto a' causa (fix round 3): a sessao congela tanto na reaplicacao recuperada que falhou quanto
    /// na gravacao de alteracoes que falhou.</summary>
    internal const string MensagemSessaoCongelada =
        "As alterações pendentes não podem ser gravadas: a sessão deste documento está congelada (o documento continua guardado para recuperação).";

    /// <summary>Script de <see cref="DescarregarChangesAsync"/>: "nada" | "pediu" | "congelada" | motivo da
    /// recusa. Roda no documento de cima e alcanca o quadro como os demais (ver <see cref="PedirSalvarAoEditor"/>).</summary>
    private const string PedirDescargaAoEditor = """
        (function () {
          try {
            var w = (typeof DesktopOfflineAppDocumentStartSave === "function") ? window : window.frames[0];
            if (!w) return "sem quadro do editor";
            if (w.__moffice && w.__moffice.congelado) return "congelada";
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api || typeof api.asc_Save !== "function") return "sem asc_Save";
            var h = w.AscCommon && w.AscCommon.History;
            var d = w.AscDesktopEditor;
            if (!h || !d) return "sem historico do editor";
            var temChanges = null;
            for (var k in h) {
              var f;
              try { f = h[k]; } catch (e) { continue; }
              if (typeof f === "function" && String(f).indexOf("LocalFileGetOpenChangesCount") >= 0) { temChanges = f; break; }
            }
            if (!temChanges) return "sem Have_Changes no historico";
            var contagem = d.LocalFileGetOpenChangesCount, salvo = d.LocalFileGetSaved, pendente;
            d.LocalFileGetOpenChangesCount = function () { return 0; };
            d.LocalFileGetSaved = function () { return true; };
            try { pendente = temChanges.call(h, true); }
            finally { d.LocalFileGetOpenChangesCount = contagem; d.LocalFileGetSaved = salvo; }
            if (!pendente) return "nada";
            api.asc_Save(true, false, true);
            return "pediu";
          } catch (e) { return "erro: " + e; }
        })()
        """;

    /// <summary>
    /// PDF de impressao do documento COMO ESTA' NO EDITOR (inclusive o que ainda nao foi salvo): descarrega os
    /// changes (<see cref="DescarregarChangesAsync"/>) e roda o x2t sobre a pasta de recuperacao (Editor.bin +
    /// changes0.json) para <c>&lt;Pastas.Impressao&gt;\&lt;guid&gt;.pdf</c>, com <see cref="OpcoesImpressao.JsonParaX2t"/>
    /// em <c>m_sJsonParams</c>. Nao salva o arquivo do usuario nem mexe em <see cref="Modificado"/>. O PDF e'
    /// apagado quando a janela fecha; se ficar pronto DEPOIS de a janela fechar, e' apagado na hora e o metodo
    /// lanca <see cref="InvalidOperationException"/> (nao ha' mais para quem entregar).
    /// </summary>
    public async Task<string> GerarPdfDeImpressaoAsync(OpcoesImpressao o)
    {
        var sessao = _sessao ?? throw new InvalidOperationException("Nenhum documento aberto.");
        if (_fechada) throw new InvalidOperationException(MensagemJanelaFechada);
        await DescarregarChangesAsync();
        var destino = PastaImpressao.NovoArquivo(Pastas.Impressao);
        try
        {
            await _conv.DeBinAsync(sessao.PastaRecovery, FormatoDocumento.Pdf, destino, jsonParams: o.JsonParaX2t);
        }
        catch
        {
            PastaImpressao.Apagar(destino);
            throw;
        }
        if (AntesDeRegistrarPdfParaTeste is { } costura) await costura();
        bool fechou;
        lock (_travaImpressao)
        {
            fechou = _fechada;
            if (!fechou) _pdfsDeImpressao.Add(destino);
        }
        if (fechou)
        {
            PastaImpressao.Apagar(destino);
            _arquivos.Log("IMPRESSAO PDF pronto depois de a janela fechar: apagado " + Path.GetFileName(destino));
            throw new InvalidOperationException(MensagemJanelaFechada);
        }
        _arquivos.Log($"IMPRESSAO PDF gerado {Path.GetFileName(destino)} ({new FileInfo(destino).Length} bytes, jsonParams {(o.JsonParaX2t is null ? "nenhum" : o.JsonParaX2t.Length + " caracteres")})");
        return destino;
    }

    /// <summary>Tratamento do <c>Print(json)</c> do editor. Fire-and-forget como os demais avisos: devolve o
    /// controle ao JS antes de tocar no WebView e nunca propaga excecao.</summary>
    private async Task ImprimirPedidoDoEditorAsync(string json)
    {
        await Task.Yield();
        OpcoesImpressao opcoes;
        string pdf;
        // A leitura do JSON tambem fica no try (fix round 1): qualquer falha vira log + ErroAoImprimir, nunca uma
        // excecao numa Task que ninguem observa.
        try
        {
            opcoes = OpcoesImpressao.Ler(json);
            pdf = await GerarPdfDeImpressaoAsync(opcoes);
        }
        catch (Exception ex)
        {
            _arquivos.Log("IMPRESSAO falhou: " + (ex is ErroConversao ec ? $"{ec.Message} [x2t {ec.CodigoSaida}] {ec.Stderr}" : ex.ToString()));
            // Janela ja' fechada: nao ha' faixa para mostrar nem usuario esperando — so' o log.
            if (_fechada) return;
            // So' mensagem pt-BR vai para o usuario (fix round 1 da Task 5): as nossas (InvalidOperationException) e as
            // do conversor (ErroConversao) ja' sao; qualquer outra (COM do WebView2, E/S...) fica no log acima.
            var mensagem = ex is InvalidOperationException or ErroConversao ? ex.Message : MensagemErroInesperado;
            try { ErroAoImprimir?.Invoke(this, mensagem); }
            catch (Exception ex2) { _arquivos.Log("IMPRESSAO erro no aviso de erro: " + ex2.Message); }
            return;
        }
        var pedido = ImprimirPedido;
        lock (_travaImpressao)
        {
            if (_fechada) return; // o Closed ja' apagou (ou apaga) o PDF; ninguem mais o manda para a impressora
            // Dono do PDF (Task 5): com assinante, o PDF passa a ser de quem imprime — a janela nao o apaga mais.
            if (pedido is not null) _pdfsDeImpressao.Remove(pdf);
        }
        if (pedido is null) { _arquivos.Log("IMPRESSAO PDF pronto, mas ninguem assina ImprimirPedido (fica com a janela)"); return; }
        _arquivos.Log("IMPRESSAO PDF entregue a quem imprime (a janela nao o apaga mais): " + Path.GetFileName(pdf));
        try { pedido.Invoke(this, opcoes, pdf); }
        catch (Exception ex) { _arquivos.Log("IMPRESSAO erro no aviso de impressao pedida: " + ex.Message); }
    }

    /// <summary>A lista de impressoras do Windows mandada ao painel no ultimo EndLoad (so' para os testes
    /// esperarem). `internal`.</summary>
    internal Task? ImpressorasEnviadasParaTeste { get; private set; }

    /// <summary>
    /// Lista de impressoras do Windows no combo "Impressora" do painel <b>Arquivo -&gt; Imprimir</b> (Task 5). O
    /// painel NAO pede a lista (medido no web-apps 9.4.0.129): o host EMPURRA <c>printer:config</c> por
    /// <c>on_native_message</c> no quadro do editor (ver <see cref="ImpressorasDoWindows"/>) — aqui, logo depois do
    /// EndLoad. A enumeracao pode levar segundos com impressoras de rede: roda fora da thread de interface. Nunca
    /// lanca (so' log): sem a lista o painel continua funcionando, com a impressora padrao.
    /// </summary>
    private async Task EnviarImpressorasAoEditorAsync(SessaoDocumento sessao)
    {
        try
        {
            var (padrao, lista) = await Task.Run(() => ImpressorasDoWindows.Listar(_arquivos.Log));
            if (_fechada || !ReferenceEquals(sessao, _sessao)) return;
            var r = (await ExecutarJsAsync(ImpressorasDoWindows.ScriptParaEditor(ImpressorasDoWindows.JsonParaEditor(padrao, lista)))).Trim('"');
            _arquivos.Log($"IMPRESSORAS {lista.Count} enviada(s) ao painel (padrao: {padrao ?? "nenhuma"}): {r}");
        }
        catch (Exception ex) { _arquivos.Log("IMPRESSORAS falha ao enviar ao painel: " + ex.Message); }
    }

    private const string MensagemJanelaFechada = "A janela foi fechada antes de o PDF de impressão ficar pronto.";

    /// <summary>Texto da falha que nao tem mensagem pt-BR propria (o detalhe tecnico fica no log da janela).</summary>
    internal const string MensagemErroInesperado = "ocorreu um erro inesperado (os detalhes ficaram no registro do mOffice).";

    // ---------------------------------------------------------------------------------------------------
    // Salvar. Caminho do app oficial, medido no sdkjs 9.4.0.129 (sdkjs/word/sdk-all.js):
    //   asc_Save (= Asc.asc_docs_api.prototype.n1) -> descarrega os changes pendentes (Fm.Umb -> HIe ->
    //   Fm.lFb, que leva a JIb.prototype.zXa -> AscDesktopEditor.LocalFileSaveChanges) -> window.Desktop
    //   OfflineAppDocumentStartSave(...) -> AscDesktopEditor.LocalFileSave(param, senha, docinfo, fileType,
    //   jsonOptions, senhaAntiga) -> o NATIVO converte e chama window.DesktopOfflineAppDocumentEndSave(
    //   erro, hash, senha) de volta (erro 0 = ok, 1 = nada gravado sem erro, 2 = falha de conversao).
    // Por isso NAO convertemos direto: pedimos o salvamento ao editor e so convertemos quando o LocalFileSave
    // chega — e' ele que garante que o ultimo lote de changes ja esta' em changes0.json.
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Salva no arquivo de origem, no mesmo formato. Devolve o que a gravacao fez (ver
    /// <see cref="ResultadoSalvar"/>); aqui <c>Gravou</c> e' sempre true — cancelamento so' existe onde ha'
    /// dialogo.</summary>
    public Task<ResultadoSalvar> SalvarAsync()
    {
        if (_sessao is null) throw new InvalidOperationException("Nenhum documento aberto.");
        // Documento novo nao tem arquivo de origem do usuario: gravar em _sessao.ArquivoOrigem escreveria por
        // cima da COPIA do modelo dentro da pasta de recuperacao — que e' apagada no fechamento. Quem chama
        // pela API do host (a casca) pergunta o destino e chama SalvarComoAsync; o Ctrl+S do editor cai no
        // mesmo lugar por SalvarPedidoDoEditorAsync, que forca o "salvar como".
        if (_sessao.Novo) throw new InvalidOperationException("Documento novo: use SalvarComoAsync.");
        return SalvarComoAsync(_sessao.ArquivoOrigem, _sessao.Formato);
    }

    /// <summary>Salva em <paramref name="destino"/> no formato <paramref name="fmt"/>, passando pelo editor
    /// (para que ele descarregue os changes pendentes antes da conversao). Devolve o que a gravacao fez —
    /// em particular se foi uma EXPORTACAO (PDF/PDF-A), que nao reponta a sessao e por isso nao serve como
    /// "salvei antes de fechar"; ver <see cref="ResultadoSalvar"/>.</summary>
    public async Task<ResultadoSalvar> SalvarComoAsync(string destino, FormatoDocumento fmt)
    {
        if (_sessao is null) throw new InvalidOperationException("Nenhum documento aberto.");
        // Fix round 2 da revisao da Task 2: sessao recuperada inconsistente nao se salva (ver FalharRecuperacao).
        if (FalhaDaRecuperacao is not null) throw new InvalidOperationException(FalhaDaRecuperacao);
        var salvamento = IniciarSalvamento();
        try { return await SalvarComoCorpoAsync(destino, fmt, salvamento.AvisosDeModificadoNoInicio); }
        finally { TerminarSalvamento(salvamento); }
    }

    private async Task<ResultadoSalvar> SalvarComoCorpoAsync(string destino, FormatoDocumento fmt, long avisosNoInicio)
    {
        var salvarComo = EhSalvarComo(_sessao!, destino, fmt);
        var tcs = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        // Cancela quem estava esperando antes (achado 8/9 da revisao 2026-09-17, fix review 1): sem isto, um
        // SalvarComoAsync que comecasse enquanto outro ainda esperava LocalFileSave deixava o TCS antigo preso
        // para sempre (ninguem mais o completa nem observa a excecao) — um vazamento silencioso de Task.
        var anterior = Interlocked.Exchange(ref _pedidoSalvar, tcs);
        anterior?.TrySetCanceled();
        string r;
        try
        {
            r = (await ExecutarJsAsync(PedirSalvarAoEditor.Replace("SALVARCOMO", salvarComo ? "true" : "false"))).Trim('"');
        }
        catch { Interlocked.CompareExchange(ref _pedidoSalvar, null, tcs); throw; }
        if (r != "ok")
        {
            Interlocked.CompareExchange(ref _pedidoSalvar, null, tcs);
            throw new InvalidOperationException("O editor não aceitou o pedido de salvamento: " + r);
        }
        // O LocalFileSave chega DENTRO da execucao do asc_Save (tudo sincrono no sdkjs), mas o await nao
        // custa nada e evita travar para sempre se um dia deixar de ser.
        await EsperarLocalFileSaveAsync(tcs);
        // Quem chamou pela API do host recebe a falha como excecao (com o codigo e o stderr do x2t intactos);
        // o editor ja foi avisado pelo EndSave(2) la dentro.
        var falha = await ConverterESinalizarAsync(destino, fmt, avisosNoInicio);
        if (falha is not null) throw falha;
        return new ResultadoSalvar(Gravou: true, Exportou: fmt.EhExportacao());
    }

    /// <summary>Espera o <c>LocalFileSave</c> do pedido de salvamento (<paramref name="tcs"/>) e mapeia falhas
    /// para excecoes com mensagem em pt-BR, liberando <see cref="_pedidoSalvar"/> se ainda apontar para este
    /// TCS. Extraido como metodo `internal` separado (achado 1 da revisao 2026-09-17, fix review 2) para dar
    /// um seam de teste: <see cref="SalvarComoAsync"/> inteiro precisa de janela real (WebView2) para rodar (o
    /// `ExecutarJsAsync` anterior pede o `asc_Save` de verdade), mas o mapeamento de excecao em torno do
    /// cancelamento — o `anterior?.TrySetCanceled()` de um <see cref="SalvarComoAsync"/> mais novo substituindo
    /// este pedido — nao precisa de WebView nenhuma, so de um <see cref="TaskCompletionSource{TResult}"/>
    /// cancelado. Sem este catch especifico, quem estava esperando via `await` recebia um
    /// <see cref="TaskCanceledException"/> cru (nao localizado, sem <see cref="UltimoErro"/> setado) em vez de
    /// um erro explicando o que aconteceu.</summary>
    internal async Task EsperarLocalFileSaveAsync(TaskCompletionSource<string> tcs)
    {
        try { await tcs.Task.WaitAsync(EsperaSalvar); }
        catch (TimeoutException)
        {
            Interlocked.CompareExchange(ref _pedidoSalvar, null, tcs);
            throw new InvalidOperationException($"O editor não chamou LocalFileSave em {EsperaSalvar.TotalSeconds:0}s.");
        }
        catch (OperationCanceledException)
        {
            // TaskCanceledException (o que TrySetCanceled() produz) deriva de OperationCanceledException.
            Interlocked.CompareExchange(ref _pedidoSalvar, null, tcs);
            UltimoErro = "O pedido de salvamento foi substituído por outro antes de o editor responder.";
            throw new InvalidOperationException(UltimoErro);
        }
        catch { Interlocked.CompareExchange(ref _pedidoSalvar, null, tcs); throw; }
    }

    /// <summary>Um "salvar como" de verdade e' um destino OU um formato diferentes do arquivo atual da sessao
    /// (minor da revisao 2026-09-17, fix review 2): antes, <see cref="SalvarComoAsync"/> e
    /// <see cref="ConverterESinalizarAsync"/> calculavam isto cada um do seu jeito — o primeiro com as duas
    /// condicoes (destino OU formato), o segundo so com o destino — entao um "salvar com o MESMO nome mas
    /// formato diferente" (ex.: origem.docx -> origem.odt escrito por cima, hipotetico) marcaria "salvar como"
    /// no pedido ao editor mas NAO repontaria a sessao nem avisaria o quadro. Uma unica regra usada nos dois
    /// lugares.</summary>
    private static bool EhSalvarComo(SessaoDocumento sessao, string destino, FormatoDocumento fmt) =>
        fmt != sessao.Formato
        || !string.Equals(Path.GetFullPath(destino), Path.GetFullPath(sessao.ArquivoOrigem), StringComparison.OrdinalIgnoreCase);

    /// <summary>Salvamento iniciado pelo EDITOR (Ctrl+S, botao Salvar, menu Salvar como). Fire-and-forget: roda
    /// dentro da chamada sincrona do JS, entao devolve o controle antes de tocar no WebView, e nunca propaga
    /// excecao — quem "chamou" foi o renderer, que so entende o EndSave.</summary>
    private async Task SalvarPedidoDoEditorAsync(string param, Salvamento salvamento)
    {
        await Task.Yield();
        try
        {
            if (_sessao is null) { await ChamarEndSaveAsync(2); return; }
            var (comoPedido, fileType) = PedidoSalvar.Ler(param);
            var destino = _sessao.ArquivoOrigem;
            var fmt = _sessao.Formato;
            // "Salvar como" pelo menu do editor: o sdkjs manda param = "saveas=true;" (e fileType != 0 quando o
            // chamador escolheu o formato em options.xV — asc_DownloadAs). Nos dois casos quem pergunta o
            // destino e' o host, como no app oficial.
            // Documento NOVO: o Ctrl+S do editor manda "salvar normal", mas nao ha arquivo de origem — e' o
            // host que pergunta o destino, como no "salvar como". (Task 2 do Marco 1 / Plano 2.)
            // Sessao recuperada DESVINCULADA (I2 da revisao final do Plano 3): o original mudou depois da queda, entao o
            // Ctrl+S tambem vira "Salvar como" (nome sugerido "<nome> (recuperado)").
            if (_sessao.Novo || _sessao.Desvinculada) comoPedido = true;
            var assinar = false;
            if (comoPedido || fileType != 0)
            {
                // Plano 4, Task 2: o dialogo mostra SEMPRE a linha inteira do tipo de documento (OOXML, ODF, PDF, PDF/A,
                // "PDF assinado") e o fileType do painel "Salvar como" do editor so' PRE-SELECIONA a linha (513 -> PDF,
                // 521 -> PDF/A; medidos no bundle 9.4: c_oAscFileType.PDF/PDFA). Antes, fileType 513 montava um filtro
                // so' com PDF. O painel do editor nao tem "assinado": a linha 5 so' e' escolhida no proprio dialogo.
                var escolha = EscolherDestino(_sessao.Formato, PedidoSalvar.FormatoPorCodigo(fileType));
                if (escolha is null) { await ChamarEndSaveAsync(1); return; }  // cancelou: nem sucesso nem erro
                (destino, fmt, assinar) = (escolha.Caminho, escolha.Formato, escolha.Assinar);
            }
            var falha = await ConverterESinalizarAsync(destino, fmt, salvamento.AvisosDeModificadoNoInicio);
            if (falha is null && assinar) PedirAssinatura(destino, fmt);
        }
        catch (Exception ex)
        {
            UltimoErro = ex.Message;
            _arquivos.Log("PONTE salvar pedido pelo editor falhou: " + ex.Message);
            try { await ChamarEndSaveAsync(2); } catch (Exception ex2) { _arquivos.Log("PONTE EndSave de erro tambem falhou: " + ex2.Message); }
        }
        finally { TerminarSalvamento(salvamento); }   // o inicio foi no aviso "localFileSave"; inclusive no "Salvar como" cancelado
    }

    /// <summary>Converte a pasta de recovery (Editor.bin + changes/changes0.json, com m_bFromChanges) para o
    /// destino e avisa o editor. Escrita ATOMICA: o x2t grava num temporario ao lado do destino e so entao o
    /// arquivo antigo e' substituido — uma conversao que falha no meio nao deixa o documento do usuario truncado.
    /// Devolve null se salvou, ou a excecao que derrubou o salvamento (ja sinalizada ao editor com EndSave(2)).</summary>
    private Task<Exception?> ConverterESinalizarAsync(string destino, FormatoDocumento fmt, long avisosNoInicio) =>
        ConverterESinalizarCorpoAsync(destino, fmt, avisosNoInicio);

    private async Task<Exception?> ConverterESinalizarCorpoAsync(string destino, FormatoDocumento fmt, long avisosNoInicio)
    {
        var s = _sessao!;
        var tmp = destino + ".moffice-tmp";
        Exception? falha = null;
        string? senha = null;
        long? geracaoLida = null;
        // Recusas da revisao final do Plano 3 (I1/I2): ANTES de qualquer E/S — o arquivo do usuario nao e' tocado, o
        // x2t nem roda. Saem pelos MESMOS canais de uma gravacao que falhou (EndSave 2 ao editor, Modificado
        // restaurado, excecao para quem chamou pela API do host) e, alem deles, pela faixa (AvisoAoUsuario).
        var recusa = MotivoParaRecusarGravacao(s, destino, fmt);
        if (recusa is not null)
        {
            falha = new InvalidOperationException(recusa);
            UltimoErro = recusa;
            _arquivos.Log("PONTE salvamento RECUSADO (" + Path.GetFileName(destino) + "): " + recusa);
        }
        else try
        {
            var pasta = Path.GetDirectoryName(Path.GetFullPath(destino));
            if (!string.IsNullOrEmpty(pasta)) Directory.CreateDirectory(pasta);
            // Task 7 do Plano 3: documento com senha sai com a MESMA senha (m_sSavePassword), senao o arquivo gravado
            // perderia a protecao. So' nos formatos de DOCUMENTO em que o x2t cifra (FormatoGuardaSenha: OOXML e ODF,
            // medidos); PDF/PDF-A exportado sai sem senha de proposito (PJe/e-SAJ recusam PDF com senha). A senha nao
            // aparece no log (so' o fato de ter ido).
            senha = X2tConversor.FormatoGuardaSenha(fmt) ? s.Senha : null;
            // Temporario VELHO de um salvamento que morreu no meio (processo morto, app encerrado com o salvamento em curso
            // antes do fix round 4): sai antes, mesmo somente-leitura — o x2t nao grava por cima dele.
            ApagarTemporarioVelho(tmp);
            // O que o x2t vai ler (fix rounds 3 e 4): se a ponte gravar alguma coisa depois daqui, o arquivo gravado nao tem
            // tudo — e a pasta de uma janela ja' fechada nao pode ser apagada no fim.
            geracaoLida = _ponte?.GeracaoDeGravacao;
            await _conv.DeBinAsync(s.PastaRecovery, fmt, tmp, senhaSalvar: senha);
            File.Move(tmp, destino, true);
            UltimoErro = null;
            if (DepoisDaConversaoParaTeste is { } costura) await costura();
            _arquivos.Log($"PONTE salvo {fmt} -> {Path.GetFileName(destino)} ({new FileInfo(destino).Length} bytes, changes={File.Exists(s.ArquivoChanges)}{(string.IsNullOrEmpty(senha) ? "" : ", protegido por senha")})");
        }
        catch (Exception ex)
        {
            falha = ex;
            UltimoErro = ex is ErroConversao ec ? $"{ec.Message} [x2t {ec.CodigoSaida}] {ec.Stderr}" : ex.Message;
            _arquivos.Log("PONTE erro ao salvar: " + UltimoErro);
        }
        finally { try { if (File.Exists(tmp)) { File.SetAttributes(tmp, FileAttributes.Normal); File.Delete(tmp); } } catch { } }
        // Protegido (fix round 3): com a janela fechada no meio (Ctrl+S e fechar logo em seguida) o EndSave nao tem mais
        // para quem ir, e uma excecao aqui pulava o resto — o registro da gravacao, os recentes e a limpeza da pasta.
        try { await ChamarEndSaveAsync(falha is null ? 0 : 2); }
        catch (Exception ex) { _arquivos.Log("PONTE EndSave nao entregue (a janela fechou?): " + ex.Message); }
        if (falha is not null)
        {
            // A gravacao FALHOU e o editor ja' tinha zerado Modificado no asc_Save (protocolo do Marco 0,
            // passo 6 antes do 7): sem devolver o sinal, o documento do usuario fica marcado como salvo sem
            // ter sido gravado em lugar nenhum — e o proximo Close() o descarta sem perguntar.
            RestaurarModificadoSeHouveEdicao();
            if (recusa is not null) AvisarUsuario(recusa);
            return falha;
        }
        // Documento protegido gravado num formato que nao guarda senha (fix round 1 da Task 7: hoje PDF/PDF-A): o arquivo
        // saiu SEM protecao, e o usuario precisa saber. So' avisa; nada mais muda.
        if (!string.IsNullOrEmpty(s.Senha) && !X2tConversor.FormatoGuardaSenha(fmt))
        {
            var aviso = MensagemSalvoSemSenha(fmt);
            _arquivos.Log("SENHA " + aviso);
            AvisarUsuario(aviso);
        }
        // Sessao recuperada de um protegido, ainda sem senha (I1): a gravacao (sempre noutro arquivo — o original
        // cifrado e' recusado acima) saiu SEM protecao. So' avisa.
        // Fix round 2: UMA vez por sessao e so' em salvamento de DOCUMENTO — o Ctrl+S seguinte da copia repetia o
        // aviso, e o PDF exportado nunca leva senha (a regra do PDF e' outra). A recusa de gravar o original cifrado
        // continua valendo ate' o Proteger (OriginalProtegidoSemSenha nao e' zerado aqui: ver o relatorio do fix round 2).
        else if (s.OriginalProtegidoSemSenha is not null && string.IsNullOrEmpty(senha) && !fmt.EhExportacao() && !s.AvisouSalvoSemSenha)
        {
            s.AvisouSalvoSemSenha = true;
            _arquivos.Log("SENHA " + MensagemRecuperadoSalvoSemSenha);
            AvisarUsuario(MensagemRecuperadoSalvoSemSenha);
        }
        // EXPORTACAO (PDF/PDF-A): grava o arquivo, responde EndSave(0) ao editor e para por aqui
        // (achado da revisao 2). Antes disto um "salvar como PDF" caia no mesmo caminho de um "salvar
        // como .odt": a sessao era REPONTADA para o .pdf — o titulo virava "relatorio.pdf", o
        // documento aparecia como NAO modificado embora as alteracoes so' existissem no PDF, e o
        // proximo Ctrl+S tentava regravar o PDF em vez do documento. Exportar nao e' salvar.
        if (fmt.EhExportacao())
        {
            // ANTES de avisar a casca: exportar nao grava o documento editavel, e o asc_Save ja' tinha
            // zerado Modificado (protocolo do Marco 0, passo 6 antes do 7). Sem devolver o sinal, um
            // "Salvar" que o usuario respondeu com PDF deixa o documento parecendo salvo — o titulo
            // perde o asterisco, o botao Salvar desliga e o proximo Close() apaga a recuperacao sem
            // perguntar. Quem escuta Exportado ja' le' a sessao no estado certo.
            RestaurarModificadoSeHouveEdicao();
            // Mesmo try/catch (e mesmo motivo) do ModificadoAlterado/Salvo abaixo: o EndSave(0) ja'
            // foi respondido, entao uma excecao de assinante subindo daqui viraria EndSave(2).
            try { Exportado?.Invoke(destino); }
            catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de exportado: " + ex.Message); }
            return falha;
        }

        // Fix round 5 da revisao final do Plano 3: o usuario digitou DURANTE o salvamento? (a ponte gravou changes depois de
        // o x2t ler a pasta, ou o editor avisou "modificado" depois do inicio). Entao o arquivo gravado NAO tem essas
        // alteracoes: o documento continua modificado — senao o asterisco sumia, o fechamento nao perguntava e elas se
        // perdiam em silencio.
        var editadoDurante = EditadoDuranteOSalvamento(geracaoLida, avisosNoInicio);
        if (editadoDurante) _arquivos.Log("PONTE houve edicao durante o salvamento: o documento continua modificado (o arquivo gravado nao a tem)");

        // Salvamento de DOCUMENTO que deu certo: e' o UNICO lugar que apaga o sinal do host. Daqui para
        // a frente so' uma edicao nova (AoMudarModificado(true)) o levanta de novo.
        _edicaoNaoGravada = editadoDurante;

        // "Salvar como" (achado 1 da revisao 2026-09-17, fix review 1): a sessao passa a apontar para o
        // arquivo NOVO, como um editor de verdade. Antes disto o metodo zerava `Modificado` em QUALQUER
        // sucesso enquanto `ArquivoOrigem` continuava nomeando o arquivo antigo — um `SalvarAsync()`
        // seguinte gravaria de volta no arquivo errado. Com o reponte, `Modificado = false` fica correto
        // nos dois casos (salvar normal E salvar como), porque o "arquivo atual" da sessao e' sempre o que
        // acabou de ser gravado.
        if (EhSalvarComo(s, destino, fmt))
        {
            s.Repontar(destino, fmt);
            try { await AtualizarSessaoNoEditorAsync(s.Titulo); }
            catch (Exception ex) { _arquivos.Log("PONTE atualizar sessao no quadro nao entregue (a janela fechou?): " + ex.Message); }
        }
        // Revisao final do Plano 3: a sessao passa a estar VINCULADA ao arquivo que acabou de gravar (I2 — a recusa
        // acima garante que nao foi o original alterado) e o sessao.json guarda o carimbo DELE, para uma proxima
        // queda saber se alguem mexeu no arquivo depois; e o "protegido" passa a dizer se ESTE arquivo saiu cifrado
        // (uma copia sem senha de uma sessao recuperada nao e' mais "documento protegido").
        s.Desvinculada = false;
        RegistrarGravacaoNaSessao(s, destino, protegido: !string.IsNullOrEmpty(senha));
        // O asterisco do titulo e o ModificadoAlterado(false) saem DAQUI (Task 2 do Marco 1 / Plano 2):
        // este e' o unico lugar que zera Modificado sem passar pelo onDocumentModifiedChanged do editor,
        // entao sem isto a casca continuaria mostrando "alterado" num documento recem-gravado.
        var era = s.Modificado;
        s.Modificado = editadoDurante;
        AtualizarTitulo();
        // try/catch (achado 8 da revisao 1): este metodo devolve `falha` e ja chamou EndSave(0). Uma
        // excecao de ASSINANTE subindo daqui viraria o catch de SalvarPedidoDoEditorAsync, que responde
        // EndSave(2) ao editor — um salvamento bem-sucedido reportado como erro por causa da casca.
        if (era != editadoDurante)
        {
            try { ModificadoAlterado?.Invoke(editadoDurante); }
            catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de modificado (apos salvar): " + ex.Message); }
        }
        // Mesmo try/catch e pelo mesmo motivo do ModificadoAlterado acima: este metodo ja' respondeu
        // EndSave(0) ao editor, entao uma excecao de ASSINANTE subindo daqui viraria o catch de
        // SalvarPedidoDoEditorAsync, que responderia EndSave(2) — um salvamento que deu certo reportado
        // ao editor como erro por causa da casca. Depois do Repontar: quem escuta le' a sessao ja' nova.
        try { Salvo?.Invoke(destino); }
        catch (Exception ex) { _arquivos.Log("PONTE erro no aviso de salvo: " + ex.Message); }
        ApagarRecoverySeAJanelaJaFechou(s, geracaoLida);
        return falha;
    }

    /// <summary>Regrava o <c>sessao.json</c> depois de um salvamento de documento bem-sucedido: carimbo do arquivo
    /// gravado (I2) e <see cref="MetadadosSessao.Protegido"/>. Melhor esforco, como o resto dos metadados.</summary>
    private void RegistrarGravacaoNaSessao(SessaoDocumento s, string gravado, bool protegido)
    {
        if (s.Metadados is not { } m) return;
        s.Metadados = m.ComCarimboDe(gravado) with { Protegido = protegido };
        try { s.Metadados.Gravar(s.PastaRecovery); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { _arquivos.Log("RECUPERACAO nao regravou sessao.json: " + e.Message); }
    }

    /// <summary>
    /// Fix round 3: a janela FECHOU durante este salvamento (Ctrl+S e fechar em seguida; encerramento do Windows no meio)
    /// e o <c>Closed</c> manteve a pasta porque havia salvamento em curso. Agora o salvamento de DOCUMENTO deu certo e
    /// gravou o arquivo atual da sessao: o trabalho esta' no arquivo do usuario, e a pasta — que senao reapareceria em
    /// "Documentos recuperados" com um documento ja' salvo — sai. So' sem nenhum outro estado pendente: sessao nao
    /// congelada, sem falha de recuperacao, e o <c>changes0.json</c> do mesmo tamanho de quando o x2t o leu (nada digitado
    /// no meio). A pasta e' apagada COM A TRAVA NA MAO (reaquirida): se outra janela a pegou nesse meio-tempo, fica.
    /// </summary>
    private void ApagarRecoverySeAJanelaJaFechou(SessaoDocumento s, long? geracaoLida)
    {
        if (!_fechada) return;
        if (FalhaDaRecuperacao is not null || _ponte?.Congelada == true || !ReferenceEquals(s, _sessao)) return;
        if (s.Modificado || _manterRecoveryAoFechar || RecuperacaoPendente) return;
        // Geracao de gravacao (fix round 4) no lugar do tamanho do changes0.json: truncar um change e gravar outro do
        // mesmo tamanho deixava o tamanho igual com o conteudo diferente.
        if (geracaoLida is null || _ponte?.GeracaoDeGravacao != geracaoLida)
        {
            _arquivos.Log("RECUPERACAO janela fechada durante o salvamento, mas houve alteracao depois da leitura do x2t: a pasta fica");
            return;
        }
        s.Trava ??= TravaSessao.Adquirir(s.PastaRecovery);
        if (s.Trava is null) { _arquivos.Log("RECUPERACAO janela fechada durante o salvamento: a pasta esta' com outra janela e fica"); return; }
        s.ApagarRecovery();
        _arquivos.Log("RECUPERACAO janela fechada durante o salvamento, que deu certo: a pasta sai (o trabalho esta' no arquivo)");
    }

    /// <summary>Houve edicao durante o salvamento (fix round 5): a ponte gravou changes depois do instantaneo tirado antes do
    /// x2t (<see cref="PonteDesktop.GeracaoDeGravacao"/>), ou o editor avisou "modificado" depois do inicio do salvamento.
    /// Sem instantaneo (nao chegou ao x2t) conta como edicao — o lado seguro.</summary>
    private bool EditadoDuranteOSalvamento(long? geracaoLida, long avisosNoInicio) =>
        geracaoLida is null
        || _ponte?.GeracaoDeGravacao != geracaoLida
        || Interlocked.Read(ref _avisosDeModificado) != avisosNoInicio;

    /// <summary>Costura de teste (fix round 4): roda depois de o arquivo do usuario ser gravado e antes do EndSave — o
    /// teste grava changes aqui ("alteracao depois da leitura do x2t"). `internal`.</summary>
    internal Func<Task>? DepoisDaConversaoParaTeste { get; set; }

    /// <summary>Apaga <paramref name="tmp"/> (o <c>&lt;destino&gt;.moffice-tmp</c>) se sobrou de antes, tirando o
    /// somente-leitura. Melhor esforco: se nao sair, o x2t reclama e o salvamento falha com a mensagem dele.</summary>
    private void ApagarTemporarioVelho(string tmp)
    {
        try
        {
            if (!File.Exists(tmp)) return;
            File.SetAttributes(tmp, FileAttributes.Normal);
            File.Delete(tmp);
            _arquivos.Log("PONTE temporario velho de um salvamento anterior apagado: " + Path.GetFileName(tmp));
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { _arquivos.Log("PONTE temporario velho nao saiu: " + e.Message); }
    }

    /// <summary><see cref="AvisoAoUsuario"/> com o try/catch dos demais eventos publicos.</summary>
    private void AvisarUsuario(string aviso)
    {
        try { AvisoAoUsuario?.Invoke(aviso); }
        catch (Exception ex) { _arquivos.Log("PONTE erro no aviso ao usuario: " + ex.Message); }
    }

    /// <summary>Depois de um "Salvar como" bem-sucedido, avisa o QUADRO do editor de que a sessao mudou de
    /// arquivo. ponte-desktop.js cacheia <c>window.__moffice.sessao</c> na primeira leitura (LocalFileGetSourcePath/
    /// getFontsSprite rodam antes de qualquer salvamento) — sem isto, LocalFileGetSourcePath() continuaria
    /// devolvendo o caminho antigo mesmo com <see cref="SessaoDocumento"/> ja repontada no lado C#. Limpa o
    /// cache (o proximo acesso re-le SessaoJson(), que ja reflete <see cref="SessaoDocumento.ArquivoOrigem"/>
    /// novo) e atualiza o titulo pelo mesmo caminho que o proprio editor usa (SetDocumentName).</summary>
    private async Task AtualizarSessaoNoEditorAsync(string novoTitulo)
    {
        var js = """
            (function () {
              try {
                var w = window.frames[0];
                if (!w || !w.__moffice) return "sem quadro";
                w.__moffice.sessao = null;
                if (w.AscDesktopEditor && typeof w.AscDesktopEditor.SetDocumentName === "function") w.AscDesktopEditor.SetDocumentName(TITULO);
                return "ok";
              } catch (e) { return "erro: " + e; }
            })()
            """.Replace("TITULO", System.Text.Json.JsonSerializer.Serialize(novoTitulo));
        var r = await ExecutarJsAsync(js);
        if (r.Trim('"') != "ok") _arquivos.Log("PONTE atualizar sessao no quadro falhou: " + r);
    }

    /// <summary>Costura de teste (fix round 3): o <c>EndSave</c> ao editor LANCA, como numa janela que fechou no meio do
    /// salvamento com a WebView2 ja' descartada. `internal`.</summary>
    internal bool FalharEndSaveParaTeste { get; set; }

    private async Task ChamarEndSaveAsync(int erro)
    {
        if (FalharEndSaveParaTeste) throw new InvalidOperationException("falha simulada do EndSave (teste)");
        // Como o EndLoad: DesktopOfflineAppDocumentEndSave vive no quadro do editor (sdkjs/word/sdk-all.js).
        // Sem esta chamada o editor fica preso no "long action" que o StartSave abriu (editor.by(...)).
        var js = """
            (function () {
              try {
                var w = (typeof DesktopOfflineAppDocumentEndSave === "function") ? window : window.frames[0];
                if (!w || typeof w.DesktopOfflineAppDocumentEndSave !== "function") return "sem EndSave";
                w.DesktopOfflineAppDocumentEndSave(ERRO, "", "");
                return "ok";
              } catch (e) { return "erro: " + e; }
            })()
            """.Replace("ERRO", erro.ToString());
        var r = await ExecutarJsAsync(js);
        if (r.Trim('"') != "ok") { UltimoErro = r; _arquivos.Log("PONTE EndSave falhou: " + r); }
    }

    /// <summary>asc_Save(a, b, d, f) = Asc.asc_docs_api.prototype.n1: entra se `d || b || documentoModificado`,
    /// e `b` tambem vira o isSaveAs do StartSave. Entao: salvar como = `asc_Save(false, true)` (param
    /// "saveas=true;"), salvar normal FORCADO = `asc_Save(false, false, true)` (param ""). Forcado porque um
    /// documento nao modificado simplesmente nao dispararia LocalFileSave e o host ficaria esperando.</summary>
    private const string PedirSalvarAoEditor = """
        (function () {
          try {
            var w = (typeof DesktopOfflineAppDocumentStartSave === "function") ? window : window.frames[0];
            if (!w) return "sem quadro do editor";
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api || typeof api.asc_Save !== "function") return "sem asc_Save";
            if (SALVARCOMO) api.asc_Save(false, true); else api.asc_Save(false, false, true);
            return "ok";
          } catch (e) { return "erro: " + e; }
        })()
        """;

    /// <summary>
    /// "Salvar como" COM o dialogo do host: pergunta o destino por <see cref="EscolherDestino"/> e grava.
    /// Devolve <c>Gravou = false</c> quando o usuario cancelou o dialogo (nao e' erro).
    ///
    /// O filtro do dialogo tambem oferece PDF: nesse caso o que acontece e' uma EXPORTACAO (a sessao
    /// continua no documento; ver <see cref="Exportado"/>), e o metodo devolve <c>Gravou = true</c> do mesmo
    /// jeito — gravou o que o usuario pediu — com <c>Exportou = true</c>, para quem precise distinguir (o
    /// caminho de FECHAMENTO precisa: ver <see cref="DecidirFechamentoAsync"/>).
    ///
    /// E' o unico "salvar como" que a casca deve chamar. Antes disto o mOffice.Comum montava um
    /// <c>SaveFileDialog</c> proprio, que divergia deste em tudo que importa: sem <c>Owner</c> (dialogo que
    /// se perde atras da janela), sem <c>InitialDirectory</c> (abria na ultima pasta do processo em vez da
    /// pasta do documento) e com 2 formatos em vez de 3 (sem PDF). Um unico dialogo, um unico comportamento.
    /// </summary>
    public async Task<ResultadoSalvar> SalvarComoComDialogoAsync()
    {
        if (_sessao is null) throw new InvalidOperationException("Nenhum documento aberto.");
        var escolha = EscolherDestino(_sessao.Formato);
        if (escolha is null) return ResultadoSalvar.Cancelado;
        // SalvarComoAsync LANCA quando a gravacao falha: daqui para baixo a exportacao deu certo.
        var r = await SalvarComoAsync(escolha.Caminho, escolha.Formato);
        if (escolha.Assinar) PedirAssinatura(escolha.Caminho, escolha.Formato);
        return r;
    }

    /// <summary>Dispara <see cref="AssinarPedido"/> depois de uma exportacao que DEU CERTO na linha "PDF assinado com
    /// mPDF". So' com formato de exportacao (defesa: a linha so' produz PDF, ver <see cref="DestinoDoDialogo"/>). O
    /// EndSave ja' foi respondido ao editor, entao uma excecao do assinante e' so' registrada.</summary>
    private void PedirAssinatura(string pdf, FormatoDocumento fmt)
    {
        if (!fmt.EhExportacao())
        {
            // Linha "PDF assinado com mPDF" com extensao de documento digitada (revisao final do Plano 4): gravou o
            // documento, nada vai para o mPDF — o usuario precisa saber.
            _arquivos.Log("ASSINAR pedido na linha do PDF assinado, mas o arquivo foi salvo como " + fmt.Extensao() + ": nada enviado ao mPDF");
            AvisarUsuario(MensagemAssinarSalvouDocumento(fmt));
            return;
        }
        _arquivos.Log("PDF exportado para assinar com o mPDF: " + Path.GetFileName(pdf));
        try { AssinarPedido?.Invoke(this, pdf); }
        catch (Exception ex) { _arquivos.Log("PONTE erro no pedido de assinatura: " + ex.Message); }
    }

    /// <summary>Costura de teste do dialogo de destino (achado Critico da revisao final do Plano 2): o
    /// <c>SaveFileDialog</c> e' modal — num teste de janela real ele travaria a thread STA —, entao o teste
    /// do caminho de fechamento substitui a ESCOLHA, mantendo todo o resto (filtro, gravacao, eventos) igual
    /// ao do produto. Recebe o FILTRO que o dialogo real mostraria: e' o que deixa o teste provar que o
    /// caminho de fechamento nao oferece PDF. `internal`, como <see cref="EscolherImagemParaTeste"/>: nao e'
    /// API do host.
    /// Recebe tambem o formato PRE-SELECIONADO (o que o dialogo real mostraria escolhido) e devolve o
    /// <see cref="DestinoSalvar"/> — com <c>Assinar</c>, o teste percorre a linha "PDF assinado com mPDF" (Plano 4).</summary>
    internal Func<string, FormatoDocumento, DestinoSalvar?>? EscolherDestinoParaTeste { get; set; }

    /// <summary>Dialogo de "Salvar como" do host (o editor nao tem dialogo proprio no build desktop).</summary>
    /// <param name="documento">Formato do documento da sessao: decide a linha do filtro (texto, planilha ou
    /// apresentacao — ver <see cref="FiltroSalvar"/>).</param>
    /// <param name="preSelecionado">Formato que o dialogo abre escolhido (o <c>fileType</c> do painel "Salvar como"
    /// do editor); <c>null</c> = o do documento.</param>
    /// <param name="permitirExportacao">Oferecer tambem PDF (exportacao). <c>false</c> no caminho de
    /// FECHAMENTO, onde exportar nao grava o documento editavel — ver <see cref="DecidirFechamentoAsync"/>.</param>
    private DestinoSalvar? EscolherDestino(FormatoDocumento documento, FormatoDocumento? preSelecionado = null, bool permitirExportacao = true)
    {
        var filtro = FiltroSalvar(documento, permitirExportacao);
        var sugerido = preSelecionado ?? documento;
        if (EscolherDestinoParaTeste is not null) return EscolherDestinoParaTeste(filtro, sugerido);
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title = "Salvar como",
            // Documento novo (Task 2 do Marco 1 / Plano 2): o "arquivo de origem" e' a copia do modelo dentro
            // da pasta de recuperacao — sugerir "new" como nome, na pasta de recuperacao, seria absurdo. O
            // nome vem do titulo generico ("Documento") e a pasta inicial, de Documentos do usuario.
            FileName = NomeSugerido(_sessao!),
            DefaultExt = sugerido.Extensao(),
            Filter = filtro,
            // Sem isto o diálogo abre sempre no PRIMEIRO filtro (OOXML) mesmo quando o formato sugerido é o
            // ODF ou o PDF — o usuário vê "Documento do Word" selecionado ao salvar um ODT.
            FilterIndex = IndiceDoFiltro(sugerido),
            AddExtension = true,
            OverwritePrompt = true,
        };
        var inicial = _sessao!.Novo
            ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            : Path.GetDirectoryName(Path.GetFullPath(_sessao.ArquivoOrigem));
        if (!string.IsNullOrEmpty(inicial)) dlg.InitialDirectory = inicial;
        if (dlg.ShowDialog(this) != true) return null;
        // O formato sai do INDICE da linha escolhida (as tres linhas de PDF tem a mesma extensao): ver DestinoDoDialogo.
        return DestinoDoDialogo(dlg.FileName, dlg.FilterIndex, documento, permitirExportacao);
    }

    /// <summary>
    /// Traduz o que o usuario escolheu no dialogo (<paramref name="caminho"/> + <c>FilterIndex</c> base 1) no
    /// <see cref="DestinoSalvar"/>. Puro (`internal static`, testado sem janela). Linhas de <see cref="FiltroSalvar"/>:
    /// 1 = OOXML, 2 = ODF, 3 = PDF, 4 = PDF/A, 5 = "PDF assinado com mPDF".
    /// <list type="bullet">
    /// <item>1 e 2: pela extensao, como sempre foi (extensao desconhecida = <paramref name="sugerido"/>).</item>
    /// <item>3, 4 e 5: pelo indice. Uma extensao de DOCUMENTO digitada vence (o usuario escreveu "a.odt" com a linha
    /// PDF escolhida: grava ODT, comportamento de antes); na linha 5 o <c>Assinar</c> vem junto, para o host avisar que
    /// nada foi enviado para assinar (<see cref="MensagemAssinarSalvouDocumento"/>). Senao o nome termina em <c>.pdf</c> (acrescentado se faltar
    /// — nunca grava PDF com outra extensao).</item>
    /// <item>Sem exportacao (caminho do FECHAMENTO) ou sessao que ja' e' PDF (filtro de uma linha so'): nao ha linhas
    /// de PDF por indice, tudo pela extensao — o fechamento nao muda.</item>
    /// </list>
    /// </summary>
    internal static DestinoSalvar DestinoDoDialogo(string caminho, int filterIndex, FormatoDocumento sugerido, bool permitirExportacao)
    {
        FormatoDocumento? daExtensao;
        try { daExtensao = FormatoDocumentoExtensoes.DaExtensao(caminho); }
        catch (ErroConversao) { daExtensao = null; }  // o usuario digitou o nome inteiro com extensao desconhecida
        var linhaDePdf = permitirExportacao && !sugerido.EhExportacao()
            && filterIndex is IndiceFiltroPdf or IndiceFiltroPdfA or IndiceFiltroPdfAssinado;
        if (!linhaDePdf) return new DestinoSalvar(caminho, daExtensao ?? sugerido);
        // Extensao de documento digitada na linha "PDF assinado": grava o DOCUMENTO e leva o Assinar junto so' para o
        // host AVISAR que nada foi para assinar (revisao final do Plano 4; ver PedirAssinatura).
        if (daExtensao is { } doc && !doc.EhExportacao()) return new DestinoSalvar(caminho, doc, Assinar: filterIndex == IndiceFiltroPdfAssinado);
        var pdf = daExtensao == FormatoDocumento.Pdf ? caminho : caminho + ".pdf";
        return filterIndex switch
        {
            IndiceFiltroPdfA => DestinoPdfA(pdf),
            IndiceFiltroPdfAssinado => new DestinoSalvar(pdf, FormatoDocumento.Pdf, Assinar: true),
            _ => new DestinoSalvar(pdf, FormatoDocumento.Pdf),
        };
    }

    /// <summary>
    /// Como a linha "PDF/A" do "Salvar como" gera o arquivo — o UNICO lugar dessa decisao (S6 do Plano 4).
    /// Hoje e' a decisao <b>A</b>: o proprio x2t exporta PDF/A (<see cref="FormatoDocumento.PdfA"/>), pelo mesmo
    /// caminho de exportacao do PDF (sessao nao repontada, <see cref="Exportado"/>, sem senha). O pedido real ao x2t e'
    /// <c>m_nFormatTo = 513</c> + <c>m_bIsPDFA = true</c> (<see cref="X2tConversor"/>): 521 a partir do Editor.bin sai com
    /// codigo 88 (MEDIDO). O arquivo DECLARA PDF/A-2A no XMP (<c>pdfaid:part 2</c>, <c>conformance A</c>); se ele e'
    /// conforme e' a Task 1/S6 que valida, no veraPDF, contra esse perfil declarado. Se reprovar (decisao <b>B</b>), a
    /// troca e' SO' aqui: devolver
    /// <c>FormatoDocumento.Pdf</c> com uma marca de "converter pelo mPDF (/pdfa)" no <see cref="DestinoSalvar"/>.
    /// </summary>
    private static DestinoSalvar DestinoPdfA(string pdf) => new(pdf, FormatoDocumento.PdfA);

    /// <summary>Nome (sem extensao) que o dialogo "Salvar como" sugere: o titulo generico de um documento novo; o
    /// nome do original com " (recuperado)" numa sessao recuperada desvinculada (I2 da revisao final do Plano 3 —
    /// aceitar o nome sugerido nunca sobrescreve o original alterado); senao o nome do original. `internal static`
    /// para o teste puro.</summary>
    internal static string NomeSugerido(SessaoDocumento s) =>
        s.Novo ? s.Titulo
        : s.Desvinculada ? Path.GetFileNameWithoutExtension(s.ArquivoOrigem) + " (recuperado)"
        : Path.GetFileNameWithoutExtension(s.ArquivoOrigem);

    private const string FiltroTexto = "Documento do Word (*.docx)|*.docx|Documento de texto ODF (*.odt)|*.odt";
    private const string FiltroPdf = "PDF (*.pdf)|*.pdf";
    // Plano 4, Task 2: as tres linhas de PDF tem a MESMA extensao — o formato sai do indice (DestinoDoDialogo).
    private const string FiltroPdfA = "PDF/A (*.pdf)|*.pdf";
    private const string FiltroPdfAssinado = "PDF assinado com mPDF (*.pdf)|*.pdf";
    private const int IndiceFiltroPdf = 3, IndiceFiltroPdfA = 4, IndiceFiltroPdfAssinado = 5;

    /// <summary>
    /// Filtro do <c>SaveFileDialog</c> para o formato do documento. A ordem de cada linha e' OOXML, ODF e — so'
    /// com <paramref name="permitirExportacao"/> — PDF, PDF/A e "PDF assinado com mPDF" (Plano 4, Task 2). O
    /// <see cref="IndiceDoFiltro"/> e o <see cref="DestinoDoDialogo"/> dependem dessa ordem: as tres linhas de PDF
    /// tem a mesma extensao, entao o formato escolhido so' se sabe pelo indice.
    ///
    /// <paramref name="permitirExportacao"/> e' <c>false</c> no caminho de FECHAMENTO: ali o usuario pediu
    /// "Salvar" para NAO perder o documento, e PDF nao salva documento nenhum — e' exportacao, a sessao
    /// continua no original e a janela fecharia apagando a recuperacao (achado Critico da revisao final do
    /// Plano 2). `internal` para o teste puro poder comparar as duas versoes do filtro sem abrir janela.
    ///
    /// Formato de exportacao com <paramref name="permitirExportacao"/> falso (uma sessao cujo formato e' PDF,
    /// que este editor nem abre) nao tem par editavel: cai no filtro de TEXTO, o unico destino que ainda
    /// preserva o que o usuario escreveu. Nunca devolve filtro vazio — um dialogo sem filtro nenhum seria
    /// pior do que um destino inesperado.
    /// </summary>
    internal static string FiltroSalvar(FormatoDocumento f, bool permitirExportacao = true)
    {
        var documento = f.DocType() switch
        {
            "word" => FiltroTexto,
            "cell" => "Pasta de trabalho do Excel (*.xlsx)|*.xlsx|Planilha ODF (*.ods)|*.ods",
            "slide" => "Apresentação do PowerPoint (*.pptx)|*.pptx|Apresentação ODF (*.odp)|*.odp",
            _ => null,  // PDF: nao e' formato de DOCUMENTO — so' existe como exportacao
        };
        if (!permitirExportacao) return documento ?? FiltroTexto;
        return documento is null ? FiltroPdf : documento + "|" + FiltroPdf + "|" + FiltroPdfA + "|" + FiltroPdfAssinado;
    }

    /// <summary>Posição do formato dentro de <see cref="FiltroSalvar"/> — <b>base 1</b>, que é o que o
    /// <c>SaveFileDialog.FilterIndex</c> usa: 1 = OOXML, 2 = ODF, 3 = PDF, 4 = PDF/A, 5 = PDF com
    /// <paramref name="assinar"/>. Numa sessão que já é PDF (filtro de uma linha só, que este editor nem abre) o
    /// índice de PDF passa do fim, e o diálogo cai na primeira linha — que ali é o próprio PDF.</summary>
    internal static int IndiceDoFiltro(FormatoDocumento f, bool assinar = false) => f switch
    {
        FormatoDocumento.Odt or FormatoDocumento.Ods or FormatoDocumento.Odp => 2,
        FormatoDocumento.Pdf => assinar ? IndiceFiltroPdfAssinado : IndiceFiltroPdf,
        FormatoDocumento.PdfA => IndiceFiltroPdfA,
        _ => 1,
    };
}
