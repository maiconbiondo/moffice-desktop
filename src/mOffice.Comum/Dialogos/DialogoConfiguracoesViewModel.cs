using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using mOffice.Comum.Atualizacao;

namespace mOffice.Comum.Dialogos;

/// <summary>Estado do bloco "Atualizações" das Configurações.</summary>
public enum EstadoAtualizacao { Ocioso, Verificando, Atualizado, Disponivel, Baixando, Pronto, Erro }

/// <summary>Configuracoes do app: tema (3 opcoes), o grampo de log da ponte e "Verificar atualizações". Cada set
/// grava no <see cref="Configuracoes"/> (que grava no disco na hora, sem cache) e, no caso do tema, reaplica os
/// tokens ao vivo.
///
/// Atualizacao (Plano 5, Task 2; porte do ConfiguracoesViewModel do mPDF): REDE SO' POR CLIQUE — o
/// <see cref="ServicoAtualizacao"/> so' e' construido dentro de <see cref="VerificarAtualizacoesAsync"/> (nunca no
/// construtor: abrir o dialogo nao toca a rede). "Versão Y disponível. Instalar agora?" -> Sim -> baixa e verifica o
/// SHA-256 -> <see cref="InstaladorPronto"/> e pede para o dialogo fechar; quem fecha os documentos e roda o
/// instalador e' o <see cref="AplicativoOffice"/> (pelo fluxo de fechar de sempre).</summary>
public sealed partial class DialogoConfiguracoesViewModel : ObservableObject
{
    private readonly Configuracoes _config; private readonly TemaService _tema;

    // Atribui os CAMPOS direto, sem passar pelos partial OnXChanged: abrir o dialogo nao pode reaplicar o
    // tema nem regravar o config — so' LE o que ja' esta la'.
    private readonly Func<IFonteDeReleases> _criarFonte;
    private readonly Func<string, bool> _confirmar;
    private readonly Func<IReadOnlyList<string>> _outrosAppsAbertos;
    private readonly string _nomeApp;

    /// <param name="criarFonte">De onde vem a ultima release. Padrao: a API do GitHub. Os testes passam um falso.</param>
    /// <param name="confirmar">A pergunta Sim/Nao. Padrao: MessageBox do sistema (os testes passam um falso — a padrao
    /// abre uma janela modal).</param>
    /// <param name="outrosAppsAbertos">Os OUTROS apps da suite abertos agora (producao:
    /// <see cref="OutrosAppsDaSuite.Abertos(PerfilApp)"/>). Com algum aberto, "Instalar agora" nao e' oferecido: o
    /// instalador o fecharia sem perguntar se salva. Padrao: nenhum.</param>
    /// <param name="nomeApp">O app deste dialogo (mTexto, mPlanilha, mSlides), para a pergunta.</param>
    public DialogoConfiguracoesViewModel(Configuracoes config, TemaService tema,
        Func<IFonteDeReleases>? criarFonte = null, Func<string, bool>? confirmar = null,
        Func<IReadOnlyList<string>>? outrosAppsAbertos = null, string nomeApp = "mOffice")
    {
        _config = config; _tema = tema; _modoTema = config.Tema; _grampo = config.Grampo;
        _criarFonte = criarFonte ?? ServicoAtualizacao.CriarFontePadrao;
        _confirmar = confirmar ?? ConfirmarNoSistema;
        _outrosAppsAbertos = outrosAppsAbertos ?? (() => []);
        _nomeApp = nomeApp;
    }

    [ObservableProperty] private ModoTema _modoTema;
    partial void OnModoTemaChanged(ModoTema value)
    {
        _config.Tema = value; _tema.Aplicar(value);
        // Roda DEPOIS da atribuicao: as tres booleanas dos RadioButton ja' refletem o valor novo.
        OnPropertyChanged(nameof(Sistema)); OnPropertyChanged(nameof(Claro)); OnPropertyChanged(nameof(Escuro));
    }

    [ObservableProperty] private bool _grampo;

    /// <summary>Grava a configuracao E aplica ao estado vivo do processo (achado da revisao final do Plano 2):
    /// sem a segunda linha, ligar o grampo nas Configuracoes so' tinha efeito no proximo arranque do app —
    /// <c>JanelaEditor.GrampoLigado</c> e' lido na subida (<c>AplicativoOffice.OnStartup</c>) e nunca mais.
    /// O efeito vale so' para JANELAS NOVAS: o grampo e' injetado em <c>InicializarAsync</c>, uma vez por
    /// janela, entao as janelas ja' abertas continuam como estavam ate' serem fechadas e reabertas.</summary>
    partial void OnGrampoChanged(bool value)
    {
        _config.Grampo = value;
        mOffice.Nucleo.JanelaEditor.GrampoLigado = value;
    }

    // Os tres RadioButton ligam aqui (sem converter): cada set e' um Aplicar.
    public bool Sistema { get => ModoTema == ModoTema.Sistema; set { if (value) ModoTema = ModoTema.Sistema; } }
    public bool Claro { get => ModoTema == ModoTema.Claro; set { if (value) ModoTema = ModoTema.Claro; } }
    public bool Escuro { get => ModoTema == ModoTema.Escuro; set { if (value) ModoTema = ModoTema.Escuro; } }

    // ---- Atualizacoes ---------------------------------------------------------------------------------

    /// <summary>A versao instalada, para o rotulo do bloco.</summary>
    public string VersaoInstalada => ServicoAtualizacao.VersaoAtualTexto();

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(VerificarAtualizacoesCommand))]
    private EstadoAtualizacao _estadoAtualizacao = EstadoAtualizacao.Ocioso;

    /// <summary>O que o bloco mostra abaixo do botao (versao mais recente, erro, "Baixando…"). Nulo = nada.</summary>
    [ObservableProperty] private string? _mensagemAtualizacao;

    /// <summary>Preenchido quando o usuario disse "Sim" e o instalador ja' esta' baixado e VERIFICADO: quem abriu o
    /// dialogo le' isto depois que ele fecha e fecha o app para instalar.</summary>
    public ServicoAtualizacao.ArquivoVerificado? InstaladorPronto { get; private set; }

    /// <summary>O dialogo deve fechar (o instalador esta' pronto). O code-behind assina.</summary>
    public event Action? PedirFechar;

    private bool PodeVerificar() => EstadoAtualizacao is not (EstadoAtualizacao.Verificando or EstadoAtualizacao.Baixando or EstadoAtualizacao.Pronto);

    /// <summary>"Verificar atualizações": o UNICO ponto deste VM que constroi o <see cref="ServicoAtualizacao"/>.</summary>
    [RelayCommand(CanExecute = nameof(PodeVerificar))]
    private async Task VerificarAtualizacoesAsync()
    {
        EstadoAtualizacao = EstadoAtualizacao.Verificando;
        MensagemAtualizacao = "Verificando…";
        using var servico = new ServicoAtualizacao(_criarFonte());
        ResultadoVerificacao r;
        try { r = await servico.VerificarAsync(); }
        catch (Exception) { r = ResultadoVerificacao.Erro(TipoErroAtualizacao.SemRede, ServicoAtualizacao.MensagemSemRede); }

        switch (r.Status)
        {
            case StatusVerificacao.Atualizado:
                EstadoAtualizacao = EstadoAtualizacao.Atualizado;
                MensagemAtualizacao = MensagemMaisRecente(VersaoInstalada);
                return;
            case StatusVerificacao.Disponivel:
                var versao = ServicoAtualizacao.VersaoDaTag(r.Info!.TagVersao);
                EstadoAtualizacao = EstadoAtualizacao.Disponivel;
                MensagemAtualizacao = $"Versão {versao} disponível.";
                // Outro app da suite aberto: nem pergunta (o instalador o fecharia sem perguntar se salva).
                if (_outrosAppsAbertos() is { Count: > 0 } outros)
                {
                    MensagemAtualizacao = $"Versão {versao} disponível. " + OutrosAppsDaSuite.MensagemFecharOutros(outros);
                    return;
                }
                if (!_confirmar(PerguntaInstalar(versao, _nomeApp))) return;
                await BaixarAsync(servico, r.Info, versao);
                return;
            default:
                EstadoAtualizacao = EstadoAtualizacao.Erro;
                MensagemAtualizacao = r.MensagemErro;
                return;
        }
    }

    // O download (o instalador tem ~1 GB) e' cancelado se o dialogo fechar no meio: sem isto ele seguiria em segundo
    // plano, sem ninguem para instalar, e o arquivo ficaria no %TEMP%. O servico apaga o parcial no cancelamento.
    private CancellationTokenSource? _cancelarDownload;

    /// <summary>O dialogo fechou: cancela um download em curso (nada acontece fora dele).</summary>
    public void CancelarDownload() => _cancelarDownload?.Cancel();

    private async Task BaixarAsync(ServicoAtualizacao servico, InfoAtualizacao info, string versao)
    {
        EstadoAtualizacao = EstadoAtualizacao.Baixando;
        MensagemAtualizacao = $"Baixando a versão {versao}…";
        var progresso = new Progress<(long Baixados, long Total)>(p =>
        {
            if (EstadoAtualizacao == EstadoAtualizacao.Baixando && p.Total > 0)
                MensagemAtualizacao = $"Baixando a versão {versao}… {Math.Min(100, p.Baixados * 100 / p.Total)}%";
        });
        using var cts = _cancelarDownload = new CancellationTokenSource();
        ResultadoDownload d;
        try { d = await servico.BaixarEVerificarAsync(info, progresso, cts.Token); }
        catch (OperationCanceledException)
        {
            EstadoAtualizacao = EstadoAtualizacao.Disponivel;
            MensagemAtualizacao = "Download cancelado.";
            return;
        }
        catch (Exception) { d = ResultadoDownload.Falhou("Não foi possível baixar a atualização. Confira a conexão com a internet."); }
        finally { _cancelarDownload = null; }

        if (d.Status != StatusDownload.Verificado || d.Arquivo is null)
        {
            EstadoAtualizacao = EstadoAtualizacao.Erro;
            MensagemAtualizacao = d.MensagemErro;
            return;
        }
        InstaladorPronto = d.Arquivo;
        EstadoAtualizacao = EstadoAtualizacao.Pronto;
        MensagemAtualizacao = "Fechando o mOffice para instalar…";
        PedirFechar?.Invoke();
    }

    internal static string MensagemMaisRecente(string versao) => $"Você está com a versão mais recente ({versao}).";

    /// <summary>So' promete o que acontece: as janelas DESTE app fecham pelo fluxo normal, o Windows pede permissao, o
    /// instalador mostra o progresso (<c>/SILENT</c>) e reabre este app no fim (<c>/REABRIR=</c>). Pede para nao abrir o
    /// mOffice no meio (revisao final do Plano 5, I-1: abrir durante a instalacao trava arquivos e quebra a instalacao).</summary>
    internal static string PerguntaInstalar(string versao, string nomeApp) =>
        $"Versão {versao} disponível. Instalar agora?\n\n" +
        $"As janelas do {nomeApp} serão fechadas (documentos com alterações perguntam se devem ser salvos) e o Windows vai pedir permissão para instalar. " +
        $"O instalador mostra o progresso e o {nomeApp} reabre sozinho em um ou dois minutos. Não abra o mOffice enquanto isso.";

    private static bool ConfirmarNoSistema(string pergunta)
    {
        var dono = Application.Current?.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive);
        const string titulo = "Atualização do mOffice";
        var r = dono is null
            ? MessageBox.Show(pergunta, titulo, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No)
            : MessageBox.Show(dono, pergunta, titulo, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No);
        return r == MessageBoxResult.Yes;
    }
}
