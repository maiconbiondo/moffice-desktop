using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Comum.Telas;

/// <summary>Tela de boas-vindas (estado sem documento). Todo efeito colateral — criar um documento, abrir um
/// arquivo, escolher um arquivo no dialogo do Windows, listar/recuperar/descartar documentos recuperados — e' um
/// DELEGADO injetado, entao o VM roda em teste sem janela e sem tocar disco alem do arquivo de recentes. UMA
/// excecao: a pergunta de descarte (<c>confirmarDescarte</c>) tem PADRAO — o MessageBox do sistema
/// (<see cref="ConfirmarDescarteNoSistema"/>) —, entao um teste que exercite o Descartar SEM passar um falso abre
/// uma janela modal e trava. Os testes passam sempre um falso.</summary>
public sealed partial class TelaBoasVindasViewModel : ObservableObject
{
    private readonly PerfilApp _perfil; private readonly Recentes _recentes;
    private readonly Func<Task> _novo; private readonly Func<string, Task> _abrir; private readonly Func<Task<string?>> _escolher;

    private readonly Action? _configuracoes; private readonly Action? _sobre;

    private readonly Func<IReadOnlyList<DocumentoRecuperavel>>? _recuperaveis;
    private readonly Func<DocumentoRecuperavel, Task>? _recuperar;
    private readonly Action<DocumentoRecuperavel>? _descartar;
    private readonly Func<DocumentoRecuperavel, bool> _confirmarDescarte;

    /// <param name="configuracoes">Abre as Configuracoes do mOffice; null = link escondido.</param>
    /// <param name="sobre">Abre o Sobre (atribuicao de licencas); null = link escondido.</param>
    /// <param name="recuperaveis">Documentos que uma sessao anterior deixou sem salvar (Task 3 do Plano 3); null =
    /// secao "Documentos recuperados" sempre escondida.</param>
    /// <param name="recuperar">Reabre um documento recuperavel numa janela de documento.</param>
    /// <param name="descartar">Apaga um documento recuperavel (so' depois de <paramref name="confirmarDescarte"/>).</param>
    /// <param name="confirmarDescarte">Pergunta antes de descartar; true = pode apagar. Null = a pergunta do
    /// sistema (<see cref="ConfirmarDescarteNoSistema"/> sem dono, MessageBox Sim/Nao). Os testes SEMPRE passam um
    /// falso: a padrao abre uma janela modal.</param>
    public TelaBoasVindasViewModel(PerfilApp perfil, Recentes recentes, Func<Task> novo, Func<string, Task> abrir,
        Func<Task<string?>> escolherArquivo, Action? configuracoes = null, Action? sobre = null,
        Func<IReadOnlyList<DocumentoRecuperavel>>? recuperaveis = null, Func<DocumentoRecuperavel, Task>? recuperar = null,
        Action<DocumentoRecuperavel>? descartar = null, Func<DocumentoRecuperavel, bool>? confirmarDescarte = null)
    {
        _perfil = perfil; _recentes = recentes; _novo = novo; _abrir = abrir; _escolher = escolherArquivo;
        _configuracoes = configuracoes; _sobre = sobre;
        _recuperaveis = recuperaveis; _recuperar = recuperar; _descartar = descartar;
        _confirmarDescarte = confirmarDescarte ?? (d => ConfirmarDescarteNoSistema(null, d));
        Recarregar();
    }

    /// <summary>O perfil do app, para o code-behind filtrar o que foi SOLTO na tela (PrimeiroAceito).</summary>
    public PerfilApp Perfil => _perfil;
    public string Titulo => _perfil.NomeExibicao;
    public string Subtitulo => _perfil.Subtitulo;
    public string ChaveLogo => _perfil.ChaveLogo;
    public ObservableCollection<ItemRecente> Recentes { get; } = [];
    [ObservableProperty] private string? _mensagemErro;

    /// <summary>Recarrega as duas listas (recentes e recuperados): a boas-vindas chama isto toda vez que volta
    /// a aparecer — enquanto ela estava escondida, uma janela pode ter recuperado, salvo ou caido.</summary>
    public void Recarregar()
    {
        Recentes.Clear(); foreach (var i in _recentes.Listar()) Recentes.Add(i);
        RecarregarRecuperados();
    }

    // ---------------------------------------------------------------------------------------------------
    // Documentos recuperados (Task 3 do Plano 3)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Documentos que o app fechou sem salvar (queda), do mais recente para o mais antigo.</summary>
    public ObservableCollection<DocumentoRecuperavel> Recuperados { get; } = [];

    /// <summary>A secao "Documentos recuperados" aparece?</summary>
    public bool TemRecuperados => Recuperados.Count > 0;

    private void RecarregarRecuperados()
    {
        Recuperados.Clear();
        // A fonte e' o disco (Recuperaveis.Listar, que nunca lanca); o try so' protege a tela de um delegado
        // de teste ou de uma falha inesperada — a lista de recuperados nao pode derrubar a boas-vindas.
        try { if (_recuperaveis is not null) foreach (var d in _recuperaveis()) Recuperados.Add(d); }
        catch (Exception ex) { MensagemErro = $"Não foi possível listar os documentos recuperados: {ex.Message}"; }
        OnPropertyChanged(nameof(TemRecuperados));
    }

    /// <summary>Reabre o documento. Depois (deu certo ou nao) a lista e' relida do disco: a janela que recuperou
    /// segura a trava da pasta, entao o item some sozinho; se falhou, ele continua ali, com Descartar.</summary>
    [RelayCommand] private async Task Recuperar(DocumentoRecuperavel d)
    {
        MensagemErro = null;
        try { if (_recuperar is not null) await _recuperar(d); }
        catch (Exception ex) { MensagemErro = $"Não foi possível recuperar o documento: {ex.Message}"; }
        finally { RecarregarRecuperados(); }
    }

    /// <summary>Pergunta e, com o "Sim", apaga. Relido do disco depois: se o item ainda estiver la' (arquivo
    /// preso, ou alguem abriu a pasta no meio), avisa em vez de fingir que sumiu.</summary>
    [RelayCommand] private void Descartar(DocumentoRecuperavel d)
    {
        if (_descartar is null) return;
        MensagemErro = null;
        if (!_confirmarDescarte(d)) return;
        try { _descartar(d); }
        catch (Exception ex) { MensagemErro = $"Não foi possível descartar \"{d.Titulo}\": {ex.Message}"; }
        RecarregarRecuperados();
        if (MensagemErro is null && Recuperados.Any(r => string.Equals(r.Pasta, d.Pasta, StringComparison.OrdinalIgnoreCase)))
            MensagemErro = $"Não foi possível descartar \"{d.Titulo}\" agora. Tente de novo em instantes.";
    }

    /// <summary>O texto da pergunta de descarte (spec da Task 3 do Plano 3).</summary>
    public static string PerguntaDescartar(DocumentoRecuperavel d) =>
        $"Descartar as alterações não salvas de {d.Titulo}? Esta ação não pode ser desfeita.";

    /// <summary>A pergunta padrao: MessageBox do sistema, Sim/Nao, com o "Nao" como padrao (descartar e'
    /// irreversivel). <paramref name="dono"/>: a janela sobre a qual a pergunta abre (sem dono ela pode ficar
    /// ATRAS da boas-vindas).</summary>
    public static bool ConfirmarDescarteNoSistema(System.Windows.Window? dono, DocumentoRecuperavel d)
    {
        const string titulo = "Descartar documento recuperado";
        var r = dono is null
            ? System.Windows.MessageBox.Show(PerguntaDescartar(d), titulo, System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.No)
            : System.Windows.MessageBox.Show(dono, PerguntaDescartar(d), titulo, System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning, System.Windows.MessageBoxResult.No);
        return r == System.Windows.MessageBoxResult.Yes;
    }

    /// <summary>UNICO caminho para abrir um arquivo daqui (arrastar-e-soltar, o botao Abrir… e a lista de
    /// recentes passam por aqui). Uma falha vira MENSAGEM na tela, nunca excecao solta: o <c>Drop</c> do
    /// code-behind e' <c>async void</c>, e uma excecao ali nao tem ninguem para observa-la — morre como
    /// excecao nao tratada do Dispatcher, derrubando o app inteiro por um arquivo corrompido.</summary>
    public async Task AbrirCaminhoAsync(string caminho)
    {
        MensagemErro = null;
        // Catch largo de PROPOSITO: e' a fronteira de UI. Abrir um arquivo passa por disco, conversor e
        // editor — nao da' para enumerar aqui o que cada um lanca, e engolir calado seria pior.
        try { await _abrir(caminho); }
        catch (Exception ex) { MensagemErro = $"Não foi possível abrir o arquivo: {ex.Message}"; }
    }

    [RelayCommand] private Task Novo() { MensagemErro = null; return _novo(); }

    [RelayCommand] private async Task Abrir()
    {
        MensagemErro = null;
        var caminho = await _escolher();
        if (!string.IsNullOrWhiteSpace(caminho)) await AbrirCaminhoAsync(caminho);
    }

    [RelayCommand] private async Task AbrirRecente(ItemRecente item)
    {
        if (!File.Exists(item.Caminho))
        {
            // Some da lista E avisa na propria tela (faixa vermelha) — nunca um MessageBox: o item ja' nao
            // existe, insistir com uma janela modal so' atrapalha quem esta tentando abrir outra coisa.
            _recentes.Remover(item.Caminho); Recarregar();
            MensagemErro = $"O arquivo não existe mais: {item.Caminho}";
            return;
        }
        await AbrirCaminhoAsync(item.Caminho);
    }

    /// <summary>Links discretos no rodape: desde que a barra de comandos saiu (2026-09-25), aqui e o menu da
    /// janela sao os dois lugares de Configuracoes e Sobre.</summary>
    public bool TemConfiguracoes => _configuracoes is not null;
    public bool TemSobre => _sobre is not null;
    [RelayCommand] private void Configuracoes() => _configuracoes?.Invoke();
    [RelayCommand] private void Sobre() => _sobre?.Invoke();

    [RelayCommand] private void Fixar(ItemRecente item) { _recentes.Fixar(item.Caminho, !item.Fixado); Recarregar(); }
    [RelayCommand] private void Remover(ItemRecente item) { _recentes.Remover(item.Caminho); Recarregar(); }
}
