using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using mOffice.Nucleo;
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

    private readonly Func<IReadOnlyList<ModeloDisponivel>>? _modelos;
    private readonly Func<string, Task>? _novoDoModelo;

    private readonly Func<IReadOnlyList<DocumentoRecuperavel>>? _naoSalvos;
    private readonly Func<string, Task<EstadoDoRecente>> _verificarRecente;

    /// <param name="configuracoes">Abre as Configuracoes do mOffice; null = link escondido.</param>
    /// <param name="sobre">Abre o Sobre (atribuicao de licencas); null = link escondido.</param>
    /// <param name="recuperaveis">Documentos que uma sessao anterior deixou sem salvar (Task 3 do Plano 3); null =
    /// secao "Documentos recuperados" sempre escondida.</param>
    /// <param name="recuperar">Reabre um documento recuperavel numa janela de documento.</param>
    /// <param name="descartar">Apaga um documento recuperavel (so' depois de <paramref name="confirmarDescarte"/>).</param>
    /// <param name="modelos">Lista os modelos da galeria "Novo" (Plano 6, Task 9); roda FORA da thread da tela, com tempo limite de
    /// <see cref="TempoLimiteModelos"/> (uma pasta de rede fora do ar nao pode travar a boas-vindas). Null = sem galeria.</param>
    /// <param name="novoDoModelo">Cria um documento novo a partir do arquivo de modelo (o modelo nunca e' o destino do Salvar).</param>
    /// <param name="naoSalvos">"Não salvos recentemente" (Plano 6, Task 10): o que o "Não salvar" do fechamento guardou nos descartados (4 dias); null = seção escondida.
    /// "Abrir" usa o mesmo <paramref name="recuperar"/> e "Excluir" o mesmo <paramref name="descartar"/> (com a mesma pergunta).</param>
    /// <param name="verificarRecente">Confere o recente ao clicar (fora da thread da tela, com limite); padrão = <see cref="VerificadorDeRecente"/> com <see cref="TempoLimiteRecente"/>.
    /// Os testes passam um falso.</param>
    /// <param name="confirmarDescarte">Pergunta antes de descartar; true = pode apagar. Null = a pergunta do
    /// sistema (<see cref="ConfirmarDescarteNoSistema"/> sem dono, MessageBox Sim/Nao). Os testes SEMPRE passam um
    /// falso: a padrao abre uma janela modal.</param>
    public TelaBoasVindasViewModel(PerfilApp perfil, Recentes recentes, Func<Task> novo, Func<string, Task> abrir,
        Func<Task<string?>> escolherArquivo, Action? configuracoes = null, Action? sobre = null,
        Func<IReadOnlyList<DocumentoRecuperavel>>? recuperaveis = null, Func<DocumentoRecuperavel, Task>? recuperar = null,
        Action<DocumentoRecuperavel>? descartar = null, Func<DocumentoRecuperavel, bool>? confirmarDescarte = null,
        Func<IReadOnlyList<ModeloDisponivel>>? modelos = null, Func<string, Task>? novoDoModelo = null,
        Func<IReadOnlyList<DocumentoRecuperavel>>? naoSalvos = null, Func<string, Task<EstadoDoRecente>>? verificarRecente = null)
    {
        _modelos = modelos; _novoDoModelo = novoDoModelo;
        _naoSalvos = naoSalvos;
        _verificarRecente = verificarRecente ?? (c => VerificadorDeRecente.VerificarAsync(c, TempoLimiteRecente));
        _perfil = perfil; _recentes = recentes; _novo = novo; _abrir = abrir; _escolher = escolherArquivo;
        _configuracoes = configuracoes; _sobre = sobre;
        _recuperaveis = recuperaveis; _recuperar = recuperar; _descartar = descartar;
        _confirmarDescarte = confirmarDescarte ?? (d => ConfirmarDescarteNoSistema(null, d));
        Recarregar();
    }

    /// <summary>O perfil do app, para o code-behind filtrar o que foi SOLTO na tela (PrimeiroAceito).</summary>
    public PerfilApp Perfil => _perfil;
    public string Titulo => _perfil.NomeExibicao;
    public string RotuloNovo => _perfil.RotuloNovo;
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
        RecarregarNaoSalvos();
        _ = RecarregarModelosAsync();
    }

    // ---------------------------------------------------------------------------------------------------
    // Galeria de modelos (Plano 6, Task 9)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Tempo maximo esperando a listagem dos modelos (pasta de rede inacessivel).</summary>
    public static readonly TimeSpan TempoLimiteModelos = TimeSpan.FromSeconds(2);

    /// <summary>Quantos modelos a boas-vindas mostra antes do "Mais modelos…": a galeria nao empurra "Documentos recuperados" nem os recentes para fora da tela.</summary>
    public const int LimiteInicialDeModelos = 6;

    /// <summary>Todos os modelos (usuario primeiro, depois os do motor).</summary>
    public ObservableCollection<ModeloDisponivel> Modelos { get; } = [];
    /// <summary>O que a tela lista: os <see cref="LimiteInicialDeModelos"/> primeiros, ou todos depois de "Mais modelos…".</summary>
    public ObservableCollection<ModeloDisponivel> ModelosVisiveis { get; } = [];
    public bool TemModelos => Modelos.Count > 0;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(TemMaisModelos))] private bool _todosOsModelos;
    public bool TemMaisModelos => !TodosOsModelos && Modelos.Count > LimiteInicialDeModelos;
    [RelayCommand] private void MaisModelos() { TodosOsModelos = true; AtualizarModelosVisiveis(); }

    private void AtualizarModelosVisiveis()
    {
        ModelosVisiveis.Clear();
        foreach (var m in TodosOsModelos ? Modelos : Modelos.Take(LimiteInicialDeModelos)) ModelosVisiveis.Add(m);
        OnPropertyChanged(nameof(TemMaisModelos));
    }

    /// <summary>Pedido para a tela rolar ate' a galeria (o "Criar a partir de modelo" do editor).</summary>
    public event Action? RolarParaModelos;

    private int _geracaoModelos;
    private Task<IReadOnlyList<ModeloDisponivel>>? _listagemEmVoo;
    private readonly object _trava = new();
    /// <summary>Quantas listagens de verdade foram iniciadas (a pasta de rede lenta nao acumula threads: ver <see cref="RecarregarModelosAsync"/>).</summary>
    public int ListagensIniciadas { get; private set; }

    /// <summary>Le' os modelos fora da thread da tela e, se passar do limite ou falhar, deixa a lista VAZIA — nunca lanca. A ultima
    /// chamada vence. Uma UNICA listagem fica em voo: se uma pasta de rede lenta ainda nao respondeu, as chamadas seguintes esperam a
    /// MESMA listagem (com o mesmo limite) em vez de empilhar outra thread.</summary>
    public async Task RecarregarModelosAsync()
    {
        if (_modelos is null) return;
        var minha = Interlocked.Increment(ref _geracaoModelos);
        IReadOnlyList<ModeloDisponivel> lista = [];
        try
        {
            Task<IReadOnlyList<ModeloDisponivel>> tarefa;
            lock (_trava)
            {
                if (_listagemEmVoo is not { IsCompleted: false }) { ListagensIniciadas++; _listagemEmVoo = Task.Run(_modelos); }
                tarefa = _listagemEmVoo;
            }
            if (await Task.WhenAny(tarefa, Task.Delay(TempoLimiteModelos)) == tarefa) lista = await tarefa;
            else _ = tarefa.ContinueWith(t => _ = t.Exception, TaskContinuationOptions.OnlyOnFaulted); // observa a falha tardia
        }
        catch (Exception) { lista = []; }
        if (minha != Volatile.Read(ref _geracaoModelos)) return;
        Modelos.Clear(); foreach (var m in lista) Modelos.Add(m);
        AtualizarModelosVisiveis();
        OnPropertyChanged(nameof(TemModelos));
    }

    /// <summary>Mostra a galeria COMPLETA (recarregada) e pede para a tela rolar ate' ela.</summary>
    public async Task MostrarModelosAsync() { TodosOsModelos = true; await RecarregarModelosAsync(); RolarParaModelos?.Invoke(); }

    [RelayCommand] private async Task NovoDoModelo(ModeloDisponivel modelo)
    {
        if (_novoDoModelo is null) return;
        MensagemErro = null;
        try { await _novoDoModelo(modelo.Caminho); }
        catch (Exception ex) { MensagemErro = $"Não foi possível criar o documento a partir do modelo: {ex.Message}"; }
    }

    // ---------------------------------------------------------------------------------------------------
    // Documentos recuperados (Task 3 do Plano 3)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Documentos que o app fechou sem salvar (queda), do mais recente para o mais antigo.</summary>
    public ObservableCollection<DocumentoRecuperavel> Recuperados { get; } = [];

    /// <summary>A secao "Documentos recuperados" aparece?</summary>
    public bool TemRecuperados => Recuperados.Count > 0;

    // ---------------------------------------------------------------------------------------------------
    // Não salvos recentemente (Plano 6, Task 10)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>O que o usuário mandou "Não salvar" nos últimos 4 dias (<c>Descartados.Retencao</c>), do mais recente para o mais antigo.</summary>
    public ObservableCollection<DocumentoRecuperavel> NaoSalvos { get; } = [];
    public bool TemNaoSalvos => NaoSalvos.Count > 0;

    private void RecarregarNaoSalvos()
    {
        NaoSalvos.Clear();
        try { if (_naoSalvos is not null) foreach (var d in _naoSalvos()) NaoSalvos.Add(d); }
        catch (Exception ex) { MensagemErro = $"Não foi possível listar os documentos não salvos: {ex.Message}"; }
        OnPropertyChanged(nameof(TemNaoSalvos));
    }

    /// <summary>"Abrir" de um não salvo: reabre a pasta como uma recuperação (o documento volta modificado).</summary>
    [RelayCommand] private async Task AbrirNaoSalvo(DocumentoRecuperavel d)
    {
        MensagemErro = null;
        try { if (_recuperar is not null) await _recuperar(d); }
        catch (Exception ex) { MensagemErro = $"Não foi possível abrir o documento: {ex.Message}"; }
        finally { RecarregarNaoSalvos(); }
    }

    /// <summary>"Excluir": SEMPRE ato do usuário (com a mesma pergunta do Descartar). Nada mais apaga um não salvo antes dos 4 dias.</summary>
    [RelayCommand] private void ExcluirNaoSalvo(DocumentoRecuperavel d)
    {
        if (_descartar is null) return;
        MensagemErro = null;
        if (!_confirmarDescarte(d)) return;
        try { _descartar(d); }
        catch (Exception ex) { MensagemErro = $"Não foi possível excluir \"{d.Titulo}\": {ex.Message}"; }
        RecarregarNaoSalvos();
        if (MensagemErro is null && NaoSalvos.Any(r => string.Equals(r.Pasta, d.Pasta, StringComparison.OrdinalIgnoreCase)))
            MensagemErro = $"Não foi possível excluir \"{d.Titulo}\" agora. Tente de novo em instantes.";
    }

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

    /// <summary>Tempo máximo esperando a sonda de um recente (um UNC fora do ar não pode travar a tela).</summary>
    public static readonly TimeSpan TempoLimiteRecente = TimeSpan.FromSeconds(2);

    /// <summary>O recente que não abriu (sumiu ou a rede não respondeu): a tela oferece "Remover da lista" para ele. Nunca é tirado sozinho.</summary>
    [ObservableProperty, NotifyPropertyChangedFor(nameof(TemRecenteAusente))] private ItemRecente? _recenteAusente;
    public bool TemRecenteAusente => RecenteAusente is not null;

    private bool _definindoAusente;
    /// <summary>Qualquer OUTRA mensagem de erro (outra acao falhou, ou a faixa foi limpa) apaga a oferta "Remover da lista": ela so' vale junto da mensagem do recente ausente.
    /// (O clique repetido durante a sonda ja' e' ignorado: o AsyncRelayCommand nao roda duas vezes ao mesmo tempo.)</summary>
    partial void OnMensagemErroChanged(string? value) { if (!_definindoAusente) RecenteAusente = null; }

    [RelayCommand] private async Task AbrirRecente(ItemRecente item)
    {
        // Guarda de ocupado: durante a sonda (ate' 2 s) um segundo clique/Enter nao dispara outra sonda nem outra abertura. (O botao tambem fica
        // desabilitado pelo AsyncRelayCommand, mas Execute() programatico nao olha CanExecute.)
        if (_verificando) return;
        _verificando = true;
        try { await AbrirRecenteVerificandoAsync(item); }
        finally { _verificando = false; }
    }

    private bool _verificando;

    private async Task AbrirRecenteVerificandoAsync(ItemRecente item)
    {
        MensagemErro = null; RecenteAusente = null;
        // A sonda roda FORA da thread da tela e com limite; o recente NUNCA sai da lista por aqui (nem o fixado, nem o de uma pasta de
        // rede que só está fora do ar): quem tira é o usuário, pelo "Remover da lista". Faixa vermelha na própria tela, nunca MessageBox.
        var estado = await _verificarRecente(item.Caminho);
        if (estado == EstadoDoRecente.Existe) { await AbrirCaminhoAsync(item.Caminho); return; }
        _definindoAusente = true;
        try
        {
            MensagemErro = estado == EstadoDoRecente.NaoExiste
                ? $"Não foi possível encontrar \"{Path.GetFileName(item.Caminho)}\". Ele foi movido, renomeado ou apagado?"
                : $"Não foi possível acessar \"{PastaDe(item.Caminho)}\". A pasta de rede está acessível?";
            RecenteAusente = item;
        }
        finally { _definindoAusente = false; }
    }

    private static string PastaDe(string caminho) { try { return Path.GetDirectoryName(caminho) ?? caminho; } catch (ArgumentException) { return caminho; } }

    /// <summary>"Remover da lista" do recente ausente: só o usuário tira.</summary>
    [RelayCommand] private void RemoverRecenteAusente()
    {
        if (RecenteAusente is not { } item) return;
        _recentes.Remover(item.Caminho);
        RecenteAusente = null; MensagemErro = null;
        Recarregar();
        RecentesAlterados?.Invoke();
    }

    /// <summary>Links discretos no rodape: desde que a barra de comandos saiu (2026-09-25), aqui e o menu da
    /// janela sao os dois lugares de Configuracoes e Sobre.</summary>
    public bool TemConfiguracoes => _configuracoes is not null;
    public bool TemSobre => _sobre is not null;
    [RelayCommand] private void Configuracoes() => _configuracoes?.Invoke();
    [RelayCommand] private void Sobre() => _sobre?.Invoke();

    /// <summary>A lista de recentes mudou por ação do usuário aqui (fixar/remover): a casca refaz a Jump List.</summary>
    public event Action? RecentesAlterados;

    [RelayCommand] private void Fixar(ItemRecente item) { _recentes.Fixar(item.Caminho, !item.Fixado); Recarregar(); RecentesAlterados?.Invoke(); }
    [RelayCommand] private void Remover(ItemRecente item) { _recentes.Remover(item.Caminho); Recarregar(); RecentesAlterados?.Invoke(); }
}
