namespace mOffice.Comum;

/// <summary>
/// Uma janela de documento, vista pela politica de janelas. O nucleo (<c>JanelaEditor</c>) NAO
/// implementa esta interface de proposito: o nucleo nao conhece o Comum. Quem faz a ponte e' o adaptador
/// privado <c>AplicativoOffice.JanelaAdaptada</c> — e e' o que deixa o <see cref="GerenteDeJanelas"/>
/// inteiro testavel sem WPF.
///
/// A criacao e' SINCRONA e o carregamento do documento e' um passo SEPARADO
/// (<see cref="CarregarAsync"/>) por um motivo de correcao, nao de estilo: ver
/// <see cref="GerenteDeJanelas"/>.
/// </summary>
public interface IJanelaDocumento
{
    /// <summary>Arquivo que a janela tem ABERTO AGORA, ou null. E' o estado do editor, entao vale so'
    /// depois de o documento carregar — e muda sozinho num "salvar como". Para saber que arquivo a janela
    /// foi criada para abrir (inclusive enquanto ainda carrega), quem manda e' o registro do gerente.</summary>
    string? Caminho { get; }

    void Ativar();

    /// <summary>Carrega o documento (abrir arquivo ou criar novo). Chamado pelo gerente DEPOIS de a janela
    /// ja' estar registrada e com <see cref="Fechado"/> assinado.</summary>
    Task CarregarAsync();

    /// <summary>Fecha a janela. So' o gerente chama: quando o carregamento falhou — ver o
    /// <c>catch</c> de <c>CriarAsync</c>, que explica por que o fechamento tem de vir DEPOIS de a janela
    /// sair do registro — e em <see cref="GerenteDeJanelas.FecharTodas"/> (atualizacao). E' o fechamento
    /// NORMAL da janela: um documento modificado pergunta se salva e pode recusar.</summary>
    void Fechar();

    event Action? Fechado;
}

/// <summary>Politica de janelas do app (spec §4): uma janela por documento no mesmo processo; abrir o mesmo
/// arquivo de novo so' ativa; a boas-vindas some quando ha documento e a ultima janela fechada encerra.</summary>
public sealed class GerenteDeJanelas
{
    /// <summary>Uma janela aberta MAIS o caminho que foi PEDIDO ao cria-la.
    ///
    /// O caminho pedido existe desde o instante da criacao; o <see cref="IJanelaDocumento.Caminho"/> so'
    /// aparece quando o documento termina de carregar (e' lido da sessao do editor). Guardar os dois e' o
    /// que faz a deduplicacao funcionar NOS DOIS SENTIDOS: durante o carregamento vale o pedido (dois
    /// duplos cliques seguidos no mesmo arquivo sao uma janela so'), e depois de um "salvar como" vale o
    /// da janela (o arquivo mudou e o pedido antigo ficou obsoleto). Documento novo tem pedido null e
    /// caminho null ate' ser salvo — por isso dois "Novo documento" continuam sendo duas janelas.</summary>
    private sealed record Aberta(IJanelaDocumento Janela, string? CaminhoPedido);

    private readonly Func<string?, IJanelaDocumento> _criar;
    private readonly Action _mostrar, _esconder, _encerrar;
    private readonly List<Aberta> _abertas = [];

    public GerenteDeJanelas(Func<string?, IJanelaDocumento> criarJanela, Action mostrarBoasVindas, Action esconderBoasVindas, Action encerrar)
    { _criar = criarJanela; _mostrar = mostrarBoasVindas; _esconder = esconderBoasVindas; _encerrar = encerrar; }

    public int Abertas => _abertas.Count;

    public async Task AbrirAsync(string caminho)
    {
        // Normaliza ANTES de comparar e de guardar (achado da revisao 1): o mesmo arquivo pode chegar por
        // caminhos escritos de formas diferentes — a linha de comando pode trazer relativo, o pipe da
        // segunda instancia traz absoluto, os recentes trazem o que foi gravado. Sem GetFullPath, "a.docx"
        // e "C:\x\a.docx" abririam DUAS janelas do mesmo documento (e as duas gravando por cima uma da
        // outra). A comparacao continua sem maiusculas: o Windows nao distingue.
        var normalizado = Path.GetFullPath(caminho);
        var existente = _abertas.FirstOrDefault(a => Corresponde(a, normalizado));
        if (existente is not null) { existente.Janela.Ativar(); return; }
        await CriarAsync(normalizado);
    }

    /// <summary>Esta janela e' a do arquivo pedido? Olha o caminho PEDIDO e o caminho ATUAL — ver
    /// <see cref="Aberta"/> para por que os dois sao necessarios.</summary>
    // O caminho REAL da janela (a sessao do editor) manda assim que existe: depois de um "salvar como",
    // a janela passa a ser do arquivo novo e o caminho pedido na abertura fica obsoleto — reabrir o
    // original tem de abrir OUTRA janela (achado da re-revisao 2: com `||` o original ativava a janela
    // que ja' era do arquivo novo, e um Ctrl+S ali gravaria por cima do arquivo errado). O caminho
    // pedido so' vale ENQUANTO a janela carrega (Caminho ainda nulo).
    private static bool Corresponde(Aberta a, string caminho) =>
        Igual(a.Janela.Caminho ?? a.CaminhoPedido, caminho);

    private static bool Igual(string? a, string b) =>
        a is not null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    public Task NovoAsync() => CriarAsync(null);

    /// <summary>A janela de documento que vem depois de <paramref name="atual"/> na ordem de abertura, voltando a primeira depois da ultima
    /// (Ctrl+F6 do Office). Com uma so' devolve ela mesma; <paramref name="atual"/> fora do registro, a primeira; sem janelas, <c>null</c>.</summary>
    public IJanelaDocumento? ProximaDepois(IJanelaDocumento atual)
    {
        if (_abertas.Count == 0) return null;
        var i = _abertas.FindIndex(a => ReferenceEquals(a.Janela, atual));
        return _abertas[(i + 1) % _abertas.Count].Janela;
    }

    /// <summary>
    /// Janela criada por uma fabrica PROPRIA — a reabertura de um documento recuperado (Task 3 do Plano 3), que
    /// carrega de uma pasta de recuperacao e nao de um caminho. Mesma politica de <see cref="NovoAsync"/>:
    /// esconde a boas-vindas, registra antes de carregar e, se o carregamento falhar, fecha a janela e lanca
    /// (quem chama mostra o erro). Sem caminho pedido: o caminho da janela aparece quando o documento carrega.
    /// </summary>
    public Task AbrirPorAsync(Func<IJanelaDocumento> criarJanela) => CriarAsync(null, criarJanela);

    /// <summary>Ha' uma janela com este arquivo (aberto, ou pedido e ainda carregando)? Mesma comparacao de
    /// <see cref="AbrirAsync"/>. A recuperacao usa para nao abrir o MESMO arquivo em duas janelas (o usuario
    /// reabriu o original depois da queda e agora pede para recuperar).</summary>
    public bool EstaAberto(string caminho)
    {
        string normalizado;
        try { normalizado = Path.GetFullPath(caminho); }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return false; }
        return _abertas.Any(a => Corresponde(a, normalizado));
    }

    /// <summary>
    /// Tira a janela do registro SEM fecha-la: ela vai fechar sozinha logo em seguida e esse fechamento nao pode
    /// contar como "ultima janela fechada" (encerraria o app). E' a reabertura recuperada que falhou
    /// (<c>JanelaEditor.RecuperacaoFalhou</c>): a janela se fecha e quem chama mostra a boas-vindas com o erro —
    /// o mesmo que o <c>catch</c> de <see cref="CriarAsync"/> faz com uma falha DURANTE o carregamento.
    /// Devolve se ela estava registrada.
    /// </summary>
    public bool Soltar(IJanelaDocumento j) => Remover(j);

    /// <summary>Subida do app (e volta ao estado sem documento): so' MOSTRA, nunca esconde. Esconder e'
    /// responsabilidade de <see cref="CriarAsync"/>, que ja' o faz em TODA criacao de janela — um
    /// <c>else _esconder()</c> aqui seria um segundo aviso de "esconda" para cada chamada feita com
    /// documento aberto, sem nenhum estado novo para comunicar (e e' o que o teste
    /// <c>MostrarBoasVindasSeVazio</c> fixa: dois eventos, nao tres).</summary>
    public void MostrarBoasVindasSeVazio() { if (_abertas.Count == 0) _mostrar(); }

    /// <summary>
    /// Cria a janela, REGISTRA (com o caminho pedido), assina o fechamento e so' entao carrega o documento.
    ///
    /// A ordem e' o contrato (achado CRITICO da revisao 1): abrir um documento leva segundos (mais ainda na
    /// primeira execucao, que gera o catalogo de fontes), e a janela fica na tela esse tempo todo. Enquanto
    /// o registro acontecia DEPOIS do <c>await</c>, um fechamento nesse intervalo disparava
    /// <see cref="IJanelaDocumento.Fechado"/> sem ninguem assinado; a janela morta era registrada logo em
    /// seguida e ficava la' para sempre — <see cref="Abertas"/> nunca voltava a zero e o processo
    /// sobrevivia sem nenhuma janela na tela (reproduzido com WM_CLOSE durante o carregamento).
    /// </summary>
    private async Task CriarAsync(string? caminho, Func<IJanelaDocumento>? criarJanela = null)
    {
        _esconder();
        var j = criarJanela is null ? _criar(caminho) : criarJanela();
        _abertas.Add(new Aberta(j, caminho));
        j.Fechado += () => AoFechar(j);
        try
        {
            await j.CarregarAsync();
        }
        catch
        {
            // Tira do registro ANTES de fechar: fechar com a janela ainda registrada faria o AoFechar
            // tratar isto como "ultima janela fechada" e ENCERRAR o app, quando o certo e' voltar a'
            // boas-vindas com a mensagem de erro. Se a janela ja' foi fechada durante o carregamento, o
            // AoFechar ja' a removeu (e ja' encerrou, se era a ultima) e nao ha' o que fechar — por isso o
            // fechamento fica dentro do `if` do Remover.
            if (Remover(j)) j.Fechar();
            throw;
        }
    }

    /// <summary>
    /// Pede o fechamento de TODAS as janelas abertas pelo fluxo normal de cada uma (Plano 5, Task 2: "instalar a
    /// atualizacao agora"). Nao espera nem forca: um documento modificado pergunta se salva e o usuario pode cancelar;
    /// as que fecham saem do registro por <see cref="AoFechar"/>, e a ultima encerra o app como sempre. Itera sobre uma
    /// COPIA: o fechamento de uma janela sem alteracao e' sincrono e mexe na lista no meio do laco.
    /// </summary>
    public void FecharTodas()
    {
        foreach (var a in _abertas.ToList()) a.Janela.Fechar();
    }

    private bool Remover(IJanelaDocumento j) => _abertas.RemoveAll(a => ReferenceEquals(a.Janela, j)) > 0;

    public void AoFechar(IJanelaDocumento j)
    {
        // So' conta como "uma janela a menos" se ela ESTAVA registrada. Sem esta guarda, o fechamento de
        // uma janela que o proprio gerente ja' tirou do registro (falha de carregamento, acima) zeraria a
        // contagem de novo e encerraria o app em vez de mostrar o erro.
        if (!Remover(j)) return;
        if (_abertas.Count == 0) _encerrar();
    }
}
