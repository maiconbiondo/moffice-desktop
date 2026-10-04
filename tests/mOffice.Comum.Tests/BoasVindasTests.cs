using mOffice.Comum;
using mOffice.Comum.Telas;
using mOffice.Conversor;
using mOffice.Nucleo;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Comum.Tests;

public class BoasVindasTests
{
    private static Recentes Rec() => new(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "r.json"));

    [Fact] // perfis: nome tecnico, nome de exibicao, formato do novo e filtro coerentes
    public void Perfis_Coerentes()
    {
        Assert.Equal("mtexto", PerfilApp.Texto.Nome); Assert.Equal("mTexto", PerfilApp.Texto.NomeExibicao);
        Assert.Equal(FormatoDocumento.Docx, PerfilApp.Texto.FormatoNovo); Assert.Contains("*.odt", PerfilApp.Texto.FiltroAbrir);
        Assert.Equal(FormatoDocumento.Xlsx, PerfilApp.Planilha.FormatoNovo); Assert.Equal(FormatoDocumento.Pptx, PerfilApp.Slides.FormatoNovo);
        Assert.Equal(["Logo.mTexto", "Logo.mPlanilha", "Logo.mSlides"], new[] { PerfilApp.Texto, PerfilApp.Planilha, PerfilApp.Slides }.Select(p => p.ChaveLogo));
    }

    [Fact] // Abrir: escolhe o arquivo e chama abrir; cancelar no dialogo nao chama abrir
    public async Task Abrir_ChamaEscolherDepoisAbrir()
    {
        var abertos = new List<string>(); string? escolha = @"C:\x\a.docx";
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, c => { abertos.Add(c); return Task.CompletedTask; }, () => Task.FromResult<string?>(escolha));
        await vm.AbrirCommand.ExecuteAsync(null);
        Assert.Equal([@"C:\x\a.docx"], abertos);
        escolha = null; await vm.AbrirCommand.ExecuteAsync(null);
        Assert.Single(abertos);
    }

    [Fact] // Plano 6, Task 10 (SO-10): recente ausente NAO sai sozinho (nem o fixado): mensagem, "Remover da lista" visivel, abrir NAO e' chamado; o comando tira
    public async Task RecenteAusente_NaoRemoveSozinho_OfereceRemover()
    {
        var rec = Rec(); var caminho = @"C:\pasta\Relatorio.docx";
        rec.Registrar(caminho); rec.Fixar(caminho, true);
        var abertos = 0;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => { abertos++; return Task.CompletedTask; }, () => Task.FromResult<string?>(null),
            verificarRecente: _ => Task.FromResult(EstadoDoRecente.NaoExiste));
        Assert.Single(vm.Recentes); Assert.False(vm.TemRecenteAusente);
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.Equal(0, abertos);
        Assert.Single(vm.Recentes); Assert.Single(rec.Listar()); Assert.True(rec.Listar()[0].Fixado);   // nada saiu, nem do arquivo
        Assert.Equal("Não foi possível encontrar \"Relatorio.docx\". Ele foi movido, renomeado ou apagado?", vm.MensagemErro);
        Assert.True(vm.TemRecenteAusente);
        vm.RemoverRecenteAusenteCommand.Execute(null);   // so' o usuario tira
        Assert.Empty(vm.Recentes); Assert.Empty(rec.Listar());
        Assert.False(vm.TemRecenteAusente); Assert.Null(vm.MensagemErro);
    }

    [Fact] // menor: qualquer outra mensagem de erro (outra acao) apaga a oferta "Remover da lista" do recente ausente — ela nao fica pendurada
    public async Task RecenteAusente_LimpaQuandoOutraMensagemSubstitui()
    {
        var rec = Rec(); rec.Registrar(@"C:\a.docx");
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => throw new IOException("falhou"), () => Task.FromResult<string?>(null),
            verificarRecente: _ => Task.FromResult(EstadoDoRecente.NaoExiste));
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.True(vm.TemRecenteAusente);
        await vm.AbrirCaminhoAsync(@"C:\outro.docx");   // outra acao falha: mensagem nova
        Assert.Contains("Não foi possível abrir", vm.MensagemErro);
        Assert.False(vm.TemRecenteAusente);
    }

    [Fact] // menor: durante a sonda o comando fica indisponivel (clique repetido ignorado: nada de varias sondas/aberturas em paralelo) e volta depois
    public async Task AbrirRecente_DuranteASonda_ComandoIndisponivel()
    {
        var rec = Rec(); rec.Registrar(@"C:\a.docx");
        var liberar = new TaskCompletionSource<EstadoDoRecente>(); var sondas = 0; var abertos = 0;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => { abertos++; return Task.CompletedTask; }, () => Task.FromResult<string?>(null),
            verificarRecente: _ => { sondas++; return liberar.Task; });
        var item = vm.Recentes[0];
        Assert.True(vm.AbrirRecenteCommand.CanExecute(item));
        var primeiro = vm.AbrirRecenteCommand.ExecuteAsync(item);
        Assert.False(vm.AbrirRecenteCommand.CanExecute(item));
        vm.AbrirRecenteCommand.Execute(item);   // clique repetido: ignorado
        Assert.Equal(1, sondas);
        liberar.SetResult(EstadoDoRecente.Existe); await primeiro;
        Assert.Equal(1, abertos);
        Assert.True(vm.AbrirRecenteCommand.CanExecute(item));
    }

    [Fact] // rede sem resposta: mensagem propria (nao "foi apagado"), o recente fica, nada e' aberto
    public async Task RecenteSemResposta_MensagemDeRede()
    {
        var rec = Rec(); rec.Registrar(@"\\srv\docs\Contrato.docx");
        var abertos = 0;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => { abertos++; return Task.CompletedTask; }, () => Task.FromResult<string?>(null),
            verificarRecente: _ => Task.FromResult(EstadoDoRecente.SemResposta));
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.Equal(0, abertos); Assert.Single(rec.Listar());
        Assert.Equal("Não foi possível acessar \"\\\\srv\\docs\". A pasta de rede está acessível?", vm.MensagemErro);
        Assert.True(vm.TemRecenteAusente);
    }

    [Fact] // recente que existe abre normalmente e limpa a oferta de remover de uma tentativa anterior
    public async Task RecenteExistente_Abre_LimpaOferta()
    {
        var rec = Rec(); rec.Registrar(@"C:\a.docx");
        var estado = EstadoDoRecente.NaoExiste; var abertos = new List<string>();
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, c => { abertos.Add(c); return Task.CompletedTask; }, () => Task.FromResult<string?>(null),
            verificarRecente: _ => Task.FromResult(estado));
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.True(vm.TemRecenteAusente);
        estado = EstadoDoRecente.Existe;
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.Equal([@"C:\a.docx"], abertos); Assert.False(vm.TemRecenteAusente); Assert.Null(vm.MensagemErro);
    }

    [Fact] // a sonda padrao e' a de verdade: arquivo que nao existe vira a mensagem, sem remover (e o clique volta sem esperar mais que o limite)
    public async Task RecenteAusente_SondaPadrao_NaoRemove()
    {
        var rec = Rec(); var caminho = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N") + ".docx");
        rec.Registrar(caminho);
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null));
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.Single(rec.Listar()); Assert.Contains("Não foi possível encontrar", vm.MensagemErro);
    }

    [Fact] // a data dos recentes (e dos recuperados/nao salvos) leva o ANO: dd/MM/yyyy HH:mm (o ano some em abril de um ano e aparece em maio do outro)
    public void DataDosRecentes_ComAno()
    {
        var xaml = File.ReadAllText(Path.Combine(AplicativoOfficeTests.RaizRepo(), "src", "mOffice.Comum", "Telas", "TelaBoasVindas.xaml"));
        Assert.Contains("{Binding Quando, StringFormat=dd/MM/yyyy HH:mm}", xaml);
        Assert.DoesNotContain("StringFormat=dd/MM HH:mm", xaml);
    }

    [Fact] // "Nao salvos recentemente": a secao lista o que o delegado devolve; Abrir recupera (mesmo delegado da recuperacao); Excluir pergunta e so' entao apaga
    public async Task NaoSalvos_ListaAbreEExclui()
    {
        var d1 = new DocumentoRecuperavel(@"C:\d\1", "Carta", @"C:\x\carta.docx", false, DateTime.UtcNow.AddHours(-2));
        var d2 = new DocumentoRecuperavel(@"C:\d\2", "Documento1", "", true, DateTime.UtcNow.AddDays(-1));
        var lista = new List<DocumentoRecuperavel> { d1, d2 };
        DocumentoRecuperavel? aberto = null; var excluidos = new List<string>(); var confirma = false;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null),
            recuperar: d => { aberto = d; return Task.CompletedTask; }, descartar: d => { excluidos.Add(d.Pasta); lista.Remove(d); }, confirmarDescarte: _ => confirma,
            naoSalvos: () => lista.ToList());
        Assert.True(vm.TemNaoSalvos); Assert.Equal(2, vm.NaoSalvos.Count);
        Assert.False(vm.TemRecuperados);   // as duas secoes sao independentes
        await vm.AbrirNaoSalvoCommand.ExecuteAsync(d1);
        Assert.Equal(d1, aberto);
        vm.ExcluirNaoSalvoCommand.Execute(d2);          // pergunta -> Nao
        Assert.Empty(excluidos); Assert.Equal(2, vm.NaoSalvos.Count);
        confirma = true;
        vm.ExcluirNaoSalvoCommand.Execute(d2);          // pergunta -> Sim
        Assert.Equal([@"C:\d\2"], excluidos); Assert.Single(vm.NaoSalvos); Assert.Null(vm.MensagemErro);
    }

    [Fact] // sem delegado de nao salvos a secao nunca aparece
    public void NaoSalvos_SemDelegado_Escondida() =>
        Assert.False(new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null)).TemNaoSalvos);

    [Fact] // fixar/remover refletem no arquivo e na colecao observavel
    public void FixarERemover_AtualizamListaEDisco()
    {
        var rec = Rec(); rec.Registrar(@"C:\a.docx"); rec.Registrar(@"C:\b.docx");
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null));
        vm.FixarCommand.Execute(vm.Recentes.Single(i => i.Caminho == @"C:\a.docx"));
        Assert.True(vm.Recentes[0].Fixado); Assert.Equal(@"C:\a.docx", vm.Recentes[0].Caminho);
        vm.RemoverCommand.Execute(vm.Recentes[1]);
        Assert.Single(vm.Recentes); Assert.Single(rec.Listar());
    }

    [Fact] // abrir que falha vira MENSAGEM na tela, nunca excecao solta (o Drop do code-behind e' async void)
    public async Task AbrirCaminho_QuandoAbrirLanca_ViraMensagemENaoPropaga()
    {
        var falhar = true;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask,
            _ => falhar ? throw new InvalidOperationException("formato corrompido") : Task.CompletedTask,
            () => Task.FromResult<string?>(null));

        await vm.AbrirCaminhoAsync(@"C:\x\a.docx"); // nao lanca: se lancasse, este await propagaria
        Assert.Contains("Não foi possível abrir o arquivo", vm.MensagemErro);
        Assert.Contains("formato corrompido", vm.MensagemErro);

        falhar = false;
        await vm.AbrirCaminhoAsync(@"C:\x\a.docx"); // a mensagem velha nao pode ficar na tela
        Assert.Null(vm.MensagemErro);
    }

    [Fact] // solto na tela: so' extensoes do perfil contam; a primeira valida abre (funcao pura no perfil, sem WPF)
    public void ArquivoSolto_FiltraPorPerfil()
    {
        Assert.Equal(@"C:\b.odt", PerfilApp.Texto.PrimeiroAceito([@"C:\a.xlsx", @"C:\b.odt", @"C:\c.docx"]));
        Assert.Null(PerfilApp.Texto.PrimeiroAceito([@"C:\a.xlsx"]));
    }

    // ---------------------------------------------------------------------------------------------------
    // Documentos recuperados (Task 3 do Plano 3). A pergunta de descarte e' SEMPRE um falso aqui: a padrao e'
    // um MessageBox do sistema, que travaria o teste.
    // ---------------------------------------------------------------------------------------------------

    [Fact] // recuperados aparecem; Recuperar chama o delegado; Descartar pergunta, chama o delegado e tira da lista
    public async Task Recuperados_RecuperarEDescartar()
    {
        var d1 = new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false, DateTime.UtcNow);
        var d2 = new DocumentoRecuperavel(@"C:\r\2", "Documento1", "", true, DateTime.UtcNow);
        var lista = new List<DocumentoRecuperavel> { d1, d2 };
        DocumentoRecuperavel? recuperado = null, perguntado = null;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask,
            () => Task.FromResult<string?>(null), null, null,
            () => lista, d => { recuperado = d; return Task.CompletedTask; }, d => lista.Remove(d),
            d => { perguntado = d; return true; });
        Assert.True(vm.TemRecuperados); Assert.Equal(2, vm.Recuperados.Count);
        await vm.RecuperarCommand.ExecuteAsync(d1); Assert.Same(d1, recuperado);
        vm.DescartarCommand.Execute(d2); Assert.Same(d2, perguntado); Assert.Equal([d1], vm.Recuperados);
        Assert.Null(vm.MensagemErro);
    }

    [Fact] // "Nao" na pergunta: nada e' apagado; a lista (e o aviso de TemRecuperados) acompanham a fonte
    public void Recuperados_DescartarSemConfirmar_NaoApaga()
    {
        var d1 = new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false, DateTime.UtcNow);
        var lista = new List<DocumentoRecuperavel> { d1 };
        var descartes = 0; var confirmar = false;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask,
            () => Task.FromResult<string?>(null), recuperaveis: () => lista, recuperar: _ => Task.CompletedTask,
            descartar: d => { descartes++; lista.Remove(d); }, confirmarDescarte: _ => confirmar);
        var avisos = new List<string?>(); vm.PropertyChanged += (_, e) => avisos.Add(e.PropertyName);
        vm.DescartarCommand.Execute(d1);
        Assert.Equal(0, descartes); Assert.Equal([d1], vm.Recuperados);
        confirmar = true; vm.DescartarCommand.Execute(d1);
        Assert.Equal(1, descartes); Assert.Empty(vm.Recuperados); Assert.False(vm.TemRecuperados);
        Assert.Contains(nameof(TelaBoasVindasViewModel.TemRecuperados), avisos);
    }

    [Fact] // descartar que nao apagou (pasta presa) avisa em vez de sumir calado; recuperar que falhou vira mensagem
    public async Task Recuperados_Falhas_ViramMensagem()
    {
        var d1 = new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false, DateTime.UtcNow);
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask,
            () => Task.FromResult<string?>(null), recuperaveis: () => [d1],
            recuperar: _ => throw new InvalidOperationException("Este documento já está aberto em outra janela."),
            descartar: _ => { }, confirmarDescarte: _ => true);
        vm.DescartarCommand.Execute(d1);
        Assert.Contains("Não foi possível descartar", vm.MensagemErro); Assert.Single(vm.Recuperados);
        await vm.RecuperarCommand.ExecuteAsync(d1);
        Assert.Contains("já está aberto em outra janela", vm.MensagemErro); Assert.Single(vm.Recuperados);
    }

    [Fact] // sem fonte de recuperados (null): secao escondida; o texto da pergunta de descarte e' o da spec
    public void Recuperados_SemFonte_EscondidoETextoDaPergunta()
    {
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null));
        Assert.False(vm.TemRecuperados); Assert.Empty(vm.Recuperados);
        Assert.Equal("Descartar as alterações não salvas de peticao.docx? Esta ação não pode ser desfeita.",
            TelaBoasVindasViewModel.PerguntaDescartar(new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", "", false, DateTime.UtcNow)));
    }

    private static TelaBoasVindasViewModel ComModelos(Func<IReadOnlyList<ModeloDisponivel>> modelos, Func<string, Task>? novo = null) =>
        new(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null),
            modelos: modelos, novoDoModelo: novo);

    [Fact] // Plano 6, Task 9: a galeria lista os modelos (usuario primeiro) e o clique cria o documento a partir do arquivo do modelo
    public async Task Modelos_ListaEClicaCriaDoModelo()
    {
        var criados = new List<string>();
        var vm = ComModelos(() => [new("Peticao", @"C:\m\Peticao.dotx", true), new("Ordem de compra", @"C:\mot\x.dotx", false)], c => { criados.Add(c); return Task.CompletedTask; });
        await vm.RecarregarModelosAsync();
        Assert.True(vm.TemModelos); Assert.Equal(["Peticao", "Ordem de compra"], vm.Modelos.Select(m => m.Nome));
        await vm.NovoDoModeloCommand.ExecuteAsync(vm.Modelos[1]);
        Assert.Equal([@"C:\mot\x.dotx"], criados);
    }

    [Fact] // listagem que lanca ou passa do tempo limite: galeria vazia, nunca excecao nem trava
    public async Task Modelos_ListagemQueLancaOuDemora_FicaVazia()
    {
        var vm = ComModelos(() => throw new IOException("rede fora do ar"));
        await vm.RecarregarModelosAsync();
        Assert.False(vm.TemModelos);
        using var solta = new ManualResetEventSlim();
        var lento = ComModelos(() => { solta.Wait(TimeSpan.FromSeconds(30)); return [new("x", "x.dotx", true)]; });
        var t = System.Diagnostics.Stopwatch.StartNew();
        await lento.RecarregarModelosAsync();
        solta.Set();
        Assert.True(t.Elapsed < TimeSpan.FromSeconds(10), "esperou a listagem lenta");
        Assert.False(lento.TemModelos);
    }

    [Fact] // criar a partir do modelo que falha vira mensagem na tela; MostrarModelosAsync pede a rolagem
    public async Task Modelos_FalhaViraMensagem_ERolagem()
    {
        var vm = ComModelos(() => [new("A", "a.dotx", true)], _ => throw new InvalidOperationException("modelo corrompido"));
        var rolou = 0; vm.RolarParaModelos += () => rolou++;
        await vm.MostrarModelosAsync();
        Assert.Equal(1, rolou); Assert.Single(vm.Modelos);
        await vm.NovoDoModeloCommand.ExecuteAsync(vm.Modelos[0]);
        Assert.Contains("modelo corrompido", vm.MensagemErro);
    }

    [Theory] // create:new do web-apps: "template:*" leva a' galeria; o resto cria em branco
    [InlineData("template:word", true)] [InlineData("template:cell", true)] [InlineData("template:slide", true)]
    [InlineData("", false)] [InlineData("word", false)] [InlineData(null, false)]
    public void CreateNew_Template_PedeModelo(string? json, bool esperado) => Assert.Equal(esperado, AplicativoOffice.PedeModelo(json));

    [Fact] // a galeria mostra so' os primeiros (os recuperados e recentes ficam na primeira tela); "Mais modelos…" mostra todos; o pedido do editor ja' mostra todos
    public async Task Modelos_LimiteInicial_MaisModelos()
    {
        var todos = Enumerable.Range(1, 10).Select(i => new ModeloDisponivel("M" + i, $@"C:\m\{i}.dotx", i <= 2)).ToList();
        var vm = ComModelos(() => todos);
        await vm.RecarregarModelosAsync();
        Assert.Equal(10, vm.Modelos.Count); Assert.Equal(TelaBoasVindasViewModel.LimiteInicialDeModelos, vm.ModelosVisiveis.Count);
        Assert.True(vm.TemMaisModelos);
        vm.MaisModelosCommand.Execute(null);
        Assert.Equal(10, vm.ModelosVisiveis.Count); Assert.False(vm.TemMaisModelos);
        var outro = ComModelos(() => todos); await outro.MostrarModelosAsync();
        Assert.Equal(10, outro.ModelosVisiveis.Count);
        var pouco = ComModelos(() => todos.Take(3).ToList()); await pouco.RecarregarModelosAsync();
        Assert.False(pouco.TemMaisModelos); Assert.Equal(3, pouco.ModelosVisiveis.Count);
    }

    [Fact] // listagem lenta (pasta de rede): chamadas seguintes esperam a MESMA listagem, nao empilham outra
    public async Task Modelos_UmaListagemEmVoo()
    {
        using var solta = new ManualResetEventSlim(); var chamadas = 0;
        var vm = ComModelos(() => { Interlocked.Increment(ref chamadas); solta.Wait(TimeSpan.FromSeconds(20)); return [new("x", "x.dotx", true)]; });
        // o construtor ja' iniciou a primeira; mais duas chamadas passam do limite de 2 s sem iniciar outra
        await vm.RecarregarModelosAsync(); await vm.RecarregarModelosAsync();
        Assert.Equal(1, vm.ListagensIniciadas); Assert.Equal(1, Volatile.Read(ref chamadas));
        solta.Set();
        await Task.Delay(300);
        await vm.RecarregarModelosAsync();
        Assert.Equal(2, vm.ListagensIniciadas); Assert.Single(vm.Modelos);
    }
}
