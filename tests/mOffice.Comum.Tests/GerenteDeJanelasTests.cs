using mOffice.Comum;

namespace mOffice.Comum.Tests;

public class GerenteDeJanelasTests
{
    private sealed class JanelaFake(string? caminho) : IJanelaDocumento
    {
        private readonly string? _caminho = caminho;
        public int Ativacoes;
        public int Fechamentos;

        /// <summary>Quando nao e' nulo, <see cref="CarregarAsync"/> fica PENDENTE ate' o teste completar —
        /// e' assim que se simula o intervalo (segundos, na vida real) em que a janela ja' esta' na tela e
        /// o documento ainda esta' abrindo.</summary>
        public TaskCompletionSource? Carga;

        /// <summary>null ENQUANTO carrega, como na janela de verdade: o adaptador le' o caminho da sessao
        /// do editor, e sessao so' existe depois de o documento abrir. E' justamente por isso que o
        /// gerente nao pode deduplicar so' por esta propriedade.</summary>
        public string? Caminho => Carga is null || Carga.Task.IsCompletedSuccessfully ? (CaminhoAtual ?? _caminho) : null;

        /// <summary>Simula um "salvar como": a sessao da janela passa a apontar para OUTRO arquivo.</summary>
        public string? CaminhoAtual;

        public event Action? Fechado;
        public void Ativar() => Ativacoes++;
        public Task CarregarAsync() => Carga?.Task ?? Task.CompletedTask;
        /// <summary>O gerente mandando fechar (falha de carregamento). Numa janela de verdade, fechar dispara
        /// <see cref="Fechado"/> — o fake faz o mesmo, senao o teste nao veria o efeito colateral real.</summary>
        public void Fechar() { Fechamentos++; if (!Recusar) Disparar(); }
        /// <summary>O documento modificado cujo usuario respondeu "Cancelar": o pedido de fechar nao fecha.</summary>
        public bool Recusar;
        /// <summary>O USUARIO fechando a janela.</summary>
        public void Disparar() => Fechado?.Invoke();
    }

    private static (GerenteDeJanelas g, List<JanelaFake> janelas, List<string> eventos) Montar(bool suspenderCarga = false)
    {
        var janelas = new List<JanelaFake>(); var eventos = new List<string>();
        var g = new GerenteDeJanelas(
            c => { var j = new JanelaFake(c) { Carga = suspenderCarga ? new TaskCompletionSource() : null }; janelas.Add(j); return j; },
            () => eventos.Add("mostrar"), () => eventos.Add("esconder"), () => eventos.Add("encerrar"));
        return (g, janelas, eventos);
    }

    [Fact] // Ctrl+F6: cicla entre as abertas na ordem de abertura; da ultima volta a primeira; com uma so' devolve ela mesma
    public async Task ProximaDepois_CiclaEntreAsAbertas()
    {
        var (g, janelas, _) = Montar();
        Assert.Null(g.ProximaDepois(new JanelaFake(null)));            // sem janelas
        await g.AbrirAsync(@"C:\a.docx");
        Assert.Same(janelas[0], g.ProximaDepois(janelas[0]));          // uma so'
        await g.AbrirAsync(@"C:\b.docx"); await g.AbrirAsync(@"C:\c.docx");
        Assert.Same(janelas[1], g.ProximaDepois(janelas[0]));
        Assert.Same(janelas[2], g.ProximaDepois(janelas[1]));
        Assert.Same(janelas[0], g.ProximaDepois(janelas[2]));          // da ultima volta a primeira
        janelas[1].Disparar();                                         // fechou uma: o ciclo se refaz sobre as que restam
        Assert.Same(janelas[2], g.ProximaDepois(janelas[0]));
    }

    [Fact] // abrir esconde a boas-vindas e cria a janela; o mesmo caminho so' ativa
    public async Task Abrir_CriaUmaVezEAtivaDepois()
    {
        var (g, janelas, eventos) = Montar();
        await g.AbrirAsync(@"C:\a.docx"); await g.AbrirAsync(@"c:\A.DOCX");
        Assert.Single(janelas); Assert.Equal(1, janelas[0].Ativacoes); Assert.Equal(1, g.Abertas);
        Assert.Equal(["esconder"], eventos);
    }

    [Fact] // o MESMO arquivo escrito de formas diferentes (relativo x absoluto) e' UMA janela so'
    public async Task Abrir_NormalizaOCaminho()
    {
        var (g, janelas, _) = Montar();
        var pasta = Directory.GetCurrentDirectory();
        await g.AbrirAsync(Path.Combine(pasta, "a.docx"));
        await g.AbrirAsync("a.docx");                                  // relativo ao MESMO diretorio
        await g.AbrirAsync(Path.Combine(pasta, ".", "a.docx"));        // com um componente redundante
        Assert.Single(janelas); Assert.Equal(2, janelas[0].Ativacoes); Assert.Equal(1, g.Abertas);
        Assert.Equal(Path.GetFullPath(Path.Combine(pasta, "a.docx")), janelas[0].Caminho);
    }

    [Fact] // dois duplos cliques seguidos no MESMO arquivo, antes de ele abrir, sao UMA janela (achado da revisao 2)
    public async Task Abrir_MesmoCaminhoDuranteCarregamento_SoAtiva()
    {
        var (g, janelas, _) = Montar(suspenderCarga: true);
        var primeiro = g.AbrirAsync(@"C:\a.docx");   // NAO aguardado: fica preso no carregamento
        Assert.False(primeiro.IsCompleted);
        Assert.Null(janelas[0].Caminho);             // a janela ainda nao sabe que arquivo tem

        var segundo = g.AbrirAsync(@"c:\A.DOCX");    // o segundo pedido chega antes de o documento abrir

        // Sem `await` no `segundo` ANTES de conferir (de proposito): se a deduplicacao falhar, ele cria uma
        // SEGUNDA janela e fica preso no carregamento dela — um `await` aqui penduraria a suite inteira em
        // vez de falhar. Deduplicado, ele so' ativa a janela existente e ja' nasce completo.
        Assert.True(segundo.IsCompleted, "o segundo pedido criou outra janela em vez de ativar a que ja' existia");
        Assert.Single(janelas);                      // uma janela so' — sem segunda sessao de recuperacao
        Assert.Equal(1, janelas[0].Ativacoes);       // o segundo pedido so' ATIVOU a que ja' existia
        Assert.Equal(1, g.Abertas);

        await segundo;
        janelas[0].Carga!.SetResult();
        await primeiro;
    }

    [Fact] // depois de "salvar como" b.docx, reabrir a.docx abre OUTRA janela (o pedido original ficou obsoleto);
           // reabrir b.docx so' ativa a janela que ja' e' dele (re-revisao 2 da Task 6)
    public async Task Abrir_DepoisDeSalvarComoOutroNome_ReabreOOriginalEmJanelaNova()
    {
        var (g, janelas, _) = Montar();
        await g.AbrirAsync(@"C:\a.docx");
        janelas[0].CaminhoAtual = @"C:\b.docx";        // a janela agora e' do b.docx
        await g.AbrirAsync(@"C:\a.docx");
        Assert.Equal(2, janelas.Count);                 // a.docx continua no disco: janela nova
        Assert.Equal(0, janelas[0].Ativacoes);
        await g.AbrirAsync(@"c:\B.DOCX");
        Assert.Equal(2, janelas.Count); Assert.Equal(1, janelas[0].Ativacoes); Assert.Equal(2, g.Abertas);
    }

    [Fact] // novo cria janela sem caminho; duas "novas" sao duas janelas
    public async Task Novo_SempreCriaJanela()
    {
        var (g, janelas, _) = Montar();
        await g.NovoAsync(); await g.NovoAsync();
        Assert.Equal(2, janelas.Count); Assert.All(janelas, j => Assert.Null(j.Caminho));
    }

    [Fact] // fechar a ultima janela encerra o processo (spec §4)
    public async Task FecharUltima_Encerra()
    {
        var (g, janelas, eventos) = Montar();
        await g.AbrirAsync(@"C:\a.docx"); await g.NovoAsync();
        janelas[0].Disparar(); Assert.DoesNotContain("encerrar", eventos);
        janelas[1].Disparar(); Assert.Contains("encerrar", eventos); Assert.Equal(0, g.Abertas);
    }

    [Fact] // fechar DURANTE o carregamento conta: a janela ja' esta' registrada antes do await (achado CRITICO da revisao 1)
    public async Task FecharDuranteCarregamento_EncerraUmaVez()
    {
        var (g, janelas, eventos) = Montar(suspenderCarga: true);
        var abrindo = g.AbrirAsync(@"C:\a.docx");

        Assert.Single(janelas);
        Assert.Equal(1, g.Abertas);              // registrada ANTES de o documento terminar de carregar
        Assert.False(abrindo.IsCompleted);

        janelas[0].Disparar();                   // o usuario fecha a janela enquanto o documento abre
        Assert.Equal(0, g.Abertas);
        Assert.Single(eventos, e => e == "encerrar");

        janelas[0].Carga!.SetResult();           // so' entao o carregamento termina
        await abrindo;

        // A janela morta NAO volta para o registro, e o encerramento aconteceu UMA vez so'.
        Assert.Equal(0, g.Abertas);
        Assert.Single(eventos, e => e == "encerrar");
    }

    [Fact] // carregamento que falha: a janela sai do registro, e' fechada e a excecao sobe — sem encerrar o app
    public async Task CarregamentoQueFalha_TiraDoRegistroEFechaSemEncerrar()
    {
        var (g, janelas, eventos) = Montar(suspenderCarga: true);
        var abrindo = g.AbrirAsync(@"C:\a.docx");
        janelas[0].Carga!.SetException(new InvalidOperationException("arquivo corrompido"));

        var erro = await Assert.ThrowsAsync<InvalidOperationException>(() => abrindo);
        Assert.Equal("arquivo corrompido", erro.Message);
        Assert.Equal(0, g.Abertas);
        Assert.Equal(1, janelas[0].Fechamentos);          // o gerente mandou fechar
        // NAO encerra: quem trata e' a casca, mostrando a boas-vindas com a mensagem de erro. Se o
        // fechamento acontecesse com a janela ainda registrada, o AoFechar veria "ultima janela fechada".
        Assert.DoesNotContain("encerrar", eventos);
    }

    [Fact] // sem janelas na subida: mostra a boas-vindas; com janela: nao
    public async Task MostrarBoasVindasSeVazio()
    {
        var (g, _, eventos) = Montar();
        g.MostrarBoasVindasSeVazio(); Assert.Equal(["mostrar"], eventos);
        await g.NovoAsync(); g.MostrarBoasVindasSeVazio(); Assert.Equal(["mostrar", "esconder"], eventos);
    }

    // --- Recuperacao (Task 3 do Plano 3) -------------------------------------------------------------------

    [Fact] // janela de fabrica propria (recuperar): mesma politica — esconde, registra, fechar a ultima encerra
    public async Task AbrirPor_MesmaPoliticaDoNovo()
    {
        var (g, janelas, eventos) = Montar();
        var recuperada = new JanelaFake(@"C:\x\peticao.docx");
        await g.AbrirPorAsync(() => recuperada);
        Assert.Empty(janelas);                     // a fabrica padrao nao foi usada
        Assert.Equal(1, g.Abertas); Assert.Equal(["esconder"], eventos);
        Assert.True(g.EstaAberto(@"c:\X\PETICAO.DOCX"));
        await g.AbrirAsync(@"C:\x\peticao.docx"); // o mesmo arquivo depois de carregado: so' ativa
        Assert.Equal(1, recuperada.Ativacoes); Assert.Empty(janelas);
        recuperada.Disparar();
        Assert.Equal(["esconder", "encerrar"], eventos);
        Assert.False(g.EstaAberto(@"C:\x\peticao.docx"));
    }

    [Fact] // a recuperacao que falhou depois de carregar: a janela SOLTA fecha sozinha sem encerrar o app
    public async Task Soltar_FechamentoNaoEncerra()
    {
        var (g, _, eventos) = Montar();
        var recuperada = new JanelaFake(null);
        await g.AbrirPorAsync(() => recuperada);
        Assert.True(g.Soltar(recuperada));
        Assert.Equal(0, g.Abertas);
        recuperada.Disparar();                     // a janela se fecha (JanelaEditor.FalharRecuperacao)
        Assert.DoesNotContain("encerrar", eventos);
        Assert.False(g.Soltar(recuperada));        // ja' fora do registro
    }

    [Fact] // controle do caso acima: SEM soltar, o mesmo fechamento da ultima janela encerra
    public async Task SemSoltar_FechamentoDaUltimaEncerra()
    {
        var (g, _, eventos) = Montar();
        var recuperada = new JanelaFake(null);
        await g.AbrirPorAsync(() => recuperada);
        recuperada.Disparar();
        Assert.Contains("encerrar", eventos);
    }

    [Fact] // Plano 5 Task 2 ("instalar agora"): pede o fechamento de TODAS; a ultima a fechar encerra UMA vez
    public async Task FecharTodas_FechaCadaUmaEEncerra()
    {
        var (g, janelas, eventos) = Montar();
        await g.AbrirAsync(@"C:\a.docx"); await g.AbrirAsync(@"C:\b.docx"); await g.NovoAsync();
        g.FecharTodas();
        Assert.All(janelas, j => Assert.Equal(1, j.Fechamentos));
        Assert.Equal(0, g.Abertas);
        Assert.Single(eventos, e => e == "encerrar");
    }

    [Fact] // um documento recusa (Cancelar na pergunta de salvar): o app NAO encerra e a janela continua registrada
    public async Task FecharTodas_UmaRecusa_NaoEncerra()
    {
        var (g, janelas, eventos) = Montar();
        await g.AbrirAsync(@"C:\a.docx"); await g.AbrirAsync(@"C:\b.docx");
        janelas[1].Recusar = true;
        g.FecharTodas();
        Assert.Equal(1, g.Abertas);
        Assert.DoesNotContain("encerrar", eventos);
    }
}
