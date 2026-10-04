using mOffice.Comum;

namespace mOffice.Comum.Tests;

public class InstanciaUnicaTests
{
    private static (string mutex, string pipe) Nomes() { var g = Guid.NewGuid().ToString("N"); return ($"mOffice-teste-{g}-mutex", $"mOffice-teste-{g}-pipe"); }

    [Fact] // a primeira assume; a segunda nao assume e entrega o caminho a primeira
    public async Task Segunda_EncaminhaCaminhoParaAPrimeira()
    {
        var (m, p) = Nomes();
        using var primeira = new InstanciaUnica(m, p);
        var recebido = new TaskCompletionSource<string>();
        primeira.CaminhoRecebido += c => recebido.TrySetResult(c);
        Assert.True(primeira.TentarAssumir(null));
        using var segunda = new InstanciaUnica(m, p);
        Assert.False(segunda.TentarAssumir(@"C:\pasta com espaço\petição.docx"));
        Assert.Equal(@"C:\pasta com espaço\petição.docx", await recebido.Task.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact] // fix round 1 da Task 5: a segunda SEM arquivo pede so' "mostre-se" (o app pode estar vivo sem janela,
           // esperando uma impressao) — AtivacaoRecebida, e nunca CaminhoRecebido
    public async Task Segunda_SemArquivo_PedeAtivacao()
    {
        var (m, p) = Nomes();
        using var primeira = new InstanciaUnica(m, p);
        var ativou = new TaskCompletionSource();
        var caminhos = new List<string>();
        primeira.AtivacaoRecebida += () => ativou.TrySetResult();
        primeira.CaminhoRecebido += caminhos.Add;
        Assert.True(primeira.TentarAssumir(null));
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir(null));
        await ativou.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Empty(caminhos);
        Assert.False(Path.IsPathRooted(InstanciaUnica.PedidoDeAtivacao)); // primaria antiga (so' caminho absoluto) ignora
    }

    [Fact] // caminho relativo e' ignorado pela primeira (resolveria contra o cwd errado)
    public async Task CaminhoRelativo_Ignorado()
    {
        var (m, p) = Nomes();
        using var primeira = new InstanciaUnica(m, p);
        var recebidos = new List<string>();
        primeira.CaminhoRecebido += recebidos.Add;
        Assert.True(primeira.TentarAssumir(null));
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir("relativo.docx"));
        await Task.Delay(500);
        Assert.Empty(recebidos);
    }

    [Fact] // Plano 6, Task 2: a linha "?modelo|<caminho absoluto>" passa pelo pipe e chega a CaminhoRecebido intacta; "?modelo|relativo" nao
    public async Task PipeAceitaPedidoComModo()
    {
        var (m, p) = Nomes();
        using var primeira = new InstanciaUnica(m, p);
        var recebidos = new List<string>();
        var primeiro = new TaskCompletionSource<string>();
        primeira.CaminhoRecebido += c => { lock (recebidos) recebidos.Add(c); primeiro.TrySetResult(c); };
        Assert.True(primeira.TentarAssumir(null));
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir(@"?modelo|m.dotx"));   // relativo: descartado
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir(@"?modelo|C:\x\m.dotx"));
        Assert.Equal(@"?modelo|C:\x\m.dotx", await primeiro.Task.WaitAsync(TimeSpan.FromSeconds(5)));
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir(@"?apresentar|C:\x\s.ppsx"));
        for (var i = 0; i < 50; i++) { lock (recebidos) if (recebidos.Count >= 2) break; await Task.Delay(100); }
        lock (recebidos) Assert.Equal([@"?modelo|C:\x\m.dotx", @"?apresentar|C:\x\s.ppsx"], recebidos);
    }

    [Fact] // FMT-05: a segunda instancia manda a rejeicao e a primaria a recebe (RejeicaoRecebida), sem abrir nada
    public async Task SegundaInstancia_RejeicaoChegaAPrimaria()
    {
        var (m, p) = Nomes();
        using var primeira = new InstanciaUnica(m, p);
        var rej = new TaskCompletionSource<string>(); var abertos = new List<string>();
        primeira.RejeicaoRecebida += c => rej.TrySetResult(c);
        primeira.CaminhoRecebido += abertos.Add;
        Assert.True(primeira.TentarAssumir(null));
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir(PedidoDeAbertura.RejeicaoParaPipe(@"C:\x\a.xlsx")));
        Assert.Equal(@"C:\x\a.xlsx", await rej.Task.WaitAsync(TimeSpan.FromSeconds(5)));
        Assert.Empty(abertos);
    }

    [Fact] // um modo inventado ("?formatar|...") nunca chega a CaminhoRecebido
    public async Task PipeIgnoraModoDesconhecido()
    {
        var (m, p) = Nomes();
        using var primeira = new InstanciaUnica(m, p);
        var recebidos = new List<string>();
        primeira.CaminhoRecebido += recebidos.Add;
        Assert.True(primeira.TentarAssumir(null));
        using (var s = new InstanciaUnica(m, p)) Assert.False(s.TentarAssumir(@"?formatar|C:\x\a.docx"));
        await Task.Delay(500);
        Assert.Empty(recebidos);
    }

    [Fact] public void Nomes_CarregamAppUsuarioESessao()
    {
        Assert.Contains("mtexto", NomesDeInstancia.Mutex("mtexto")); Assert.Contains(Environment.UserName.Replace('\\', '_'), NomesDeInstancia.Pipe("mtexto"));
        Assert.NotEqual(NomesDeInstancia.Pipe("mtexto"), NomesDeInstancia.Pipe("mplanilha"));
    }

    [Fact] // fail-open: excecao ao assumir e' logada e o app continua; sem excecao, devolve o que TentarAssumir devolveu
    public void DeveContinuar_ExcecaoLogaEContinua()
    {
        Exception? logada = null;
        Assert.True(InstanciaUnica.DeveContinuar(() => throw new IOException("pipe indisponivel"), e => logada = e));
        Assert.IsType<IOException>(logada);
        Assert.False(InstanciaUnica.DeveContinuar(() => false, _ => Assert.Fail("nao deveria logar")));
    }

    [Fact] // liberado no Dispose: a proxima instancia assume
    public void Dispose_LiberaOMutex()
    {
        var (m, p) = Nomes();
        var a = new InstanciaUnica(m, p); Assert.True(a.TentarAssumir(null)); a.Dispose();
        using var b = new InstanciaUnica(m, p); Assert.True(b.TentarAssumir(null));
    }
}
