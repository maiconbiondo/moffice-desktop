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
