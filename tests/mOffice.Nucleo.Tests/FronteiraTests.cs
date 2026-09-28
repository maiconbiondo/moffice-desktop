namespace mOffice.Nucleo.Tests;

/// <summary>Guardas de fronteira do produto (spec Marco 1 §4): tipos do WebView2 so' dentro do nucleo (e do
/// executavel de apoio); nenhum iText em lugar nenhum. Varre o TEXTO dos .cs e .csproj de src/ — o compilador
/// nao impede um app de referenciar o pacote, so' esta varredura impede.</summary>
public class FronteiraTests
{
    // internal static (nao mais private): RedeTests (Task 7) reusa a mesma varredura de fontes em vez de
    // duplicar a logica de achar a raiz do repo e enumerar arquivos.
    internal static string RaizRepo()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d is not null && !File.Exists(Path.Combine(d.FullName, "mOffice.slnx"))) d = d.Parent;
        return d?.FullName ?? throw new InvalidOperationException("mOffice.slnx não encontrado acima do assembly de testes");
    }

    internal static IEnumerable<string> Fontes(string subpasta, params string[] extensoes) =>
        Directory.EnumerateFiles(Path.Combine(RaizRepo(), subpasta), "*.*", SearchOption.AllDirectories)
            .Where(f => extensoes.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase))
            .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}") && !f.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}"));

    [Fact] // WebView2 so' em src/mOffice.Nucleo e src/mOffice.Sonda (Task 7: Comum e os tres apps ficam proibidos)
    public void WebView2_SoNoNucleoENaSonda()
    {
        var permitidos = new[] { "mOffice.Nucleo", "mOffice.Sonda" };
        var arquivos = Fontes("src", ".cs", ".csproj", ".xaml").ToList();
        // Piso de varredura nao vazia (licao do AgplGuardTests.SrcTree_ScanIsNotEmpty do mPDF): sem isto, um
        // caminho errado ou um filtro quebrado faria a varredura devolver zero arquivos e os dois testes
        // abaixo passariam por vacuidade, sem provar nada. Exige explicitamente que mOffice.Comum — um dos
        // projetos novos da Task 6 — tenha sido ENUMERADO.
        Assert.Contains(arquivos, a => a.Contains("mOffice.Comum"));
        var violacoes = arquivos
            .Where(f => !permitidos.Any(p => f.Contains($"{Path.DirectorySeparatorChar}{p}{Path.DirectorySeparatorChar}")))
            .Where(f => File.ReadAllText(f).Contains("Microsoft.Web.WebView2"))
            .ToList();
        Assert.True(violacoes.Count == 0, "WebView2 fora do nucleo: " + string.Join(", ", violacoes));
    }

    [Fact] // iText (AGPL do mPDF) nunca entra no mOffice — a assinatura e' pelo mPDF instalado
    public void IText_EmLugarNenhum()
    {
        // Exclui este proprio arquivo: o nome do guard e' texto na fronteira que ele mesmo varre, e citar a
        // palavra proibida em prosa (comentarios, mensagem de assert) nao e' uma violacao de verdade.
        var violacoes = Fontes("src", ".cs", ".csproj").Concat(Fontes("tests", ".cs", ".csproj"))
            .Where(f => Path.GetFileName(f) != nameof(FronteiraTests) + ".cs")
            .Where(f => File.ReadAllText(f).Contains("itext", StringComparison.OrdinalIgnoreCase)).ToList();
        Assert.True(violacoes.Count == 0, "iText encontrado: " + string.Join(", ", violacoes));
    }

    [Fact] // os projetos de produto existem com os nomes da spec (falha se a promocao ficou pela metade)
    public void ProjetosDoProduto_Existem()
    {
        foreach (var p in new[]
                 {
                     "src/mOffice.Conversor/mOffice.Conversor.csproj",
                     "src/mOffice.Nucleo/mOffice.Nucleo.csproj",
                     "src/mOffice.Sonda/mOffice.Sonda.csproj",
                     // Task 7: os tres apps e o Comum da Task 6, mais o Directory.Build.props compartilhado —
                     // falha se a promocao do Plano 2 ficou pela metade.
                     "src/mOffice.Comum/mOffice.Comum.csproj",
                     "src/mTexto/mTexto.csproj",
                     "src/mPlanilha/mPlanilha.csproj",
                     "src/mSlides/mSlides.csproj",
                     "src/Directory.Build.props",
                 })
            Assert.True(File.Exists(Path.Combine(RaizRepo(), p)), p + " ausente");
        Assert.False(Directory.Exists(Path.Combine(RaizRepo(), "poc", "mOffice.Poc.Host")), "poc/mOffice.Poc.Host ainda existe");
    }
}
