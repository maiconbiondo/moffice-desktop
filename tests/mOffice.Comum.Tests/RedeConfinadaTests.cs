using System.Text.RegularExpressions;

namespace mOffice.Comum.Tests;

/// <summary>
/// Plano 5, Task 2 — REDE CONFINADA A UM ARQUIVO (porte de <c>UpdateNetworkConfinementTests</c> do mPDF). O mOffice tem
/// egresso zero no uso normal; a unica rede do produto e' a atualizacao sob demanda, e ela mora INTEIRA em
/// <c>src/mOffice.Comum/Atualizacao/ServicoAtualizacao.cs</c>. Duas guardas textuais, independentes da
/// <c>RedeTests</c> do nucleo (outra lista de tokens, outra varredura — a mesma fronteira vista duas vezes):
///
///   1. nenhum token da familia de rede da BCL em QUALQUER <c>.cs</c> de <c>src/</c> (todos os projetos) fora do
///      arquivo permitido. <c>Contains</c> simples de proposito: qualquer MENCAO ao tipo fora dali ja' e' suspeita,
///      inclusive formas qualificadas. Comentario de linha inteira e' pulado (residuo aceito, como no mPDF).
///   2. o <see cref="Atualizacao.ServicoAtualizacao"/> e' CONSTRUIDO em exatamente dois sitios deliberados — o botao
///      das Configuracoes (<c>DialogoConfiguracoesViewModel</c>) e os verbos do RMM (<c>ExecutorDeAtualizacao</c>) —,
///      com deteccao de alias (<c>using U = ...ServicoAtualizacao;</c>), como no mPDF.
///
/// RESIDUO ACEITO (o mesmo do mPDF): reflexao (<c>Activator.CreateInstance</c>) e P/Invoke de rede nativo.
/// </summary>
public class RedeConfinadaTests
{
    private static string Raiz => AplicativoOfficeTests.RaizRepo();

    private static string Permitido => Path.Combine(Raiz, "src", "mOffice.Comum", "Atualizacao", "ServicoAtualizacao.cs");

    private static bool EhArtefato(string caminho)
    {
        var partes = caminho.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return partes.Contains("obj") || partes.Contains("bin");
    }

    private static List<string> Fontes() =>
        Directory.EnumerateFiles(Path.Combine(Raiz, "src"), "*.cs", SearchOption.AllDirectories).Where(f => !EhArtefato(f)).ToList();

    /// <summary>A familia de rede da BCL (lista do mPDF + UdpClient/HttpListener/HttpMessageHandler e os namespaces).
    /// "Socket" nu cobre TcpClient/ClientWebSocket/SocketsHttpHandler; "WebRequest" cobre HttpWebRequest.</summary>
    internal static readonly string[] Tokens =
    [
        "HttpClient", "HttpMessageHandler", "HttpListener", "WebRequest", "WebClient", "Socket", "TcpClient", "UdpClient",
        "Dns", "System.Net.Http", "System.Net.Sockets",
    ];

    [Fact] // piso: a varredura enxerga os projetos do produto (senao a guarda passaria por vacuidade)
    public void Varredura_NaoVazia()
    {
        var fontes = Fontes();
        Assert.True(fontes.Count >= 40, $"varredura quase vazia ({fontes.Count} arquivos) = guarda morta");
        foreach (var projeto in new[] { "mOffice.Comum", "mOffice.Nucleo", "mOffice.Conversor", "mTexto", "mPlanilha", "mSlides" })
            Assert.Contains(fontes, f => f.Contains(Path.DirectorySeparatorChar + projeto + Path.DirectorySeparatorChar));
        Assert.Contains(Permitido, fontes, StringComparer.OrdinalIgnoreCase);
    }

    internal static string? Ofensa(string linha)
    {
        var t = linha.TrimStart();
        if (t.StartsWith("//", StringComparison.Ordinal)) return null;
        return Tokens.FirstOrDefault(x => t.Contains(x, StringComparison.Ordinal));
    }

    [Fact]
    public void Rede_SoNoArquivoDaAtualizacao()
    {
        var ofensores = new List<string>();
        foreach (var f in Fontes())
        {
            if (string.Equals(f, Permitido, StringComparison.OrdinalIgnoreCase)) continue;
            foreach (var (linha, n) in File.ReadLines(f).Select((l, i) => (l, i + 1)))
                if (Ofensa(linha) is { } token) { ofensores.Add($"{Path.GetRelativePath(Raiz, f)}:{n}: [{token}] {linha.Trim()}"); break; }
        }
        Assert.True(ofensores.Count == 0,
            "rede fora da fronteira (so' src/mOffice.Comum/Atualizacao/ServicoAtualizacao.cs pode):\n" + string.Join("\n", ofensores));
    }

    [Fact] // o arquivo permitido usa rede de verdade (a excecao nao isenta um arquivo vazio de rede por engano)
    public void ArquivoPermitido_UsaRede() =>
        Assert.Contains(File.ReadLines(Permitido), l => Ofensa(l) == "HttpClient");

    [Theory] // capacidade do detector: pega o uso tipico...
    [InlineData("        using var c = new HttpClient();")]
    [InlineData("        var c = new System.Net.Http.HttpClient();")]
    [InlineData("using System.Net.Http;")]
    [InlineData("        var r = WebRequest.Create(u);")]
    [InlineData("        using var w = new WebClient();")]
    [InlineData("        using var t = new TcpClient(h, 443);")]
    [InlineData("        using var u = new UdpClient();")]
    [InlineData("        var s = new Socket(a, b, c);")]
    [InlineData("        var ip = Dns.GetHostAddresses(h);")]
    [InlineData("        var ws = new ClientWebSocket();")]
    public void Detector_PegaUsoTipico(string linha) => Assert.NotNull(Ofensa(linha));

    [Theory] // ...e nao pega comentario de linha inteira nem codigo inocente
    [InlineData("// new HttpClient() em comentario")]
    [InlineData("    /// <summary>HttpClient em doc</summary>")]
    [InlineData("        var p = new NamedPipeServerStream(nome);")]
    [InlineData("        var x = Path.Combine(a, b);")]
    public void Detector_IgnoraComentarioEInocente(string linha) => Assert.Null(Ofensa(linha));

    // ---- guarda 2: ServicoAtualizacao construido em exatamente 2 sitios (com deteccao de alias) --------------

    private static readonly Regex Alias = new(
        @"^\s*(?:global\s+)?using\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(?:global::)?(?:[A-Za-z_][A-Za-z0-9_]*\.)*ServicoAtualizacao\s*;",
        RegexOptions.Compiled);

    internal static HashSet<string> Nomes(IEnumerable<string> linhas)
    {
        var nomes = new HashSet<string> { "ServicoAtualizacao" };
        foreach (var l in linhas) if (Alias.Match(l) is { Success: true } m) nomes.Add(m.Groups[1].Value);
        return nomes;
    }

    internal static int Construcoes(IEnumerable<string> nomes, string linha)
    {
        var alt = string.Join("|", nomes.Select(Regex.Escape));
        return Regex.Matches(linha, $@"(?<![\w.])new\s+(?:global::)?(?:[A-Za-z_][A-Za-z0-9_]*\.)*(?:{alt})\s*\(").Count;
    }

    [Fact]
    public void ServicoAtualizacao_ConstruidoSoNosDoisSitios()
    {
        var fontes = Fontes();
        var nomes = Nomes(fontes.SelectMany(File.ReadLines));
        var sitios = new List<string>();
        foreach (var f in fontes)
            foreach (var (linha, n) in File.ReadLines(f).Select((l, i) => (l, i + 1)))
            {
                if (linha.TrimStart().StartsWith("//", StringComparison.Ordinal)) continue;
                for (var k = Construcoes(nomes, linha); k > 0; k--) sitios.Add($"{Path.GetRelativePath(Raiz, f)}:{n}");
            }
        Assert.True(sitios.Count == 2, "ServicoAtualizacao deve ser construido em EXATAMENTE 2 sitios (botao e verbos do RMM): " + string.Join(", ", sitios));
        Assert.Contains(sitios, s => s.Contains("DialogoConfiguracoesViewModel.cs"));
        Assert.Contains(sitios, s => s.Contains("ExecutorDeAtualizacao.cs"));
    }

    [Theory]
    [InlineData("var x = new ServicoAtualizacao(f);")]
    [InlineData("using var x = new  ServicoAtualizacao(f);")]
    [InlineData("var x = new mOffice.Comum.Atualizacao.ServicoAtualizacao(f);")]
    [InlineData("var x = new global::mOffice.Comum.Atualizacao.ServicoAtualizacao(f);")]
    public void Construcoes_ReconheceGrafias(string linha) => Assert.Equal(1, Construcoes(["ServicoAtualizacao"], linha));

    [Theory]
    [InlineData("var x = new OutroServicoAtualizacao(f);")]
    [InlineData("var v = ServicoAtualizacao.VersaoAtualTexto();")]
    public void Construcoes_NaoPegaMencao(string linha) => Assert.Equal(0, Construcoes(["ServicoAtualizacao"], linha));

    [Fact] // alias: `using U = ...ServicoAtualizacao;` + `new U(` conta; alias de outro tipo nao
    public void Construcoes_ComAlias()
    {
        string[] comAlias = ["using U = mOffice.Comum.Atualizacao.ServicoAtualizacao;", "var x = new U(f);"];
        Assert.Equal(1, comAlias.Sum(l => Construcoes(Nomes(comAlias), l)));
        string[] outro = ["using U = System.String;", "var x = new U();"];
        Assert.Equal(0, outro.Sum(l => Construcoes(Nomes(outro), l)));
    }
}
