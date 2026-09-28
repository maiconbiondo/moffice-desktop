namespace mOffice.Nucleo.Tests;

/// <summary>
/// Guarda de fronteira do produto (Task 7): nenhum cliente de rede em <c>src/</c> — a suite fala com o mundo
/// so' pelo WebView2 (Nucleo/Sonda) e pelo x2t (processo local). A UNICA excecao, desde o Plano 5 (Task 2), e' o
/// arquivo da atualizacao sob demanda (<see cref="Permitidos"/>), por NOME EXATO, como no mPDF. A mesma fronteira
/// tem uma segunda varredura independente em <c>mOffice.Comum.Tests/RedeConfinadaTests</c>.
///
/// O que esta guarda REALMENTE faz (e o que nao faz): varre o TEXTO dos <c>.cs</c> e dos <c>.csproj</c> de
/// <c>src/</c> procurando (a) os tipos de <see cref="Proibidos"/> em qualquer linha de codigo e (b) os
/// pacotes de <see cref="PacotesProibidos"/> SO' em linhas de <c>PackageReference</c> de <c>.csproj</c>.
/// Linhas de comentario — de linha (<c>//</c>, <c>///</c>) e de bloco (<c>/* … */</c>,
/// <c>&lt;!-- … --&gt;</c>, inclusive multilinha) — sao puladas por <see cref="LinhasDeCodigo"/>.
///
/// Nao e' analise semantica: nao enxerga reflexao, nem um alias de `using`, nem uma dependencia
/// TRANSITIVA que abra soquete por baixo, nem um comentario no meio de uma linha de codigo. E' uma barreira
/// contra a introducao DISTRAIDA de codigo de rede — o mesmo papel (e os mesmos limites) das demais guardas
/// de <see cref="FronteiraTests"/>.
///
/// <c>.csproj</c> entra na varredura porque um <c>PackageReference</c> de cliente HTTP e' exatamente a forma
/// mais facil de a rede voltar sem nenhum <c>.cs</c> mudar.
/// </summary>
public class RedeTests
{
    /// <summary>Clientes e primitivas de rede da BCL. <c>Socket(</c> com o parentese de proposito: o token
    /// nu casaria com <c>SocketException</c>, <c>UnixDomainSocketEndPoint</c> e por ai afora. Os pipes
    /// nomeados de <c>InstanciaUnica</c> NAO estao aqui: <c>NamedPipeServerStream</c> nao e' soquete e nao
    /// sai da maquina — e' o canal da instancia unica, parte do produto.</summary>
    internal static readonly string[] Proibidos =
    [
        "HttpClient", "HttpMessageHandler", "HttpListener", "WebRequest",
        "TcpClient", "UdpClient", "Socket(", "Dns.",
    ];

    /// <summary>Pacotes NuGet que trazem um cliente HTTP junto. Conferidos SO' em linhas de
    /// <c>PackageReference</c> de <c>.csproj</c> (achado da revisao 2): <c>System.Net.Http</c> casaria com o
    /// <c>using System.Net.Http;</c> de qualquer arquivo, e num <c>.cs</c> quem denuncia o uso sao os tipos de
    /// <see cref="Proibidos"/> — e' no <c>.csproj</c> que a DEPENDENCIA entra sem nenhum <c>.cs</c> mudar.</summary>
    internal static readonly string[] PacotesProibidos =
    [
        "System.Net.Http", "RestSharp", "Flurl", "Refit", "Polly.Extensions.Http", "Microsoft.Extensions.Http",
    ];

    /// <summary>Os UNICOS arquivos de <c>src/</c> que podem usar rede, pelo caminho relativo EXATO (nunca por pasta
    /// nem por prefixo: um arquivo novo ao lado nao herda a excecao). Plano 5, Task 2: a atualizacao sob demanda
    /// (GitHub Releases), acionada so' pelo botao das Configuracoes e pelos verbos do RMM.</summary>
    internal static readonly string[] Permitidos =
    [
        Path.Combine("src", "mOffice.Comum", "Atualizacao", "ServicoAtualizacao.cs"),
    ];

    private static bool EhPermitido(string raiz, string arquivo) =>
        Permitidos.Any(p => string.Equals(Path.GetRelativePath(raiz, arquivo), p, StringComparison.OrdinalIgnoreCase));

    [Fact] // nenhum cliente de rede em src/, fora o arquivo da atualizacao (Plano 5), permitido por nome exato
    public void NenhumClienteDeRedeEmSrc()
    {
        var raiz = FronteiraTests.RaizRepo();
        var arquivos = FronteiraTests.Fontes("src", ".cs", ".csproj").Where(a => !EhPermitido(raiz, a)).ToList();
        // Piso de varredura nao vazia (mesma licao do FronteiraTests): sem isto um filtro quebrado devolveria
        // zero arquivos e a guarda passaria por vacuidade. Exige os dois tipos que ela diz varrer.
        Assert.Contains(arquivos, a => a.EndsWith(".cs", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(arquivos, a => a.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase));

        var ofensores = arquivos
            .SelectMany(a => LinhasDeCodigo(a).Select(x => (arquivo: a, x.linha, x.numero)))
            .Where(x => Proibidos.Any(p => x.linha.Contains(p, StringComparison.Ordinal))
                     || (EhPackageReference(x.arquivo, x.linha)
                         && PacotesProibidos.Any(p => x.linha.Contains(p, StringComparison.Ordinal))))
            .Select(x => $"{Path.GetRelativePath(raiz, x.arquivo)}:{x.numero}")
            .ToList();
        Assert.True(ofensores.Count == 0, "codigo de rede em src/: " + string.Join(", ", ofensores));
    }

    [Fact] // a excecao e' UMA, por nome exato, e esta' viva: o arquivo existe, usa rede de verdade e esta' fora da varredura
    public void Permitido_UmSoArquivo_ExisteEUsaRede()
    {
        var raiz = FronteiraTests.RaizRepo();
        var permitido = Assert.Single(Permitidos);
        var caminho = Path.Combine(raiz, permitido);
        Assert.True(File.Exists(caminho), "o arquivo permitido sumiu (excecao morta): " + permitido);
        // Controle positivo: sem a excecao ele SERIA pego (senao a excecao nao isenta nada e so' esconde um erro de nome).
        Assert.Contains(LinhasDeCodigo(caminho), x => Proibidos.Any(p => x.linha.Contains(p, StringComparison.Ordinal)));
        Assert.True(EhPermitido(raiz, caminho));
        // Nome EXATO: um vizinho na mesma pasta, ou o mesmo nome em outro projeto, nao herda a excecao.
        Assert.False(EhPermitido(raiz, Path.Combine(raiz, "src", "mOffice.Comum", "Atualizacao", "ExecutorDeAtualizacao.cs")));
        Assert.False(EhPermitido(raiz, Path.Combine(raiz, "src", "mOffice.Nucleo", "ServicoAtualizacao.cs")));
        Assert.False(EhPermitido(raiz, Path.Combine(raiz, "src", "mOffice.Comum", "Atualizacao", "ServicoAtualizacao.cs.bak")));
    }

    /// <summary>Linha de <c>PackageReference</c> de um <c>.csproj</c> — o unico lugar onde
    /// <see cref="PacotesProibidos"/> vale.</summary>
    private static bool EhPackageReference(string arquivo, string linha) =>
        arquivo.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase)
        && linha.Contains("PackageReference", StringComparison.Ordinal);

    /// <summary>
    /// As linhas do arquivo que sao CODIGO, com o numero (base 1): fora as de comentario de linha
    /// (<c>//</c>, <c>///</c>) e as que estao dentro de um bloco <c>/* … */</c> ou <c>&lt;!-- … --&gt;</c>.
    /// Prosa que CITA um tipo de rede — como esta propria classe faz — nao e' uma violacao.
    ///
    /// E' uma maquina de estados DELIBERADAMENTE simples: conta so' abertura e fechamento de bloco no comeco
    /// e no fim da linha. Nao entende comentario dentro de string literal, nem <c>/* */</c> no meio de uma
    /// linha de codigo. Para o que esta guarda faz — achar a introducao distraida de um cliente de rede —
    /// isso basta, e os dois modos de erro sao aceitaveis: no maximo ela deixa passar uma linha exotica
    /// (o que um humano veria no diff) ou acusa uma linha comentada (que o autor arruma em 10 segundos).
    /// </summary>
    private static IEnumerable<(string linha, int numero)> LinhasDeCodigo(string arquivo)
    {
        var emBloco = false;
        var n = 0;
        foreach (var bruta in File.ReadLines(arquivo))
        {
            n++;
            var t = bruta.TrimStart();
            if (emBloco)
            {
                // A linha que FECHA o bloco ainda e' comentario da parte que interessa aqui.
                if (t.Contains("*/", StringComparison.Ordinal) || t.Contains("-->", StringComparison.Ordinal)) emBloco = false;
                continue;
            }
            if (t.StartsWith("//", StringComparison.Ordinal)) continue;
            if ((t.StartsWith("/*", StringComparison.Ordinal) && !t.Contains("*/", StringComparison.Ordinal))
                || (t.StartsWith("<!--", StringComparison.Ordinal) && !t.Contains("-->", StringComparison.Ordinal)))
            {
                emBloco = true;
                continue;
            }
            // Bloco que abre e fecha na MESMA linha (<!-- x --> ou /* x */): e' so' comentario.
            if (t.StartsWith("/*", StringComparison.Ordinal) || t.StartsWith("<!--", StringComparison.Ordinal)) continue;
            yield return (bruta, n);
        }
    }

    [Fact] // a guarda enxerga de verdade cada token da lista (controle positivo, sem tocar em src/)
    public void Proibidos_CasamComOUsoTipicoENaoComComentario()
    {
        var usos = new[]
        {
            "        using var c = new HttpClient();",
            "        var h = new SocketsHttpHandler() as HttpMessageHandler;",
            "        var l = new HttpListener();",
            "        var r = WebRequest.Create(u);",
            "        using var t = new TcpClient(host, 443);",
            "        using var u2 = new UdpClient();",
            "        var s = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);",
            "        var ip = Dns.GetHostEntry(nome);",
        };
        foreach (var uso in usos)
            Assert.True(Proibidos.Any(p => uso.Contains(p, StringComparison.Ordinal)), "nao pegou: " + uso);

        // E NAO casa com o que so' se parece: excecao de soquete e o pipe nomeado da instancia unica.
        foreach (var inocente in new[]
                 {
                     "        catch (SocketException) { }",
                     "        var p = new NamedPipeServerStream(nome, PipeDirection.In);",
                     "        var e = new UnixDomainSocketEndPoint(caminho);",
                 })
            Assert.False(Proibidos.Any(p => inocente.Contains(p, StringComparison.Ordinal)), "falso positivo: " + inocente);
    }

    [Fact] // controle positivo do ramo .csproj: cada pacote proibido e' pego numa linha de PackageReference DE VERDADE
    public void PacotesProibidos_SaoPegosEmLinhaDePackageReference()
    {
        foreach (var pacote in PacotesProibidos)
        {
            // A linha e' a que um desenvolvedor escreveria de verdade — o nome do PACOTE no Include, sem
            // nenhum comentario para o casamento se apoiar (era o vicio do controle anterior, que casava
            // pelo "<!-- HttpClient por baixo -->" e nao pelo nome do pacote: provava a coisa errada).
            var linha = $"    <PackageReference Include=\"{pacote}\" Version=\"9.0.0\" />";
            Assert.True(EhPackageReference("x.csproj", linha), "nao reconheceu como PackageReference: " + linha);
            Assert.True(PacotesProibidos.Any(p => linha.Contains(p, StringComparison.Ordinal)), "nao pegou: " + linha);
        }

        // O ramo e' ESTREITO de proposito: pacote so' vale em .csproj E so' em linha de PackageReference.
        var usando = "using System.Net.Http;";
        Assert.False(EhPackageReference("A.cs", usando));                                       // .cs nunca entra
        Assert.False(EhPackageReference("x.csproj", "    <ProjectReference Include=\"..\\a\\a.csproj\" />"));
        // Um pacote inocente continua passando (a lista nao e' "qualquer PackageReference").
        var inocente = "    <PackageReference Include=\"xunit\" Version=\"2.9.3\" />";
        Assert.DoesNotContain(PacotesProibidos, p => inocente.Contains(p, StringComparison.Ordinal));
    }

    [Fact] // a varredura pula comentario de linha E bloco (/* */ e <!-- -->), inclusive multilinha
    public void LinhasDeCodigo_IgnoramComentarioDeLinhaEDeBloco()
    {
        var arquivo = Path.Combine(DadosDeTeste.RaizDeTestes, Guid.NewGuid().ToString("N") + ".cs");
        Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!);
        try
        {
            File.WriteAllLines(arquivo,
            [
                "// HttpClient em comentario de linha",
                "/// <summary>WebRequest em doc comment</summary>",
                "/*",
                "   TcpClient dentro de bloco",
                "*/",
                "<!-- UdpClient dentro de bloco XML",
                "     HttpListener na segunda linha do bloco -->",
                "/* Dns. num bloco de uma linha so' */",
                "var x = 1;                     // codigo de verdade",
            ]);
            var codigo = LinhasDeCodigo(arquivo).ToList();
            var unica = Assert.Single(codigo);
            Assert.Contains("var x = 1;", unica.linha);
            Assert.Equal(9, unica.numero);                       // o numero da linha continua o do ARQUIVO
            // Controle negativo do proprio teste: as linhas puladas REALMENTE tinham token proibido, ou seja
            // o que prova o pulo nao e' a lista estar vazia.
            var puladas = File.ReadLines(arquivo).Count(l => Proibidos.Any(p => l.Contains(p, StringComparison.Ordinal)));
            Assert.Equal(6, puladas);
        }
        finally { try { File.Delete(arquivo); } catch (IOException) { } }
    }
}
