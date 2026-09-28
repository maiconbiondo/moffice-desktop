using mOffice.Nucleo.Assinatura;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 4, Task 5 — deteccao do mPDF pela chave de desinstalacao do instalador dele. Testes PUROS: o registro entra
/// pelo leitor falso (<c>lerChave</c>), o exe e' um arquivo vazio numa pasta temporaria. Nenhum teste abre o mPDF.
/// </summary>
public class AssinaturaTests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("moffice-mpdf-").FullName;

    public void Dispose()
    {
        try { Directory.Delete(_pasta, true); } catch { }
    }

    /// <summary>Uma "instalacao" de mentira: pasta com um <c>mPdf.App.exe</c> vazio. Devolve a pasta com a barra
    /// final, como o Inno grava o InstallLocation (<c>C:\Program Files\mPDF\</c>).</summary>
    private string Instalacao(string nome)
    {
        var d = Directory.CreateDirectory(Path.Combine(_pasta, nome)).FullName;
        File.WriteAllBytes(Path.Combine(d, DetectorMpdf.NomeExe), []);
        return d + Path.DirectorySeparatorChar;
    }

    private static Func<string, (string? Local, string? Versao)?> Leitor(
        (string? Local, string? Versao)? hklm, (string? Local, string? Versao)? hcu = null) =>
        origem => origem switch
        {
            "HKLM64" => hklm,
            "HKCU64" => hcu,
            "HKLM32" or "HKCU32" => null,
            _ => throw new ArgumentException(origem),
        };

    /// <summary>Leitor com as 4 origens explicitas (<see cref="DetectorMpdf.Origens"/>).</summary>
    private static Func<string, (string? Local, string? Versao)?> Leitor4(
        Dictionary<string, (string? Local, string? Versao)> porOrigem) =>
        origem => porOrigem.TryGetValue(origem, out var v) ? v : null;

    [Fact]
    public void Detectar_HklmValido_Encontra()
    {
        var local = Instalacao("mPDF");
        var m = DetectorMpdf.Detectar(Leitor((local, "2.13.0")));
        Assert.NotNull(m);
        Assert.Equal(Path.Combine(local, "mPdf.App.exe"), m.Exe);
        Assert.Equal(new Version(2, 13, 0), m.Versao);
    }

    [Fact] // revisao final do Plano 4: vence a MAIOR versao valida entre HKLM e HKCU (antes HKLM ganhava sempre)
    public void Detectar_MaiorVersaoVence()
    {
        var maquina = Instalacao("maquina");
        var usuario = Instalacao("usuario");
        var m = DetectorMpdf.Detectar(Leitor((maquina, "2.13.0"), (usuario, "2.14.0")))!;
        Assert.Equal(Path.Combine(usuario, "mPdf.App.exe"), m.Exe);
        Assert.Equal(new Version(2, 14, 0), m.Versao);
        // Ao contrario: a de maquina e' a mais nova.
        Assert.Equal(Path.Combine(maquina, "mPdf.App.exe"), DetectorMpdf.Detectar(Leitor((maquina, "2.15.0"), (usuario, "2.14.0")))!.Exe);
        // Empate: fica a primeira na ordem (HKLM antes de HKCU).
        Assert.Equal(Path.Combine(maquina, "mPdf.App.exe"), DetectorMpdf.Detectar(Leitor((maquina, "2.14.0"), (usuario, "2.14.0")))!.Exe);
        // So' HKCU (instalacao por usuario): vale.
        Assert.Equal(new Version(2, 14, 0), DetectorMpdf.Detectar(Leitor(null, (usuario, "2.14.0")))!.Versao);
        // Versao maior mas invalida (exe apagado) nao esconde uma menor boa.
        Assert.Equal(Path.Combine(usuario, "mPdf.App.exe"),
            DetectorMpdf.Detectar(Leitor((Path.Combine(_pasta, "sumiu") + "\\", "2.15.0"), (usuario, "2.14.0")))!.Exe);
    }

    [Fact] // as 4 origens sao lidas; uma chave com valor faltando/invalido nao encerra a busca nas outras visoes/hives
    public void Detectar_VisoesEHives_ChaveInvalidaNaoParaABusca()
    {
        var v32 = Instalacao("hklm32");
        var cu32 = Instalacao("hkcu32");
        Assert.Equal(["HKLM64", "HKLM32", "HKCU64", "HKCU32"], DetectorMpdf.Origens);
        // HKLM64 presente sem DisplayVersion, HKLM32 boa.
        var m = DetectorMpdf.Detectar(Leitor4(new() { ["HKLM64"] = (Instalacao("sem-versao"), null), ["HKLM32"] = (v32, "2.13.0") }))!;
        Assert.Equal(Path.Combine(v32, "mPdf.App.exe"), m.Exe);
        // HKLM64 com versao ilegivel e HKLM32 sem InstallLocation: a de usuario (visao 32) vale.
        m = DetectorMpdf.Detectar(Leitor4(new()
        {
            ["HKLM64"] = (Instalacao("ilegivel"), "dois ponto treze"),
            ["HKLM32"] = (null, "2.20.0"),
            ["HKCU32"] = (cu32, "2.13.1"),
        }))!;
        Assert.Equal(Path.Combine(cu32, "mPdf.App.exe"), m.Exe);
        Assert.Equal(new Version(2, 13, 1), m.Versao);
        // Maior entre as 4.
        m = DetectorMpdf.Detectar(Leitor4(new()
        {
            ["HKLM64"] = (Instalacao("a"), "2.13.0"),
            ["HKLM32"] = (Instalacao("b"), "2.13.5"),
            ["HKCU64"] = (Instalacao("c"), "2.13.2"),
            ["HKCU32"] = (Instalacao("d"), "2.14.0"),
        }))!;
        Assert.Equal(new Version(2, 14, 0), m.Versao);
        Assert.EndsWith(Path.Combine("d", "mPdf.App.exe"), m.Exe);
    }

    [Fact]
    public void Detectar_SemChave_Nulo()
    {
        Assert.Null(DetectorMpdf.Detectar(Leitor(null, null)));
        Assert.Null(DetectorMpdf.Detectar(Leitor((null, null), (null, null))));
    }

    [Fact]
    public void Detectar_ExeInexistente_Nulo()
    {
        var vazia = Directory.CreateDirectory(Path.Combine(_pasta, "sem-exe")).FullName;
        Assert.Null(DetectorMpdf.Detectar(Leitor((vazia, "2.13.0"))));
        Assert.Null(DetectorMpdf.Detectar(Leitor((Path.Combine(_pasta, "nao-existe"), "2.13.0"))));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("dois ponto treze")]
    [InlineData("v")]
    [InlineData("vv2.13.0")]
    [InlineData("2")]
    public void Detectar_VersaoIlegivel_Nulo(string versao)
        => Assert.Null(DetectorMpdf.Detectar(Leitor((Instalacao("mPDF"), versao))));

    [Theory]
    [InlineData(@"\\servidor\apps\mPDF\")]
    [InlineData(@"//servidor/apps/mPDF/")]
    [InlineData(@"\\?\C:\Program Files\mPDF\")]
    [InlineData(@"\\.\C:\Program Files\mPDF\")]
    public void Detectar_InstallLocationUnc_Nulo(string local)
        => Assert.Null(DetectorMpdf.Detectar(Leitor((local, "2.13.0"))));

    [Fact]
    public void Detectar_InstallLocationRelativo_Nulo()
    {
        // Relativo ao diretorio de trabalho: mesmo com o exe la', nao vale (o diretorio de trabalho nao e' confiavel). A
        // instalacao de mentira fica DENTRO do diretorio de trabalho (sem troca-lo: e' global do processo e a suite roda
        // classes em paralelo), para o relativo apontar de fato para um exe existente — o controle.
        var rel = "moffice-mpdf-rel-" + Guid.NewGuid().ToString("N");
        var dir = Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory, rel)).FullName;
        try
        {
            File.WriteAllBytes(Path.Combine(dir, DetectorMpdf.NomeExe), []);
            Assert.True(File.Exists(Path.Combine(rel, "mPdf.App.exe")));   // o controle: o exe existe pelo relativo
            Assert.NotNull(DetectorMpdf.Detectar(Leitor((dir, "2.13.0"))));  // ...e pelo absoluto, vale
            Assert.Null(DetectorMpdf.Detectar(Leitor((rel, "2.13.0"))));
            Assert.Null(DetectorMpdf.Detectar(Leitor((rel + "\\", "2.13.0"))));
            Assert.Null(DetectorMpdf.Detectar(Leitor((@".\" + rel, "2.13.0"))));
            Assert.Null(DetectorMpdf.Detectar(Leitor((dir[2..], "2.13.0"))));                 // "\pasta\...": raiz da unidade atual, sem letra
            Assert.Null(DetectorMpdf.Detectar(Leitor((dir[..2] + rel, "2.13.0"))));           // "D:pasta": relativo a' unidade
        }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }

    [Theory]
    [InlineData("2.13.0-rc1", 2, 13, 0)]
    [InlineData("2.13.1 beta", 2, 13, 1)]
    [InlineData(" 2.12.2 ", 2, 12, 2)]
    [InlineData("2.13", 2, 13, 0)]          // sem Build: vira 2.13.0 (Version(2,13) seria MENOR que a minima)
    [InlineData("2.13.0.4", 2, 13, 0)]
    [InlineData("v2.13.0", 2, 13, 0)]       // "v" inicial aceito (revisao final do Plano 4)
    [InlineData("V2.14", 2, 14, 0)]
    public void Detectar_VersaoComSufixo(string texto, int maior, int menor, int build)
        => Assert.Equal(new Version(maior, menor, build), DetectorMpdf.Detectar(Leitor((Instalacao("mPDF"), texto)))!.Versao);

    [Fact]
    public void Minima_2_13_0()
    {
        Assert.Equal(new Version(2, 13, 0), DetectorMpdf.Minima);
        Assert.True(DetectorMpdf.AbreAssinatura(new MpdfInstalado(@"C:\x\mPdf.App.exe", new Version(2, 13, 0))));
        Assert.True(DetectorMpdf.AbreAssinatura(new MpdfInstalado(@"C:\x\mPdf.App.exe", new Version(3, 0, 0))));
        Assert.False(DetectorMpdf.AbreAssinatura(new MpdfInstalado(@"C:\x\mPdf.App.exe", new Version(2, 12, 2))));
    }

    [Fact] // leitor que lanca (acesso negado ao registro): nao derruba, passa para a proxima raiz
    public void Detectar_LeitorQueLanca_PassaAdiante()
    {
        var usuario = Instalacao("usuario");
        var m = DetectorMpdf.Detectar(origem => origem.StartsWith("HKLM", StringComparison.Ordinal) ? throw new UnauthorizedAccessException("negado") : (usuario, "2.13.0"));
        Assert.Equal(Path.Combine(usuario, "mPdf.App.exe"), m!.Exe);
    }

    [Fact] // o pedido de processo: exe direto (sem shell), "/assinar" e o PDF como DOIS argumentos separados — um nome
           // com aspas, espaco e & continua UM argumento (nunca uma linha de comando concatenada)
    public void Lancador_ArgumentList_SemShell()
    {
        var mpdf = new MpdfInstalado(@"C:\Program Files\mPDF\mPdf.App.exe", new Version(2, 13, 0));
        var pdf = @"C:\Peças & Cia\contrato ""final"" -x.pdf";
        var psi = LancadorMpdf.Pedido(mpdf, pdf);
        Assert.Equal(mpdf.Exe, psi.FileName);
        Assert.False(psi.UseShellExecute);
        Assert.Equal(["/assinar", pdf], psi.ArgumentList);
        Assert.Equal("", psi.Arguments);
        Assert.Equal(@"C:\Program Files\mPDF", psi.WorkingDirectory);
    }

    [Fact] // a leitura do registro de verdade: so' leitura, nao lanca. O valor desta maquina nao e' conferido.
    public void DetectarNoWindows_NaoLanca()
    {
        var m = DetectorMpdf.DetectarNoWindows();
        if (m is not null) Assert.True(File.Exists(m.Exe));
    }
}

/// <summary>
/// Plano 4, Task 5, passo 4 — janela real: "Salvar como" do host com a linha "PDF assinado com mPDF"
/// (<c>DestinoSalvar(…, Pdf, Assinar: true)</c> pela costura do dialogo) e um lancador FALSO ligado ao
/// <see cref="JanelaEditor.AssinarPedido"/> como a casca liga (o <c>PedidoDeAssinatura</c> do Comum e' testado em
/// <c>AplicativoOfficeTests</c>; a costura do dialogo e' <c>internal</c> do Nucleo, por isso este teste mora aqui).
/// </summary>
[Collection(JanelaReal.Nome)]
public class AssinaturaJanelaTests
{
    private sealed class LancadorFalso : ILancadorMpdf
    {
        public List<(MpdfInstalado Mpdf, string Pdf)> Chamadas { get; } = [];
        public void Assinar(MpdfInstalado mpdf, string pdf) => Chamadas.Add((mpdf, pdf));
    }

    [SkippableFact]
    public void SalvarComoPdfAssinado_ExportaELancaOMpdfComOPdf_SessaoNoDocx()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-assinar-" + Guid.NewGuid().ToString("N"))).FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Contrato de prestacao");
        var pdf = Path.Combine(pasta, "contrato & anexo 'final'.pdf");
        var mpdf = new MpdfInstalado(Path.Combine(pasta, "nao-executa", DetectorMpdf.NomeExe), new Version(2, 13, 0));
        var lancador = new LancadorFalso();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.AssinarPedido += (_, p) => lancador.Assinar(mpdf, p);
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                j.EscolherDestinoParaTeste = (_, _) => new DestinoSalvar(pdf, mOffice.Conversor.FormatoDocumento.Pdf, Assinar: true);
                var r = await j.SalvarComoComDialogoAsync();
                Assert.True(r.Exportou);
                Assert.Null(j.UltimoErro);
                Assert.Equal(docx, j.SessaoAtual!.ArquivoOrigem);   // assinar nao reponta a sessao: continua no DOCX
                Assert.Equal(mOffice.Conversor.FormatoDocumento.Docx, j.SessaoAtual!.Formato);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Equal([(mpdf, pdf)], lancador.Chamadas);
            Assert.Equal("%PDF-", System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(pdf), 0, 5));
            Assert.True(mOffice.Conversor.Tests.TextoDePdf.Paginas(pdf) >= 1);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
