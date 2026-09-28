using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

public class X2tConversorTests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("moffice-x2t-").FullName;
    private readonly X2tConversor _conv = new(MotorDeTeste.Localizar());
    public const string Frase = "mOffice PoC round-trip 2026";

    [Fact] // DOCX -> Editor.bin -> DOCX preserva o texto (abrir + salvar sem changes)
    public async Task Docx_RoundTrip_PreservaTexto()
    {
        var origem = FabricaDocumentos.Docx(_pasta, Frase);
        var recovery = Path.Combine(_pasta, "rec");
        var bin = await _conv.ParaBinAsync(origem, recovery);
        Assert.True(File.Exists(bin) && new FileInfo(bin).Length > 0, "Editor.bin vazio");
        var destino = Path.Combine(_pasta, "volta.docx");
        await _conv.DeBinAsync(recovery, FormatoDocumento.Docx, destino);
        Assert.Contains(Frase, FabricaDocumentos.TextoDocx(destino));
    }

    [Fact] // XLSX -> bin -> XLSX preserva A1
    public async Task Xlsx_RoundTrip_PreservaA1()
    {
        var origem = FabricaDocumentos.Xlsx(_pasta, Frase);
        var recovery = Path.Combine(_pasta, "rec");
        await _conv.ParaBinAsync(origem, recovery);
        var destino = Path.Combine(_pasta, "volta.xlsx");
        await _conv.DeBinAsync(recovery, FormatoDocumento.Xlsx, destino);
        Assert.Equal(Frase, FabricaDocumentos.TextoXlsxA1(destino));
    }

    [Fact] // PPTX -> bin -> PPTX preserva o texto do slide
    public async Task Pptx_RoundTrip_PreservaTexto()
    {
        var origem = FabricaDocumentos.Pptx(_pasta, Frase);
        var recovery = Path.Combine(_pasta, "rec");
        await _conv.ParaBinAsync(origem, recovery);
        var destino = Path.Combine(_pasta, "volta.pptx");
        await _conv.DeBinAsync(recovery, FormatoDocumento.Pptx, destino);
        Assert.Contains(Frase, FabricaDocumentos.TextoPptx(destino));
    }

    [Theory] // OOXML -> ODF direto (arquivo -> arquivo): o texto sobrevive em content.xml
    [InlineData("docx", FormatoDocumento.Odt)]
    [InlineData("xlsx", FormatoDocumento.Ods)]
    [InlineData("pptx", FormatoDocumento.Odp)]
    public async Task Ooxml_ParaOdf_PreservaTexto(string tipo, FormatoDocumento alvo)
    {
        var origem = tipo switch { "docx" => FabricaDocumentos.Docx(_pasta, Frase), "xlsx" => FabricaDocumentos.Xlsx(_pasta, Frase), _ => FabricaDocumentos.Pptx(_pasta, Frase) };
        var destino = Path.Combine(_pasta, "saida." + alvo.Extensao());
        await _conv.ConverterAsync(origem, alvo, destino);
        Assert.Contains(Frase, FabricaDocumentos.TextoOdf(destino));
    }

    [Theory] // ODF -> bin -> ODF (o caminho que o editor usara para ODT/ODS/ODP)
    [InlineData(FormatoDocumento.Odt)]
    [InlineData(FormatoDocumento.Ods)]
    [InlineData(FormatoDocumento.Odp)]
    public async Task Odf_RoundTripPeloBin_PreservaTexto(FormatoDocumento fmt)
    {
        var ooxml = fmt switch { FormatoDocumento.Odt => FabricaDocumentos.Docx(_pasta, Frase), FormatoDocumento.Ods => FabricaDocumentos.Xlsx(_pasta, Frase), _ => FabricaDocumentos.Pptx(_pasta, Frase) };
        var odf = Path.Combine(_pasta, "origem." + fmt.Extensao());
        await _conv.ConverterAsync(ooxml, fmt, odf);
        var recovery = Path.Combine(_pasta, "rec");
        await _conv.ParaBinAsync(odf, recovery);
        var destino = Path.Combine(_pasta, "volta." + fmt.Extensao());
        await _conv.DeBinAsync(recovery, fmt, destino);
        Assert.Contains(Frase, FabricaDocumentos.TextoOdf(destino));
    }

    [Fact] // PDF e PDF/A sao EXPORTACAO (saida so' de leitura); os formatos de documento, nao
    public void EhExportacao_SoPdf()
    {
        Assert.True(FormatoDocumento.Pdf.EhExportacao());
        Assert.True(FormatoDocumento.PdfA.EhExportacao());
        foreach (var f in new[] { FormatoDocumento.Docx, FormatoDocumento.Odt, FormatoDocumento.Xlsx,
                                  FormatoDocumento.Ods, FormatoDocumento.Pptx, FormatoDocumento.Odp })
            Assert.False(f.EhExportacao(), f + " nao e' exportacao: o editor reabre este formato");
    }

    [Fact] // DOCX -> PDF gera um PDF (assinatura %PDF-)
    public async Task Docx_ParaPdf_GeraPdf()
    {
        var origem = FabricaDocumentos.Docx(_pasta, Frase);
        var destino = Path.Combine(_pasta, "saida.pdf");
        await _conv.ConverterAsync(origem, FormatoDocumento.Pdf, destino);
        var cabecalho = new byte[5]; using (var fs = File.OpenRead(destino)) fs.ReadExactly(cabecalho);
        Assert.Equal("%PDF-", System.Text.Encoding.ASCII.GetString(cabecalho));
    }

    [Fact] // DOCX -> PDF/A arquivo a arquivo (Plano 4, fix round 1 da Task 2): mesmo pedido do DeBinAsync (513 + m_bIsPDFA),
           // o arquivo sai e DECLARA PDF/A no XMP (pdfaid:part); o PDF comum da mesma origem nao declara (controle negativo)
    public async Task Docx_ParaPdfA_DeclaraPdfA()
    {
        var origem = FabricaDocumentos.Docx(_pasta, Frase);
        var pdfa = Path.Combine(_pasta, "saida-a.pdf");
        await _conv.ConverterAsync(origem, FormatoDocumento.PdfA, pdfa);
        var texto = System.Text.Encoding.Latin1.GetString(await File.ReadAllBytesAsync(pdfa));
        Assert.StartsWith("%PDF-", texto);
        Assert.Contains("pdfaid:part", texto);
        var pdf = Path.Combine(_pasta, "saida.pdf");
        await _conv.ConverterAsync(origem, FormatoDocumento.Pdf, pdf);
        Assert.DoesNotContain("pdfaid:part", System.Text.Encoding.Latin1.GetString(await File.ReadAllBytesAsync(pdf)));
    }

    [Fact] // arquivo corrompido -> ErroConversao com mensagem pt-BR, nunca excecao crua do Process
    public async Task ArquivoCorrompido_LancaErroConversao()
    {
        var lixo = Path.Combine(_pasta, "lixo.docx"); await File.WriteAllBytesAsync(lixo, new byte[] { 1, 2, 3 });
        var ex = await Assert.ThrowsAsync<ErroConversao>(() => _conv.ParaBinAsync(lixo, Path.Combine(_pasta, "rec")));
        Assert.Contains("não foi possível converter", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact] // ElementosComuns cria m_sTempDir por chamada (finding do review): a pasta-pai de PastaTemp nao pode acumular subpastas orfas, no sucesso nem no erro
    public async Task Conversao_NaoDeixaPastaTemporaria()
    {
        var pastaTempPai = Path.Combine(_pasta, "tmp-x2t");
        var conv = new X2tConversor(MotorDeTeste.Localizar() with { PastaTemp = pastaTempPai });

        var origem = FabricaDocumentos.Docx(_pasta, Frase);
        var destino = Path.Combine(_pasta, "saida-temp.pdf");
        await conv.ConverterAsync(origem, FormatoDocumento.Pdf, destino);
        Assert.True(File.Exists(destino));
        Assert.Empty(Directory.Exists(pastaTempPai) ? Directory.GetDirectories(pastaTempPai) : Array.Empty<string>());

        var lixo = Path.Combine(_pasta, "lixo-temp.docx");
        await File.WriteAllBytesAsync(lixo, new byte[] { 1, 2, 3 });
        await Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(lixo, Path.Combine(_pasta, "rec-temp")));
        Assert.Empty(Directory.Exists(pastaTempPai) ? Directory.GetDirectories(pastaTempPai) : Array.Empty<string>());
    }

    [Fact] // x2t ausente -> ErroConversao claro (nao FileNotFound do Process)
    public async Task X2tAusente_LancaErroConversao()
    {
        var conv = new X2tConversor(new ConfiguracaoMotor(Path.Combine(_pasta, "motor-inexistente")));
        var ex = await Assert.ThrowsAsync<ErroConversao>(() => conv.ConverterAsync(FabricaDocumentos.Docx(_pasta, Frase), FormatoDocumento.Pdf, Path.Combine(_pasta, "x.pdf")));
        Assert.Contains("x2t.exe", ex.Message);
    }

    [Fact] // medida (nao limiar): tempos de ParaBin/DeBin do DOCX sintetico em poc/relatorios/medidas-x2t.txt
    public async Task Medida_TemposSintetico()
    {
        var origem = FabricaDocumentos.Docx(_pasta, Frase);
        var recovery = Path.Combine(_pasta, "rec");
        var sw = System.Diagnostics.Stopwatch.StartNew(); await _conv.ParaBinAsync(origem, recovery); var abrir = sw.Elapsed;
        sw.Restart(); await _conv.DeBinAsync(recovery, FormatoDocumento.Docx, Path.Combine(_pasta, "v.docx")); var salvar = sw.Elapsed;
        var relatorios = Path.GetFullPath(Path.Combine(MotorDeTeste.Localizar().PastaMotor, "..", "..", "poc", "relatorios"));
        Directory.CreateDirectory(relatorios);
        await File.AppendAllTextAsync(Path.Combine(relatorios, "medidas-x2t.txt"), $"{DateTime.Now:s} sintetico.docx abrir={abrir.TotalMilliseconds:0}ms salvar={salvar.TotalMilliseconds:0}ms\n");
    }

    public void Dispose() { try { Directory.Delete(_pasta, true); } catch { } }
}
