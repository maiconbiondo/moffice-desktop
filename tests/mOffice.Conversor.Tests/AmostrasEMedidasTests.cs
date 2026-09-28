using mOffice.Conversor;
using Xunit.Abstractions;

namespace mOffice.Conversor.Tests;

/// <summary>
/// Task 7 — geracao das amostras "reais sinteticas" e as medidas de x2t em arquivos GRANDES (5/20/50 MB).
///
/// NAO roda na suite normal: gera dezenas de MB e chama o x2t varias vezes (minutos). E' ligado pela variavel
/// de ambiente <c>MOFFICE_MEDIDAS=1</c>, que e' o que `tools/medir.ps1` faz. Um `[Trait]` sozinho nao
/// bastaria — `dotnet test` sem filtro rodaria tudo assim mesmo.
/// </summary>
[Trait("Categoria", "Medida")]
public class AmostrasEMedidasTests
{
    private readonly ITestOutputHelper _saida;
    public AmostrasEMedidasTests(ITestOutputHelper saida) { _saida = saida; }

    public static string PastaAmostras => Path.GetFullPath(Path.Combine(
        MotorDeTeste.Localizar().PastaMotor, "..", "..", "poc", "amostras", "sinteticas"));

    public static string PastaRelatorios => Path.GetFullPath(Path.Combine(
        MotorDeTeste.Localizar().PastaMotor, "..", "..", "poc", "relatorios"));

    private static void PularSeDesligado()
        => Skip.If(Environment.GetEnvironmentVariable("MOFFICE_MEDIDAS") != "1",
                   "medidas pesadas: rode com MOFFICE_MEDIDAS=1 (tools/medir.ps1)");

    [SkippableFact] // gera poc/amostras/sinteticas/{rico.docx,rico.xlsx,rico.pptx} (fora do git)
    public void GerarAmostrasRicas()
    {
        PularSeDesligado();
        Directory.CreateDirectory(PastaAmostras);
        foreach (var caminho in new[]
                 {
                     AmostrasSinteticas.DocxRico(PastaAmostras),
                     AmostrasSinteticas.XlsxRico(PastaAmostras),
                     AmostrasSinteticas.PptxRico(PastaAmostras),
                 })
        {
            var tamanho = new FileInfo(caminho).Length;
            _saida.WriteLine($"AMOSTRA {Path.GetFileName(caminho)} {tamanho} bytes");
            Assert.True(tamanho > 0, caminho);
        }
        // Prova que o que foi gerado abre de verdade (e nao e' um zip invalido que so o x2t descobriria depois).
        Assert.Contains(AmostrasSinteticas.Marca, FabricaDocumentos.TextoDocx(Path.Combine(PastaAmostras, "rico.docx")));
        Assert.Contains(AmostrasSinteticas.Marca, FabricaDocumentos.TextoXlsx(Path.Combine(PastaAmostras, "rico.xlsx")));
        Assert.Contains(AmostrasSinteticas.Marca, FabricaDocumentos.TextoPptx(Path.Combine(PastaAmostras, "rico.pptx")));

        // As tres ODF saem das OOXML pelo x2t — e' assim que o escritorio recebe ODF na pratica, e da' as
        // linhas "sintetico rico" de ODT/ODS/ODP na matriz do protocolo.
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        foreach (var (origem, alvo) in new[]
                 {
                     ("rico.docx", FormatoDocumento.Odt),
                     ("rico.xlsx", FormatoDocumento.Ods),
                     ("rico.pptx", FormatoDocumento.Odp),
                 })
        {
            var destino = Path.Combine(PastaAmostras, "rico." + alvo.Extensao());
            conv.ConverterAsync(Path.Combine(PastaAmostras, origem), alvo, destino).GetAwaiter().GetResult();
            _saida.WriteLine($"AMOSTRA {Path.GetFileName(destino)} {new FileInfo(destino).Length} bytes");
            Assert.Contains(AmostrasSinteticas.Marca, FabricaDocumentos.TextoOdf(destino));
        }
    }

    [SkippableTheory] // x2t abrir (ParaBin) + salvar (DeBin) nos tamanhos dos limiares; grava medidas-x2t-grandes.txt
    [InlineData("docx", 5)]
    [InlineData("docx", 20)]
    [InlineData("docx", 50)]
    [InlineData("docx-texto", 5)]
    [InlineData("docx-texto", 20)]
    [InlineData("docx-texto", 50)]
    [InlineData("xlsx", 5)]
    [InlineData("xlsx", 20)]
    [InlineData("xlsx", 50)]
    public async Task Medida_X2tArquivoGrande(string tipo, int alvoMb)
    {
        PularSeDesligado();
        Directory.CreateDirectory(PastaAmostras);
        var arquivo = tipo switch
        {
            "docx" => AmostrasSinteticas.DocxGrande(PastaAmostras, alvoMb),
            "docx-texto" => AmostrasSinteticas.DocxGrandeTexto(PastaAmostras, alvoMb),
            _ => AmostrasSinteticas.XlsxGrande(PastaAmostras, alvoMb),
        };
        var mbReal = new FileInfo(arquivo).Length / 1024.0 / 1024.0;
        var formato = FormatoDocumentoExtensoes.DaExtensao(arquivo);
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        var recovery = Path.Combine(Path.GetTempPath(), "mOffice", "medida", Guid.NewGuid().ToString("N"));
        try
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            await conv.ParaBinAsync(arquivo, recovery);
            var abrir = sw.Elapsed;
            sw.Restart();
            await conv.DeBinAsync(recovery, formato, Path.Combine(recovery, "volta." + formato.Extensao()));
            var salvar = sw.Elapsed;

            var linha = $"{DateTime.Now:s} {Path.GetFileName(arquivo)} {mbReal:0.0}MB abrir={abrir.TotalMilliseconds:0}ms salvar={salvar.TotalMilliseconds:0}ms";
            _saida.WriteLine(linha);
            Directory.CreateDirectory(PastaRelatorios);
            await File.AppendAllTextAsync(Path.Combine(PastaRelatorios, "medidas-x2t-grandes.txt"), linha + Environment.NewLine);

            // Medida, nao limiar: o veredito sobre os limiares e' do relatorio (medidas-<data>.md). Aqui so
            // garantimos que a conversao REALMENTE aconteceu — um x2t que falhasse rapido daria "tempo otimo".
            Assert.True(File.Exists(Path.Combine(recovery, "volta." + formato.Extensao())), "o x2t nao gravou o arquivo de volta");
        }
        finally { try { Directory.Delete(recovery, true); } catch { } }
    }
}
