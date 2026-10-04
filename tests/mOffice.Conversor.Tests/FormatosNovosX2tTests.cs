using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

/// <summary>
/// Plano 6, Task 2 — cada formato NOVO que o usuario passa a poder gravar (modelos, macros, apresentacao de slides, ODF de
/// modelo, RTF, HTML) tem de: sair do x2t, passar no <see cref="ValidadorDePacote"/> (senao o salvar o recusaria) e voltar a abrir
/// com o texto. Roda o x2t de verdade (so' o motor; nenhuma janela, entao vale tambem em sessao remota).
/// </summary>
public class FormatosNovosX2tTests
{
    private const string Marca = "marca-formatos-novos";

    private static async Task<string> TextoReaberto(X2tConversor conv, string pasta, string arquivo, FormatoDocumento moderno)
    {
        var rec = Path.Combine(pasta, "reabrir-" + Guid.NewGuid().ToString("N")[..6]);
        await conv.ParaBinAsync(arquivo, rec);
        var volta = Path.Combine(pasta, "volta-" + Guid.NewGuid().ToString("N")[..6] + "." + moderno.Extensao());
        await conv.DeBinAsync(rec, moderno, volta);
        return moderno.DocType() switch
        {
            "cell" => FabricaDocumentos.TextoXlsx(volta),
            "slide" => FabricaDocumentos.TextoPptx(volta),
            _ => FabricaDocumentos.TextoDocx(volta),
        };
    }

    /// <summary>DOCX sintetico COM <c>sectPr</c>: o gravador de RTF do x2t descarta o corpo de um documento sem configuracao de
    /// pagina (MEDIDO: 905 bytes, sem texto) — todo documento real tem uma. Os outros formatos nao dependem disso.</summary>
    private static string DocxComSecao(string pasta)
    {
        var f = FabricaDocumentos.Docx(pasta, Marca);
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(f, true);
        doc.MainDocumentPart!.Document.Body!.Append(new DocumentFormat.OpenXml.Wordprocessing.SectionProperties(
            new DocumentFormat.OpenXml.Wordprocessing.PageSize { Width = 11906, Height = 16838 }));
        doc.MainDocumentPart.Document.Save();
        return f;
    }

    [SkippableTheory] // grava, valida e reabre com o texto
    [InlineData(FormatoDocumento.Docm)] [InlineData(FormatoDocumento.Dotx)] [InlineData(FormatoDocumento.Dotm)]
    [InlineData(FormatoDocumento.Ott)] [InlineData(FormatoDocumento.Rtf)]
    [InlineData(FormatoDocumento.Xlsm)] [InlineData(FormatoDocumento.Xltx)] [InlineData(FormatoDocumento.Xltm)] [InlineData(FormatoDocumento.Ots)]
    [InlineData(FormatoDocumento.Pptm)] [InlineData(FormatoDocumento.Potx)] [InlineData(FormatoDocumento.Potm)]
    [InlineData(FormatoDocumento.Ppsx)] [InlineData(FormatoDocumento.Ppsm)] [InlineData(FormatoDocumento.Otp)]
    public async Task NovoFormato_GravaValidaEReabre(FormatoDocumento alvo)
    {
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        var pasta = Directory.CreateTempSubdirectory("moffice-fmt-novos-").FullName;
        try
        {
            var tipo = alvo.DocType();
            var origem = tipo switch
            {
                "cell" => FabricaDocumentos.Xlsx(pasta, Marca),
                "slide" => FabricaDocumentos.Pptx(pasta, Marca),
                _ => DocxComSecao(pasta),
            };
            var rec = Path.Combine(pasta, "rec");
            await conv.ParaBinAsync(origem, rec);
            var destino = Path.Combine(pasta, "saida" + TabelaFormatos.Principal(alvo).Extensao);
            await conv.DeBinAsync(rec, alvo, destino);
            ValidadorDePacote.Validar(destino, alvo);   // o salvar so' substitui o arquivo do usuario se isto passar
            Assert.Contains(Marca, await TextoReaberto(conv, pasta, destino, TabelaFormatos.Moderno(tipo)));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    private static readonly byte[] Vba = [.. "VBA-FALSO-MOFFICE-PLANO6"u8.ToArray(), .. Enumerable.Range(0, 256).Select(i => (byte)i)];

    /// <summary>Pacote com macro: o sintetico do tipo, convertido para MacroEnabled e com um <c>vbaProject.bin</c> marcador (nao e' um
    /// projeto VBA valido — so' prova que o x2t preserva os bytes).</summary>
    private static string ComVba(string pasta, string tipo)
    {
        switch (tipo)
        {
            case "cell":
            {
                var f = Path.Combine(pasta, "macro.xlsm"); File.Move(FabricaDocumentos.Xlsx(pasta, Marca), f);
                using var d = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(f, true);
                d.ChangeDocumentType(DocumentFormat.OpenXml.SpreadsheetDocumentType.MacroEnabledWorkbook);
                using var ms = new MemoryStream(Vba); d.WorkbookPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.VbaProjectPart>().FeedData(ms);
                return f;
            }
            case "slide":
            {
                var f = Path.Combine(pasta, "macro.pptm"); File.Move(FabricaDocumentos.Pptx(pasta, Marca), f);
                using var d = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(f, true);
                d.ChangeDocumentType(DocumentFormat.OpenXml.PresentationDocumentType.MacroEnabledPresentation);
                using var ms = new MemoryStream(Vba); d.PresentationPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.VbaProjectPart>().FeedData(ms);
                return f;
            }
            default:
            {
                var f = Path.Combine(pasta, "macro.docm"); File.Move(DocxComSecao(pasta), f);
                using var d = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(f, true);
                d.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.MacroEnabledDocument);
                using var ms = new MemoryStream(Vba); d.MainDocumentPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.VbaProjectPart>().FeedData(ms);
                return f;
            }
        }
    }

    private static byte[]? VbaDoPacote(string pacote)
    {
        using var z = System.IO.Compression.ZipFile.OpenRead(pacote);
        var e = z.Entries.FirstOrDefault(x => x.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
        if (e is null) return null;
        using var s = e.Open(); using var ms = new MemoryStream(); s.CopyTo(ms);
        return ms.ToArray();
    }

    [SkippableTheory] // macro: docm/xlsm/pptm abertos e salvos de volta no MESMO formato mantem o vbaProject.bin byte a byte (e o tipo de conteudo de macro)
    [InlineData(FormatoDocumento.Docm)] [InlineData(FormatoDocumento.Xlsm)] [InlineData(FormatoDocumento.Pptm)]
    public async Task Macro_MesmoFormato_PreservaVbaProjectByteAByte(FormatoDocumento fmt)
    {
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        var pasta = Directory.CreateTempSubdirectory("moffice-fmt-vba-").FullName;
        try
        {
            var origem = ComVba(pasta, fmt.DocType());
            Assert.Equal(Vba, VbaDoPacote(origem));   // linha de base: a fixture tem o VBA
            var rec = Path.Combine(pasta, "rec");
            await conv.ParaBinAsync(origem, rec);
            var destino = Path.Combine(pasta, "volta" + TabelaFormatos.Principal(fmt).Extensao);
            await conv.DeBinAsync(rec, fmt, destino);
            ValidadorDePacote.Validar(destino, fmt);
            Assert.Equal(Vba, VbaDoPacote(destino));
            using var z = System.IO.Compression.ZipFile.OpenRead(destino);
            using var r = new StreamReader(z.GetEntry("[Content_Types].xml")!.Open());
            Assert.Contains("macroEnabled", r.ReadToEnd());
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // HTML com imagem (medido): o x2t embute a figura como data URI e NAO cria arquivos/pasta de recursos ao lado do destino
    public async Task Html_ComImagem_EmbuteESemArquivosAoLado()
    {
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        var pasta = Directory.CreateTempSubdirectory("moffice-fmt-html-img-").FullName;
        try
        {
            var rec = Path.Combine(pasta, "rec");
            await conv.ParaBinAsync(FabricaDocumentos.DocxComImagem(pasta, Marca), rec);
            var saida = Directory.CreateDirectory(Path.Combine(pasta, "saida")).FullName;
            var destino = Path.Combine(saida, "pagina.html.moffice-tmp");   // o nome temporario que o salvar usa
            await conv.DeBinAsync(rec, FormatoDocumento.Html, destino);
            Assert.Equal([destino], Directory.EnumerateFileSystemEntries(saida).ToArray());
            var html = File.ReadAllText(destino);
            Assert.Contains("data:image/png;base64,", html);
            Assert.Contains(Marca, html);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // HTML ("Pagina da Web"): so' grava (nao reabre); o x2t gera HTML de verdade, com o texto, e o validador aceita
    public async Task Html_GravaValidaEContemOTexto()
    {
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        var pasta = Directory.CreateTempSubdirectory("moffice-fmt-html-").FullName;
        try
        {
            var rec = Path.Combine(pasta, "rec");
            await conv.ParaBinAsync(DocxComSecao(pasta), rec);
            var destino = Path.Combine(pasta, "pagina.html");
            await conv.DeBinAsync(rec, FormatoDocumento.Html, destino);
            ValidadorDePacote.Validar(destino, FormatoDocumento.Html);
            Assert.Contains(Marca, File.ReadAllText(destino));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
