using System.IO.Compression;
using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

public class ValidadorDePacoteTests : IDisposable
{
    private readonly List<string> _pastas = [];

    private string Pasta()
    {
        var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(p);
        _pastas.Add(p);
        return p;
    }

    public void Dispose() { foreach (var p in _pastas) try { Directory.Delete(p, true); } catch { } }

    private static string Zip(string pasta, string nome, params (string Parte, string Conteudo)[] partes)
    {
        var arquivo = Path.Combine(pasta, nome);
        using var z = ZipFile.Open(arquivo, ZipArchiveMode.Create);
        foreach (var (parte, conteudo) in partes)
        {
            using var w = new StreamWriter(z.CreateEntry(parte).Open());
            w.Write(conteudo);
        }
        return arquivo;
    }

    [Fact] // docx do FabricaDocumentos (pacote completo) passa
    public void DocxCompleto_Passa() => ValidadorDePacote.Validar(FabricaDocumentos.Docx(Pasta(), "ok"), FormatoDocumento.Docx);

    [Fact] // zip sem a parte principal (o caso "10 de 12 partes" medido) falha com mensagem pt-BR
    public void DocxSemParteprincipal_Falha()
    {
        var a = Zip(Pasta(), "a.docx", ("[Content_Types].xml", "<Types/>"));
        var e = Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(a, FormatoDocumento.Docx));
        Assert.Contains("incompleto", e.Message);
    }

    [Fact] // zip vazio de 22 bytes (o outro caso medido) falha
    public void ZipVazio_Falha()
    {
        var a = Zip(Pasta(), "a.xlsx");
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(a, FormatoDocumento.Xlsx));
    }

    [Fact] // ODF exige mimetype e content.xml
    public void OdtSemContent_Falha()
    {
        var a = Zip(Pasta(), "a.odt", ("mimetype", "application/vnd.oasis.opendocument.text"));
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(a, FormatoDocumento.Odt));
    }

    [Fact] // PDF sem %%EOF no fim (truncado) falha; com cabecalho e fim, passa
    public void Pdf_ExigeCabecalhoEFim()
    {
        var p = Pasta();
        var ruim = Path.Combine(p, "a.pdf"); File.WriteAllText(ruim, "%PDF-1.7\n1 0 obj");
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(ruim, FormatoDocumento.Pdf));
        var bom = Path.Combine(p, "b.pdf"); File.WriteAllText(bom, "%PDF-1.7\n...\n%%EOF\n");
        ValidadorDePacote.Validar(bom, FormatoDocumento.PdfA);
    }

    [Fact] // documento protegido por senha sai como conteiner OLE (d0cf11e0, medido no Plano 3): aceito so' com cifrado=true
    public void ConteinerOle_SoComCifrado()
    {
        var a = Path.Combine(Pasta(), "a.docx");
        File.WriteAllBytes(a, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0]);
        ValidadorDePacote.Validar(a, FormatoDocumento.Docx, cifrado: true);
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(a, FormatoDocumento.Docx));
    }

    [Fact] // arquivo vazio falha em qualquer formato
    public void Vazio_Falha()
    {
        var a = Path.Combine(Pasta(), "a.pptx"); File.WriteAllBytes(a, []);
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(a, FormatoDocumento.Pptx));
    }

    // ---- Plano 6, Task 2: os formatos novos que o usuario passa a poder gravar ----

    [Theory] // OOXML de modelo, macro e apresentacao de slides: pela parte principal do TIPO (word/cell/slide), nao do formato
    [InlineData(FormatoDocumento.Docm, "word/document.xml")] [InlineData(FormatoDocumento.Dotx, "word/document.xml")]
    [InlineData(FormatoDocumento.Dotm, "word/document.xml")] [InlineData(FormatoDocumento.Xlsm, "xl/workbook.xml")]
    [InlineData(FormatoDocumento.Xltx, "xl/workbook.xml")] [InlineData(FormatoDocumento.Xltm, "xl/workbook.xml")]
    [InlineData(FormatoDocumento.Pptm, "ppt/presentation.xml")] [InlineData(FormatoDocumento.Potx, "ppt/presentation.xml")]
    [InlineData(FormatoDocumento.Ppsx, "ppt/presentation.xml")] [InlineData(FormatoDocumento.Ppsm, "ppt/presentation.xml")]
    [InlineData(FormatoDocumento.Potm, "ppt/presentation.xml")]
    public void OoxmlNovo_ExigeAParteDoTipo(FormatoDocumento f, string parte)
    {
        var p = Pasta();
        ValidadorDePacote.Validar(Zip(p, "ok", ("[Content_Types].xml", "<Types/>"), (parte, "<x/>")), f);
        var e = Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(Zip(p, "ruim", ("[Content_Types].xml", "<Types/>")), f));
        Assert.Contains("incompleto", e.Message);
    }

    [Theory] // ODF de modelo: mimetype + content.xml, como o ODF comum
    [InlineData(FormatoDocumento.Ott)] [InlineData(FormatoDocumento.Ots)] [InlineData(FormatoDocumento.Otp)]
    public void OdfDeModelo_ExigeMimetypeEContent(FormatoDocumento f)
    {
        var p = Pasta();
        ValidadorDePacote.Validar(Zip(p, "ok", ("mimetype", "application/x"), ("content.xml", "<x/>")), f);
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(Zip(p, "ruim", ("mimetype", "application/x")), f));
    }

    [Fact] // RTF: tem de comecar com {\rtf; um zip ou lixo no lugar e' recusado
    public void Rtf_ExigeAbertura()
    {
        var p = Pasta();
        var bom = Path.Combine(p, "a.rtf"); File.WriteAllText(bom, @"{\rtf1\ansi\deff0 oi\par}");
        ValidadorDePacote.Validar(bom, FormatoDocumento.Rtf);
        var ruim = Path.Combine(p, "b.rtf"); File.WriteAllText(ruim, "PK nao e rtf");
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(ruim, FormatoDocumento.Rtf));
        var zip = Zip(p, "c.rtf", ("word/document.xml", "<x/>"));
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(zip, FormatoDocumento.Rtf));
    }

    [Fact] // HTML: nao vazio e com uma marca de HTML no comeco; um zip (o x2t gravando o formato errado) e' recusado
    public void Html_ExigeMarcaDeHtml()
    {
        var p = Pasta();
        var bom = Path.Combine(p, "a.html"); File.WriteAllText(bom, "<!DOCTYPE html><html><body>oi</body></html>");
        ValidadorDePacote.Validar(bom, FormatoDocumento.Html);
        var semDoctype = Path.Combine(p, "b.html"); File.WriteAllText(semDoctype, "  \n<HTML><BODY>oi</BODY></HTML>");
        ValidadorDePacote.Validar(semDoctype, FormatoDocumento.Html);
        var ruim = Path.Combine(p, "c.html"); File.WriteAllText(ruim, "texto qualquer sem marca");
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(ruim, FormatoDocumento.Html));
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(Zip(p, "d.html", ("a", "b")), FormatoDocumento.Html));
    }

    [Theory] // TXT e CSV (Plano 6, Task 3): texto de verdade — nao vazio, sem NUL (binario/UTF-16 por engano) e nao um zip (o x2t gravando o formato errado)
    [InlineData(FormatoDocumento.Txt)] [InlineData(FormatoDocumento.Csv)]
    public void TxtECsv_ExigemTextoPuro(FormatoDocumento f)
    {
        var p = Pasta();
        var bom = Path.Combine(p, "a.txt"); File.WriteAllBytes(bom, [0xEF, 0xBB, 0xBF, .. System.Text.Encoding.UTF8.GetBytes("Nome;Valor\nJoão;1,5\n")]);
        ValidadorDePacote.Validar(bom, f);
        var ansi = Path.Combine(p, "b.txt"); File.WriteAllBytes(ansi, System.Text.Encoding.Latin1.GetBytes("Ação;São Paulo\n"));
        ValidadorDePacote.Validar(ansi, f);
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(Zip(p, "c.txt", ("a", "b")), f));
        var comNul = Path.Combine(p, "d.txt"); File.WriteAllBytes(comNul, [(byte)65, 0, (byte)66, 0]);
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(comNul, f));
        var vazio = Path.Combine(p, "e.txt"); File.WriteAllBytes(vazio, []);
        Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(vazio, f));
    }

    [Theory] // formato so' de leitura (.doc/.xls/.ppt/.xlsb): o x2t nunca os grava, entao validar um deles e' bug de quem chamou — falha FECHADA
    [InlineData(FormatoDocumento.Doc)] [InlineData(FormatoDocumento.Xls)] [InlineData(FormatoDocumento.Ppt)] [InlineData(FormatoDocumento.Xlsb)]
    public void FormatoSoDeLeitura_Recusado(FormatoDocumento f)
    {
        var a = Path.Combine(Pasta(), "a.bin"); File.WriteAllBytes(a, [1, 2, 3, 4]);
        var e = Assert.Throws<ErroConversao>(() => ValidadorDePacote.Validar(a, f));
        Assert.Contains("não pode ser gravado", e.Message);
    }

    [Theory] // a familia vem da tabela
    [InlineData(FormatoDocumento.Docm, "ooxml")] [InlineData(FormatoDocumento.Ott, "odf")] [InlineData(FormatoDocumento.Rtf, "rtf")]
    [InlineData(FormatoDocumento.Html, "texto")] [InlineData(FormatoDocumento.Txt, "texto")] [InlineData(FormatoDocumento.Csv, "texto")] [InlineData(FormatoDocumento.PdfA, "pdf")] [InlineData(FormatoDocumento.Doc, "binario")]
    public void Familia_VemDaTabela(FormatoDocumento f, string familia) => Assert.Equal(familia, ValidadorDePacote.Familia(f));
}
