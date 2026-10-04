using System.IO.Compression;
using System.Text.RegularExpressions;
using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 9: os tres modelos vazios proprios do mOffice (gerados por tools/modelos/gerar-modelos.py).</summary>
public class ModelosVaziosTests
{
    private static string Parte(string pacote, string parte) { using var z = ZipFile.OpenRead(pacote); using var r = new StreamReader(z.GetEntry(parte)!.Open()); return r.ReadToEnd(); }
    private static string Proprio(string nome) => Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo", "Modelos", "pt-BR", nome);

    [Fact] // decisao do dono: Calibri 11, 1,08 (line 259), 8 pt depois (after 160); A4 e margens 2,5/3 cm continuam
    public void Docx_PadraoWord2016()
    {
        var a = Proprio("new.docx");
        var estilos = Parte(a, "word/styles.xml"); var tema = Parte(a, "word/theme/theme1.xml"); var doc = Parte(a, "word/document.xml");
        Assert.Contains("<w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/>", estilos);
        Assert.DoesNotContain("w:line=\"276\"", estilos);
        Assert.Contains("<w:sz w:val=\"22\"/>", estilos);
        Assert.Matches("<a:minorFont><a:latin typeface=\"Calibri\"", tema);
        Assert.Matches("<a:majorFont><a:latin typeface=\"Calibri Light\"", tema);
        Assert.DoesNotContain("<a:latin typeface=\"Arial\"", tema);
        Assert.Contains("w:w=\"11906\"", doc); Assert.Contains("w:left=\"1701\"", doc); Assert.Contains("w:top=\"1417\"", doc);
        Assert.Contains("w:val=\"pt-BR\"", estilos);
    }

    [Fact] // os 11 layouts com os nomes do PowerPoint pt-BR (pelo atributo type, nao pela ordem); tema Calibri
    public void Pptx_LayoutsPtBr()
    {
        var a = Proprio("new.pptx");
        using var z = ZipFile.OpenRead(a);
        var nomes = z.Entries.Where(e => Regex.IsMatch(e.FullName, @"^ppt/slideLayouts/slideLayout\d+\.xml$"))
            .Select(e => { using var r = new StreamReader(e.Open()); var x = r.ReadToEnd();
                return (Regex.Match(x, "<p:sldLayout[^>]* type=\"([^\"]+)\"").Groups[1].Value, Regex.Match(x, "<p:cSld name=\"([^\"]+)\"").Groups[1].Value); })
            .ToDictionary(t => t.Item1, t => t.Item2);
        Assert.Equal(new Dictionary<string, string> {
            ["title"] = "Slide de Título", ["obj"] = "Título e Conteúdo", ["secHead"] = "Cabeçalho da Seção", ["twoObj"] = "Duas Partes de Conteúdo",
            ["twoTxTwoObj"] = "Comparação", ["titleOnly"] = "Somente Título", ["blank"] = "Em branco", ["objTx"] = "Conteúdo com Legenda",
            ["picTx"] = "Imagem com Legenda", ["vertTx"] = "Título e Texto Vertical", ["vertTitleAndTx"] = "Título Vertical e Texto" }, nomes);
        Assert.Contains("<vt:lpstr>Calibri</vt:lpstr>", Parte(a, "docProps/app.xml")); // "Fonts Used"
        Assert.DoesNotContain("<vt:lpstr>Arial</vt:lpstr>", Parte(a, "docProps/app.xml"));
        foreach (var t in z.Entries.Where(e => Regex.IsMatch(e.FullName, @"^ppt/theme/theme\d+\.xml$")))
        {
            using var r = new StreamReader(t.Open()); var x = r.ReadToEnd();
            Assert.Matches("<a:minorFont><a:latin typeface=\"Calibri\"", x);
            Assert.Matches("<a:majorFont><a:latin typeface=\"Calibri Light\"", x);
        }
    }

    [Fact] // a primeira aba do xlsx novo se chama "Planilha1" (Excel pt-BR), nao "Folha1"; sem o caminho de rede do autor do modelo do motor
    public void Xlsx_PrimeiraAbaPlanilha1()
    {
        var wb = Parte(Proprio("new.xlsx"), "xl/workbook.xml");
        Assert.Contains("name=\"Planilha1\"", wb);
        Assert.DoesNotContain("absPath", wb);
        Assert.DoesNotContain("Folha1", wb);
    }

    [SkippableFact] // so' as trocas esperadas: o resto do pacote tem as mesmas partes do motor, na mesma ordem
    public void Gerados_TemAsMesmasPartesDoMotor()
    {
        var motor = Path.Combine(FronteiraTests.RaizRepo(), "third_party", "onlyoffice", "converter", "empty", "pt-BR");
        Skip.If(!Directory.Exists(motor), "motor nao obtido");
        foreach (var n in new[] { "new.docx", "new.xlsx", "new.pptx" })
        {
            using var a = ZipFile.OpenRead(Path.Combine(motor, n)); using var b = ZipFile.OpenRead(Proprio(n));
            Assert.Equal(a.Entries.Select(e => e.FullName), b.Entries.Select(e => e.FullName));
        }
    }

    [SkippableFact] // ModelosVazios usa os modelos proprios antes dos do motor
    public void Caminho_PreferePropriosDoMOffice()
    {
        var motor = mOffice.Conversor.Tests.MotorDeTeste.Localizar();
        foreach (var f in new[] { FormatoDocumento.Docx, FormatoDocumento.Xlsx, FormatoDocumento.Pptx, FormatoDocumento.Odt })
            Assert.StartsWith(Path.Combine(AppContext.BaseDirectory, "Modelos"), ModelosVazios.Caminho(motor, f));
    }
}
