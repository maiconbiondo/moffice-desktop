using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

public class TabelaFormatosTests
{
    [Fact] // a lista literal de extensoes que ABREM em cada app (cardinalidade conferida no plano, nunca comparada com a propria tabela)
    public void Abre_ListaLiteralPorApp()
    {
        string[] Abre(string t) => TabelaFormatos.DoTipo(t).Where(i => i.Abre).Select(i => i.Extensao).Order(StringComparer.Ordinal).ToArray();
        Assert.Equal([".doc", ".docm", ".docx", ".dot", ".dotm", ".dotx", ".odt", ".ott", ".rtf", ".txt"], Abre("word"));
        Assert.Equal([".csv", ".ods", ".ots", ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltm", ".xltx"], Abre("cell"));
        Assert.Equal([".odp", ".otp", ".pot", ".potm", ".potx", ".pps", ".ppsm", ".ppsx", ".ppt", ".pptm", ".pptx"], Abre("slide"));
    }

    [Theory] // o x2t 9.4 NAO grava .doc/.xls/.ppt (80/88, medido) e o .xlsb relido perde os textos: nunca "Grava"
    [InlineData(".doc")] [InlineData(".dot")] [InlineData(".xls")] [InlineData(".xlt")] [InlineData(".xlsb")]
    [InlineData(".ppt")] [InlineData(".pps")] [InlineData(".pot")]
    public void Legado_NaoGrava(string ext)
    {
        var i = TabelaFormatos.DaExtensao(ext)!;
        Assert.True(i.Legado); Assert.False(i.Grava);
    }

    [Theory] // legado grava como o moderno do mesmo tipo
    [InlineData(".doc", FormatoDocumento.Docx)] [InlineData(".dot", FormatoDocumento.Docx)]
    [InlineData(".xls", FormatoDocumento.Xlsx)] [InlineData(".xlsb", FormatoDocumento.Xlsx)]
    [InlineData(".ppt", FormatoDocumento.Pptx)] [InlineData(".pps", FormatoDocumento.Pptx)]
    [InlineData(".docx", FormatoDocumento.Docx)] [InlineData(".dotm", FormatoDocumento.Dotm)]
    public void GravaComo(string ext, FormatoDocumento esperado) =>
        Assert.Equal(esperado, TabelaFormatos.GravaComo(TabelaFormatos.DaExtensao(ext)!));

    [Theory] // papeis de cada extensao especial
    [InlineData(".dotx", true, false, false)] [InlineData(".dotm", true, true, false)] [InlineData(".dot", true, false, false)]
    [InlineData(".docm", false, true, false)] [InlineData(".ppsx", false, false, true)] [InlineData(".ppsm", false, true, true)]
    [InlineData(".pps", false, false, true)] [InlineData(".potx", true, false, false)] [InlineData(".xltm", true, true, false)]
    public void Papeis(string ext, bool modelo, bool macro, bool apresentar)
    {
        var i = TabelaFormatos.DaExtensao(ext)!;
        Assert.Equal((modelo, macro, apresentar), (i.Modelo, i.Macro, i.Apresentar));
    }

    [Fact] // Principal = a primeira linha do formato; Extensao() e DocType() saem dela
    public void Principal_ExtensaoEDocType()
    {
        Assert.Equal("doc", FormatoDocumento.Doc.Extensao());
        Assert.Equal("pptx", FormatoDocumento.Pptx.Extensao());
        Assert.Equal("pdf", FormatoDocumento.PdfA.Extensao());
        Assert.Equal("slide", FormatoDocumento.Ppsx.DocType());
        Assert.Equal("cell", FormatoDocumento.Xlsb.DocType());
        Assert.Equal("pdf", FormatoDocumento.Pdf.DocType());
    }

    [Fact] // extensao desconhecida: mensagem sem "PDF" e sem codigo do x2t
    public void DaExtensao_Desconhecida_MensagemSemPdf()
    {
        var e = Assert.Throws<ErroConversao>(() => FormatoDocumentoExtensoes.DaExtensao(@"C:\x\a.pages"));
        Assert.DoesNotContain("PDF", e.Message);
        Assert.Contains(".pages", e.Message);
    }

    [Fact] // os seis ProgIDs de antes continuam com o MESMO nome (o Windows guarda a escolha do usuario por ProgID)
    public void ProgIdsAntigos_Mantidos()
    {
        foreach (var (ext, progId) in new[] { (".docx", "mOffice.Texto.docx"), (".odt", "mOffice.Texto.odt"), (".xlsx", "mOffice.Planilha.xlsx"),
                     (".ods", "mOffice.Planilha.ods"), (".pptx", "mOffice.Slides.pptx"), (".odp", "mOffice.Slides.odp") })
            Assert.Equal(progId, TabelaFormatos.DaExtensao(ext)!.ProgId);
    }

    [Fact] // toda extensao que abre tem ProgID (aparece no "Abrir com"); a que nao abre, nao
    public void ProgId_SoNasQueAbrem() =>
        Assert.All(TabelaFormatos.Todas, i => Assert.Equal(i.Abre && i.Extensao != ".txt", i.ProgId is not null));

    [Fact] // .txt: o Bloco de Notas continua o dono (nenhuma associacao); .csv aparece no "Abrir com" com o ProgID proprio — o Excel segue o padrao
    public void TxtSemProgId_CsvComProgId()
    {
        Assert.Null(TabelaFormatos.DaExtensao(".txt")!.ProgId);
        Assert.Equal("mOffice.Planilha.csv", TabelaFormatos.DaExtensao(".csv")!.ProgId);
        Assert.All(new[] { ".txt", ".csv" }, e => Assert.False(TabelaFormatos.DaExtensao(e)!.PadraoNoRmm));
    }

    [Fact] // TXT e CSV abrem e gravam (Task 3); PDF continua so' exportacao (o mTexto o abre como documento pelo caminho proprio, nao pela tabela)
    public void TxtCsv_AbremEGravam_PdfSoGrava()
    {
        foreach (var e in new[] { ".txt", ".csv" }) { var i = TabelaFormatos.DaExtensao(e)!; Assert.True(i.Abre); Assert.True(i.Grava); Assert.True(i.PerdeFormatacao); Assert.Equal("texto", i.Familia); }
        var pdf = TabelaFormatos.DaExtensao(".pdf")!; Assert.False(pdf.Abre); Assert.True(pdf.Grava);
    }

    [Fact] // extensao repetida na tabela faria DaExtensao devolver sempre a primeira em silencio
    public void Extensoes_SemRepeticao() =>
        Assert.Equal(TabelaFormatos.Todas.Count, TabelaFormatos.Todas.Select(i => i.Extensao).Distinct(StringComparer.Ordinal).Count());

    [Fact] // toda extensao: ponto na frente e minuscula (o contrato do record)
    public void Extensoes_ComPontoEMinusculas() =>
        Assert.All(TabelaFormatos.Todas, i => { Assert.StartsWith(".", i.Extensao); Assert.Equal(i.Extensao.ToLowerInvariant(), i.Extensao); });

    [Theory] // a extensao digitada em maiusculas e o caminho inteiro resolvem
    [InlineData(@"C:\x\A.DOC", ".doc")] [InlineData(".XLSM", ".xlsm")] [InlineData("a.pdf", ".pdf")]
    public void DaExtensao_IgnoraCaixa(string entrada, string esperada) => Assert.Equal(esperada, TabelaFormatos.DaExtensao(entrada)!.Extensao);

    [Theory] // todo codigo do enum e' o c_oAscFileType do sdkjs (valores do brief, conferidos no bundle 9.4)
    [InlineData(FormatoDocumento.Docx, 65)] [InlineData(FormatoDocumento.Doc, 66)] [InlineData(FormatoDocumento.Rtf, 68)]
    [InlineData(FormatoDocumento.Html, 70)] [InlineData(FormatoDocumento.Docm, 75)] [InlineData(FormatoDocumento.Dotx, 76)]
    [InlineData(FormatoDocumento.Pptx, 129)] [InlineData(FormatoDocumento.Ppsx, 132)] [InlineData(FormatoDocumento.Xlsx, 257)]
    [InlineData(FormatoDocumento.Xlsm, 261)] [InlineData(FormatoDocumento.Xlsb, 264)] [InlineData(FormatoDocumento.Pdf, 513)]
    [InlineData(FormatoDocumento.PdfA, 521)]
    public void Codigos_ConferemComOSdk(FormatoDocumento f, int codigo) => Assert.Equal(codigo, (int)f);

    [Fact] // todo formato do enum que alguem possa gravar tem uma linha principal na tabela (Extensao()/DocType() nao quebram)
    public void TodoFormatoTemPrincipal() =>
        Assert.All(Enum.GetValues<FormatoDocumento>(), f => Assert.NotNull(TabelaFormatos.Principal(f)));
}
