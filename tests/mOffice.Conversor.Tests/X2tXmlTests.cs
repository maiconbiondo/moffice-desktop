using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

public class X2tXmlTests
{
    private static X2tConversor Conversor() => new(MotorDeTeste.Localizar());

    /// <summary>Senha que so' este teste usa (fix round 1 da Task 7): quem procura "a senha nao sobrou em lugar
    /// nenhum" (%TEMP%, stderr) nao pode ser enganado por outro teste rodando em paralelo com a mesma senha.</summary>
    internal static string SenhaUnica(string rotulo) => $"{rotulo}-{Guid.NewGuid():N}";

    [Fact] // PL-01: o salvar (Editor.bin -> arquivo) leva m_nLcid=1046, senao o cache das formulas de texto sai em ingles
    public void XmlDeBin_LevaLcidPtBr()
    {
        var xml = Conversor().XmlDeBin(@"C:\r\Editor.bin", FormatoDocumento.Xlsx, @"C:\d\a.xlsx", temChanges: true, @"C:\t", jsonParams: null);
        Assert.Equal("1046", xml.Element("m_nLcid")?.Value);
    }

    [Fact] // a constante e' a do pt-BR (literal, nao lida de outro lugar do produto)
    public void LcidPadrao_EhPtBr() => Assert.Equal(1046, X2tConversor.LcidPadrao);

    [Fact] // sem jsonParams o XML do salvar fica identico ao de hoje (nenhum m_sJsonParams)
    public void XmlDeBin_SemJson_SemElemento()
        => Assert.Null(Conversor().XmlDeBin("a.bin", FormatoDocumento.Pdf, "b.pdf", true, "t", null).Element("m_sJsonParams"));

    [Fact] // com jsonParams o elemento vai com o JSON literal
    public void XmlDeBin_ComJson()
        => Assert.Equal("""{"documentLayout":{"isPrint":true}}""",
            (string?)Conversor().XmlDeBin("a.bin", FormatoDocumento.Pdf, "b.pdf", true, "t", """{"documentLayout":{"isPrint":true}}""").Element("m_sJsonParams"));

    [Fact] // a extracao nao reordenou nada: a ordem dos elementos e' a de antes, com m_sJsonParams logo depois de
           // m_nDoctParams (e os comuns no fim)
    public void XmlDeBin_OrdemDosElementos()
    {
        var nomes = Conversor().XmlDeBin("a.bin", FormatoDocumento.Pdf, "b.pdf", false, "t", "{}").Elements().Select(e => e.Name.LocalName).ToList();
        Assert.Equal(["m_sFileFrom", "m_sFileTo", "m_nFormatTo", "m_bFromChanges", "m_bDontSaveAdditional", "m_sAllFontsPath",
            "m_nCsvTxtEncoding", "m_nCsvDelimiter", "m_nDoctParams", "m_nLcid", "m_sJsonParams", "m_sFontDir", "m_sTempDir"], nomes.Take(13));
    }

    [Fact] // PDF/A (Plano 4, Task 2): o x2t 9.4 nao converte Editor.bin com m_nFormatTo 521 (codigo 88, MEDIDO); o pedido
           // certo e' PDF (513) + m_bIsPDFA. PDF comum e documentos nao levam a marca.
    public void XmlDeBin_PdfA_VaiComoPdfComMarca()
    {
        var a = Conversor().XmlDeBin("a.bin", FormatoDocumento.PdfA, "b.pdf", true, "t", null);
        Assert.Equal("513", (string?)a.Element("m_nFormatTo"));
        Assert.Equal("true", (string?)a.Element("m_bIsPDFA"));
        foreach (var f in new[] { FormatoDocumento.Pdf, FormatoDocumento.Docx, FormatoDocumento.Xlsx })
        {
            var x = Conversor().XmlDeBin("a.bin", f, "b", true, "t", null);
            Assert.Equal(((int)f).ToString(), (string?)x.Element("m_nFormatTo"));
            Assert.Null(x.Element("m_bIsPDFA"));
        }
    }

    [Fact] // senha de salvar vai em m_sSavePassword; sem senha (nula ou vazia), nenhum elemento
    public void XmlDeBin_SenhaSalvar()
    {
        Assert.Equal("s3gr3d0", (string?)Conversor().XmlDeBin("a.bin", FormatoDocumento.Docx, "b.docx", true, "t", null, "s3gr3d0").Element("m_sSavePassword"));
        Assert.Null(Conversor().XmlDeBin("a.bin", FormatoDocumento.Docx, "b.docx", true, "t", null, null).Element("m_sSavePassword"));
        Assert.Null(Conversor().XmlDeBin("a.bin", FormatoDocumento.Docx, "b.docx", true, "t", null, "").Element("m_sSavePassword"));
    }

    [Fact] // m_sSavePassword entra logo depois de m_sJsonParams, antes dos elementos comuns; o texto vai XML-escapado
    public void XmlDeBin_SenhaSalvar_PosicaoEEscape()
    {
        var xml = Conversor().XmlDeBin("a.bin", FormatoDocumento.Docx, "b.docx", false, "t", "{}", "a&b<c>");
        var nomes = xml.Elements().Select(e => e.Name.LocalName).ToList();
        Assert.Equal(nomes.IndexOf("m_sJsonParams") + 1, nomes.IndexOf("m_sSavePassword"));
        Assert.Equal("a&b<c>", (string?)xml.Element("m_sSavePassword"));
        Assert.Contains("<m_sSavePassword>a&amp;b&lt;c&gt;</m_sSavePassword>", xml.ToString());
    }

    [SkippableFact] // o x2t cifra com m_sSavePassword, recusa sem senha (90), recusa senha errada (91) e abre com a certa
    public async Task Senha_CicloCompleto()
    {
        Skip.If(!File.Exists(MotorDeTeste.Localizar().X2t), "motor nao obtido");
        var pasta = Directory.CreateTempSubdirectory("moffice-senha-").FullName;
        // Senhas UNICAS deste teste (fix round 1): as buscas por elas no texto nao podem casar com o que outro teste
        // (outro projeto rodando em paralelo) gravou.
        var certa = SenhaUnica("certa");
        var errada = SenhaUnica("errada");
        try
        {
            var cifrado = await FabricaDocumentos.DocxComSenha(pasta, "conteudo-sigiloso", certa);
            // Cifrado de verdade: OOXML protegido e' um arquivo composto OLE (D0 CF 11 E0), nao um ZIP (PK).
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, File.ReadAllBytes(cifrado).Take(4).ToArray());
            var conv = Conversor();
            var e90 = await Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r1")));
            Assert.Equal(90, e90.CodigoSaida);
            Assert.Equal(X2tConversor.CodigoProtegido, e90.CodigoSaida);
            var e91 = await Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r2"), senha: errada));
            Assert.Equal(91, e91.CodigoSaida);
            Assert.Equal(X2tConversor.CodigoSenhaErrada, e91.CodigoSaida);
            // A senha nunca volta no que o x2t devolve (stderr+stdout vao para UltimoErro e para o log do salvar).
            Assert.DoesNotContain(errada, e91.Stderr + e91.Message);
            var bin = await conv.ParaBinAsync(cifrado, Path.Combine(pasta, "r3"), senha: certa);
            Assert.True(new FileInfo(bin).Length > 0);
            // E o conteudo e' o do documento: de volta a DOCX sem senha, o texto esta' la'.
            var claro = Path.Combine(pasta, "claro.docx");
            await conv.DeBinAsync(Path.Combine(pasta, "r3"), FormatoDocumento.Docx, claro);
            Assert.Contains("conteudo-sigiloso", FabricaDocumentos.TextoDocx(claro));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableTheory] // cada formato de FormatoGuardaSenha sai cifrado do x2t (reabrir sem senha = 90, com a senha abre)
    [InlineData(FormatoDocumento.Docx)]
    [InlineData(FormatoDocumento.Odt)]
    [InlineData(FormatoDocumento.Xlsx)]
    [InlineData(FormatoDocumento.Ods)]
    [InlineData(FormatoDocumento.Pptx)]
    [InlineData(FormatoDocumento.Odp)]
    // Plano 6, Task 2: macro, modelo e apresentacao de slides (medidos: todos cifram)
    [InlineData(FormatoDocumento.Docm)] [InlineData(FormatoDocumento.Dotx)] [InlineData(FormatoDocumento.Dotm)] [InlineData(FormatoDocumento.Ott)]
    [InlineData(FormatoDocumento.Xlsm)] [InlineData(FormatoDocumento.Xltx)] [InlineData(FormatoDocumento.Xltm)] [InlineData(FormatoDocumento.Ots)]
    [InlineData(FormatoDocumento.Pptm)] [InlineData(FormatoDocumento.Potx)] [InlineData(FormatoDocumento.Potm)] [InlineData(FormatoDocumento.Ppsx)]
    [InlineData(FormatoDocumento.Ppsm)] [InlineData(FormatoDocumento.Otp)]
    public async Task Senha_SalvarMantemProtecao_PorFormato(FormatoDocumento alvo)
    {
        Skip.If(!File.Exists(MotorDeTeste.Localizar().X2t), "motor nao obtido");
        Assert.True(X2tConversor.FormatoGuardaSenha(alvo));
        var pasta = Directory.CreateTempSubdirectory("moffice-senha-fmt-").FullName;
        try
        {
            var conv = Conversor();
            var origem = alvo.DocType() switch
            {
                "cell" => FabricaDocumentos.Xlsx(pasta, "celula-sigilosa"),
                "slide" => FabricaDocumentos.Pptx(pasta, "slide-sigiloso"),
                _ => FabricaDocumentos.Docx(pasta, "texto-sigiloso"),
            };
            var rec = Path.Combine(pasta, "rec");
            await conv.ParaBinAsync(origem, rec);
            var destino = Path.Combine(pasta, "saida." + alvo.Extensao());
            await conv.DeBinAsync(rec, alvo, destino, senhaSalvar: "s3gr3d0");
            // E o validador aceita o conteiner cifrado (OLE) deste formato — senao o salvar o recusaria.
            ValidadorDePacote.Validar(destino, alvo, cifrado: true);
            var e = await Assert.ThrowsAsync<ErroConversao>(() => conv.ParaBinAsync(destino, Path.Combine(pasta, "sem")));
            Assert.Equal(X2tConversor.CodigoProtegido, e.CodigoSaida);
            var bin = await conv.ParaBinAsync(destino, Path.Combine(pasta, "com"), senha: "s3gr3d0");
            Assert.True(new FileInfo(bin).Length > 0);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [Theory] // RTF e HTML nao guardam senha (MEDIDO: o x2t ignora o m_sSavePassword); legado nunca e' gravado
    [InlineData(FormatoDocumento.Rtf)] [InlineData(FormatoDocumento.Html)] [InlineData(FormatoDocumento.Doc)]
    [InlineData(FormatoDocumento.Xls)] [InlineData(FormatoDocumento.Ppt)] [InlineData(FormatoDocumento.Xlsb)] [InlineData(FormatoDocumento.Txt)]
    [InlineData(FormatoDocumento.Csv)]
    public void FormatoGuardaSenha_NaoParaRtfHtmlELegado(FormatoDocumento f) => Assert.False(X2tConversor.FormatoGuardaSenha(f));

    [Fact] // PDF e PDF/A nunca levam a senha do documento (fix round 1: PJe/e-SAJ recusam PDF com senha; PDF/A proibe)
    public void FormatoGuardaSenha_PdfNao()
    {
        Assert.False(X2tConversor.FormatoGuardaSenha(FormatoDocumento.Pdf));
        Assert.False(X2tConversor.FormatoGuardaSenha(FormatoDocumento.PdfA));
    }

    [SkippableFact] // o x2t CIFRARIA o PDF com m_sSavePassword (/Encrypt, medido) — e' por isso que FormatoGuardaSenha
                    // tem de excluir o PDF; sem senha o PDF sai aberto
    public async Task Senha_ExportarPdf_SoCifraSeReceberSenha()
    {
        Skip.If(!File.Exists(MotorDeTeste.Localizar().X2t), "motor nao obtido");
        var pasta = Directory.CreateTempSubdirectory("moffice-senha-pdf-").FullName;
        try
        {
            var conv = Conversor();
            var rec = Path.Combine(pasta, "rec");
            await conv.ParaBinAsync(FabricaDocumentos.Docx(pasta, "texto"), rec);
            var com = Path.Combine(pasta, "com.pdf");
            var sem = Path.Combine(pasta, "sem.pdf");
            await conv.DeBinAsync(rec, FormatoDocumento.Pdf, com, senhaSalvar: "s3gr3d0");
            await conv.DeBinAsync(rec, FormatoDocumento.Pdf, sem);
            Assert.Contains("/Encrypt", System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(com)));   // controle
            Assert.DoesNotContain("/Encrypt", System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(sem)));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // o XML temporario (que leva a senha em claro) e' apagado mesmo quando o x2t falha
    public async Task XmlTemporario_ApagadoMesmoComErro()
    {
        Skip.If(!File.Exists(MotorDeTeste.Localizar().X2t), "motor nao obtido");
        var pasta = Directory.CreateTempSubdirectory("moffice-senha-xml-").FullName;
        var certa = SenhaUnica("certa");
        var errada = SenhaUnica("errada");
        try
        {
            var cifrado = await FabricaDocumentos.DocxComSenha(pasta, "x", certa);
            await Assert.ThrowsAsync<ErroConversao>(() => Conversor().ParaBinAsync(cifrado, Path.Combine(pasta, "r"), senha: errada));
            var sobras = Directory.GetFiles(Path.GetTempPath(), "moffice-x2t-*.xml")
                .Where(f => { try { var t = File.ReadAllText(f); return t.Contains(certa) || t.Contains(errada); } catch (IOException) { return false; } })
                .ToList();
            Assert.Empty(sobras);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [Fact] // salvar CSV com as opcoes escolhidas substitui os 46/4 fixos e leva o lcid
    public void XmlDeBin_ComOpcoesTexto_UsaAsOpcoes()
    {
        var xml = Conversor().XmlDeBin(@"C:\r\Editor.bin", FormatoDocumento.Csv, @"C:\d\a.csv", true, @"C:\t", null, null, OpcoesTexto.CsvExcelPtBr);
        Assert.Equal(("44", "2", "1046"), (xml.Element("m_nCsvTxtEncoding")?.Value, xml.Element("m_nCsvDelimiter")?.Value, xml.Element("m_nLcid")?.Value));
        Assert.Equal("260", xml.Element("m_nFormatTo")?.Value);
    }

    [Fact] // sem opcoes de texto o salvar continua com os fixos de antes (inofensivos para os outros formatos)
    public void XmlDeBin_SemOpcoesTexto_MantemOsFixos()
    {
        var xml = Conversor().XmlDeBin(@"C:\r\Editor.bin", FormatoDocumento.Xlsx, @"C:\d\a.xlsx", true, @"C:\t", null);
        Assert.Equal(("46", "4", "1046"), (xml.Element("m_nCsvTxtEncoding")?.Value, xml.Element("m_nCsvDelimiter")?.Value, xml.Element("m_nLcid")?.Value));
    }

    [Fact] // TXT: UTF-8 sem delimitador (0)
    public void XmlDeBin_Txt_Utf8SemDelimitador()
    {
        var xml = Conversor().XmlDeBin("a.bin", FormatoDocumento.Txt, "b.txt", true, "t", null, null, OpcoesTexto.TxtUtf8);
        Assert.Equal(("46", "0"), (xml.Element("m_nCsvTxtEncoding")?.Value, xml.Element("m_nCsvDelimiter")?.Value));
    }

    [Fact] // o "Outro" do dialogo (m_nCsvDelimiterChar) vai no XML de abrir e de salvar, so' quando existe
    public void XmlComDelimitadorOutro()
    {
        var o = new OpcoesTexto(44, OpcoesTexto.SemDelimitador, DelimitadorChar: "|");
        Assert.Equal("|", Conversor().XmlParaBin(@"C:\u\a.csv", @"C:\r\Editor.bin", @"C:\t", null, o).Element("m_nCsvDelimiterChar")?.Value);
        Assert.Equal("|", Conversor().XmlDeBin("a.bin", FormatoDocumento.Csv, "b.csv", true, "t", null, null, o).Element("m_nCsvDelimiterChar")?.Value);
        Assert.Null(Conversor().XmlParaBin(@"C:\u\a.csv", @"C:\r\Editor.bin", @"C:\t", null, OpcoesTexto.CsvExcelPtBr).Element("m_nCsvDelimiterChar"));
        Assert.Null(Conversor().XmlDeBin("a.bin", FormatoDocumento.Csv, "b.csv", true, "t", null, null, OpcoesTexto.CsvExcelPtBr).Element("m_nCsvDelimiterChar"));
    }

    [Fact] // abrir CSV leva codificacao/delimitador/lcid no XML do ParaBin; sem opcoes, nenhum dos tres
    public void XmlParaBin_ComOpcoesTexto()
    {
        var com = Conversor().XmlParaBin(@"C:\u\a.csv", @"C:\r\Editor.bin", @"C:\t", null, OpcoesTexto.CsvExcelPtBr);
        Assert.Equal(("44", "2", "1046"), (com.Element("m_nCsvTxtEncoding")?.Value, com.Element("m_nCsvDelimiter")?.Value, com.Element("m_nLcid")?.Value));
        // MEDIDO: com m_nFormatFrom o x2t sai com 88 (sem ele deduz CSV/TXT pela extensao): nunca vai no XML
        Assert.Null(com.Element("m_nFormatFrom"));
        var sem = Conversor().XmlParaBin(@"C:\u\a.docx", @"C:\r\Editor.bin", @"C:\t", null, null);
        Assert.Null(sem.Element("m_nCsvTxtEncoding")); Assert.Null(sem.Element("m_nCsvDelimiter")); Assert.Null(sem.Element("m_nLcid")); Assert.Null(sem.Element("m_nFormatFrom"));
        Assert.Equal("8192", sem.Element("m_nFormatTo")?.Value);
    }

    [Theory] // Plano 6, Task 8: convertFile pede o canvas do editor (8193 word, 8194 cell); fora de 8192-8195 vira o generico
    [InlineData(8193, "8193")] [InlineData(8194, "8194")] [InlineData(8195, "8195")] [InlineData(8192, "8192")]
    [InlineData(513, "8192")] [InlineData(8196, "8192")] [InlineData(0, "8192")]
    public void XmlParaBin_FormatoCanvas(int pedido, string esperado) =>
        Assert.Equal(esperado, Conversor().XmlParaBin(@"C:\u\a.docx", @"C:\r\Editor.bin", @"C:\t", null, null, pedido).Element("m_nFormatTo")?.Value);
}
