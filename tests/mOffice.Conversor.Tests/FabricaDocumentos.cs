using System.Buffers.Binary;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using S = DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace mOffice.Conversor.Tests;

/// <summary>Gera DOCX/XLSX minimos com um texto conhecido, para round-trip pelo x2t. PPTX vem de fixture (ver abaixo).</summary>
public static class FabricaDocumentos
{
    public static string Docx(string pasta, string texto)
    {
        var caminho = Path.Combine(pasta, "sintetico.docx");
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(texto)))));
        main.Document.Save();
        return caminho;
    }

    /// <summary>
    /// DOCX PROTEGIDO POR SENHA (Task 7 do Plano 3), gerado pelo proprio x2t: nenhuma biblioteca daqui cifra OOXML
    /// (o pacote cifrado e' um arquivo composto OLE com EncryptionInfo/EncryptedPackage), e o x2t nao cria arquivo
    /// com senha do nada. O caminho e' o mesmo do salvar do produto: <see cref="Docx"/> -&gt; <c>ParaBinAsync</c> numa
    /// pasta de recuperacao temporaria -&gt; <c>DeBinAsync(..., senhaSalvar)</c> -&gt; <c>protegido.docx</c>. A pasta
    /// temporaria e' apagada no fim; o arquivo fica em <paramref name="pasta"/>.
    /// </summary>
    public static async Task<string> DocxComSenha(string pasta, string texto, string senha)
    {
        var conv = new X2tConversor(MotorDeTeste.Localizar());
        var trabalho = Path.Combine(pasta, "gerar-protegido-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(trabalho);
        try
        {
            var aberto = Docx(trabalho, texto);
            var recovery = Path.Combine(trabalho, "recovery");
            await conv.ParaBinAsync(aberto, recovery);
            var destino = Path.Combine(pasta, "protegido.docx");
            await conv.DeBinAsync(recovery, FormatoDocumento.Docx, destino, senhaSalvar: senha);
            return destino;
        }
        finally { try { Directory.Delete(trabalho, true); } catch { } }
    }

    /// <summary>DOCX com um paragrafo de texto MAIS uma figura inline (PNG). E' o documento minimo que a
    /// Task 5 (S4) precisa: so' um documento COM imagem prova que a midia da pasta de recovery chega a tela —
    /// o DOCX so' de texto do <see cref="Docx"/> abre igual com a midia quebrada.
    /// O PNG padrao e' um quadrado vermelho solido de 64x64 (<see cref="PngSolido"/>): cor que nao existe em
    /// nenhum outro lugar da tela do editor, para a captura poder conta-la (ver Pixels nos testes do Nucleo).</summary>
    public static string DocxComImagem(string pasta, string texto, byte[]? png = null, int ladoPx = 64)
    {
        png ??= PngSolido(ladoPx, 255, 0, 0);
        var caminho = Path.Combine(pasta, "sintetico-imagem.docx");
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new W.Document(new W.Body());
        var corpo = main.Document.Body!;
        corpo.AppendChild(new W.Paragraph(new W.Run(new W.Text(texto))));

        var parte = main.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(png)) parte.FeedData(ms);
        corpo.AppendChild(new W.Paragraph(new W.Run(DesenhoInline(main.GetIdOfPart(parte), ladoPx, ladoPx))));
        main.Document.Save();
        return caminho;
    }

    /// <summary>`w:drawing` inline apontando para a relacao da imagem — o desenho minimo que o Word (e o x2t)
    /// aceitam: extensao, propriedades do desenho, o `pic:pic` com o blip e a geometria retangular.</summary>
    private static W.Drawing DesenhoInline(string idRelacao, int larguraPx, int alturaPx)
    {
        const long EmuPorPixel = 9525; // 914400 EMU por polegada / 96 dpi
        long cx = larguraPx * EmuPorPixel, cy = alturaPx * EmuPorPixel;
        return new W.Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = 1U, Name = "figura" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "figura.png" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = idRelacao },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
    }

    /// <summary>PNG de cor solida, escrito na mao (IHDR + IDAT zlib + IEND, 8 bits truecolor). Sem
    /// <c>System.Drawing</c> nem WPF: este projeto de teste e' `net10.0` puro, e um PNG de cor solida sao ~30
    /// linhas — menos do que trazer uma dependencia grafica so' para gerar um quadrado.</summary>
    public static byte[] PngSolido(int lado, byte r, byte g, byte b)
    {
        var bruto = new byte[lado * (1 + lado * 3)];
        var p = 0;
        for (var y = 0; y < lado; y++)
        {
            bruto[p++] = 0; // tipo de filtro da linha: None
            for (var x = 0; x < lado; x++) { bruto[p++] = r; bruto[p++] = g; bruto[p++] = b; }
        }
        byte[] comprimido;
        using (var saida = new MemoryStream())
        {
            using (var z = new ZLibStream(saida, CompressionLevel.Optimal, leaveOpen: true)) z.Write(bruto, 0, bruto.Length);
            comprimido = saida.ToArray();
        }
        var ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(0), lado);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4), lado);
        ihdr[8] = 8;  // bits por amostra
        ihdr[9] = 2;  // tipo de cor: truecolor (RGB)
        using var png = new MemoryStream();
        png.Write([0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A]);
        EscreverBloco(png, "IHDR", ihdr);
        EscreverBloco(png, "IDAT", comprimido);
        EscreverBloco(png, "IEND", []);
        return png.ToArray();
    }

    private static void EscreverBloco(Stream destino, string tipo, byte[] dados)
    {
        Span<byte> tamanho = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(tamanho, dados.Length);
        destino.Write(tamanho);
        var corpo = new byte[4 + dados.Length];
        for (var i = 0; i < 4; i++) corpo[i] = (byte)tipo[i];
        dados.CopyTo(corpo, 4);
        destino.Write(corpo);
        Span<byte> crc = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crc, Crc32(corpo));
        destino.Write(crc);
    }

    private static readonly uint[] TabelaCrc = CriarTabelaCrc();
    private static uint[] CriarTabelaCrc()
    {
        var t = new uint[256];
        for (uint n = 0; n < 256; n++)
        {
            var c = n;
            for (var k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            t[n] = c;
        }
        return t;
    }

    private static uint Crc32(byte[] dados)
    {
        var c = 0xFFFFFFFFu;
        foreach (var b in dados) c = TabelaCrc[(c ^ b) & 0xFF] ^ (c >> 8);
        return c ^ 0xFFFFFFFFu;
    }

    public static string Xlsx(string pasta, string textoA1)
    {
        var caminho = Path.Combine(pasta, "sintetico.xlsx");
        using var doc = SpreadsheetDocument.Create(caminho, SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart(); wb.Workbook = new S.Workbook();
        var ws = wb.AddNewPart<WorksheetPart>();
        var celula = new S.Cell { CellReference = "A1", DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text(textoA1)) };
        ws.Worksheet = new S.Worksheet(new S.SheetData(new S.Row(celula) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new S.Sheets(new S.Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Plan1" }));
        wb.Workbook.Save();
        return caminho;
    }

    /// <summary>PPTX exige master/layout/tema (dezenas de linhas de OpenXML). Fixture: poc/mOffice.Conversor.Tests/Fixtures/sintetico.pptx,
    /// 1 slide com a caixa de texto "mOffice PoC round-trip 2026", salvo pelo PowerPoint (sintetico, sem conteudo do escritorio; commitado).</summary>
    public static string Pptx(string pasta, string texto)
    {
        var origem = Path.Combine(AppContext.BaseDirectory, "Fixtures", "sintetico.pptx");
        if (!File.Exists(origem)) throw new FileNotFoundException("Fixture sintetico.pptx ausente (crie no PowerPoint com o texto da constante Frase).", origem);
        var destino = Path.Combine(pasta, "sintetico.pptx"); File.Copy(origem, destino, true);
        // A fixture tem UMA caixa de texto; o texto dela e' reescrito com o `texto` pedido (a Task 7 precisa de
        // uma frase de origem propria, diferente da usada pelos testes do x2t).
        using (var doc = PresentationDocument.Open(destino, true))
        {
            var t = doc.PresentationPart!.SlideParts.First().Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().First();
            t.Text = texto;
            doc.PresentationPart.SlideParts.First().Slide.Save();
        }
        return destino;
    }

    public static string TextoDocx(string caminho) { using var d = WordprocessingDocument.Open(caminho, false); return d.MainDocumentPart!.Document.Body!.InnerText; }

    /// <summary>Bytes de cada imagem embutida num DOCX (as <c>ImagePart</c> da parte principal). Usado para
    /// provar que uma figura COLADA no editor sobreviveu ao salvamento — o texto nao diz nada sobre isso.</summary>
    public static IReadOnlyList<byte[]> ImagensDocx(string caminho)
    {
        using var d = WordprocessingDocument.Open(caminho, false);
        var imagens = new List<byte[]>();
        foreach (var parte in d.MainDocumentPart!.ImageParts)
        {
            using var fluxo = parte.GetStream();
            using var ms = new MemoryStream();
            fluxo.CopyTo(ms);
            imagens.Add(ms.ToArray());
        }
        return imagens;
    }
    public static string TextoXlsxA1(string caminho) => TextoXlsxCelula(caminho, "A1");

    /// <summary>Texto de UMA celula da PRIMEIRA planilha, resolvendo shared string. Devolve string vazia se a
    /// celula nao existe (antes da edicao, B1 nao existe) — quem chama assere o conteudo, nao a ausencia de
    /// excecao. Generalizado na fix review 1 da Task 7: o teste dos seis formatos precisa afirmar que A1
    /// continua sendo EXATAMENTE o texto de origem e que o texto inserido foi para B1, em vez de olhar so o
    /// texto integral da planilha (que ficaria verde mesmo se a edicao tivesse sobrescrito A1).</summary>
    public static string TextoXlsxCelula(string caminho, string referencia)
    {
        using var d = SpreadsheetDocument.Open(caminho, false);
        var ws = d.WorkbookPart!.WorksheetParts.First();
        var c = ws.Worksheet.Descendants<S.Cell>().FirstOrDefault(x => x.CellReference == referencia);
        if (c is null) return "";
        if (c.DataType?.Value == S.CellValues.SharedString && int.TryParse(c.CellValue?.Text, out var i))
            return d.WorkbookPart.SharedStringTablePart!.SharedStringTable.ElementAt(i).InnerText;
        return c.InnerText;
    }
    /// <summary>Texto integral de um XLSX: todas as planilhas, todas as celulas (inline, shared string e valor
    /// direto). Necessario a partir da Task 7 porque o x2t regrava a planilha com sharedStrings e a edicao pelo
    /// editor cai numa celula diferente de A1 — <see cref="TextoXlsxA1"/> so olha A1 de UMA planilha.</summary>
    public static string TextoXlsx(string caminho)
    {
        using var d = SpreadsheetDocument.Open(caminho, false);
        var wbp = d.WorkbookPart!;
        var compartilhadas = wbp.SharedStringTablePart?.SharedStringTable;
        var sb = new System.Text.StringBuilder();
        foreach (var ws in wbp.WorksheetParts)
            foreach (var c in ws.Worksheet.Descendants<S.Cell>())
            {
                if (c.DataType?.Value == S.CellValues.SharedString && compartilhadas is not null && int.TryParse(c.CellValue?.Text, out var i))
                    sb.Append(compartilhadas.ElementAt(i).InnerText);
                else sb.Append(c.InnerText);
                sb.Append(' ');
            }
        return sb.ToString();
    }

    public static string TextoPptx(string caminho) { using var d = PresentationDocument.Open(caminho, false); return string.Concat(d.PresentationPart!.SlideParts.Select(s => s.Slide.InnerText)); }
    /// <summary>DOCX do teste de janela real da fonte de simbolos livre (Task 1 do Marco 1, decisao 1 do gate):
    /// duas tabulacoes, uma quebra de linha e dois controles de conteudo de caixa de selecao (um marcado, um
    /// nao). Com o ¶ ligado, o editor desenha as setas de tabulacao e a marca de quebra com a familia interna
    /// "ASCW3" — que, depois do patch de tools/onlyoffice/aplicar-patches.ps1, e' a NOSSA fonte
    /// (tools/fontes/mOffice-Simbolos.ttf). As caixas usam a fonte <paramref name="fonteDasCaixas"/>:
    /// "MS Gothic" e' o que o Word grava; "ASCW3" faz o editor desenha-las com a fonte interna tambem.</summary>
    public static string DocxMarcasDeFormatacao(string pasta, string fonteDasCaixas = "MS Gothic", int tamanhoPt = 12)
    {
        var caminho = Path.Combine(pasta, "marcas.docx");
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        // Tamanho do corpo explicito (12 pt). Nao adianta aumentar para engordar as marcas: o motor desenha as
        // setas/quebras SEMPRE a 10 pt (`Ub:10` na chamada que escolhe a fonte ASCW3) — quem as amplia na
        // captura e' o zoom do editor, nao o tamanho do texto.
        var corpo = new W.Body(
            new W.Paragraph(Run(tamanhoPt,
                Texto("Tabulacao"), new W.TabChar(), Texto("depois da tabulacao"), new W.TabChar(), Texto("fim"))),
            new W.Paragraph(Run(tamanhoPt, Texto("Quebra de linha"), new W.Break(), Texto("segunda linha"))),
            new W.Paragraph(
                CaixaDeSelecao(true, fonteDasCaixas, tamanhoPt), Run(tamanhoPt, Texto(" marcada   ")),
                CaixaDeSelecao(false, fonteDasCaixas, tamanhoPt), Run(tamanhoPt, Texto(" desmarcada"))));
        main.Document = new W.Document(corpo);
        main.Document.Save();
        return caminho;
    }

    private static W.Text Texto(string s) => new(s) { Space = SpaceProcessingModeValues.Preserve };

    private static W.Run Run(int tamanhoPt, params OpenXmlElement[] filhos)
    {
        var r = new W.Run(new W.RunProperties(new W.FontSize { Val = (tamanhoPt * 2).ToString() }));
        foreach (var f in filhos) r.Append(f);
        return r;
    }

    /// <summary>Controle de conteudo de caixa de selecao (w14:checkbox), como o Word 2010+ grava: o simbolo
    /// visivel e' U+2611 (marcada) ou U+2610 (vazia) num run com a fonte pedida.</summary>
    private static W.SdtRun CaixaDeSelecao(bool marcada, string fonte, int tamanhoPt)
    {
        var props = new W.SdtProperties(new W14.SdtContentCheckBox(
            new W14.Checked { Val = marcada ? W14.OnOffValues.One : W14.OnOffValues.Zero },
            new W14.CheckedState { Val = new HexBinaryValue("2611"), Font = fonte },
            new W14.UncheckedState { Val = new HexBinaryValue("2610"), Font = fonte }));
        var run = new W.Run(
            new W.RunProperties(
                new W.RunFonts { Ascii = fonte, HighAnsi = fonte, EastAsia = fonte, ComplexScript = fonte },
                new W.FontSize { Val = (tamanhoPt * 2).ToString() }),
            Texto(marcada ? "☑" : "☐"));
        return new W.SdtRun(props, new W.SdtContentRun(run));
    }

    public static string TextoOdf(string caminho) // ODT/ODS/ODP sao ZIP com content.xml
    {
        using var zip = System.IO.Compression.ZipFile.OpenRead(caminho);
        using var sr = new StreamReader(zip.GetEntry("content.xml")!.Open());
        return sr.ReadToEnd();
    }
}
