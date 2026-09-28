using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace mOffice.Conversor.Tests;

/// <summary>
/// Amostras "reais sinteticas" da Task 7. `poc/amostras/` esta VAZIA (o escritorio ainda nao entregou arquivos
/// reais) e `FabricaDocumentos` so gera o minimo para round-trip; para medir tempo e fidelidade de verdade
/// precisamos de documentos com a estrutura de um arquivo de trabalho: titulos, tabela, imagem, cabecalho e
/// rodape; 3 planilhas com formulas; 3 slides.
///
/// Tambem gera os arquivos GRANDES (5/20/50 MB) dos limiares do x2t, em tres familias, porque "20 MB" nao diz
/// nada sozinho — o que custa ao x2t nao e' o tamanho, e' o que tem dentro:
///   - DOCX com imagens (PNG de RUIDO, incompressivel): e' a peticao com fotos, o caso do limiar. O x2t
///     basicamente copia a midia, entao e' o extremo BARATO.
///   - DOCX so de paragrafos: ~5 mil paginas num arquivo de 20 MB. Nao e' um documento realista, mas e' o
///     extremo CARO — mede o motor processando texto de verdade.
///   - XLSX por LINHAS, que e' como planilha grande cresce de verdade.
/// O relatorio traz os tres; o veredito do limiar olha o caso realista e registra os outros como contexto.
///
/// Nada aqui sai do git: os arquivos vao para `poc/amostras/sinteticas/` (`poc/amostras/` esta no .gitignore).
/// </summary>
public static class AmostrasSinteticas
{
    public const string Marca = "mOffice PoC amostra sintetica rica";

    // ----------------------------------------------------------------------------------------------------
    // PNG de ruido: bytes aleatorios NAO comprimem, entao o tamanho do arquivo e' previsivel. Escrito na mao
    // (assinatura + IHDR + IDAT + IEND) para nao trazer dependencia de imagem so para gerar amostra.
    // ----------------------------------------------------------------------------------------------------
    public static byte[] PngRuido(int lado, int semente)
    {
        var rnd = new Random(semente);
        // Linhas do PNG: 1 byte de filtro (0 = None) + largura*3 bytes RGB.
        var cru = new byte[lado * (1 + lado * 3)];
        for (var y = 0; y < lado; y++)
        {
            var i = y * (1 + lado * 3);
            cru[i] = 0;
            rnd.NextBytes(cru.AsSpan(i + 1, lado * 3));
        }
        using var comprimido = new MemoryStream();
        using (var z = new System.IO.Compression.ZLibStream(comprimido, System.IO.Compression.CompressionLevel.Fastest, true))
            z.Write(cru, 0, cru.Length);
        var idat = comprimido.ToArray();

        using var png = new MemoryStream();
        png.Write([0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A]);
        var ihdr = new byte[13];
        EscreverBigEndian(ihdr, 0, lado); EscreverBigEndian(ihdr, 4, lado);
        ihdr[8] = 8;   // bits por canal
        ihdr[9] = 2;   // truecolor RGB
        Bloco(png, "IHDR", ihdr);
        Bloco(png, "IDAT", idat);
        Bloco(png, "IEND", []);
        return png.ToArray();
    }

    private static void EscreverBigEndian(byte[] destino, int pos, int valor)
    {
        destino[pos] = (byte)(valor >> 24); destino[pos + 1] = (byte)(valor >> 16);
        destino[pos + 2] = (byte)(valor >> 8); destino[pos + 3] = (byte)valor;
    }

    private static void Bloco(Stream saida, string tipo, byte[] dados)
    {
        var tamanho = new byte[4]; EscreverBigEndian(tamanho, 0, dados.Length);
        saida.Write(tamanho);
        var corpo = new byte[4 + dados.Length];
        for (var i = 0; i < 4; i++) corpo[i] = (byte)tipo[i];
        dados.CopyTo(corpo, 4);
        saida.Write(corpo);
        var crc = new byte[4]; EscreverBigEndian(crc, 0, unchecked((int)Crc32(corpo)));
        saida.Write(crc);
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

    // ----------------------------------------------------------------------------------------------------
    // Amostras RICAS (estrutura de documento de trabalho)
    // ----------------------------------------------------------------------------------------------------

    /// <summary>DOCX com titulos, paragrafos, uma tabela 3x3, uma imagem PNG embutida e cabecalho + rodape.</summary>
    public static string DocxRico(string pasta, string nome = "rico.docx")
    {
        var caminho = Path.Combine(pasta, nome);
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var corpo = new W.Body();
        main.Document = new W.Document(corpo);
        main.AddNewPart<StyleDefinitionsPart>().Styles = EstilosDeTitulo();

        corpo.Append(Paragrafo(Marca, "Heading1"));
        corpo.Append(Paragrafo("Este documento existe para medir tempo de abertura e fidelidade de round-trip. Ele tem título, subtítulo, tabela, imagem, cabeçalho e rodapé."));
        corpo.Append(Paragrafo("Seção 1 — texto corrido", "Heading2"));
        for (var i = 1; i <= 12; i++)
            corpo.Append(Paragrafo($"Parágrafo {i}: acentuação (ção, ã, é, ô), números 1.234,56 e uma citação entre aspas."));
        corpo.Append(Paragrafo("Seção 2 — tabela", "Heading2"));
        corpo.Append(Tabela3x3());
        corpo.Append(Paragrafo("Seção 3 — imagem", "Heading2"));
        corpo.Append(ParagrafoComImagem(main, PngRuido(96, 7), 96, 96));
        corpo.Append(new W.SectionProperties(CabecalhoERodape(main, corpo)));
        main.Document.Save();
        return caminho;
    }

    /// <summary>XLSX com 3 planilhas: dados, fórmulas (SUM/AVERAGE/IF) e resumo que referencia as outras.</summary>
    public static string XlsxRico(string pasta, string nome = "rico.xlsx")
    {
        var caminho = Path.Combine(pasta, nome);
        using var doc = SpreadsheetDocument.Create(caminho, SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart(); wb.Workbook = new S.Workbook();
        var planilhas = new S.Sheets();

        var dados = wb.AddNewPart<WorksheetPart>();
        var linhas = new S.SheetData();
        linhas.Append(Linha(1, [Texto("A1", Marca), Texto("B1", "Valor"), Texto("C1", "Quantidade")]));
        for (uint i = 2; i <= 11; i++)
            linhas.Append(Linha(i, [Texto($"A{i}", $"Item {i - 1}"), Numero($"B{i}", (i - 1) * 10.5), Numero($"C{i}", i - 1)]));
        dados.Worksheet = new S.Worksheet(linhas);
        planilhas.Append(new S.Sheet { Id = wb.GetIdOfPart(dados), SheetId = 1, Name = "Dados" });

        var calculo = wb.AddNewPart<WorksheetPart>();
        calculo.Worksheet = new S.Worksheet(new S.SheetData(
            Linha(1, [Texto("A1", "Soma"), Formula("B1", "SUM(Dados!B2:B11)")]),
            Linha(2, [Texto("A2", "Média"), Formula("B2", "AVERAGE(Dados!B2:B11)")]),
            Linha(3, [Texto("A3", "Condicional"), Formula("B3", "IF(B1>100,\"alto\",\"baixo\")")])));
        planilhas.Append(new S.Sheet { Id = wb.GetIdOfPart(calculo), SheetId = 2, Name = "Calculo" });

        var resumo = wb.AddNewPart<WorksheetPart>();
        resumo.Worksheet = new S.Worksheet(new S.SheetData(
            Linha(1, [Texto("A1", "Total geral"), Formula("B1", "Calculo!B1*2")])));
        planilhas.Append(new S.Sheet { Id = wb.GetIdOfPart(resumo), SheetId = 3, Name = "Resumo" });

        wb.Workbook.AppendChild(planilhas);
        wb.Workbook.Save();
        return caminho;
    }

    /// <summary>PPTX com 3 slides (a fixture tem 1; os outros dois sao clones dela com texto proprio).</summary>
    public static string PptxRico(string pasta, string nome = "rico.pptx")
    {
        var caminho = FabricaDocumentos.Pptx(pasta, Marca);
        var destino = Path.Combine(pasta, nome);
        if (!string.Equals(caminho, destino, StringComparison.OrdinalIgnoreCase)) { File.Move(caminho, destino, true); }
        using var doc = PresentationDocument.Open(destino, true);
        var pres = doc.PresentationPart!;
        var modelo = pres.SlideParts.First();
        var lista = pres.Presentation.SlideIdList!;
        var maiorId = lista.Elements<P.SlideId>().Max(s => s.Id!.Value);
        foreach (var texto in new[] { "Slide 2 — segundo slide da amostra", "Slide 3 — terceiro slide da amostra" })
        {
            var novo = pres.AddNewPart<SlidePart>();
            using (var origem = modelo.GetStream(FileMode.Open, FileAccess.Read)) novo.FeedData(origem);
            novo.AddPart(modelo.SlideLayoutPart!);
            novo.Slide.Descendants<A.Text>().First().Text = texto;
            novo.Slide.Save();
            lista.Append(new P.SlideId { Id = ++maiorId, RelationshipId = pres.GetIdOfPart(novo) });
        }
        pres.Presentation.Save();
        return destino;
    }

    // ----------------------------------------------------------------------------------------------------
    // Amostras GRANDES (limiares do x2t)
    // ----------------------------------------------------------------------------------------------------

    /// <summary>DOCX de aproximadamente <paramref name="alvoMb"/> MB: corpo de texto + imagens PNG de ruido
    /// ate atingir o tamanho. Devolve o caminho.</summary>
    public static string DocxGrande(string pasta, int alvoMb, string? nome = null)
    {
        var caminho = Path.Combine(pasta, nome ?? $"grande-{alvoMb}mb.docx");
        // Cada imagem de 512x512 RGB de ruido da' ~0,78 MB no arquivo (ruido nao comprime).
        const int Lado = 512;
        var porImagem = (double)PngRuido(Lado, 1).Length;
        var quantasImagens = Math.Max(1, (int)Math.Ceiling(alvoMb * 1024.0 * 1024.0 / porImagem));
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var corpo = new W.Body();
        main.Document = new W.Document(corpo);
        main.AddNewPart<StyleDefinitionsPart>().Styles = EstilosDeTitulo();
        corpo.Append(Paragrafo($"{Marca} — arquivo grande de {alvoMb} MB", "Heading1"));
        for (var i = 1; i <= 400; i++)
            corpo.Append(Paragrafo($"Parágrafo {i} do corpo de texto, com acentuação e números (1.234,56), para o x2t ter texto de verdade além das imagens."));
        for (var i = 0; i < quantasImagens; i++)
        {
            corpo.Append(Paragrafo($"Figura {i + 1}", "Heading2"));
            corpo.Append(ParagrafoComImagem(main, PngRuido(Lado, 1000 + i), 240, 240));
        }
        main.Document.Save();
        return caminho;
    }

    /// <summary>DOCX de <paramref name="paginas"/> paginas (quebra de pagina explicita), cada uma com um titulo
    /// "Página N", um paragrafo de texto e uma imagem PNG de ruido DIFERENTE (160x160) — o documento grande da
    /// medida do visualizador de PDF (Task 5 do Plano 3). A ULTIMA pagina termina com
    /// <c>ultima-pagina-{marca}</c>, para provar que a impressao chegou ao fim.</summary>
    public static string DocxPaginasComImagens(string pasta, int paginas, string marca, string? nome = null)
    {
        var caminho = Path.Combine(pasta, nome ?? $"paginas-{paginas}.docx");
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var corpo = new W.Body();
        main.Document = new W.Document(corpo);
        for (var i = 1; i <= paginas; i++)
        {
            corpo.Append(Paragrafo($"Página {i}"));
            corpo.Append(Paragrafo($"Texto da página {i}, com acentuação e números (1.234,56), para o PDF ter texto de verdade em toda página."));
            corpo.Append(ParagrafoComImagem(main, PngRuido(160, 5000 + i), 160, 160));
            if (i < paginas) corpo.Append(new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })));
        }
        corpo.Append(Paragrafo("ultima-pagina-" + marca));
        main.Document.Save();
        return caminho;
    }

    /// <summary>DOCX de aproximadamente <paramref name="alvoMb"/> MB feito SO de paragrafos (sem imagem), escrito
    /// em fluxo. E' o outro extremo do arquivo grande: o de imagens mede quanto o x2t copia midia, este mede
    /// quanto ele processa TEXTO — e os dois tempos sao muito diferentes, entao o relatorio traz os dois.</summary>
    public static string DocxGrandeTexto(string pasta, int alvoMb, string? nome = null)
    {
        var caminho = Path.Combine(pasta, nome ?? $"grande-{alvoMb}mb-texto.docx");
        // ~62 bytes por paragrafo depois do zip (MEDIDO: texto pseudoaleatorio comprime pouco). O tamanho real
        // de cada arquivo entra no relatorio, entao o veredito nunca depende desta estimativa.
        var paragrafos = Math.Max(100, (int)(alvoMb * 1024.0 * 1024.0 / 62.0));
        var rnd = new Random(7);
        const string Alfabeto = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 ";
        using var doc = WordprocessingDocument.Create(caminho, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        using (var w = OpenXmlWriter.Create(main))
        {
            w.WriteStartElement(new W.Document());
            w.WriteStartElement(new W.Body());
            var buffer = new char[90];
            for (var i = 0; i < paragrafos; i++)
            {
                for (var c = 0; c < buffer.Length; c++) buffer[c] = Alfabeto[rnd.Next(Alfabeto.Length)];
                w.WriteStartElement(new W.Paragraph());
                w.WriteStartElement(new W.Run());
                w.WriteElement(new W.Text(new string(buffer)) { Space = SpaceProcessingModeValues.Preserve });
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        }
        return caminho;
    }

    /// <summary>XLSX de aproximadamente <paramref name="alvoMb"/> MB: linhas de verdade (8 colunas + 1 formula),
    /// escritas em fluxo (OpenXmlWriter) para nao montar a arvore inteira em memoria.</summary>
    public static string XlsxGrande(string pasta, int alvoMb, string? nome = null)
    {
        var caminho = Path.Combine(pasta, nome ?? $"grande-{alvoMb}mb.xlsx");
        // ~71 bytes por linha depois do zip (MEDIDO nesta maquina: 616.809 linhas deram 41,7 MB). O tamanho
        // real de cada arquivo entra no relatorio de medidas, entao o veredito nunca depende desta estimativa —
        // ela so evita gerar um arquivo com o dobro ou a metade do alvo.
        var linhas = Math.Max(100, (int)(alvoMb * 1024.0 * 1024.0 / 71.0));
        using var doc = SpreadsheetDocument.Create(caminho, SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        var ws = wb.AddNewPart<WorksheetPart>();
        using (var w = OpenXmlWriter.Create(ws))
        {
            w.WriteStartElement(new S.Worksheet());
            w.WriteStartElement(new S.SheetData());
            var rnd = new Random(42);
            for (var r = 1; r <= linhas; r++)
            {
                w.WriteStartElement(new S.Row { RowIndex = (uint)r });
                for (var c = 0; c < 8; c++)
                {
                    w.WriteStartElement(new S.Cell { CellReference = $"{(char)('A' + c)}{r}" });
                    w.WriteElement(new S.CellValue((rnd.Next(1000000) / 100.0).ToString(System.Globalization.CultureInfo.InvariantCulture)));
                    w.WriteEndElement();
                }
                w.WriteStartElement(new S.Cell { CellReference = $"I{r}" });
                w.WriteElement(new S.CellFormula($"SUM(A{r}:H{r})"));
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndElement();
        }
        wb.Workbook = new S.Workbook(new S.Sheets(new S.Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Grande" }));
        wb.Workbook.Save();
        return caminho;
    }

    // ----------------------------------------------------------------------------------------------------

    private static W.Styles EstilosDeTitulo() => new(
        new W.Style(new W.StyleName { Val = "heading 1" }, new W.StyleParagraphProperties(new W.OutlineLevel { Val = 0 }),
                    new W.StyleRunProperties(new W.Bold(), new W.FontSize { Val = "32" }))
        { Type = W.StyleValues.Paragraph, StyleId = "Heading1" },
        new W.Style(new W.StyleName { Val = "heading 2" }, new W.StyleParagraphProperties(new W.OutlineLevel { Val = 1 }),
                    new W.StyleRunProperties(new W.Bold(), new W.FontSize { Val = "26" }))
        { Type = W.StyleValues.Paragraph, StyleId = "Heading2" });

    private static W.Paragraph Paragrafo(string texto, string? estilo = null)
    {
        var p = new W.Paragraph(new W.Run(new W.Text(texto) { Space = SpaceProcessingModeValues.Preserve }));
        if (estilo is not null) p.PrependChild(new W.ParagraphProperties(new W.ParagraphStyleId { Val = estilo }));
        return p;
    }

    private static W.Table Tabela3x3()
    {
        var t = new W.Table(new W.TableProperties(new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Size = 4 }, new W.BottomBorder { Val = W.BorderValues.Single, Size = 4 },
            new W.LeftBorder { Val = W.BorderValues.Single, Size = 4 }, new W.RightBorder { Val = W.BorderValues.Single, Size = 4 },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4 }, new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 4 })));
        for (var l = 0; l < 3; l++)
        {
            var linha = new W.TableRow();
            for (var c = 0; c < 3; c++)
                linha.Append(new W.TableCell(new W.TableCellProperties(new W.TableCellWidth { Width = "2000", Type = W.TableWidthUnitValues.Dxa }),
                                             Paragrafo(l == 0 ? $"Coluna {c + 1}" : $"L{l}C{c + 1}")));
            t.Append(linha);
        }
        return t;
    }

    private static W.Paragraph ParagrafoComImagem(MainDocumentPart main, byte[] png, int larguraPx, int alturaPx)
    {
        var parte = main.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(png)) parte.FeedData(ms);
        var id = main.GetIdOfPart(parte);
        // EMU: 1 px a 96 dpi = 9525 EMU.
        long cx = larguraPx * 9525L, cy = alturaPx * 9525L;
        var nome = "Imagem" + id;
        var desenho = new W.Drawing(new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = 1U, Name = nome },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = nome + ".png" },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(new A.Blip { Embed = id }, new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = cx, Cy = cy }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
        return new W.Paragraph(new W.Run(desenho));
    }

    private static IEnumerable<OpenXmlElement> CabecalhoERodape(MainDocumentPart main, W.Body corpo)
    {
        var cab = main.AddNewPart<HeaderPart>();
        cab.Header = new W.Header(Paragrafo("Cabeçalho da amostra sintética"));
        cab.Header.Save();
        var rod = main.AddNewPart<FooterPart>();
        rod.Footer = new W.Footer(Paragrafo("Rodapé da amostra sintética"));
        rod.Footer.Save();
        yield return new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(cab) };
        yield return new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(rod) };
    }

    private static S.Row Linha(uint indice, IEnumerable<S.Cell> celulas)
    {
        var l = new S.Row { RowIndex = indice };
        foreach (var c in celulas) l.Append(c);
        return l;
    }
    private static S.Cell Texto(string referencia, string texto)
        => new() { CellReference = referencia, DataType = S.CellValues.InlineString, InlineString = new S.InlineString(new S.Text(texto)) };
    private static S.Cell Numero(string referencia, double valor)
        => new() { CellReference = referencia, CellValue = new S.CellValue(valor.ToString(System.Globalization.CultureInfo.InvariantCulture)) };
    private static S.Cell Formula(string referencia, string formula)
        => new() { CellReference = referencia, CellFormula = new S.CellFormula(formula) };
}
