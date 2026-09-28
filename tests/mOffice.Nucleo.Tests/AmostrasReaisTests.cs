using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Microsoft.Web.WebView2.Core;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 7b — as linhas REAIS da matriz de fidelidade (documentos de trabalho, fora do git): os cinco arquivos
/// de trabalho do escritorio (mais o ODP derivado do PPTX real pelo LibreOffice, ja' que nenhum ODP foi
/// entregue) pelo MESMO caminho completo do produto que <see cref="SeisFormatosTests"/> ja prova com amostras
/// sinteticas: abrir no editor real -&gt; editar pela API publica do editor -&gt; salvar pelo asc_Save/LocalFileSave
/// -&gt; reler o arquivo do disco.
///
/// Confidencialidade: os arquivos de origem moram em <c>poc/amostras/</c> (fora do git). O ROTULO
/// (real-docx/real-xlsx/...) e' a unica identificacao usada aqui — em nome de teste, em `_saida.WriteLine` e em
/// qualquer arquivo que este teste grava. O mapeamento rotulo -&gt; nome real fica so' em
/// <c>poc/amostras/trabalho/MAPA.txt</c> (tambem fora do git). Nenhuma captura de tela do conteudo real e'
/// commitada: o PNG vai para a pasta de trabalho (git-ignored) para quem estiver rodando a suite localmente
/// inspecionar.
/// </summary>
[Collection(JanelaReal.Nome)]
public class AmostrasReaisTests
{
    /// <summary>Marcador SINTETICO inserido em cada amostra real (Task 8 do Plano 3, criterio 3 da spec §2:
    /// documentos reais EDITADOS e preservados). Ate' o Plano 2 era "mOffice PoC 2026-09-18" (o da secao "Como
    /// preencher as linhas reais" do protocolo do Marco 0); o texto e' neutro de proposito — nada do conteudo do
    /// escritorio. O espaco inicial separa o marcador da palavra onde o cursor estiver.</summary>
    public const string Marcador = " [mOffice-teste]";

    /// <summary>O que se procura no texto salvo: o marcador sem o espaco inicial (a normalizacao de espacos e o
    /// inicio de uma celula podem come-lo).</summary>
    public static readonly string MarcadorNoTexto = Marcador.Trim();

    /// <summary>Celulas candidatas para a edicao de planilha, da mais preferida a mais distante. Planilhas reais
    /// podem ter tabelas largas; procurar a primeira célula livre abaixo do intervalo usado — por isso o teste
    /// TESTA que a celula esta vazia antes de usa-la, em vez de assumir Z1 as cegas.</summary>
    private static readonly string[] CelulasCandidatas = ["Z1", "AZ1", "BZ1", "CZ1", "ZZ1"];

    private const string SemAmostras = "SEM-AMOSTRAS";

    private readonly ITestOutputHelper _saida;
    public AmostrasReaisTests(ITestOutputHelper saida) { _saida = saida; }

    public static TheoryData<string> Rotulos()
    {
        var casos = AmostrasReaisCatalogo.Obter();
        var dados = new TheoryData<string>();
        if (casos.Count == 0) { dados.Add(SemAmostras); return dados; }
        foreach (var c in casos) dados.Add(c.Rotulo);
        return dados;
    }

    [SkippableTheory] // janela real: cada amostra REAL do escritorio (mais o ODP derivado), abrir -> editar -> salvar -> reler
    [MemberData(nameof(Rotulos))]
    public void AbreEditaSalva_AmostraReal(string rotulo)
    {
        Skip.If(rotulo == SemAmostras,
            "amostras reais ausentes: poc/amostras/ nao existe ou nao tem nenhum .docx/.xlsx/.pptx/.odt/.ods " +
            "(pasta ignorada pelo git — o escritorio precisa colocar os arquivos de trabalho la)");
        EditorReal.PularSeNaoDaParaAbrirJanela();

        var caso = AmostrasReaisCatalogo.Obter().Single(c => c.Rotulo == rotulo);
        var pastaTrabalho = AmostrasReaisCatalogo.PastaTrabalho;
        Directory.CreateDirectory(pastaTrabalho);

        // Copia SEMPRE nova, com o nome NEUTRO do rotulo: o original nunca e' tocado, e nenhum nome real de
        // arquivo do escritorio aparece na pasta de trabalho a partir daqui.
        var copia = Path.Combine(pastaTrabalho, rotulo + caso.Extensao);
        File.Copy(caso.CaminhoOrigem, copia, overwrite: true);
        var salvo = Path.Combine(pastaTrabalho, rotulo + "-salvo" + caso.Extensao);
        var png = Path.Combine(pastaTrabalho, rotulo + ".png");
        if (File.Exists(salvo)) File.Delete(salvo);

        var formato = FormatoDocumentoExtensoes.DaExtensao(copia);
        var doctype = formato.DocType();
        var ehPlanilha = formato is FormatoDocumento.Xlsx or FormatoDocumento.Ods;

        // Amostra de texto do ORIGINAL, lida do ARQUIVO (nao do editor) ANTES de abrir — mesma disciplina de
        // SeisFormatosTests.Ler: a linha de base nao pode depender do editor ja ter feito nada.
        var amostraOriginal = NormalizarEspacos(TextoLegivel(copia, caso.Extensao));
        if (amostraOriginal.Length > 200) amostraOriginal = amostraOriginal[..200];
        Assert.False(string.IsNullOrWhiteSpace(amostraOriginal),
            $"[{rotulo}] nao foi possivel extrair nenhum texto do arquivo original para servir de linha de base");
        _saida.WriteLine($"[{rotulo}] amostra do original ({amostraOriginal.Length} chars): {amostraOriginal}");

        // Contagens estruturais do ORIGINAL (paragrafos, tabelas, imagens, planilhas, slides — as mesmas de
        // tools/fidelidade.ps1, estendidas ao ODF), lidas do arquivo ANTES de abrir. So' numeros: nada do
        // conteudo sai daqui.
        var contagensAntes = ContagensEstruturais(copia, caso.Extensao);

        var celula = ehPlanilha ? EscolherCelulaLivre(copia) : "";
        if (ehPlanilha) _saida.WriteLine($"[{rotulo}] celula escolhida para a edicao: {celula}");

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(copia);
            var pronto = await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300);
            _saida.WriteLine($"[{rotulo}] pronto={pronto}");
            Assert.Equal("true", pronto);

            // Captura da janela ANTES da edicao: e' o que prova visualmente que o documento abriu de verdade
            // (imagens, tabelas, layout) — o PNG fica na pasta de trabalho, nunca commitado.
            var wv = await j.ProntoAsync();
            using (var arquivoPng = File.Create(png))
                await wv.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, arquivoPng);
            _saida.WriteLine($"[{rotulo}] captura -> {Path.GetFileName(png)}");

            var podeSalvarAntes = (await j.ExecutarJsAsync(EditorReal.PodeSalvarNoEditor)).Trim('"');
            _saida.WriteLine($"[{rotulo}/{doctype}] asc_isDocumentCanSave antes de editar: {podeSalvarAntes}");

            var inserido = (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(doctype, Marcador, celula))).Trim('"');
            _saida.WriteLine($"[{rotulo}/{doctype}] insercao ({(ehPlanilha ? $"celula {celula}" : "cursor")}): {inserido}");
            Assert.StartsWith("ok", inserido);

            // Espera DETERMINISTICA (mesmo sinal de SeisFormatosTests): em cell/slide a insercao passa por
            // callback e nem sempre esta no modelo quando o script retorna.
            var podeSalvar = await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 100);
            Assert.Equal("true", podeSalvar);

            await j.SalvarComoAsync(salvo, formato);
            Assert.Null(j.UltimoErro);

            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));

        Assert.True(File.Exists(salvo), $"[{rotulo}] o PoC nao gravou o arquivo salvo em {Path.GetFileName(salvo)}");
        var textoDepois = NormalizarEspacos(TextoLegivel(salvo, caso.Extensao));
        _saida.WriteLine($"[{rotulo}] texto salvo ({textoDepois.Length} chars, apos normalizar espacos)");
        Assert.Contains(MarcadorNoTexto, textoDepois);

        // O marcador entra ONDE O CURSOR ESTA' ao abrir — e isso nao e' sempre o inicio do texto. No ODP derivado
        // (Task 6 do Plano 1, 2026-09-19) o cursor cai no FIM do primeiro paragrafo do slide, e o marcador ficou
        // no meio dos primeiros 200 caracteres: a amostra deixava de ser uma substring contigua do texto salvo e
        // a asserção caia sem nenhuma perda de conteudo. O que este teste precisa provar e' que o texto de origem
        // SOBREVIVEU, nao que o marcador foi parar antes ou depois dele — por isso a comparação e' feita com o
        // marcador removido (e os espacos renormalizados, porque a remocao pode deixar dois espacos seguidos).
        var textoDepoisSemMarcador = NormalizarEspacos(textoDepois.Replace(MarcadorNoTexto, " "));
        Assert.Contains(amostraOriginal, textoDepoisSemMarcador);

        // XLSX: a asserção por CELULA (mesma disciplina de SeisFormatosTests) — o texto integral nao distingue
        // "inseriu na celula escolhida" de "sobrescreveu outra coisa e o marcador sobrou em outro lugar".
        if (formato == FormatoDocumento.Xlsx)
            Assert.Contains(MarcadorNoTexto, FabricaDocumentos.TextoXlsxCelula(salvo, celula));

        // Criterio 3 (Task 8 do Plano 3): a edicao nao pode ter custado estrutura. Nenhuma contagem pode CAIR
        // (subir e' permitido: a celula editada passa a ter valor, o x2t pode acrescentar um paragrafo vazio).
        // A linha "RESUMO" e' o que o gate agrega por formato — so' rotulo e numeros, nunca nome nem trecho.
        var contagensDepois = ContagensEstruturais(salvo, caso.Extensao);
        var perdas = contagensAntes
            .Where(kv => contagensDepois.GetValueOrDefault(kv.Key) < kv.Value)
            .Select(kv => $"{kv.Key} {kv.Value} -> {contagensDepois.GetValueOrDefault(kv.Key)}")
            .ToList();
        _saida.WriteLine($"RESUMO [{rotulo}] {(caso.Derivada ? "(derivada) " : "")}contagens antes " +
                         $"{Descrever(contagensAntes)} | depois {Descrever(contagensDepois)} | " +
                         (perdas.Count == 0 ? "ok" : "COM PERDA: " + string.Join("; ", perdas)));
        Assert.True(perdas.Count == 0, $"[{rotulo}] contagens cairam depois de editar e salvar: {string.Join("; ", perdas)}");
    }

    private static string Descrever(IReadOnlyDictionary<string, int> c) => string.Join(", ", c.Select(kv => $"{kv.Key}={kv.Value}"));

    // -----------------------------------------------------------------------------------------------------
    // Contagens estruturais (Task 8 do Plano 3). OOXML: as MESMAS de tools/fidelidade.ps1 (Perfil-Docx/Xlsx/
    // Pptx) mais as imagens do PPTX; ODF: o equivalente no content.xml (fidelidade.ps1 so' compara texto no ODF).

    private const string NsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string NsS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string NsP = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private const string NsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private const string NsTable = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private const string NsDraw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

    /// <summary>Contagens estruturais do pacote (so' numeros). Nenhuma delas pode cair entre o original e o salvo.</summary>
    internal static IReadOnlyDictionary<string, int> ContagensEstruturais(string caminho, string ext)
    {
        using var zip = System.IO.Compression.ZipFile.OpenRead(caminho);
        System.Xml.XmlDocument? Ler(string entrada)
        {
            var e = zip.GetEntry(entrada);
            if (e is null) return null;
            using var fluxo = e.Open();
            var d = new System.Xml.XmlDocument { XmlResolver = null };
            d.Load(fluxo);
            return d;
        }
        // So' as partes diretas da pasta (xl/worksheets/sheet1.xml sim; xl/worksheets/_rels/... nao).
        List<System.Xml.XmlDocument> LerTodas(string prefixo) =>
            zip.Entries.Where(e => e.FullName.StartsWith(prefixo, StringComparison.OrdinalIgnoreCase) &&
                                   e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                                   !e.FullName[prefixo.Length..].Contains('/'))
                       .Select(e => Ler(e.FullName)!)
                       .ToList();
        static int Contar(System.Xml.XmlDocument? d, string ns, string nome) => d?.GetElementsByTagName(nome, ns).Count ?? 0;

        var c = new Dictionary<string, int>();
        switch (ext)
        {
            case ".docx":
            {
                var doc = Ler("word/document.xml");
                c["paragrafos"] = Contar(doc, NsW, "p");
                c["tabelas"] = Contar(doc, NsW, "tbl");
                c["imagens"] = Contar(doc, NsW, "drawing");
                break;
            }
            case ".xlsx":
            {
                var folhas = LerTodas("xl/worksheets/");
                c["planilhas"] = Contar(Ler("xl/workbook.xml"), NsS, "sheet");
                c["formulas"] = folhas.Sum(f => Contar(f, NsS, "f"));
                c["celulas"] = folhas.Sum(f => f.GetElementsByTagName("c", NsS).Cast<System.Xml.XmlElement>()
                    .Count(e => e.GetElementsByTagName("v", NsS).Count > 0 || e.GetElementsByTagName("is", NsS).Count > 0));
                break;
            }
            case ".pptx":
            {
                var slides = LerTodas("ppt/slides/");
                c["slides"] = slides.Count;
                c["formas"] = slides.Sum(sl => Contar(sl, NsP, "sp"));
                c["imagens"] = slides.Sum(sl => Contar(sl, NsP, "pic"));
                break;
            }
            default: // .odt / .ods / .odp: tudo no content.xml
            {
                var conteudo = Ler("content.xml");
                // Imagem = draw:frame cuja PRIMEIRA representacao e' um draw:image. As demais representacoes de
                // um frame sao substitutos: o LibreOffice grava uma imagem de pre-visualizacao atras de uma
                // tabela ou de um objeto OLE (frame = [table, image] / [object, image]) que o x2t nao regrava — a
                // tabela e o objeto continuam no arquivo salvo. Medido no real-odp derivado (Task 8): contar todo
                // draw:image acusava 2 imagens a menos so' por esses dois substitutos, sem nada sumir da tela.
                // ESCOPO da regra: ela foi calibrada no padrao OBSERVADO (substituto de pre-visualizacao do
                // LibreOffice atras de tabela/OLE). Um frame com DUAS imagens realmente distintas (a segunda nao
                // sendo substituto) conta como uma so' — a perda dessa segunda imagem passaria sem alarme.
                // Revisitar se aparecer amostra ODF real com frames de varias imagens (backlog 54).
                c["imagens"] = conteudo?.GetElementsByTagName("frame", NsDraw).Cast<System.Xml.XmlElement>()
                    .Count(f => f.ChildNodes.OfType<System.Xml.XmlElement>().FirstOrDefault() is { LocalName: "image", NamespaceURI: NsDraw }) ?? 0;
                if (ext == ".odt")
                {
                    c["paragrafos"] = Contar(conteudo, NsTextoOdf, "p") + Contar(conteudo, NsTextoOdf, "h");
                    c["tabelas"] = Contar(conteudo, NsTable, "table");
                }
                else if (ext == ".ods")
                {
                    // Planilhas = table:table filhas diretas de office:spreadsheet (uma tabela dentro de um desenho
                    // nao e' uma aba).
                    var planilha = conteudo?.GetElementsByTagName("spreadsheet", NsOffice).Cast<System.Xml.XmlElement>().FirstOrDefault();
                    c["planilhas"] = planilha?.ChildNodes.Cast<System.Xml.XmlNode>()
                        .Count(n => n.LocalName == "table" && n.NamespaceURI == NsTable) ?? 0;
                }
                else
                {
                    c["slides"] = Contar(conteudo, NsDraw, "page");
                }
                break;
            }
        }
        return c;
    }

    // -----------------------------------------------------------------------------------------------------

    /// <summary>Texto legivel do arquivo, por formato — os leitores existentes de <see cref="FabricaDocumentos"/>
    /// para OOXML (ja devolvem texto limpo) e <see cref="FabricaDocumentos.TextoOdf"/> para ODF. TextoOdf
    /// devolve o content.xml INTEIRO (estilos + corpo): pegar os primeiros ~200 caracteres DESSE texto pegaria
    /// so' cabecalho/estilos XML, que o x2t REGRAVA do zero a cada rodada (nomes de estilo mudam) — um
    /// "presente" ali nao provaria nada sobre o texto do documento, e uma falha seria um alarme falso, nao uma
    /// perda real. Por isso o texto VISIVEL (dentro de text:p/text:h) e' extraido do content.xml devolvido por
    /// TextoOdf — mesmo criterio de tools/fidelidade.ps1 (Texto-De-Odf) — antes de tirar a amostra.</summary>
    private static string TextoLegivel(string caminho, string ext) => ext switch
    {
        ".docx" => FabricaDocumentos.TextoDocx(caminho),
        ".xlsx" => FabricaDocumentos.TextoXlsx(caminho),
        ".pptx" => FabricaDocumentos.TextoPptx(caminho),
        _ => TextoVisivelOdf(FabricaDocumentos.TextoOdf(caminho)), // .odt / .ods / .odp
    };

    private const string NsTextoOdf = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    /// <summary>Texto visivel dos paragrafos (text:p/text:h) de um content.xml, respeitando a codificacao de
    /// espaco em branco do ODF: <c>&lt;text:s text:c="N"/&gt;</c> vale N espacos (1 se o atributo faltar),
    /// <c>&lt;text:tab/&gt;</c> e' uma tabulacao e <c>&lt;text:line-break/&gt;</c> uma quebra de linha.
    /// <para>Por que nao <c>InnerText</c>: esses tres elementos sao VAZIOS, e o InnerText os apaga. O LibreOffice
    /// (que gera o ODP derivado) e o x2t (que grava o arquivo salvo) escolhem representacoes DIFERENTES para o
    /// mesmo espaco — um grava o caractere literal, o outro grava <c>&lt;text:s/&gt;</c> — e com InnerText um
    /// espaco entre duas palavras "sumia" de um lado so' (medido na Task 6 do Plano 1, 2026-09-19: 679 vs 678
    /// caracteres; com esta extracao os dois textos sao identicos). Era o mesmo artefato que a fidelidade do
    /// Marco 0 registrou como "um espaco mudou de lugar na fronteira de runs" — nao ha' juncao de palavras no
    /// arquivo, so' na leitura.</para></summary>
    private static string TextoVisivelOdf(string contentXml)
    {
        var doc = new System.Xml.XmlDocument();
        doc.LoadXml(contentXml);
        var mgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
        mgr.AddNamespace("text", NsTextoOdf);
        var nos = doc.SelectNodes("//text:p | //text:h", mgr)!;
        var sb = new System.Text.StringBuilder();
        foreach (System.Xml.XmlNode n in nos) { AcumularTextoOdf(n, sb); sb.Append(' '); }
        return sb.ToString();
    }

    private static void AcumularTextoOdf(System.Xml.XmlNode no, System.Text.StringBuilder sb)
    {
        foreach (System.Xml.XmlNode filho in no.ChildNodes)
        {
            switch (filho.NodeType)
            {
                case System.Xml.XmlNodeType.Text:
                case System.Xml.XmlNodeType.CDATA:
                    sb.Append(filho.Value);
                    break;
                case System.Xml.XmlNodeType.Element when filho.NamespaceURI == NsTextoOdf && filho.LocalName == "s":
                    var c = (filho as System.Xml.XmlElement)?.GetAttribute("c", NsTextoOdf);
                    sb.Append(' ', int.TryParse(c, out var n) && n > 0 ? n : 1);
                    break;
                case System.Xml.XmlNodeType.Element when filho.NamespaceURI == NsTextoOdf && filho.LocalName == "tab":
                    sb.Append('\t');
                    break;
                case System.Xml.XmlNodeType.Element when filho.NamespaceURI == NsTextoOdf && filho.LocalName == "line-break":
                    sb.Append('\n');
                    break;
                case System.Xml.XmlNodeType.Element:
                    AcumularTextoOdf(filho, sb);
                    break;
            }
        }
    }

    private static string NormalizarEspacos(string s) => System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ").Trim();

    /// <summary>A primeira celula candidata que estiver vazia no arquivo ORIGINAL (nao no editor — mesma
    /// disciplina da leitura de linha de base). Planilhas reais de escritorio podem ir alem de Z1.</summary>
    private static string EscolherCelulaLivre(string caminho)
    {
        var ext = Path.GetExtension(caminho).ToLowerInvariant();
        if (ext != ".xlsx")
            // ODS: TextoXlsxCelula so' le OOXML. Sem um leitor de celula ODF dedicado, a primeira candidata e'
            // usada e a presenca do marcador depois de salvar e' o que prova a insercao (via texto integral).
            return CelulasCandidatas[0];
        foreach (var c in CelulasCandidatas)
            if (string.IsNullOrEmpty(FabricaDocumentos.TextoXlsxCelula(caminho, c)))
                return c;
        throw new InvalidOperationException($"nenhuma celula candidata livre encontrada ({string.Join(", ", CelulasCandidatas)}) — planilha maior do que o esperado");
    }
}
