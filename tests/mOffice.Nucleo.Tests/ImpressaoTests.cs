using System.Diagnostics;
using System.Printing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Web.WebView2.Core;
using mOffice.Conversor;
using mOffice.Nucleo.Impressao;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

public class ImpressaoTests
{
    [Fact] // opcoes do painel do editor viram o record; o que nao e' nativeOptions segue para o x2t
    public void Ler_OpcoesDoPainel()
    {
        var o = OpcoesImpressao.Ler("""{"nativeOptions":{"usesystemdialog":false,"printer":"HP LaserJet","pages":"1-3,5","copies":2,"sides":"both-long","paperSize":{"w":210,"h":297,"preset":"A4"},"paperOrientation":"landscape","colorMode":false,"quickPrint":false},"documentLayout":{"isPrint":true}}""");
        Assert.Equal("HP LaserJet", o.Impressora); Assert.Equal("1-3,5", o.Paginas); Assert.Equal(2, o.Copias);
        Assert.Equal("both-long", o.Lados); Assert.True(o.Paisagem); Assert.False(o.Colorido); Assert.False(o.Rapida);
        Assert.Equal((210d, 297d), (o.LarguraMm, o.AlturaMm));
        Assert.Equal("""{"documentLayout":{"isPrint":true}}""", o.JsonParaX2t);
    }

    [Fact] // JSON vazio, sem nativeOptions ou quebrado: padroes seguros (1 copia, A4 retrato, colorido, impressora padrao)
    public void Ler_Padroes()
    {
        foreach (var json in new[] { "", "{}", "{quebrado", """{"nativeOptions":null}""" })
        {
            var o = OpcoesImpressao.Ler(json);
            Assert.Null(o.Impressora); Assert.Equal(1, o.Copias); Assert.Equal("one", o.Lados);
            Assert.False(o.Paisagem); Assert.True(o.Colorido); Assert.Equal((210d, 297d), (o.LarguraMm, o.AlturaMm));
        }
    }

    [Fact] // impressao rapida do botao da barra: so' o flag, sem opcoes
    public void Ler_Rapida() => Assert.True(OpcoesImpressao.Ler("""{"nativeOptions":{"quickPrint":true}}""").Rapida);

    [Fact] // copias fora da faixa sao limitadas (o painel nao deixa, mas o JSON e' dado externo)
    public void Ler_CopiasLimitadas()
    {
        Assert.Equal(1, OpcoesImpressao.Ler("""{"nativeOptions":{"copies":0}}""").Copias);
        Assert.Equal(999, OpcoesImpressao.Ler("""{"nativeOptions":{"copies":100000}}""").Copias);
    }

    [Fact] // o JSON REAL do editor de texto 9.4 (forma medida em sdkjs/word NIe): "all" = todas, "auto" nao forca
           // orientacao, printer nulo = padrao, e o documentLayout inteiro segue para o x2t
    public void Ler_JsonDoEditorDeTexto()
    {
        var o = OpcoesImpressao.Ler("""{"nativeOptions":{"usesystemdialog":true,"printer":null,"colorMode":true,"pages":"all","paperSize":{"w":215.9,"h":279.4,"preset":"Letter"},"paperOrientation":"auto","copies":1,"sides":"one"},"printOptions":{"selection":1},"documentLayout":{"drawPlaceHolders":false,"drawFormHighlight":false,"isPrint":true}}""");
        Assert.True(o.DialogoDoSistema); Assert.Null(o.Impressora); Assert.Null(o.Paginas); Assert.True(o.Colorido);
        Assert.Equal("Letter", o.Papel); Assert.Equal((215.9, 279.4), (o.LarguraMm, o.AlturaMm));
        Assert.True(o.OrientacaoAutomatica); Assert.False(o.Paisagem);
        Assert.Equal("""{"printOptions":{"selection":1},"documentLayout":{"drawPlaceHolders":false,"drawFormHighlight":false,"isPrint":true}}""", o.JsonParaX2t);
    }

    [Fact] // planilha 9.4 manda SO' nativeOptions: nada sobra para o x2t (null, nao "{}"); lados desconhecido = "one"
    public void Ler_SoNativeOptions_SemJsonParaX2t()
    {
        var o = OpcoesImpressao.Ler("""{"nativeOptions":{"sides":"three","colorMode":0,"copies":"3"}}""");
        Assert.Null(o.JsonParaX2t); Assert.Equal("one", o.Lados); Assert.False(o.Colorido); Assert.Equal(3, o.Copias);
    }

    [Fact] // ChangesSalvos: dispara DEPOIS de o lote estar no disco (quem espera ja' le' o arquivo novo), tambem
           // com count 0, e NUNCA com a ponte congelada (nada foi gravado)
    public void ChangesSalvos_DepoisDeGravar_NuncaCongelada()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-changes-imp-").FullName;
        try
        {
            var sessao = new SessaoDocumento { ArquivoOrigem = "doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = pasta };
            var ponte = new PonteDesktop(sessao);
            var vistos = new List<string>();
            ponte.ChangesSalvos += () => vistos.Add(File.Exists(sessao.ArquivoChanges) ? File.ReadAllText(sessao.ArquivoChanges) : "");
            ponte.SalvarChanges("c1", 0, 1);
            ponte.SalvarChanges("", 1, 0);
            Assert.Equal(["\"c1\",", "\"c1\","], vistos);
            ponte.Congelar();
            Assert.True(ponte.Congelada);
            ponte.SalvarChanges("c2", 1, 1);
            Assert.Equal(2, vistos.Count);
            Assert.Equal("\"c1\",", File.ReadAllText(sessao.ArquivoChanges));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [Fact] // retencao: so' .pdf direto da pasta e mais velho que o limite; o resto fica
    public void PastaImpressao_Varrer_SoPdfsVelhos()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-varrer-").FullName;
        try
        {
            var velho = PastaImpressao.NovoArquivo(pasta); File.WriteAllText(velho, "x");
            File.SetLastWriteTimeUtc(velho, DateTime.UtcNow.AddDays(-2));
            var novo = PastaImpressao.NovoArquivo(pasta); File.WriteAllText(novo, "x");
            var outro = Path.Combine(pasta, "nota.txt"); File.WriteAllText(outro, "x");
            File.SetLastWriteTimeUtc(outro, DateTime.UtcNow.AddDays(-2));
            var sub = Directory.CreateDirectory(Path.Combine(pasta, "sub")).FullName;
            var velhoNaSub = Path.Combine(sub, "a.pdf"); File.WriteAllText(velhoNaSub, "x");
            File.SetLastWriteTimeUtc(velhoNaSub, DateTime.UtcNow.AddDays(-2));

            Assert.Equal(1, PastaImpressao.Varrer(pasta, PastaImpressao.IdadeMinima));
            Assert.False(File.Exists(velho));
            Assert.True(File.Exists(novo)); Assert.True(File.Exists(outro)); Assert.True(File.Exists(velhoNaSub));
            Assert.Equal(0, PastaImpressao.Varrer(Path.Combine(pasta, "nao-existe"), PastaImpressao.IdadeMinima));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [Fact] // fix round 1: "nunca lanca" vale tambem para chave duplicada (o indexador do JsonObject lanca
           // ArgumentException) — padroes, como qualquer JSON invalido
    public void Ler_ChaveDuplicada_Padroes()
    {
        foreach (var json in new[] { """{"nativeOptions":{"copies":2,"copies":3}}""", """{"a":1,"a":2}""", """{"nativeOptions":{"paperSize":{"w":1,"w":2}}}""" })
            Assert.Same(OpcoesImpressao.Padrao, OpcoesImpressao.Ler(json));
    }

    [Fact] // a pasta dos PDFs de impressao mora nos dados por usuario (spec: nada em %TEMP%), nao na recuperacao
    public void Pastas_Impressao_SobARaiz()
        => Assert.Equal(Path.Combine(Pastas.Raiz, "impressao"), Pastas.Impressao);

    // ---- Task 5: impressoras no painel e envio ----

    [Fact] // JSON do painel com a impressora padrao e os papeis em mm
    public void JsonParaEditor_Formato()
    {
        var json = ImpressorasDoWindows.JsonParaEditor("HP", [new ImpressoraInfo("HP", true, false, [new PapelImpressora("A4", 210, 297)])]);
        var o = System.Text.Json.Nodes.JsonNode.Parse(json)!;
        Assert.Equal("HP", (string?)o["current_printer"]);
        var p = o["printers"]![0]!;
        Assert.Equal("HP", (string?)p["name"]); Assert.True((bool)p["duplex_supported"]!); Assert.False((bool)p["color_supported"]!);
        Assert.Equal(210, (double)p["paper_supported"]![0]!["width"]!);
        Assert.Equal(297, (double)p["paper_supported"]![0]!["height"]!);
        Assert.Equal("A4", (string?)p["paper_supported"]![0]!["name"]);
    }

    [Fact] // nome de impressora com aspas e barra invertida nao quebra o JSON
    public void JsonParaEditor_EscapaNomes() =>
        Assert.Equal("Sala \"2\" \\ fila", (string?)System.Text.Json.Nodes.JsonNode.Parse(
            ImpressorasDoWindows.JsonParaEditor(null, [new ImpressoraInfo("Sala \"2\" \\ fila", false, true, [])]))!["printers"]![0]!["name"]);

    [Fact] // sem impressora padrao a chave current_printer nao vai (o web-apps mantem o que tinha); lista vazia = []
    public void JsonParaEditor_SemPadrao_SemImpressoras()
    {
        var o = System.Text.Json.Nodes.JsonNode.Parse(ImpressorasDoWindows.JsonParaEditor(null, []))!.AsObject();
        Assert.False(o.ContainsKey("current_printer"));
        Assert.Empty(o["printers"]!.AsArray());
    }

    [SkippableFact] // na maquina de teste existe ao menos "Microsoft Print to PDF" e a lista le sem excecao
    public void Listar_MaquinaReal()
    {
        var (padrao, lista) = ImpressorasDoWindows.Listar();
        Skip.If(lista.Count == 0, "nenhuma impressora instalada nesta maquina");
        Assert.All(lista, i => Assert.False(string.IsNullOrWhiteSpace(i.Nome)));
        Assert.Equal(lista.Count, lista.Select(i => i.Nome).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.All(lista.SelectMany(i => i.Papeis), p => Assert.True(p.LarguraMm > 0 && p.AlturaMm > 0));
        if (padrao is not null) Assert.Contains(lista, i => i.Nome == padrao);
        // A fila virtual do Windows (quando existe) tem A4 com 210 x 297 mm: prova a conversao 1/96 pol -> mm.
        if (lista.FirstOrDefault(i => i.Nome == "Microsoft Print to PDF") is { } pdf)
            Assert.Contains(pdf.Papeis, p => p.Nome == "A4" && Math.Abs(p.LarguraMm - 210) < 0.2 && Math.Abs(p.AlturaMm - 297) < 0.2);
    }

    [Fact] // nomes do System.Printing ficam legiveis no combo do painel
    public void NomeDoPapel_Legivel()
    {
        Assert.Equal("A4", ImpressorasDoWindows.NomeDoPapel(PageMediaSizeName.ISOA4));
        Assert.Equal("A3", ImpressorasDoWindows.NomeDoPapel(PageMediaSizeName.ISOA3));
        Assert.Equal("Letter", ImpressorasDoWindows.NomeDoPapel(PageMediaSizeName.NorthAmericaLetter));
        Assert.Equal("Legal", ImpressorasDoWindows.NomeDoPapel(PageMediaSizeName.NorthAmericaLegal));
        Assert.Equal("JIS B5", ImpressorasDoWindows.NomeDoPapel(PageMediaSizeName.JISB5));
        Assert.Equal("ISOA4Rotated", ImpressorasDoWindows.NomeDoPapel(PageMediaSizeName.ISOA4Rotated));
    }

    [Fact] // opcoes do painel -> PrintAsync: intervalo normalizado, copias, frente e verso, cinza, paisagem, papel
           // SEMPRE retrato em polegadas (o painel manda w>h quando a pagina e' paisagem)
    public void Montar_OpcoesDoPainel()
    {
        var o = OpcoesImpressao.Ler("""{"nativeOptions":{"printer":"HP","pages":" 1 - 3 , 5 ","copies":2,"sides":"both-long","paperSize":{"w":297,"h":210,"preset":"A4"},"paperOrientation":"landscape","colorMode":false}}""");
        var c = Impressora.Montar(o, "HP", (8.5, 11));
        Assert.Equal("HP", c.Impressora); Assert.Equal("1-3,5", c.Paginas); Assert.Equal(2, c.Copias);
        Assert.Equal(CoreWebView2PrintDuplex.TwoSidedLongEdge, c.Duplex);
        Assert.Equal(CoreWebView2PrintColorMode.Grayscale, c.Cor);
        Assert.Equal(CoreWebView2PrintOrientation.Landscape, c.Orientacao);
        Assert.Equal(210 / 25.4, c.LarguraPol!.Value, 3); Assert.Equal(297 / 25.4, c.AlturaPol!.Value, 3);

        var curto = Impressora.Montar(OpcoesImpressao.Ler("""{"nativeOptions":{"sides":"both-short","paperOrientation":"portrait","pages":"all"}}"""), "X", null);
        Assert.Equal(CoreWebView2PrintDuplex.TwoSidedShortEdge, curto.Duplex);
        Assert.Equal(CoreWebView2PrintOrientation.Portrait, curto.Orientacao);
        Assert.Null(curto.Paginas);
        Assert.Equal(CoreWebView2PrintColorMode.Color, curto.Cor);
    }

    [Fact] // "auto" do painel nao forca orientacao (cada pagina na dela)
    public void Montar_OrientacaoAutomatica_NaoForca() =>
        Assert.Null(Impressora.Montar(OpcoesImpressao.Ler("""{"nativeOptions":{"paperOrientation":"auto"}}"""), "X", null).Orientacao);

    [Fact] // impressao rapida: so' a impressora padrao, com os padroes DELA (papel dela, lados e cor Default), 1 copia, tudo
    public void Montar_Rapida_PadroesDaImpressora()
    {
        var o = OpcoesImpressao.Ler("""{"nativeOptions":{"quickPrint":true}}""") with { Copias = 5, Paginas = "2" };
        var c = Impressora.Montar(o, "Padrao", (8.27, 11.69));
        Assert.Equal("Padrao", c.Impressora); Assert.Null(c.Paginas); Assert.Equal(1, c.Copias);
        Assert.Equal(CoreWebView2PrintDuplex.Default, c.Duplex); Assert.Equal(CoreWebView2PrintColorMode.Default, c.Cor);
        Assert.Null(c.Orientacao); Assert.Equal((8.27, 11.69), (c.LarguraPol!.Value, c.AlturaPol!.Value));
    }

    [Theory] // o painel deixa digitar qualquer coisa: formato invalido vira mensagem pt-BR (nao ArgumentException do WebView2)
    [InlineData("abc")]
    [InlineData("1-")]
    [InlineData("1,,2")]
    [InlineData("-3")]
    [InlineData("1;2")]
    [InlineData("0")]   // MEDIDO: "0" trava o PrintToPdfAsync do WebView2
    [InlineData("1,0-2")]
    [InlineData("3-2")] // MEDIDO: intervalo invertido falha no WebView2
    public void NormalizarPaginas_Invalido_MensagemPtBr(string paginas)
    {
        var e = Assert.Throws<InvalidOperationException>(() => Impressora.NormalizarPaginas(paginas));
        Assert.StartsWith("Intervalo de páginas inválido", e.Message);
    }

    [Fact] // o que o usuario escolheu no dialogo Imprimir do Windows (bilhete + intervalo) -> PrintAsync
    public void DoDialogo_Bilhete()
    {
        var t = new PrintTicket
        {
            CopyCount = 3, Duplexing = Duplexing.TwoSidedShortEdge, OutputColor = OutputColor.Monochrome,
            PageOrientation = PageOrientation.Landscape, PageMediaSize = new PageMediaSize(PageMediaSizeName.ISOA4, 793.7, 1122.5),
        };
        var c = Impressora.DoDialogo("Sala 2", t, new System.Windows.Controls.PageRange(2, 4));
        Assert.Equal("Sala 2", c.Impressora); Assert.Equal("2-4", c.Paginas); Assert.Equal(3, c.Copias);
        Assert.Equal(CoreWebView2PrintDuplex.TwoSidedShortEdge, c.Duplex);
        Assert.Equal(CoreWebView2PrintColorMode.Grayscale, c.Cor);
        Assert.Equal(CoreWebView2PrintOrientation.Landscape, c.Orientacao);
        Assert.Equal(793.7 / 96, c.LarguraPol!.Value, 3);

        var vazio = Impressora.DoDialogo("X", new PrintTicket(), null);
        Assert.Null(vazio.Paginas); Assert.Equal(1, vazio.Copias);
        Assert.Equal(CoreWebView2PrintDuplex.Default, vazio.Duplex); Assert.Equal(CoreWebView2PrintColorMode.Default, vazio.Cor);
        Assert.Null(vazio.Orientacao); Assert.Null(vazio.LarguraPol);
        Assert.Equal("5", Impressora.DoDialogo("X", new PrintTicket(), new System.Windows.Controls.PageRange(5, 5)).Paginas);
    }

    [Fact] // impressora que nao existe: InvalidOperationException pt-BR COM o nome, antes de qualquer WebView2 ou fila
    public void Imprimir_ImpressoraInexistente_ErroComONome()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-inexistente-").FullName;
        try
        {
            var pdf = Path.Combine(pasta, "x.pdf"); File.WriteAllText(pdf, "%PDF-1.4");
            var nome = "Impressora que nao existe " + Guid.NewGuid().ToString("N")[..6];
            var o = OpcoesImpressao.Ler($$$"""{"nativeOptions":{"printer":"{{{nome}}}"}}""");
            InvalidOperationException? erro = null;
            JanelaSta.Executar(async () =>
            {
                erro = await Assert.ThrowsAsync<InvalidOperationException>(() => new Impressora(JanelaEditor.AppPadrao).ImprimirAsync(pdf, o));
            }, TimeSpan.FromSeconds(60));
            Assert.Equal($"A impressora \"{nome}\" não foi encontrada no Windows.", erro!.Message);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    // ---- Task 5, fix round 1 ----

    private const string UrlPdf = "file:///C:/Users/Jo%C3%A3o/AppData/Local/mOffice/impressao/abc.pdf";

    private static string Alvos(string tituloDoVisualizador, string urlDaPagina = UrlPdf, string pai = "P1") => $$"""
        {"targetInfos":[
          {"targetId":"P1","type":"page","url":"{{urlDaPagina}}","title":"abc.pdf"},
          {"targetId":"W1","type":"webview","parentId":"{{pai}}","url":"chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/edge_pdf/index.html","title":"{{tituloDoVisualizador}}"},
          {"targetId":"I1","type":"iframe","parentId":"W1","url":"{{urlDaPagina}}","title":""}
        ]}
        """;

    [Fact] // sinal do visualizador (medido no runtime fixo): o alvo "webview" filho da pagina DESTE PDF so' conta
           // quando o titulo deixa de ser vazio/url da extensao e vira o titulo do documento
    public void VisualizadorCarregou_SoComOTituloDoDocumento()
    {
        Assert.False(Impressora.VisualizadorCarregou(Alvos(""), UrlPdf));
        Assert.False(Impressora.VisualizadorCarregou(Alvos("chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/edge_pdf/index.html"), UrlPdf));
        Assert.False(Impressora.VisualizadorCarregou(Alvos(UrlPdf), UrlPdf));
        Assert.True(Impressora.VisualizadorCarregou(Alvos("abc.pdf"), UrlPdf));
        // a mesma url, escrita com e sem escape (o usuario "João"), e' o mesmo PDF
        Assert.True(Impressora.VisualizadorCarregou(Alvos("abc.pdf"), "file:///C:/Users/João/AppData/Local/mOffice/impressao/abc.pdf"));
    }

    [Fact] // outro PDF (outra impressao simultanea) ou um visualizador de outra pagina nao contam; JSON estranho = nao
    public void VisualizadorCarregou_SoOVisualizadorDestePdf()
    {
        Assert.False(Impressora.VisualizadorCarregou(Alvos("xyz.pdf", "file:///C:/outro/xyz.pdf"), UrlPdf));
        Assert.False(Impressora.VisualizadorCarregou(Alvos("abc.pdf", pai: "OUTRA"), UrlPdf));
        foreach (var lixo in new[] { "", "{}", "[]", "{\"targetInfos\":5}", "{\"targetInfos\":[1,\"x\",null]}", "{quebrado" })
            Assert.False(Impressora.VisualizadorCarregou(lixo, UrlPdf));
    }

    [Fact] // limite do visualizador proporcional ao tamanho, com teto
    public void LimiteVisualizador_ProporcionalComTeto()
    {
        Assert.Equal(TimeSpan.FromSeconds(15), Impressora.LimiteVisualizador(0));
        Assert.Equal(TimeSpan.FromSeconds(35), Impressora.LimiteVisualizador(10L * 1024 * 1024));
        Assert.Equal(TimeSpan.FromSeconds(90), Impressora.LimiteVisualizador(10L * 1024 * 1024 * 1024));
    }

    [Fact] // uma fila cujo driver lanca QUALQUER excecao e' pulada (com log); as outras ficam; todas sao descartadas
    public void Coletar_FilaQuebradaNaoDerrubaALista()
    {
        var descartadas = new List<string>();
        var log = new List<string>();
        var lista = ImpressorasDoWindows.Coletar(["A", "quebrada-nome", "quebrada-driver", "B", "a", ""],
            nome: f => f == "quebrada-nome" ? throw new NullReferenceException("driver") : f,
            capacidades: (f, n) => f == "quebrada-driver" ? throw new AccessViolationException("driver") : new ImpressoraInfo(n, false, true, []),
            descartar: f => { descartadas.Add(f); if (f == "B") throw new ObjectDisposedException("x"); },
            log.Add);
        Assert.Equal(["A", "B"], lista.Select(i => i.Nome));
        Assert.Equal(6, descartadas.Count);
        Assert.Equal(3, log.Count); // duas filas puladas + um descarte que falhou
    }

    [Fact] // o PDF que o navegador ainda segura por um instante e' apagado assim que ele solta; ausente = ok
    public async Task ApagarComEspera_EsperaSoltarOArquivo()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-apagar-").FullName;
        try
        {
            var pdf = Path.Combine(pasta, "a.pdf"); File.WriteAllText(pdf, "x");
            var preso = new FileStream(pdf, FileMode.Open, FileAccess.Read, FileShare.Read);
            var soltar = Task.Run(async () => { await Task.Delay(700); preso.Dispose(); });
            Assert.True(await PastaImpressao.ApagarComEsperaAsync(pdf, tentativas: 20, intervaloMs: 100));
            await soltar;
            Assert.False(File.Exists(pdf));
            Assert.True(await PastaImpressao.ApagarComEsperaAsync(Path.Combine(pasta, "nao-existe.pdf")));

            File.WriteAllText(pdf, "x");
            using (new FileStream(pdf, FileMode.Open, FileAccess.Read, FileShare.Read))
                Assert.False(await PastaImpressao.ApagarComEsperaAsync(pdf, tentativas: 2, intervaloMs: 50));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}

/// <summary>PDF de impressao na janela real (Task 4 do Plano 3): o que o editor mostra — inclusive o que ainda
/// nao foi salvo — e' o que vai para o PDF, sem salvar o documento.</summary>
[Collection(JanelaReal.Nome)]
public class ImpressaoJanelaTests(ITestOutputHelper saida)
{
    /// <summary>Cria a janela, abre <paramref name="arquivo"/>, espera o editor e roda o corpo. A janela FECHA no
    /// <c>finally</c> (fix round 1): uma asserção que falhe no meio nao deixa janela, trava de recuperacao nem PDF
    /// de impressao para tras.</summary>
    private static async Task NaJanela(string app, string arquivo, Func<JanelaEditor, Task> corpo, Action<JanelaEditor>? preparar = null)
    {
        var j = EditorReal.NovaJanela(app);
        var fechada = false;
        j.Closed += (_, _) => fechada = true;
        preparar?.Invoke(j);
        j.Show();
        try
        {
            await j.AbrirAsync(arquivo);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            await corpo(j);
        }
        finally { if (!fechada) j.Close(); }
    }

    /// <summary>O botao Imprimir do painel, pelo caminho REAL (web-apps/apps/documenteditor/main/app.js): um
    /// asc_CAdjustPrint com asc_setNativeOptions dentro de um asc_CDownloadOptions e api.asc_Print(o). E' o sdkjs
    /// (NIe, sdkjs/word) quem monta o JSON e chama AscDesktopEditor.Print — nao o teste.</summary>
    private static readonly string ImprimirPeloPainel = EditorReal.NoEditor("""
        var a = new w.Asc.asc_CAdjustPrint();
        a.asc_setNativeOptions({usesystemdialog: false, printer: null, colorMode: true, pages: "all",
          paperSize: {w: 210, h: 297, preset: "A4"}, paperOrientation: "portrait", copies: 2, sides: "one"});
        var o = new w.Asc.asc_CDownloadOptions();
        o.asc_setAdvancedOptions(a);
        api.asc_Print(o);
        return "ok";
        """);

    private static HashSet<string> PdfsDeImpressao() =>
        Directory.Exists(Pastas.Impressao) ? [.. Directory.GetFiles(Pastas.Impressao, "*.pdf")] : [];

    [SkippableFact] // o texto digitado (ainda nao salvo) aparece no PDF de impressao, e o documento continua modificado
    public void PdfDeImpressao_TemOTextoNaoSalvo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Base");
        var docxAntes = File.GetLastWriteTimeUtc(docx);
        string? pdf = null;
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" impresso-777"));
                pdf = await j.GerarPdfDeImpressaoAsync(OpcoesImpressao.Ler("{}"));
                Assert.Contains("impresso-777", mOffice.Conversor.Tests.TextoDePdf.Extrair(pdf));
                Assert.True(j.Modificado);
                Assert.Equal(Pastas.Impressao, Path.GetDirectoryName(pdf));
            }), TimeSpan.FromSeconds(180));
            Assert.Equal(docxAntes, File.GetLastWriteTimeUtc(docx)); // nao salvou o documento do usuario
            Assert.False(File.Exists(pdf)); // a janela apagou o PDF dela ao fechar
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // o Print da ponte (o caminho do botao Imprimir do painel) termina no evento ImprimirPedido com o
                    // PDF pronto, as opcoes lidas e o resto do JSON indo para o x2t
    public void PrintDoEditor_DisparaImprimirPedidoComOPdf()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-print-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Documento para imprimir");
        string? entregue = null;
        try
        {
            var pedido = new TaskCompletionSource<(OpcoesImpressao, string)>(TaskCreationOptions.RunContinuationsAsynchronously);
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" via-print-555"));
                Assert.Equal("\"ok\"", await j.ExecutarJsAsync(ImprimirPeloPainel));
                var (opcoes, arquivo) = await pedido.Task.WaitAsync(TimeSpan.FromSeconds(60));
                entregue = arquivo; // Task 5: com assinante, o PDF e' de quem assina (aqui, o teste) — apagado no finally
                Assert.Equal(2, opcoes.Copias);
                Assert.Equal("""{"documentLayout":{"drawPlaceHolders":false,"drawFormHighlight":false,"isPrint":true}}""", opcoes.JsonParaX2t);
                Assert.Contains("via-print-555", mOffice.Conversor.Tests.TextoDePdf.Extrair(arquivo));
                Assert.True(j.Modificado);
            }, j =>
            {
                j.ImprimirPedido += (_, o, arquivo) => pedido.TrySetResult((o, arquivo));
                j.ErroAoImprimir += (_, m) => pedido.TrySetException(new InvalidOperationException(m));
            }), TimeSpan.FromSeconds(180));
        }
        finally
        {
            if (entregue is not null) PastaImpressao.Apagar(entregue);
            try { Directory.Delete(pasta, true); } catch { }
        }
    }

    [SkippableFact] // fix round 1: a impressao que TERMINA depois de a janela fechar apaga o PDF na hora e nao avisa
                    // ninguem (nem ImprimirPedido, nem ErroAoImprimir) — a janela ja' nao existe para o usuario
    public void Print_JanelaFechaAntesDeOPdfFicarPronto_NaoAvisaENaoDeixaPdf()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-fechou-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Documento que fecha no meio");
        var antes = PdfsDeImpressao();
        var avisos = new List<string>();
        var passouPelaCostura = false;
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" fecha-no-meio"));
                Assert.Equal("\"ok\"", await j.ExecutarJsAsync(ImprimirPeloPainel));
                var impressao = j.ImpressaoEmCursoParaTeste;
                Assert.NotNull(impressao);
                await impressao!.WaitAsync(TimeSpan.FromSeconds(60)); // nunca lanca: tudo vira log/evento
            }, j =>
            {
                j.ImprimirPedido += (_, _, arquivo) => avisos.Add("ImprimirPedido " + arquivo);
                j.ErroAoImprimir += (_, m) => avisos.Add("ErroAoImprimir " + m);
                // O x2t JA' gerou o PDF; a janela fecha (de verdade: documento modificado passa pela pergunta, que
                // os testes respondem "nao salvar") exatamente antes de ele ser registrado.
                j.AntesDeRegistrarPdfParaTeste = async () =>
                {
                    passouPelaCostura = true;
                    var fechou = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                    j.Closed += (_, _) => fechou.TrySetResult();
                    j.Close();
                    await fechou.Task.WaitAsync(TimeSpan.FromSeconds(30));
                };
            }), TimeSpan.FromSeconds(180));
            Assert.True(passouPelaCostura, "o PDF nao chegou a ser gerado (teste inconclusivo)");
            Assert.Empty(avisos);
            Assert.Empty(PdfsDeImpressao().Except(antes));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableTheory] // a descarga acha o "Have_Changes" do historico pelo CONTEUDO nos TRES bundles (o nome minificado
                      // e' l5/r8/Via): editar e descarregar grava os changes na hora, sem salvar e sem esperar o autosave
    [InlineData("mtexto", "docx")]
    [InlineData("mplanilha", "xlsx")]
    [InlineData("mslides", "pptx")]
    public void Descarga_TresEditores_GravaNaHora(string app, string tipo)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-tres-").FullName;
        var arquivo = tipo switch
        {
            "xlsx" => mOffice.Conversor.Tests.FabricaDocumentos.Xlsx(pasta, "Base"),
            "pptx" => mOffice.Conversor.Tests.FabricaDocumentos.Pptx(pasta, "Base"),
            _ => mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Base"),
        };
        var antes = File.GetLastWriteTimeUtc(arquivo);
        try
        {
            JanelaSta.Executar(() => NaJanela(app, arquivo, async j =>
            {
                var changes = j.SessaoAtual!.ArquivoChanges;
                // Nada pendente logo depois de abrir.
                var relogio = Stopwatch.StartNew();
                await j.DescarregarChangesAsync();
                Assert.True(relogio.Elapsed < TimeSpan.FromSeconds(2));
                Assert.False(File.Exists(changes) && new FileInfo(changes).Length > 0);

                var r = (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(j.SessaoAtual.Formato.DocType(), " descarga-" + tipo))).Trim('"');
                Assert.StartsWith("ok", r);
                // A colagem de planilha/apresentacao e' assincrona: espera a edicao chegar ao modelo (poucos
                // centesimos de segundo) e descarrega LOGO — antes dos ~2 s do autosave (medido), entao o arquivo
                // ainda esta' vazio e quem grava e' a descarga.
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                Assert.False(File.Exists(changes) && new FileInfo(changes).Length > 0, "o autosave gravou antes da descarga (teste inconclusivo)");
                relogio.Restart();
                await j.DescarregarChangesAsync();
                var gravouAntesDoAutosave = relogio.Elapsed < TimeSpan.FromSeconds(1.5);
                saida.WriteLine($"{tipo}: descarga depois da edicao em {relogio.ElapsedMilliseconds} ms, changes0.json = {(File.Exists(changes) ? new FileInfo(changes).Length : -1)} bytes");
                Assert.True(File.Exists(changes) && new FileInfo(changes).Length > 0, "a descarga nao gravou os changes");
                Assert.True(gravouAntesDoAutosave, "a descarga demorou como se tivesse esperado o autosave");
                Assert.True(j.Modificado);
            }), TimeSpan.FromSeconds(180));
            Assert.Equal(antes, File.GetLastWriteTimeUtc(arquivo));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // sem nada pendente a descarga volta NA HORA (nao espera evento nenhum); e o autosave do proprio
                    // editor grava os changes sozinho bem antes dos 10 s (a recuperacao depende disso) — medido
    public void Descarga_SemPendente_VoltaNaHora_EAutosaveGravaSozinho()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-descarga-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Base");
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                var relogio = Stopwatch.StartNew();
                await j.DescarregarChangesAsync();
                saida.WriteLine($"descarga sem nada pendente: {relogio.ElapsedMilliseconds} ms");
                Assert.True(relogio.Elapsed < TimeSpan.FromSeconds(2), "sem pendencia a descarga nao pode esperar");
                var changes = j.SessaoAtual!.ArquivoChanges;
                Assert.False(File.Exists(changes) && new FileInfo(changes).Length > 0);

                // Autosave: digita e NAO pede nada — mede quanto o editor leva para gravar sozinho.
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" autosave-1"));
                relogio.Restart();
                while (relogio.Elapsed < TimeSpan.FromSeconds(15) && !(File.Exists(changes) && new FileInfo(changes).Length > 0))
                    await Task.Delay(50);
                saida.WriteLine($"autosave do editor: changes no disco {relogio.ElapsedMilliseconds} ms depois da edicao");
                Assert.True(relogio.Elapsed < TimeSpan.FromSeconds(10), "o autosave do editor passou de 10 s");

                // Depois do autosave nao sobra nada: a descarga volta na hora de novo.
                relogio.Restart();
                await j.DescarregarChangesAsync();
                saida.WriteLine($"descarga depois do autosave: {relogio.ElapsedMilliseconds} ms");
                Assert.True(relogio.Elapsed < TimeSpan.FromSeconds(2));
            }), TimeSpan.FromSeconds(180));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // fix round 1: fechar a janela do documento NO MEIO do envio nao mata o envio — a WebView2 oculta
                    // nao tem dono (nao morre com a janela), a thread de interface segue viva (no app, o encerramento
                    // espera as impressoes: ImpressoesEmCurso) e o PDF, que ja' e' de quem imprime, nao e' apagado
                    // pela janela
    public void PrintDoEditor_JanelaFechaNoMeioDoEnvio_ImpressaoTermina()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-fecha-envio-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Documento que fecha durante o envio");
        var destino = Path.Combine(pasta, "saida-da-impressora.pdf");
        string? pdf = null;
        ImpressaoEnviada? resultado = null;
        var noEnvio = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var portao = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        Task<ImpressaoEnviada?>? envio = null;
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" fecha-no-envio-321"));
                Assert.Equal("\"ok\"", await j.ExecutarJsAsync(ImprimirPeloPainel));
                await noEnvio.Task.WaitAsync(TimeSpan.FromSeconds(60)); // a WebView2 oculta ja' esta' no PDF, enviando
                var fechou = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                j.Closed += (_, _) => fechou.TrySetResult();
                j.Close();
                await fechou.Task.WaitAsync(TimeSpan.FromSeconds(30));
                Assert.True(File.Exists(pdf), "a janela apagou o PDF que ja' era de quem imprime");
                portao.TrySetResult();
                resultado = await envio!.WaitAsync(TimeSpan.FromSeconds(60));
            }, j =>
            {
                j.ImprimirPedido += (_, o, arquivo) =>
                {
                    pdf = arquivo;
                    envio = ImpressaoSondaTests.SemFila(destino, antes: async () => { noEnvio.TrySetResult(); await portao.Task; }).ImprimirAsync(arquivo, o);
                };
                j.ErroAoImprimir += (_, m) => noEnvio.TrySetException(new InvalidOperationException(m));
            }), TimeSpan.FromSeconds(180));
            Assert.NotNull(resultado);
            Assert.Contains("fecha-no-envio-321", mOffice.Conversor.Tests.TextoDePdf.Extrair(destino));
        }
        finally
        {
            if (pdf is not null) PastaImpressao.Apagar(pdf);
            try { Directory.Delete(pasta, true); } catch { }
        }
    }

    /// <summary>Estado do combo "Impressora" do painel, lido no CONTROLADOR do web-apps (documenteditor:
    /// <c>DE.getController("Print")._state</c>, onde <c>setPrintersInfo</c> guarda o que chegou por
    /// <c>on_native_message("printer:config", …)</c>).</summary>
    private const string ImpressorasDoPainel = """
        (function () {
          try {
            var w = window.frames[0];
            var c = w.DE && w.DE.getController("Print");
            if (!c || !c._state) return "sem controlador Print";
            return JSON.stringify({ atual: c._state.currentPrinter, nomes: (c._state.printersList || []).map(function (p) { return p.name; }),
              papeis: (c._state.printersList || []).map(function (p) { return (p.paper_supported || []).length; }) });
          } catch (e) { return "erro: " + e; }
        })()
        """;

    [SkippableFact] // Task 5: as impressoras do Windows chegam ao painel Arquivo -> Imprimir (printer:config empurrado
                    // pelo host depois do EndLoad), com a padrao selecionada
    public void PrinterConfig_ChegaAoPainelDoEditor()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (padrao, lista) = ImpressorasDoWindows.Listar();
        Skip.If(lista.Count == 0, "nenhuma impressora instalada nesta máquina");
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-painel-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Base");
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                var envio = j.ImpressorasEnviadasParaTeste;
                Assert.NotNull(envio);
                await envio!.WaitAsync(TimeSpan.FromSeconds(60));
                var json = System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(ImpressorasDoPainel))!;
                saida.WriteLine("painel: " + json);
                var estado = System.Text.Json.Nodes.JsonNode.Parse(json)!;
                var nomes = estado["nomes"]!.AsArray().Select(n => (string)n!).ToList();
                Assert.Equal(lista.Select(i => i.Nome).OrderBy(n => n), nomes.OrderBy(n => n));
                Assert.Equal(padrao, (string?)estado["atual"]);
                // Os papeis tambem chegaram (a fila virtual do Windows tem varios).
                Assert.Equal(lista.Select(i => i.Papeis.Count), estado["papeis"]!.AsArray().Select(n => (int)n!));
            }), TimeSpan.FromSeconds(180));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableTheory] // Task 5, dono do PDF: com assinante de ImprimirPedido o PDF passa a ser DELE (a janela fecha e
                      // o PDF continua la' para a impressora); sem assinante continua da janela (apagado ao fechar)
    [InlineData(true)]
    [InlineData(false)]
    public void PrintDoEditor_DonoDoPdf(bool comAssinante)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-dono-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Documento do dono do PDF");
        var antes = PdfsDeImpressao();
        string? entregue = null;
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                Assert.Equal("\"ok\"", await j.ExecutarJsAsync(ImprimirPeloPainel));
                var impressao = j.ImpressaoEmCursoParaTeste;
                Assert.NotNull(impressao);
                await impressao!.WaitAsync(TimeSpan.FromSeconds(60));
                Assert.Single(PdfsDeImpressao().Except(antes)); // o PDF foi gerado nos dois casos
                if (comAssinante) Assert.NotNull(entregue);
            }, j =>
            {
                if (comAssinante) j.ImprimirPedido += (_, _, arquivo) => entregue = arquivo;
                j.ErroAoImprimir += (_, m) => throw new InvalidOperationException(m);
            }), TimeSpan.FromSeconds(180));
            // A janela ja' fechou (NaJanela).
            var sobrou = PdfsDeImpressao().Except(antes).ToList();
            if (comAssinante) Assert.Equal([entregue!], sobrou); // de quem imprime: a janela nao apagou
            else Assert.Empty(sobrou);                            // da janela: apagado ao fechar
        }
        finally
        {
            if (entregue is not null) PastaImpressao.Apagar(entregue);
            try { Directory.Delete(pasta, true); } catch { }
        }
    }
}

/// <summary>
/// Sonda S5 (Task 4 do Plano 3) e o envio de verdade (Task 5): COMO o PDF de impressao chega a' impressora. A
/// WebView2 OCULTA e' a de producao (<see cref="Impressora.ComPdfOculto"/>): MESMO ambiente (runtime fixo, mesmas
/// flags) e MESMA guarda de egresso da janela, com o PDF local no visualizador embutido.
///
/// Partes: <see cref="S5_PrintToPdf_PaginasCertasComIntervalo"/> nao toca em fila nem em dialogo nenhum e roda em
/// toda suite; as demais mandam trabalho para a fila VIRTUAL "Microsoft Print to PDF" ou abrem dialogos do
/// Windows — so' com <c>MOFFICE_SONDA_IMPRESSAO=1</c>. Nenhum teste usa impressora fisica.
/// </summary>
[Collection(JanelaReal.Nome)]
public class ImpressaoSondaTests(ITestOutputHelper saida)
{
    private const string FilaPdf = "Microsoft Print to PDF";
    public const string VariavelSondaFila = "MOFFICE_SONDA_IMPRESSAO";

    /// <summary>PDF de 3 paginas pelo x2t (o mesmo motor do PDF de impressao), com uma marca unica no fim.</summary>
    private static (string Pasta, string Pdf, string Marca) PdfDeTresPaginas()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-s5-").FullName;
        var marca = Guid.NewGuid().ToString("N")[..8];
        var texto = string.Join(" ", Enumerable.Range(1, 1500).Select(i => $"palavra{i}")) + " fim-" + marca;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, texto);
        var pdf = Path.Combine(pasta, $"moffice-s5-{marca}.pdf");
        new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar()).ConverterAsync(docx, FormatoDocumento.Pdf, pdf).GetAwaiter().GetResult();
        Assert.Equal(3, mOffice.Conversor.Tests.TextoDePdf.Paginas(pdf));
        return (pasta, pdf, marca);
    }

    /// <summary>Sondas que mexem em fila do sistema ou abrem dialogo do Windows: so' com a variavel ligada.</summary>
    private static void SoComSondaLigada(string oQueFaz)
    {
        Skip.IfNot(Environment.GetEnvironmentVariable(VariavelSondaFila) == "1",
            $"sonda de impressão desligada: {oQueFaz}; ligue com {VariavelSondaFila}=1");
        EditorReal.PularSeNaoDaParaAbrirJanela();
    }

    private static void SoComFilaPdf()
    {
        using var servidor = new LocalPrintServer();
        Skip.IfNot(servidor.GetPrintQueues().Any(f => f.Name == FilaPdf), $"fila \"{FilaPdf}\" ausente nesta máquina");
    }

    /// <summary>Janela VISIVEL que faz o papel da janela do app (dono dos dialogos). Fecha no fim.</summary>
    private static async Task ComJanelaDono(Func<Window, Task> corpo)
    {
        var dono = new Window { Title = "mOffice teste — dono da impressão", Width = 480, Height = 320, Left = 100, Top = 100 };
        dono.Show();
        dono.Activate();
        try { await corpo(dono); }
        finally { dono.Close(); }
    }

    [SkippableFact, Trait("Categoria", "Sonda")] // sem fila e sem dialogo: roda em toda suite
    public void S5_PrintToPdf_PaginasCertasComIntervalo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, pdf, marca) = PdfDeTresPaginas();
        try
        {
            var recorte = Path.Combine(pasta, "recorte-2-3.pdf");
            JanelaSta.Executar(() => Impressora.ComPdfOculto(JanelaEditor.AppPadrao, pdf, Logs.NovoArquivo("sonda-impressao"), async (env, wv, guarda) =>
            {
                // A guarda nao pode ter bloqueado nada: o visualizador embutido nao pede nada fora do file:/// do PDF.
                saida.WriteLine("guarda: bloqueadas=" + string.Join(" | ", guarda.Bloqueadas) + " ; permitidas=" + string.Join(" | ", guarda.Permitidas));
                Assert.Empty(guarda.Bloqueadas);
                // PageRanges "2-3": o visualizador imprime as PAGINAS (nao a moldura) e respeita o intervalo.
                var cfg = env.CreatePrintSettings();
                cfg.PageRanges = "2-3";
                cfg.ShouldPrintHeaderAndFooter = false;
                var ok = await wv.PrintToPdfAsync(recorte, cfg);
                saida.WriteLine($"PrintToPdfAsync(2-3) = {ok}, paginas = {(File.Exists(recorte) ? mOffice.Conversor.Tests.TextoDePdf.Paginas(recorte) : -1)}");
                Assert.True(ok);
                Assert.Equal(2, mOffice.Conversor.Tests.TextoDePdf.Paginas(recorte));
                // As paginas certas, e como TEXTO (nao a moldura do visualizador, nem uma captura): a pagina 1
                // termina antes de "palavra641" (medido), entao o recorte comeca nela e nao tem "palavra1 ".
                var textoRecorte = mOffice.Conversor.Tests.TextoDePdf.Extrair(recorte);
                saida.WriteLine("inicio do recorte: " + textoRecorte.Trim()[..Math.Min(60, textoRecorte.Trim().Length)]);
                Assert.DoesNotContain("palavra1 ", textoRecorte);
                Assert.Contains("fim-" + marca, textoRecorte);
            }), TimeSpan.FromSeconds(180));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Envio de producao SEM fila (fix round 1): a <see cref="Impressora"/> inteira (sinal do visualizador,
    /// opcoes, limite) com o <c>PrintAsync</c> trocado por <c>PrintToPdfAsync</c> (costura) e a fila "verificada"
    /// por costura — nada toca em fila do Windows nem abre dialogo.</summary>
    public static Impressora SemFila(string destino, bool offline = false, Func<Task>? antes = null) => new(JanelaEditor.AppPadrao)
    {
        FilaParaTeste = nome => new Impressora.EstadoDaFila(nome ?? "Fila de teste", offline, false, (8.27, 11.69)),
        EnvioParaTeste = async (env, wv, s) =>
        {
            if (antes is not null) await antes();
            return await wv.PrintToPdfAsync(destino, s) ? CoreWebView2PrintStatus.Succeeded : CoreWebView2PrintStatus.OtherError;
        },
    };

    [SkippableFact] // decisao do controlador (fix round 1): fila que o Windows da' como offline NAO e' recusada — o
                    // trabalho vai e o resultado diz que ela estava offline; o intervalo do painel chega ao envio
    public void Envio_FilaOffline_EnviaEInforma()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, pdf, marca) = PdfDeTresPaginas();
        try
        {
            var destino = Path.Combine(pasta, "enviado.pdf");
            ImpressaoEnviada? r = null;
            JanelaSta.Executar(async () =>
            {
                var o = OpcoesImpressao.Ler("""{"nativeOptions":{"printer":"Brother dormindo","pages":"2-3"}}""");
                r = await SemFila(destino, offline: true).ImprimirAsync(pdf, o);
            }, TimeSpan.FromSeconds(120));
            Assert.Equal(new ImpressaoEnviada("Brother dormindo", true), r);
            Assert.Equal(2, mOffice.Conversor.Tests.TextoDePdf.Paginas(destino));
            Assert.Contains("fim-" + marca, mOffice.Conversor.Tests.TextoDePdf.Extrair(destino));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // fix round 1: envio que nao volta no limite = erro pt-BR (sem detalhe tecnico) e a WebView2 oculta
                    // e' fechada (o finally de ComPdfOculto) — quem chamou apaga o PDF
    public void Envio_TempoEsgotado_MensagemPtBr()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, pdf, _) = PdfDeTresPaginas();
        try
        {
            InvalidOperationException? erro = null;
            JanelaSta.Executar(async () =>
            {
                var nunca = new TaskCompletionSource<CoreWebView2PrintStatus>();
                var i = new Impressora(JanelaEditor.AppPadrao)
                {
                    FilaParaTeste = n => new Impressora.EstadoDaFila("Fila lenta", false, false, null),
                    EnvioParaTeste = (_, _, _) => nunca.Task,
                    LimiteEnvioParaTeste = TimeSpan.FromSeconds(2),
                };
                erro = await Assert.ThrowsAsync<InvalidOperationException>(() => i.ImprimirAsync(pdf, OpcoesImpressao.Padrao));
            }, TimeSpan.FromSeconds(120));
            Assert.Equal("A impressora \"Fila lenta\" não respondeu em 2 segundo(s); o envio foi cancelado.", erro!.Message);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>
    /// MEDIDA do visualizador de PDF (fix round 1): documento grande (300 paginas, cada uma com texto e uma imagem
    /// diferente, ~6 MB de PDF pelo x2t), ambiente FRIO (pasta de dados nova = navegador novo) em cada rodada.
    /// (a) SEM espera nenhuma depois da navegacao, (b) com a folga fixa de 2 s de antes e (c) com o sinal de producao
    /// (<see cref="Impressora.EsperarVisualizadorAsync"/>): as 300 paginas saem e o texto da ULTIMA pagina esta' la'.
    /// Os tempos vao para a saida e para o relatorio da Task 5.
    /// </summary>
    [SkippableFact]
    public void Visualizador_DocumentoGrande_TodasAsPaginas()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-imp-grande-").FullName;
        var apps = new List<string>();
        try
        {
            var marca = Guid.NewGuid().ToString("N")[..8];
            var relogio = Stopwatch.StartNew();
            var docx = mOffice.Conversor.Tests.AmostrasSinteticas.DocxPaginasComImagens(pasta, 300, marca);
            var pdf = Path.Combine(pasta, "grande.pdf");
            new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar()).ConverterAsync(docx, FormatoDocumento.Pdf, pdf).GetAwaiter().GetResult();
            saida.WriteLine($"PDF grande: {new FileInfo(pdf).Length / 1024} KB, {mOffice.Conversor.Tests.TextoDePdf.Paginas(pdf)} paginas, gerado em {relogio.ElapsedMilliseconds} ms");
            Assert.Equal(300, mOffice.Conversor.Tests.TextoDePdf.Paginas(pdf));
            foreach (var modo in new[] { "sem espera", "2 s fixos (antigo)", "com o sinal" })
            {
                var comSinal = modo == "com o sinal";
                var app = "medida-" + Guid.NewGuid().ToString("N")[..6];
                apps.Add(app);
                var saidaPdf = Path.Combine(pasta, $"impresso-{apps.Count}.pdf");
                long esperaMs = 0, impressaoMs = 0;
                JanelaSta.Executar(() => Impressora.ComPdfOculto(app, pdf, Logs.NovoArquivo("medida-impressao"), async (env, wv, _) =>
                {
                    var t = Stopwatch.StartNew();
                    Assert.True(await wv.PrintToPdfAsync(saidaPdf, env.CreatePrintSettings()));
                    impressaoMs = t.ElapsedMilliseconds;
                }, default, async (wv, nav, ct) =>
                {
                    // Com o sinal: a espera de producao, assinada ANTES da navegacao (como no envio); sem: so' a navegacao.
                    var pronto = comSinal ? Impressora.EsperarVisualizadorAsync(wv, nav, new FileInfo(pdf).Length, saida.WriteLine, ct)
                        : modo.StartsWith("2 s") ? nav.ContinueWith(_ => Task.Delay(2000, ct), ct).Unwrap() : nav;
                    await nav;
                    var t = Stopwatch.StartNew();
                    await pronto;
                    esperaMs = t.ElapsedMilliseconds;
                }), TimeSpan.FromSeconds(180));
                var paginas = mOffice.Conversor.Tests.TextoDePdf.Paginas(saidaPdf);
                var texto = mOffice.Conversor.Tests.TextoDePdf.Extrair(saidaPdf);
                saida.WriteLine($"{modo}: espera {esperaMs} ms, PrintToPdf {impressaoMs} ms, {paginas} paginas, ultima pagina presente = {texto.Contains("ultima-pagina-" + marca)}");
                Assert.Equal(300, paginas);
                Assert.Contains("ultima-pagina-" + marca, texto);
                Assert.Contains("Página 1", texto);
            }
        }
        finally
        {
            try { Directory.Delete(pasta, true); } catch { }
            // Pastas de dados das rodadas frias: o navegador de cada uma ainda segura arquivos por alguns segundos
            // depois de a WebView2 fechar — tenta de novo ate' ~15 s.
            foreach (var app in apps)
            {
                var dir = Path.GetDirectoryName(AmbienteWebView2.PastaDadosUsuario(app))!;
                for (var i = 0; i < 30 && Directory.Exists(dir); i++)
                {
                    try { Directory.Delete(dir, true); }
                    catch (Exception e) when (e is IOException or UnauthorizedAccessException) { Thread.Sleep(500); }
                }
                if (Directory.Exists(dir)) saida.WriteLine("pasta de dados da medida nao pode ser apagada: " + dir);
            }
        }
    }

    /// <summary>
    /// O ENVIO DE PRODUCAO (<see cref="Impressora.ImprimirAsync"/>) para a fila VIRTUAL "Microsoft Print to PDF"
    /// (nunca uma fisica), com as opcoes do painel ("2-3"). OPT-IN: mexe numa fila do sistema e responde o dialogo
    /// "Salvar Saída de Impressão como" dela. Mede tambem ONDE esse dialogo aparece em relacao a' janela do app.
    ///
    /// (Task 4: pausar a fila nao retem o trabalho nesta fila — MEDIDO; a prova e' a SAIDA da impressora. Um
    /// trabalho com a marca que por acaso fique na fila e' cancelado no finally.)
    /// </summary>
    [SkippableFact, Trait("Categoria", "Sonda")]
    public void S5_FilaMicrosoftPrintToPdf_ImprimeAsPaginas()
    {
        SoComSondaLigada($"manda um trabalho para \"{FilaPdf}\" e responde o diálogo dela");
        SoComFilaPdf();
        var (pasta, pdf, marca) = PdfDeTresPaginas();
        try
        {
            var impresso = Path.Combine(pasta, "impresso-fila.pdf");
            JanelaSta.Executar(() => ComJanelaDono(async dono =>
            {
                using var servidor = new LocalPrintServer();
                using var fila = servidor.GetPrintQueue(FilaPdf);
                VigiaDeDialogo? vigia = null;
                var impressora = new Impressora(JanelaEditor.AppPadrao, dono)
                {
                    // O dialogo so' e' respondido se vier do navegador DESTE envio (o processo do navegador do
                    // WebView2 ou um descendente dele) e tiver o titulo esperado — nunca o "Salvar como" de outro app.
                    AntesDoEnvioParaTeste = wv => vigia = VigiaDeDialogo.ResponderSalvar(wv.BrowserProcessId, impresso, new WindowInteropHelper(dono).Handle, saida.WriteLine),
                };
                try
                {
                    var o = OpcoesImpressao.Ler($$$"""{"nativeOptions":{"printer":"{{{FilaPdf}}}","pages":"2-3","copies":1,"sides":"one","paperSize":{"w":210,"h":297,"preset":"A4"},"paperOrientation":"portrait","colorMode":true}}""");
                    var enviada = await impressora.ImprimirAsync(pdf, o).WaitAsync(TimeSpan.FromSeconds(90));
                    saida.WriteLine($"ImprimirAsync = \"{enviada}\"; dialogo: {vigia?.Relato}");
                    Assert.Equal(FilaPdf, enviada?.Impressora);
                    Assert.True(vigia is { Respondeu: true }, "o dialogo \"Salvar saida de impressao como\" do navegador deste envio nao apareceu");
                    long tamanho = -1;
                    for (var i = 0; i < 100; i++)
                    {
                        await Task.Delay(200);
                        var agora = File.Exists(impresso) ? new FileInfo(impresso).Length : -1;
                        if (agora > 0 && agora == tamanho) break;
                        tamanho = agora;
                    }
                    Assert.True(File.Exists(impresso), "a impressora nao gerou a saida");
                    var paginas = mOffice.Conversor.Tests.TextoDePdf.Paginas(impresso);
                    saida.WriteLine($"saida da impressora: {new FileInfo(impresso).Length} bytes, {paginas} paginas");
                    Assert.Equal(2, paginas);
                }
                finally
                {
                    vigia?.Dispose();
                    try
                    {
                        fila.Refresh();
                        foreach (var j in fila.GetPrintJobInfoCollection().Where(j => j.Name.Contains(marca, StringComparison.OrdinalIgnoreCase)))
                            try { saida.WriteLine("cancelando trabalho que ficou na fila: " + j.Name); j.Cancel(); } catch (Exception e) { saida.WriteLine("cancelar: " + e.Message); }
                    }
                    catch (Exception e) { saida.WriteLine("limpeza da fila: " + e.Message); }
                }
            }), TimeSpan.FromSeconds(180));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>
    /// SONDA do <c>usesystemdialog</c> (Task 5): o <c>ShowPrintUI(System)</c> do WebView2 a partir de uma WebView2
    /// OCULTA. MEDIDO (Windows 11, runtime fixo 153.0.4234.48): abre o "Imprimir" MODERNO do Windows 11 ("Microsoft
    /// Edge WebView2 - Imprimir", <c>ApplicationFrameWindow</c> do explorer) em ~0,4 s, em primeiro plano — mas
    /// modal so' para a janela OCULTA da WebView2 (o proxy modal e' dela): a janela do app continua clicavel. E a
    /// API nao devolve NADA (nem se imprimiu, nem onde, nem quando acabou), entao nao ha' como apagar o PDF depois
    /// nem dizer "Enviado para …". Por isso a producao usa o PrintDialog do WPF (o MESMO dialogo do Windows, modal
    /// para a janela do documento) + PrintAsync. Cancelar devolve a WebView2 ao normal (medido: um PrintToPdf logo
    /// depois funciona).
    ///
    /// OPT-IN (abre dialogo do sistema, que o vigia CANCELA — so' o "Imprimir" cujo proxy modal e' de uma janela
    /// deste processo/navegador e cujo titulo e' exatamente "&lt;programa nosso&gt; - Imprimir"). Isolada numa pasta
    /// de dados PROPRIA ("sonda-showprintui"): se o dialogo ficasse aberto, o navegador ficaria preso esperando por
    /// ele (medido antes de o vigia conhecer o dialogo moderno) e travaria a criacao de WebView2 nessa pasta; por
    /// isso, no fim, o navegador DESTA sonda (o pid do navegador desta WebView2 e os descendentes, so'
    /// msedgewebview2) e' encerrado.
    /// </summary>
    [SkippableFact, Trait("Categoria", "Sonda")]
    public void Sonda_ShowPrintUISistema_NaWebViewOculta()
    {
        SoComSondaLigada("abre o diálogo Imprimir do sistema pelo WebView2 oculto (e o cancela)");
        var (pasta, pdf, _) = PdfDeTresPaginas();
        uint navegador = 0;
        // Fix round 1: so' se encerra processo cujo HANDLE foi pego com ele VIVO e cuja linha de comando e' a da
        // pasta de dados DESTA sonda — nunca por um pid guardado (o pid pode ter sido reusado por outro processo).
        using var nossos = new ProcessosDaSonda(AmbienteWebView2.PastaDadosUsuario("sonda-showprintui"));
        try
        {
            JanelaSta.Executar(() => ComJanelaDono(dono => Impressora.ComPdfOculto("sonda-showprintui", pdf, Logs.NovoArquivo("sonda-impressao"), async (env, wv, guarda) =>
            {
                navegador = wv.BrowserProcessId;
                nossos.Capturar(navegador);
                var eu = (uint)Environment.ProcessId;
                using var vigia = VigiaDeDialogo.Cancelar(() => { var p = VigiaDeDialogo.ArvoreDeProcessos(navegador); p.Add(eu); return p; }, VigiaDeDialogo.TitulosImprimir,
                    new WindowInteropHelper(dono).Handle, saida.WriteLine);
                wv.ShowPrintUI(CoreWebView2PrintDialogKind.System);
                var relogio = Stopwatch.StartNew();
                while (!vigia.Respondeu && relogio.Elapsed < TimeSpan.FromSeconds(20)) await Task.Delay(200);
                saida.WriteLine($"ShowPrintUI(System) na WebView2 oculta: dialogo {(vigia.Respondeu ? "APARECEU" : "NAO apareceu")} em {relogio.ElapsedMilliseconds} ms; {vigia.Relato}");
                if (!vigia.Respondeu) saida.WriteLine("janelas do navegador: " + string.Join(" ; ", vigia.JanelasDosProcessos()));
                Assert.True(vigia.Respondeu, "o dialogo Imprimir do ShowPrintUI(System) nao foi visto");
                Assert.Equal("moderno", vigia.Forma);
                Assert.False(vigia.DonoEAJanelaDoApp, "o dialogo do ShowPrintUI ficou modal para a janela do app: rever a decisao da Task 5");

                // Depois do cancelamento a WebView2 volta a responder (um PDF de 1 pagina, sem fila).
                await Task.Delay(1000);
                var depois = Path.Combine(pasta, "depois-do-cancelar.pdf");
                var cfg = env.CreatePrintSettings(); cfg.PageRanges = "1";
                nossos.Capturar(navegador); // de novo, com os filhos que nasceram depois
                Assert.True(await wv.PrintToPdfAsync(depois, cfg).WaitAsync(TimeSpan.FromSeconds(30)));
            })), TimeSpan.FromSeconds(180));
        }
        finally
        {
            // Rede de seguranca: se o dialogo tivesse ficado aberto, o navegador desta sonda estaria preso.
            foreach (var linha in nossos.EncerrarOsQueSobraram()) saida.WriteLine(linha);
            try { Directory.Delete(pasta, true); } catch { }
        }
    }

    /// <summary>
    /// O dialogo "Imprimir" do Windows do caminho de producao do <c>usesystemdialog</c>
    /// (<see cref="Impressora.EscolherNoDialogoDoSistemaAsync"/>, <c>PrintDialog</c> do WPF): aparece e cancelar devolve
    /// <c>null</c> (nada e' impresso). OPT-IN (abre dialogo do sistema; o vigia so' age no "Imprimir" DESTE processo).
    ///
    /// O DONO nao da' para provar aqui: o WPF usa <c>Application.Current.MainWindow</c> como dono do
    /// <c>PrintDlgEx</c> (conferido no codigo do PresentationFramework 10.0.10) e o processo de teste nao tem
    /// <c>Application</c> — o dono vira a area de trabalho e o Windows mostra o "Imprimir" CLASSICO (#32770). Com
    /// <c>Application</c> (o app de verdade) a producao troca a <c>MainWindow</c> pela janela do documento durante o
    /// dialogo, e o Windows 11 mostra o "Imprimir" MODERNO modal para ELA — medido num programa WPF avulso (com
    /// controle negativo: sem a troca, o proxy modal fica na boas-vindas escondida e a janela do documento continua
    /// clicavel). Ver o relatorio da Task 5.
    /// </summary>
    [SkippableFact, Trait("Categoria", "Sonda")]
    public void DialogoDoSistema_Cancelado_NaoImprime()
    {
        SoComSondaLigada("abre o diálogo Imprimir do Windows (e o cancela)");
        JanelaSta.Executar(() => ComJanelaDono(async dono =>
        {
            await Task.Yield();
            var eu = (uint)Environment.ProcessId;
            using var vigia = VigiaDeDialogo.Cancelar(() => [eu], VigiaDeDialogo.TitulosImprimir, new WindowInteropHelper(dono).Handle, saida.WriteLine);
            var o = OpcoesImpressao.Ler("""{"nativeOptions":{"usesystemdialog":true,"pages":"2-3","copies":2}}""");
            var escolhida = await new Impressora(JanelaEditor.AppPadrao, dono).EscolherNoDialogoDoSistemaAsync(o);
            saida.WriteLine($"dialogo do sistema: {vigia.Relato}");
            Assert.True(vigia.Respondeu, "o dialogo Imprimir do Windows nao apareceu");
            Assert.Null(escolhida);
        }), TimeSpan.FromSeconds(120));
    }
}

/// <summary>
/// Processos do navegador do WebView2 de UMA sonda, para encerrar os que sobrarem (fix round 1 da Task 5). Regra:
/// so' entra quem esta' VIVO no momento da captura (o <see cref="Process"/> fica guardado — o handle aberto impede
/// confundir com outro processo que reuse o pid), e' <c>msedgewebview2</c> e tem na LINHA DE COMANDO a pasta de
/// dados desta sonda (<c>--user-data-dir</c>). Nada e' encerrado por pid guardado.
/// </summary>
internal sealed class ProcessosDaSonda(string pastaDeDados) : IDisposable
{
    private readonly List<Process> _capturados = [];

    /// <summary>Captura <paramref name="raiz"/> e descendentes que passam na regra (os ja' capturados ficam).</summary>
    public void Capturar(uint raiz)
    {
        foreach (var pid in VigiaDeDialogo.ArvoreDeProcessos(raiz))
        {
            if (_capturados.Any(p => p.Id == pid)) continue;
            Process? p = null;
            try
            {
                p = Process.GetProcessById((int)pid);
                if (!p.HasExited && p.ProcessName.Equals("msedgewebview2", StringComparison.OrdinalIgnoreCase)
                    && LinhaDeComando(p) is { } linha && linha.Contains(pastaDeDados, StringComparison.OrdinalIgnoreCase))
                { _capturados.Add(p); p = null; }
            }
            catch (Exception e) when (e is ArgumentException or InvalidOperationException or System.ComponentModel.Win32Exception) { }
            finally { p?.Dispose(); }
        }
    }

    /// <summary>Encerra os capturados que ainda estao vivos (pelo handle guardado). Devolve o relato.</summary>
    public List<string> EncerrarOsQueSobraram()
    {
        var relato = new List<string>();
        foreach (var p in _capturados)
            try
            {
                if (p.HasExited) continue;
                p.Kill(); p.WaitForExit(5000);
                relato.Add($"navegador da sonda encerrado: pid {p.Id}");
            }
            catch (Exception e) when (e is InvalidOperationException or System.ComponentModel.Win32Exception) { relato.Add("encerrar: " + e.Message); }
        return relato;
    }

    public void Dispose() { foreach (var p in _capturados) p.Dispose(); _capturados.Clear(); }

    /// <summary>Linha de comando de outro processo (NtQueryInformationProcess, classe 60), ou null.</summary>
    private static string? LinhaDeComando(Process p)
    {
        NtQueryInformationProcess(p.Handle, 60, IntPtr.Zero, 0, out var tamanho);
        if (tamanho <= 0) return null;
        var buffer = Marshal.AllocHGlobal(tamanho);
        try
        {
            if (NtQueryInformationProcess(p.Handle, 60, buffer, tamanho, out _) != 0) return null;
            var us = Marshal.PtrToStructure<UNICODE_STRING>(buffer);
            return us.Buffer == IntPtr.Zero ? null : Marshal.PtrToStringUni(us.Buffer, us.Length / 2);
        }
        finally { Marshal.FreeHGlobal(buffer); }
    }

    [StructLayout(LayoutKind.Sequential)] private struct UNICODE_STRING { public ushort Length, MaximumLength; public IntPtr Buffer; }
    [DllImport("ntdll.dll")] private static extern int NtQueryInformationProcess(IntPtr processo, int classe, IntPtr info, int tamanho, out int devolvido);
}

/// <summary>
/// Vigia de UM dialogo do Windows num teste (Task 4, fix round 1; generalizado na Task 5). Duas formas de dialogo,
/// as duas com escopo FECHADO — nunca o "Salvar como"/"Imprimir" de outro app:
/// <list type="bullet">
/// <item><b>Classico</b>: janela VISIVEL de classe #32770, com um dos TITULOS esperados, de um dos PROCESSOS
/// permitidos (o navegador do WebView2 DESTE teste e seus descendentes, ou o proprio processo de teste).</item>
/// <item><b>"Imprimir" moderno do Windows 11</b> (medido na Task 5): quando <c>PrintDlgEx</c> recebe uma janela
/// dona, o Windows mostra o dialogo novo — uma <c>ApplicationFrameWindow</c> do explorer com o titulo
/// "&lt;processo&gt; - Imprimir" — e deixa no processo que pediu uma <c>ProxyModalWindow</c> cujo DONO e' a janela
/// que fica bloqueada. So' conta um par em que o dono do proxy e' de um dos PROCESSOS permitidos e o titulo e'
/// EXATAMENTE "&lt;nome desse processo&gt; - Imprimir"/"- Print".</item>
/// </list>
/// Ao achar, mede onde o dialogo esta' (dono, primeiro plano, acima da janela do app) e responde: preenche o nome e
/// da' OK (<see cref="ResponderSalvar"/>) ou cancela (<see cref="Cancelar"/>). No <c>Dispose</c>, um dialogo do
/// mesmo escopo que tenha sobrado aberto e' fechado (= cancelar).
/// </summary>
internal sealed class VigiaDeDialogo : IDisposable
{
    public static readonly string[] TitulosSalvar = ["Salvar Saída de Impressão como", "Save Print Output As"];
    public static readonly string[] TitulosImprimir = ["Imprimir", "Print"];

    private readonly CancellationTokenSource _cts = new();
    private readonly Task _vigia;
    private readonly Func<HashSet<uint>> _processos;
    private readonly string[] _titulos;
    private readonly bool _moderno;
    public bool Respondeu { get; private set; }
    public uint PidDoDialogo { get; private set; }
    public string Relato { get; private set; } = "(dialogo nao visto)";
    /// <summary>"classico" (#32770) ou "moderno" (Imprimir do Windows 11).</summary>
    public string? Forma { get; private set; }
    /// <summary>O dono do dialogo (cadeia de donos no classico; dono do proxy modal no moderno) e' a janela do app?</summary>
    public bool DonoEAJanelaDoApp { get; private set; }

    private VigiaDeDialogo(Func<HashSet<uint>> processos, string[] titulos, bool moderno, IntPtr janelaDoApp, Func<Achado, Task> responder, Action<string> log)
    {
        _processos = processos;
        _titulos = titulos;
        _moderno = moderno;
        _vigia = Task.Run(async () =>
        {
            while (!_cts.IsCancellationRequested)
            {
                var a = AcharDialogo();
                if (a is { } dlg)
                {
                    // Primeiro plano: amostra por ate' 1 s (o dialogo pode levar alguns ms para ser ativado).
                    var primeiroPlano = false;
                    for (var i = 0; i < 10 && !primeiroPlano; i++) { primeiroPlano = GetForegroundWindow() == dlg.Janela; if (!primeiroPlano) await Task.Delay(100); }
                    var donos = new List<IntPtr>();
                    for (var h = dlg.Dono; h != IntPtr.Zero && donos.Count < 10; h = GetWindow(h, GW_OWNER)) donos.Add(h);
                    DonoEAJanelaDoApp = janelaDoApp != IntPtr.Zero && donos.Contains(janelaDoApp);
                    var acima = janelaDoApp != IntPtr.Zero && AcimaDe(dlg.Janela, janelaDoApp);
                    var bloqueada = janelaDoApp != IntPtr.Zero && !IsWindowEnabled(janelaDoApp);
                    Forma = dlg.Moderno ? "moderno" : "classico";
                    Relato = $"{Forma} \"{Titulo(dlg.Janela)}\" [{Classe(dlg.Janela)}] pid={dlg.Pid}; donos=[{string.Join(" > ", donos.Select(d => $"0x{d:X} '{Titulo(d)}' {Classe(d)} visivel={IsWindowVisible(d)}"))}]"
                        + $"; dono e' a janela do app={DonoEAJanelaDoApp}; janela do app bloqueada (modal)={bloqueada}; primeiro plano={primeiroPlano}; acima da janela do app={acima}";
                    log("dialogo achado: " + Relato);
                    await responder(dlg);
                    PidDoDialogo = dlg.Pid;
                    Respondeu = true;
                    return;
                }
                await Task.Delay(200);
            }
        });
    }

    private readonly record struct Achado(IntPtr Janela, uint Pid, IntPtr Dono, bool Moderno);

    /// <summary>Todas as janelas de nivel superior dos processos vigiados (visiveis ou nao), para o relato.</summary>
    public List<string> JanelasDosProcessos()
    {
        var nossos = _processos(); var r = new List<string>();
        EnumWindows((h, _) => { GetWindowThreadProcessId(h, out var pid); if (nossos.Contains(pid) && Classe(h) != "IME" && Classe(h) != "MSCTFIME UI") { GetWindowRect(h, out var rc); r.Add($"{pid} vis={IsWindowVisible(h)} [{Classe(h)}] '{Titulo(h)}' ({rc.L},{rc.T},{rc.R},{rc.B}) dono=0x{GetWindow(h, GW_OWNER):X}"); } return true; }, IntPtr.Zero);
        return r;
    }

    /// <summary>"Salvar Saída de Impressão como" (classico) do navegador <paramref name="pidNavegador"/> (e
    /// descendentes): caminho <paramref name="destino"/> + OK.</summary>
    public static VigiaDeDialogo ResponderSalvar(uint pidNavegador, string destino, IntPtr janelaDoApp, Action<string> log) =>
        new(() => ArvoreDeProcessos(pidNavegador), TitulosSalvar, moderno: false, janelaDoApp, async dlg =>
        {
            var edit = AcharEditDoNome(dlg.Janela);
            if (edit == IntPtr.Zero) { log("dialogo de salvar sem campo de nome"); PostMessage(dlg.Janela, WM_CLOSE, IntPtr.Zero, IntPtr.Zero); return; }
            SendMessage(edit, WM_SETTEXT, IntPtr.Zero, destino);
            await Task.Delay(200);
            var botao = GetDlgItem(dlg.Janela, 1);
            if (botao != IntPtr.Zero) SendMessage(botao, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            else PostMessage(dlg.Janela, WM_COMMAND, (IntPtr)1, IntPtr.Zero);
        }, log);

    /// <summary>O dialogo "Imprimir" (classico OU moderno) dos <paramref name="processos"/>: CANCELA (classico:
    /// IDCANCEL; moderno: WM_CLOSE na moldura, que o Windows trata como Cancelar).</summary>
    public static VigiaDeDialogo Cancelar(Func<HashSet<uint>> processos, string[] titulos, IntPtr janelaDoApp, Action<string> log) =>
        new(processos, titulos, moderno: true, janelaDoApp, dlg =>
        {
            if (dlg.Moderno) PostMessage(dlg.Janela, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            else PostMessage(dlg.Janela, WM_COMMAND, (IntPtr)IDCANCEL, IntPtr.Zero);
            return Task.CompletedTask;
        }, log);

    public void Dispose()
    {
        _cts.Cancel();
        try { _vigia.Wait(2000); } catch { }
        if (AcharDialogo() is { } aberto) PostMessage(aberto.Janela, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }

    private Achado? AcharDialogo()
    {
        var nossos = _processos();
        Achado? achado = null;
        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h) || Classe(h) != "#32770") return true;
            if (!_titulos.Contains(Titulo(h))) return true;
            GetWindowThreadProcessId(h, out var pid);
            if (!nossos.Contains(pid)) return true;
            achado = new Achado(h, pid, GetWindow(h, GW_OWNER), false); return false;
        }, IntPtr.Zero);
        if (achado is not null || !_moderno) return achado;

        // Moderno: um proxy modal cujo dono e' dos nossos processos + a moldura com o titulo EXATO desse processo.
        var esperados = new Dictionary<string, IntPtr>(StringComparer.Ordinal);
        EnumWindows((h, _) =>
        {
            if (Classe(h) != "ProxyModalWindow") return true;
            var dono = GetWindow(h, GW_OWNER);
            GetWindowThreadProcessId(dono, out var pidDono);
            if (dono == IntPtr.Zero || !nossos.Contains(pidDono)) return true;
            // O titulo leva o nome do programa que PEDIU o dialogo — que pode ser outro dos nossos processos (o
            // ShowPrintUI do WebView2 pede do navegador, "Microsoft Edge WebView2", com o proxy na janela do host).
            foreach (var pid in nossos)
                foreach (var nome in NomesDoPrograma(pid))
                    foreach (var t in _titulos) esperados[$"{nome} - {t}"] = dono;
            return true;
        }, IntPtr.Zero);
        if (esperados.Count == 0) return null;
        EnumWindows((h, _) =>
        {
            if (!IsWindowVisible(h) || Classe(h) != "ApplicationFrameWindow") return true;
            if (!esperados.TryGetValue(Titulo(h), out var dono)) return true;
            GetWindowThreadProcessId(h, out var pid);
            achado = new Achado(h, pid, dono, true); return false;
        }, IntPtr.Zero);
        return achado;
    }

    /// <summary>Descricao do executavel (o que o Windows poe no titulo do "Imprimir" moderno) e nome do processo.</summary>
    private static IEnumerable<string> NomesDoPrograma(uint pid)
    {
        var nomes = new List<string>();
        try
        {
            using var p = Process.GetProcessById((int)pid);
            nomes.Add(p.ProcessName);
            try { if (p.MainModule?.FileVersionInfo.FileDescription is { Length: > 0 } d) nomes.Add(d); }
            catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException) { }
        }
        catch (Exception e) when (e is ArgumentException or InvalidOperationException) { }
        return nomes;
    }

    /// <summary><paramref name="a"/> esta' acima de <paramref name="b"/> na ordem Z (EnumWindows vai de cima para baixo).</summary>
    private static bool AcimaDe(IntPtr a, IntPtr b)
    {
        int ia = -1, ib = -1, i = 0;
        EnumWindows((h, _) => { if (h == a) ia = i; if (h == b) ib = i; i++; return ia < 0 || ib < 0; }, IntPtr.Zero);
        return ia >= 0 && ib >= 0 && ia < ib;
    }

    /// <summary><paramref name="raiz"/> e todos os descendentes dela (instantaneo do Toolhelp32).</summary>
    public static HashSet<uint> ArvoreDeProcessos(uint raiz)
    {
        var filhos = new Dictionary<uint, List<uint>>();
        var snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snap == IntPtr.Zero || snap == new IntPtr(-1)) return [raiz];
        try
        {
            var e = new PROCESSENTRY32 { dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32>() };
            for (var ok = Process32First(snap, ref e); ok; ok = Process32Next(snap, ref e))
            {
                if (!filhos.TryGetValue(e.th32ParentProcessID, out var l)) filhos[e.th32ParentProcessID] = l = [];
                l.Add(e.th32ProcessID);
            }
        }
        finally { CloseHandle(snap); }
        var arvore = new HashSet<uint> { raiz };
        var fila = new Queue<uint>([raiz]);
        while (fila.Count > 0)
            if (filhos.TryGetValue(fila.Dequeue(), out var l))
                foreach (var f in l) if (arvore.Add(f)) fila.Enqueue(f);
        return arvore;
    }

    private static IntPtr AcharEditDoNome(IntPtr dlg)
    {
        var achado = IntPtr.Zero;
        EnumChildWindows(dlg, (h, _) =>
        {
            if (Classe(h) == "Edit" && Classe(GetParent(h)) == "ComboBox") { achado = h; return false; }
            return true;
        }, IntPtr.Zero);
        return achado;
    }

    private static string Classe(IntPtr h) { var sb = new StringBuilder(256); GetClassName(h, sb, sb.Capacity); return sb.ToString(); }
    private static string Titulo(IntPtr h) { var sb = new StringBuilder(256); GetWindowText(h, sb, sb.Capacity); return sb.ToString(); }

    private const int WM_SETTEXT = 0x000C, WM_CLOSE = 0x0010, WM_COMMAND = 0x0111, BM_CLICK = 0x00F5, IDCANCEL = 2;
    private const uint TH32CS_SNAPPROCESS = 0x2, GW_OWNER = 4;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PROCESSENTRY32
    {
        public uint dwSize, cntUsage, th32ProcessID;
        public IntPtr th32DefaultHeapID;
        public uint th32ModuleID, cntThreads, th32ParentProcessID;
        public int pcPriClassBase;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szExeFile;
    }

    private delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint pid);
    [DllImport("kernel32.dll", EntryPoint = "Process32FirstW", CharSet = CharSet.Unicode)] private static extern bool Process32First(IntPtr snap, ref PROCESSENTRY32 e);
    [DllImport("kernel32.dll", EntryPoint = "Process32NextW", CharSet = CharSet.Unicode)] private static extern bool Process32Next(IntPtr snap, ref PROCESSENTRY32 e);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc f, IntPtr l);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr pai, EnumProc f, IntPtr l);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr h, uint cmd);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr h);
    private struct RC { public int L, T, R, B; }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr h, out RC r);
    [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr h, int id);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
}
