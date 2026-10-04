using System.IO.Compression;
using System.Text.RegularExpressions;
using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 9 na janela real: formulas e constantes pt-BR do mPlanilha (e o xlsx gravado em ingles), nome da aba nova,
/// documento novo em Calibri 11 com "Título 1", layouts do PowerPoint pt-BR.</summary>
[Collection(JanelaReal.Nome)]
public class TextosPtBrJanelaTests
{
    private static string NomeDoApp(string app) => app switch { "mplanilha" => "mPlanilha", "mslides" => "mSlides", _ => "mTexto" };

    private static async Task<JanelaEditor> AbrirNovaAsync(string app, FormatoDocumento f)
    {
        var j = EditorReal.PrepararParaTeste(new JanelaEditor(app) { NomeApp = NomeDoApp(app) });
        j.Show();
        await j.NovoAsync(f);
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
        return j;
    }

    /// <summary>Avalia <paramref name="corpo"/> no editor (variaveis <c>w</c>, o quadro, e <c>api</c>) e devolve o texto sem as aspas do JSON.</summary>
    private static async Task<string> NoEditor(JanelaEditor j, string corpo) => (await j.ExecutarJsAsync(EditorReal.NoEditor(corpo))).Trim('"').Replace("\\\"", "\"");

    private static string Parte(string pacote, string parte) { using var z = ZipFile.OpenRead(pacote); using var r = new StreamReader(z.GetEntry(parte)!.Open()); return r.ReadToEnd(); }

    // =A1..A9: o que o usuario do Excel pt-BR digita. DATADIF leva a unidade "y" (as unidades do DATEDIF nao sao traduzidas pelo motor;
    // o brief previa "a" e o motor responde #NUM! — medido ao vivo e registrado no relatorio).
    private static readonly string[] Formulas =
    [
        "=SE(VERDADEIRO;1;2)",                                 // A1 = 1
        "=1/0",                                                // A2 = #DIV/0!
        "=XPTO()",                                             // A3 = #NAME? (exibido #NOME?)
        "=PROCV(9;A1:B2;2;FALSO)",                             // A4 = #N/A  (exibido #N/D)
        "=DATADIF(DATA(2020;1;15);DATA(2026;1;15);\"y\")",     // A5 = 6
        "=SE(FALSO;1;2)",                                      // A6 = 2
        "=TIPO.ERRO(XPTO())",                                  // A7 = 5 (#NAME?)
        "=ÉERROS(#N/D)",                                       // A8 = TRUE: o literal "#N/D" digitado e' lido como erro (e gravado como #N/A)
        "=SE(ÉERROS(#VALOR!);\"sim\";\"nao\")",                // A9 = "sim"
        "=ÉERROS(#NÚM!)",                                      // A10..A13: literais de erro digitados em pt-BR sao gravados em ingles
        "=ÉERROS(#NULO!)",                                     // A11
        "=ÉERROS(#DIV/0!)",                                    // A12
        "=ÉERROS(#REF!)",                                      // A13
    ];

    [SkippableFact] // planilha: as formulas e constantes em pt-BR funcionam; o xlsx gravado tem DATEDIF, IF, TRUE/FALSE e os erros em ingles
    public void Planilha_FormulasEConstantesPtBr()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-ptbr-").FullName;
        var destino = Path.Combine(pasta, "formulas.xlsx");
        var reaberto = new List<string>();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = await AbrirNovaAsync("mplanilha", FormatoDocumento.Xlsx);
                for (var i = 0; i < Formulas.Length; i++)
                    Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoCell(Formulas[i], "A" + (i + 1)))).Trim('"'));
                await Task.Delay(500);
                var r = await j.SalvarComoAsync(destino, FormatoDocumento.Xlsx);
                Assert.True(File.Exists(destino), "xlsx nao gravado");
                j.Close();

                // ida e volta: o xlsx (formulas em ingles) reabre mostrando os nomes em pt-BR, com os mesmos resultados
                var k = EditorReal.PrepararParaTeste(new JanelaEditor("mplanilha") { NomeApp = "mPlanilha" });
                k.Show();
                await k.AbrirAsync(destino);
                Assert.Equal("true", await EditorReal.Esperar(k, EditorReal.ProntoNoEditor, "true", 300));
                for (var i = 0; i < Formulas.Length; i++)
                    reaberto.Add(await NoEditor(k, "api.asc_findCell('A" + (i + 1) + "'); return api.asc_getCellInfo().asc_getText();"));
                k.Close();
            }, TimeSpan.FromSeconds(300));
            Assert.Equal(Formulas, reaberto);

            var xml = Parte(destino, "xl/worksheets/sheet1.xml");
            string Celula(string c) => Regex.Match(xml, "<c r=\"" + c + "\"[^>]*>.*?</c>", RegexOptions.Singleline).Value;
            // formula em ingles + valor calculado
            Assert.Matches(@"<f>IF\(TRUE\(?\)?,1,2\)</f>", Celula("A1")); Assert.Contains("<v>1</v>", Celula("A1"));
            Assert.Contains("<f>1/0</f>", Celula("A2")); Assert.Contains("#DIV/0!", Celula("A2"));
            Assert.Contains("<f>XPTO()</f>", Celula("A3")); Assert.Contains("#NAME?", Celula("A3"));
            Assert.Contains("VLOOKUP(9,A1:B2,2,FALSE", Celula("A4")); Assert.Contains("#N/A", Celula("A4"));
            Assert.Contains("DATEDIF(DATE(2020,1,15),DATE(2026,1,15),&quot;y&quot;)", Celula("A5")); Assert.Contains("<v>6</v>", Celula("A5"));
            Assert.Matches(@"<f>IF\(FALSE\(?\)?,1,2\)</f>", Celula("A6")); Assert.Contains("<v>2</v>", Celula("A6"));
            Assert.Contains("ERROR.TYPE(XPTO())", Celula("A7")); Assert.Contains("<v>5</v>", Celula("A7"));
            // Os literais de erro digitados em pt-BR sao gravados em INGLES, o canonico do OOXML.
            Assert.Contains("<f>ISERROR(#NUM!)</f>", Celula("A10")); Assert.Contains("<f>ISERROR(#NULL!)</f>", Celula("A11"));
            Assert.Contains("<f>ISERROR(#DIV/0!)</f>", Celula("A12")); Assert.Contains("<f>ISERROR(#REF!)</f>", Celula("A13"));
            Assert.Contains("ISERROR(#N/A)", Celula("A8")); Assert.Contains("<v>1</v>", Celula("A8"));
            Assert.Contains("IF(ISERROR(#VALUE!),&quot;sim&quot;,&quot;nao&quot;)", Celula("A9")); Assert.Contains("sim", Celula("A9"));
            foreach (var pt in new[] { "SE(", "VERDADEIRO", "FALSO", "DATADIF", "PROCV", "#NOME", "#N/D", "#VALOR", "#NÚM", "#NULO" }) Assert.DoesNotContain(pt, xml);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // planilha nova: a 1a aba e' "Planilha1" e a proxima que o editor propoe e' "Planilha2" (nao "Folha1")
    public void Planilha_NovaAbaPlanilha()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string primeira = "", proxima = "", prefixo = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mplanilha", FormatoDocumento.Xlsx);
            primeira = await NoEditor(j, "return api.asc_getWorksheetName(0);");
            // o botao "+" das abas chama Statusbar.createSheetName(): strSheet (txtSheet traduzido) + o primeiro numero livre
            proxima = await NoEditor(j, "return w.SSE.getController('Statusbar').createSheetName();");
            prefixo = await NoEditor(j, "return w.SSE.getController('Statusbar').strSheet;");
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.Equal("Planilha1", primeira);
        Assert.Equal("Planilha2", proxima);
        Assert.Equal("Planilha", prefixo);
    }

    [SkippableFact] // documento novo: Calibri 11 no primeiro paragrafo; o estilo de titulo se chama "Título 1"
    public void Documento_NovoCalibriETitulo1()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string fonte = "", tamanho = "", titulo = "", idioma = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mtexto", FormatoDocumento.Docx);
            fonte = await NoEditor(j, "var p=api.get_TextProps().get_TextPr(); return p.get_FontFamily().get_Name();");
            tamanho = await NoEditor(j, "var p=api.get_TextProps().get_TextPr(); return String(p.get_FontSize());");
            titulo = await NoEditor(j, "return w.DE.getController('Main').txtStyle_Heading_1;");
            idioma = await NoEditor(j, "return String(api.asc_getDefaultLanguage());");
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.Equal("Calibri", fonte);
        Assert.Equal("11", tamanho);
        Assert.Equal("Título 1", titulo);
        Assert.Equal("1046", idioma); // pt-BR (o w:lang do modelo)
    }

    [SkippableFact] // apresentacao nova: o menu de layouts lista os nomes do PowerPoint pt-BR
    public void Slides_LayoutsPtBr()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string nomes = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mslides", FormatoDocumento.Pptx);
            // o mesmo evento de onde o web-apps preenche o menu de layouts (asc_onUpdateLayout); UpdateInterfaceState o redispara
            await NoEditor(j, "window.__layoutsDoTeste=null; api.asc_registerCallback('asc_onUpdateLayout', function(l){ window.__layoutsDoTeste = (l||[]).map(function(x){return x.get_Name();}); }); api.UpdateInterfaceState(); return 'ok';");
            await Task.Delay(1000);
            nomes = await j.ExecutarJsAsync("JSON.stringify(window.__layoutsDoTeste || null)");
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.Contains("Slide de Título", nomes);
        Assert.Contains("Título e Conteúdo", nomes);
        Assert.DoesNotContain("Diapositivo", nomes);
    }

    private static string SalvarComA1(string formula, string pasta)
    {
        var destino = Path.Combine(pasta, "literal.xlsx");
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mplanilha", FormatoDocumento.Xlsx);
            await j.ExecutarJsAsync(EditorReal.InserirTextoCell(formula, "A1"));
            await Task.Delay(500);
            await j.SalvarComoAsync(destino, FormatoDocumento.Xlsx);
            j.Close();
        }, TimeSpan.FromSeconds(180));
        return Parte(destino, "xl/worksheets/sheet1.xml");
    }

    private static readonly HashSet<string> LiteraisCanonicos = ["#NAME?", "#N/A", "#VALUE!", "#NUM!", "#NULL!", "#DIV/0!", "#REF!"];
    private static IEnumerable<string> LiteraisNasFormulas(string xml) => Regex.Matches(xml, "<f>(.*?)</f>").SelectMany(m => Regex.Matches(m.Groups[1].Value, @"#[^,()+\-*/ &]*(?:/0!|/[ADad])?[!?]?").Select(x => x.Value));

    [SkippableTheory] // literal de erro digitado em MAIUSCULAS (como o Excel pt-BR) nunca vira formula invalida: ou e' recusado, ou o <f> traz so' o canonico em INGLES.
                      // "#NOME?"/"#NOME"/"#NOM"/"#NAME": recusados (o "name" do pt-br.json fica como o original, #NAME\?; ver MODIFICACOES-ONLYOFFICE.md).
    [InlineData("#NOME?")] [InlineData("#NOME")] [InlineData("#NOM")] [InlineData("#NAME")] [InlineData("#NAME?")]
    [InlineData("#N/D")] [InlineData("#VALOR!")] [InlineData("#NÚM!")] [InlineData("#NULO!")] [InlineData("#DIV/0!")] [InlineData("#REF!")]
    public void Planilha_LiteralDeErroDigitado_NuncaGravaFormulaInvalida(string literal)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-nome-").FullName;
        try
        {
            var xml = SalvarComA1("=ÉERROS(" + literal + ")", pasta);
            Assert.DoesNotContain("ÉERRO", string.Concat(Regex.Matches(xml, "<f>(.*?)</f>").Select(m => m.Groups[1].Value))); // nome de funcao em pt-BR nunca vai para o arquivo
            Assert.All(LiteraisNasFormulas(xml), l => Assert.Contains(l, LiteraisCanonicos));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableTheory] // LIMITACAO CONHECIDA (Task 12 / diferencas-conhecidas): o motor casa o literal de erro SEM diferenciar maiusculas e, se o texto digitado nao for o
                      // canonico exato, grava o digitado CRU. Com o motor original isso ja' acontecia com "#n/a", "#value!" (INGLES minusculo, medido sem o patch 05);
                      // com o patch 05 acontece tambem com "#n/d", "#valor!", "#núm!", "#nulo!" minusculos (e o valor calculado vira #NAME?). Digitar erro em minuscula numa formula e' raro.
                      // O teste afirma o comportamento ATUAL (nao o desejavel): se o motor passar a normalizar, ele falha e a nota some.
    [InlineData("#n/d")] [InlineData("#valor!")] [InlineData("#núm!")] [InlineData("#nulo!")] [InlineData("#nome?")] [InlineData("#div/0!")] [InlineData("#ref!")]
    public void Planilha_LiteralDeErroEmMinuscula_GravaCruComoOMotorOriginal(string literal)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-min-").FullName;
        try
        {
            var xml = SalvarComA1("=ÉERROS(" + literal + ")", pasta);
            var achados = LiteraisNasFormulas(xml).ToList();
            if (literal == "#nome?") Assert.Empty(achados);               // recusado (o "?" exige o canonico #NAME?; "#name?" em ingles minusculo sai cru, como no original)
            else Assert.Equal([literal], achados);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
