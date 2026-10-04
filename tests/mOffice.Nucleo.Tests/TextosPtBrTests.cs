using System.Text.Json;
using System.Text.RegularExpressions;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 9 (A7/A8/PL-05/PL-06/PL-11): os textos pt-BR do motor patchado (patches 03, 04 e 05). Sem motor: pulam.</summary>
public class TextosPtBrTests
{
    private static string Apps => Path.Combine(FronteiraTests.RaizRepo(), "third_party", "onlyoffice", "editors", "web-apps", "apps");

    private static Dictionary<string, string> Locale(string app)
    {
        var arq = Path.Combine(Apps, app, "main", "locale", "pt.json");
        Skip.If(!File.Exists(arq), "motor nao obtido (rode tools/onlyoffice/obter-motor.ps1)");
        using var d = JsonDocument.Parse(File.ReadAllBytes(arq));
        return d.RootElement.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString() ?? "");
    }

    [SkippableTheory] // chaves corrigidas (valor exato depois do patch)
    [InlineData("documenteditor", "DE.Controllers.Main.txtStyle_Heading_1", "Título 1")]
    [InlineData("documenteditor", "DE.Controllers.Main.txtStyle_Heading_9", "Título 9")]
    [InlineData("documenteditor", "DE.Controllers.Statusbar.textTrackChanges", "O documento foi aberto com o modo Controlar Alterações ativado")]
    [InlineData("documenteditor", "DE.Controllers.Statusbar.textHasChanges", "Novas alterações foram controladas")]
    [InlineData("documenteditor", "DE.Controllers.Main.txtArt", "Seu texto aqui")]
    [InlineData("spreadsheeteditor", "SSE.Controllers.Main.txtSheet", "Planilha")]
    [InlineData("spreadsheeteditor", "SSE.Views.FileMenuPanels.CreateNew.txtBlank", "Pasta de trabalho em branco")]
    [InlineData("spreadsheeteditor", "SSE.Views.Toolbar.textTabTableDesign", "Design da Tabela")]
    [InlineData("spreadsheeteditor", "SSE.Views.DataValidationDialog.txtDate", "Data")]
    [InlineData("spreadsheeteditor", "SSE.Views.DataValidationDialog.txtTime", "Hora")]
    [InlineData("spreadsheeteditor", "SSE.Views.DocumentHolder.txtTime", "Hora")]
    [InlineData("spreadsheeteditor", "SSE.Views.DataValidationDialog.txtTextLength", "Comprimento do texto")]
    [InlineData("spreadsheeteditor", "SSE.Views.DataValidationDialog.txtOther", "Personalizado")]
    [InlineData("spreadsheeteditor", "SSE.Controllers.Main.txtRowLbls", "Rótulos de Linha")]
    [InlineData("spreadsheeteditor", "SSE.Controllers.Main.txtStyle_Bad", "Incorreto")]
    [InlineData("spreadsheeteditor", "SSE.Controllers.LeftMenu.textItemEntireCell", "Coincidir conteúdo da célula inteira")]
    [InlineData("spreadsheeteditor", "Common.Views.SearchPanel.textItemEntireCell", "Coincidir conteúdo da célula inteira")]
    [InlineData("spreadsheeteditor", "SSE.Views.FileMenuPanels.DocumentInfo.txtComment", "Comentário")]
    [InlineData("spreadsheeteditor", "SSE.Views.DocumentHolder.txtCellFormat", "Formatar células")]
    [InlineData("spreadsheeteditor", "SSE.Views.ProtectDialog.txtFormatCells", "Formatar células")]
    [InlineData("spreadsheeteditor", "SSE.Views.Toolbar.textFormatCells", "Formatar células")]
    [InlineData("spreadsheeteditor", "SSE.Views.ProtectDialog.txtFormatCols", "Formatar colunas")]
    [InlineData("spreadsheeteditor", "SSE.Views.TableDesignTab.txtSlicer", "Segmentação de Dados")]
    public void Locale_Corrigido(string app, string chave, string esperado) => Assert.Equal(esperado, Locale(app)[chave]);

    [SkippableFact] // a mensagem de planilha protegida fala em "planilha", nao em "pagina"
    public void Planilha_ProtegidaFalaEmPlanilha()
    {
        var v = Locale("spreadsheeteditor")["SSE.Controllers.Main.errorChangeOnProtectedSheet"];
        Assert.Contains("planilha protegida", v); Assert.Contains("desproteja a planilha", v); Assert.DoesNotContain("página", v);
    }

    [SkippableFact] // nenhum valor do pt.json da planilha fala em "Folha"/"folha" (59 hoje), nem sobrou "Tempo"/"Mau"/"Fatiador"
    public void Planilha_SemFolha()
    {
        var l = Locale("spreadsheeteditor");
        Assert.Empty(l.Where(p => Regex.IsMatch(p.Value, "folha", RegexOptions.IgnoreCase)).Select(p => p.Key));
        Assert.Empty(l.Where(p => p.Value is "Tempo" or "Mau" or "Fatiador" or "Formatar celulas").Select(p => p.Key));
    }

    [SkippableFact] // o documento nao tem "Cabeçalho N" como estilo de titulo (cabecalho de pagina nao e' afetado)
    public void Documento_TitulosNaoViramCabecalho()
    {
        var l = Locale("documenteditor");
        Assert.Empty(l.Where(p => p.Key.StartsWith("DE.Controllers.Main.txtStyle_Heading_", StringComparison.Ordinal) && p.Value.StartsWith("Cabeçalho")).Select(p => p.Key));
    }

    [SkippableFact] // formula-lang pt-br: VERDADEIRO/FALSO, #NOME?, #N/D, #VALOR!, #NÚM!, #NULO!, DATADIF
    public void Formulas_PtBr()
    {
        var arq = Path.Combine(Apps, "spreadsheeteditor", "main", "resources", "formula-lang", "pt-br.json");
        Skip.If(!File.Exists(arq), "motor nao obtido");
        using var d = JsonDocument.Parse(File.ReadAllBytes(arq));
        var r = d.RootElement.GetProperty("LocalFormulaOperands"); // constantes e erros ficam aqui; as funcoes (DATEDIF) na raiz
        Assert.Equal("VERDADEIRO", r.GetProperty("CONST_TRUE_FALSE").GetProperty("t").GetString());
        Assert.Equal("FALSO", r.GetProperty("CONST_TRUE_FALSE").GetProperty("f").GetString());
        var e = r.GetProperty("CONST_ERROR");
        Assert.Equal("#NULO!", e.GetProperty("nil").GetString());
        Assert.Equal("#DIV/0!", e.GetProperty("div").GetString());
        Assert.Equal("#VALOR!", e.GetProperty("value").GetString());
        Assert.Equal("#REF!", e.GetProperty("ref").GetString());
        Assert.Equal(@"#NAME\?", e.GetProperty("name").GetString()); // como o original (display em ingles): "#NOME?"/"#NOME(?=\?)" gravam cru, aceitam #NOME/#NOM ou exibem a regex (Task 9, fixes 1 e 2)
        Assert.Equal("#NÚM!", e.GetProperty("num").GetString());
        Assert.Equal("#N/D", e.GetProperty("na").GetString());
        Assert.Equal("DATADIF", d.RootElement.GetProperty("DATEDIF").GetString());
    }
}
