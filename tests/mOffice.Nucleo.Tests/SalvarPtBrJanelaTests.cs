using DocumentFormat.OpenXml.Packaging;
using mOffice.Conversor;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class SalvarPtBrJanelaTests
{
    /// <summary>O VALOR EM CACHE (&lt;v&gt;) de uma celula — o que o mPlanilha mostra ao reabrir sem recalcular.</summary>
    private static string Cache(string xlsx, string referencia)
    {
        using var d = SpreadsheetDocument.Open(xlsx, false);
        var c = d.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<S.Cell>().First(x => x.CellReference == referencia);
        return c.CellValue?.Text ?? "";
    }

    [SkippableFact] // PL-01: TEXTO(DATA(...);"mmmm") e concatenacao com numero gravam o cache em pt-BR
    public void Planilha_FormulasDeTexto_CacheEmPortugues()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-ptbr-").FullName;
        var destino = Path.Combine(pasta, "ptbr.xlsx");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha"); j.Show();
                await j.NovoAsync(FormatoDocumento.Xlsx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                await j.ExecutarJsAsync(EditorReal.InserirTextoCell("=TEXTO(DATA(2026;12;31);\"mmmm\")", "A1"));
                await j.ExecutarJsAsync(EditorReal.InserirTextoCell("=\"x\"&ARRED(10/7;2)", "A2"));
                await j.SalvarComoAsync(destino, FormatoDocumento.Xlsx);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Equal("dezembro", Cache(destino, "A1"));
            Assert.Equal("x1,43", Cache(destino, "A2"));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Pilha de undo do editor (tamanho das listas do historico + canUndo): a restauracao do "modificado" nao
    /// pode acrescentar passo nenhum.</summary>
    private static readonly string EstadoDoEditor = EditorReal.NoEditor("""
        var H = w.AscCommon.History; var pts = {};
        for (var k in H) if (Array.isArray(H[k])) pts[k] = H[k].length;
        return JSON.stringify({ pts: pts, canUndo: api.asc_getCanUndo() });
        """);

    [SkippableFact] // gravacao parcial (5b): Ctrl+S do editor -> UM aviso pt-BR na faixa, o temporario sobrevive, o original esta no .moffice-bak
    public void GravacaoParcial_CtrlS_AvisaUmaVez_ETemporarioSobrevive()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-5b-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var bak = docx + GravacaoNoDestino.SufixoBackup; var tmp = docx + ".moffice-tmp";
        var avisos = new List<string>(); var avisosDoFechamento = new List<string>();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.AvisoAoUsuario += a => avisos.Add(a);
                j.ErroAoFechar += a => avisosDoFechamento.Add(a);
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" parcial"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                // 1177 simulado + o retorno do backup impossivel (um diretorio no lugar do destino)
                j.SubstituirParaTeste = (_, alvo, b) => { File.Move(alvo, b); Directory.CreateDirectory(alvo); throw new IOException("1177 simulado"); };
                await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"));
                for (var i = 0; i < 100 && avisos.Count == 0; i++) await Task.Delay(100);
                await Task.Delay(500);
                Assert.True(File.Exists(tmp), "o temporario (versao nova) foi apagado com o destino ausente");
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Single(avisos);
            Assert.Equal($"O arquivo original ficou em {bak}. A versão nova continua aberta no mOffice: salve de novo ou use Salvar como.", avisos[0]);
            Assert.Empty(avisosDoFechamento);                       // o caminho do Ctrl+S nao passa pelo ErroAoFechar: sem duplicar
            Assert.Contains("Texto base", mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(bak));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableTheory] // PL-02: depois de exportar PDF o editor volta a "modificado" + CanSave (undo intacto) e o Ctrl+S (SalvarAsync) grava o documento
    [InlineData("mtexto", FormatoDocumento.Docx)]
    [InlineData("mplanilha", FormatoDocumento.Xlsx)]
    [InlineData("mslides", FormatoDocumento.Pptx)]
    public void ExportarPdf_DeixaOEditorModificado_ESalvarGrava(string app, FormatoDocumento f)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-pl02-").FullName;
        var doc = Path.Combine(pasta, "doc." + f.Extensao());
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela(app); j.Show();
                await j.NovoAsync(f);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                await j.SalvarComoAsync(doc, f);
                await j.ExecutarJsAsync(EditorReal.InserirPorDocType(f.DocType(), "depois-do-pdf"));
                var antes = File.GetLastWriteTimeUtc(doc);
                var pilhaAntes = await j.ExecutarJsAsync(EstadoDoEditor);
                await j.SalvarComoAsync(Path.Combine(pasta, "export.pdf"), FormatoDocumento.Pdf);
                // PL-02: o editor volta a "modificado" E a "pode salvar" (botao Salvar ligado), sem passo de undo novo
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.NoEditor(
                    "return String(api.isDocumentModified() && api.asc_isDocumentCanSave());"), "true", 50));
                Assert.Equal(pilhaAntes, await j.ExecutarJsAsync(EstadoDoEditor));
                await Task.Delay(1100); // resolucao do carimbo de escrita
                await j.SalvarAsync();
                Assert.True(File.GetLastWriteTimeUtc(doc) > antes, "o documento nao foi regravado depois da exportacao");
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
