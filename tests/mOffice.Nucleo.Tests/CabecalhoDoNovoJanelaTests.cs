using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using mOffice.Conversor;
using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// I1 da revisao final do Plano 6 (A10, backlog 71) na janela real: o cabecalho do editor de um documento NOVO virava <c>new.docx</c> (ou o
/// nome do modelo) depois da 1a edicao, porque o sdkjs rele' <c>LocalFileGetSourcePath</c> (a copia do modelo na recovery). MEDIDO ao vivo
/// (porta 9622): com a origem "Documento1.docx" o sdkjs poe "Documento1.docx" no cabecalho; a ponte devolve "Documento1" (como o Word).
/// </summary>
[Collection(JanelaReal.Nome)]
public class CabecalhoDoNovoJanelaTests
{
    private const string Cabecalho = "window.frames[0].document.querySelector('#title-doc-name').value";
    private const string Origem = "window.frames[0].AscDesktopEditor.LocalFileGetSourcePath()";

    private static string NovaPasta(string prefixo) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, prefixo + Guid.NewGuid().ToString("N"))).FullName;

    private static async Task<string> Js(JanelaEditor j, string expr) =>
        (await j.ExecutarJsAsync("(function(){try{return String(" + expr + ");}catch(e){return 'erro: '+e;}})()")).Trim('"');

    /// <summary>Edita, espera o "modificado" e uma janela maior que o ciclo do sdkjs, e devolve (cabecalho, origem vista pelo editor).</summary>
    private static async Task<(string Cabecalho, string Origem)> EditarELerAsync(JanelaEditor j, string doctype)
    {
        Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(doctype, "edicao-a10"))).Trim('"'));
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
        await Task.Delay(2500);
        return (await Js(j, Cabecalho), await Js(j, Origem));
    }

    [SkippableTheory] // Ctrl+N (NovoAsync) -> editar: o cabecalho continua "<Rotulo><n>", sem "new." e sem extensao, nos tres apps
    [InlineData("mtexto", FormatoDocumento.Docx, "word")]
    [InlineData("mplanilha", FormatoDocumento.Xlsx, "cell")]
    [InlineData("mslides", FormatoDocumento.Pptx, "slide")]
    public void Novo_AposEditar_CabecalhoSemNew(string app, FormatoDocumento f, string doctype)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string titulo = "", cab = "", origem = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela(app);
            try
            {
                j.Show();
                await j.NovoAsync(f);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                titulo = j.SessaoAtual!.Titulo;
                (cab, origem) = await EditarELerAsync(j, doctype);
            }
            finally { j.Close(); }
        }, TimeSpan.FromSeconds(240));
        Assert.Equal(titulo, cab);
        Assert.DoesNotContain("new.", cab, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("new.", Path.GetFileName(origem), StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact] // novo a partir de um modelo (.dotx): o cabecalho nao vira o nome do modelo depois de editar
    public void NovoDoModelo_AposEditar_CabecalhoSemNomeDoModelo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-a10-modelo-");
        var docx = FabricaDocumentos.Docx(pasta, "Texto do modelo");
        var dotx = Path.Combine(pasta, "modelo-ata.dotx");
        File.Move(docx, dotx);
        using (var d = WordprocessingDocument.Open(dotx, true)) d.ChangeDocumentType(WordprocessingDocumentType.Template);
        string titulo = "", cab = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                try
                {
                    j.Show();
                    await j.NovoDoModeloAsync(dotx);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    titulo = j.SessaoAtual!.Titulo;
                    (cab, _) = await EditarELerAsync(j, "word");
                }
                finally { j.Close(); }
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
        Assert.StartsWith("Documento", titulo);
        Assert.Equal(titulo, cab);
        Assert.DoesNotContain("modelo", cab, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact] // .ppsx RECUPERADO (LocalFileSetSourcePath antes do EndLoad) continua abrindo para edicao: nao inicia a apresentacao
    public void PpsxRecuperado_NaoApresenta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-a10-ppsx-");
        var pptx = FabricaDocumentos.Pptx(pasta, "Slide recuperado");
        var ppsx = Path.Combine(pasta, "show.ppsx");
        File.Move(pptx, ppsx);
        using (var d = PresentationDocument.Open(ppsx, true)) d.ChangeDocumentType(PresentationDocumentType.Slideshow);
        string recovery = "", origem = "";
        var telaCheia = true;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mslides"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
                await j.AbrirAsync(ppsx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                recovery = j.SessaoAtual!.PastaRecovery;
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType("slide", " recuperado-ppsx"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                await j.DescarregarChangesAsync();
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.True(Directory.Exists(recovery));
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mslides");
                try
                {
                    j.Show(); j.Activate();
                    await j.RecuperarAsync(recovery);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    await Task.Delay(2000);
                    telaCheia = j.EmTelaCheia;
                    origem = await Js(j, Origem);
                }
                finally
                {
                    if (j.EmTelaCheia) j.DefinirTelaCheia(false);
                    j.Close();
                }
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } try { if (recovery.Length > 0) Directory.Delete(recovery, true); } catch { } }
        Assert.False(telaCheia, "o .ppsx recuperado iniciou a apresentacao sozinho");
        Assert.EndsWith(".pptx", origem, StringComparison.OrdinalIgnoreCase);
    }
}
