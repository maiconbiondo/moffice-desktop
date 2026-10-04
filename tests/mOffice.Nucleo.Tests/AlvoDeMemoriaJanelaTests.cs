using Microsoft.Web.WebView2.Core;
using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 11 (PL-10) na janela real: sem foco o WebView2 vai para o alvo de memoria Low; com foco volta a Normal, e
/// o documento continua editavel. Tambem: um salvamento iniciado e a janela perdendo o foco no meio terminam bem.</summary>
[Collection(JanelaReal.Nome)]
public class AlvoDeMemoriaJanelaTests
{
    [SkippableFact] // Low sem foco, Normal com foco; o documento segue editavel (pronto + insercao) depois de voltar
    public void AlvoDeMemoria_BaixoSemFoco_NormalComFoco()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string[] lidos = [];
        string insercao = "", pronto = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            await j.NovoAsync(FormatoDocumento.Docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            var antes = j.Web.CoreWebView2.MemoryUsageTargetLevel;
            j.AplicarAlvoDeMemoria(false);
            var semFoco = j.Web.CoreWebView2.MemoryUsageTargetLevel;
            j.AplicarAlvoDeMemoria(true);
            var comFoco = j.Web.CoreWebView2.MemoryUsageTargetLevel;
            lidos = [antes.ToString(), semFoco.ToString(), comFoco.ToString()];
            pronto = await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 50);
            insercao = (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("depois do foco"))).Trim('"');
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.Equal(nameof(CoreWebView2MemoryUsageTargetLevel.Normal), lidos[0]);
        Assert.Equal(nameof(CoreWebView2MemoryUsageTargetLevel.Low), lidos[1]);
        Assert.Equal(nameof(CoreWebView2MemoryUsageTargetLevel.Normal), lidos[2]);
        Assert.Equal("true", pronto);
        Assert.StartsWith("ok", insercao);
    }

    [SkippableFact] // salvar iniciado e a janela perdendo o foco no meio: o arquivo sai inteiro e o documento segue editavel
    public void AlvoDeMemoria_SalvarComFocoPerdidoNoMeio_Termina()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-alvo-").FullName;
        var destino = Path.Combine(pasta, "salvo.docx");
        string insercao = "", nivelFinal = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.Show();
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("texto-do-salvar-cx9"))).Trim('"'));
                var salvando = j.SalvarComoAsync(destino, FormatoDocumento.Docx);   // comeca...
                j.AplicarAlvoDeMemoria(false);                                      // ...e a janela perde o foco no meio
                await salvando;
                j.AplicarAlvoDeMemoria(true);
                nivelFinal = j.Web.CoreWebView2.MemoryUsageTargetLevel.ToString();
                insercao = (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("editavel-depois"))).Trim('"');
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.True(File.Exists(destino) && new FileInfo(destino).Length > 0);
            using var zip = System.IO.Compression.ZipFile.OpenRead(destino);
            using var r = new StreamReader(zip.GetEntry("word/document.xml")!.Open());
            Assert.Contains("texto-do-salvar-cx9", r.ReadToEnd());
            Assert.Equal(nameof(CoreWebView2MemoryUsageTargetLevel.Normal), nivelFinal);
            Assert.StartsWith("ok", insercao);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // M5 da revisao final do Plano 6 (backlog 74): sem foco ANTES do "pronto" o alvo continua Normal (nada de Low no meio da
                    // carga); depois do "pronto", sem foco vai para Low
    public void AlvoDeMemoria_SemFocoAntesDoPronto_ContinuaNormal()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string antesDoPronto = "", depoisDoPronto = ""; var prontoAntes = true; string log = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            await j.ProntoAsync();
            var abrindo = j.NovoAsync(FormatoDocumento.Docx);
            prontoAntes = j.DocumentoDaSessaoPronto;
            j.AplicarAlvoDeMemoria(false);                                   // Deactivated durante a carga
            antesDoPronto = j.Web.CoreWebView2.MemoryUsageTargetLevel.ToString();
            await abrindo;
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            for (var i = 0; i < 50 && !j.DocumentoDaSessaoPronto; i++) await Task.Delay(100);
            Assert.True(j.DocumentoDaSessaoPronto);
            j.AplicarAlvoDeMemoria(false);                                   // Deactivated com o documento pronto
            depoisDoPronto = j.Web.CoreWebView2.MemoryUsageTargetLevel.ToString();
            j.AplicarAlvoDeMemoria(true);
            log = File.ReadAllText(j.ArquivoDeLog);
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.False(prontoAntes);
        Assert.Equal(nameof(CoreWebView2MemoryUsageTargetLevel.Normal), antesDoPronto);
        Assert.Equal(nameof(CoreWebView2MemoryUsageTargetLevel.Low), depoisDoPronto);
        Assert.Contains("MEMORIA alvo low adiado: a janela perdeu o foco com o documento carregando", log);
        Assert.Contains("MEMORIA alvo low (documento pronto)", log);
    }
}
