using mOffice.Nucleo;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class HostFumacaTests
{
    [SkippableFact] // janela real: host object le um arquivo via base64 E a guarda bloqueia o fetch externo (403 + log)
    public void HostObject_LeArquivo_EGuardaBloqueiaEgresso()
    {
        Skip.If(JanelaSta.SessaoRemota, "teste de janela real: só em sessão local (não RDP)");
        var arquivo = Path.Combine(Path.GetTempPath(), $"moffice-{Guid.NewGuid():N}.txt");
        File.WriteAllText(arquivo, "ola");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela(); j.Show();
            await j.ProntoAsync();
            // "arquivos" nao vai mais para a janela por padrao (fix review 2); so este teste de fumaca precisa
            // dele, para a pagina Ponte/ponte-teste.html exercitar LerBase64/GravarBase64/GuardaEgresso.
            j.ExporArquivosParaTeste();
            await j.NavegarAsync(new Uri("https://moffice-poc.local/ponte-teste.html?arquivo=" + Uri.EscapeDataString(arquivo)));
            string resultado = "";
            for (var i = 0; i < 50 && !resultado.StartsWith("\"lido:"); i++) { await Task.Delay(100); resultado = await j.ExecutarJsAsync("window.resultado"); }
            Assert.Equal("\"lido:ola\"", resultado);
            string egresso = "";
            for (var i = 0; i < 50 && egresso != "\"status:403\""; i++) { await Task.Delay(100); egresso = await j.ExecutarJsAsync("window.egresso"); }
            // Prova o 403 de verdade (nao so que a URI foi anotada em Bloqueadas): o fetch da pagina recebeu
            // a resposta que a guarda construiu.
            Assert.Equal("\"status:403\"", egresso);
            Assert.Contains(j.Guarda.Bloqueadas, u => u.StartsWith("https://example.com/"));
            Assert.DoesNotContain(j.Guarda.Permitidas, u => u.Contains("example.com"));
            // Observado (revisao 2026-09-16 §3): o subrecurso ponte-teste.js carrega normalmente (a pagina
            // roda), mas NEM ele nem a navegacao de nivel superior para moffice-poc.local aparecem em
            // Guarda.Permitidas nem no log — ver o comentario de classe de GuardaEgresso. Por isso nao
            // afirmamos aqui que a guarda audita esse trafego; so que ele nao e bloqueado.
            Assert.Equal("true", await j.ExecutarJsAsync("window.subrecursoCarregado"));
            j.Close();
        });
        File.Delete(arquivo);
    }
}
