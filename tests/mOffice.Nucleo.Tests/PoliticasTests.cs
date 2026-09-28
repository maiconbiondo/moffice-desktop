using mOffice.Nucleo.Egresso;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 3 (S2 — egresso zero), camada 2: as politicas de registro do WebView2. A lista vive no C# porque
/// quem as aplica em producao e' o INSTALADOR; tools/egresso/politicas.ps1 e' a mesma lista para a bancada
/// de medida, e um teste aqui prova que os dois nao divergiram.
/// </summary>
public class PoliticasTests
{
    [Fact] // a politica que corta o servico de experimentacao/configuracao esta na lista com valor 0 (RestrictedMode)
    public void Recomendadas_ContemExperimentationRestrito()
    {
        var p = Politicas.Recomendadas.Single(x => x.Nome == "ExperimentationAndConfigurationServiceControl");
        Assert.Equal(@"HKLM\SOFTWARE\Policies\Microsoft\Edge\WebView2", p.Chave);
        Assert.Equal(0, p.Valor);
    }

    [Fact] // os comandos gerados sao `reg add` idempotentes (/f) — o instalador os executa
    public void ComandosReg_SaoIdempotentes()
    {
        var c = Politicas.ComandosReg();
        Assert.Contains("reg add \"HKLM\\SOFTWARE\\Policies\\Microsoft\\Edge\\WebView2\" /v ExperimentationAndConfigurationServiceControl /t REG_DWORD /d 0 /f", c);
    }

    [Fact] // a desinstalacao tem de DESFAZER o que a instalacao fez: um `reg delete` por politica aplicada
    public void ComandosRegRemover_CobremTodasAsPoliticas()
    {
        var c = Politicas.ComandosRegRemover();
        foreach (var p in Politicas.Recomendadas)
            Assert.Contains($"reg delete \"{p.Chave}\" /v {p.Nome} /f", c);
        Assert.Equal(Politicas.Recomendadas.Count, c.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length);
    }

    [Fact] // guarda de divergencia: o script da bancada aplica EXATAMENTE as politicas desta lista, nem mais nem menos
    public void ScriptDaBancada_NaoDivergiuDaLista()
    {
        var script = Path.Combine(RaizDoRepositorio, "tools", "egresso", "politicas.ps1");
        Assert.True(File.Exists(script), $"script nao encontrado: {script}");
        var texto = File.ReadAllText(script);
        foreach (var p in Politicas.Recomendadas)
        {
            Assert.Contains(p.Nome, texto);
            Assert.Contains(p.Chave, texto);
        }
        // O caminho inverso (nada no script fora da lista): nenhum outro nome de valor sob a chave de politica.
        var nomesNoScript = System.Text.RegularExpressions.Regex.Matches(texto, @"/v\s+(\w+)")
            .Select(m => m.Groups[1].Value).Distinct().OrderBy(n => n);
        Assert.Equal(Politicas.Recomendadas.Select(p => p.Nome).OrderBy(n => n), nomesNoScript);
    }

    /// <summary>Raiz do repositorio (a pasta do mOffice.slnx) - nao depende do motor, que fica fora do git.</summary>
    private static string RaizDoRepositorio => FronteiraTests.RaizRepo();
}
