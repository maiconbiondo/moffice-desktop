using System.Diagnostics;

namespace mOffice.Nucleo.Tests;

/// <summary>Os <c>.orig</c> que o aplicar-patches.ps1 deixa no motor (copia do ORIGINAL, que inclui a fonte ASCW3 da Ascensio
/// que o patch da fonte existe para manter fora) nunca podem ir para o pacote/instalador. Tres camadas: a copia exclui,
/// a conferencia do pacote aborta, o .iss exclui.</summary>
public class PacoteSemOrigTests
{
    private static string Raiz => FronteiraTests.RaizRepo();
    private static string Pacote => File.ReadAllText(Path.Combine(Raiz, "tools", "installer", "montar-pacote.ps1"));

    [Fact] // camada 3: o instalador exclui *.orig
    public void Iss_ExcluiOrig()
    {
        var linha = File.ReadAllLines(Path.Combine(Raiz, "tools", "installer", "moffice.iss"))
            .Single(l => l.StartsWith("Source: \"{#PacoteDir}", StringComparison.Ordinal));
        Assert.Contains("*.orig", linha);
    }

    [Fact] // camada 2: a conferencia final do pacote reprova .orig
    public void MontarPacote_ConferenciaReprovaOrig() =>
        Assert.Matches(@"\$proibidos\s*=.*\.orig", Pacote.Replace("\r", ""));

    [Fact] // camada 1: a copia da arvore do motor nao leva .orig (roda a funcao REAL num arvore temporaria)
    public void CopiarArvore_NaoCopiaOrig()
    {
        var tmp = Directory.CreateTempSubdirectory("moffice-copia-").FullName;
        try
        {
            var origem = Path.Combine(tmp, "o"); Directory.CreateDirectory(Path.Combine(origem, "sdkjs", "word"));
            File.WriteAllText(Path.Combine(origem, "sdkjs", "word", "sdk-all.js"), "x");
            File.WriteAllText(Path.Combine(origem, "sdkjs", "word", "sdk-all.js.orig"), "ASCW3");
            File.WriteAllText(Path.Combine(origem, "index.html.orig"), "orig");
            File.WriteAllText(Path.Combine(origem, "index.html"), "novo");
            var destino = Path.Combine(tmp, "d");
            var cmd = "$ast=[System.Management.Automation.Language.Parser]::ParseFile('" + Path.Combine(Raiz, "tools", "installer", "montar-pacote.ps1") + "',[ref]$null,[ref]$null);"
                + "$f=$ast.Find({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Copiar-Arvore'},$true);"
                + "Invoke-Expression $f.Extent.Text; Copiar-Arvore '" + origem + "' '" + destino + "'";
            var psi = new ProcessStartInfo("powershell.exe") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
            foreach (var a in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd }) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi)!;
            var so = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
            Assert.True(p.WaitForExit(60_000));
            Assert.True(p.ExitCode == 0, so);
            Assert.True(File.Exists(Path.Combine(destino, "sdkjs", "word", "sdk-all.js")));
            Assert.True(File.Exists(Path.Combine(destino, "index.html")));
            Assert.Empty(Directory.GetFiles(destino, "*.orig", SearchOption.AllDirectories));
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }
}
