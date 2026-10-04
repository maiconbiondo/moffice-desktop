using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Principal;
using mOffice.Conversor;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 5, Task 1, fix round 1: o <c>AllFonts.js</c> chega ao editor pelo host do catalogo
/// (<see cref="CatalogoNoEditor"/>), sem existir dentro do motor. Unitarios do mapeamento URL -> arquivo.</summary>
public class CatalogoNoEditorTests
{
    private const string Pasta = @"C:\catalogo-de-teste";

    [Fact] // a URL exata -> <pasta>\AllFonts.js
    public void Resolver_UrlExata_DevolveOArquivoDoCatalogo()
        => Assert.Equal(Path.Combine(Pasta, "AllFonts.js"), CatalogoNoEditor.Resolver(CatalogoNoEditor.UrlAllFonts, Pasta));

    [Theory] // qualquer outra coisa -> null (e nada da requisicao vira caminho)
    [InlineData("https://moffice-fontes.local/AllFonts.js?v=1")]
    [InlineData("https://moffice-fontes.local/AllFonts.js#x")]
    [InlineData("https://moffice-fontes.local/allfonts.js")]
    [InlineData("https://moffice-fontes.local/AllFonts.js/")]
    [InlineData("https://moffice-fontes.local/")]
    [InlineData("https://moffice-fontes.local/font_selection.bin")]
    [InlineData("https://moffice-fontes.local/C:/Windows/win.ini")]
    [InlineData("https://moffice-fontes.local/%2e%2e/AllFonts.js")]
    [InlineData("https://moffice-fontes.local/sub/AllFonts.js")]
    [InlineData("http://moffice-fontes.local/AllFonts.js")]
    [InlineData("https://moffice.local/sdkjs/common/AllFonts.js")]
    [InlineData("ascdesktop://fonts/AllFonts.js")]
    [InlineData("")]
    public void Resolver_OutrasUrls_Null(string uri) => Assert.Null(CatalogoNoEditor.Resolver(uri, Pasta));

    [Theory]
    [InlineData("https://moffice-fontes.local/AllFonts.js", true)]
    [InlineData("https://MOFFICE-FONTES.local/qualquer", true)]
    [InlineData("https://moffice-fontes.localhost/AllFonts.js", false)]
    [InlineData("https://moffice.local/sdkjs/common/AllFonts.js", false)]
    public void EhDoHost(string uri, bool esperado) => Assert.Equal(esperado, CatalogoNoEditor.EhDoHost(uri));

    [Fact] // a guarda de egresso deixa o host passar (e so' ele entre os novos); o script aponta para a URL servida
    public void GuardaEScriptApontamParaOMesmoHost()
    {
        Assert.Contains(CatalogoNoEditor.Origem, GuardaEgresso.OrigensPermitidas);
        Assert.Contains("\"" + CatalogoNoEditor.UrlAllFonts + "\"", CatalogoNoEditor.ScriptReescrita);
        Assert.Contains("u.hostname === \"moffice.local\"", CatalogoNoEditor.ScriptReescrita);
        Assert.Contains(@"\/sdkjs\/common\/AllFonts\.js$", CatalogoNoEditor.ScriptReescrita);
    }
}

/// <summary>Janela real sobre um motor SEM <c>editors\sdkjs\common\AllFonts.js</c> e com essa pasta SO' LEITURA
/// (ACL que nega escrita ao usuario atual — o caso do Program Files): o editor tem de carregar as fontes pelo
/// host do catalogo, e nada pode ser escrito no motor.</summary>
[Collection(JanelaReal.Nome)]
public class CatalogoNoEditorJanelaTests
{
    private readonly ITestOutputHelper _saida;
    public CatalogoNoEditorJanelaTests(ITestOutputHelper saida) { _saida = saida; }

    [SkippableFact]
    public void MotorSemAllFontsESoLeitura_EditorCarregaAsFontesPeloHost()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var real = mOffice.Conversor.Tests.MotorDeTeste.Localizar();
        var pasta = Directory.CreateTempSubdirectory("moffice-catalogo-host-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Catalogo pelo host 2026");
        using var motor = MotorSoLeituraSemAllFonts.Criar(real);
        var comum = Path.Combine(motor.Pasta, "editors", "sdkjs", "common");
        var noMotor = Path.Combine(comum, "AllFonts.js");
        try
        {
            // Controles do arranjo: o arquivo NAO esta' no motor e a pasta REALMENTE nega escrita.
            Assert.False(File.Exists(noMotor));
            Assert.Throws<UnauthorizedAccessException>(() => File.WriteAllText(Path.Combine(comum, "sonda-escrita.txt"), "x"));

            var config = new ConfiguracaoMotor(motor.Pasta);
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.PrepararParaTeste(new JanelaEditor("teste-catalogo-host", config));
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Contains("Catalogo pelo host 2026", await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor));

                // o catalogo veio pela URL do host (o motor nao tem o arquivo) ...
                Assert.Contains(CatalogoNoEditor.UrlAllFonts, j.Guarda.Permitidas);
                // ... e o editor o CONSUMIU: os caminhos que ele pede por ascdesktop://fonts/ saem do __fonts_files
                // do AllFonts.js (o sdkjs apaga as globais depois de ler — `delete a.__fonts_files` —, entao elas
                // nao servem de prova depois do carregamento).
                var fontes = j.Guarda.Permitidas.Count(u => u.StartsWith("ascdesktop://fonts/", StringComparison.OrdinalIgnoreCase));
                _saida.WriteLine("fontes pedidas por ascdesktop://: " + fontes);
                Assert.True(fontes > 0, "o editor nao pediu nenhuma fonte do catalogo");
                Assert.Empty(j.Guarda.Bloqueadas);
                j.Close();
            });
            Assert.False(File.Exists(noMotor), "nada pode ser escrito no motor");
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
}

/// <summary>Motor temporario (padrao do <c>MotorDeControle</c> de <c>FonteSimbolosTests</c>): <c>converter</c>,
/// <c>fonts</c>, <c>dictionaries</c> como juncoes; <c>editors</c> COPIADO (o mapeamento de pasta nao segue juncao)
/// so' com o que o editor de texto pede, SEM <c>sdkjs\common\AllFonts.js</c>, e com <c>editors</c> negando
/// escrita ao usuario atual (ACL herdavel). O Dispose tira a negacao antes de apagar.</summary>
internal sealed class MotorSoLeituraSemAllFonts : IDisposable
{
    public string Pasta { get; }
    private readonly List<string> _juncoes = [];
    private FileSystemAccessRule? _negacao;
    private string Editors => Path.Combine(Pasta, "editors");

    private MotorSoLeituraSemAllFonts(string pasta) { Pasta = pasta; }

    public static MotorSoLeituraSemAllFonts Criar(ConfiguracaoMotor real)
    {
        var m = new MotorSoLeituraSemAllFonts(Directory.CreateTempSubdirectory("moffice-motor-sem-allfonts-").FullName);
        try
        {
            foreach (var d in new[] { "converter", "fonts", "dictionaries" }) m.Juncao(Path.Combine(real.PastaMotor, d), Path.Combine(m.Pasta, d));
            var editorsReal = Path.Combine(real.PastaMotor, "editors");
            foreach (var p in new[] { "web-apps", "webext", Path.Combine("sdkjs", "common"), Path.Combine("sdkjs", "word") })
                Copiar(Path.Combine(editorsReal, p), Path.Combine(m.Editors, p));
            var allFonts = Path.Combine(m.Editors, "sdkjs", "common", "AllFonts.js");
            if (File.Exists(allFonts)) File.Delete(allFonts);

            var sid = WindowsIdentity.GetCurrent().User!;
            m._negacao = new FileSystemAccessRule(sid,
                FileSystemRights.CreateFiles | FileSystemRights.AppendData | FileSystemRights.WriteData | FileSystemRights.WriteAttributes
                | FileSystemRights.WriteExtendedAttributes | FileSystemRights.Delete | FileSystemRights.DeleteSubdirectoriesAndFiles,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Deny);
            var di = new DirectoryInfo(m.Editors);
            var acl = di.GetAccessControl();
            acl.AddAccessRule(m._negacao);
            di.SetAccessControl(acl);
            return m;
        }
        catch { m.Dispose(); throw; }
    }

    private static void Copiar(string origem, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var f in Directory.GetFiles(origem))
            if (!f.EndsWith(".orig", StringComparison.OrdinalIgnoreCase)) File.Copy(f, Path.Combine(destino, Path.GetFileName(f)));
        foreach (var d in Directory.GetDirectories(origem)) Copiar(d, Path.Combine(destino, Path.GetFileName(d)));
    }

    private void Juncao(string alvo, string link)
    {
        var psi = new ProcessStartInfo("cmd.exe", $"/c mklink /J \"{link}\" \"{alvo}\"") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
        using var p = Process.Start(psi)!;
        var saida = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
        p.WaitForExit();
        if (p.ExitCode != 0 || !Directory.Exists(link)) throw new InvalidOperationException($"mklink /J falhou para {link}: {saida}");
        _juncoes.Add(link);
    }

    public void Dispose()
    {
        if (_negacao is not null && Directory.Exists(Editors))
        {
            try
            {
                var di = new DirectoryInfo(Editors);
                var acl = di.GetAccessControl();
                acl.RemoveAccessRule(_negacao);
                di.SetAccessControl(acl);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return; } // nao arrisca apagar pela metade
        }
        // Juncoes primeiro, sem recursao: apagar recursivo por cima de uma juncao desceria no motor real.
        foreach (var j in _juncoes)
            try { if (Directory.Exists(j)) Directory.Delete(j, recursive: false); } catch (IOException) { }
        if (_juncoes.Any(Directory.Exists)) return;
        try { Directory.Delete(Pasta, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
    }
}
