using mOffice.Conversor;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 5, Task 1 — o PACOTE montado por <c>tools/installer/montar-pacote.ps1</c> basta para abrir um
/// documento: motor enxuto (sem <c>sdkjs-plugins</c>, so' os dicionarios pt_BR e en_US, sem <c>ASC.ttf</c>, sem o
/// <c>AllFonts.js</c> da maquina de build) e o runtime WebView2 do pacote.
///
/// Opt-in: <c>MOFFICE_PACOTE=&lt;pasta do pacote&gt;</c>. Rode SOZINHO
/// (<c>dotnet test --filter FullyQualifiedName~PacoteTests</c>): o teste aponta <c>MOFFICE_MOTOR_DIR</c> para
/// <c>&lt;pacote&gt;\motor</c> no processo, e o runtime fixo e' calculado uma vez por processo
/// (<see cref="AmbienteWebView2.PastaRuntimeFixo"/>) — numa suite inteira ele ja' teria sido fixado no
/// runtime de desenvolvimento antes deste teste.
/// </summary>
[Collection(JanelaReal.Nome)]
public class PacoteTests
{
    public const string VariavelPacote = "MOFFICE_PACOTE";

    private readonly ITestOutputHelper _saida;
    public PacoteTests(ITestOutputHelper saida) { _saida = saida; }

    [SkippableFact]
    public void PacoteMontado_AbreDocumentoPeloMotorDoPacote()
    {
        var pacote = Environment.GetEnvironmentVariable(VariavelPacote);
        Skip.If(string.IsNullOrWhiteSpace(pacote), $"teste do pacote: defina {VariavelPacote}=<pasta montada por tools/installer/montar-pacote.ps1>");
        pacote = Path.GetFullPath(pacote!);
        var motorDoPacote = Path.Combine(pacote, "motor");

        // O subconjunto que este teste diz provar TEM de ser o subconjunto enxuto — senao "abriu" nao prova nada.
        Assert.True(File.Exists(Path.Combine(motorDoPacote, "converter", "x2t.exe")), "pacote sem motor\\converter\\x2t.exe");
        Assert.False(Directory.Exists(Path.Combine(motorDoPacote, "editors", "sdkjs-plugins")), "o pacote nao deveria ter sdkjs-plugins");
        Assert.Equal(new[] { "en_US", "pt_BR" }, Directory.GetDirectories(Path.Combine(motorDoPacote, "dictionaries"))
            .Select(d => Path.GetFileName(d)!).Order(StringComparer.Ordinal).ToArray());
        Assert.Empty(Directory.GetFiles(pacote, "ASC.ttf", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(pacote, "ponte-teste.*", SearchOption.AllDirectories));
        // <script src> estatico do AllFonts.js, fora do alcance da reescrita do CatalogoNoEditor
        Assert.Empty(Directory.GetFiles(motorDoPacote, "preload.html", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(motorDoPacote, "cache-scripts.html", SearchOption.AllDirectories));
        var allFontsNoMotor = Path.Combine(motorDoPacote, "editors", "sdkjs", "common", "AllFonts.js");
        Assert.False(File.Exists(allFontsNoMotor), "o pacote nao deveria trazer o AllFonts.js da maquina de build");

        var anterior = Environment.GetEnvironmentVariable(ConfiguracaoMotor.VariavelMotor);
        Environment.SetEnvironmentVariable(ConfiguracaoMotor.VariavelMotor, motorDoPacote);
        var pasta = Directory.CreateTempSubdirectory("moffice-pacote-").FullName;
        try
        {
            Assert.Equal(motorDoPacote, ConfiguracaoMotor.Localizar().PastaMotor, ignoreCase: true);
            EditorReal.PularSeNaoDaParaAbrirJanela();

            var runtime = AmbienteWebView2.PastaRuntimeFixo;
            _saida.WriteLine($"runtime fixo: {runtime}");
            Assert.NotNull(runtime);
            Assert.StartsWith(Path.Combine(pacote, "WebView2Runtime"), runtime!, StringComparison.OrdinalIgnoreCase);

            var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Pacote do mOffice 2026");
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela();
                j.Show();
                await j.AbrirAsync(docx);
                var pronto = await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300);
                Assert.Equal("true", pronto);
                var texto = await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor);
                Assert.Contains("Pacote do mOffice 2026", texto);
                // fix round 1: o AllFonts.js veio do catalogo pelo host, nao de dentro do motor
                Assert.Contains(CatalogoNoEditor.UrlAllFonts, j.Guarda.Permitidas);
                j.Close();
            });
            Assert.False(File.Exists(allFontsNoMotor), "a janela nao pode escrever no motor do pacote");
        }
        finally
        {
            Environment.SetEnvironmentVariable(ConfiguracaoMotor.VariavelMotor, anterior);
            try { Directory.Delete(pasta, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }
}
