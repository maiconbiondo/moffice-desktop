namespace mOffice.Comum.Tests;

public class ManifestoTests
{
    [Fact] // SO-13: os tres exe declaram PerMonitorV2 pelo MESMO manifesto
    public void TresApps_PerMonitorV2()
    {
        var raiz = AplicativoOfficeTests.RaizRepo();
        Assert.Contains("<dpiAwareness xmlns=\"http://schemas.microsoft.com/SMI/2016/WindowsSettings\">PerMonitorV2</dpiAwareness>",
            File.ReadAllText(Path.Combine(raiz, "src", "app.manifest")));
        foreach (var app in new[] { "mTexto", "mPlanilha", "mSlides" })
            Assert.Contains(@"<ApplicationManifest>..\app.manifest</ApplicationManifest>", File.ReadAllText(Path.Combine(raiz, "src", app, app + ".csproj")));
    }
}
