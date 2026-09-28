namespace mOffice.Nucleo.Tests;

/// <summary>A constante <see cref="Motor.Versao"/> e' o que o dialogo Sobre mostra (o motor INSTALADO nao
/// carrega o VERSAO.txt do repo). Nesta maquina de desenvolvimento, porem, third_party/onlyoffice existe —
/// entao o teste amarra a constante ao arquivo de verdade; numa maquina sem o motor extraido ele e' pulado.</summary>
public class MotorTests
{
    private static string RaizRepo()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d is not null && !File.Exists(Path.Combine(d.FullName, "mOffice.slnx"))) d = d.Parent;
        return d?.FullName ?? throw new InvalidOperationException("mOffice.slnx não encontrado acima do assembly de testes");
    }

    [SkippableFact] // a constante bate com third_party/onlyoffice/VERSAO.txt quando o motor esta extraido
    public void Versao_BateComVersaoTxt()
    {
        var arquivo = Path.Combine(RaizRepo(), "third_party", "onlyoffice", "VERSAO.txt");
        Skip.If(!File.Exists(arquivo), "third_party/onlyoffice/VERSAO.txt ausente — rode tools/onlyoffice/obter-motor.ps1");
        Assert.Contains(Motor.Versao, File.ReadAllText(arquivo));
    }
}
