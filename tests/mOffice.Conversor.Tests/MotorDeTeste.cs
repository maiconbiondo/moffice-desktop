using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

/// <summary>O motor ONLYOFFICE (<c>third_party/onlyoffice</c>) fica FORA do git: num clone recem-feito do repositorio
/// publico ele nao existe ate alguem rodar <c>tools/onlyoffice/obter-motor.ps1</c>. Os testes que precisam dele
/// chamam <see cref="Localizar"/> em vez de <see cref="ConfiguracaoMotor.Localizar()"/>: com o motor presente o
/// resultado e' o mesmo; sem ele, num CLONE PUBLICO, o teste e' PULADO (SkipException) em vez de falhar.
///
/// Na arvore de DESENVOLVIMENTO (existe <c>third_party/</c> na raiz do repositorio) o motor ausente e' FALHA: ali ele
/// tem de estar, e um skip silencioso deixaria a suite verde sem testar nada. <see cref="VariavelPermitir"/>=1 libera o
/// skip mesmo assim (maquina de desenvolvimento sem o motor, de proposito).
///
/// Os tres projetos de teste trocam <c>[Fact]</c>/<c>[Theory]</c> por <c>[SkippableFact]</c>/<c>[SkippableTheory]</c>
/// por alias global (csproj), para o SkipException valer em qualquer teste - inclusive quando ele vem de um helper
/// compartilhado. So' a AUSENCIA do motor vira skip: qualquer outro erro continua falhando.</summary>
public static class MotorDeTeste
{
    public const string MotivoAusente = "motor ONLYOFFICE ausente desta arvore (third_party/onlyoffice fora do git) - rode tools/onlyoffice/obter-motor.ps1";
    public const string VariavelPermitir = "MOFFICE_PERMITIR_SEM_MOTOR";

    public static ConfiguracaoMotor Localizar() => OuPular(ConfiguracaoMotor.Localizar);

    public static T OuPular<T>(Func<T> localizar) =>
        OuPular(localizar, RaizRepositorio(), Environment.GetEnvironmentVariable(VariavelPermitir) == "1");

    /// <summary>Costura de teste: raiz do repositorio e a liberacao entram por parametro.</summary>
    internal static T OuPular<T>(Func<T> localizar, string? raizRepositorio, bool permitirSemMotor)
    {
        try { return localizar(); }
        catch (ErroConversao e) when (e.Message.StartsWith("Motor ONLYOFFICE n", StringComparison.Ordinal))
        {
            var arvoreDeDesenvolvimento = raizRepositorio is not null && Directory.Exists(Path.Combine(raizRepositorio, "third_party"));
            if (arvoreDeDesenvolvimento && !permitirSemMotor) throw; // no desenvolvimento o motor tem de estar
            throw new SkipException(MotivoAusente);
        }
    }

    /// <summary>A pasta do mOffice.slnx acima do assembly de testes, ou null.</summary>
    internal static string? RaizRepositorio()
    {
        var d = new DirectoryInfo(AppContext.BaseDirectory);
        while (d is not null && !File.Exists(Path.Combine(d.FullName, "mOffice.slnx"))) d = d.Parent;
        return d?.FullName;
    }
}
