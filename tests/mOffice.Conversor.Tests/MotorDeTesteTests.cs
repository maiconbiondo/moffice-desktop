using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

/// <summary><see cref="MotorDeTeste"/>: motor ausente PULA num clone publico (sem third_party/) e FALHA na arvore de
/// desenvolvimento (com third_party/), salvo <see cref="MotorDeTeste.VariavelPermitir"/>=1. Qualquer outro erro passa.</summary>
public class MotorDeTesteTests
{
    private static ConfiguracaoMotor SemMotor() =>
        throw new ErroConversao("Motor ONLYOFFICE não encontrado (converter\\x2t.exe). Procurado em: teste", 0, "");

    private static string Arvore(bool comThirdParty)
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-motordeteste-").FullName;
        if (comThirdParty) Directory.CreateDirectory(Path.Combine(raiz, "third_party"));
        return raiz;
    }

    [Fact]
    public void ClonePublico_SemMotor_Pula()
    {
        var raiz = Arvore(comThirdParty: false);
        try { Assert.Throws<SkipException>(() => MotorDeTeste.OuPular(SemMotor, raiz, permitirSemMotor: false)); }
        finally { Directory.Delete(raiz, true); }
    }

    [Fact]
    public void Desenvolvimento_SemMotor_Falha()
    {
        var raiz = Arvore(comThirdParty: true);
        try
        {
            var e = Assert.Throws<ErroConversao>(() => MotorDeTeste.OuPular(SemMotor, raiz, permitirSemMotor: false));
            Assert.StartsWith("Motor ONLYOFFICE n", e.Message);
        }
        finally { Directory.Delete(raiz, true); }
    }

    [Fact]
    public void Desenvolvimento_SemMotor_ComLiberacao_Pula()
    {
        var raiz = Arvore(comThirdParty: true);
        try { Assert.Throws<SkipException>(() => MotorDeTeste.OuPular(SemMotor, raiz, permitirSemMotor: true)); }
        finally { Directory.Delete(raiz, true); }
    }

    [Fact] // so' o "motor nao encontrado" vira skip: outro ErroConversao continua falhando, mesmo num clone publico
    public void OutroErro_NuncaViraSkip()
    {
        var raiz = Arvore(comThirdParty: false);
        try
        {
            Assert.Throws<ErroConversao>(() => MotorDeTeste.OuPular<ConfiguracaoMotor>(
                () => throw new ErroConversao("x2t saiu com erro", 1, ""), raiz, permitirSemMotor: true));
        }
        finally { Directory.Delete(raiz, true); }
    }

    [Fact] // esta arvore (desenvolvimento) e' reconhecida como tal
    public void RaizRepositorio_Encontrada() => Assert.NotNull(MotorDeTeste.RaizRepositorio());
}
