using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

/// <summary>Plano 5, Task 1: <see cref="ConfiguracaoMotor.Localizar()"/> procura o motor em tres lugares, nesta
/// ordem — <c>&lt;base&gt;\motor</c> (producao), <c>MOFFICE_MOTOR_DIR</c> (teste do pacote) e a subida ate
/// <c>third_party\onlyoffice</c> (desenvolvimento). Os testes usam a costura com a base e a variavel por
/// parametro: mexer no ambiente do processo afetaria os outros testes que rodam em paralelo.</summary>
public class ConfiguracaoMotorTests
{
    private static string MotorFalso(string raiz)
    {
        var motor = Path.Combine(raiz, "motor");
        Directory.CreateDirectory(Path.Combine(motor, "converter"));
        File.WriteAllText(Path.Combine(motor, "converter", "x2t.exe"), "falso");
        return motor;
    }

    [Fact] // producao primeiro: <base>\motor vence a variavel; a variavel vence a subida ate third_party
    public void Localizar_PrefereMotorInstalado()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-localizar-").FullName;
        try
        {
            var instalado = MotorFalso(Path.Combine(raiz, "app"));
            var daVariavel = MotorFalso(Path.Combine(raiz, "pacote"));

            // (1) instalado: <base>\motor, mesmo com a variavel apontando para outro motor valido
            Assert.Equal(Path.GetFullPath(instalado),
                ConfiguracaoMotor.Localizar(Path.Combine(raiz, "app"), daVariavel).PastaMotor);

            // (2) variavel: a base e' a pasta dos testes, cuja subida acha third_party\onlyoffice nesta arvore
            //     de desenvolvimento — a variavel tem de vir ANTES dela
            Assert.Equal(Path.GetFullPath(daVariavel),
                ConfiguracaoMotor.Localizar(AppContext.BaseDirectory, daVariavel).PastaMotor);

            // controle: sem a variavel, a mesma base cai no motor de desenvolvimento (third_party\onlyoffice)
            var dev = MotorDeTeste.OuPular(() => ConfiguracaoMotor.Localizar(AppContext.BaseDirectory, null)).PastaMotor;
            Assert.EndsWith(Path.Combine("third_party", "onlyoffice"), dev, StringComparison.OrdinalIgnoreCase);
        }
        finally { Directory.Delete(raiz, true); }
    }

    [Fact] // variavel apontando para pasta sem x2t e' ignorada (nao devolve um motor quebrado)
    public void Localizar_VariavelSemX2t_Ignorada()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-localizar-").FullName;
        try
        {
            var dev = MotorDeTeste.OuPular(() => ConfiguracaoMotor.Localizar(AppContext.BaseDirectory, raiz)).PastaMotor;
            Assert.EndsWith(Path.Combine("third_party", "onlyoffice"), dev, StringComparison.OrdinalIgnoreCase);
        }
        finally { Directory.Delete(raiz, true); }
    }

    [Fact] // nada em lugar nenhum: a mensagem (pt-BR) cita os tres lugares procurados
    public void Localizar_SemNada_MensagemCitaOsTresLugares()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-localizar-").FullName;
        try
        {
            var inexistente = Path.Combine(raiz, "nao-existe");
            var erro = Assert.Throws<ErroConversao>(() => ConfiguracaoMotor.Localizar(raiz, inexistente));
            Assert.Contains(Path.Combine(raiz, "motor"), erro.Message);
            Assert.Contains(ConfiguracaoMotor.VariavelMotor, erro.Message);
            Assert.Contains(inexistente, erro.Message);
            Assert.Contains("third_party", erro.Message);
            Assert.Contains("não encontrado", erro.Message);

            var semVariavel = Assert.Throws<ErroConversao>(() => ConfiguracaoMotor.Localizar(raiz, null));
            Assert.Contains("não definida", semVariavel.Message);
        }
        finally { Directory.Delete(raiz, true); }
    }
}
