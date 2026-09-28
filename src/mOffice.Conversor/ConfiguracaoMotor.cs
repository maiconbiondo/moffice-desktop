namespace mOffice.Conversor;

/// <summary>Caminhos do motor ONLYOFFICE: <c>&lt;base&gt;\motor</c> instalado, ou third_party/onlyoffice no desenvolvimento (ver <see cref="Localizar()"/>).</summary>
public sealed record ConfiguracaoMotor(string PastaMotor)
{
    public string X2t => Path.Combine(PastaMotor, "converter", "x2t.exe");

    /// <summary>Item 7 da revisão 1 (Marco 1/Plano 1/Task 2): a sondagem de escrita em
    /// <c>%ProgramData%\mOffice\fontes</c> (<see cref="SondarPastaFontes"/>) faz três operações de disco
    /// (<c>CreateDirectory</c> + escrever + apagar um arquivo de teste) — e rodava de novo a CADA construção
    /// de <see cref="ConfiguracaoMotor"/>, porque este é um <c>record</c> com inicializador de propriedade
    /// (que chamadores como <c>mOffice.Sonda.App.PastaRelatorios</c> disparam a cada acesso). <see cref="Lazy{T}"/>
    /// com o valor padrão (thread-safe, executa uma única vez) garante que a sondagem roda uma vez por
    /// PROCESSO, não uma vez por instância.</summary>
    private static readonly Lazy<string> _pastaFontesCache = new(SondarPastaFontes);

    /// <summary>
    /// Pasta com AllFonts.js + font_selection.bin + fonts_thumbnail*.png (m_sFontDir / m_sAllFontsPath) —
    /// o catálogo DO mOffice, gerado por ele mesmo (<c>mOffice.Nucleo.CatalogoFontes</c>, Task 2 do Plano 1).
    ///
    /// Até a Task 2 esta propriedade apontava para o cache do ONLYOFFICE **instalado**
    /// (<c>%LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\data\fonts</c>), porque o pacote 9.4.0.129 do Desktop
    /// Editors não traz <c>allfontsgen.exe</c> nem o catálogo pronto — ou seja, o produto não rodava num PC
    /// sem o ONLYOFFICE oficial. Agora o mOffice extrai o <c>allfontsgen.exe</c> do pacote do Document
    /// Server (mesma versão, mesmo core AGPL) e gera o catálogo por conta própria, nesta pasta.
    /// </summary>
    public string PastaFontes { get; init; } = _pastaFontesCache.Value;

    /// <summary>Temas de apresentacao (m_sThemeDir). HIPOTESE: editors/sdkjs/slide/themes — confirmar com Get-ChildItem; se nao existir, o elemento nao e emitido.</summary>
    public string PastaTemas { get; init; } = Path.Combine(PastaMotor, "editors", "sdkjs", "slide", "themes");
    public string PastaTemp { get; init; } = PastaTempPadrao;

    /// <summary>Pasta-pai das pastas temporarias por chamada do x2t (<see cref="PastaTemp"/> padrao). Estatica para a
    /// varredura da subida do app (fix round 2 da revisao final do Plano 3), que nao precisa localizar o motor.</summary>
    public static string PastaTempPadrao => Path.Combine(Path.GetTempPath(), "mOffice", "x2t");

    /// <summary>Por MÁQUINA (<c>%ProgramData%\mOffice\fontes</c>) quando dá para escrever lá — um catálogo só
    /// para todos os usuários do PC, gerado uma vez; senão, por USUÁRIO
    /// (<c>%LOCALAPPDATA%\mOffice\fontes</c>), que é sempre gravável. A escolha é feita tentando escrever de
    /// verdade: <c>Directory.CreateDirectory</c> em <c>%ProgramData%</c> costuma passar mesmo sem elevação
    /// (a ACL padrão deixa criar), e é o arquivo que revela a falta de permissão.</summary>
    private static string SondarPastaFontes()
    {
        var porMaquina = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "mOffice", "fontes");
        try
        {
            Directory.CreateDirectory(porMaquina);
            var teste = Path.Combine(porMaquina, ".teste-escrita");
            File.WriteAllText(teste, "");
            File.Delete(teste);
            return porMaquina;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "mOffice", "fontes");
        }
    }

    /// <summary>Variável de ambiente que aponta direto para uma pasta de motor (a que contém
    /// <c>converter\x2t.exe</c>). Usada pelo teste do pacote montado (<c>PacoteTests</c>); a produção não a define.</summary>
    public const string VariavelMotor = "MOFFICE_MOTOR_DIR";

    /// <summary>Plano 5, Task 1: três lugares, nesta ordem —
    /// (1) PRODUÇÃO: <c>&lt;pasta do exe&gt;\motor</c> (layout do pacote/instalador, <c>tools/installer/montar-pacote.ps1</c>);
    /// (2) TESTE DO PACOTE: <see cref="VariavelMotor"/>;
    /// (3) DESENVOLVIMENTO: sobe pastas a partir do exe até achar <c>third_party\onlyoffice</c>.</summary>
    public static ConfiguracaoMotor Localizar() =>
        Localizar(AppContext.BaseDirectory, Environment.GetEnvironmentVariable(VariavelMotor));

    /// <summary>Costura de teste de <see cref="Localizar()"/>: a pasta base e o valor da variável entram por
    /// parâmetro, para o teste não mexer no ambiente do processo (outros testes em paralelo chamam
    /// <see cref="Localizar()"/>).</summary>
    internal static ConfiguracaoMotor Localizar(string pastaBase, string? motorDaVariavel)
    {
        static bool TemX2t(string pasta) => File.Exists(Path.Combine(pasta, "converter", "x2t.exe"));

        var instalado = Path.Combine(pastaBase, "motor");
        if (TemX2t(instalado)) return new ConfiguracaoMotor(Path.GetFullPath(instalado));

        if (!string.IsNullOrWhiteSpace(motorDaVariavel) && TemX2t(motorDaVariavel))
            return new ConfiguracaoMotor(Path.GetFullPath(motorDaVariavel));

        var dir = new DirectoryInfo(pastaBase);
        while (dir is not null)
        {
            var candidato = Path.Combine(dir.FullName, "third_party", "onlyoffice");
            if (TemX2t(candidato)) return new ConfiguracaoMotor(candidato);
            dir = dir.Parent;
        }
        var variavel = string.IsNullOrWhiteSpace(motorDaVariavel) ? "não definida" : $"\"{motorDaVariavel}\"";
        throw new ErroConversao(
            "Motor ONLYOFFICE não encontrado (converter\\x2t.exe). Procurado em: " +
            $"(1) instalação: {instalado} — reinstale o mOffice; " +
            $"(2) variável {VariavelMotor} ({variavel}); " +
            "(3) desenvolvimento: third_party\\onlyoffice subindo a partir da pasta do programa — rode tools/onlyoffice/obter-motor.ps1.",
            0, "");
    }
}
