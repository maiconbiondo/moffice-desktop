using Microsoft.Win32;

namespace mOffice.Nucleo.Assinatura;

/// <summary>O mPDF instalado nesta maquina: o exe (caminho absoluto local, existente) e a versao declarada pelo
/// instalador (<c>DisplayVersion</c>, normalizada para <c>Major.Minor.Build</c>).</summary>
public sealed record MpdfInstalado(string Exe, Version Versao);

/// <summary>
/// Acha o mPDF (app irmao da suite) pela chave de desinstalacao que o instalador Inno Setup dele grava
/// (<c>AppId</c> de <c>tools/installer/mpdf.iss</c> no repositorio do mPDF): <c>InstallLocation</c> + <c>DisplayVersion</c>.
/// Plano 4, Task 5. So' LE o registro; nunca lanca.
/// </summary>
public static class DetectorMpdf
{
    /// <summary>Primeira versao do mPDF com o verbo <c>/assinar</c> (mPDF 2.13.0). Mais antiga: o mOffice so' avisa.</summary>
    public static readonly Version Minima = new(2, 13, 0);

    /// <summary>Chave de desinstalacao do mPDF (relativa a HKLM/HKCU).</summary>
    public const string ChaveDesinstalacao = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\{C267C7BD-85F1-449C-8B1D-69DA6013BC8D}_is1";

    /// <summary>Nome do exe do mPDF dentro de <c>InstallLocation</c>.</summary>
    public const string NomeExe = "mPdf.App.exe";

    /// <summary>Onde a chave pode estar, na ordem de desempate: por maquina antes de por usuario, visao de 64 bits antes
    /// da de 32 (o Inno grava numa ou noutra conforme o modo do instalador; pode sobrar chave velha na outra).</summary>
    public static readonly IReadOnlyList<string> Origens = ["HKLM64", "HKLM32", "HKCU64", "HKCU32"];

    /// <summary>
    /// Detecta com o leitor de registro dado. <paramref name="lerChave"/> recebe uma das <see cref="Origens"/> e devolve
    /// (<c>InstallLocation</c>, <c>DisplayVersion</c>) da <see cref="ChaveDesinstalacao"/> naquela origem, ou null sem
    /// chave. TODAS as origens sao lidas e vence a MAIOR versao valida (revisao final do Plano 4: antes o HKLM valido
    /// ganhava sempre, e uma instalacao por usuario mais nova ficava escondida atras de uma por maquina velha); empate fica
    /// com a primeira na ordem de <see cref="Origens"/>. Chave presente com valor faltando/invalido, ou leitura negada,
    /// so' pula AQUELA origem. Valido = pasta absoluta LOCAL (nunca UNC, <c>\\?\</c>, relativa ou unidade de rede) com o
    /// <see cref="NomeExe"/> existente, e versao legivel (<c>v</c> inicial aceito; sufixo como <c>-rc1</c> e' cortado). A
    /// versao MINIMA nao e' filtrada aqui: quem chama decide (<see cref="AbreAssinatura"/>), para poder dizer ao usuario
    /// qual versao antiga esta' instalada.
    /// </summary>
    public static MpdfInstalado? Detectar(Func<string, (string? Local, string? Versao)?> lerChave)
    {
        MpdfInstalado? melhor = null;
        foreach (var origem in Origens)
        {
            (string? Local, string? Versao)? lido;
            try { lido = lerChave(origem); }
            catch (Exception e) when (e is System.Security.SecurityException or UnauthorizedAccessException or IOException) { continue; }
            if (lido is not { } v) continue;
            var versao = LerVersao(v.Versao);
            if (versao is null || (melhor is not null && versao <= melhor.Versao)) continue;
            var exe = ExeValido(v.Local);
            if (exe is not null) melhor = new MpdfInstalado(exe, versao);
        }
        return melhor;
    }

    /// <summary>Detecta pelo registro do Windows: HKLM e HKCU, cada um nas visoes de 64 e de 32 bits. So' leitura.</summary>
    public static MpdfInstalado? DetectarNoWindows() => Detectar(LerDoRegistro);

    /// <summary>Esta versao abre a assinatura sozinha (<c>/assinar</c>)?</summary>
    public static bool AbreAssinatura(MpdfInstalado mpdf) => mpdf.Versao >= Minima;

    /// <summary><c>DisplayVersion</c> -> <see cref="Version"/> com 3 partes (<c>2.13</c> vira <c>2.13.0</c>: sem isto
    /// <c>Version(2,13)</c> seria MENOR que <see cref="Minima"/>). Corta no primeiro caractere que nao e' digito nem ponto
    /// (<c>2.13.0-rc1</c>, <c>2.13.0 beta</c>); um <c>v</c> inicial (<c>v2.13.0</c>) e' aceito. Ilegivel = null.</summary>
    internal static Version? LerVersao(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto)) return null;
        var t = texto.Trim();
        if (t.Length > 1 && (t[0] == 'v' || t[0] == 'V') && char.IsAsciiDigit(t[1])) t = t[1..];   // "v2.13.0"
        var fim = 0;
        while (fim < t.Length && (char.IsAsciiDigit(t[fim]) || t[fim] == '.')) fim++;
        if (!Version.TryParse(t[..fim].TrimEnd('.'), out var v)) return null;
        return new Version(v.Major, v.Minor, Math.Max(0, v.Build));
    }

    /// <summary>O caminho do exe, se <paramref name="local"/> e' uma pasta absoluta local com o exe dentro; senao null.</summary>
    internal static string? ExeValido(string? local)
    {
        if (string.IsNullOrWhiteSpace(local)) return null;
        var l = local.Trim().Trim('"');
        // UNC (\\servidor\...), caminho de dispositivo (\\?\, \\.\) e relativos ficam de fora ANTES de tocar no disco.
        if (l.StartsWith(@"\\", StringComparison.Ordinal) || l.StartsWith("//", StringComparison.Ordinal)) return null;
        if (l.Length < 3 || !char.IsAsciiLetter(l[0]) || l[1] != ':' || (l[2] != '\\' && l[2] != '/')) return null;
        try
        {
            if (!Path.IsPathFullyQualified(l)) return null;
            if (new DriveInfo(l[..1]).DriveType is DriveType.Network or DriveType.NoRootDirectory) return null;
            var exe = Path.GetFullPath(Path.Combine(l, NomeExe));
            return File.Exists(exe) ? exe : null;
        }
        catch (Exception e) when (e is ArgumentException or IOException or NotSupportedException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            return null;
        }
    }

    private static (string? Local, string? Versao)? LerDoRegistro(string origem)
    {
        var hive = origem.StartsWith("HKLM", StringComparison.Ordinal) ? RegistryHive.LocalMachine : RegistryHive.CurrentUser;
        var visao = origem.EndsWith("64", StringComparison.Ordinal) ? RegistryView.Registry64 : RegistryView.Registry32;
        using var baseKey = RegistryKey.OpenBaseKey(hive, visao);
        using var k = baseKey.OpenSubKey(ChaveDesinstalacao, writable: false);
        if (k is null) return null;
        return (k.GetValue("InstallLocation") as string, k.GetValue("DisplayVersion") as string);
    }
}
