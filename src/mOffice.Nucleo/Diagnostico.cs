namespace mOffice.Nucleo;

/// <summary>Modo de diagnostico (<c>MOFFICE_DIAGNOSTICO=1</c>, ligado por <c>tools/ao-vivo/lancar.ps1 -Diagnostico</c>): religa DevTools e as
/// teclas aceleradoras do navegador, que em uso normal ficam desligadas (SO-04/PL-16). Lido uma vez por processo.</summary>
public static class Diagnostico
{
    public static bool Ligado { get; } = Environment.GetEnvironmentVariable("MOFFICE_DIAGNOSTICO") == "1";
}
