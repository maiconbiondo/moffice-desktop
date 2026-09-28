using Microsoft.Win32;

namespace mOffice.Comum;

/// <summary>Tema claro/escuro dos APLICATIVOS no Windows (Configuracoes > Personalizacao > Cores). O mPDF nao tem
/// isto; aqui e' o padrao (ModoTema.Sistema).</summary>
public static class TemaDoSistema
{
    private const string Chave = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

    public static bool AppsClaros()
    {
        try
        {
            using var k = Registry.CurrentUser.OpenSubKey(Chave);
            return k?.GetValue("AppsUseLightTheme") is not int v || v != 0;
        }
        catch (Exception e) when (e is System.Security.SecurityException or IOException or UnauthorizedAccessException) { return true; }
    }

    private static event Action? _alterado;
    public static event Action? Alterado
    {
        add { if (_alterado is null) SystemEvents.UserPreferenceChanged += AoMudarPreferencia; _alterado += value; }
        remove { _alterado -= value; if (_alterado is null) SystemEvents.UserPreferenceChanged -= AoMudarPreferencia; }
    }
    private static void AoMudarPreferencia(object sender, UserPreferenceChangedEventArgs e)
    { if (e.Category == UserPreferenceCategory.General) _alterado?.Invoke(); }
}
