namespace mOffice.Nucleo;

/// <summary>Tema do EDITOR (web-apps): ids lidos de web-apps/apps/common/main/lib/controller/Themes.js na tag
/// 9.4.0.129 — theme-classic-light e theme-dark sao os dois que o app oficial usa como padrao.</summary>
public enum TemaEditor { Claro, Escuro }

public static class TemaEditorExtensoes
{
    public static string Id(this TemaEditor t) => t == TemaEditor.Escuro ? "theme-dark" : "theme-classic-light";
    public static string Tipo(this TemaEditor t) => t == TemaEditor.Escuro ? "dark" : "light";
}
