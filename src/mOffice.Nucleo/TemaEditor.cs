namespace mOffice.Nucleo;

/// <summary>Tema do EDITOR (web-apps): ids lidos de web-apps/apps/common/main/lib/controller/Themes.js na tag
/// 9.4.0.129 — theme-classic-light e theme-dark sao os dois que o app oficial usa como padrao.</summary>
public enum TemaEditor { Claro, Escuro }

public static class TemaEditorExtensoes
{
    public static string Id(this TemaEditor t) => t == TemaEditor.Escuro ? "theme-dark" : "theme-classic-light";
    public static string Tipo(this TemaEditor t) => t == TemaEditor.Escuro ? "dark" : "light";
}

/// <summary>Os 7 temas do motor (Plano 6, Task 4) e o tipo (claro/escuro) de cada um, para lembrar a escolha feita dentro do
/// editor (<c>uitheme:changed</c>). Um id fora desta lista nunca vira tema (vem do editor, nao do nosso codigo).</summary>
public static class TemasDoEditor
{
    private static readonly Dictionary<string, string> Tipos = new(StringComparer.Ordinal)
    {
        ["theme-classic-light"] = "light", ["theme-light"] = "light", ["theme-white"] = "light", ["theme-gray"] = "light",
        ["theme-dark"] = "dark", ["theme-contrast-dark"] = "dark", ["theme-night"] = "dark",
    };

    public static readonly IReadOnlySet<string> Conhecidos = Tipos.Keys.ToHashSet(StringComparer.Ordinal);

    /// <summary>"light" ou "dark". Id desconhecido = "light" (o padrao do editor).</summary>
    public static string Tipo(string id) => Tipos.TryGetValue(id, out var t) ? t : "light";
}
