using System.Windows;

namespace mOffice.Comum.Tests;

public static class JanelaWpf
{
    public static ResourceDictionary Tokens(bool escuro) => new()
    {
        Source = new Uri($"pack://application:,,,/mOffice.Comum;component/Temas/Tokens.{(escuro ? "Escuro" : "Claro")}.xaml")
    };
    public static ResourceDictionary Tema() => new() { Source = new Uri("pack://application:,,,/mOffice.Comum;component/Temas/mOfficeTema.xaml") };

    /// <summary>Recursos completos (tokens + tema) para um elemento de teste sem Application.Current.</summary>
    public static void Aplicar(FrameworkElement alvo, bool escuro)
    {
        alvo.Resources.MergedDictionaries.Insert(0, Tokens(escuro));
        alvo.Resources.MergedDictionaries.Add(Tema());
    }
}
