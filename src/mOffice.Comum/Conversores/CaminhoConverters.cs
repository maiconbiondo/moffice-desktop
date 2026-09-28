using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace mOffice.Comum.Conversores;

/// Conversores de CHROME da casca (copiados de src/mPdf.App/Converters/ShellConverters.cs do mPDF, que
/// ja' os usa na mesma tela de boas-vindas). Nenhum toca o editor — sao utilitarios de binding.

/// <summary>Caminho completo -&gt; nome do arquivo (linha de cima de cada item de "Recentes").
/// <c>Path.GetFileName</c> nunca lanca para uma string qualquer.</summary>
public sealed class CaminhoParaNomeConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is string p && p.Length > 0 ? Path.GetFileName(p) : "";

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>Caminho completo -&gt; pasta (linha muda de baixo de cada item de "Recentes").
/// <c>Path.GetDirectoryName</c> devolve <c>null</c> para um caminho sem pasta (vira "").</summary>
public sealed class CaminhoParaPastaConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is string p && p.Length > 0 ? (Path.GetDirectoryName(p) ?? "") : "";

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>Visibilidade por PRESENCA: visivel quando o valor nao e' nulo (nem string vazia), colapsado
/// quando e'. E' o que esconde a faixa de aviso da barra e a faixa de erro das boas-vindas sem um bool
/// extra no ViewModel.</summary>
public sealed class NuloParaColapsadoConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is null || (value is string s && s.Length == 0) ? Visibility.Collapsed : Visibility.Visible;

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}

/// <summary>CHAVE de recurso (ex.: "Logo.mPlanilha") -&gt; o recurso em si. Existe porque o logo varia por
/// app (<see cref="PerfilApp.ChaveLogo"/>) e <c>StaticResource</c> nao aceita uma chave vinda de binding.
/// Recebe DOIS valores: [0] a chave e [1] um <see cref="FrameworkElement"/> pelo qual procurar — e' esse
/// elemento (nao <c>Application.Current</c>) que resolve nos TESTES, onde nao existe Application e o tema
/// esta mesclado no proprio UserControl/Window.</summary>
public sealed class ChaveParaRecursoConverter : IMultiValueConverter
{
    public object? Convert(object?[] values, Type targetType, object? parameter, CultureInfo culture)
    {
        if (values.Length == 0 || values[0] is not string chave || chave.Length == 0) return null;
        var elemento = values.Length > 1 ? values[1] as FrameworkElement : null;
        return elemento?.TryFindResource(chave) ?? Application.Current?.TryFindResource(chave);
    }

    public object?[] ConvertBack(object? value, Type[] targetTypes, object? parameter, CultureInfo culture) =>
        throw new NotSupportedException();
}
