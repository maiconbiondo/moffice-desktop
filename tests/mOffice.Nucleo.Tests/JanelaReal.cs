namespace mOffice.Nucleo.Tests;

/// <summary>
/// Testes que abrem uma janela WPF + WebView2 de verdade rodam em serie: duas instancias ao mesmo tempo
/// disputam a MESMA pasta de dados do usuario do WebView2 (JanelaEditor.PastaDadosWebView) e o xunit
/// paraleliza classes diferentes por padrao.
/// </summary>
[CollectionDefinition(Nome)]
public sealed class JanelaReal
{
    public const string Nome = "janela real";
}
