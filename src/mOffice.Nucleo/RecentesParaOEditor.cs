using System.Text.Json;
using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>
/// "Abrir recente" do menu Arquivo do editor (Plano 6, Task 8). O web-apps 9.4 (<c>Common.Controllers.Desktop</c>, app.js dos tres
/// editores) define <c>window.onupdaterecents(lista)</c> e passa cada item por <c>S(t)</c>: so' entra o item cujo <c>type</c> o
/// <c>matchFileFormat</c> reconhece (<c>c_oAscFileType</c>) e cujo <c>path</c> termina em nome.extensao; o <c>path</c> passa por
/// <c>$("&lt;div&gt;").html(path).text()</c> (por isso vai com &amp;, &lt; e &gt; codificados) e vira <c>url</c> do item — o mesmo que o
/// <c>open:recent</c> devolve a' casca. O nucleo nao conhece a lista: quem a fornece e' a casca (<see cref="JanelaEditor.FonteDeRecentes"/>).
/// </summary>
public static class RecentesParaOEditor
{
    public static string Json(IEnumerable<string> caminhos)
    {
        var itens = new List<object>();
        foreach (var c in caminhos)
        {
            if (string.IsNullOrWhiteSpace(c) || TabelaFormatos.DaExtensao(c) is not { Abre: true } info) continue;
            itens.Add(new { id = itens.Count, type = (int)info.Formato, path = CodificarHtml(c) });
        }
        return JsonSerializer.Serialize(itens);
    }

    private static string CodificarHtml(string s) => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
