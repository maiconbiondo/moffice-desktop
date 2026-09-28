using System.Text.Json;

namespace mOffice.Nucleo;

public enum DecisaoFechar { Salvar, NaoSalvar, Cancelar }

/// <summary>Um execCommand que o web-apps mandou ao host (create:new, open:recent, recent:forget, editor:event,
/// go:folder, title:button, uitheme:changed, …). Comando e Json chegam brutos; a casca decide.</summary>
public sealed record ComandoDoEditor(string Comando, string Json)
{
    /// <summary>
    /// Le o <c>{cmd, json}</c> que o <c>ponte-desktop.js</c> serializa em <c>execCommand</c>. Devolve
    /// <c>null</c> quando o parametro nao e' JSON, ou e' JSON que nao e' objeto — nesses casos nao ha comando
    /// nenhum para entregar a casca.
    ///
    /// O campo <c>json</c> NEM SEMPRE e' string: o web-apps chama, por exemplo,
    /// <c>execCommand("go:folder", info.data.status)</c>, e o <c>JSON.stringify</c> do lado JS guarda ali um
    /// numero, um booleano, um objeto ou <c>null</c>. Por isso o valor e' coagido para o TEXTO BRUTO dele
    /// (<c>GetRawText</c>) em vez de passar por <c>GetString</c>, que lancaria — a casca recebe o que o editor
    /// mandou, sem o nucleo interpretar. <c>null</c>/ausente vira <c>""</c>, como no lado JS.
    /// </summary>
    public static ComandoDoEditor? Ler(string param)
    {
        try
        {
            using var doc = JsonDocument.Parse(param);
            if (doc.RootElement.ValueKind != JsonValueKind.Object) return null;
            var cmd = doc.RootElement.TryGetProperty("cmd", out var c) && c.ValueKind == JsonValueKind.String
                ? c.GetString() ?? ""
                : "";
            var json = doc.RootElement.TryGetProperty("json", out var j) ? Bruto(j) : "";
            return new ComandoDoEditor(cmd, json);
        }
        catch (JsonException) { return null; }
    }

    private static string Bruto(JsonElement e) => e.ValueKind switch
    {
        JsonValueKind.String => e.GetString() ?? "",
        JsonValueKind.Null or JsonValueKind.Undefined => "",
        _ => e.GetRawText(),
    };

    /// <summary>Para "editor:event": o campo action ("file:open", "file:close"); null nos demais ou se nao for JSON.</summary>
    public string? AcaoDoEvento()
    {
        if (Comando != "editor:event") return null;
        try
        {
            using var doc = JsonDocument.Parse(Json);
            return doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("action", out var a)
                ? a.GetString()
                : null;
        }
        catch (JsonException) { return null; }
    }
}
