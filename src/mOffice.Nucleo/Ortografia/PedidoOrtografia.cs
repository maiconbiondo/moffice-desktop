using System.Text.Json;
using System.Text.Json.Nodes;

namespace mOffice.Nucleo.Ortografia;

/// <summary>Resposta ao <c>AscDesktopEditor.SpellCheck</c> do sdkjs 9.4, no formato do spellchecker.cpp do app
/// oficial: ecoa o pedido inteiro (a planilha le' cellsInfo/wordsIndex de volta) e acrescenta usrCorrect e
/// usrSuggest. <c>null</c> = nao responder (a string "clear" de reinicio e' tratada no JS, que so' precisa
/// devolve-la na ordem).
///
/// Lado do editor (MEDIDO em sdkjs/word/sdk-all-min.js 9.4.0.129): <c>window.asc_nativeOnSpellCheck(w)</c>
/// com <c>w === "clear"</c> libera o corretor; qualquer outra coisa vai para o tratador que le'
/// <c>w.ParagraphId</c>/<c>w.RecalcId</c>/<c>w.usrCorrect</c>/<c>w.usrSuggest[0]</c> — ou seja, um OBJETO.
/// Uma resposta sem <c>ParagraphId</c> (add/clear) e' ignorada por ele sem erro.</summary>
public static class PedidoOrtografia
{
    /// <summary>
    /// Teto de palavras com SUGESTOES por pedido "suggest" (M5 da revisao final do Plano 3, refeito no fix round 2). O
    /// pedido vem da PAGINA e a resposta roda na thread da interface: sugestao e' a operacao cara do Hunspell (ms por
    /// palavra), e o editor pede uma palavra por vez. Acima do teto as palavras seguintes recebem lista VAZIA — o pedido
    /// e' SEMPRE respondido.
    ///
    /// "spell" NAO tem teto de palavras (fix round 2): o editor de planilha manda TODAS as celulas de texto de uma coluna
    /// num pedido so', e um pedido sem resposta deixa a marca de "esperando" dele presa — o corretor da planilha parava
    /// de vez. Conferir uma palavra custa microssegundos; o que protege a janela e' o teto de TEXTO abaixo.
    /// </summary>
    public const int MaximoPalavrasComSugestao = 100;

    /// <summary>Teto do TEXTO do pedido, checado antes de parsear o JSON: acima dele nao ha' resposta (um pedido de
    /// 4 MiB nao sai de um editor em uso normal).</summary>
    public const int MaximoCaracteresPorPedido = 4 * 1024 * 1024;

    public static string? Responder(string pedido, Func<string, int, bool> correta,
        Func<string, int, IReadOnlyList<string>> sugestoes, DicionarioUsuario usuario, Action<string>? log = null)
    {
        if (pedido.Length > MaximoCaracteresPorPedido)
        {
            log?.Invoke($"ORTOGRAFIA pedido ignorado: {pedido.Length} caracteres (teto {MaximoCaracteresPorPedido})");
            return null;
        }
        JsonObject o;
        try { o = JsonNode.Parse(pedido) as JsonObject ?? throw new JsonException(); }
        catch (JsonException) { return AddSemEscape(pedido, usuario); }

        string? tipo;
        try { tipo = (string?)o["type"]; }
        catch (Exception e) when (e is InvalidOperationException or FormatException) { return null; }

        switch (tipo)
        {
            case "spell":
            {
                if (!Listas(o, out var palavras, out var idiomas)) return null;
                var certas = new JsonArray();
                for (var i = 0; i < palavras.Count; i++) certas.Add(usuario.Contem(palavras[i]) || correta(palavras[i], idiomas[i]));
                o["usrCorrect"] = certas; o["usrSuggest"] = new JsonArray();
                return o.ToJsonString();
            }
            case "suggest":
            {
                if (!Listas(o, out var palavras, out var idiomas)) return null;
                if (palavras.Count > MaximoPalavrasComSugestao)
                    log?.Invoke($"ORTOGRAFIA suggest com {palavras.Count} palavras: so' as {MaximoPalavrasComSugestao} primeiras recebem sugestoes");
                var todas = new JsonArray();
                for (var i = 0; i < palavras.Count; i++)
                    todas.Add(i < MaximoPalavrasComSugestao
                        ? new JsonArray(sugestoes(palavras[i], idiomas[i]).Select(s => (JsonNode?)JsonValue.Create(s)).ToArray())
                        : new JsonArray());
                o["usrCorrect"] = new JsonArray(); o["usrSuggest"] = todas;
                return o.ToJsonString();
            }
            case "add":
                if (o["usrWords"] is JsonArray novas)
                    foreach (var p in novas)
                        if (p is JsonValue v && v.TryGetValue<string>(out var s)) usuario.Adicionar(s);
                return """{"type":"add"}""";
            case "clear":
                usuario.Limpar();
                return """{"type":"clear"}""";
            default:
                return null;
        }
    }

    /// <summary>Palavras e idiomas em paralelo. Tamanhos diferentes (ou tipos errados) = pedido descartado,
    /// como no oficial: nao ha' como saber qual idioma vale para qual palavra.</summary>
    private static bool Listas(JsonObject o, out List<string> palavras, out List<int> idiomas)
    {
        palavras = []; idiomas = [];
        if (o["usrWords"] is not JsonArray ws || o["usrLang"] is not JsonArray ls || ws.Count != ls.Count) return false;
        foreach (var n in ws) palavras.Add(n is JsonValue v && v.TryGetValue<string>(out var s) ? s : "");
        foreach (var n in ls) idiomas.Add(n is JsonValue v && v.TryGetValue<int>(out var l) ? l : 0);
        return true;
    }

    /// <summary>O sdkjs monta o "add" concatenando a palavra: <c>'{"type":"add","usrWords":["' + w + '"]}'</c>
    /// (MEDIDO: sdk-all-min.js, <c>A9i</c>). Uma palavra com aspas ou barra invertida quebra o JSON —
    /// recupera pelo prefixo e sufixo fixos.</summary>
    private static string? AddSemEscape(string pedido, DicionarioUsuario usuario)
    {
        const string ini = "{\"type\":\"add\",\"usrWords\":[\"", fim = "\"]}";
        if (pedido.Length < ini.Length + fim.Length
            || !pedido.StartsWith(ini, StringComparison.Ordinal) || !pedido.EndsWith(fim, StringComparison.Ordinal)) return null;
        usuario.Adicionar(pedido[ini.Length..^fim.Length]);
        return """{"type":"add"}""";
    }
}
