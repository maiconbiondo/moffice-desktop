using System.Text;
using System.Text.Json;
using Microsoft.Web.WebView2.Core;

namespace mOffice.Nucleo;

/// <summary>
/// Instrumentação da sonda (Task 4), ligada SÓ no modo <c>--sonda</c>.
///
/// Motivo de existir: a <see cref="GuardaEgresso"/> não enxerga o que o mapeamento de host virtual serve (ver
/// o comentário de classe dela) — 404 dentro de <c>moffice.local</c> não aparecem no log de egresso. O único
/// jeito de ver o que o build desktop pede (e o que o console dele grita) é pelo DevTools Protocol.
///
/// Escreve linhas em <c>poc/relatorios/sonda-&lt;doctype&gt;.log</c>: <c>HTTP &lt;status&gt;</c> por resposta,
/// <c>FALHA</c> por carregamento abortado (com a URL recuperada do <c>requestWillBeSent</c>, que o
/// <c>loadingFailed</c> não traz), <c>CONSOLE</c>, <c>EXCECAO</c> e <c>LOG</c>.
///
/// Achado desta sonda (2026-09-16), medido numa rodada AD HOC (o código do experimento Fetch.enable NÃO está commitado; reproduzir antes de depender disto): o <c>Fetch</c> do DevTools Protocol
/// (<c>Fetch.enable</c> com padrão exato E com <c>*AllFonts.js*</c>, <c>requestStage: Request</c>) TAMBÉM
/// não intercepta o que o mapeamento de host virtual serve — o <c>Fetch.requestPaused</c> nunca dispara e o
/// pedido segue direto para o disco (404 <c>ERR_FILE_NOT_FOUND</c>). Ou seja, o mapeamento resolve abaixo do
/// ponto em que tanto o <see cref="CoreWebView2.WebResourceRequested"/> quanto o <c>Fetch</c> agem: o
/// <c>Network</c> só OBSERVA (relata a falha), não substitui. Consequência para a Task 5A: recurso que o app
/// oficial gera em tempo de execução (AllFonts.js) tem de estar no disco dentro da pasta mapeada, ou ser
/// servido de outra origem/esquema (ex.: <c>ascdesktop://</c>), não dá para "costurar" por interceptação.
/// </summary>
public sealed class SondaInstrumentacao
{
    private readonly string _arquivo;
    private readonly object _trava = new();
    private readonly Dictionary<string, string> _urlPorRequisicao = [];
    // Os receptores precisam continuar vivos enquanto a sonda roda: sem guardar a referência, o GC pode
    // recolher o receptor e os eventos param de chegar.
    private readonly List<CoreWebView2DevToolsProtocolEventReceiver> _receptores = [];

    public string Arquivo => _arquivo;

    public SondaInstrumentacao(string arquivo)
    {
        _arquivo = arquivo;
        Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!);
        File.WriteAllText(arquivo, $"=== sonda {Path.GetFileNameWithoutExtension(arquivo)} {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===\n");
    }

    public async Task InstalarAsync(CoreWebView2 wv)
    {
        await wv.CallDevToolsProtocolMethodAsync("Network.enable", "{}");
        await wv.CallDevToolsProtocolMethodAsync("Runtime.enable", "{}");
        await wv.CallDevToolsProtocolMethodAsync("Log.enable", "{}");

        Assinar(wv, "Network.requestWillBeSent", p =>
        {
            var id = Texto(p, "requestId");
            var url = Texto(p, "request", "url");
            if (id is not null && url is not null) lock (_trava) _urlPorRequisicao[id] = url;
            return null; // ruído: a resposta já identifica o recurso
        });
        Assinar(wv, "Network.responseReceived", p =>
            $"HTTP {Texto(p, "response", "status")} {Texto(p, "type")} {Texto(p, "response", "url")}");
        Assinar(wv, "Network.loadingFailed", p =>
        {
            var id = Texto(p, "requestId");
            string? url = null;
            if (id is not null) lock (_trava) _urlPorRequisicao.TryGetValue(id, out url);
            return $"FALHA {Texto(p, "errorText")} {Texto(p, "type")} {url ?? "(url desconhecida: " + id + ")"}";
        });
        Assinar(wv, "Runtime.exceptionThrown", p =>
            $"EXCECAO {Texto(p, "exceptionDetails", "exception", "description") ?? Texto(p, "exceptionDetails", "text")}"
            + $" @ {Texto(p, "exceptionDetails", "url")}:{Texto(p, "exceptionDetails", "lineNumber")}");
        Assinar(wv, "Runtime.consoleAPICalled", p => $"CONSOLE {Texto(p, "type")} {Argumentos(p)}");
        Assinar(wv, "Log.entryAdded", p =>
            $"LOG {Texto(p, "entry", "level")} {Texto(p, "entry", "text")} @ {Texto(p, "entry", "url")}");
    }


    public void Registrar(string linha)
    {
        lock (_trava) File.AppendAllText(_arquivo, $"{DateTime.Now:HH:mm:ss.fff} {linha}\n");
    }

    private void Assinar(CoreWebView2 wv, string evento, Func<JsonElement, string?> formatar)
    {
        var receptor = wv.GetDevToolsProtocolEventReceiver(evento);
        receptor.DevToolsProtocolEventReceived += (_, e) =>
        {
            string linha;
            try
            {
                using var doc = JsonDocument.Parse(e.ParameterObjectAsJson);
                var formatada = formatar(doc.RootElement);
                if (formatada is null) return;
                linha = formatada;
            }
            catch (Exception ex)
            {
                linha = $"{evento} (falha ao ler: {ex.Message}) {Cortar(e.ParameterObjectAsJson, 500)}";
            }
            Registrar(linha);
        };
        _receptores.Add(receptor);
    }

    private static string Argumentos(JsonElement p)
    {
        if (!p.TryGetProperty("args", out var args) || args.ValueKind != JsonValueKind.Array) return "";
        var sb = new StringBuilder();
        foreach (var a in args.EnumerateArray())
        {
            if (sb.Length > 0) sb.Append(" | ");
            sb.Append(Texto(a, "value") ?? Texto(a, "description") ?? Texto(a, "type") ?? "?");
        }
        return Cortar(sb.ToString(), 1000);
    }

    private static string Cortar(string s, int max) => s.Length <= max ? s : s[..max] + "…";

    /// <summary>Lê um campo aninhado do JSON do evento sem exigir um DTO por evento do protocolo.</summary>
    private static string? Texto(JsonElement e, params string[] caminho)
    {
        foreach (var passo in caminho)
        {
            if (e.ValueKind != JsonValueKind.Object || !e.TryGetProperty(passo, out var filho)) return null;
            e = filho;
        }
        return e.ValueKind switch
        {
            JsonValueKind.String => e.GetString(),
            JsonValueKind.Null or JsonValueKind.Undefined => null,
            _ => e.ToString(),
        };
    }
}
