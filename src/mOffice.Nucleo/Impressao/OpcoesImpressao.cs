using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace mOffice.Nucleo.Impressao;

/// <summary>
/// O que o painel <b>Arquivo -&gt; Imprimir</b> do editor (ou o botao de impressao rapida) pediu, lido do JSON que o
/// sdkjs manda em <c>AscDesktopEditor.Print(json)</c> (Task 4 do Plano 3).
///
/// Protocolo MEDIDO nos bundles 9.4.0.129 (o bundle vence o brief):
/// <list type="bullet">
/// <item><c>web-apps/apps/{documenteditor,presentationeditor,spreadsheeteditor}/main/app.js</c> montam
/// <c>asc_setNativeOptions({usesystemdialog, printer, colorMode, pages, paperSize:{w,h,preset}, paperOrientation,
/// copies, sides})</c>; a impressao rapida manda so' <c>{quickPrint:true}</c>. Valores: <c>colorMode</c> e'
/// BOOLEANO (<c>"color" === combo</c>); <c>pages</c> e' <c>"all"</c>, o numero da pagina atual (<c>"3"</c>) ou o
/// texto digitado (<c>"1-3,5"</c>); <c>paperOrientation</c> e' <c>"portrait"</c>, <c>"landscape"</c> ou
/// <c>"auto"</c> (texto e planilha; a apresentacao nao manda); <c>sides</c> e' <c>"one"</c>, <c>"both-long"</c>
/// ou <c>"both-short"</c>; <c>paperSize</c> em milimetros; <c>printer</c> pode ser <c>null</c>.</item>
/// <item>O sdkjs embrulha: texto (<c>sdkjs/word/sdk-all-min.js</c>, <c>NIe</c>) manda
/// <c>{nativeOptions, printOptions:{selection:1}?, documentLayout:{drawPlaceHolders:false, drawFormHighlight:false,
/// isPrint:true}}</c>; apresentacao (<c>W9f</c>) manda <c>{printOptions?, nativeOptions}</c>; planilha
/// (<c>OGg</c>) manda SO' <c>{nativeOptions}</c> — sem <c>spreadsheetLayout</c> (divergencia do brief: a area de
/// impressao da planilha nao vem neste JSON).</item>
/// </list>
///
/// Tudo o que NAO e' <c>nativeOptions</c> segue intacto para o <c>x2t</c> (<see cref="JsonParaX2t"/>, que vira o
/// <c>m_sJsonParams</c>): e' o que faz o PDF sair "como impresso" (sem marcadores de campo, so' a selecao...).
/// </summary>
/// <param name="DialogoDoSistema"><c>usesystemdialog</c>: o usuario pediu o dialogo de impressao do Windows.</param>
/// <param name="Rapida"><c>quickPrint</c>: botao de impressao rapida (impressora padrao, sem perguntar).</param>
/// <param name="Impressora">Nome da impressora, ou <c>null</c> = padrao.</param>
/// <param name="Paginas">Intervalo digitado/escolhido (<c>"1-3,5"</c>, <c>"3"</c>), ou <c>null</c> = todas.</param>
/// <param name="Lados"><c>"one"</c> | <c>"both-long"</c> | <c>"both-short"</c>.</param>
/// <param name="Papel">Nome do tamanho (<c>preset</c>, ex.: <c>"A4"</c>), quando o painel informou.</param>
public sealed record OpcoesImpressao(bool DialogoDoSistema, bool Rapida, string? Impressora, string? Paginas, int Copias,
    string Lados, string? Papel, double LarguraMm, double AlturaMm, bool Paisagem, bool Colorido, string? JsonParaX2t)
{
    /// <summary>Teto de copias: o painel nao passa disso na pratica, mas o JSON e' dado vindo da pagina.</summary>
    public const int MaximoCopias = 999;

    /// <summary>Padroes seguros: impressora padrao, todas as paginas, 1 copia, um lado, A4 retrato, colorido.</summary>
    public static OpcoesImpressao Padrao { get; } = new(false, false, null, null, 1, "one", null, 210, 297, false, true, null);

    /// <summary><c>paperOrientation = "auto"</c> (acrescimo ao contrato do brief, so' leitura): cada pagina fica na
    /// orientacao dela — quem manda a impressao NAO deve forcar retrato/paisagem. <see cref="Paisagem"/> fica
    /// <c>false</c> nesse caso.</summary>
    public bool OrientacaoAutomatica { get; init; }

    private static readonly JsonSerializerOptions SaidaJson = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    /// <summary>Le o JSON do <c>Print</c>. Campo ausente ou de tipo inesperado = padrao; JSON invalido (ou que
    /// nao seja objeto, ou com chave duplicada) = <see cref="Padrao"/>. Nunca lanca.</summary>
    public static OpcoesImpressao Ler(string? json)
    {
        // O JSON e' dado da PAGINA. Alem do JsonException do parse, o JsonObject so' monta o dicionario no
        // primeiro acesso e ai' lanca ArgumentException com chave duplicada (fix round 1 da revisao da Task 4):
        // a leitura inteira fica dentro do try.
        try { return LerObjeto(string.IsNullOrWhiteSpace(json) ? null : JsonNode.Parse(json) as JsonObject); }
        catch (Exception e) when (e is JsonException or ArgumentException or InvalidOperationException) { return Padrao; }
    }

    private static OpcoesImpressao LerObjeto(JsonObject? raiz)
    {
        if (raiz is null) return Padrao;

        var n = raiz["nativeOptions"] as JsonObject;
        var papel = n?["paperSize"] as JsonObject;
        var largura = Numero(papel?["w"]) is { } w && w > 0 ? w : Padrao.LarguraMm;
        var altura = Numero(papel?["h"]) is { } h && h > 0 ? h : Padrao.AlturaMm;
        var orientacao = Texto(n?["paperOrientation"]);
        var paginas = Texto(n?["pages"])?.Trim();
        if (string.IsNullOrEmpty(paginas) || paginas.Equals("all", StringComparison.OrdinalIgnoreCase)) paginas = null;
        var lados = Texto(n?["sides"]) is "both-long" or "both-short" ? Texto(n?["sides"])! : "one";
        var copias = Numero(n?["copies"]) is { } c ? (int)Math.Clamp(Math.Round(c), 1, MaximoCopias) : 1;
        var impressora = Texto(n?["printer"]);

        return new OpcoesImpressao(
            DialogoDoSistema: Booleano(n?["usesystemdialog"]) ?? false,
            Rapida: Booleano(n?["quickPrint"]) ?? false,
            Impressora: string.IsNullOrWhiteSpace(impressora) ? null : impressora,
            Paginas: paginas,
            Copias: copias,
            Lados: lados,
            Papel: Texto(papel?["preset"]) is { Length: > 0 } p ? p : null,
            LarguraMm: largura,
            AlturaMm: altura,
            Paisagem: string.Equals(orientacao, "landscape", StringComparison.OrdinalIgnoreCase),
            Colorido: Booleano(n?["colorMode"]) ?? true,
            JsonParaX2t: SemOpcoesNativas(raiz))
        {
            OrientacaoAutomatica = string.Equals(orientacao, "auto", StringComparison.OrdinalIgnoreCase),
        };
    }

    /// <summary>O JSON recebido SEM <c>nativeOptions</c> (as opcoes da impressora nao sao do x2t), ou <c>null</c>
    /// quando nao sobra nada.</summary>
    private static string? SemOpcoesNativas(JsonObject raiz)
    {
        var copia = (JsonObject)raiz.DeepClone();
        copia.Remove("nativeOptions");
        return copia.Count == 0 ? null : copia.ToJsonString(SaidaJson);
    }

    private static string? Texto(JsonNode? no) =>
        no is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;

    private static double? Numero(JsonNode? no)
    {
        if (no is not JsonValue v) return null;
        if (v.TryGetValue<double>(out var d)) return double.IsFinite(d) ? d : null;
        if (v.TryGetValue<string>(out var s) && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && double.IsFinite(d)) return d;
        return null;
    }

    /// <summary>Booleano tolerante: <c>true/false</c>, numero (diferente de zero = verdadeiro) ou o texto
    /// <c>"true"/"false"</c> (<c>colorMode</c> e' booleano no bundle 9.4, mas outros hosts do protocolo mandam 0/1).</summary>
    private static bool? Booleano(JsonNode? no)
    {
        if (no is not JsonValue v) return null;
        if (v.TryGetValue<bool>(out var b)) return b;
        if (v.TryGetValue<double>(out var d)) return d != 0;
        if (v.TryGetValue<string>(out var s) && bool.TryParse(s, out b)) return b;
        return null;
    }
}
