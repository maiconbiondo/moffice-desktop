using System.Globalization;
using System.Printing;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace mOffice.Nucleo.Impressao;

/// <summary>Um tamanho de papel que a impressora aceita, em milimetros (1 casa).</summary>
public sealed record PapelImpressora(string Nome, double LarguraMm, double AlturaMm);

/// <summary>Uma impressora do Windows como o painel <b>Arquivo -&gt; Imprimir</b> do editor a mostra.
/// <paramref name="Nome"/> e' o <c>FullName</c> da fila (<c>\\servidor\fila</c> nas de rede) — o mesmo nome que
/// volta no <c>nativeOptions.printer</c> e que o <see cref="Impressora"/> manda ao WebView2.</summary>
public sealed record ImpressoraInfo(string Nome, bool FrenteEVerso, bool Cor, IReadOnlyList<PapelImpressora> Papeis);

/// <summary>
/// As impressoras instaladas no Windows, para o combo "Impressora" do painel de impressao do editor (Task 5 do
/// Plano 3). SO' LEITURA: enumera filas e capacidades; nunca pausa, configura nem manda trabalho.
///
/// Protocolo MEDIDO no web-apps 9.4.0.129 (<c>apps/*/main/app.js</c>, o bundle vence o brief): o painel NAO pede a
/// lista — o controlador <c>Common.Controllers.Desktop</c> define <c>window.on_native_message(cmd, param)</c> no
/// quadro do editor e, para <c>/printer:config/</c>, faz <c>JSON.parse(param)</c> e chama
/// <c>getController("Print").setPrintersInfo(p.current_printer, p.printers)</c> (a lista e' ACUMULADA por nome,
/// <c>_.union</c> + <c>_.uniq</c>). A vista (<c>code.js</c>, <c>updateCmbPrinter</c>) le de cada impressora
/// <c>name</c>, <c>paper_supported</c> (lista de <c>{name, width, height}</c> em MILIMETROS — o combo mostra
/// <c>width/10</c> cm), <c>duplex_supported</c> e <c>color_supported</c>; <c>current_printer</c> so' vira a
/// selecao se bater com um <c>name</c> da lista. Mensagens que chegam antes de o controlador existir podem ir em
/// <c>window.native_message_cmd[cmd]</c>, que ele reexecuta ao iniciar.
/// </summary>
public static class ImpressorasDoWindows
{
    private static readonly JsonSerializerOptions SaidaJson = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    /// <summary>
    /// Impressora padrao (ou <c>null</c>) e a lista. Pode levar SEGUNDOS com impressoras de rede: chamar sempre
    /// fora da thread de interface. Nunca lanca (nenhuma excecao, fix round 1): uma fila cujas capacidades nao
    /// podem ser lidas (offline, driver quebrado) entra com papeis vazios, sem frente e verso e colorida; uma fila
    /// que falha por inteiro e' pulada (<paramref name="log"/>); falha ao enumerar = lista vazia.
    /// </summary>
    public static (string? Padrao, IReadOnlyList<ImpressoraInfo> Lista) Listar(Action<string>? log = null)
    {
        string? padrao = null;
        var lista = new List<ImpressoraInfo>();
        try
        {
            using var servidor = new LocalPrintServer();
            try
            {
                using var fila = LocalPrintServer.GetDefaultPrintQueue();
                padrao = string.IsNullOrWhiteSpace(fila.FullName) ? null : fila.FullName;
            }
            catch (Exception e) { log?.Invoke("IMPRESSORAS sem padrao: " + e.Message); }

            PrintQueueCollection filas;
            try { filas = servidor.GetPrintQueues([EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections]); }
            catch (Exception e)
            {
                log?.Invoke("IMPRESSORAS enumeracao falhou: " + e.Message);
                return (padrao, lista);
            }
            using (filas)
                lista = Coletar(filas, f => f.FullName, (f, nome) => Capacidades(f, nome, log), f => f.Dispose(), log);
        }
        catch (Exception e)
        {
            // "Nunca lanca" vale para qualquer excecao (fix round 1): sem lista o painel so' fica sem o combo.
            log?.Invoke("IMPRESSORAS servidor de impressao indisponivel: " + e);
        }
        return (padrao, lista);
    }

    /// <summary>
    /// Uma entrada por fila (nome nao vazio, sem repetir). Fix round 1: QUALQUER excecao de uma fila (driver de
    /// terceiro quebrado lanca de tudo) pula so' essa fila, com log — a lista das outras nunca se perde; toda fila e'
    /// descartada, inclusive a que falhou. Generico para o teste poder usar filas falsas.
    /// </summary>
    internal static List<ImpressoraInfo> Coletar<T>(IEnumerable<T> filas, Func<T, string?> nome, Func<T, string, ImpressoraInfo> capacidades,
        Action<T> descartar, Action<string>? log)
    {
        var lista = new List<ImpressoraInfo>();
        foreach (var fila in filas)
        {
            try
            {
                var n = nome(fila);
                if (string.IsNullOrWhiteSpace(n) || lista.Any(i => string.Equals(i.Nome, n, StringComparison.OrdinalIgnoreCase))) continue;
                lista.Add(capacidades(fila, n));
            }
            catch (Exception e) { log?.Invoke("IMPRESSORAS fila pulada: " + e); }
            finally
            {
                try { descartar(fila); } catch (Exception e) { log?.Invoke("IMPRESSORAS descarte da fila: " + e.Message); }
            }
        }
        return lista;
    }

    private static ImpressoraInfo Capacidades(PrintQueue fila, string nome, Action<string>? log)
    {
        PrintCapabilities cap;
        try { cap = fila.GetPrintCapabilities(); }
        catch (Exception e)
        {
            // Offline / driver indisponivel: entra na lista mesmo assim (o usuario reconhece o nome), sem papeis.
            log?.Invoke($"IMPRESSORAS capacidades de \"{nome}\" indisponiveis: {e.Message}");
            return new ImpressoraInfo(nome, false, true, []);
        }
        var papeis = new List<PapelImpressora>();
        foreach (var m in cap.PageMediaSizeCapability)
        {
            if (m.PageMediaSizeName is not { } nomePapel || m.Width is not { } w || m.Height is not { } h || w <= 0 || h <= 0) continue;
            var papel = new PapelImpressora(NomeDoPapel(nomePapel), Mm(w), Mm(h));
            if (!papeis.Any(p => p.Nome == papel.Nome)) papeis.Add(papel);
        }
        return new ImpressoraInfo(nome,
            FrenteEVerso: cap.DuplexingCapability.Contains(Duplexing.TwoSidedLongEdge),
            Cor: cap.OutputColorCapability.Contains(OutputColor.Color),
            Papeis: papeis);
    }

    /// <summary>1/96 de polegada (unidade do System.Printing) -&gt; milimetros, 1 casa.</summary>
    private static double Mm(double unidades96) => Math.Round(unidades96 * 25.4 / 96, 1);

    /// <summary>Nome legivel no combo: <c>ISOA4</c> -&gt; <c>A4</c>, <c>NorthAmericaLetter</c> -&gt; <c>Letter</c>,
    /// <c>JISB5</c> -&gt; <c>JIS B5</c>; o resto fica como o nome do <see cref="PageMediaSizeName"/>.</summary>
    internal static string NomeDoPapel(PageMediaSizeName nome)
    {
        var s = nome.ToString();
        var iso = Regex.Match(s, @"^ISO([ABC]\d+)$");
        if (iso.Success) return iso.Groups[1].Value;
        var jis = Regex.Match(s, @"^JIS([AB]\d+)$");
        if (jis.Success) return "JIS " + jis.Groups[1].Value;
        return s.StartsWith("NorthAmerica", StringComparison.Ordinal) && s.Length > "NorthAmerica".Length ? s["NorthAmerica".Length..] : s;
    }

    /// <summary>
    /// O parametro de <c>on_native_message("printer:config", ...)</c> (chaves conferidas no bundle, ver o resumo
    /// da classe): <c>{"current_printer":"…","printers":[{"name":"…","paper_supported":[{"name":"A4","width":210,
    /// "height":297}],"duplex_supported":true,"color_supported":true}]}</c>. <c>current_printer</c> so' vai quando
    /// ha' padrao (o web-apps ignora <c>null</c> e mantem o que tinha).
    /// </summary>
    public static string JsonParaEditor(string? padrao, IReadOnlyList<ImpressoraInfo> lista)
    {
        var raiz = new JsonObject();
        if (!string.IsNullOrEmpty(padrao)) raiz["current_printer"] = padrao;
        var impressoras = new JsonArray();
        foreach (var i in lista)
        {
            var papeis = new JsonArray();
            foreach (var p in i.Papeis)
                papeis.Add(new JsonObject { ["name"] = p.Nome, ["width"] = p.LarguraMm, ["height"] = p.AlturaMm });
            impressoras.Add(new JsonObject
            {
                ["name"] = i.Nome,
                ["paper_supported"] = papeis,
                ["duplex_supported"] = i.FrenteEVerso,
                ["color_supported"] = i.Cor,
            });
        }
        raiz["printers"] = impressoras;
        return raiz.ToJsonString(SaidaJson);
    }

    /// <summary>Script que entrega o JSON ao painel no quadro do editor (mesma forma de alcancar o quadro que o
    /// <c>EndLoad</c>): <c>on_native_message</c> quando o controlador ja' existe ("entregue"); senao guarda em
    /// <c>native_message_cmd</c>, que o controlador reexecuta ao iniciar ("adiado").</summary>
    public static string ScriptParaEditor(string json)
    {
        var texto = JsonSerializer.Serialize(json);
        return string.Create(CultureInfo.InvariantCulture, $$"""
            (function () {
              try {
                var w = (typeof DesktopOfflineAppDocumentEndLoad === "function") ? window : window.frames[0];
                if (!w) return "sem quadro do editor";
                if (typeof w.on_native_message === "function") { w.on_native_message("printer:config", {{texto}}); return "entregue"; }
                w.native_message_cmd = w.native_message_cmd || {};
                w.native_message_cmd["printer:config"] = {{texto}};
                return "adiado";
              } catch (e) { return "erro: " + e; }
            })()
            """);
    }
}
