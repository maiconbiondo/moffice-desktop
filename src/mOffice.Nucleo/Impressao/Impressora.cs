using System.Printing;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using Microsoft.Web.WebView2.Core;

namespace mOffice.Nucleo.Impressao;

/// <summary>
/// O que vai para o <c>PrintAsync</c> do WebView2, ja' decidido (Task 5 do Plano 3). Separado do envio para ser
/// testavel sem fila nem janela. <c>Default</c> nos enums = "o padrao da impressora" (impressao rapida).
/// </summary>
/// <param name="Paginas">No formato do WebView2 (<c>"1-3,5"</c>), ou <c>null</c> = todas.</param>
/// <param name="Orientacao"><c>null</c> = nao forcar (orientacao "auto" do painel, impressao rapida).</param>
/// <param name="LarguraPol">Papel em polegadas, SEMPRE no sentido retrato (menor lado); <c>null</c> = nao informar.</param>
internal sealed record ConfiguracaoEnvio(string Impressora, string? Paginas, int Copias, CoreWebView2PrintDuplex Duplex,
    CoreWebView2PrintColorMode Cor, CoreWebView2PrintOrientation? Orientacao, double? LarguraPol, double? AlturaPol)
{
    /// <summary>A fila pergunta onde salvar (porta <c>PORTPROMPT:</c> — Microsoft Print to PDF, XPS): o envio espera
    /// o usuario, entao o limite de tempo e' o de <see cref="Impressora.LimiteComDialogo"/>.</summary>
    public bool PerguntaOndeSalvar { get; init; }

    /// <summary>O Windows diz que a fila esta' offline/desligada (fix round 1: NAO e' erro — impressora de rede
    /// dormindo aceita o trabalho; ele fica na fila).</summary>
    public bool Offline { get; init; }
}

/// <summary>Resultado de um envio que chegou a' fila: a impressora e se o Windows a dava como offline.</summary>
public sealed record ImpressaoEnviada(string Impressora, bool Offline);

/// <summary>
/// Manda o PDF de impressao (gerado pelo x2t, Task 4) para uma impressora do Windows — decisao A da sonda S5:
/// uma <c>CoreWebView2</c> OCULTA, no MESMO ambiente da janela (runtime fixo, mesmas flags) e com a MESMA guarda de
/// egresso, abre o PDF local no visualizador embutido e chama <c>PrintAsync</c> com as opcoes do painel do editor.
///
/// <para><b>Diálogo do sistema</b> (<c>usesystemdialog</c>): o dialogo "Imprimir" do Windows via
/// <see cref="PrintDialog"/> do WPF (o mesmo <c>PrintDlgEx</c>; no Windows 11, o "Imprimir" moderno), modal para a
/// janela do documento, e depois o mesmo <c>PrintAsync</c> com o que o usuario escolheu. O
/// <c>ShowPrintUI(System)</c> do WebView2 NAO e' usado — MEDIDO na Task 5 (<c>Sonda_ShowPrintUISistema_*</c>): ele
/// abre o mesmo dialogo, mas modal so' para a janela OCULTA (a do documento continua clicavel) e com o titulo
/// "Microsoft Edge WebView2 - Imprimir"; e nao devolve NADA (nem se imprimiu, nem em qual impressora, nem quando
/// acabou) — sem isso nao ha' como apagar o PDF depois da impressao nem dizer "Enviado para …".</para>
///
/// <para><b>Mensagens</b> (fix round 1): toda <see cref="InvalidOperationException"/> que sai daqui tem mensagem
/// pt-BR para a faixa do usuario; o detalhe tecnico (mensagem da excecao original, HRESULT) vai so' para o log.</para>
///
/// <para>Roda na thread de interface (o ambiente do WebView2 e' por thread de interface, ver
/// <see cref="AmbienteWebView2.ObterAsync"/>). Nao apaga o PDF: quem pediu a impressao e' dono dele.</para>
/// </summary>
public sealed class Impressora
{
    private readonly string _app;
    private readonly Window? _dono;
    private string? _log;

    /// <param name="app">App dono do processo (pasta de dados do WebView2 se o ambiente ainda nao existir).</param>
    /// <param name="dono">Janela sobre a qual o dialogo "Imprimir" do Windows abre.</param>
    public Impressora(string app, Window? dono = null)
    {
        _app = app;
        _dono = dono;
    }

    /// <summary>Log desta impressao (a guarda de egresso da WebView2 oculta escreve nele). Criado no 1o uso.</summary>
    public string ArquivoLog => _log ??= Logs.NovoArquivo(_app + "-impressao");

    /// <summary>Limite do <c>PrintAsync</c> numa fila comum (fix round 1): generoso — um documento grande num driver
    /// lento leva minutos para ser entregue ao spooler, mas nao dez.</summary>
    public static readonly TimeSpan LimiteEnvio = TimeSpan.FromMinutes(10);

    /// <summary>Limite quando a fila pergunta onde salvar (PORTPROMPT): o envio espera o usuario responder.</summary>
    public static readonly TimeSpan LimiteComDialogo = TimeSpan.FromMinutes(60);

    /// <summary>Costura de teste: roda com a WebView2 oculta ja' no PDF, logo antes do envio; o que devolver e'
    /// descartado depois dele (ex.: o vigia de um dialogo restrito ao navegador DESTE envio). `internal`.</summary>
    internal Func<CoreWebView2, IDisposable?>? AntesDoEnvioParaTeste { get; set; }

    /// <summary>Costura de teste (fix round 1): troca o <c>PrintAsync</c> (fila de verdade) por outro envio — os
    /// testes da suite normal usam <c>PrintToPdfAsync</c>, que nao toca em fila nenhuma. `internal`.</summary>
    internal Func<CoreWebView2Environment, CoreWebView2, CoreWebView2PrintSettings, Task<CoreWebView2PrintStatus>>? EnvioParaTeste { get; set; }

    /// <summary>Costura de teste: fila "verificada" sem consultar o Windows (nome, offline, pergunta onde salvar). `internal`.</summary>
    internal Func<string?, EstadoDaFila>? FilaParaTeste { get; set; }

    /// <summary>Costura de teste: limite do envio no lugar de <see cref="LimiteEnvio"/>. `internal`.</summary>
    internal TimeSpan? LimiteEnvioParaTeste { get; set; }

    /// <summary>
    /// Imprime <paramref name="pdf"/> conforme o painel. Impressora <c>null</c> ou impressao rapida = a padrao do
    /// Windows. Devolve a impressora que recebeu o trabalho (e se o Windows a dava como offline), ou <c>null</c>
    /// quando o usuario cancelou o dialogo "Imprimir" do Windows.
    /// </summary>
    /// <exception cref="InvalidOperationException">Mensagem pt-BR: impressora inexistente, nenhuma impressora padrao,
    /// intervalo de paginas invalido, visualizador que nao carregou, tempo esgotado, ou a impressora recusou/falhou.</exception>
    public async Task<ImpressaoEnviada?> ImprimirAsync(string pdf, OpcoesImpressao o, CancellationToken ct = default)
    {
        ct.ThrowIfCancellationRequested();
        if (!File.Exists(pdf)) throw new InvalidOperationException("O PDF de impressão não existe mais.");
        ConfiguracaoEnvio cfg;
        if (o.DialogoDoSistema)
        {
            var escolhida = await EscolherNoDialogoDoSistemaAsync(o);
            if (escolhida is null) { Log("IMPRESSAO cancelada no dialogo do sistema"); return null; }
            cfg = escolhida;
        }
        else
        {
            var nome = o.Rapida ? null : o.Impressora;
            var fila = FilaParaTeste is { } f ? f(nome) : await Task.Run(() => VerificarFila(nome, Log), ct);
            Log($"IMPRESSAO fila \"{fila.Nome}\" (offline={fila.Offline}, pergunta onde salvar={fila.PerguntaOndeSalvar}, papel padrao {(fila.PapelPadraoPol is { } p ? $"{p.L:0.##}x{p.A:0.##} pol" : "desconhecido")})");
            cfg = Montar(o, fila.Nome, fila.PapelPadraoPol) with { Offline = fila.Offline, PerguntaOndeSalvar = fila.PerguntaOndeSalvar };
        }
        ct.ThrowIfCancellationRequested();
        await EnviarAsync(pdf, cfg, ct);
        return new ImpressaoEnviada(cfg.Impressora, cfg.Offline);
    }

    // ---------------------------------------------------------------------------------------------------
    // Decisao (pura)
    // ---------------------------------------------------------------------------------------------------

    private static readonly Regex FormatoPaginas = new(@"^\s*\d+\s*(-\s*\d+\s*)?(,\s*\d+\s*(-\s*\d+\s*)?)*$", RegexOptions.CultureInvariant);

    /// <summary>
    /// As opcoes do painel na forma do WebView2. Impressao rapida = so' a impressora padrao, com os padroes DELA
    /// (papel <paramref name="papelPadraoPol"/>, frente/verso e cor <c>Default</c>, 1 copia, todas as paginas).
    /// Papel sempre no sentido retrato (menor lado primeiro) e a orientacao a parte; "auto" nao forca orientacao.
    /// </summary>
    /// <exception cref="InvalidOperationException">Intervalo de paginas fora do formato <c>"1-3,5"</c>.</exception>
    internal static ConfiguracaoEnvio Montar(OpcoesImpressao o, string impressora, (double L, double A)? papelPadraoPol)
    {
        if (o.Rapida)
            return new ConfiguracaoEnvio(impressora, null, 1, CoreWebView2PrintDuplex.Default, CoreWebView2PrintColorMode.Default,
                null, papelPadraoPol?.L, papelPadraoPol?.A);
        var paginas = NormalizarPaginas(o.Paginas);
        var (l, a) = Retrato(o.LarguraMm / 25.4, o.AlturaMm / 25.4);
        return new ConfiguracaoEnvio(impressora, paginas, o.Copias,
            o.Lados switch { "both-long" => CoreWebView2PrintDuplex.TwoSidedLongEdge, "both-short" => CoreWebView2PrintDuplex.TwoSidedShortEdge, _ => CoreWebView2PrintDuplex.OneSided },
            o.Colorido ? CoreWebView2PrintColorMode.Color : CoreWebView2PrintColorMode.Grayscale,
            o.OrientacaoAutomatica ? null : o.Paisagem ? CoreWebView2PrintOrientation.Landscape : CoreWebView2PrintOrientation.Portrait,
            l, a);
    }

    /// <summary>O que o usuario escolheu no dialogo "Imprimir" do Windows (fila, bilhete e intervalo), na forma do
    /// WebView2. O que o bilhete nao diz fica com o padrao da impressora.</summary>
    internal static ConfiguracaoEnvio DoDialogo(string impressora, PrintTicket t, PageRange? intervalo)
    {
        (double, double)? papel = t.PageMediaSize is { Width: { } w, Height: { } h } && w > 0 && h > 0 ? Retrato(w / 96, h / 96) : null;
        return new ConfiguracaoEnvio(impressora,
            intervalo is { } r && r.PageFrom > 0 ? (r.PageTo > r.PageFrom ? $"{r.PageFrom}-{r.PageTo}" : $"{r.PageFrom}") : null,
            Math.Clamp(t.CopyCount ?? 1, 1, OpcoesImpressao.MaximoCopias),
            t.Duplexing switch
            {
                Duplexing.TwoSidedLongEdge => CoreWebView2PrintDuplex.TwoSidedLongEdge,
                Duplexing.TwoSidedShortEdge => CoreWebView2PrintDuplex.TwoSidedShortEdge,
                Duplexing.OneSided => CoreWebView2PrintDuplex.OneSided,
                _ => CoreWebView2PrintDuplex.Default,
            },
            t.OutputColor switch
            {
                OutputColor.Color => CoreWebView2PrintColorMode.Color,
                OutputColor.Grayscale or OutputColor.Monochrome => CoreWebView2PrintColorMode.Grayscale,
                _ => CoreWebView2PrintColorMode.Default,
            },
            t.PageOrientation switch
            {
                PageOrientation.Landscape or PageOrientation.ReverseLandscape => CoreWebView2PrintOrientation.Landscape,
                PageOrientation.Portrait or PageOrientation.ReversePortrait => CoreWebView2PrintOrientation.Portrait,
                _ => null,
            },
            papel?.Item1, papel?.Item2);
    }

    /// <summary><c>"1-3, 5"</c> -&gt; <c>"1-3,5"</c>; vazio -&gt; <c>null</c> (todas). O painel deixa digitar qualquer
    /// coisa: formato invalido vira mensagem ao usuario, nao um <c>ArgumentException</c> do WebView2. Pagina 0 e
    /// intervalo invertido tambem sao recusados AQUI — MEDIDO no WebView2 (Task 5): <c>"3-2"</c> falha e
    /// <c>"0"</c> TRAVA o <c>PrintToPdfAsync</c> (nunca volta). Intervalo alem do fim do documento: <c>"2-9"</c> num
    /// documento de 3 paginas imprime 2-3; <c>"5-9"</c> falha (vira a mensagem de falha do envio).</summary>
    internal static string? NormalizarPaginas(string? paginas)
    {
        if (string.IsNullOrWhiteSpace(paginas)) return null;
        var invalido = new InvalidOperationException($"Intervalo de páginas inválido: \"{paginas.Trim()}\". Use, por exemplo, 1-3,5.");
        if (!FormatoPaginas.IsMatch(paginas)) throw invalido;
        var partes = new List<string>();
        foreach (var parte in paginas.Split(','))
        {
            var numeros = parte.Split('-').Select(n => int.TryParse(n.Trim(), out var v) ? v : 0).ToArray();
            if (numeros.Any(n => n < 1) || (numeros.Length == 2 && numeros[1] < numeros[0])) throw invalido;
            partes.Add(string.Join("-", numeros));
        }
        return string.Join(",", partes);
    }

    private static (double, double) Retrato(double x, double y) => (Math.Min(x, y), Math.Max(x, y));

    // ---------------------------------------------------------------------------------------------------
    // Fila do Windows (so' leitura)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>O que o envio precisa saber da fila: o nome como o Windows o conhece (<c>FullName</c>), se ela esta'
    /// offline (so' informativo), se a porta pergunta onde salvar (<c>PORTPROMPT:</c>) e o papel padrao dela.</summary>
    internal sealed record EstadoDaFila(string Nome, bool Offline, bool PerguntaOndeSalvar, (double L, double A)? PapelPadraoPol);

    /// <summary>
    /// A fila pelo nome (<c>null</c> = padrao do Windows). So' leitura. Fila que NAO existe (ou nenhuma padrao) =
    /// erro pt-BR com o nome; fila offline NAO e' erro (fix round 1, decisao do controlador: impressora de rede
    /// dormindo aceita o trabalho — e' so' informada ao usuario). Falha do sistema de impressao = erro pt-BR
    /// generico; o detalhe tecnico vai para <paramref name="log"/>.
    /// </summary>
    internal static EstadoDaFila VerificarFila(string? nome, Action<string>? log = null)
    {
        try
        {
            using var servidor = new LocalPrintServer();
            if (nome is null)
            {
                PrintQueue? padrao = null;
                try { padrao = LocalPrintServer.GetDefaultPrintQueue(); }
                catch (Exception e) when (e is PrintQueueException or PrintSystemException) { log?.Invoke("IMPRESSAO sem fila padrao: " + e.Message); }
                if (padrao is null || string.IsNullOrWhiteSpace(padrao.FullName))
                {
                    padrao?.Dispose();
                    throw new InvalidOperationException("Nenhuma impressora padrão está configurada no Windows.");
                }
                using (padrao) return Estado(padrao, log);
            }
            // A fila e' lida ENQUANTO a colecao (e o servidor) estao vivos; todas sao descartadas no fim.
            using var filas = servidor.GetPrintQueues([EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections]);
            try
            {
                var fila = filas.FirstOrDefault(f => string.Equals(f.FullName, nome, StringComparison.OrdinalIgnoreCase))
                    ?? filas.FirstOrDefault(f => string.Equals(f.Name, nome, StringComparison.OrdinalIgnoreCase));
                if (fila is null) throw new InvalidOperationException($"A impressora \"{nome}\" não foi encontrada no Windows.");
                return Estado(fila, log);
            }
            finally { foreach (var f in filas) f.Dispose(); }
        }
        catch (Exception e) when (e is not InvalidOperationException)
        {
            log?.Invoke("IMPRESSAO falha ao consultar a fila: " + e);
            throw new InvalidOperationException(nome is null
                ? "Não foi possível consultar a impressora padrão do Windows."
                : $"Não foi possível consultar a impressora \"{nome}\".", e);
        }
    }

    private static EstadoDaFila Estado(PrintQueue fila, Action<string>? log)
    {
        var offline = false;
        try { fila.Refresh(); offline = fila.IsOffline; }
        catch (Exception e) when (e is PrintQueueException or PrintSystemException or COMException) { log?.Invoke("IMPRESSAO estado da fila ilegivel: " + e.Message); }
        var pergunta = false;
        try { pergunta = string.Equals(fila.QueuePort?.Name, "PORTPROMPT:", StringComparison.OrdinalIgnoreCase); }
        catch (Exception e) when (e is PrintQueueException or PrintSystemException or COMException) { log?.Invoke("IMPRESSAO porta da fila ilegivel: " + e.Message); }
        (double, double)? papel = null;
        try
        {
            if (fila.DefaultPrintTicket?.PageMediaSize is { Width: { } w, Height: { } h } && w > 0 && h > 0) papel = Retrato(w / 96, h / 96);
        }
        catch (Exception e) when (e is PrintQueueException or PrintSystemException or COMException) { log?.Invoke("IMPRESSAO bilhete padrao ilegivel: " + e.Message); }
        return new EstadoDaFila(fila.FullName, offline, pergunta, papel);
    }

    // ---------------------------------------------------------------------------------------------------
    // Dialogo "Imprimir" do Windows (usesystemdialog)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>
    /// O dialogo "Imprimir" do Windows, pre-preenchido com o painel, dono = <see cref="_dono"/>. A fila do painel e'
    /// procurada FORA da thread de interface (fix round 1: a enumeracao leva segundos com impressoras de rede); a
    /// fila escolhida e' descartada depois de lida. O <see cref="PrintDialog"/> do WPF usa
    /// <c>Application.Current.MainWindow</c> como dono do <c>PrintDlgEx</c> (nao tem parametro de dono; conferido no
    /// PresentationFramework 10.0.10): durante o <c>ShowDialog</c> a janela do documento vira a <c>MainWindow</c> e
    /// volta ao que era depois. MEDIDO (Task 5, Windows 11): com a troca, o "Imprimir" moderno fica modal para a
    /// janela do documento (o proxy modal e' dela, ela fica bloqueada) e em primeiro plano; SEM a troca, o proxy fica
    /// na boas-vindas escondida e a janela do documento continua clicavel. Sem <c>Application</c> (testes) o dono e'
    /// a area de trabalho e o Windows mostra o "Imprimir" classico. <c>null</c> = cancelado.
    /// </summary>
    internal async Task<ConfiguracaoEnvio?> EscolherNoDialogoDoSistemaAsync(OpcoesImpressao o)
    {
        (LocalPrintServer Servidor, PrintQueue Fila)? preSelecionada = null;
        if (o.Impressora is { } nome)
            preSelecionada = await Task.Run(() => AcharFila(nome, Log));
        try
        {
            var d = new PrintDialog { UserPageRangeEnabled = true, MinPage = 1, MaxPage = 9999 };
            try
            {
                if (preSelecionada is { } p) d.PrintQueue = p.Fila;
                PreencherBilhete(d, o);
            }
            catch (Exception e) when (e is PrintQueueException or PrintSystemException or COMException) { Log("IMPRESSAO pre-preenchimento do dialogo: " + e.Message); }

            var app = Application.Current;
            var principal = app?.MainWindow;
            var trocou = false;
            if (app is not null && _dono is not null && !ReferenceEquals(principal, _dono)) { app.MainWindow = _dono; trocou = true; }
            bool? ok;
            try { ok = d.ShowDialog(); }
            finally { if (trocou) app!.MainWindow = principal; }
            if (ok != true) return null;
            using var escolhida = d.PrintQueue;
            var estado = Estado(escolhida, Log);
            return DoDialogo(estado.Nome, d.PrintTicket, d.PageRangeSelection == PageRangeSelection.UserPages ? d.PageRange : null)
                with { Offline = estado.Offline, PerguntaOndeSalvar = estado.PerguntaOndeSalvar };
        }
        finally
        {
            if (preSelecionada is { } p) { p.Fila.Dispose(); p.Servidor.Dispose(); }
        }
    }

    /// <summary>A fila <paramref name="nome"/> e o servidor dela (vivos para o dialogo; quem chama descarta os dois),
    /// ou <c>null</c>. As demais filas enumeradas sao descartadas aqui. Nunca lanca.</summary>
    private static (LocalPrintServer, PrintQueue)? AcharFila(string nome, Action<string> log)
    {
        LocalPrintServer? servidor = null;
        try
        {
            servidor = new LocalPrintServer();
            using var filas = servidor.GetPrintQueues([EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections]);
            PrintQueue? achada = null;
            foreach (var f in filas)
            {
                if (achada is null && string.Equals(f.FullName, nome, StringComparison.OrdinalIgnoreCase)) achada = f;
                else f.Dispose();
            }
            if (achada is not null) return (servidor, achada);
        }
        catch (Exception e) { log("IMPRESSAO fila do painel nao achada para o dialogo: " + e.Message); }
        servidor?.Dispose();
        return null;
    }

    /// <summary>O dialogo abre com o que o painel ja' escolheu: copias, lados, cor, orientacao (se nao "auto") e um
    /// intervalo SIMPLES ("2-5" ou "3" — o PrintDialog do WPF so' guarda um intervalo; "1-3,5" fica "todas").</summary>
    private static void PreencherBilhete(PrintDialog d, OpcoesImpressao o)
    {
        var t = d.PrintTicket;
        t.CopyCount = o.Copias;
        t.Duplexing = o.Lados switch { "both-long" => Duplexing.TwoSidedLongEdge, "both-short" => Duplexing.TwoSidedShortEdge, _ => Duplexing.OneSided };
        t.OutputColor = o.Colorido ? OutputColor.Color : OutputColor.Grayscale;
        if (!o.OrientacaoAutomatica) t.PageOrientation = o.Paisagem ? PageOrientation.Landscape : PageOrientation.Portrait;
        if (o.Paginas is { } p && Regex.Match(p, @"^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$") is { Success: true } m)
        {
            var de = int.Parse(m.Groups[1].Value);
            var ate = m.Groups[2].Success ? int.Parse(m.Groups[2].Value) : de;
            if (de >= 1 && ate >= de) { d.PageRangeSelection = PageRangeSelection.UserPages; d.PageRange = new PageRange(de, ate); }
        }
    }

    // ---------------------------------------------------------------------------------------------------
    // Envio pela WebView2 oculta
    // ---------------------------------------------------------------------------------------------------

    private async Task EnviarAsync(string pdf, ConfiguracaoEnvio cfg, CancellationToken ct)
    {
        Log($"IMPRESSAO abrindo {Path.GetFileName(pdf)} ({new FileInfo(pdf).Length / 1024} KB) na WebView2 oculta para \"{cfg.Impressora}\"");
        await ComPdfOculto(_app, pdf, ArquivoLog, async (env, wv, guarda) =>
        {
            if (guarda.Bloqueadas.Count > 0) Log("IMPRESSAO guarda bloqueou no visualizador: " + string.Join(" | ", guarda.Bloqueadas));
            var s = env.CreatePrintSettings();
            s.PrinterName = cfg.Impressora;
            s.Copies = cfg.Copias;
            s.Duplex = cfg.Duplex;
            s.ColorMode = cfg.Cor;
            if (cfg.Orientacao is { } orientacao) s.Orientation = orientacao;
            if (cfg.LarguraPol is { } l && cfg.AlturaPol is { } a) { s.PageWidth = l; s.PageHeight = a; }
            if (cfg.Paginas is not null) s.PageRanges = cfg.Paginas;
            s.ShouldPrintHeaderAndFooter = false;
            s.ShouldPrintBackgrounds = true;
            // O "Salvar Saída de Impressão como" das impressoras de porta PORTPROMPT: (Microsoft Print to PDF, XPS)
            // e' aberto pelo PROCESSO DO NAVEGADOR do WebView2 (medido na S5) e fica sobre a janela que estiver em
            // primeiro plano quando o trabalho comeca (MEDIDO na Task 5: o dono dele e' essa janela, de qualquer
            // processo) — normalmente a do documento, que acabou de pedir a impressao. Esta permissao deixa o
            // navegador (outro processo) trazer o dialogo para a frente em vez de ele nascer atras do app.
            AllowSetForegroundWindow(wv.BrowserProcessId);
            using var vigia = AntesDoEnvioParaTeste?.Invoke(wv);
            ct.ThrowIfCancellationRequested();
            var limite = LimiteEnvioParaTeste ?? (cfg.PerguntaOndeSalvar ? LimiteComDialogo : LimiteEnvio);
            CoreWebView2PrintStatus status;
            try
            {
                var envio = EnvioParaTeste is { } teste ? teste(env, wv, s) : wv.PrintAsync(s);
                status = await envio.WaitAsync(limite, ct);
            }
            catch (TimeoutException)
            {
                // A WebView2 oculta e' fechada pelo finally de ComPdfOculto (o que cancela o trabalho no navegador) e
                // quem chamou apaga o PDF.
                Log($"IMPRESSAO tempo esgotado ({limite.TotalMinutes:0.#} min) esperando \"{cfg.Impressora}\"; envio cancelado");
                throw new InvalidOperationException($"A impressora \"{cfg.Impressora}\" não respondeu em {Minutos(limite)}; o envio foi cancelado.");
            }
            catch (ArgumentException e)
            {
                Log("IMPRESSAO PrintAsync recusou as opcoes: " + e);
                throw new InvalidOperationException($"Intervalo de páginas inválido: \"{cfg.Paginas}\".", e);
            }
            catch (COMException e)
            {
                Log($"IMPRESSAO PrintAsync falhou (HRESULT 0x{e.HResult:X8}): {e}");
                throw new InvalidOperationException($"A impressora \"{cfg.Impressora}\" não aceitou o trabalho.", e);
            }
            Log($"IMPRESSAO PrintAsync({cfg.Impressora}, paginas={cfg.Paginas ?? "todas"}, copias={cfg.Copias}, {cfg.Duplex}, {cfg.Cor}, {cfg.Orientacao?.ToString() ?? "auto"}, offline={cfg.Offline}) = {status}");
            if (status == CoreWebView2PrintStatus.PrinterUnavailable)
                throw new InvalidOperationException($"A impressora \"{cfg.Impressora}\" não está disponível.");
            if (status != CoreWebView2PrintStatus.Succeeded)
                throw new InvalidOperationException(cfg.Paginas is null
                    ? $"A impressora \"{cfg.Impressora}\" não conseguiu imprimir o documento."
                    : $"A impressora \"{cfg.Impressora}\" não conseguiu imprimir o documento (confira o intervalo de páginas \"{cfg.Paginas}\").");
        }, ct, (wv, nav, c) => EsperarVisualizadorAsync(wv, nav, new FileInfo(pdf).Length, Log, c));
    }

    private static string Minutos(TimeSpan t) => t.TotalMinutes >= 1 ? $"{t.TotalMinutes:0} minuto(s)" : $"{t.TotalSeconds:0} segundo(s)";

    // ---------------------------------------------------------------------------------------------------
    // Visualizador de PDF: sinal de "documento carregado" (fix round 1)
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Limite para o visualizador carregar o PDF, proporcional ao tamanho: 15 s + 2 s por MB, no maximo 90 s.
    /// MEDIDO (Task 5, fix round 1): o sinal chega em 0,45–1,8 s tanto para 3 paginas (98 KB) quanto para 300
    /// paginas com imagens (6,2 MB), com o ambiente frio.</summary>
    internal static TimeSpan LimiteVisualizador(long bytes) =>
        TimeSpan.FromSeconds(Math.Min(90, 15 + 2.0 * bytes / (1024 * 1024)));

    /// <summary>
    /// Espera o visualizador de PDF embutido terminar de abrir o documento (fix round 1: substitui a folga fixa de
    /// 2 s). MEDIDO no runtime fixo 153.0.4234.48 (<c>Target.getTargets</c> pelo DevTools Protocol da propria
    /// WebView2): a pagina do PDF (<c>type "page"</c>, url = o PDF) ganha um alvo filho <c>type "webview"</c> — o
    /// visualizador <c>chrome-extension://…/edge_pdf/index.html</c> — cujo TITULO passa de vazio para a url da
    /// extensao e, quando o documento foi lido, para o titulo do PDF (o nome do arquivo). Nao ha' evento para essa
    /// ultima troca (<c>Target.targetInfoChanged</c> NAO dispara para ela — medido), entao a lista e' consultada a
    /// cada 100 ms. Esgotado <see cref="LimiteVisualizador"/>: erro pt-BR — nao se manda para a impressora uma
    /// pagina que pode ser so' a moldura vazia.
    ///
    /// Medido tambem: o <c>PrintToPdfAsync</c> chamado logo no <c>ContentLoading</c> (antes do sinal) ja' sai com as
    /// 300 paginas — o Chromium espera o visualizador por conta propria. O sinal e' a garantia explicita, nao a
    /// correcao de uma falha observada.
    /// </summary>
    internal static async Task EsperarVisualizadorAsync(CoreWebView2 wv, Task navegacao, long bytes, Action<string>? log, CancellationToken ct)
    {
        var url = "";
        void AoIniciar(object? s, CoreWebView2NavigationStartingEventArgs e) => url = e.Uri;
        wv.NavigationStarting += AoIniciar;
        try
        {
            await navegacao;
            if (url.Length == 0) url = wv.Source; // chamado depois de a navegacao comecar: a url atual e' a do PDF
            var relogio = System.Diagnostics.Stopwatch.StartNew();
            var limite = LimiteVisualizador(bytes);
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                var alvos = await wv.CallDevToolsProtocolMethodAsync("Target.getTargets", "{}");
                if (VisualizadorCarregou(alvos, url))
                {
                    log?.Invoke($"IMPRESSAO visualizador pronto em {relogio.ElapsedMilliseconds} ms depois da navegacao ({bytes / 1024} KB)");
                    return;
                }
                if (relogio.Elapsed > limite)
                {
                    log?.Invoke($"IMPRESSAO visualizador nao carregou em {limite.TotalSeconds:0} s; alvos: {(alvos.Length > 2000 ? alvos[..2000] : alvos)}");
                    throw new InvalidOperationException("O visualizador de PDF não terminou de abrir o documento de impressão.");
                }
                await Task.Delay(100, ct);
            }
        }
        finally { wv.NavigationStarting -= AoIniciar; }
    }

    /// <summary>A resposta de <c>Target.getTargets</c> mostra o visualizador com o documento <paramref name="urlDoPdf"/>
    /// carregado? Exige o alvo <c>page</c> DESTE PDF (outra impressao simultanea tem outro arquivo) e um alvo
    /// <c>webview</c> filho dele, da extensao, com titulo que nao e' vazio, nem url de extensao, nem a propria url.
    /// Nunca lanca (JSON inesperado = nao).</summary>
    internal static bool VisualizadorCarregou(string alvosJson, string urlDoPdf)
    {
        try
        {
            using var d = JsonDocument.Parse(alvosJson);
            if (d.RootElement.ValueKind != JsonValueKind.Object || !d.RootElement.TryGetProperty("targetInfos", out var lista) || lista.ValueKind != JsonValueKind.Array) return false;
            string? pagina = null;
            foreach (var t in lista.EnumerateArray())
                if (t.ValueKind == JsonValueKind.Object && Texto(t, "type") == "page" && MesmaUrl(Texto(t, "url"), urlDoPdf)) { pagina = Texto(t, "targetId"); break; }
            if (pagina is null) return false;
            foreach (var t in lista.EnumerateArray())
            {
                if (t.ValueKind != JsonValueKind.Object || Texto(t, "type") != "webview" || Texto(t, "parentId") != pagina) continue;
                var u = Texto(t, "url") ?? "";
                var titulo = Texto(t, "title") ?? "";
                if (!u.StartsWith("chrome-extension://", StringComparison.OrdinalIgnoreCase)) continue;
                if (titulo.Length > 0 && !titulo.StartsWith("chrome-extension://", StringComparison.OrdinalIgnoreCase) && titulo != u && !MesmaUrl(titulo, urlDoPdf))
                    return true;
            }
            return false;
        }
        catch (JsonException) { return false; }
    }

    private static string? Texto(JsonElement e, string campo) =>
        e.TryGetProperty(campo, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static bool MesmaUrl(string? a, string b) =>
        a is not null && b.Length > 0 && string.Equals(Uri.UnescapeDataString(a), Uri.UnescapeDataString(b), StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// WebView2 OCULTA com o PDF aberto no visualizador embutido: janela WPF nunca mostrada, fora da tela e SEM dono
    /// (uma janela com dono morre junto com ele — fechar o documento no meio da impressao abortaria o envio),
    /// controller invisivel, no MESMO ambiente da thread (runtime fixo, flags) e com a guarda de egresso instalada
    /// (a guarda recebe o log <paramref name="arquivoLog"/>). O visualizador so' pede o <c>file:///</c> do PDF (S5:
    /// zero bloqueios). <paramref name="esperarVisualizador"/> nasce ANTES da navegacao (recebe a tarefa dela) e
    /// diz quando o documento esta' pronto; o padrao e' <see cref="EsperarVisualizadorAsync"/>. Tudo e' fechado no
    /// fim, com sucesso ou nao. Usado pelo envio e pelas sondas dos testes.
    /// </summary>
    internal static async Task ComPdfOculto(string app, string pdf, string arquivoLog,
        Func<CoreWebView2Environment, CoreWebView2, GuardaEgresso, Task> corpo, CancellationToken ct = default,
        Func<CoreWebView2, Task, CancellationToken, Task>? esperarVisualizador = null)
    {
        var janela = new Window { Width = 800, Height = 600, ShowInTaskbar = false, WindowStyle = WindowStyle.None, Left = -20000, Top = -20000, ShowActivated = false };
        var hwnd = new WindowInteropHelper(janela).EnsureHandle();
        CoreWebView2Controller? ctl = null;
        try
        {
            var env = await AmbienteWebView2.ObterAsync(app);
            ctl = await env.CreateCoreWebView2ControllerAsync(hwnd);
            ctl.Bounds = new System.Drawing.Rectangle(0, 0, 800, 600);
            ctl.IsVisible = false;
            var wv = ctl.CoreWebView2;
            wv.Settings.AreDefaultContextMenusEnabled = false;
            wv.Settings.IsReputationCheckingRequired = false;
            var guarda = new GuardaEgresso();
            guarda.Instalar(wv, arquivoLog);
            var nav = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            wv.NavigationCompleted += (_, e) => nav.TrySetResult(e.IsSuccess);
            async Task NavegouAsync()
            {
                bool ok;
                try { ok = await nav.Task.WaitAsync(TimeSpan.FromSeconds(30), ct); }
                catch (TimeoutException) { ok = false; }
                if (!ok) throw new InvalidOperationException("O visualizador de PDF não abriu o documento de impressão.");
            }
            var navegou = NavegouAsync();
            var bytes = new FileInfo(pdf).Length;
            var pronto = (esperarVisualizador ?? ((w, n, c) => EsperarVisualizadorAsync(w, n, bytes, null, c)))(wv, navegou, ct);
            wv.Navigate(new Uri(pdf).AbsoluteUri);
            await pronto;
            await corpo(env, wv, guarda);
        }
        finally
        {
            ctl?.Close();
            janela.Close();
        }
    }

    private void Log(string linha)
    {
        try { File.AppendAllText(ArquivoLog, $"{DateTime.Now:yyyy-MM-ddTHH:mm:ss.fff} {linha}{Environment.NewLine}"); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    [DllImport("user32.dll")] private static extern bool AllowSetForegroundWindow(uint pid);
}
