using System.Windows;
using Microsoft.Web.WebView2.Core;
using mOffice.Nucleo;
using mOffice.Conversor;

namespace mOffice.Sonda;

/// <summary>
/// Sem argumentos a janela abre vazia. Com um caminho de arquivo, abre o documento (Task 5A). Com
/// <c>--sonda &lt;word|cell|slide&gt;</c> roda a sonda da Task 4. Com <c>&lt;arquivo&gt; --medir</c> grava uma
/// linha de tempo de abertura em <c>poc/relatorios/medidas-host.txt</c> (Task 7, usado por tools/medir.ps1).
/// (O <c>StartupUri</c> saiu do App.xaml: quem cria a janela é o <see cref="OnStartup"/>.)
/// </summary>
public partial class App : Application
{
    /// <summary>Cronometro da medida "abertura do editor" (Task 7): comeca no <see cref="OnStartup"/> e e' lido
    /// quando o editor sinaliza documento pronto. Nao cobre o tempo que o .NET leva ate chegar aqui — por isso a
    /// linha gravada tambem traz `processo=`, contado desde <see cref="System.Diagnostics.Process.StartTime"/>,
    /// que e' o numero honesto de "a frio" para o usuario (ele clica no arquivo, nao no OnStartup).</summary>
    private static readonly System.Diagnostics.Stopwatch Relogio = new();

    /// <summary>Retencao de recuperacao DESTA ferramenta de bancada — explicitamente diferente da do produto
    /// (<c>AplicativoOffice.IdadeMinimaRecovery</c>, 30 dias). Ver o comentario em <see cref="OnStartup"/>.</summary>
    private static readonly TimeSpan IdadeMinimaRecoveryBancada = TimeSpan.FromDays(1);

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        Relogio.Restart();
        // Varredura de retencao (achado da revisao final 2026-09-17): pastas de recovery de execucoes
        // anteriores que nao foram apagadas no fechamento (queda, kill do processo) — cada uma guarda o
        // documento inteiro em claro. O frescor usado pela varredura (SessaoDocumento.VarrerRecoveryAntigo)
        // conta tambem a data dos ARQUIVOS dentro da pasta, e nao so' a da pasta em si (achado 2 da revisao
        // final, fix review 2): a pasta NAO muda de LastWriteTime quando o editor apenda
        // changes/changes0.json ou o x2t regrava Editor.bin.
        //
        // 1 dia (e nao os 30 do PRODUTO, AplicativoOffice.IdadeMinimaRecovery) e' uma escolha DE BANCADA: a
        // Sonda roda medidas de minutos e acumularia lixo pesado a cada rodada. Atencao (correcao da revisao
        // final do Plano 2): o comentario antigo daqui afirmava que uma pasta em uso "mesmo NUMA OUTRA
        // instancia deste app" nao seria apagada — isso e' FALSO enquanto a instancia viva estiver so'
        // ociosa. A data dos arquivos protege a pasta enquanto o editor GRAVA (changes, Editor.bin); um
        // documento apenas aberto, sem edicao, para de escrever e envelhece normalmente. Como a raiz de
        // recuperacao e' compartilhada, esta varredura de 1 dia so' e' aceitavel porque a Sonda e' ferramenta
        // de bancada — nao se roda a Sonda numa maquina com documento de trabalho aberto.
        // Nunca lanca.
        SessaoDocumento.VarrerRecoveryAntigo(SessaoDocumento.RaizRecovery, IdadeMinimaRecoveryBancada);
        // Varredura dos logs por janela (Task 1 do Marco 1 / Plano 2, backlog 15): mesma politica de retencao
        // por idade da varredura de recovery acima, so' que para Pastas.Logs. Nunca lanca.
        Logs.Varrer(Pastas.Logs, Logs.RetencaoPadrao);
        // --grampo (Task 1 do Marco 1 / Plano 2): liga a injecao do grampo.js (loga toda chamada da ponte) —
        // desligado por padrao no produto. Tem de ser lido ANTES de qualquer `new JanelaEditor(...)` abaixo
        // (inclusive do modo --abrir), porque GrampoLigado e' lido em JanelaEditor.InicializarAsync.
        if (e.Args.Contains("--grampo")) JanelaEditor.GrampoLigado = true;
        // Modo `--abrir <a> <b> <c> [--segundos N]` (Task 3/S2): abre N documentos em janelas do MESMO
        // processo para medir o ganho de memoria do ambiente unico do WebView2. Nao passa por
        // `new JanelaEditor()` abaixo porque ele abriria uma janela vazia a mais.
        if (e.Args.Length >= 1 && e.Args[0] == "--abrir") { _ = AbrirVariasAsync(e.Args); return; }

        var janela = new JanelaEditor(NomeApp);
        janela.Show();
        if (e.Args.Length >= 1 && e.Args[0] == "--sonda")
            _ = SondarAsync(janela, e.Args.Length >= 2 ? e.Args[1] : "word");
        else if (e.Args.Length >= 1 && !e.Args[0].StartsWith("--"))
        {
            var iSalvar = Array.IndexOf(e.Args, "--salvar");
            var destino = iSalvar >= 0 && iSalvar + 1 < e.Args.Length ? e.Args[iSalvar + 1] : null;
            _ = AbrirAsync(janela, e.Args[0], e.Args.Contains("--medir"), destino);
        }
    }

    /// <summary>Nome do app da suite a que a Sonda pertence: decide a pasta de dados do WebView2
    /// (%LOCALAPPDATA%\mOffice\sonda\WebView2) e aparece no relatorio de egresso.</summary>
    public const string NomeApp = "sonda";

    /// <summary>Pasta de ARTEFATOS DE BANCADA da Sonda (medidas de abertura/memoria, capturas da sondagem
    /// Task 4) — nao e' dado do produto (backlog 5/15), fica ao lado do motor extraido, em poc/relatorios,
    /// tal como antes da Task 1 do Marco 1 / Plano 2. So a Sonda (executavel de desenvolvimento) grava aqui;
    /// o nucleo (JanelaEditor) nao conhece mais este caminho — o log por janela agora e' <see cref="Logs"/>.</summary>
    public static string PastaRelatorios => Path.GetFullPath(Path.Combine(mOffice.Conversor.ConfiguracaoMotor.Localizar().PastaMotor, "..", "..", "poc", "relatorios"));

    /// <summary>
    /// `--abrir &lt;arq1&gt; [&lt;arq2&gt; ...] [--segundos N]` (Task 3/S2): abre uma janela por arquivo no MESMO
    /// processo, espera os documentos ficarem prontos, grava em `poc/relatorios/medidas-host.txt` a memoria da
    /// arvore de processos (ver <see cref="MedirArvore"/>) e sai. E' a medida do ganho do
    /// ambiente unico: no Marco 0, com um ambiente por janela, cada documento custava ~530 MB de arvore
    /// propria.
    /// </summary>
    private static async Task AbrirVariasAsync(string[] args)
    {
        // Fire-and-forget do OnStartup: uma excecao aqui nao tem quem a pegue e deixaria o processo de pe',
        // com janelas abertas, esperando para sempre — e a medida nunca seria gravada.
        try { await AbrirVariasNucleoAsync(args); }
        catch (Exception ex) { RegistrarLinha("ABRIR-VARIAS ERRO " + ex.Message); }
        finally { Current.Shutdown(); }
    }

    private static async Task AbrirVariasNucleoAsync(string[] args)
    {
        var iSegundos = Array.IndexOf(args, "--segundos");
        var segundos = iSegundos >= 0 && iSegundos + 1 < args.Length && int.TryParse(args[iSegundos + 1], out var n) ? n : 20;
        var arquivos = args.Skip(1).TakeWhile(a => !a.StartsWith("--")).ToArray();
        var prontas = new List<Task>();
        var janelas = new List<JanelaEditor>();
        foreach (var arquivo in arquivos)
        {
            var j = new JanelaEditor(NomeApp);
            janelas.Add(j);
            j.Show();
            var tcs = new TaskCompletionSource();
            j.DocumentoPronto += () => tcs.TrySetResult();
            prontas.Add(tcs.Task);
            try { await j.AbrirAsync(arquivo); }
            catch (Exception ex) { RegistrarLinha($"ABRIR-VARIAS ERRO {Path.GetFileName(arquivo)} {ex.Message}"); tcs.TrySetResult(); }
        }
        // Espera todos prontos (ou o tempo pedido, o que vier primeiro): a memoria so' vale medida com os
        // documentos carregados.
        await Task.WhenAny(Task.WhenAll(prontas), Task.Delay(TimeSpan.FromSeconds(segundos)));
        await Task.Delay(TimeSpan.FromSeconds(5)); // assentar o uso de memoria depois do carregamento
        RegistrarLinha(MedirArvore(arquivos.Length));
        foreach (var j in janelas) j.Close();
    }

    /// <summary>
    /// Memoria da "arvore": este processo mais TODO <c>msedgewebview2.exe</c> que esteja rodando a partir da
    /// NOSSA copia do runtime fixo. A identificacao e' pelo CAMINHO do executavel, e NAO por ascendencia do
    /// processo atual (fix da revisao 1, achado 8 — decisao registrada aqui e no relatorio da task, secao 6,
    /// "Fix (review 1)"): nenhuma outra instalacao da maquina roda o binario daquela pasta, e ler o caminho
    /// do modulo principal nao precisa de WMI nem de pacote extra (ao contrario de
    /// tools/egresso-monitor.ps1, que ja' tem <c>Get-CadeiaDeAncestrais</c> disponivel via
    /// <c>Get-CimInstance Win32_Process</c> — WMI que este processo C# nao quer trazer so' para esta medida
    /// de diagnostico). LIMITACAO CONHECIDA E ACEITA: se DUAS instancias desta Sonda rodarem ao mesmo tempo
    /// na mesma maquina (nao acontece na medida da Task 3, nem no uso normal do produto — cada app abre uma
    /// janela por vez), a medida por caminho contaria os processos <c>msedgewebview2.exe</c> da OUTRA
    /// instancia tambem, porque as duas rodam da mesma pasta de runtime fixo; a medida por ascendencia nao
    /// teria esse problema. Aceito porque este metodo so' e' chamado pelo modo `--abrir` de MEDIDA manual
    /// (nunca em producao) e quem roda a medida controla ter uma unica instancia no ar. So' numeros e nomes
    /// de processo entram na linha gravada: ela pode virar evidencia.
    /// </summary>
    private static string MedirArvore(int janelas)
    {
        var raiz = mOffice.Nucleo.AmbienteWebView2.PastaRuntimeFixo;
        long bytes = System.Diagnostics.Process.GetCurrentProcess().WorkingSet64;
        var contados = 1;
        foreach (var p in System.Diagnostics.Process.GetProcessesByName("msedgewebview2"))
        {
            try
            {
                var caminho = p.MainModule?.FileName;
                if (raiz is null || caminho is null || !caminho.StartsWith(raiz, StringComparison.OrdinalIgnoreCase)) continue;
                bytes += p.WorkingSet64;
                contados++;
            }
            // Processo que morreu entre a listagem e a leitura, ou de outro usuario/nivel de integridade.
            catch (Exception e) when (e is InvalidOperationException or System.ComponentModel.Win32Exception) { }
        }
        var fixo = raiz is null ? "evergreen" : "fixo";
        return $"ABRIR-VARIAS janelas={janelas} runtime={fixo} processos={contados} workingset={bytes / 1024 / 1024}MB";
    }

    private static void RegistrarLinha(string linha)
    {
        try
        {
            Directory.CreateDirectory(PastaRelatorios);
            File.AppendAllText(Path.Combine(PastaRelatorios, "medidas-host.txt"),
                $"{DateTime.Now:yyyy-MM-ddTHH:mm:ss} {linha}{Environment.NewLine}");
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    private static async Task AbrirAsync(JanelaEditor janela, string caminho, bool medir, string? destinoSalvar = null)
    {
        var uma = 0;
        if (medir || destinoSalvar is not null)
            janela.DocumentoPronto += () =>
            {
                if (Interlocked.Exchange(ref uma, 1) != 0) return;
                // Dentro da chamada sincrona do JS: so' le' o cronometro e escreve uma linha curta.
                if (medir) RegistrarMedida(caminho);
                // Salvar chama o WebView de volta (asc_Save) — NAO pode acontecer aqui dentro; volta pela fila
                // do Dispatcher, como o ResponderAberturaAsync ja faz com o Task.Yield().
                if (destinoSalvar is not null)
                    janela.Dispatcher.BeginInvoke(new Func<Task>(() => SalvarESairAsync(janela, destinoSalvar)));
            };
        try { await janela.AbrirAsync(caminho); }
        catch (Exception ex) { MessageBox.Show(ex.Message, "mOffice", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    /// <summary>Modo `&lt;arquivo&gt; --salvar &lt;destino&gt;` (Task 7): abre, pede o salvamento ao editor e sai.
    /// E' o que tools/fidelidade.ps1 usa para produzir o "salvo pelo PoC" de cada amostra — round-trip pelo
    /// EDITOR (asc_Save -> LocalFileSave -> x2t), nao pelo x2t direto.</summary>
    private static async Task SalvarESairAsync(JanelaEditor janela, string destino)
    {
        var codigo = 0;
        try
        {
            var fmt = mOffice.Conversor.FormatoDocumentoExtensoes.DaExtensao(destino);
            await janela.SalvarComoAsync(destino, fmt);
        }
        catch (Exception ex)
        {
            codigo = 2;
            try { await File.AppendAllTextAsync(Path.Combine(PastaRelatorios, "medidas-host.txt"),
                    $"{DateTime.Now:yyyy-MM-ddTHH:mm:ss} {Path.GetFileName(destino)} ERRO-SALVAR {ex.Message}{Environment.NewLine}"); }
            catch (Exception e2) when (e2 is IOException or UnauthorizedAccessException) { }
        }
        Current.Shutdown(codigo);
    }

    /// <summary>Uma linha por abertura em `poc/relatorios/medidas-host.txt`. So o NOME do arquivo entra (o
    /// arquivo pode ser commitado; caminho pessoal nao vai para o git).</summary>
    private static void RegistrarMedida(string caminho)
    {
        try
        {
            var desdeOnStartup = (long)Relogio.Elapsed.TotalMilliseconds;
            long desdeProcesso;
            try { desdeProcesso = (long)(DateTime.Now - System.Diagnostics.Process.GetCurrentProcess().StartTime).TotalMilliseconds; }
            catch (Exception e) when (e is InvalidOperationException or System.ComponentModel.Win32Exception) { desdeProcesso = -1; }
            var linha = $"{DateTime.Now:yyyy-MM-ddTHH:mm:ss} {Path.GetFileName(caminho)} pronto={desdeOnStartup}ms processo={desdeProcesso}ms";
            Directory.CreateDirectory(PastaRelatorios);
            File.AppendAllText(Path.Combine(PastaRelatorios, "medidas-host.txt"), linha + Environment.NewLine);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* medida e' diagnostico: nunca derruba a abertura */ }
    }

    /// <summary>
    /// Sonda time-boxed (Task 4): navega para a MESMA página de entrada que o app oficial usa
    /// (<c>index.html?placement=desktop…</c>, sem <c>?url=</c> — então <c>document.url</c> vira
    /// <c>'_offline_'</c>) e observa o que o build desktop pede à ponte.
    ///
    /// Como ninguém está olhando a tela quando isto roda (sessão de agente), a observação é gravada em
    /// arquivo: PNG da janela em 8 s e 20 s (<see cref="CoreWebView2.CapturePreviewAsync"/>), despejo de
    /// título/texto/iframes por JS, e rede+console pelo <see cref="SondaInstrumentacao"/>.
    ///
    /// As chamadas de ponte saem pelo grampo.js. Medido nesta sonda: dentro do iframe do editor o
    /// <c>arquivos.Log</c> NÃO funciona (o <c>AddHostObjectToScript</c> do CoreWebView2 só alcança o documento
    /// de nível superior — usar o objeto de dentro do quadro dá <c>0x80070490</c>, "Elemento não encontrado"),
    /// então o grampo cai no <c>console.log</c> e as linhas aparecem como <c>CONSOLE log [grampo] …</c> no log
    /// da sonda; no log de egresso só entra a linha do quadro de cima. A Task 5A precisa expor a ponte ao
    /// quadro (<c>CoreWebView2Frame.AddHostObjectToScriptWithOrigins</c>, no evento <c>FrameCreated</c>).
    /// </summary>
    private static async Task SondarAsync(JanelaEditor janela, string doctype)
    {
        var sonda = new SondaInstrumentacao(Path.Combine(PastaRelatorios, $"sonda-{doctype}.log"));
        try
        {
            var wv = await janela.ProntoAsync();
            await sonda.InstalarAsync(wv);
            // O motor copiado não tem sdkjs/common/AllFonts.js — o app oficial GERA esse catálogo em
            // %LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\data\fonts no primeiro lançamento (Preparação item 2 do
            // protocolo), e sem ele o requirejs morre em "allfonts". Interceptar o host mapeado não funciona (ver
            // o comentário de SondaInstrumentacao). Desde o fix round 1 da Task 1 do Plano 5 a JanelaEditor
            // reescreve o <script src> do AllFonts.js para um host NÃO mapeado e o serve do catálogo
            // (CatalogoNoEditor) — nada é copiado para dentro do motor.
            var filetype = doctype switch { "cell" => "xlsx", "slide" => "pptx", _ => "docx" };
            var url = $"https://moffice.local/web-apps/apps/api/documents/index.html?placement=desktop&doctype={doctype}"
                    + $"&filetype={filetype}&title=sonda&lang=pt-BR&mode=edit&userid=1&username=mOffice";
            sonda.Registrar("SONDA navegando " + url);
            await janela.NavegarAsync(new Uri(url));
            sonda.Registrar("SONDA navegacao de nivel superior concluida");
            var decorrido = 0;
            foreach (var t in new[] { 8, 20 })
            {
                await Task.Delay(TimeSpan.FromSeconds(t - decorrido));
                decorrido = t;
                var png = Path.Combine(PastaRelatorios, $"sonda-{doctype}-{t}s.png");
                using (var arquivo = File.Create(png))
                    await wv.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, arquivo);
                sonda.Registrar($"SONDA captura {t}s -> {Path.GetFileName(png)}");
                sonda.Registrar($"SONDA estado {t}s " + await wv.ExecuteScriptAsync(Diagnostico));
            }
        }
        catch (Exception ex)
        {
            sonda.Registrar("SONDA ERRO " + ex);
        }
        finally
        {
            sonda.Registrar("SONDA fim");
            Current.Shutdown();
        }
    }

    /// <summary>Despejo de estado da página: o editor vive num iframe de mesma origem, então dá para ler o corpo dele.</summary>
    private const string Diagnostico = """
        (function () {
          function corpo(d) { try { return d && d.body ? d.body.innerText.replace(/\s+/g, " ").slice(0, 2000) : ""; } catch (e) { return "erro: " + e.message; } }
          var quadros = Array.prototype.map.call(document.querySelectorAll("iframe"), function (f) {
            var d = null; try { d = f.contentDocument; } catch (e) {}
            // Prova se a ponte existe DE DENTRO do quadro do editor (e nao so no documento de cima):
            // tenta de verdade chamar o objeto de host. Se der certo, a linha aparece no log de egresso.
            var ponte = "?";
            try {
              var ho = d.defaultView.chrome.webview.hostObjects;
              ho.sync.arquivos.Log("sonda: chamada do objeto de host de dentro do iframe do editor");
              ponte = "chamada do iframe = ok";
            } catch (e) { ponte = "chamada do iframe = erro: " + e.message; }
            return { src: f.src, titulo: d ? d.title : null, pronto: d ? d.readyState : null, ponte: ponte, texto: corpo(d),
                     mascara: d && d.querySelector("#loading-mask, .asc-loadmask, #loadmask") ? true : false };
          });
          return JSON.stringify({ titulo: document.title, texto: corpo(document), iframes: quadros.length, quadros: quadros });
        })()
        """;

}
