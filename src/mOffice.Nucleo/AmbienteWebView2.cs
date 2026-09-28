using System.Diagnostics;
using Microsoft.Web.WebView2.Core;

namespace mOffice.Nucleo;

/// <summary>
/// UM <see cref="CoreWebView2Environment"/> por THREAD DE INTERFACE, compartilhado por todas as janelas dela,
/// sobre o runtime WebView2 FIXO embarcado. No produto a suite tem UMA thread de interface, entao isto e' um
/// ambiente por PROCESSO — que e' o objetivo. A chave e' a thread, e nao o processo, por uma razao medida na
/// Task 3: o <c>CoreWebView2Environment</c> e' um objeto COM de apartamento, preso a' thread que o criou; um
/// cache estatico simples devolvia a uma SEGUNDA thread de interface um RCW ja' desligado
/// ("COM object that has been separated from its underlying RCW"). Acontece na bancada de testes, onde cada
/// teste de janela real roda numa thread STA propria que morre no fim — e aconteceria em qualquer app que
/// abrisse uma segunda thread de interface.
///
/// Duas razoes, as duas medidas:
///   1. Egresso (spec 7.2, camada 1 da Task 3). Com o runtime Evergreen o processo do navegador do WebView2
///      abria conexoes TLS/443 para a Microsoft POR JANELA, mesmo com --disable-background-networking: e' o
///      runtime do sistema, ligado ao servico de atualizacao do Edge. O runtime fixo e' uma copia nossa,
///      sem atualizador.
///   2. Memoria. Um ambiente por janela significa uma ARVORE de processos do navegador por janela. Um
///      ambiente unico faz todas as janelas dividirem o mesmo processo de navegador.
///
/// Se o runtime fixo nao estiver presente (maquina de desenvolvimento onde ninguem rodou
/// tools/webview2/obter-runtime.ps1), cai no Evergreen do sistema — mas com AVISO em <see cref="Aviso"/> e no
/// log de quem chamar. Nunca em silencio: a diferenca entre os dois e' exatamente o que a Task 3 mediu.
/// </summary>
public static class AmbienteWebView2
{
    private static readonly SemaphoreSlim Trava = new(1, 1);
    // Chave: a thread de interface dona. O objeto Thread (e nao so' o id) fica guardado porque ids de thread
    // sao REAPROVEITADOS depois que a thread morre — comparar so' o id entregaria um ambiente morto a uma
    // thread nova que calhasse de herdar o numero.
    private static readonly Dictionary<int, (Thread Dona, CoreWebView2Environment Ambiente)> _ambientes = [];

    /// <summary>Verdadeiro depois que <see cref="ObterAsync"/> criou o ambiente sobre o runtime embarcado.</summary>
    public static bool UsandoRuntimeFixo { get; private set; }

    /// <summary>Mensagem em pt-BR quando o ambiente NAO subiu sobre o runtime fixo; <c>null</c> quando subiu.</summary>
    public static string? Aviso { get; private set; }

    /// <summary>Origens locais que podem falar com o esquema nativo <c>ascdesktop://</c> e com os objetos de
    /// host da ponte. UMA lista so' para os dois: o registro do esquema e' por AMBIENTE (mora aqui) e o
    /// AddHostObjectToScript e' por janela (JanelaEditor le daqui) — se as duas divergissem, o editor
    /// carregaria as fontes de uma origem que nao pode chamar a ponte, ou o contrario.
    /// <c>https://moffice-doc.local</c> (a pasta de recovery: Editor.bin e a midia que o x2t gerou, inclusive
    /// SVG derivado de EMF/WMF do documento) fica DE FORA de proposito (fix final da revisao do branch): e'
    /// conteudo derivado do documento, so' e' buscado por fetch/img a partir da pagina do editor e nao precisa
    /// nem da ponte nem das fontes — deixa-la aqui daria os objetos de host "ponte"/"log" e o esquema
    /// ascdesktop:// a qualquer quadro que viesse a rodar naquela origem.</summary>
    public static readonly string[] OrigensPonte = ["https://moffice.local", "https://moffice-poc.local"];

    /// <summary>Cache de <see cref="PastaRuntimeFixo"/> — calculado uma vez por PROCESSO (fix da revisao 1,
    /// achado 7). A pasta de um runtime fixo ja' presente no disco nao muda em tempo de execucao, e o calculo
    /// varre diretorios (e, no caso "instalado ausente", ainda chama <see cref="mOffice.Conversor.ConfiguracaoMotor.Localizar"/>,
    /// que teria custo repetido a cada leitura). <see cref="Lazy{T}"/> com o valor padrao e' thread-safe e
    /// executa o calculo uma unica vez.</summary>
    private static readonly Lazy<string?> _pastaRuntimeFixo = new(CalcularPastaRuntimeFixo);

    /// <summary>
    /// Pasta que contem o <c>msedgewebview2.exe</c> embarcado, ou <c>null</c> se nao houver nenhum.
    /// Instalado: <c>&lt;pasta do exe&gt;\WebView2Runtime</c> (ou uma subpasta dela, que e' como o .cab
    /// extrai). Desenvolvimento: <c>third_party\webview2\&lt;versao&gt;</c>, a versao mais alta encontrada.
    /// NUNCA lanca (fix da revisao 1, achado 7): "sem runtime fixo" e' um estado valido (cai no Evergreen com
    /// aviso, ver <see cref="ObterAsync"/>), inclusive numa maquina sem o motor ONLYOFFICE extraido.
    /// </summary>
    public static string? PastaRuntimeFixo => _pastaRuntimeFixo.Value;

    private static string? CalcularPastaRuntimeFixo()
    {
        var instalado = Path.Combine(AppContext.BaseDirectory, "WebView2Runtime");
        var noInstalado = ProcurarExe(instalado);
        if (noInstalado is not null) return noInstalado;

        string dev;
        try { dev = Path.GetFullPath(Path.Combine(mOffice.Conversor.ConfiguracaoMotor.Localizar().PastaMotor, "..", "webview2")); }
        // ConfiguracaoMotor.Localizar() lanca ErroConversao quando third_party/onlyoffice nao existe (motor
        // ONLYOFFICE ausente). Isso nao diz nada sobre o runtime WebView2 — antes desta correcao, uma maquina
        // com o runtime fixo baixado mas sem o motor extraido (ou os testes rodando antes de qualquer um dos
        // dois) via' esta propriedade LANCAR, quebrando qualquer chamador que so' queria saber "tem runtime
        // fixo?" (inclusive FirewallTests, que precisa ler isto sem exigir o motor).
        catch (mOffice.Conversor.ErroConversao) { return null; }

        // Plano 5, Task 1: motor de um PACOTE montado apontado por MOFFICE_MOTOR_DIR (PacoteTests) — o runtime
        // do pacote e' irmao do motor (<pacote>\WebView2Runtime), e o teste do pacote tem de usar ESSE runtime,
        // nao cair no Evergreen.
        // FRAGIL: `dev` e' <motor>\..\webview2 so' LEXICAMENTE, entao `dev\..\WebView2Runtime` = <motor>\..\WebView2Runtime; se `dev` mudar de forma, recalcular a partir de PastaMotor.
        var irmaoDoMotor = ProcurarExe(Path.GetFullPath(Path.Combine(dev, "..", "WebView2Runtime")));
        if (irmaoDoMotor is not null) return irmaoDoMotor;

        if (!Directory.Exists(dev)) return null;
        // Ordenacao por VERSAO, nao por texto: "9.0.0.0" e' maior que "10.0.0.0" em ordem alfabetica.
        foreach (var v in Directory.EnumerateDirectories(dev).OrderByDescending(d =>
                     Version.TryParse(Path.GetFileName(d), out var ver) ? ver : new Version(0, 0)))
        {
            var achado = ProcurarExe(v);
            if (achado is not null) return achado;
        }
        return null;
    }

    /// <summary>Versao do <c>msedgewebview2.exe</c> embarcado (<c>null</c> se nao houver runtime fixo). E' o
    /// numero que <c>CoreWebView2Environment.BrowserVersionString</c> tem de devolver depois do
    /// <see cref="ObterAsync"/> — a prova de que o ambiente subiu sobre a NOSSA copia, e nao sobre o Edge do PC.</summary>
    public static string? VersaoRuntimeFixo
    {
        get
        {
            var pasta = PastaRuntimeFixo;
            if (pasta is null) return null;
            try { return FileVersionInfo.GetVersionInfo(Path.Combine(pasta, "msedgewebview2.exe")).FileVersion; }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return null; }
        }
    }

    /// <summary>
    /// Pasta de dados do WebView2, uma por app da suite, em <c>&lt;Pastas.Raiz&gt;\&lt;app&gt;\WebView2</c>
    /// (normalmente <c>%LOCALAPPDATA%\mOffice\&lt;app&gt;\WebView2</c>). Nunca em <c>%TEMP%</c>: o WebView2
    /// guarda ali cache, IndexedDB e o perfil, e um %TEMP% varrido por limpador de disco derrubaria a sessao
    /// em uso. Passa por <see cref="Pastas.Raiz"/> (Task 1 do Marco 1 / Plano 2): com
    /// <see cref="Pastas.VariavelDeAmbiente"/> definida, o perfil do WebView2 tambem vai para a pasta de
    /// teste, junto com o resto dos dados por usuario.
    /// </summary>
    /// <exception cref="ArgumentException">se <paramref name="app"/> nao for um nome simples de pasta.</exception>
    public static string PastaDadosUsuario(string app)
    {
        if (string.IsNullOrWhiteSpace(app) || app.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || app is "." or "..")
            throw new ArgumentException($"nome de app invalido para pasta de dados: '{app}'", nameof(app));
        return Path.Combine(Pastas.Raiz, app, "WebView2");
    }

    /// <summary>
    /// Linha de comando do processo do navegador, UMA flag por item, cada uma com o motivo escrito. Todos os
    /// nomes foram CONFERIDOS no binario do runtime pinado (153.0.4234.48): um <c>--disable-features=</c> com
    /// nome errado nao da' erro nenhum — e' ignorado em silencio, e vira culto de carga no codigo. (O
    /// <c>InterestFeedContentSuggestions</c> da hipotese do brief NAO existe neste build; saiu da lista.)
    /// </summary>
    public static IReadOnlyList<(string Flag, string Motivo)> Flags { get; } =
    [
        ("--disable-background-networking",
            "desliga o trafego de fundo do Chromium (atualizador de componentes, lista do Safe Browsing, sync, sondagem de captive portal)"),
        ("--disable-component-update",
            "nenhum componente (CRLSet, Widevine, dicionarios) e' baixado; com runtime fixo, tambem nao ha' o que atualizar"),
        ("--disable-domain-reliability",
            "nao envia relatorio de confiabilidade de dominio para a Google/Microsoft"),
        ("--no-pings",
            "nao envia o ping de auditoria de hiperlink (<a ping>) — o editor gera HTML e nao ha' motivo para essa saida"),
        ("--disable-sync",
            "nada de sincronizacao de perfil; o mOffice nao tem conta e nao quer nenhuma"),
        ("--disable-breakpad",
            "relatorio de falha do Chromium nao e' enviado — uma queda do editor nao pode virar upload"),
        ("--disable-features=msSmartScreenProtection,OptimizationHints,OptimizationHintsFetching,OptimizationGuideModelExecution,NetworkTimeServiceQuerying,AutofillServerCommunication,msImplicitSignin",
            "recursos que falam com servidor por conta propria: SmartScreen (consulta de reputacao de URL — desligado tambem pela CONFIGURACAO oficial da API, wv.Settings.IsReputationCheckingRequired = false, em JanelaEditor.InicializarAsync; as DUAS formas juntas, flag E configuracao, sao a base do aviso dispensado na Secao 8 da licenca do runtime — fix da revisao 1, achado 2: uma mudanca futura do Chromium que pare de respeitar a flag ainda encontra a configuracao), Optimization Guide (baixa dicas e modelos), servico de hora do Chromium, preenchimento automatico servidor-a-servidor e login implicito com a conta do Windows"),
    ];

    /// <summary>A linha montada a partir de <see cref="Flags"/> — nada entra por fora da lista documentada.</summary>
    public static string ArgumentosNavegador => string.Join(' ', Flags.Select(f => f.Flag));

    /// <summary>
    /// Devolve o ambiente da thread de interface que chama (no produto, um so' para o processo inteiro),
    /// criando-o na primeira chamada. Chamadas seguintes da MESMA thread devolvem o MESMO objeto, mesmo com
    /// <paramref name="app"/> diferente — a pasta de dados e' a do primeiro chamador, que e' o app dono do
    /// processo. <paramref name="log"/> recebe o aviso quando o runtime fixo nao esta' presente.
    /// </summary>
    public static async Task<CoreWebView2Environment> ObterAsync(string app, Action<string>? log = null)
    {
        // A thread e' capturada ANTES do await: e' ela que vai ser dona do objeto COM criado abaixo. Este
        // metodo e' chamado do `Loaded` da janela, ou seja, sempre com o contexto do Dispatcher da interface,
        // que devolve a continuacao a esta mesma thread.
        var dona = Thread.CurrentThread;
        await Trava.WaitAsync();
        try
        {
            if (_ambientes.TryGetValue(dona.ManagedThreadId, out var existente) && ReferenceEquals(existente.Dona, dona))
                return existente.Ambiente;

            var esquemaFontes = new CoreWebView2CustomSchemeRegistration("ascdesktop")
            {
                TreatAsSecure = true,
                // AllowedOrigins de uma registration nova chega NULO no WebView2 1.0.2903.40: tem de ser
                // ATRIBUIDO (`= [...]`), nunca inicializado com `{ ... }` (viraria NullReferenceException).
                HasAuthorityComponent = true,
                AllowedOrigins = [.. OrigensPonte],
            };
            // CustomSchemeRegistrations e' somente-leitura E chega NULA quando as opcoes sao criadas com
            // inicializador de objeto (o unico construtor recebe a lista) — por isso tudo vai pelo construtor.
            var opcoes = new CoreWebView2EnvironmentOptions(
                additionalBrowserArguments: ArgumentosNavegador,
                language: "pt-BR",
                customSchemeRegistrations: [esquemaFontes]);

            var fixo = PastaRuntimeFixo;
            UsandoRuntimeFixo = fixo is not null;
            Aviso = fixo is not null ? null
                : "runtime fixo ausente — usando o WebView2 do sistema (Evergreen). O egresso zero da Task 3 "
                  + "NAO esta' garantido: rode tools/webview2/obter-runtime.ps1.";
            if (Aviso is not null) log?.Invoke("AMBIENTE WEBVIEW2: " + Aviso);

            var ambiente = await CoreWebView2Environment.CreateAsync(fixo, PastaDadosUsuario(app), opcoes);
            // Entradas de threads que ja' morreram nunca mais serao usadas (e a comparacao por referencia
            // acima ja' as ignora); tira-las daqui evita que o dicionario cresca a cada janela de teste.
            foreach (var morto in _ambientes.Where(e => !e.Value.Dona.IsAlive).Select(e => e.Key).ToArray())
                _ambientes.Remove(morto);
            _ambientes[dona.ManagedThreadId] = (dona, ambiente);
            return ambiente;
        }
        finally { Trava.Release(); }
    }

    /// <summary>Pasta dada, ou a primeira subpasta dela, que contenha o <c>msedgewebview2.exe</c> — e' assim
    /// que o .cab do runtime fixo extrai (uma pasta <c>Microsoft.WebView2.FixedVersionRuntime.&lt;v&gt;.x64</c>
    /// por dentro). Nao varre recursivamente a arvore inteira: a arvore do runtime tem dezenas de milhares de
    /// arquivos e isto e' lido na subida da janela.</summary>
    private static string? ProcurarExe(string pasta)
    {
        if (!Directory.Exists(pasta)) return null;
        if (File.Exists(Path.Combine(pasta, "msedgewebview2.exe"))) return pasta;
        foreach (var sub in Directory.EnumerateDirectories(pasta))
            if (File.Exists(Path.Combine(sub, "msedgewebview2.exe"))) return sub;
        return null;
    }
}
