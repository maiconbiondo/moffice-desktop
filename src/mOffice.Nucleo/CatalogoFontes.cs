using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>
/// Catalogo de fontes do motor (<c>AllFonts.js</c> + <c>font_selection.bin</c> + <c>fonts_thumbnail*.png</c>),
/// GERADO pelo mOffice com o <c>allfontsgen.exe</c> que ele embarca — sem depender do ONLYOFFICE instalado
/// na maquina (obrigacao 7.1 da spec do Marco 1).
///
/// De onde vem o gerador (sondagem S1a, Task 2 do Plano 1): o pacote do Desktop Editors 9.4.0.129 que o
/// mOffice embarca NAO traz o <c>allfontsgen.exe</c> — o app oficial gera o catalogo por dentro, no primeiro
/// lancamento, em <c>%LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\data\fonts</c>. O pacote do ONLYOFFICE Docs
/// (Document Server) para Windows, da MESMA versao 9.4.0.129 e do mesmo core AGPL-3.0, traz o gerador como
/// executavel solto; <c>tools/onlyoffice/obter-allfontsgen.ps1</c> o extrai sem instalar nada.
///
/// Opcoes REAIS do <c>allfontsgen.exe</c> 9.4.0.129 (o <c>--help</c> nao imprime nada — a lista saiu das
/// strings UTF-16 do binario e foi confirmada rodando): <c>--input</c> (repetivel, uma pasta por ocorrencia),
/// <c>--allfonts</c>, <c>--allfonts-web</c>, <c>--selection</c>, <c>--images</c>, <c>--output-web</c>,
/// <c>--use-system</c> (<c>true</c>/<c>false</c>), <c>--use-system-user-fonts</c>, <c>--font-format</c>.
///
/// Por que <c>--use-system=false</c> + <c>--input</c> explicito para cada pasta, e nao <c>--use-system=true</c>:
/// a impressao digital (que decide se regenera) e' calculada sobre as pastas de <see cref="_pastas"/>. Com
/// <c>--use-system=true</c> o gerador varre pastas que a impressao digital nao conhece — uma fonte nova numa
/// delas nunca dispararia regeneracao. MEDIDO nesta maquina: as duas formas dao o MESMO catalogo (638 fontes),
/// e a explicita varre metade dos arquivos (860 x 1.496 linhas no <c>fonts.log</c>, porque a forma
/// <c>use-system</c> visita as fontes da nuvem duas vezes).
///
/// Nunca lanca para o chamador: devolve <see cref="ResultadoCatalogo.Erro"/> em pt-BR (a tela de boas-vindas
/// mostra; nunca "Carregando documento" para sempre).
///
/// Revisao 1 (fix review 1, Marco 1/Plano 1/Task 2): o timeout interno agora MATA o processo do gerador em
/// vez de so' desistir dele (item 1); a impressao digital passa a incluir a identidade do PROPRIO gerador e
/// uma versao de formato (item 3); a regeneracao inteira roda sob um mutex nomeado entre processos (item 4).
/// </summary>
public sealed class CatalogoFontes
{
    private readonly string _gerador;
    private readonly IReadOnlyList<string> _pastas;
    private readonly Action<string>? _registrar;
    private readonly int _versaoFormato;

    /// <summary>Extensoes que contam para a impressao digital. Sao as mesmas que o motor sabe ler; um
    /// <c>fonts.log</c> ou um <c>.tmp</c> na pasta de fontes NAO pode disparar regeneracao (7,5 s por
    /// abertura, medidos nesta maquina).</summary>
    private static readonly string[] ExtensoesDeFonte = [".ttf", ".otf", ".ttc", ".fon"];

    /// <summary>Versao do FORMATO do catalogo que este <see cref="CatalogoFontes"/> espera. Dobra na
    /// impressao digital (<see cref="ImpressaoDigitalAtual"/>): mudar este numero no codigo forca regenerar
    /// TODO catalogo existente na proxima abertura, mesmo que nenhuma fonte tenha mudado — util se um dia os
    /// argumentos passados ao <c>allfontsgen</c> mudarem de um jeito que produz um <c>AllFonts.js</c>
    /// incompativel com o que o mOffice espera. E' parametro de construtor (ver <see cref="_versaoFormato"/>)
    /// em vez de so' esta constante, para o teste poder provar o efeito do bump sem <c>InternalsVisibleTo</c>
    /// nem reflexao.</summary>
    private const int VersaoFormatoCatalogo = 1;

    /// <summary>Nome do mutex nomeado entre processos que serializa a regeneracao (item 4 da revisao 1).
    /// <c>Global\</c> (e nao o prefixo de sessao padrao) porque a trava tem de valer entre processos de
    /// USUARIOS diferentes tambem: <c>%ProgramData%\mOffice\fontes</c> e' compartilhado pela maquina
    /// inteira.</summary>
    private const string NomeMutex = @"Global\mOffice-catalogo-fontes";

    public string PastaCatalogo { get; }

    private string ArquivoDigital => Path.Combine(PastaCatalogo, "impressao-digital.txt");
    private string ArquivoAllFonts => Path.Combine(PastaCatalogo, "AllFonts.js");
    private string ArquivoSelecao => Path.Combine(PastaCatalogo, "font_selection.bin");

    public CatalogoFontes(string pastaCatalogo, string allfontsgenExe, IReadOnlyList<string> pastasDeFontes,
        Action<string>? registrar = null, int versaoFormato = VersaoFormatoCatalogo)
    {
        PastaCatalogo = pastaCatalogo;
        _gerador = allfontsgenExe;
        _pastas = pastasDeFontes;
        _registrar = registrar;
        _versaoFormato = versaoFormato;
    }

    /// <summary>Os dois arquivos que o motor exige existem E a impressao digital gravada bate com a atual.</summary>
    public bool Atualizado
    {
        get
        {
            try
            {
                return File.Exists(ArquivoAllFonts) && File.Exists(ArquivoSelecao) && File.Exists(ArquivoDigital)
                    && File.ReadAllText(ArquivoDigital).Trim() == ImpressaoDigitalAtual();
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return false; }
        }
    }

    /// <summary>SHA256 de (versao do formato | identidade do gerador | caminho|tamanho|mtime de cada fonte),
    /// em ordem estavel.</summary>
    public string ImpressaoDigitalAtual()
    {
        var sb = new StringBuilder();

        // Item 3 da revisao 1: a versao do FORMATO do catalogo entra na conta. Sem isto, mudar os argumentos
        // do allfontsgen (ou o layout que o mOffice espera do AllFonts.js) nao forcaria regenerar catalogos
        // ja existentes cujas fontes nao mudaram.
        sb.Append("versaoFormato=").Append(_versaoFormato).Append('\n');

        // Idem para a identidade do PROPRIO gerador: o allfontsgen.exe pode ser trocado (nova versao pinada,
        // correcao de bug no gerador) sem que uma unica fonte do disco mude. "Tamanho + mtime" e' a mesma
        // heuristica barata usada abaixo para as fontes; nao compensa um hash de 32 KB a cada chamada so'
        // para isto. Falta o arquivo (ex.: teste que aponta para um gerador inexistente de proposito) nao e'
        // erro aqui — so' significa que essa parte da impressao digital fica vazia.
        try
        {
            var gerador = new FileInfo(_gerador);
            if (gerador.Exists)
                sb.Append("gerador=").Append(_gerador.ToLowerInvariant()).Append('|').Append(gerador.Length).Append('|').Append(gerador.LastWriteTimeUtc.Ticks).Append('\n');
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }

        foreach (var pasta in _pastas.Where(Directory.Exists).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            IEnumerable<string> arquivos;
            try { arquivos = Directory.EnumerateFiles(pasta, "*.*", SearchOption.AllDirectories).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray(); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { continue; }
            foreach (var f in arquivos)
            {
                if (!ExtensoesDeFonte.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)) continue;
                FileInfo i;
                try { i = new FileInfo(f); _ = i.Length; }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { continue; }
                sb.Append(f.ToLowerInvariant()).Append('|').Append(i.Length).Append('|').Append(i.LastWriteTimeUtc.Ticks).Append('\n');
            }
        }
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString()))).ToLowerInvariant();
    }

    /// <summary>Mensagem pt-BR para a espera do gerador que nao terminou a tempo — distingue o cancelamento
    /// PEDIDO PELO CHAMADOR (<paramref name="cancelamentoDoChamador"/> = <c>true</c>, o <c>ct</c> de
    /// <see cref="GarantirAtualizadoAsync"/>) do teto INTERNO de 5 minutos (item 1 da revisao 1). Extraido
    /// como funcao pura `internal` para o teste poder provar a distincao sem precisar de um processo real
    /// pendurado por 5 minutos.</summary>
    internal static string MensagemDeCancelamento(bool cancelamentoDoChamador) => cancelamentoDoChamador
        ? "geração do catálogo cancelada"
        : "O gerador de catálogo de fontes demorou mais de 5 minutos e foi interrompido.";

    /// <summary>Roda o <c>allfontsgen</c> se o catalogo faltar ou estiver defasado. Nunca lanca.</summary>
    public async Task<ResultadoCatalogo> GarantirAtualizadoAsync(CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();
        if (Atualizado) return new ResultadoCatalogo(false, ContarFontes(), sw.Elapsed, null);
        if (!File.Exists(_gerador))
            return new ResultadoCatalogo(false, 0, sw.Elapsed,
                $"O gerador de catálogo de fontes não foi encontrado em '{_gerador}'. Reinstale o mOffice.");

        try
        {
            // Toda a secao critica (mutex + processo + arquivos) roda numa UNICA thread de pool, do inicio ao
            // fim, sem nenhum 'await' no meio (ver GerarComMutex/GerarSincrono): Mutex.ReleaseMutex() so' pode
            // ser chamado pela MESMA thread que fez o WaitOne(), e a continuacao de um 'await' comum nao tem
            // essa garantia — o pool pode retomar a continuacao numa thread diferente, e o ReleaseMutex
            // estouraria (ou pior, liberaria o mutex errado). Isolar num delegate sincrono despachado por
            // Task.Run e' o jeito de manter essa afinidade sem reescrever Process/File para bloquear a
            // thread de quem chamou este metodo publico.
            return await Task.Run(() => GerarComMutex(sw, ct), CancellationToken.None);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            return new ResultadoCatalogo(false, 0, sw.Elapsed, "Não foi possível gerar o catálogo de fontes: " + e.Message);
        }
    }

    /// <summary>Adquire o mutex nomeado entre processos (item 4 da revisao 1) e, se conseguir, refaz a
    /// checagem de <see cref="Atualizado"/> antes de regenerar: outro processo (ou outra janela do mesmo
    /// processo) pode ter terminado enquanto este esperava a vez. Sem este recheck, dois chamadores
    /// concorrentes regenerariam os dois — nao e' so' trabalho em dobro, e' dois <c>allfontsgen</c> escrevendo
    /// o MESMO <c>AllFonts.js</c>/<c>font_selection.bin</c> ao mesmo tempo.</summary>
    private ResultadoCatalogo GerarComMutex(Stopwatch sw, CancellationToken ct)
    {
        using var mutex = new Mutex(false, NomeMutex);
        bool comMutex;
        try { comMutex = mutex.WaitOne(TimeSpan.FromMinutes(6)); }
        catch (AbandonedMutexException)
        {
            // O dono anterior morreu (crash, Kill externo, queda de energia) com o mutex tomado. O catalogo
            // dele pode estar pela metade — e' exatamente o que o recheck de Atualizado abaixo e o "apaga o
            // marcador primeiro" de GerarSincrono cobrem: um catalogo truncado nao passa em Atualizado, e a
            // regeneracao acontece de novo, desta vez ate o fim.
            comMutex = true;
        }
        if (!comMutex)
            return new ResultadoCatalogo(false, 0, sw.Elapsed,
                "Não foi possível obter acesso exclusivo ao catálogo de fontes: outro processo está gerando o catálogo há mais de 6 minutos.");
        try
        {
            if (Atualizado) return new ResultadoCatalogo(false, ContarFontes(), sw.Elapsed, null);
            return GerarSincrono(sw, ct);
        }
        finally { mutex.ReleaseMutex(); }
    }

    /// <summary>O trabalho de verdade: roda o <c>allfontsgen</c>, espera (observando o cancelamento do
    /// chamador e um teto interno de 5 minutos), e grava o catalogo. Chamado SOMENTE de dentro do mutex de
    /// <see cref="GerarComMutex"/>, numa unica thread sincrona (nenhum <c>await</c> aqui dentro).</summary>
    private ResultadoCatalogo GerarSincrono(Stopwatch sw, CancellationToken ct)
    {
        try
        {
            Directory.CreateDirectory(PastaCatalogo);

            // Minor 8 da revisao 1: apaga o marcador ANTES de regenerar (nao so' escreve o novo DEPOIS de
            // terminar). Fecha o buraco do marcador obsoleto: se o processo for interrompido (timeout, Kill,
            // queda) DEPOIS de o allfontsgen ja ter sobrescrito o AllFonts.js mas ANTES deste metodo gravar o
            // marcador novo, um marcador ANTIGO que por coincidencia bata com a impressao digital atual (as
            // fontes nao mudaram entre as duas tentativas) faria Atualizado mentir "sim" para um catalogo
            // truncado a meio caminho. Apagar antes garante que qualquer interrupcao deixa Atualizado == false.
            try { File.Delete(ArquivoDigital); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }

            // A impressao digital e' tirada ANTES de gerar: se uma fonte for instalada durante a geracao, a
            // digital gravada continua sendo a do conjunto que o catalogo realmente descreve, e a proxima
            // abertura regenera. Ao contrario (tirar depois), o catalogo ficaria eternamente marcado como
            // atual sem conter a fonte nova.
            var digital = ImpressaoDigitalAtual();

            var psi = new ProcessStartInfo(_gerador)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                WorkingDirectory = Path.GetDirectoryName(Path.GetFullPath(_gerador))!,
            };
            psi.ArgumentList.Add("--use-system=false");
            foreach (var p in _pastas.Where(Directory.Exists)) psi.ArgumentList.Add($"--input={p}");
            psi.ArgumentList.Add($"--allfonts={ArquivoAllFonts}");
            psi.ArgumentList.Add($"--selection={ArquivoSelecao}");
            psi.ArgumentList.Add($"--images={PastaCatalogo}");

            using var p2 = Process.Start(psi) ?? throw new InvalidOperationException("Process.Start devolveu null");
            // As leituras comecam JA, antes de esperar o processo sair: se o allfontsgen encher o buffer do
            // pipe antes de terminar, esperar sem ninguem lendo os streams trava os dois lados (o classico
            // deadlock de Process com streams redirecionados).
            var err = p2.StandardError.ReadToEndAsync();
            var saida = p2.StandardOutput.ReadToEndAsync();

            // Espera com um pe' em cada relogio: o cancelamento do CHAMADOR (ct) e o teto INTERNO de 5
            // minutos. WaitForExit(int) nao aceita CancellationToken; o poll curto (250ms) e' o jeito de
            // observar os dois sem um wrapper assincrono aqui dentro (este metodo roda inteiro numa unica
            // thread de pool, sem 'await' nenhum, de proposito — ver GarantirAtualizadoAsync).
            var prazo = DateTime.UtcNow + TimeSpan.FromMinutes(5);
            var saiu = false;
            while (DateTime.UtcNow < prazo)
            {
                if (ct.IsCancellationRequested) break;
                if (p2.WaitForExit(250)) { saiu = true; break; }
            }

            if (!saiu)
            {
                // Registrado ANTES do Kill: depois dele o processo pode ja nao existir mais para consultar nada.
                var cancelamentoDoChamador = ct.IsCancellationRequested;
                // Item 1 da revisao 1: sem isto o allfontsgen ficava vivo (preso ou nao) depois de o host
                // desistir dele — um processo orfao por timeout/cancelamento que continua escrevendo nos
                // MESMOS arquivos de saida que uma proxima tentativa (retry do chamador, ou a proxima
                // abertura) tambem vai escrever. entireProcessTree porque o allfontsgen pode ter filhos.
                try { p2.Kill(entireProcessTree: true); } catch { }
                try { p2.WaitForExit(5000); } catch { }
                // Drena os streams do processo morto: sem isto as Tasks de ReadToEndAsync ficam penduradas
                // esperando o pipe fechar e viram excecao nao observada quando o GC as coletar — em qualquer
                // dos dois casos (cancelado pelo chamador ou pelo teto interno).
                try { err.GetAwaiter().GetResult(); } catch { }
                try { saida.GetAwaiter().GetResult(); } catch { }
                return new ResultadoCatalogo(false, 0, sw.Elapsed, MensagemDeCancelamento(cancelamentoDoChamador));
            }

            var erroTexto = err.GetAwaiter().GetResult();
            _ = saida.GetAwaiter().GetResult(); // observado so' para nao virar Task perdida; o stdout do allfontsgen nao traz diagnostico util (fica no fonts.log dele mesmo)

            if (p2.ExitCode != 0 || !File.Exists(ArquivoAllFonts) || !File.Exists(ArquivoSelecao))
                return new ResultadoCatalogo(false, 0, sw.Elapsed,
                    $"O gerador de catálogo de fontes falhou (código {p2.ExitCode}). {erroTexto}".Trim());

            // O marcador so' e' escrito por ULTIMO, depois de confirmar os dois arquivos de saida no disco
            // (fecha o mesmo buraco do "apaga primeiro" acima, do outro lado: nunca existe um instante em que
            // o marcador aponte para uma digital cujo catalogo ainda nao terminou de ser gravado).
            File.WriteAllText(ArquivoDigital, digital);
            var fontes = ContarFontes();
            _registrar?.Invoke($"catálogo de fontes regenerado em {sw.Elapsed.TotalSeconds:0.0}s ({fontes} fontes) em {PastaCatalogo}");
            return new ResultadoCatalogo(true, fontes, sw.Elapsed, null);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            return new ResultadoCatalogo(false, 0, sw.Elapsed, "Não foi possível gerar o catálogo de fontes: " + e.Message);
        }
    }

    /// <summary>Quantos arquivos de fonte o <c>AllFonts.js</c> gerado nomeia (mesma leitura que a whitelist
    /// do <see cref="ServidorAscDesktop"/> faz — uma conta so', nao duas definicoes de "fonte").</summary>
    private int ContarFontes() => ServidorAscDesktop.LerCatalogo(ArquivoAllFonts).Count;

    /// <summary>Catalogo desta maquina: a pasta e' a mesma que o resto do produto usa
    /// (<see cref="ConfiguracaoMotor.PastaFontes"/>, por maquina quando gravavel, senao por usuario), e as
    /// pastas de fontes sao as quatro que importam no Windows:
    /// <list type="bullet">
    /// <item>fontes do sistema (<c>C:\Windows\Fonts</c>);</item>
    /// <item>fontes instaladas "so para mim" (<c>%LOCALAPPDATA%\Microsoft\Windows\Fonts</c>);</item>
    /// <item>fontes da nuvem do Microsoft 365 (<c>%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts</c>) — a
    ///   Task 7 mostrou que sem elas uma PPTX com tema Aptos simplesmente nao abria;</item>
    /// <item>as fontes que o motor traz (<c>&lt;PastaMotor&gt;\fonts</c>: Open Sans, Caladea, Carlito, Asana;
    ///   <c>ASC.ttf</c> excluída do pacote, S3) — o equivalente do <c>core-fonts</c> que o app oficial instala
    ///   em Program Files.</item>
    /// </list>
    /// </summary>
    public static CatalogoFontes Padrao(ConfiguracaoMotor motor, Action<string>? registrar = null)
    {
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return new CatalogoFontes(
            motor.PastaFontes,
            Path.Combine(motor.PastaMotor, "converter", "allfontsgen.exe"),
            [
                Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                Path.Combine(local, "Microsoft", "Windows", "Fonts"),
                Path.Combine(local, "Microsoft", "FontCache", "4", "CloudFonts"),
                Path.Combine(motor.PastaMotor, "fonts"),
            ],
            registrar);
    }

    // -------------------------------------------------------------------------------------------------
    // O editor recebe o AllFonts.js pelo host (CatalogoNoEditor, fix round 1 da Task 1 do Plano 5), direto do
    // catalogo — nada mais e' copiado para dentro do motor (que, instalado, fica em Program Files, so' leitura).
    // -------------------------------------------------------------------------------------------------

    /// <summary>Confere, ANTES da navegacao, que o <c>AllFonts.js</c> que o host vai servir ao editor
    /// (<see cref="CatalogoNoEditor"/>) existe no catalogo (<see cref="ConfiguracaoMotor.PastaFontes"/>).
    /// Falha FECHADO (item 2 da revisao 1 do Plano 1): sem o arquivo o modulo "allfonts" do requirejs morre e o
    /// usuario veria "Carregando documento" para sempre — quem chama para' e mostra <see cref="ResultadoCatalogoEditor.Mensagem"/>
    /// (pt-BR) em vez de navegar. Nao escreve nada, em lugar nenhum. Nunca lanca.</summary>
    public static ResultadoCatalogoEditor ConferirParaOEditor(ConfiguracaoMotor motor)
    {
        var arquivo = Path.Combine(motor.PastaFontes, CatalogoNoEditor.NomeArquivo);
        try
        {
            if (File.Exists(arquivo))
                return new ResultadoCatalogoEditor(true, $"AllFonts.js servido do catalogo pelo host ({new FileInfo(arquivo).Length} bytes, {arquivo})");
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            return new ResultadoCatalogoEditor(false, $"Não foi possível ler o catálogo de fontes ({arquivo}: {e.Message}) — o editor não vai carregar.");
        }
        return new ResultadoCatalogoEditor(false, $"AllFonts.js AUSENTE em {motor.PastaFontes} — o editor nao vai carregar");
    }
}

/// <summary><see cref="Erro"/> em pt-BR, nunca excecao para o chamador.</summary>
public sealed record ResultadoCatalogo(bool Regenerou, int Fontes, TimeSpan Duracao, string? Erro);

/// <summary>Resultado de <see cref="CatalogoFontes.ConferirParaOEditor"/>: <see cref="Mensagem"/> em pt-BR
/// sempre (para log); <see cref="Ok"/> diz se quem chamou PODE navegar (fail-closed).</summary>
public readonly record struct ResultadoCatalogoEditor(bool Ok, string Mensagem);
