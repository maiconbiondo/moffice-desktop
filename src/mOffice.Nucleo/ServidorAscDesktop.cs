namespace mOffice.Nucleo;

/// <summary>
/// Servidor do esquema nativo <c>ascdesktop://fonts/&lt;caminho absoluto&gt;</c> — o mesmo que o build desktop
/// do ONLYOFFICE pede (medido na sonda da Task 4: a planilha pediu
/// <c>ascdesktop://fonts/C:/WINDOWS/Fonts/arial.ttf</c>). Aqui ele so RESOLVE o caminho; quem responde e' a
/// <see cref="GuardaEgresso"/> (um unico ponto de resposta para o WebResourceRequested).
///
/// Contencao, em duas listas (o caminho pedido vira caminho absoluto normalizado com
/// <see cref="Path.GetFullPath(string)"/> antes de qualquer comparacao):
/// 1. RAIZES — servido se estiver DENTRO de uma raiz permitida, com o separador final incluido na comparacao,
///    senao uma pasta irma de mesmo prefixo (<c>...\fontes-fora</c> para a raiz <c>...\fontes</c>) passaria.
/// 2. CATALOGO — a lista EXATA de arquivos que o <c>AllFonts.js</c> nomeia (<c>window["__fonts_files"]</c>).
///    Ate a Task 2 do Plano 1 esse arquivo vinha do ONLYOFFICE OFICIAL instalado na maquina; desde entao e' o
///    mOffice quem o gera, com o <c>allfontsgen.exe</c> que ele embarca, em <see cref="mOffice.Conversor.
///    ConfiguracaoMotor.PastaFontes"/> (ver <see cref="CatalogoFontes"/>) — a raiz servida mudou de dono, mas
///    a whitelist por ARQUIVO continua igual. Foi a Task 7 que mostrou por que ela e' necessaria: as
///    fontes do Microsoft 365 moram em <c>%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\&lt;familia&gt;</c> e
///    as do proprio ONLYOFFICE instalado em <c>C:\Program Files\ONLYOFFICE\DesktopEditors\fonts</c> — nenhuma
///    das duas cai nas raizes fixas, e por isso uma PPTX com tema Aptos NAO ABRIA: o sdkjs recebia null no
///    lugar do arquivo da fonte e estourava dentro do desenho (<c>Cannot read properties of null</c>), sem
///    erro na tela — o editor so ficava em "Carregando documento" para sempre. Whitelist por ARQUIVO, nao por
///    raiz: e' mais estreita (nada alem do que o catalogo nomeia) e sai da configuracao real da maquina em vez
///    de uma lista de pastas adivinhada.
/// Alem das duas listas, valem em TODA resolucao (nos dois ramos): so' caminho LOCAL (UNC recusado — ver
/// <c>EhLocal</c>) e so' as extensoes de <c>ExtensoesServiveis</c>. Sem esse filtro, a raiz do motor
/// (que tambem guarda <c>x2t.exe</c> e as DLLs do core) e a pasta de recovery do documento entregavam
/// qualquer arquivo seu a' pagina.
/// Qualquer outra coisa devolve <c>null</c> (a guarda responde 403).
/// </summary>
public sealed class ServidorAscDesktop
{
    public const string PrefixoFontes = "ascdesktop://fonts/";
    private readonly string[] _raizesFixas;
    private readonly HashSet<string> _doCatalogo;

    /// <summary>Raiz que muda a cada documento (pasta de recovery). Delegate porque a sessao nasce depois do servidor.</summary>
    public Func<string?>? RaizExtra { get; set; }

    public ServidorAscDesktop(IEnumerable<string> raizes, IEnumerable<string>? arquivosDoCatalogo = null)
    {
        _raizesFixas = raizes.Where(r => !string.IsNullOrWhiteSpace(r)).Select(Normalizar).ToArray();
        _doCatalogo = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var a in arquivosDoCatalogo ?? [])
        {
            try { _doCatalogo.Add(Path.GetFullPath(a)); }
            catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { }
        }
    }

    /// <summary>Quantos arquivos o catalogo trouxe (para log/diagnostico).</summary>
    public int ArquivosDoCatalogo => _doCatalogo.Count;

    /// <summary>Extensoes que o catalogo pode autorizar. MEDIDO no <c>AllFonts.js</c> real desta maquina: 630
    /// entradas, 595 <c>.ttf</c> + 18 <c>.otf</c> + 17 <c>.ttc</c> — nada mais. O filtro nao perde nenhuma fonte
    /// de verdade e impede que uma entrada de outro tipo (um <c>.dll</c>, um <c>.exe</c>) vire arquivo servido
    /// por um catalogo adulterado: o AllFonts.js e' um arquivo gravado por OUTRO programa, em pasta gravavel
    /// pelo usuario, e a whitelist e' a unica coisa entre ele e uma leitura de arquivo arbitrario.</summary>
    private static readonly string[] ExtensoesDeFonte = [".ttf", ".otf", ".ttc"];

    /// <summary>Extensoes que este servidor pode devolver, aplicadas na RESOLUCAO — nos dois ramos (raiz E
    /// catalogo), e nao so' na leitura do catalogo (achado da revisao final 2026-09-17).
    ///
    /// Por que na resolucao tambem: o ramo das RAIZES nao tinha filtro nenhum. Uma das raizes e'
    /// <c>PastaMotor\fonts</c>, e a pasta do motor tambem guarda <c>x2t.exe</c> e as DLLs do core; a raiz
    /// dinamica e' a pasta de recovery do documento (<c>Editor.bin</c>, <c>changes/</c>). Sem filtro, qualquer
    /// arquivo dentro de uma raiz era servido a' pagina — inclusive executavel. O filtro nao muda nada do que
    /// o editor realmente pede.
    ///
    /// MEDIDO nos logs de egresso commitados desta maquina (4.429 <c>.ttf</c> + 311 <c>.png</c> + 3 <c>.ttc</c>,
    /// nenhuma outra extensao): alem das tres de fonte, o editor pede <c>fonts_thumbnail*.png</c> da pasta do
    /// catalogo (medido quando essa pasta ainda era o cache do ONLYOFFICE oficial; desde a Task 2 do Plano 1
    /// e' <see cref="mOffice.Conversor.ConfiguracaoMotor.PastaFontes"/>, o catalogo que o proprio mOffice gera
    /// — mesmo papel de raiz, dono diferente). <c>.bin</c>
    /// NAO entra (achado 1 da revisao final, fix review 2): a mesma medicao mostra que <c>ascdesktop://</c>
    /// nunca foi pedido para um <c>.bin</c> nesta rodada — e a raiz DINAMICA e' justamente a pasta de recovery
    /// do documento, que guarda <c>Editor.bin</c> (o documento inteiro do usuario, em claro). Admitir a
    /// extensao aqui reabriria, pelo esquema de fontes, a mesma leitura arbitraria que este filtro existe para
    /// fechar — quem le/grava o Editor.bin e' <c>PonteDesktop</c>/<c>X2tConversor</c>, nunca a guarda de
    /// egresso. O <c>.js</c> NAO entra: o <c>AllFonts.js</c> chega ao editor pelo host do catalogo
    /// (<c>CatalogoNoEditor</c>, uma URL exata, desde o fix round 1 da Task 1 do Plano 5), nunca por este esquema — deixa-lo
    /// passar seria abrir leitura de script arbitrario sem ganho nenhum.</summary>
    private static readonly string[] ExtensoesServiveis = [".ttf", ".otf", ".ttc", ".png"];

    /// <summary>Raizes do app: fontes do Windows, fontes do motor copiado e o catalogo do mOffice
    /// (<see cref="mOffice.Conversor.ConfiguracaoMotor.PastaFontes"/>, gerado pelo proprio mOffice desde a
    /// Task 2 do Plano 1 — antes era o cache do ONLYOFFICE oficial instalado — de onde vem o
    /// <c>fonts_thumbnail*.png</c>).</summary>
    public static ServidorAscDesktop Padrao(mOffice.Conversor.ConfiguracaoMotor motor, Action<string>? registrar = null) => new(
    [
        Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
        Path.Combine(motor.PastaMotor, "fonts"),
        motor.PastaFontes,
    ], LerCatalogo(Path.Combine(motor.PastaFontes, "AllFonts.js"), registrar));

    /// <summary>Le os caminhos de <c>window["__fonts_files"] = [ "...", ... ]</c> do AllFonts.js. Nunca lanca:
    /// sem catalogo o servidor continua valendo so pelas raizes (e a janela ja loga o catalogo ausente).
    ///
    /// Falha FECHADO (achado 1 da revisao 2026-09-17, fix review 1): se o terminador <c>];</c> do array nao
    /// aparecer — arquivo truncado, ou lido no meio da regravacao pelo app oficial — a varredura NAO continua
    /// ate o fim do arquivo. Antes ela continuava, e entao qualquer caminho absoluto que estivesse em qualquer
    /// lugar do arquivo (inclusive fora do array de fontes) entrava na whitelist. O catalogo mora em pasta
    /// gravavel pelo usuario e e' escrito por outro programa: o modo de falha certo e' "nenhuma fonte extra
    /// autorizada" (o servidor volta a valer so pelas raizes fixas), nunca "autoriza o arquivo inteiro".</summary>
    /// <param name="registrar">Recebe uma linha de diagnostico quando o catalogo esta malformado. Opcional
    /// porque o parser tambem e' chamado fora da janela (testes, ferramentas).</param>
    public static IReadOnlyList<string> LerCatalogo(string allFontsJs, Action<string>? registrar = null)
    {
        try
        {
            if (!File.Exists(allFontsJs)) return [];
            var texto = File.ReadAllText(allFontsJs);
            var inicio = texto.IndexOf("__fonts_files", StringComparison.Ordinal);
            if (inicio < 0) return [];
            var abre = texto.IndexOf('[', inicio);
            if (abre < 0) return [];
            // O fim do array NAO e' o proximo ']': ha nomes de fonte com colchete no nome — nesta maquina,
            // `UbuntuMono[wght].ttf`. Procurar ']' cru cortava a lista em 47 de 630 arquivos (MEDIDO), e o
            // efeito disso e' invisivel ate um documento pedir uma fonte da parte perdida. O fim e' o
            // fechamento do COMANDO: `]` seguido de `;`.
            var fim = System.Text.RegularExpressions.Regex.Match(texto[abre..], @"\]\s*;");
            if (!fim.Success)
            {
                registrar?.Invoke("catálogo de fontes malformado — nenhuma fonte extra autorizada");
                return [];
            }
            var fecha = abre + fim.Index;
            // Aspas escapadas (`\"`) nao precisam de tratamento: `"` e' caractere PROIBIDO em nome de arquivo
            // no Windows, entao nenhum caminho real pode conte-lo. Virgula, sim, e' permitida — e o `[^"]+`
            // ja a preserva, porque a delimitacao e' pelas aspas, nao pela virgula (caso coberto por teste).
            return System.Text.RegularExpressions.Regex
                .Matches(texto[abre..fecha], "\"([^\"\r\n]+)\"")
                .Select(m => m.Groups[1].Value)
                .Where(Path.IsPathFullyQualified)
                .Where(EhLocal)
                .Where(c => ExtensoesDeFonte.Contains(Path.GetExtension(c), StringComparer.OrdinalIgnoreCase))
                .ToArray();
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return []; }
    }

    /// <summary>Caminho de arquivo a servir, ou <c>null</c> se a URI nao e' deste servidor ou nao passa na contencao.</summary>
    public string? ResolverArquivo(string uri)
    {
        if (!uri.StartsWith(PrefixoFontes, StringComparison.OrdinalIgnoreCase)) return null;
        var bruto = uri[PrefixoFontes.Length..];
        var corte = bruto.IndexOfAny(['?', '#']);
        if (corte >= 0) bruto = bruto[..corte];
        try { bruto = Uri.UnescapeDataString(bruto); } catch (UriFormatException) { return null; }
        bruto = bruto.Replace('/', Path.DirectorySeparatorChar);
        if (bruto.Length == 0 || !Path.IsPathFullyQualified(bruto)) return null;
        // ANTES do GetFullPath: uma forma UNC ja escrita assim nao pode nem chegar a ser normalizada.
        if (!EhLocal(bruto)) return null;
        string cheio;
        try { cheio = Path.GetFullPath(bruto); } catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return null; }
        // DEPOIS do GetFullPath tambem: a normalizacao pode trocar a grafia (\\?\UNC\ <-> \\), e a decisao tem
        // de valer para o caminho que sera' realmente aberto.
        if (!EhLocal(cheio)) return null;
        if (!ExtensoesServiveis.Contains(Path.GetExtension(cheio), StringComparer.OrdinalIgnoreCase)) return null;
        if (!_doCatalogo.Contains(cheio) && !Raizes().Any(r => cheio.StartsWith(r, StringComparison.OrdinalIgnoreCase))) return null;
        return File.Exists(cheio) ? cheio : null;
    }

    /// <summary>Caminho de disco LOCAL (nao UNC). Recusar UNC nao e' redundante com a comparacao de raiz: o
    /// ramo do CATALOGO nao compara com raiz nenhuma, entao uma entrada <c>\\servidor\share\x.ttf</c> num
    /// AllFonts.js adulterado chegaria ao <see cref="File.Exists(string)"/> — e um <c>File.Exists</c> num host
    /// inexistente fica SEGUNDOS no timeout de SMB, na thread da interface (medido na revisao final: 5,4 s).
    /// Ou seja: nao e' so' leitura de arquivo remoto, e' travar a janela e vazar uma tentativa de conexao pela
    /// rede a partir de um arquivo que outro programa escreve. Por isso a checagem existe nos DOIS lugares —
    /// <see cref="LerCatalogo"/> (a entrada nem entra na whitelist) e aqui (o pedido nem e' resolvido).
    ///
    /// As duas grafias Win32 de UNC: <c>\\servidor\share\...</c> e a estendida <c>\\?\UNC\servidor\share\...</c>
    /// (idem <c>\\.\UNC\</c>). Um <c>\\?\C:\...</c> continua local e passa.</summary>
    private static bool EhLocal(string caminho)
    {
        var p = caminho.Replace('/', '\\');
        if (p.StartsWith(@"\\?\", StringComparison.Ordinal) || p.StartsWith(@"\\.\", StringComparison.Ordinal))
            return !p[4..].StartsWith(@"UNC\", StringComparison.OrdinalIgnoreCase);
        return !p.StartsWith(@"\\", StringComparison.Ordinal);
    }

    public IEnumerable<string> Raizes()
    {
        foreach (var r in _raizesFixas) yield return r;
        var extra = RaizExtra?.Invoke();
        if (!string.IsNullOrWhiteSpace(extra)) yield return Normalizar(extra);
    }

    // O separador final faz parte da contencao: sem ele "...\fontes-fora\x.ttf" comeca com "...\fontes".
    private static string Normalizar(string p) => Path.TrimEndingDirectorySeparator(Path.GetFullPath(p)) + Path.DirectorySeparatorChar;
}
