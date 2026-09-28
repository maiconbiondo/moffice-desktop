using mOffice.Conversor;
using WeCantSpell.Hunspell;

namespace mOffice.Nucleo.Ortografia;

/// <summary>
/// Corretor ortografico local (Task 1 do Plano 3) sobre os dicionarios Hunspell que o proprio motor traz
/// (<c>third_party/onlyoffice/dictionaries/&lt;nome&gt;/&lt;nome&gt;.dic</c> + <c>.aff</c>), lidos pelo
/// WeCantSpell.Hunspell (porte .NET puro: sem codigo nativo, sem rede).
///
/// So' pt-BR e en-US (spec §2/§8: os unicos dicionarios que o pacote do mOffice distribui). Qualquer outro
/// idioma — e um dicionario ausente ou ilegivel — conta como "sem dicionario": toda palavra e' correta e nao
/// ha' sugestao, que e' tambem o que o app oficial faz (o editor manda palavras de TODOS os idiomas do texto).
///
/// Carga preguicosa, uma vez por LCID, thread-safe (<see cref="Lazy{T}"/> com ExecutionAndPublication): o
/// pt_BR tem ~318 mil entradas e leva ~0,7 s para carregar (medido na sondagem da Task 1). Por isso tambem a
/// instancia e' UMA por processo (<see cref="Compartilhado"/>) — nao uma por janela.
/// </summary>
public sealed class Corretor
{
    /// <summary>LCID (o <c>usrLang</c> do editor) -&gt; pasta/arquivo do dicionario.</summary>
    public static readonly IReadOnlyDictionary<int, string> Dicionarios = new Dictionary<int, string>
    {
        [1046] = "pt_BR",
        [1033] = "en_US",
    };

    private readonly Dictionary<int, Lazy<WordList?>> _listas;
    private readonly Action<string>? _log;

    public Corretor(string pastaDicionarios, Action<string>? log = null)
    {
        _log = log;
        _listas = Dicionarios.ToDictionary(
            d => d.Key,
            d => new Lazy<WordList?>(() => Carregar(pastaDicionarios, d.Value), LazyThreadSafetyMode.ExecutionAndPublication));
    }

    public static Corretor Padrao(ConfiguracaoMotor motor, Action<string>? log = null) =>
        new(Path.Combine(motor.PastaMotor, "dictionaries"), log);

    private static readonly object _travaCompartilhado = new();
    private static Corretor? _compartilhado;

    /// <summary>A instancia do PROCESSO. O primeiro chamador decide a pasta do motor (todas as janelas de um
    /// processo usam o mesmo motor; a costura de teste com motor temporario liga <c>dictionaries</c> por juncao
    /// ao motor real, entao os dicionarios sao os mesmos). O log do primeiro chamador e' o que registra falha
    /// de carga.</summary>
    public static Corretor Compartilhado(ConfiguracaoMotor motor, Action<string>? log = null)
    {
        lock (_travaCompartilhado) return _compartilhado ??= Padrao(motor, log);
    }

    private WordList? Carregar(string pasta, string nome)
    {
        var dic = Path.Combine(pasta, nome, nome + ".dic");
        try
        {
            if (!File.Exists(dic) || !File.Exists(Path.ChangeExtension(dic, ".aff")))
            {
                _log?.Invoke($"ORTOGRAFIA sem dicionario {nome} em {Path.Combine(pasta, nome)} (idioma tratado como sem corretor)");
                return null;
            }
            var inicio = System.Diagnostics.Stopwatch.StartNew();
            var lista = WordList.CreateFromFiles(dic);
            _log?.Invoke($"ORTOGRAFIA dicionario {nome} carregado em {inicio.ElapsedMilliseconds} ms");
            return lista;
        }
        catch (Exception ex)
        {
            _log?.Invoke($"ORTOGRAFIA falha ao carregar {nome}: {ex.Message}");
            return null;
        }
    }

    private WordList? Lista(int lcid) => _listas.TryGetValue(lcid, out var l) ? l.Value : null;

    public bool Correta(string palavra, int lcid)
    {
        if (string.IsNullOrEmpty(palavra)) return true;
        var lista = Lista(lcid);
        return lista is null || lista.Check(palavra);
    }

    public IReadOnlyList<string> Sugestoes(string palavra, int lcid)
    {
        if (string.IsNullOrEmpty(palavra)) return [];
        var lista = Lista(lcid);
        return lista is null ? [] : lista.Suggest(palavra).ToList();
    }

    /// <summary>Carrega os dicionarios fora da thread da interface (a janela dispara isto antes de navegar,
    /// para o primeiro SpellCheck do editor nao pagar a carga). Nunca lanca: falha de carga vira log e idioma
    /// sem corretor (ver <see cref="Carregar"/>).</summary>
    public Task PreCarregarAsync() => Task.Run(() => { foreach (var l in _listas.Values) _ = l.Value; });
}
