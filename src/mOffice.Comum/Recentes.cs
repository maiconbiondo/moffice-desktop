using System.Text.Json;
using mOffice.Nucleo;

namespace mOffice.Comum;

public sealed record ItemRecente(string Caminho, bool Fixado, DateTimeOffset Quando);

/// <summary>recentes-&lt;app&gt;.json: ate 25 itens, fixados primeiro e nunca cortados. Sem cache (le a cada chamada).</summary>
public sealed class Recentes
{
    public const int Maximo = 25;
    private readonly string _arquivo;
    public Recentes(string arquivo) { _arquivo = arquivo; Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!); }
    public static Recentes DoApp(string app) => new(Pastas.Recentes(app));

    /// <summary>Linha de log opcional (a casca liga ao log do processo).</summary>
    public static Action<string>? Log { get; set; }

    private static void Logar(string linha) { try { Log?.Invoke(linha); } catch (Exception) { } }

    public IReadOnlyList<ItemRecente> Listar() => Ordenar(Ler().Lista);

    public void Registrar(string caminho) => Alterar("registrar", lista =>
    {
        var existente = lista.FirstOrDefault(i => Igual(i.Caminho, caminho));
        lista.RemoveAll(i => Igual(i.Caminho, caminho));
        lista.Insert(0, new ItemRecente(caminho, existente?.Fixado ?? false, DateTimeOffset.Now));
        var ordenada = Ordenar(lista).ToList();
        while (ordenada.Count > Maximo)
        {
            var vitima = ordenada.LastOrDefault(i => !i.Fixado);
            if (vitima is null) break;
            ordenada.Remove(vitima);
        }
        return ordenada;
    });

    public void Fixar(string caminho, bool fixado)
        => Alterar("fixar", l => l.Select(i => Igual(i.Caminho, caminho) ? i with { Fixado = fixado } : i).ToList());

    /// <summary>So' o usuario tira um recente da lista (o "Tirar da lista" / "Remover da lista"): nada aqui, nem a casca, remove sozinho um arquivo que "sumiu" — um UNC fora do ar volta amanha.</summary>
    public void Remover(string caminho) => Alterar("remover", l => { l.RemoveAll(i => Igual(i.Caminho, caminho)); return l; });

    /// <summary>
    /// Leitura-modificacao-escrita sob o mutex. Nunca lanca (I2 da revisao final do Plano 6: chamado de eventos da interface). Arquivo existente que
    /// nao le (M1: lixo, vazio, preso) NAO e' sobrescrito — antes, a leitura devolvia [] e o proximo Registrar gravava a lista de um item por cima,
    /// perdendo os fixados; agora a mudanca e' descartada (log) e o arquivo fica como estava, como no config.json.
    /// </summary>
    private void Alterar(string acao, Func<List<ItemRecente>, List<ItemRecente>> mudanca)
    {
        try
        {
            ArquivoAtomico.ComExclusao(_arquivo, () =>
            {
                var (lista, legivel) = Ler();
                if (!legivel) { Logar("RECENTES " + Path.GetFileName(_arquivo) + " ilegivel: " + acao + " nao gravado (o arquivo nao foi alterado)"); return 0; }
                ArquivoAtomico.Escrever(_arquivo, JsonSerializer.Serialize(mudanca(lista)));
                return 0;
            });
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            Logar("RECENTES " + acao + " nao gravado: " + e.Message);
        }
    }

    private static bool Igual(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    private static List<ItemRecente> Ordenar(IEnumerable<ItemRecente> l) => l.OrderByDescending(i => i.Fixado).ThenByDescending(i => i.Quando).ToList();

    /// <summary>(lista, legivel): ausente = vazia e legivel; existente que nao le (JSON invalido, vazio, <c>null</c>, E/S) = vazia e <c>false</c>.</summary>
    private (List<ItemRecente> Lista, bool Legivel) Ler()
    {
        try
        {
            if (!File.Exists(_arquivo)) return ([], true);
            var l = JsonSerializer.Deserialize<List<ItemRecente>>(ArquivoAtomico.Ler(_arquivo));
            return l is null ? ([], false) : (l.Where(i => i is not null && i.Caminho is not null).ToList(), true);
        }
        catch (JsonException) { return ([], false); }
        catch (IOException) { return ([], false); }
        catch (UnauthorizedAccessException) { return ([], false); }
    }
}
