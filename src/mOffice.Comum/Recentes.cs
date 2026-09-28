using System.Text.Json;
using mOffice.Nucleo;

namespace mOffice.Comum;

public sealed record ItemRecente(string Caminho, bool Fixado, DateTimeOffset Quando);

/// <summary>recentes-&lt;app&gt;.json: ate 10 itens, fixados primeiro e nunca cortados. Sem cache (le a cada chamada).</summary>
public sealed class Recentes
{
    public const int Maximo = 10;
    private readonly string _arquivo;
    public Recentes(string arquivo) { _arquivo = arquivo; Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!); }
    public static Recentes DoApp(string app) => new(Pastas.Recentes(app));

    public IReadOnlyList<ItemRecente> Listar() => Ordenar(Ler());

    public void Registrar(string caminho)
    {
        var lista = Ler();
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
        Gravar(ordenada);
    }

    public void Fixar(string caminho, bool fixado)
        => Gravar(Ler().Select(i => Igual(i.Caminho, caminho) ? i with { Fixado = fixado } : i).ToList());

    public void Remover(string caminho) { var l = Ler(); l.RemoveAll(i => Igual(i.Caminho, caminho)); Gravar(l); }

    private static bool Igual(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    private static List<ItemRecente> Ordenar(IEnumerable<ItemRecente> l) => l.OrderByDescending(i => i.Fixado).ThenByDescending(i => i.Quando).ToList();
    private List<ItemRecente> Ler()
    {
        try { return File.Exists(_arquivo) ? JsonSerializer.Deserialize<List<ItemRecente>>(File.ReadAllText(_arquivo)) ?? [] : []; }
        catch (JsonException) { return []; }
        catch (IOException) { return []; }
    }
    private void Gravar(List<ItemRecente> l) => File.WriteAllText(_arquivo, JsonSerializer.Serialize(l));
}
