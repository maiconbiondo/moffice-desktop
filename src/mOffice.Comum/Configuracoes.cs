using System.Text.Json;
using mOffice.Nucleo;

namespace mOffice.Comum;

/// <summary>config.json por usuario. Le a cada get e grava a cada set (como AppConfig do mPDF): sem cache, uma
/// mudanca feita com o app aberto vale na hora. Campo ausente = padrao do record (sem migracao).</summary>
public sealed class Configuracoes
{
    // Padrao Claro (o "Clássico claro" do editor), escolha do usuario em 2026-09-25: a aparencia do
    // ONLYOFFICE que o escritorio ja' conhece, independente do tema do Windows. "Seguir o Windows" continua
    // disponivel em Configuracoes.
    private sealed record Dados(ModoTema Tema = ModoTema.Claro, bool Grampo = false);
    private readonly string _arquivo;

    public Configuracoes(string arquivo) { _arquivo = arquivo; Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!); }
    public static Configuracoes Padrao() => new(Pastas.Config);

    public ModoTema Tema { get => Ler().Tema; set => Gravar(Ler() with { Tema = value }); }
    public bool Grampo { get => Ler().Grampo; set => Gravar(Ler() with { Grampo = value }); }

    private Dados Ler()
    {
        try { return File.Exists(_arquivo) ? JsonSerializer.Deserialize<Dados>(File.ReadAllText(_arquivo)) ?? new Dados() : new Dados(); }
        catch (JsonException) { return new Dados(); }
        catch (IOException) { return new Dados(); }
    }
    private void Gravar(Dados d) => File.WriteAllText(_arquivo, JsonSerializer.Serialize(d));
}
