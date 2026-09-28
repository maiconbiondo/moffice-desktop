using System.Text.Json;
using mOffice.Comum;

namespace mOffice.Comum.Tests;

public class RecentesTests
{
    private static Recentes Novo() => new(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "recentes-mtexto.json"));

    [Fact] // registrar deduplica sem diferenciar maiusculas e poe o mais recente primeiro
    public void Registrar_DeduplicaEOrdena()
    {
        var r = Novo();
        r.Registrar(@"C:\a\um.docx"); r.Registrar(@"C:\a\dois.docx"); r.Registrar(@"C:\A\UM.DOCX");
        var lista = r.Listar();
        Assert.Equal(2, lista.Count);
        Assert.Equal(@"C:\A\UM.DOCX", lista[0].Caminho);
    }

    [Fact] // corta em 10, mas nunca remove um fixado
    public void Registrar_CortaEm10_PreservandoFixados()
    {
        var r = Novo();
        r.Registrar(@"C:\fixo.docx"); r.Fixar(@"C:\fixo.docx", true);
        for (var i = 0; i < 12; i++) r.Registrar($@"C:\d{i}.docx");
        var lista = r.Listar();
        Assert.Equal(Recentes.Maximo, lista.Count);
        Assert.True(lista[0].Fixado); Assert.Equal(@"C:\fixo.docx", lista[0].Caminho);
        Assert.DoesNotContain(lista, i => i.Caminho == @"C:\d0.docx");
    }

    [Fact] // fixados vem primeiro; registrar de novo preserva Fixado
    public void Fixados_Primeiro_EPreservadosAoRegistrar()
    {
        var r = Novo();
        r.Registrar(@"C:\a.docx"); r.Registrar(@"C:\b.docx"); r.Fixar(@"C:\a.docx", true); r.Registrar(@"C:\a.docx");
        Assert.Equal(@"C:\a.docx", r.Listar()[0].Caminho); Assert.True(r.Listar()[0].Fixado);
        r.Fixar(@"C:\a.docx", false);
        Assert.False(r.Listar().Single(i => i.Caminho == @"C:\a.docx").Fixado);
    }

    [Fact] // guarda "nunca corta um fixado" so e' alcancavel com MAIS de 10 fixados NO ARQUIVO -- via
           // Registrar/Fixar isso nunca acontece sozinho: um caminho NOVO sempre entra nao-fixado
           // (Fixar so' roda DEPOIS que Registrar retorna), entao o proprio corte de Registrar remove
           // o item recem-inserido antes que ele possa virar fixado, sempre que ja' existem 10
           // fixados -- o estado ">10 fixados" so' chega ao Recentes por uma fonte EXTERNA ao objeto
           // (edicao manual do json, migracao de uma versao com Maximo diferente, etc.), entao
           // este teste semeia o arquivo direto, como essa fonte externa faria. Com 11 fixados no
           // arquivo + registra 1 nao-fixado: o loop so acha o nao-fixado pra cortar
           // (LastOrDefault(i => !i.Fixado)), remove so ele, e para (a proxima volta nao acha mais
           // nenhum nao-fixado) -- a lista final fica com os 11 fixados, ACIMA do Maximo (10).
    public void Registrar_MaisDe10Fixados_NuncaCorta()
    {
        var arquivo = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "recentes-mtexto.json");
        Directory.CreateDirectory(Path.GetDirectoryName(arquivo)!);
        var fixados = Enumerable.Range(0, 11)
            .Select(i => new ItemRecente($@"C:\f{i}.docx", Fixado: true, DateTimeOffset.Now.AddMinutes(-i)))
            .ToList();
        File.WriteAllText(arquivo, JsonSerializer.Serialize(fixados));

        var r = new Recentes(arquivo);
        r.Registrar(@"C:\naofixo.docx");

        var lista = r.Listar();
        Assert.Equal(11, lista.Count);
        Assert.All(lista, i => Assert.True(i.Fixado));
        Assert.DoesNotContain(lista, i => i.Caminho == @"C:\naofixo.docx");
    }

    [Fact] public void Remover_Apaga() { var r = Novo(); r.Registrar(@"C:\a.docx"); r.Remover(@"c:\A.DOCX"); Assert.Empty(r.Listar()); }

    [Fact] // arquivo corrompido = lista vazia, sem lancar
    public void Corrompido_ListaVazia()
    {
        var a = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "r.json");
        Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, "[{");
        Assert.Empty(new Recentes(a).Listar());
    }
}
