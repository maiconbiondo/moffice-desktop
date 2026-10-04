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

    [Fact] // Plano 6, Task 10: o limite e' 25 e os 3 fixados MAIS ANTIGOS sobrevivem a 30 registros
    public void Maximo_25_FixadosNuncaCortados()
    {
        Assert.Equal(25, Recentes.Maximo);
        var r = Novo();
        for (var i = 0; i < 3; i++) { r.Registrar($@"C:\fixo{i}.docx"); r.Fixar($@"C:\fixo{i}.docx", true); }
        for (var i = 0; i < 27; i++) r.Registrar($@"C:\d{i}.docx");   // 30 registros no total
        var lista = r.Listar();
        Assert.Equal(25, lista.Count);
        for (var i = 0; i < 3; i++) Assert.Contains(lista, x => x.Caminho == $@"C:\fixo{i}.docx" && x.Fixado);
        Assert.Contains(lista, x => x.Caminho == @"C:\d26.docx");          // o mais novo fica
        Assert.DoesNotContain(lista, x => x.Caminho == @"C:\d0.docx");     // o mais velho nao-fixado sai
    }

    [Fact] // gravar e' atomico (Plano 6, Task 10): nenhum temporario sobra na pasta do arquivo
    public void Gravar_Atomico_SemTemporarios()
    {
        var arquivo = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "recentes-mtexto.json");
        var r = new Recentes(arquivo);
        r.Registrar(@"C:\a.docx"); r.Registrar(@"C:\b.docx"); r.Fixar(@"C:\a.docx", true);
        Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(arquivo)!, "*.tmp"));
        Assert.Equal(2, r.Listar().Count);
    }

    [Fact] // corta em Maximo, mas nunca remove um fixado
    public void Registrar_CortaNoMaximo_PreservandoFixados()
    {
        var r = Novo();
        r.Registrar(@"C:\fixo.docx"); r.Fixar(@"C:\fixo.docx", true);
        for (var i = 0; i < Recentes.Maximo + 2; i++) r.Registrar($@"C:\d{i}.docx");
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
        var fixados = Enumerable.Range(0, Recentes.Maximo + 1)
            .Select(i => new ItemRecente($@"C:\f{i}.docx", Fixado: true, DateTimeOffset.Now.AddMinutes(-i)))
            .ToList();
        File.WriteAllText(arquivo, JsonSerializer.Serialize(fixados));

        var r = new Recentes(arquivo);
        r.Registrar(@"C:\naofixo.docx");

        var lista = r.Listar();
        Assert.Equal(Recentes.Maximo + 1, lista.Count);
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

    private static string ArquivoNovo()
    {
        var a = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "recentes-mtexto.json");
        Directory.CreateDirectory(Path.GetDirectoryName(a)!);
        return a;
    }

    [Theory] // M1 da revisao final do Plano 6: recentes ILEGIVEL (lixo, 0 byte, "null") — Registrar/Fixar/Remover NAO gravam por cima (antes: a
             // leitura dava [] e o Registrar gravava uma lista de um item, perdendo os fixados); quando o arquivo volta a ser valido, os fixados estao la'
    [InlineData("[{")] [InlineData("")] [InlineData("null")] [InlineData("\0\0")]
    public void Ilegivel_NaoEhSobrescrito_FixadosPreservados(string lixo)
    {
        var a = ArquivoNovo();
        var r = new Recentes(a);
        r.Registrar(@"C:\fixado.docx"); r.Fixar(@"C:\fixado.docx", true);
        var valido = File.ReadAllBytes(a);
        File.WriteAllText(a, lixo);
        var antes = File.ReadAllBytes(a);
        r.Registrar(@"C:\novo.docx");
        r.Fixar(@"C:\novo.docx", true);
        r.Remover(@"C:\fixado.docx");
        Assert.Equal(antes, File.ReadAllBytes(a));
        File.WriteAllBytes(a, valido);                      // o arquivo volta a ser legivel (restaurado pelo usuario/backup)
        r.Registrar(@"C:\depois.docx");
        var lista = r.Listar();
        Assert.Contains(lista, i => i.Caminho == @"C:\fixado.docx" && i.Fixado);
        Assert.Contains(lista, i => i.Caminho == @"C:\depois.docx");
    }

    [Fact] // I2: recentes SOMENTE LEITURA — Registrar/Fixar/Remover nao lancam (eram chamados do DocumentoPronto e derrubavam o app)
    public void SomenteLeitura_NaoLanca()
    {
        var a = ArquivoNovo();
        var r = new Recentes(a);
        r.Registrar(@"C:\a.docx");
        var antes = File.ReadAllBytes(a);
        File.SetAttributes(a, FileAttributes.ReadOnly);
        try
        {
            r.Registrar(@"C:\b.docx");
            r.Fixar(@"C:\a.docx", true);
            r.Remover(@"C:\a.docx");
            Assert.Equal(antes, File.ReadAllBytes(a));
        }
        finally { File.SetAttributes(a, FileAttributes.Normal); }
    }
}
