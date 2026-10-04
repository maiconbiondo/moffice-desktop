using mOffice.Nucleo.Recuperacao;

namespace mOffice.Nucleo.Tests;

public class DescartadosTests
{
    private static string Pasta() { var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }

    [Fact] // guardar move a pasta inteira para os descartados (o original some do lugar)
    public void Guardar_Move()
    {
        var rec = Pasta(); File.WriteAllText(Path.Combine(rec, "Editor.bin"), "x");
        var raiz = Pasta();
        var nova = Descartados.Guardar(rec, raiz)!;
        Assert.False(Directory.Exists(rec));
        Assert.True(File.Exists(Path.Combine(nova, "Editor.bin")));
    }

    [Fact] // varredura: so' o que passou de 4 dias sai
    public void Varrer_SoAsVelhas()
    {
        var raiz = Pasta();
        var velha = Directory.CreateDirectory(Path.Combine(raiz, "velha")).FullName; File.WriteAllText(Path.Combine(velha, "a"), "x");
        File.SetLastWriteTimeUtc(Path.Combine(velha, "a"), DateTime.UtcNow.AddDays(-5)); Directory.SetLastWriteTimeUtc(velha, DateTime.UtcNow.AddDays(-5));
        var nova = Directory.CreateDirectory(Path.Combine(raiz, "nova")).FullName; File.WriteAllText(Path.Combine(nova, "a"), "x");
        Assert.Equal(1, Descartados.Varrer(raiz, Descartados.Retencao));
        Assert.False(Directory.Exists(velha)); Assert.True(Directory.Exists(nova));
    }

    [Fact] // a retencao conta do DESCARTE, nao da ultima edicao: um documento editado ha' 6 dias e descartado agora fica os 4 dias
    public void Guardar_RetencaoContaDoDescarte()
    {
        var rec = Pasta(); var arq = Path.Combine(rec, "Editor.bin"); File.WriteAllText(arq, "x");
        File.SetLastWriteTimeUtc(arq, DateTime.UtcNow.AddDays(-6)); Directory.SetLastWriteTimeUtc(rec, DateTime.UtcNow.AddDays(-6));
        var raiz = Pasta();
        var nova = Descartados.Guardar(rec, raiz)!;
        Assert.Equal(0, Descartados.Varrer(raiz, Descartados.Retencao));
        Assert.True(Directory.Exists(nova));
    }

    [Fact] // pasta velha mas com a trava presa (recuperada e aberta numa janela) nunca e' varrida
    public void Varrer_NaoTocaPastaComTravaPresa()
    {
        var raiz = Pasta();
        var viva = Directory.CreateDirectory(Path.Combine(raiz, "viva")).FullName; File.WriteAllText(Path.Combine(viva, "a"), "x");
        using var trava = TravaSessao.Adquirir(viva)!;
        foreach (var f in Directory.GetFiles(viva)) if (!f.EndsWith(TravaSessao.NomeArquivo)) File.SetLastWriteTimeUtc(f, DateTime.UtcNow.AddDays(-9));
        File.SetLastWriteTimeUtc(Path.Combine(viva, TravaSessao.NomeArquivo), DateTime.UtcNow.AddDays(-9)); Directory.SetLastWriteTimeUtc(viva, DateTime.UtcNow.AddDays(-9));
        Assert.Equal(0, Descartados.Varrer(raiz, Descartados.Retencao));
        Assert.True(Directory.Exists(viva)); Assert.True(File.Exists(Path.Combine(viva, "a")), "o conteudo da pasta com a trava presa nao pode ser apagado");
    }

    [Fact] // guardar de uma pasta que nao existe / destino com o mesmo nome: nunca lanca, nunca sobrescreve
    public void Guardar_AusenteEColisao()
    {
        var raiz = Pasta();
        Assert.Null(Descartados.Guardar(Path.Combine(Pasta(), "nao-existe"), raiz));
        var rec = Pasta(); File.WriteAllText(Path.Combine(rec, "a"), "1");
        var ocupado = Directory.CreateDirectory(Path.Combine(raiz, Path.GetFileName(rec))).FullName; File.WriteAllText(Path.Combine(ocupado, "b"), "2");
        var nova = Descartados.Guardar(rec, raiz)!;
        Assert.NotEqual(ocupado, nova);
        Assert.True(File.Exists(Path.Combine(ocupado, "b"))); Assert.True(File.Exists(Path.Combine(nova, "a")));
    }

    [Fact] // a pasta dos descartados fica ao lado da recuperacao, na raiz de dados (mesma unidade: o Move e' um rename)
    public void Pastas_Descartados_EmRaizDeDados() =>
        Assert.Equal(Path.Combine(@"C:\dados", "descartados"), Pastas.DescartadosEm(@"C:\dados"));

    [Fact] // reabrir dos nao salvos: a pasta volta para a recuperacao SEM a marca e sem lixo; o original some dos descartados
    public void Repor_MoveDeVolta_SemMarca()
    {
        var rec = Pasta(); File.WriteAllText(Path.Combine(rec, "Editor.bin"), "x");
        var desc = Pasta(); var rec2 = Pasta();
        var guardada = Descartados.Guardar(rec, desc)!;
        Assert.True(File.Exists(Path.Combine(guardada, Descartados.NomeMarca)));
        var nova = Descartados.Repor(guardada, rec2)!;
        Assert.False(Directory.Exists(guardada));
        Assert.True(File.Exists(Path.Combine(nova, "Editor.bin"))); Assert.False(File.Exists(Path.Combine(nova, Descartados.NomeMarca)));
        Assert.Equal(rec2, Path.GetDirectoryName(nova), ignoreCase: true);
    }

    [Fact] // o move falha (alguem segura um arquivo): abre uma COPIA na recuperacao e deixa o descartado intocado (nunca perde)
    public void Repor_MoveFalha_CopiaEDeixaOriginal()
    {
        var desc = Pasta(); var pasta = Directory.CreateDirectory(Path.Combine(desc, "x1")).FullName;
        File.WriteAllText(Path.Combine(pasta, "Editor.bin"), "dados"); Directory.CreateDirectory(Path.Combine(pasta, "changes")); File.WriteAllText(Path.Combine(pasta, "changes", "changes0.json"), "\"c\",");
        File.WriteAllText(Path.Combine(pasta, Descartados.NomeMarca), "agora");
        var rec2 = Pasta();
        using var seguro = new FileStream(Path.Combine(pasta, "changes", "changes0.json"), FileMode.Open, FileAccess.Read, FileShare.Read);  // impede o rename da pasta
        var nova = Descartados.Repor(pasta, rec2)!;
        Assert.NotEqual(pasta, nova, StringComparer.OrdinalIgnoreCase);
        Assert.Equal("dados", File.ReadAllText(Path.Combine(nova, "Editor.bin"))); Assert.False(File.Exists(Path.Combine(nova, Descartados.NomeMarca)));
        Assert.Equal("dados", File.ReadAllText(Path.Combine(pasta, "Editor.bin")));     // original intacto
        Assert.True(File.Exists(Path.Combine(pasta, Descartados.NomeMarca)));
    }

    [Fact] // so' repoe o que esta' DENTRO da raiz dos descartados (qualquer outra pasta passa direto)
    public void EstaNosDescartados_SoDentroDaRaiz()
    {
        var raiz = Pasta();
        Assert.True(Descartados.Contem(raiz, Path.Combine(raiz, "abc")));
        Assert.False(Descartados.Contem(raiz, Path.Combine(Pasta(), "abc")));
        Assert.False(Descartados.Contem(raiz, raiz + "-outra\\abc"));
    }
}
