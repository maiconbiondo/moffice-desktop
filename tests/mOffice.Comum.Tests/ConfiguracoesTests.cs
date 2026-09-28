using mOffice.Comum;

namespace mOffice.Comum.Tests;

public class ConfiguracoesTests
{
    private static string Arquivo() => Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "config.json");

    [Fact] public void Padroes_ClaroESemGrampo() { var c = new Configuracoes(Arquivo()); Assert.Equal(ModoTema.Claro, c.Tema); Assert.False(c.Grampo); }

    [Fact] // set grava e um segundo objeto le o mesmo arquivo
    public void Set_PersisteNoDisco()
    {
        var a = Arquivo(); var c = new Configuracoes(a);
        c.Tema = ModoTema.Escuro; c.Grampo = true;
        var d = new Configuracoes(a);
        Assert.Equal(ModoTema.Escuro, d.Tema); Assert.True(d.Grampo);
    }

    [Fact] // arquivo corrompido = padroes, sem lancar; o proximo set sobrescreve
    public void Corrompido_UsaPadroes()
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, "{nao e json");
        var c = new Configuracoes(a);
        Assert.Equal(ModoTema.Claro, c.Tema);
        c.Tema = ModoTema.Escuro;
        Assert.Equal(ModoTema.Escuro, new Configuracoes(a).Tema);
    }

    [Fact] // config antiga sem o campo Grampo continua lendo (sem migracao)
    public void CampoAusente_CaiNoPadrao()
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, "{\"Tema\":2}");
        var c = new Configuracoes(a);
        Assert.Equal(ModoTema.Escuro, c.Tema); Assert.False(c.Grampo);
    }
}
