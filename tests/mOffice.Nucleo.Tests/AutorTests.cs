namespace mOffice.Nucleo.Tests;

public class AutorTests
{
    [Fact] // id estavel por SID, sem revelar o SID: "u" + 16 hex minusculos
    public void IdDoSid_EstavelEOpaco()
    {
        var a = Autor.IdDoSid("S-1-5-21-111-222-333-1001");
        Assert.Equal(a, Autor.IdDoSid("S-1-5-21-111-222-333-1001"));
        Assert.Matches("^u[0-9a-f]{16}$", a);
        Assert.NotEqual(a, Autor.IdDoSid("S-1-5-21-111-222-333-1002"));
        Assert.DoesNotContain("1001", a);
    }

    [Fact] // nome configurado vence tudo (e espacos sao aparados); vazio nao vale
    public void DoWindows_NomeConfiguradoVence()
    {
        Assert.Equal("Dra. Ana Souza", Autor.DoWindows("  Dra. Ana Souza ").Nome);
        Assert.False(string.IsNullOrWhiteSpace(Autor.DoWindows("   ").Nome));
    }

    [Fact] // sem nome configurado sempre ha' um nome (Office, Windows ou usuario) e o id e' o do SID da conta
    public void DoWindows_SemConfiguracao_TemNomeEIdOpaco()
    {
        var a = Autor.DoWindows();
        Assert.False(string.IsNullOrWhiteSpace(a.Nome));
        Assert.Matches("^u[0-9a-f]{16}$", a.Id);
        Assert.Equal(a.Id, Autor.DoWindows("outro nome").Id); // o id nao depende do nome
    }

    [Fact] // a URL leva o autor escapado e nunca mais "mOffice"/1
    public void Url_LevaAutorEscapado()
    {
        var url = JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento.Docx, "a.docx", TemaEditor.Claro, new Autor("uabc", "João & Cia"));
        Assert.Contains("&userid=uabc&username=Jo%C3%A3o%20%26%20Cia", url);
        Assert.DoesNotContain("username=mOffice", url);
    }

    [Fact] // a sobrecarga de 3 argumentos (testes antigos) tambem usa o autor real, nao "mOffice"/1
    public void Url_SobrecargaAntiga_UsaAutorDoWindows()
    {
        var url = JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento.Docx, "a.docx", TemaEditor.Claro);
        Assert.DoesNotContain("userid=1&", url);
        Assert.Contains("&userid=" + Autor.DoWindows().Id + "&", url);
    }

    [Fact] // o tema escolhido dentro do editor (id do motor) vence o do mOffice na URL
    public void Url_TemaIdVenceOTema()
    {
        var url = JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento.Docx, "a.docx", TemaEditor.Claro, new Autor("u1", "A"), "theme-night");
        Assert.Contains("uitheme=theme-night", url);
    }

    [Theory] // os 7 temas do motor e o tipo de cada um; desconhecido nao e' tema
    [InlineData("theme-classic-light", "light")]
    [InlineData("theme-light", "light")]
    [InlineData("theme-white", "light")]
    [InlineData("theme-gray", "light")]
    [InlineData("theme-dark", "dark")]
    [InlineData("theme-contrast-dark", "dark")]
    [InlineData("theme-night", "dark")]
    public void TemasDoEditor_SeteDoMotor(string id, string tipo)
    {
        Assert.Contains(id, TemasDoEditor.Conhecidos);
        Assert.Equal(tipo, TemasDoEditor.Tipo(id));
        Assert.Equal(7, TemasDoEditor.Conhecidos.Count);
    }

    [Fact] // minor: sem nome configurado o autor e' calculado UMA vez por processo (GetUserNameEx/registro/SID)
    public void DoWindows_SemConfiguracao_EhCacheado()
    {
        Assert.Same(Autor.DoWindows(), Autor.DoWindows());
        Assert.Same(Autor.DoWindows(), Autor.DoWindows("   "));
        Assert.NotSame(Autor.DoWindows(), Autor.DoWindows("Outro"));
    }

    [Fact] // minor: o id do tema vai escapado na URL
    public void Url_TemaIdEscapado()
    {
        var url = JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento.Docx, "a.docx", TemaEditor.Claro, new Autor("u1", "A"), "a b&c");
        Assert.Contains("uitheme=a%20b%26c", url);
    }

    [Fact] public void TemasDoEditor_Desconhecido_NaoEhConhecido() => Assert.DoesNotContain("theme-system", TemasDoEditor.Conhecidos);
}
