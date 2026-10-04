namespace mOffice.Nucleo.Tests;

public class PoliticaDeLinkTests
{
    [Theory]
    [InlineData("https://www.tjsp.jus.br/", AcaoDeLink.Navegador)]
    [InlineData("http://exemplo.com.br/a?b=1", AcaoDeLink.Navegador)]
    [InlineData("mailto:fulano@exemplo.com.br", AcaoDeLink.Navegador)]
    [InlineData("https://helpcenter.onlyoffice.com/userguides/docs-de.aspx?lang=pt", AcaoDeLink.Descartar)]
    [InlineData("https://feedback.onlyoffice.com/", AcaoDeLink.Descartar)]
    [InlineData("https://support.onlyoffice.com", AcaoDeLink.Descartar)]
    [InlineData("https://ONLYOFFICE.com/", AcaoDeLink.Descartar)]
    [InlineData("https://support.onlyoffice.com./x", AcaoDeLink.Descartar)]
    [InlineData("https://onlyoffice.com.evil.example/", AcaoDeLink.Navegador)]
    [InlineData(@"file:///C:/Processos/peticao.pdf", AcaoDeLink.Arquivo)]
    [InlineData(@"file://servidor/pasta/contrato.docx", AcaoDeLink.Recusar)]   // UNC: nunca
    [InlineData(@"file://localhost/C:/Processos/peticao.pdf", AcaoDeLink.Arquivo)]
    [InlineData(@"file:///C:/pasta/que/nao/existe/sem-extensao", AcaoDeLink.Recusar)] // sem extensao e nao e' pasta existente
    [InlineData(@"file:///C:/Users/x/Downloads/virus.exe", AcaoDeLink.Recusar)]
    [InlineData(@"file:///C:/x/script.ps1", AcaoDeLink.Recusar)]
    [InlineData("javascript:alert(1)", AcaoDeLink.Recusar)]
    [InlineData("ftp://x/y", AcaoDeLink.Recusar)]
    [InlineData("about:blank", AcaoDeLink.Descartar)]
    [InlineData("nao e uma uri", AcaoDeLink.Recusar)]
    [InlineData("", AcaoDeLink.Recusar)]
    public void Classificar(string uri, AcaoDeLink esperado) => Assert.Equal(esperado, PoliticaDeLink.Classificar(uri));

    [Fact] // pasta que EXISTE abre (Explorer); arquivo sem extensao que existe NAO vai ao shell
    public void Classificar_SemExtensao_SoPastaExistente()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-link-").FullName;
        try
        {
            var arquivo = Path.Combine(pasta, "semtipo"); File.WriteAllText(arquivo, "x");
            Assert.Equal(AcaoDeLink.Arquivo, PoliticaDeLink.Classificar(new Uri(pasta).AbsoluteUri));
            Assert.Equal(AcaoDeLink.Recusar, PoliticaDeLink.Classificar(new Uri(arquivo).AbsoluteUri));
        }
        finally { Directory.Delete(pasta, true); }
    }

    [Theory] // M4 da revisao final do Plano 6: nome sugerido pela PAGINA no <a download> — saneado, so' extensoes da lista
    [InlineData("saida.png", "saida.png")]
    [InlineData("Relatorio.HTML", "Relatorio.html")]
    [InlineData("foto\u202Egnp.exe", null)]                   // RLO: "fotoexe.png" na tela, .exe de verdade -> recusado
    [InlineData("foto\u202Eexe.png", "fotoexe.png")]          // controle de direcao some, a extensao real (png) fica
    [InlineData("a\u200Fb\u2066c.txt", "abc.txt")]
    [InlineData("dados.csv. . .", "dados.csv")]               // pontos/espacos no fim
    [InlineData("CON.txt", "_CON.txt")]                       // nome reservado do Windows
    [InlineData("nul.docx", "_nul.docx")]
    [InlineData("com1.backup.pdf", "_com1.backup.pdf")]
    [InlineData(".png", "arquivo.png")]
    [InlineData("a<b>c.svg", "a_b_c.svg")]
    [InlineData("script.js", null)]
    [InlineData("instalar.msi", null)]
    [InlineData("atalho.lnk", null)]
    [InlineData("sem-extensao", null)]
    [InlineData("", null)]
    public void NomeDeDownloadSeguro(string sugerido, string? esperado) => Assert.Equal(esperado, JanelaEditor.NomeDeDownloadSeguro(sugerido));
}
