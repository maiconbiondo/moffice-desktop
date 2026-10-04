using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

public class AtalhosTests
{
    [Theory]
    [InlineData("abrir", AcaoDeAtalho.Abrir)] [InlineData("fechar", AcaoDeAtalho.Fechar)] [InlineData("salvarComo", AcaoDeAtalho.SalvarComo)]
    [InlineData("salvar", AcaoDeAtalho.Salvar)] [InlineData("proximaJanela", AcaoDeAtalho.ProximaJanela)]
    public void DoAviso_Conhecidos(string nome, AcaoDeAtalho esperado) => Assert.Equal(esperado, AtalhosDoHost.DoAviso(nome));

    [Theory] [InlineData("")] [InlineData("formatar")] [InlineData("ABRIR")]
    public void DoAviso_Desconhecido_Nulo(string nome) => Assert.Null(AtalhosDoHost.DoAviso(nome));

    [Fact] // numeracao por tipo e por processo, sem reuso (como o Office)
    public void ContadorDeNovos_PorTipo()
    {
        var a = ContadorDeNovos.Proximo("word"); var b = ContadorDeNovos.Proximo("word"); var c = ContadorDeNovos.Proximo("cell");
        Assert.Equal(a + 1, b);
        Assert.True(c >= 1);
    }

    [Theory] // titulo do novo como o Office pt-BR
    [InlineData(FormatoDocumento.Docx, 1, "Documento1")]
    [InlineData(FormatoDocumento.Xlsx, 2, "Pasta2")]
    [InlineData(FormatoDocumento.Pptx, 3, "Apresentação3")]
    public void Titulo_DoNovo(FormatoDocumento f, int n, string esperado) =>
        Assert.Equal(esperado, new SessaoDocumento { ArquivoOrigem = @"C:\r\new.x", Formato = f, PastaRecovery = @"C:\r", Novo = true, NumeroDoNovo = n }.Titulo);

    [Fact] // sem numero (sessao criada fora da janela): so' o rotulo, e a planilha passa de "Planilha" a "Pasta"
    public void Titulo_DoNovo_SemNumero() =>
        Assert.Equal("Pasta", new SessaoDocumento { ArquivoOrigem = "a.xlsx", Formato = FormatoDocumento.Xlsx, PastaRecovery = "p", Novo = true }.Titulo);

    [Fact] // PDF aberto como documento: o nome do PDF vence o numero
    public void Titulo_TituloDoNovo_VenceONumero() =>
        Assert.Equal("contrato.pdf", new SessaoDocumento { ArquivoOrigem = "a.docx", Formato = FormatoDocumento.Docx, PastaRecovery = "p", Novo = true, NumeroDoNovo = 4, TituloDoNovo = "contrato.pdf" }.Titulo);

    [Fact] // a ponte manda os 5 nomes e o host so' conhece esses (varredura do JS: nenhum nome novo sem o enum)
    public void Ponte_AvisaSomenteNomesConhecidos()
    {
        var js = File.ReadAllText(Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo", "Ponte", "ponte-desktop.js"));
        foreach (var nome in new[] { "abrir", "fechar", "salvarComo", "salvar", "proximaJanela" })
        {
            Assert.Contains("acao = \"" + nome + "\"", js);
            Assert.NotNull(AtalhosDoHost.DoAviso(nome));
        }
        Assert.Contains("avisar(\"atalho\", acao)", js);
    }
}
