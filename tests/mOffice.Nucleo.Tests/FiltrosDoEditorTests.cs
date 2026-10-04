namespace mOffice.Nucleo.Tests;

public class FiltrosDoEditorTests
{
    [Theory] // cada filtro do motor ganha titulo, filtro e consentimento proprios (MvP-15/PL-08)
    [InlineData("images", "Inserir imagem", Consentimento.Imagem)]
    [InlineData("word", "Selecionar documento", Consentimento.Conversao)]
    [InlineData("cell", "Selecionar planilha", Consentimento.Conversao)]
    [InlineData("(*.xml)", "Importar XML", Consentimento.Conversao)]
    [InlineData("csv/txt", "Importar texto ou CSV", Consentimento.Leitura)]
    [InlineData("any", "Selecionar arquivo", Consentimento.Nenhum)]
    public void Para_Conhecidos(string filtro, string titulo, Consentimento c)
    {
        var p = FiltrosDoEditor.Para(filtro)!;
        Assert.Equal((titulo, c), (p.Titulo, p.Consentimento));
    }

    [Theory] // midia nao tem suporte (IsSupportMedia = false): nenhum dialogo
    [InlineData("audio")] [InlineData("video")] [InlineData("inventado")]
    public void Para_SemSuporte_Nulo(string filtro) => Assert.Null(FiltrosDoEditor.Para(filtro));

    [Fact] // "any" (hiperlink para arquivo) nao filtra por tipo
    public void Any_TodosOsArquivos() => Assert.Equal("Todos os arquivos|*.*", FiltrosDoEditor.Para("any")!.Filtro);
}
