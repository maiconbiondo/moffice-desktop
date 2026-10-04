using System.Text.Json;

namespace mOffice.Nucleo.Tests;

public class ConfigEditorTests
{
    [Fact] // o que o host injeta: region pt-BR e o que desliga ajuda/feedback/sugestao/macros/dicas
    public void Json_TemOsCamposDecididos()
    {
        using var d = JsonDocument.Parse(ConfigEditor.Json());
        var c = d.RootElement.GetProperty("customization");
        Assert.Equal("pt-BR", d.RootElement.GetProperty("region").GetString());
        foreach (var campo in new[] { "help", "feedback", "suggestFeature", "macros" }) Assert.False(c.GetProperty(campo).GetBoolean(), campo);
        Assert.False(c.GetProperty("features").GetProperty("featuresTips").GetBoolean());
    }

    [Fact] // Dono de cada campo: a URL de entrada (user, lang, uiTheme, mode) NAO pode ser repetida nem sobrescrita pela config injetada
    public void Json_NaoDisputaCamposDaUrlDeEntrada()
    {
        using var d = JsonDocument.Parse(ConfigEditor.Json());
        foreach (var campo in new[] { "user", "lang", "mode", "createUrl" }) Assert.False(d.RootElement.TryGetProperty(campo, out _), campo);
        Assert.False(d.RootElement.GetProperty("customization").TryGetProperty("uiTheme", out _));
        Assert.False(d.RootElement.GetProperty("customization").TryGetProperty("goback", out _)); // goback e' da ponte (depende de &arquivo=)
    }

    [Theory] // &arquivo= acompanha a URL de entrada: 0 = documento novo (sem "Abrir Local do Arquivo"), 1 = veio de arquivo
    [InlineData(true, "&arquivo=1&")]
    [InlineData(false, "&arquivo=0&")]
    public void UrlDeEntrada_LevaArquivo(bool temArquivo, string esperado)
    {
        var url = JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento.Docx, "a.docx", TemaEditor.Claro, new Autor("u1", "A"), null, temArquivo);
        Assert.Contains(esperado, url + "&");
    }
}
