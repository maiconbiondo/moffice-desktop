using System.Text.Json;

namespace mOffice.Nucleo.Tests;

public class RecentesParaOEditorTests
{
    [Fact] // formato que o web-apps le' (Desktop.js, S(t)): id, type = c_oAscFileType, path com extensao
    public void Json_FormatoDoWebApps()
    {
        using var d = JsonDocument.Parse(RecentesParaOEditor.Json([@"C:\x\peticao.docx", @"\\srv\p\custos.xlsx"]));
        var a = d.RootElement[0];
        Assert.Equal((0, 65, @"C:\x\peticao.docx"), (a.GetProperty("id").GetInt32(), a.GetProperty("type").GetInt32(), a.GetProperty("path").GetString()));
        Assert.Equal(257, d.RootElement[1].GetProperty("type").GetInt32());
    }

    [Fact] // extensao que o mOffice nao conhece fica de fora (o web-apps a descartaria de qualquer jeito)
    public void Json_IgnoraDesconhecidos() =>
        Assert.Equal(0, JsonDocument.Parse(RecentesParaOEditor.Json([@"C:\x\a.pages"])).RootElement.GetArrayLength());

    [Fact] // o web-apps passa o path por $("<div>").html(path).text(): "&" de um nome de arquivo vai codificado (sem isso "a&amp;b" viraria "a&b")
    public void Json_CodificaHtmlDoCaminho()
    {
        using var d = JsonDocument.Parse(RecentesParaOEditor.Json([@"C:\x\a&amp;b.docx"]));
        Assert.Equal(@"C:\x\a&amp;amp;b.docx", d.RootElement[0].GetProperty("path").GetString());
    }
}
