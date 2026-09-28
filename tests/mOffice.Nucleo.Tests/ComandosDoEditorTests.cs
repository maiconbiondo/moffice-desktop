using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class ComandosDoEditorTests
{
    [Theory] // o web-apps manda "create:new" com o tipo, "open:recent" com JSON, "editor:event" com action
    [InlineData("create:new", "word", "create:new", "word")]
    [InlineData("editor:event", "{\"action\":\"file:open\"}", "editor:event", "{\"action\":\"file:open\"}")]
    public void ComandoDoEditor_GuardaComandoEJsonBrutos(string cmd, string json, string esperadoCmd, string esperadoJson)
    {
        var c = new ComandoDoEditor(cmd, json);
        Assert.Equal(esperadoCmd, c.Comando); Assert.Equal(esperadoJson, c.Json);
    }

    [Fact] // helper: "editor:event" com file:open / file:close
    public void ComandoDoEditor_AcaoDoEvento()
    {
        Assert.Equal("file:close", new ComandoDoEditor("editor:event", "{\"action\":\"file:close\"}").AcaoDoEvento());
        Assert.Null(new ComandoDoEditor("create:new", "word").AcaoDoEvento());
        Assert.Null(new ComandoDoEditor("editor:event", "nao-e-json").AcaoDoEvento());
    }

    [Theory] // Ler(): o {cmd,json} bruto que o ponte-desktop.js serializa em execCommand
    // O campo json NEM SEMPRE e' string: o web-apps chama execCommand("go:folder", info.data.status), e o
    // JSON.stringify do lado JS guarda ali numero/booleano/objeto/null (achado 9 da revisao 1).
    [InlineData("{\"cmd\":\"create:new\",\"json\":\"word\"}", "create:new", "word")]
    [InlineData("{\"cmd\":\"editor:event\",\"json\":\"{\\\"action\\\":\\\"file:open\\\"}\"}", "editor:event", "{\"action\":\"file:open\"}")]
    [InlineData("{\"cmd\":\"go:folder\",\"json\":3}", "go:folder", "3")]
    [InlineData("{\"cmd\":\"title:button\",\"json\":true}", "title:button", "true")]
    [InlineData("{\"cmd\":\"open:recent\",\"json\":{\"id\":7}}", "open:recent", "{\"id\":7}")]
    [InlineData("{\"cmd\":\"recent:forget\",\"json\":null}", "recent:forget", "")]
    [InlineData("{\"cmd\":\"webapps:entry\"}", "webapps:entry", "")]
    public void Ler_GuardaComandoEJsonBrutos(string param, string esperadoCmd, string esperadoJson)
    {
        var c = ComandoDoEditor.Ler(param);
        Assert.NotNull(c);
        Assert.Equal(esperadoCmd, c!.Comando); Assert.Equal(esperadoJson, c.Json);
    }

    [Theory] // param que nao e' objeto JSON nao e' comando nenhum: null, e o host registra no log
    [InlineData("nao-e-json")]
    [InlineData("")]
    [InlineData("\"so-uma-string\"")]
    [InlineData("123")]
    [InlineData("[1,2]")]
    public void Ler_ParamInvalido_DevolveNull(string param) => Assert.Null(ComandoDoEditor.Ler(param));

    [SkippableFact] // um execCommand disparado pela pagina chega como ComandoRecebido na thread de UI
    public void ExecCommand_ChegaComoComandoRecebido()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            // LISTA, nao um slot de "ultimo recebido" (achado 2 da revisao 1): o web-apps dispara sozinho
            // doc:onready, editor:onready, editor:config, webapps:entry e title:button durante a carga, entao
            // um slot assinado antes do AbrirAsync ja chega nao-nulo e o teste passaria sem provar nada.
            var recebidos = new List<ComandoDoEditor>();
            j.ComandoRecebido += c => recebidos.Add(c);
            await j.AbrirAsync(doc);
            var pronto = await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150);
            Assert.Equal("true", pronto);
            await j.ExecutarJsAsync("window.AscDesktopEditor.execCommand('create:new', 'word')");
            ComandoDoEditor? nosso = null;
            for (var i = 0; i < 50 && nosso is null; i++)
            {
                nosso = recebidos.FirstOrDefault(c => c.Comando == "create:new" && c.Json == "word");
                if (nosso is null) await Task.Delay(100);
            }
            Assert.True(nosso is not null,
                "nenhum create:new/word entre os comandos recebidos: " + string.Join(" | ", recebidos.Select(c => c.Comando + "=" + c.Json)));
            j.Close();
        });
    }
}
