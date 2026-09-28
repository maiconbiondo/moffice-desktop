using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class TemaEditorTests
{
    [Fact] public void Ids_DoWebApps() { Assert.Equal("theme-dark", TemaEditor.Escuro.Id()); Assert.Equal("theme-classic-light", TemaEditor.Claro.Id()); }
    [Fact] public void Tipos_DoWebApps() { Assert.Equal("dark", TemaEditor.Escuro.Tipo()); Assert.Equal("light", TemaEditor.Claro.Tipo()); }

    [Fact] // a URL de entrada carrega o tema (uitheme) — montada por uma funcao pura, testavel sem janela
    public void UrlDeEntrada_CarregaUitheme()
    {
        var url = JanelaEditor.MontarUrlDeEntrada(mOffice.Conversor.FormatoDocumento.Docx, "peticao.docx", TemaEditor.Escuro);
        Assert.Contains("uitheme=theme-dark", url);
        Assert.Contains("doctype=word", url);
        Assert.Contains("lang=pt-BR", url);
    }

    /// <summary>Classe do &lt;body&gt; do QUADRO do editor. O tema e' aplicado por
    /// <c>web-apps/apps/documenteditor/main/app.js</c>, que roda DENTRO do iframe — o documento de cima (onde
    /// o <c>ExecutarJsAsync</c> executa) nunca ganha classe de tema nenhuma, entao ler
    /// <c>document.body.className</c> dali seria uma asserção vazia (achado crítico 1 da revisão 1).</summary>
    private const string ClasseDoBodyDoEditor = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w || !w.document || !w.document.body) return "sem quadro";
            return String(w.document.body.className);
          } catch (e) { return "erro: " + e; }
        })()
        """;

    /// <summary>Versao booleana da anterior, para o <see cref="EditorReal.Esperar"/> (que compara por
    /// igualdade): a classe do body traz outros tokens junto com a do tema.</summary>
    private const string TemaAplicadoNoEditor = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w || !w.document || !w.document.body) return "sem quadro";
            return String(String(w.document.body.className).indexOf("theme-dark") >= 0);
          } catch (e) { return "erro: " + e; }
        })()
        """;

    /// <summary><c>window.uitheme</c> do quadro: e' o que o <c>index.html</c> monta a partir do
    /// <c>uitheme=</c> da URL / do <c>RendererProcessVariable</c> e o que o app le' para escolher o tema.</summary>
    private const string UiThemeDoEditor = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w) return "sem quadro";
            return String(w.uitheme && (w.uitheme.id || w.uitheme));
          } catch (e) { return "erro: " + e; }
        })()
        """;

    [SkippableFact] // com Tema = Escuro o editor carrega theme-dark: o body DO QUADRO ganha a classe do tema
    public void TemaEscuro_ChegaAoEditor()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { Tema = TemaEditor.Escuro }); j.Show();
            await j.AbrirAsync(doc);
            var pronto = await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150);
            Assert.Equal("true", pronto);
            // O que realmente prova o tema: a classe no body DO QUADRO do editor.
            var aplicado = await EditorReal.Esperar(j, TemaAplicadoNoEditor, "true", 50);
            var classe = await j.ExecutarJsAsync(ClasseDoBodyDoEditor);
            Assert.True(aplicado == "true", $"body do quadro sem theme-dark; className = {classe}");
            // Secundario: o uitheme que o bootstrap do editor montou (do uitheme= da URL / do
            // RendererProcessVariable que a ponte injeta).
            var uitheme = await j.ExecutarJsAsync(UiThemeDoEditor);
            Assert.Contains("theme-dark", uitheme);
            j.Close();
        });
    }
}
