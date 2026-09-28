using System.Text.Json.Nodes;
using mOffice.Conversor;
using mOffice.Nucleo;
using mOffice.Nucleo.Ortografia;

namespace mOffice.Nucleo.Tests;

public class OrtografiaTests
{
    private static DicionarioUsuario Usuario() =>
        new(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "dic.txt"));

    // Corretor falso: so' "errada" e' incorreta; sugestao fixa.
    private static bool Correta(string p, int l) => p != "errada";
    private static IReadOnlyList<string> Sugestoes(string p, int l) => p == "errada" ? ["certa", "acertada"] : [];

    [Fact] // spell: ecoa TODOS os campos do pedido e acrescenta usrCorrect por palavra e usrSuggest vazio
    public void Spell_EcoaCamposEMarcaPorPalavra()
    {
        var r = PedidoOrtografia.Responder(
            """{"type":"spell","ParagraphId":"662","RecalcId":7,"ElementId":0,"usrWords":["boa","errada"],"usrLang":[1046,1046],"cellsInfo":[{"a":1}]}""",
            Correta, Sugestoes, Usuario());
        var o = JsonNode.Parse(r!)!.AsObject();
        Assert.Equal("662", (string?)o["ParagraphId"]); Assert.Equal(7, (int)o["RecalcId"]!); Assert.NotNull(o["cellsInfo"]);
        Assert.Equal([true, false], o["usrCorrect"]!.AsArray().Select(n => (bool)n!));
        Assert.Empty(o["usrSuggest"]!.AsArray());
    }

    [Fact] // suggest: usrCorrect vazio e uma lista de sugestoes por palavra
    public void Suggest_DevolveListaPorPalavra()
    {
        var r = PedidoOrtografia.Responder("""{"type":"suggest","ParagraphId":"1","RecalcId":1,"ElementId":3,"usrWords":["errada"],"usrLang":[1046]}""",
            Correta, Sugestoes, Usuario());
        var o = JsonNode.Parse(r!)!.AsObject();
        Assert.Empty(o["usrCorrect"]!.AsArray());
        Assert.Equal(["certa", "acertada"], o["usrSuggest"]![0]!.AsArray().Select(n => (string)n!));
        Assert.Equal(3, (int)o["ElementId"]!);
    }

    [Fact] // tamanhos diferentes de palavras e idiomas: pedido descartado (o oficial tambem descarta)
    public void Spell_TamanhosDiferentes_NaoResponde() =>
        Assert.Null(PedidoOrtografia.Responder("""{"type":"spell","usrWords":["a","b"],"usrLang":[1046]}""", Correta, Sugestoes, Usuario()));

    [Fact] // palavra no dicionario do usuario e' correta mesmo que o corretor diga que nao
    public void Spell_PalavraDoUsuario_Correta()
    {
        var u = Usuario(); u.Adicionar("errada");
        var r = PedidoOrtografia.Responder("""{"type":"spell","usrWords":["errada"],"usrLang":[1046]}""", Correta, Sugestoes, u);
        Assert.Equal([true], JsonNode.Parse(r!)!["usrCorrect"]!.AsArray().Select(n => (bool)n!));
    }

    [Fact] // add: grava a palavra (mesmo com aspas no meio, que o editor concatena sem escape) e responde {"type":"add"}
    public void Add_GravaERespondeTipo()
    {
        var u = Usuario();
        Assert.Equal("""{"type":"add"}""", PedidoOrtografia.Responder("""{"type":"add","usrWords":["Zorvelquim"]}""", Correta, Sugestoes, u));
        Assert.True(u.Contem("Zorvelquim"));
        Assert.Equal("""{"type":"add"}""", PedidoOrtografia.Responder("{\"type\":\"add\",\"usrWords\":[\"d\"Avila\"]}", Correta, Sugestoes, u));
        Assert.True(u.Contem("d\"Avila"));
    }

    [Fact] // clear (JSON) limpa o dicionario do usuario; a string "clear" (reinicio) NAO e' tratada aqui (vai direto no JS)
    public void Clear_LimpaDicionario()
    {
        var u = Usuario(); u.Adicionar("x");
        Assert.Equal("""{"type":"clear"}""", PedidoOrtografia.Responder("""{"type":"clear"}""", Correta, Sugestoes, u));
        Assert.False(u.Contem("x"));
        Assert.Null(PedidoOrtografia.Responder("clear", Correta, Sugestoes, u));
    }

    [Fact] // o dicionario do usuario sobrevive a outra instancia (persistido em disco, uma palavra por linha)
    public void DicionarioUsuario_Persiste()
    {
        var arquivo = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "dic.txt");
        new DicionarioUsuario(arquivo).Adicionar("Beltrão");
        Assert.True(new DicionarioUsuario(arquivo).Contem("Beltrão"));
    }

    [Fact] // idioma sem dicionario e' sempre correto e sem sugestoes (o editor manda todas as palavras de todos os idiomas)
    public void Corretor_IdiomaSemDicionario_Correta()
    {
        var c = new Corretor(Path.Combine(Path.GetTempPath(), "mOffice-testes", "nao-existe"));
        Assert.True(c.Correta("qualquer", 2070)); Assert.Empty(c.Sugestoes("qualquer", 2070));
        Assert.Equal(new Dictionary<int, string> { [1046] = "pt_BR", [1033] = "en_US" }, Corretor.Dicionarios);
    }

    [Fact] // Pastas.DicionarioUsuario mora na raiz de dados por usuario (spec §8), compartilhado pelos tres apps
    public void Pastas_DicionarioUsuario_NaRaiz() =>
        Assert.Equal(Path.Combine(Pastas.Raiz, "dicionario-usuario.txt"), Pastas.DicionarioUsuario);

    [SkippableFact] // dicionarios reais do motor: pt-BR marca erro e sugere; en-US tambem
    public void Corretor_DicionariosReais()
    {
        var pasta = Path.Combine(mOffice.Conversor.Tests.MotorDeTeste.Localizar().PastaMotor, "dictionaries");
        Skip.If(!File.Exists(Path.Combine(pasta, "pt_BR", "pt_BR.dic")), "dicionarios do motor nao obtidos (rode tools/onlyoffice/obter-motor.ps1)");
        var c = new Corretor(pasta);
        Assert.True(c.Correta("advogado", 1046));
        Assert.False(c.Correta("advogaddo", 1046));
        Assert.Contains("advogado", c.Sugestoes("advogaddo", 1046));
        Assert.True(c.Correta("lawyer", 1033));
        Assert.False(c.Correta("lawyerr", 1033));
    }
}

[Collection(JanelaReal.Nome)]
public class OrtografiaJanelaTests
{
    [SkippableFact] // documento novo pt-BR: digitar uma palavra errada faz o editor receber usrCorrect=false
    public void DocumentoNovo_PalavraErrada_RecebeIncorreta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            await j.NovoAsync(FormatoDocumento.Docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord("O advogaddo assinou "));
            var r = await EditorReal.Esperar(j, "String((window.__moffice.ultimaOrtografia||'').indexOf('false')>=0)", "true", 150);
            Assert.Equal("true", r);
            j.Close();
        }, TimeSpan.FromSeconds(180));
    }
}
