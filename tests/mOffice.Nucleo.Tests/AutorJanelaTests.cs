using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 4 (SO-03/CE-1/MvP-01) na janela real: comentario e alteracao controlada levam o nome do autor da
/// janela (e nao "mOffice"), no editor e no .docx gravado.</summary>
[Collection(JanelaReal.Nome)]
public class AutorJanelaTests
{
    private const string Nome = "Dra. Teste";

    /// <summary>Liga o controle de alteracoes, escreve e comenta — pelas APIs publicas do motor, do mesmo jeito que o web-apps
    /// (o comentario leva o usuario de <c>appOptions.user</c>, que veio da URL de entrada: <c>userid</c>/<c>username</c>).</summary>
    private static readonly string EscreverEComentar = EditorReal.NoEditor("""
        var u = w.DE.getController('Main').appOptions.user;
        api.asc_SetTrackRevisions(true);
        if (typeof api.asc_RemoveSelection === "function") api.asc_RemoveSelection();
        api.asc_AddText("texto revisado ");
        var d = new w.Asc.asc_CCommentDataWord(null);
        d.asc_putText("comentario de teste"); d.asc_putTime(String(new Date().getTime())); d.asc_putOnlyOfficeTime(String(new Date().getTime()));
        d.asc_putUserId(u.id); d.asc_putUserName(u.fullname); d.asc_putSolved(false); d.asc_putQuoteText("x");
        api.asc_addComment(d);
        return "ok";
        """);

    /// <summary>Autores lidos DO EDITOR: o do comentario (pluginMethod_GetAllComments) e o da alteracao (asc_GetRevisionsChangesStack,
    /// que so' enxerga o que esta selecionado — por isso o select all). Devolve "comentario|revisao" ou "" enquanto nao ha' os dois.</summary>
    private static readonly string AutoresNoEditor = EditorReal.NoEditor("""
        var cs = api.pluginMethod_GetAllComments();
        api.asc_EditSelectAll();
        var rev = api.asc_GetRevisionsChangesStack();
        api.asc_RemoveSelection();
        if (!cs || !cs.length || !rev || !rev.length) return "";
        return cs[0].Data.UserName + "|" + rev[0].get_UserName();
        """);

    [SkippableFact] // comentario e alteracao controlada levam o nome do autor da janela; persistem ao salvar e reabrir
    public void Comentario_ERevisao_ComONomeDoAutor()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        mOffice.Conversor.Tests.MotorDeTeste.Localizar();
        var pasta = Directory.CreateTempSubdirectory("moffice-autor-").FullName;
        var destino = Path.Combine(pasta, "autor.docx");
        string noEditor = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { Autor = new Autor("uteste", Nome) });
                j.Show();
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EscreverEComentar)).Trim('"'));
                for (var i = 0; i < 50 && noEditor == ""; i++)
                {
                    await Task.Delay(200);
                    var r = (await j.ExecutarJsAsync(AutoresNoEditor)).Trim('"');
                    if (r.Contains('|') && !r.StartsWith("erro")) noEditor = r;
                }
                await j.SalvarComoAsync(destino, FormatoDocumento.Docx);
                j.Close();
            }, TimeSpan.FromSeconds(180));

            Assert.Equal($"{Nome}|{Nome}", noEditor);
            // Lido do ZIP cru (e nao pelo SDK tipado): o motor grava o comentario em word/comments.xml com um tipo de relacao
            // que o SDK nao reconhece como "comentarios" — o que importa aqui e' o atributo w:author nos dois lugares.
            using var zip = System.IO.Compression.ZipFile.OpenRead(destino);
            System.Xml.Linq.XDocument Ler(string parte) { using var s = zip.GetEntry(parte)!.Open(); return System.Xml.Linq.XDocument.Load(s); }
            System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var autoresDosComentarios = Ler(zip.Entries.Single(e => e.FullName.StartsWith("word/comments", StringComparison.Ordinal) && !e.FullName.Contains("Ext") && !e.FullName.Contains("Ids") && e.FullName.EndsWith(".xml")).FullName).Descendants(w + "comment").Select(c => (string?)c.Attribute(w + "author")).ToList();
            Assert.Equal([Nome], autoresDosComentarios);
            var autoresDasInsercoes = Ler("word/document.xml").Descendants(w + "ins").Select(c => (string?)c.Attribute(w + "author")).ToList();
            Assert.NotEmpty(autoresDasInsercoes);
            Assert.All(autoresDasInsercoes, a => Assert.Equal(Nome, a));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
