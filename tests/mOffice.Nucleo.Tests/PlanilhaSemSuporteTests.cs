using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Revisao final do Plano 4 — com <c>isOffline = true</c> (Task 2b) a planilha passou a oferecer dois caminhos do app oficial
/// que chamam a ponte SEM testar se a funcao existe: "Mover ou copiar planilha" -> "(Criar nova planilha)"
/// (<c>asc_moveWorksheet/asc_copyWorksheet(..., [])</c> -> <c>AscDesktopEditor.OpenWorkbook</c>) e "Vinculos externos" ->
/// "Abrir origem" (<c>asc_openExternalReference(ref)</c> -> <c>AscDesktopEditor.openExternalReference</c>). Sem a funcao:
/// TypeError -> asc_onError CRITICO e o documento travado. Aqui o api e' dirigido como o web-apps faz, e a prova e': nenhum
/// erro de JS, nenhum asc_onError critico, a edicao continua funcionando, e a faixa pt-BR no OpenWorkbook.
/// </summary>
[Collection(JanelaReal.Nome)]
public class PlanilhaSemSuporteTests
{
    private readonly ITestOutputHelper _saida;
    public PlanilhaSemSuporteTests(ITestOutputHelper saida) { _saida = saida; }

    /// <summary>Registra os asc_onError do editor em <c>w.__errosApi</c> ("id:nivel"; nivel -1 = Critical).</summary>
    private static readonly string EscutarErros = EditorReal.NoEditor("""
            w.__errosApi = [];
            api.asc_registerCallback("asc_onError", function (id, nivel) { w.__errosApi.push(id + ":" + nivel); });
            return "ok";
        """);

    private static readonly string LerErros = EditorReal.NoEditor("""
            return JSON.stringify({ api: w.__errosApi || [], js: (w.__moffice && w.__moffice.errosJs) || [] });
        """);

    /// <summary>O "ok" do dialogo "Mover ou copiar" com "(Criar nova planilha)" (web-apps spreadsheeteditor app.js:
    /// <c>"new" === o ? a = [] ...; i ? api.asc_copyWorksheet(pos, nomes, folhas, a) : api.asc_moveWorksheet(pos, folhas, nomes, a)</c>).</summary>
    private static string MoverOuCopiarParaNova(bool copiar) => EditorReal.NoEditor($$"""
            var folhas = [api.asc_getActiveWorksheetIndex()];
            var nomes = [api.asc_getWorksheetName(folhas[0])];
            if ({{(copiar ? "true" : "false")}}) api.asc_copyWorksheet(0, nomes, folhas, []);
            else api.asc_moveWorksheet(0, folhas, nomes, []);
            return "ok " + api.asc_getWorksheetsCount();
        """);

    [SkippableFact]
    public void MoverOuCopiarParaPastaNova_E_AbrirOrigemDeVinculo_NaoTravamOEditor()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-plan-semsuporte-" + Guid.NewGuid().ToString("N"))).FullName;
        var xlsx = FabricaDocumentos.Xlsx(pasta, "Base");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha");
                var avisos = new List<string>();
                j.AvisoAoUsuario += m => avisos.Add(m);
                j.Show();
                await j.AbrirAsync(xlsx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EscutarErros)).Trim('"'));

                // 1) Mover e copiar para "(Criar nova planilha)": OpenWorkbook -> faixa, e a pasta de trabalho fica como estava.
                Assert.Equal("ok 1", (await j.ExecutarJsAsync(MoverOuCopiarParaNova(copiar: false))).Trim('"'));
                for (var i = 0; i < 50 && avisos.Count < 1; i++) await Task.Delay(100);
                Assert.Equal("ok 1", (await j.ExecutarJsAsync(MoverOuCopiarParaNova(copiar: true))).Trim('"'));
                for (var i = 0; i < 50 && avisos.Count < 2; i++) await Task.Delay(100);
                Assert.Equal([JanelaEditor.MensagemNovaPastaDeTrabalho, JanelaEditor.MensagemNovaPastaDeTrabalho], avisos);

                // 1b) "Dados > Do texto/CSV > De arquivo" (web-apps onDataFromText("file"): api.asc_TextFromFileOrUrl(opcoes, cb)):
                //     o dialogo e' o "csv/txt" do host (Plano 6, Task 8), nunca o de imagem; cancelado, nenhuma faixa. O caso com
                //     arquivo escolhido esta' em PonteArquivosDoEditorJanelaTests.
                avisos.Clear();
                j.EscolherImagemParaTeste = () => throw new InvalidOperationException("nao devia abrir o dialogo de imagem");
                var filtros = new List<string>();
                j.EscolherArquivoParaTeste = (f, _) => { filtros.Add(f); return null; };
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        api.asc_TextFromFileOrUrl(new w.Asc.asc_CTextOptions(), function () { });
                        return "ok";
                    """))).Trim('"'));
                for (var i = 0; i < 50 && filtros.Count < 1; i++) await Task.Delay(100);
                Assert.Equal(["csv/txt"], filtros);
                Assert.Empty(avisos);
                j.EscolherImagemParaTeste = null;
                j.EscolherArquivoParaTeste = null;
                avisos.Clear();
                // Fix 1 da revisao da Task 8: o caminho do vinculo vem do documento -> o host pergunta antes de qualquer E/S; o usuario nega aqui.
                var perguntas = new List<string>();
                j.ConfirmarVinculoParaTeste = m => { perguntas.Add(m); return false; };

                // 2) Vinculo externo: uma formula que aponta para outra pasta de trabalho cria a referencia; "Abrir origem"
                //    (web-apps openLink: api.asc_openExternalReference(ref)) vai a' ponte, que responde "nao abriu" (cb(0)).
                var vinculo = (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        api.asc_findCell("D1");
                        api.pluginMethod_InputText("='C:\\moffice-inexistente\\[outra.xlsx]Planilha1'!A1");
                        api.asc_closeCellEditor();
                        return "ok";
                    """))).Trim('"');
                Assert.Equal("ok", vinculo);
                var refs = await EditorReal.Esperar(j, EditorReal.NoEditor("""
                        var r = api.asc_getExternalReferences(); return String(r ? r.length : -1);
                    """), "1", 50);
                _saida.WriteLine("vinculos externos: " + refs);
                Assert.Equal("1", refs);
                Assert.Equal("null", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var r = api.asc_getExternalReferences()[0];
                        return String(api.asc_openExternalReference(r));
                    """))).Trim('"'));
                await Task.Delay(500);
                _saida.WriteLine("perguntas do vinculo: " + string.Join(" | ", perguntas));
                Assert.Single(perguntas);
                Assert.Contains(@"C:\moffice-inexistente\outra.xlsx", perguntas[0]);

                var erros =(await j.ExecutarJsAsync(LerErros)).Trim('"').Replace("\\\"", "\"");
                _saida.WriteLine("erros: " + erros);
                Assert.Contains("\"js\":[]", erros);                  // nenhum TypeError
                Assert.DoesNotContain(":-1", erros);                  // nenhum asc_onError critico
                Assert.Matches("\"api\":\\[\"-?\\d+:0\"\\]", erros);  // o aviso NAO critico do proprio sdkjs ("nao abriu")

                // O aviso NAO critico e' uma janela modal do web-apps (Common.UI.warning): le' o texto e fecha no botao,
                // como o usuario faria (enquanto ela esta' aberta o editor nao recebe teclado — e' o normal).
                var modal = (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var d = w.document, jan = d.querySelectorAll(".asc-window.modal-dlg, .asc-window.alert");
                        var abertas = [].filter.call(jan, function (e) { return e.offsetParent !== null || w.getComputedStyle(e).display !== "none"; });
                        if (!abertas.length) return "sem modal";
                        var m = abertas[abertas.length - 1];
                        var texto = (m.querySelector(".info-box, .body, .box") || m).textContent.replace(/\s+/g, " ").trim();
                        var b = m.querySelector(".footer button, button.primary, button");
                        if (b) b.click();
                        return "modal: " + texto;
                    """))).Trim('"');
                _saida.WriteLine(modal);
                Assert.StartsWith("modal:", modal);   // o aviso "nao abriu" do sdkjs, fechado no OK

                // 3) A edicao continua funcionando.
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoCell("depois-do-semsuporte", "B2"))).Trim('"'));
                Assert.Equal("depois-do-semsuporte", await EditorReal.Esperar(j,
                    EditorReal.NoEditor("api.asc_findCell('B2'); return String(api.asc_getCellInfo().asc_getText());"), "depois-do-semsuporte", 50));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                Assert.True((await j.SalvarAsync()).Gravou);
                Assert.Null(j.UltimoErro);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            _saida.WriteLine("xlsx salvo: " + FabricaDocumentos.TextoXlsx(xlsx));
            Assert.Contains("depois-do-semsuporte", FabricaDocumentos.TextoXlsx(xlsx));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
