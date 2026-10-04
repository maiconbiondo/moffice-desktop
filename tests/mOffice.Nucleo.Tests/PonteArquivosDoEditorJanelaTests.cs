using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 6, Task 8 — recursos do motor que pedem ARQUIVO ao host, na janela real. Os dialogos sao SEMPRE pela costura
/// <see cref="JanelaEditor.EscolherArquivoParaTeste"/> (nunca o dialogo nativo do Windows); o resto e' o caminho do produto: o editor chama a ponte
/// (OpenFilenameDialog -&gt; convertFile/CompareDocumentFile/loadLocalFile), o host converte com o x2t e devolve no quadro.
/// </summary>
[Collection(JanelaReal.Nome)]
public class PonteArquivosDoEditorJanelaTests(ITestOutputHelper saida)
{
    private static string PastaNova(string nome) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-t8-" + nome + "-" + Guid.NewGuid().ToString("N"))).FullName;

    private static async Task NaJanela(string app, string arquivo, Func<JanelaEditor, Task> corpo, Action<JanelaEditor>? preparar = null,
        JanelaEditor? janela = null)
    {
        var j = janela ?? EditorReal.NovaJanela(app);
        var fechada = false;
        j.Closed += (_, _) => fechada = true;
        preparar?.Invoke(j);
        j.Show();
        try
        {
            await j.AbrirAsync(arquivo);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            await corpo(j);
        }
        finally { if (!fechada) j.Close(); }
    }

    private static async Task EsperarConversoes(JanelaEditor j, int n)
    {
        for (var i = 0; i < 300 && j.ConversoesTerminadas < n; i++) await Task.Delay(100);
        Assert.True(j.ConversoesTerminadas >= n, $"conversoes terminadas: {j.ConversoesTerminadas} (esperado {n})");
    }

    private static readonly string ErrosJs = EditorReal.NoEditor("return JSON.stringify((w.__moffice && w.__moffice.errosJs) || []);");

    private static readonly string TextoInteiro = EditorReal.NoEditor("""
            api.asc_EditSelectAll(); var t = String(api.asc_GetSelectedText()); api.asc_RemoveSelection(); return t;
        """);

    [SkippableFact] // Inserir > Texto do arquivo: um DOCX sintetico COM IMAGEM entra no documento; salvar e reabrir mantem texto e imagem
    public void InserirTextoDoArquivo_ComImagem()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("inserir");
        var docx = FabricaDocumentos.Docx(Directory.CreateDirectory(Path.Combine(pasta, "base")).FullName, "Documento base");
        var fonte = FabricaDocumentos.DocxComImagem(Directory.CreateDirectory(Path.Combine(pasta, "fonte")).FullName, "Trecho inserido do arquivo",
            FabricaDocumentos.PngSolido(64, 255, 0, 0));
        var fonteAntes = File.ReadAllBytes(fonte);
        var filtros = new List<string>();
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                // O caminho do botao Inserir > Texto do arquivo > Do arquivo (web-apps btnTextFromFile -> api.asc_insertTextFromFile();
                // sdkjs: OpenFilenameDialog("word", true) -> convertFile(caminho, 8193) -> insere o binario).
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_insertTextFromFile(); return 'ok';"))).Trim('"'));
                await EsperarConversoes(j, 1);
                var texto = "";
                for (var i = 0; i < 50 && !texto.Contains("Trecho inserido do arquivo"); i++)
                {
                    await Task.Delay(200);
                    texto = (await j.ExecutarJsAsync(TextoInteiro)).Trim('"');
                }
                saida.WriteLine("texto: " + texto);
                Assert.Contains("Trecho inserido do arquivo", texto);
                Assert.Contains("Documento base", texto);
                Assert.Equal("[]", System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(ErrosJs)));
                var media = Directory.GetFiles(Path.Combine(j.SessaoAtual!.PastaRecovery, "media")).Select(Path.GetFileName).ToList();
                saida.WriteLine("media/: " + string.Join(", ", media));
                Assert.NotEmpty(media);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                Assert.True((await j.SalvarAsync()).Gravou);
                Assert.Null(j.UltimoErro);
                // A pasta da conversao sai no close() que o sdkjs chama no fim da insercao.
                var conv = Path.Combine(j.SessaoAtual!.PastaRecovery, ConversaoParaEditor.SubpastaConvertidos);
                Assert.True(!Directory.Exists(conv) || !Directory.EnumerateFileSystemEntries(conv).Any(), "a pasta da conversao ficou");
            }, j => j.EscolherArquivoParaTeste = (f, multi) => { filtros.Add(f + (multi ? ":multi" : "")); return [fonte]; }), TimeSpan.FromSeconds(240));
            Assert.Equal(["word:multi"], filtros);
            var salvo = FabricaDocumentos.TextoDocx(docx);
            saida.WriteLine("salvo: " + salvo);
            Assert.Contains("Trecho inserido do arquivo", salvo);
            Assert.Single(FabricaDocumentos.ImagensDocx(docx));
            Assert.Equal(fonteAntes, File.ReadAllBytes(fonte));   // o arquivo escolhido e' so' lido
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    private static string ComparacaoOuCombinacao(bool combinar) => EditorReal.NoEditor($$"""
            var o = w.AscCommonWord && w.AscCommonWord.ComparisonOptions ? new w.AscCommonWord.ComparisonOptions() : undefined;
            api.{{(combinar ? "asc_MergeDocumentFile" : "asc_CompareDocumentFile")}}(o);
            return "ok";
        """);

    /// <summary>Alteracoes controladas do documento inteiro (asc_GetRevisionsChangesStack so' enxerga a selecao: por isso o select all, como em
    /// AutorJanelaTests).</summary>
    private static readonly string Revisoes = EditorReal.NoEditor("""
            api.asc_EditSelectAll(); var r = api.asc_GetRevisionsChangesStack(); api.asc_RemoveSelection(); return String(r ? r.length : -1);
        """);

    [SkippableFact] // Revisao > Comparar com documento do arquivo: o editor mostra as diferencas como alteracoes controladas
    public void Comparar_DocumentoDoArquivo_GeraRevisoes() => CompararOuCombinar(combinar: false);

    [SkippableFact] // Combinar: idem com onDocumentMerge
    public void Combinar_DocumentoDoArquivo() => CompararOuCombinar(combinar: true);

    private void CompararOuCombinar(bool combinar)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova(combinar ? "combinar" : "comparar");
        var docx = FabricaDocumentos.Docx(Directory.CreateDirectory(Path.Combine(pasta, "base")).FullName, "Texto original da peticao");
        var outro = FabricaDocumentos.Docx(Directory.CreateDirectory(Path.Combine(pasta, "outro")).FullName, "Texto revisado da peticao com acrescimo");
        var filtros = new List<string>();
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                Assert.Equal("0", (await j.ExecutarJsAsync(Revisoes)).Trim('"'));
                Assert.Equal("ok", (await j.ExecutarJsAsync(ComparacaoOuCombinacao(combinar))).Trim('"'));
                await EsperarConversoes(j, 1);
                var n = "0";
                for (var i = 0; i < 60 && n is "0" or "-1"; i++)
                {
                    await Task.Delay(200);
                    n = (await j.ExecutarJsAsync(Revisoes)).Trim('"');
                }
                saida.WriteLine((combinar ? "combinar" : "comparar") + ": revisoes = " + n + "; texto = " + (await j.ExecutarJsAsync(TextoInteiro)).Trim('"'));
                Assert.True(int.TryParse(n, out var k) && k > 0, "nenhuma alteracao controlada depois de " + (combinar ? "combinar" : "comparar") + ": " + n);
                Assert.Equal("[]", System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(ErrosJs)));
                // Nada da conversao sobra na pasta de recuperacao (o binario foi no script e a pasta e' apagada logo depois).
                var conv = Path.Combine(j.SessaoAtual!.PastaRecovery, ConversaoParaEditor.SubpastaConvertidos);
                Assert.True(!Directory.Exists(conv) || !Directory.EnumerateFileSystemEntries(conv).Any());
            }, j => j.EscolherArquivoParaTeste = (f, _) => { filtros.Add(f); return [outro]; }), TimeSpan.FromSeconds(240));
            Assert.Equal(["word"], filtros);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // convertFile com caminho NAO consentido (vindo da pagina) devolve null e nao roda o x2t
    public void ConvertFile_SemConsentimento_Nulo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("semconsent");
        var docx = FabricaDocumentos.Docx(Directory.CreateDirectory(Path.Combine(pasta, "base")).FullName, "Base");
        // Um arquivo QUE EXISTE: o consentimento e' a unica barreira entre a pagina e o x2t.
        var alvo = FabricaDocumentos.Docx(Directory.CreateDirectory(Path.Combine(pasta, "alvo")).FullName, "Segredo de outro arquivo");
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                var js = EditorReal.NoEditor("""
                        w.__cb = "nao chamado";
                        w.AscDesktopEditor.convertFile(ALVO, 8193, function (o) { w.__cb = o === null ? "null" : "objeto"; });
                        return "ok";
                    """.Replace("ALVO", System.Text.Json.JsonSerializer.Serialize(alvo)));
                Assert.Equal("ok", (await j.ExecutarJsAsync(js)).Trim('"'));
                Assert.Equal("null", await EditorReal.Esperar(j, EditorReal.NoEditor("return String(w.__cb);"), "null", 50));
                await Task.Delay(1000);
                Assert.Equal(0, j.ConversoesTerminadas);
                Assert.False(Directory.Exists(Path.Combine(j.SessaoAtual!.PastaRecovery, ConversaoParaEditor.SubpastaConvertidos)), "o x2t rodou sem consentimento");
                // Idem comparar/combinar e loadLocalFile com o mesmo caminho.
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        w.AscDesktopEditor.CompareDocumentFile(ALVO, {}); w.AscDesktopEditor.MergeDocumentFile(ALVO, {});
                        w.__lido = "nao chamado"; w.AscDesktopEditor.loadLocalFile(ALVO, function (u) { w.__lido = u === null ? "null" : "bytes"; });
                        return "ok";
                    """.Replace("ALVO", System.Text.Json.JsonSerializer.Serialize(alvo))))).Trim('"'));
                Assert.Equal("null", await EditorReal.Esperar(j, EditorReal.NoEditor("return String(w.__lido);"), "null", 50));
                await Task.Delay(1000);
                Assert.Equal(0, j.ConversoesTerminadas);
                Assert.False(Directory.Exists(Path.Combine(j.SessaoAtual!.PastaRecovery, ConversaoParaEditor.SubpastaConvertidos)));
            }), TimeSpan.FromSeconds(120));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // Dados > De texto/CSV: loadLocalFile entrega os bytes do CSV escolhido (os acentos chegam ao editor)
    public void ImportarCsv_LoadLocalFile()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("csv");
        var xlsx = FabricaDocumentos.Xlsx(Directory.CreateDirectory(Path.Combine(pasta, "base")).FullName, "Base");
        var csv = Path.Combine(pasta, "municipios.csv");
        File.WriteAllText(csv, "Município;Ação\r\nSão José;Execução fiscal\r\n", new System.Text.UTF8Encoding(true));
        var filtros = new List<string>();
        try
        {
            JanelaSta.Executar(() => NaJanela("mplanilha", xlsx, async j =>
            {
                // Espiao em volta de loadLocalFile (so' observa: a chamada segue para a ponte de verdade) e o caminho do web-apps
                // Dados > Do texto/CSV > De arquivo: api.asc_TextFromFileOrUrl(opcoes, cb) -> OpenFilenameDialog("csv/txt") -> loadLocalFile.
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var d = w.AscDesktopEditor, orig = d.loadLocalFile;
                        w.__lido = "nao chamado"; w.__cbr = "nao chamado";
                        d.loadLocalFile = function (c, cb) {
                          return orig.call(d, c, function (u) { w.__lido = u ? new TextDecoder("utf-8").decode(u) : "null"; cb(u); });
                        };
                        api.asc_TextFromFileOrUrl(new w.Asc.asc_CTextOptions(), function (r) { w.__cbr = r ? "objeto" : "null"; });
                        return "ok";
                    """))).Trim('"'));
                var lido = "";
                for (var i = 0; i < 100 && !lido.Contains("Execução"); i++)
                {
                    await Task.Delay(100);
                    lido = System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(EditorReal.NoEditor("return String(w.__lido);"))) ?? "";
                }
                saida.WriteLine("lido: " + lido);
                Assert.Contains("Município;Ação", lido);
                Assert.Contains("São José;Execução fiscal", lido);
                Assert.Equal("objeto", await EditorReal.Esperar(j, EditorReal.NoEditor("return String(w.__cbr);"), "objeto", 50));
                Assert.Equal("[]", System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(ErrosJs)));
            }, j => j.EscolherArquivoParaTeste = (f, _) => { filtros.Add(f); return [csv]; }), TimeSpan.FromSeconds(120));
            Assert.Equal(["csv/txt"], filtros);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // hiperlink "Selecionar arquivo": dialogo "any" (nao o de imagem), sem consentimento de copia para media/
    public void Hiperlink_SelecionarArquivo_NaoEhDialogoDeImagem()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("any");
        var xlsx = FabricaDocumentos.Xlsx(Directory.CreateDirectory(Path.Combine(pasta, "base")).FullName, "Base");
        // Um PNG de verdade: se o dialogo "any" desse consentimento de imagem, LocalFileGetImageUrl o copiaria para media/.
        var alvo = Path.Combine(pasta, "anexo.png");
        File.WriteAllBytes(alvo, FabricaDocumentos.PngSolido(8, 0, 0, 255));
        var filtros = new List<string>();
        try
        {
            JanelaSta.Executar(() => NaJanela("mplanilha", xlsx, async j =>
            {
                j.EscolherImagemParaTeste = () => throw new InvalidOperationException("nao devia abrir o dialogo de imagem");
                // web-apps (hiperlink -> "Selecionar arquivo"): api.asc_getFilePath(cb) -> OpenFilenameDialog("any") -> cb(caminho).
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        w.__p = "nao chamado"; api.asc_getFilePath(function (p) { w.__p = String(p); }); return "ok";
                    """))).Trim('"'));
                var p = "";
                for (var i = 0; i < 50 && p != alvo; i++)
                {
                    await Task.Delay(100);
                    p = System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(EditorReal.NoEditor("return String(w.__p);"))) ?? "";
                }
                Assert.Equal(alvo, p);
                Assert.Equal("error", (await j.ExecutarJsAsync(EditorReal.NoEditor(
                    "return String(w.AscDesktopEditor.LocalFileGetImageUrl(" + System.Text.Json.JsonSerializer.Serialize(alvo) + "));"))).Trim('"'));
                var media = Path.Combine(j.SessaoAtual!.PastaRecovery, "media");
                Assert.True(!Directory.Exists(media) || !Directory.EnumerateFiles(media).Any());
            }, j => j.EscolherArquivoParaTeste = (f, _) => { filtros.Add(f); return [alvo]; }), TimeSpan.FromSeconds(120));
            Assert.Equal(["any"], filtros);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // hifenizacao automatica: o hyph_pt_BR.dic e' servido com 200 pelo moffice-dic.local
    public void Hifenizacao_DicionarioServido()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("hifen");
        var docx = FabricaDocumentos.Docx(pasta, string.Join(" ", Enumerable.Repeat("extraordinariamente inconstitucionalissimamente", 30)));
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                // O pedido do PROPRIO motor (vkg em word/sdk-all-min.js: XHR a getDictionariesPath() + "pt_BR/hyph_pt_BR.dic") depois de ligar a
                // hifenizacao automatica com o texto em pt-BR, medido pelo Resource Timing do quadro (status HTTP incluso).
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        api.asc_EditSelectAll();
                        if (typeof api.put_TextPrLang === "function") api.put_TextPrLang(1046);
                        api.asc_RemoveSelection();
                        api.asc_setAutoHyphenation(true);
                        return "ok";
                    """))).Trim('"'));
                const string Pedido = """
                    (function () {
                      try {
                        var w = window.frames[0];
                        var e = w.performance.getEntriesByType("resource").filter(function (r) { return /hyph_pt_BR\.dic$/.test(r.name); });
                        return e.length ? e[0].name + " " + e[0].responseStatus : "nenhum";
                      } catch (x) { return "erro: " + x; }
                    })()
                    """;
                var r = "nenhum";
                for (var i = 0; i < 100 && r == "nenhum"; i++) { await Task.Delay(100); r = (await j.ExecutarJsAsync(Pedido)).Trim('"'); }
                saida.WriteLine("pedido do motor: " + r);
                Assert.Equal("https://moffice-dic.local/pt_BR/hyph_pt_BR.dic 200", r);
                Assert.Empty(j.Guarda.Bloqueadas);
            }), TimeSpan.FromSeconds(120));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // "Abrir recente" do menu Arquivo lista os recentes que a casca forneceu
    public void AbrirRecente_ListaPreenchida()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("recentes");
        var docx = FabricaDocumentos.Docx(pasta, "Base");
        string[] recentes = [@"C:\Processos\peticao inicial.docx", @"\\servidor\juridico\contestacao.docx", @"C:\x\ignorado.pages"];
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                // Common.Controllers.Desktop guarda a lista que window.onupdaterecents recebeu (recentFiles()) e a manda ao menu (update:recents).
                const string Lista = """
                    (function () {
                      try {
                        var w = window.frames[0], d = w.Common && w.Common.Controllers && w.Common.Controllers.Desktop;
                        if (!d || typeof d.recentFiles !== "function") return "sem Desktop";
                        return JSON.stringify((d.recentFiles() || []).map(function (r) { return r.url; }));
                      } catch (e) { return "erro: " + e; }
                    })()
                    """;
                var lista = "";
                for (var i = 0; i < 100 && !lista.Contains("contestacao"); i++)
                {
                    await Task.Delay(100);
                    lista = System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(Lista)) ?? "";
                }
                saida.WriteLine("recentes no editor: " + lista);
                Assert.Equal(System.Text.Json.JsonSerializer.Serialize(recentes.Take(2)), lista);
            }, janela: EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { FonteDeRecentes = () => recentes })), TimeSpan.FromSeconds(120));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // MvP-13: inserir grafico -> Editar dados -> mudar um valor -> salvar e reabrir: o valor persiste e o quadro aninhado nao tem erro de JS
    public void Grafico_EditarDados_Persiste()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("grafico");
        var docx = FabricaDocumentos.Docx(pasta, "Documento com grafico");
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                // O editor de dados do grafico e' um quadro ANINHADO (spreadsheeteditor/main/index_internal.html) dentro de uma janela do web-apps;
                // o controlador dele so' existe depois do "script:loaded" do web-apps.
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.NoEditor(
                    "return String(!!(w.Common && w.Common.Controllers.LaunchController && w.Common.Controllers.LaunchController.isScriptLoaded()));"), "true", 150));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        api.asc_addChartDrawingObject(w.Asc.c_oAscChartTypeSettings.barNormal, undefined, true); return "ok";
                    """))).Trim('"'));
                await Task.Delay(1000);
                // Grafico > Editar dados (web-apps ChartTab setEditData -> api.asc_editChartInFrameEditor()).
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_editChartInFrameEditor(); return 'ok';"))).Trim('"'));
                // Os dados do grafico chegam ao quadro aninhado depois do "documentReady" dele: a planilha de dados do grafico novo do motor
                // se chama "Sheet1" (uma pasta vazia do editor em pt-BR seria "Planilha1"). (asc_getCellInfo().asc_getText() devolve "" para
                // numero neste quadro — MEDIDO ao vivo —, por isso o valor e' conferido no arquivo salvo.)
                const string Aninhado = """
                    var f = w.frames[0];
                    if (!f || !f.Asc || !f.Asc.editor || typeof f.Asc.editor.asc_getWorksheetName !== "function") return "carregando";
                    try { return String(f.Asc.editor.asc_getWorksheetName(0)); } catch (e) { return "carregando"; }
                    """;
                var folha = "carregando";
                for (var i = 0; i < 150 && folha != "Sheet1"; i++) { await Task.Delay(200); folha = (await j.ExecutarJsAsync(EditorReal.NoEditor(Aninhado))).Trim('"'); }
                saida.WriteLine("planilha do editor de dados: " + folha);
                if (folha != "Sheet1")
                {
                    saida.WriteLine("diagnostico: " + await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var f = w.frames[0], r = { ifr: [].map.call(w.document.querySelectorAll("iframe"), function (x) { return (x.src || "").slice(0, 140); }) };
                        try { r.api = !!(f && f.Asc && f.Asc.editor); r.erros = f && f.__moffice && f.__moffice.errosJs; r.folhas = f.Asc.editor.asc_getWorksheetsCount(); } catch (e) { r.e = String(e); }
                        return JSON.stringify(r);
                        """)));
                    saida.WriteLine("log: " + string.Join("\n", File.ReadAllLines(j.ArquivoDeLog).Where(l => l.Contains("ERROR") || l.Contains("SEM SUPORTE") || l.Contains("REJECTION") || l.Contains("falhou")).TakeLast(20)));
                }
                Assert.Equal("Sheet1", folha);
                await Task.Delay(1000);
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var a = w.frames[0].Asc.editor;
                        a.asc_findCell("B2"); a.pluginMethod_InputText("99"); a.asc_closeCellEditor();
                        return "ok";
                    """))).Trim('"'));
                await Task.Delay(1000);
                // Fecha a janela do editor de dados no X (o web-apps grava os dados no grafico ao fechar).
                Assert.Equal("fechou", (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var js = [].filter.call(w.document.querySelectorAll(".asc-window"), function (e) {
                          return w.getComputedStyle(e).display !== "none" && e.textContent.indexOf("Editor de gr") >= 0; });
                        if (!js.length) return "sem janela";
                        var b = js[0].querySelector(".tool.close"); if (!b) return "sem botao";
                        b.click(); return "fechou";
                    """))).Trim('"'));
                await Task.Delay(1500);
                var erros = (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                        var f = w.frames[0];
                        return JSON.stringify({ editor: (w.__moffice && w.__moffice.errosJs) || [], aninhado: (f && f.__moffice && f.__moffice.errosJs) || [] });
                    """))).Trim('"').Replace("\\\"", "\"");
                saida.WriteLine("erros: " + erros);
                Assert.Equal("{\"editor\":[],\"aninhado\":[]}", erros);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                Assert.True((await j.SalvarAsync()).Gravou);
                Assert.Null(j.UltimoErro);
            }), TimeSpan.FromSeconds(240));
            using var z = System.IO.Compression.ZipFile.OpenRead(docx);
            var grafico = z.Entries.FirstOrDefault(e => e.FullName.StartsWith("word/charts/chart", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal));
            Assert.NotNull(grafico);
            using var leitor = new StreamReader(grafico!.Open());
            var xml = leitor.ReadToEnd();
            Assert.Contains("<c:v>99</c:v>", xml);
            Assert.DoesNotContain("<c:v>46</c:v>", xml);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // Fix 1 (seguranca): vinculo externo so' le' o arquivo depois do SIM do host; NAO -> nada de x2t e o motor recebe "nao";
                    // a decisao vale por caminho na sessao (sem segunda pergunta); dispositivo -> recusado com faixa, sem pergunta
    public void Vinculo_ConfirmacaoDoHost_NegarPermitirLembrar()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("vinculo");
        var xlsx = FabricaDocumentos.Xlsx(Directory.CreateDirectory(Path.Combine(pasta, "base")).FullName, "Base");
        var negada = FabricaDocumentos.Xlsx(Directory.CreateDirectory(Path.Combine(pasta, "negada")).FullName, "Salarios");
        var permitida = FabricaDocumentos.Xlsx(Directory.CreateDirectory(Path.Combine(pasta, "permitida")).FullName, "Custos");
        var perguntas = new List<string>();
        var resposta = false;
        var abertos = new List<string>();
        var avisos = new List<string>();
        try
        {
            JanelaSta.Executar(() => NaJanela("mplanilha", xlsx, async j =>
            {
                j.AbrirPedido += c => abertos.Add(c);
                j.AvisoAoUsuario += m => avisos.Add(m);
                string Converter(string c) => EditorReal.NoEditor("""
                        w.__cb = "nao chamado";
                        w.AscDesktopEditor.convertFile(ALVO, 8194, function (o) {
                          if (o === null) { w.__cb = "null"; return; }
                          var n = new Uint8Array(o.get()).length; o.close(); w.__cb = "bytes " + (n > 0);
                        });
                        return "ok";
                    """.Replace("ALVO", System.Text.Json.JsonSerializer.Serialize(c)));
                string Abrir(string c) => EditorReal.NoEditor("""
                        w.__ref = "nao chamado";
                        w.AscDesktopEditor.openExternalReference(ALVO, function (r) { w.__ref = String(r); });
                        return "ok";
                    """.Replace("ALVO", System.Text.Json.JsonSerializer.Serialize(c)));
                var cb = EditorReal.NoEditor("return String(w.__cb);");
                var refCb = EditorReal.NoEditor("return String(w.__ref);");
                var conv = Path.Combine(j.SessaoAtual!.PastaRecovery, ConversaoParaEditor.SubpastaConvertidos);

                // 1) NAO: o motor recebe null, o x2t nao roda; o mesmo caminho de novo nao pergunta outra vez.
                Assert.Equal("ok", (await j.ExecutarJsAsync(Converter(negada))).Trim('"'));
                Assert.Equal("null", await EditorReal.Esperar(j, cb, "null", 50));
                Assert.Single(perguntas);
                Assert.Contains(negada, perguntas[0]);
                Assert.False(Directory.Exists(conv), "o x2t rodou depois do NAO");
                Assert.Equal("ok", (await j.ExecutarJsAsync(Converter(negada))).Trim('"'));
                Assert.Equal("null", await EditorReal.Esperar(j, cb, "null", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(Abrir(negada))).Trim('"'));
                Assert.Equal("0", await EditorReal.Esperar(j, refCb, "0", 50));
                Assert.Single(perguntas);

                // 2) SIM: converte e entrega; segunda conversao e "Abrir origem" do mesmo caminho sem nova pergunta.
                resposta = true;
                Assert.Equal("ok", (await j.ExecutarJsAsync(Converter(permitida))).Trim('"'));
                Assert.Equal("bytes true", await EditorReal.Esperar(j, cb, "bytes true", 150));
                Assert.Equal(2, perguntas.Count);
                Assert.Equal("ok", (await j.ExecutarJsAsync(Converter(permitida))).Trim('"'));
                Assert.Equal("bytes true", await EditorReal.Esperar(j, cb, "bytes true", 150));
                Assert.Equal("ok", (await j.ExecutarJsAsync(Abrir(permitida))).Trim('"'));
                Assert.Equal("1", await EditorReal.Esperar(j, refCb, "1", 50));
                for (var i = 0; i < 20 && abertos.Count == 0; i++) await Task.Delay(100);
                Assert.Equal([permitida], abertos);
                Assert.Equal(2, perguntas.Count);

                // 3) Dispositivo: recusado sem pergunta, com faixa; o motor recebe null/0.
                avisos.Clear();
                Assert.Equal("ok", (await j.ExecutarJsAsync(Converter(@"\\?\" + permitida))).Trim('"'));
                Assert.Equal("null", await EditorReal.Esperar(j, cb, "null", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(Abrir(@"\\.\" + permitida))).Trim('"'));
                Assert.Equal("0", await EditorReal.Esperar(j, refCb, "0", 50));
                for (var i = 0; i < 20 && avisos.Count < 2; i++) await Task.Delay(100);
                Assert.Equal(2, perguntas.Count);
                Assert.Equal(2, avisos.Count);
                Assert.All(avisos, a => Assert.StartsWith("O mOffice não lê vínculos", a));
                Assert.Equal("[]", System.Text.Json.JsonSerializer.Deserialize<string>(await j.ExecutarJsAsync(ErrosJs)));
            }, j => j.ConfirmarVinculoParaTeste = m => { perguntas.Add(m); return resposta; }), TimeSpan.FromSeconds(180));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // painel Imprimir abre em "Impressao colorida" com uma impressora que suporta cor (printer:config injetado)
    public void Imprimir_ColoridoPorPadrao()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = PastaNova("cor");
        var docx = FabricaDocumentos.Docx(pasta, "Base");
        try
        {
            JanelaSta.Executar(() => NaJanela("mtexto", docx, async j =>
            {
                var json = mOffice.Nucleo.Impressao.ImpressorasDoWindows.JsonParaEditor("Colorida de teste",
                    [new mOffice.Nucleo.Impressao.ImpressoraInfo("Colorida de teste", false, true, [new mOffice.Nucleo.Impressao.PapelImpressora("A4", 210, 297)])]);
                saida.WriteLine("printer:config: " + (await j.ExecutarJsAsync(mOffice.Nucleo.Impressao.ImpressorasDoWindows.ScriptParaEditor(json))));
                // Abre o painel Arquivo -> Imprimir pelo mesmo caminho do Ctrl+P do web-apps (Main.onPrint) e le' o seletor de cor.
                const string Cor = """
                    (function () {
                      try {
                        var w = window.frames[0], c = w.DE && w.DE.getController("Print");
                        if (!c) return "sem controlador Print";
                        var v = c.printSettings || (c.getView && c.getView("PrintWithPreview"));
                        if (!v || !v.cmbColorPrinting) { w.DE.getController("Main").onPrint(); return "abrindo"; }
                        return String(v.cmbColorPrinting.getValue()) + (v.cmbColorPrinting.isDisabled() ? " desligado" : "");
                      } catch (e) { return "erro: " + e; }
                    })()
                    """;
                var r = "abrindo";
                for (var i = 0; i < 100 && r is "abrindo" or "sem controlador Print"; i++) { await Task.Delay(100); r = (await j.ExecutarJsAsync(Cor)).Trim('"'); }
                saida.WriteLine("seletor de cor: " + r);
                Assert.Equal("color", r);
            }), TimeSpan.FromSeconds(120));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
