using mOffice.Conversor;
using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 7 na janela real: as teclas chegam pelo mesmo canal do teclado (CDP Input.dispatchKeyEvent no editor) e passam pela
/// captura da ponte, nao por um atalho do teste.</summary>
[Collection(JanelaReal.Nome)]
public class AtalhosJanelaTests
{
    private const int Ctrl = 2, Shift = 8;
    private const string Api = "(w.Asc && w.Asc.editor ? w.Asc.editor : w.editor)";
    private const string CabecalhoDoEditor = "w.document.querySelector('#title-doc-name').value";

    private static string NovaPasta() { var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }

    private static async Task<bool> AguardarAsync(Func<bool> condicao, int tentativas = 100)
    {
        for (var i = 0; i < tentativas && !condicao(); i++) await Task.Delay(200);
        return condicao();
    }

    private static async Task<bool> AguardarAsync(Func<Task<bool>> condicao, int tentativas = 100)
    {
        for (var i = 0; i < tentativas; i++) { if (await condicao()) return true; await Task.Delay(200); }
        return await condicao();
    }

    private static string NomeDoApp(string app) => app switch { "mplanilha" => "mPlanilha", "mslides" => "mSlides", _ => "mTexto" };

    private static async Task<JanelaEditor> AbrirNovaAsync(string app, FormatoDocumento f)
    {
        var j = EditorReal.PrepararParaTeste(new JanelaEditor(app) { NomeApp = NomeDoApp(app) });
        j.Show();
        await j.NovoAsync(f);
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
        return j;
    }

    /// <summary>Avalia <paramref name="expressao"/> no QUADRO do editor (variavel <c>w</c>) e devolve o texto, sem as aspas do JSON.</summary>
    private static async Task<string> NoQuadro(JanelaEditor j, string expressao) =>
        (await j.ExecutarJsAsync("(function(){try{var w=window.frames[0];return String(" + expressao + ");}catch(e){return 'erro: '+e;}})()")).Trim('"');

    [SkippableFact] // Ctrl+O e Ctrl+F6 viram AtalhoRecebido; Ctrl+W fecha (documento sem alteracao: sem pergunta)
    public void CtrlO_CtrlF6_CtrlW()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var recebidos = new List<AcaoDeAtalho>(); var fechou = false;
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mtexto", FormatoDocumento.Docx);
            j.AtalhoRecebido += a => { lock (recebidos) recebidos.Add(a); };
            j.Fechado += () => fechou = true;
            await j.EnviarTeclaParaTesteAsync("o", "KeyO", 0x4F, Ctrl);
            Assert.True(await AguardarAsync(() => { lock (recebidos) return recebidos.Contains(AcaoDeAtalho.Abrir); }), "Ctrl+O nao chegou ao host");
            await j.EnviarTeclaParaTesteAsync("F6", "F6", 0x75, Ctrl);
            Assert.True(await AguardarAsync(() => { lock (recebidos) return recebidos.Contains(AcaoDeAtalho.ProximaJanela); }), "Ctrl+F6 nao chegou ao host");
            Assert.False(fechou);
            await j.EnviarTeclaParaTesteAsync("w", "KeyW", 0x57, Ctrl);
            Assert.True(await AguardarAsync(() => fechou), "Ctrl+W nao fechou a janela");
            Assert.Contains(AcaoDeAtalho.Fechar, recebidos);
        }, TimeSpan.FromSeconds(180));
    }

    [SkippableFact] // Ctrl+F4 e Ctrl+W sao o mesmo "fechar"
    public void CtrlF4_Fecha()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var fechou = false;
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mtexto", FormatoDocumento.Docx);
            j.Fechado += () => fechou = true;
            await j.EnviarTeclaParaTesteAsync("F4", "F4", 0x73, Ctrl);
            Assert.True(await AguardarAsync(() => fechou), "Ctrl+F4 nao fechou a janela");
        }, TimeSpan.FromSeconds(180));
    }

    [SkippableFact] // F12 chama o Salvar como (EscolherDestinoParaTeste chamado); Shift+F12 grava o arquivo atual
    public void F12_SalvarComo_ShiftF12_Salvar()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta();
        var docx = FabricaDocumentos.Docx(pasta, "original");
        var dialogos = 0;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { NomeApp = "mTexto" });
                j.EscolherDestinoParaTeste = (_, _) => { dialogos++; return null; };   // F12: o usuario cancela o dialogo
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                await j.EnviarTeclaParaTesteAsync("F12", "F12", 0x7B, 0);
                Assert.True(await AguardarAsync(() => dialogos > 0), "F12 nao abriu o Salvar como");
                Assert.Equal(1, dialogos);
                Assert.Equal(docx, j.SessaoAtual!.ArquivoOrigem);   // cancelou: nada mudou

                await j.ExecutarJsAsync(EditorReal.InserirTextoWord("texto-do-shift-f12"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                await j.EnviarTeclaParaTesteAsync("F12", "F12", 0x7B, Shift);
                var gravou = await AguardarAsync(() => { try { return FabricaDocumentos.TextoDocx(docx).Contains("texto-do-shift-f12"); } catch (IOException) { return false; } }, 150);
                Assert.True(gravou, "Shift+F12 nao gravou o arquivo atual");
                Assert.Equal(1, dialogos);   // Salvar de um arquivo existente nao pergunta destino
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // mSlides: F5 entra em tela cheia do inicio; Esc sai; Shift+F5 comeca do slide atual (2o slide selecionado)
    public void Slides_F5_EShiftF5()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string pedidosF5 = "", pedidosShiftF5 = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mslides", FormatoDocumento.Pptx);
            try
            {
                j.Activate();
                await NoQuadro(j, Api + ".AddSlide()");   // o slide novo fica selecionado (o 2o)
                // O motor nao expoe a pagina da demonstracao: registra o pedido que a ponte faz ao painel de apresentacao (preview:start, pagina, ...).
                await NoQuadro(j, "(w.__inicios = [], w.Common.NotificationCenter.on('preview:start', function () { w.__inicios.push([].slice.call(arguments).join(',')); }), 1)");
                Assert.Equal("1", await NoQuadro(j, Api + ".getCurrentPage()"));

                // Shift+F5 primeiro: sair da apresentacao devolve a selecao ao 1o slide.
                await j.EnviarTeclaParaTesteAsync("F5", "F5", 0x74, Shift);
                Assert.True(await AguardarAsync(() => j.EmTelaCheia), "Shift+F5 nao iniciou a apresentacao");
                pedidosShiftF5 = await NoQuadro(j, "w.__inicios.join('|')");
                await j.EnviarTeclaParaTesteAsync("Escape", "Escape", 0x1B, 0);
                Assert.True(await AguardarAsync(() => !j.EmTelaCheia), "Esc nao saiu da tela cheia");

                await j.EnviarTeclaParaTesteAsync("F5", "F5", 0x74, 0);
                Assert.True(await AguardarAsync(() => j.EmTelaCheia), "F5 nao iniciou a apresentacao");
                pedidosF5 = await NoQuadro(j, "w.__inicios.join('|')");
            }
            finally
            {
                if (j.EmTelaCheia) j.DefinirTelaCheia(false);   // nunca deixar a tela presa, mesmo com o teste falhando
                j.Close();
            }
        }, TimeSpan.FromSeconds(240));
        Assert.Contains("0", pedidosF5.Split('|'));                       // F5: do inicio (pagina 0)
        Assert.Contains("1,false,false,true", pedidosShiftF5.Split('|')); // Shift+F5: do slide atual (o 2o) / "do slide atual"
    }

    [SkippableFact] // mPlanilha: Ctrl+G e F5 poem o foco na Caixa de Nome
    public void Planilha_CtrlG_F5_FocoNaCaixaDeNome()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string depoisCtrlG = "", depoisF5 = "";
        JanelaSta.Executar(async () =>
        {
            var j = await AbrirNovaAsync("mplanilha", FormatoDocumento.Xlsx);
            const string Ativo = "w.document.activeElement.id";
            await NoQuadro(j, "(w.document.getElementById('editor_sdk') || w.document.body).focus()");
            await j.EnviarTeclaParaTesteAsync("g", "KeyG", 0x47, Ctrl);
            Assert.True(await AguardarAsync(async () => await NoQuadro(j, Ativo) == "ce-cell-name"), "Ctrl+G nao focou a Caixa de Nome");
            depoisCtrlG = await NoQuadro(j, Ativo);
            await NoQuadro(j, "(function(){ w.document.getElementById('ce-cell-name').blur(); w.document.getElementById('editor_sdk').focus(); return 1; })()");
            Assert.NotEqual("ce-cell-name", await NoQuadro(j, Ativo));
            await j.EnviarTeclaParaTesteAsync("F5", "F5", 0x74, 0);
            Assert.True(await AguardarAsync(async () => await NoQuadro(j, Ativo) == "ce-cell-name"), "F5 nao focou a Caixa de Nome");
            depoisF5 = await NoQuadro(j, Ativo);
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.Equal("ce-cell-name", depoisCtrlG);
        Assert.Equal("ce-cell-name", depoisF5);
    }

    [SkippableFact] // documento novo: titulo "Documento<n>" na janela e no cabecalho do editor; depois do Salvar como, o nome do arquivo nos dois
    public void Titulo_NovoNumerado_ECabecalhoAposSalvarComo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta();
        string cabecalhoNovo = "", cabecalhoSalvo = "", tituloNovo = "", tituloSalvo = ""; var numero = 0;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = await AbrirNovaAsync("mtexto", FormatoDocumento.Docx);
                numero = j.SessaoAtual!.NumeroDoNovo;
                tituloNovo = j.Title;
                Assert.True(await AguardarAsync(async () => await NoQuadro(j, CabecalhoDoEditor) == "Documento" + numero), "cabecalho do editor nao mostrou o titulo do novo");
                cabecalhoNovo = await NoQuadro(j, CabecalhoDoEditor);
                await j.SalvarComoAsync(Path.Combine(pasta, "contrato.docx"), FormatoDocumento.Docx);
                tituloSalvo = j.Title;
                Assert.True(await AguardarAsync(async () => await NoQuadro(j, CabecalhoDoEditor) == "contrato.docx"), "cabecalho do editor nao mostrou o nome salvo");
                cabecalhoSalvo = await NoQuadro(j, CabecalhoDoEditor);
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
        Assert.True(numero >= 1);
        Assert.Equal($"Documento{numero} — mTexto", tituloNovo);
        Assert.Equal($"Documento{numero}", cabecalhoNovo);
        Assert.Equal("contrato.docx — mTexto", tituloSalvo);
        Assert.Equal("contrato.docx", cabecalhoSalvo);
    }
}
