using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 6, Task 2 — formatos legados, modelos, macros e apresentacao de slides pela JANELA real (o x2t e o editor de
/// verdade): legado abre em Modo de compatibilidade e o primeiro Ctrl+S vira "Salvar como" em OOXML sem tocar no original;
/// modelo aberto "do Explorer" vira documento novo; .docm salva no mesmo formato com o vbaProject.bin intacto;
/// .ppsx aberto para apresentar entra em tela cheia. Em sessao remota (RDP) estes testes sao ignorados (EditorReal).
/// </summary>
[Collection(JanelaReal.Nome)]
public class FormatosLegadosJanelaTests
{
    private readonly ITestOutputHelper _saida;
    public FormatosLegadosJanelaTests(ITestOutputHelper saida) { _saida = saida; }

    private static string NovaPasta(string prefixo) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, prefixo + Guid.NewGuid().ToString("N"))).FullName;

    private static string Sha(string arquivo) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(arquivo)));

    /// <summary>Ctrl+S do editor (o mesmo asc_Save que o atalho dispara; o teclado em si e' testado em SalvarComoDesktopTests).</summary>
    private const string SalvarPeloEditor = "api.asc_Save(false, false, true); return \"ok\";";

    private static async Task<bool> Aguardar(Func<bool> condicao, int tentativas = 150)
    {
        for (var i = 0; i < tentativas && !condicao(); i++) await Task.Delay(100);
        return condicao();
    }

    private static string Texto(string arquivo, string tipo) => tipo switch
    {
        "cell" => FabricaDocumentos.TextoXlsx(arquivo),
        "slide" => FabricaDocumentos.TextoPptx(arquivo),
        _ => FabricaDocumentos.TextoDocx(arquivo),
    };

    private static string Ext(FormatoDocumento f) => TabelaFormatos.Principal(f).Extensao;

    [SkippableTheory] // .doc/.xls/.ppt sinteticos abrem, a sessao fica em compatibilidade e o Ctrl+S vira Salvar como em OOXML
    [InlineData("mtexto", "amostra.doc", FormatoDocumento.Docx)]
    [InlineData("mplanilha", "amostra.xls", FormatoDocumento.Xlsx)]
    [InlineData("mslides", "amostra.ppt", FormatoDocumento.Pptx)]
    public void Legado_AbreEmCompatibilidade_ESalvarComoModerno(string app, string fixture, FormatoDocumento moderno)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-legado-");
        var original = Path.Combine(pasta, fixture);
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "legado", fixture), original);
        var shaAntes = Sha(original);
        var saida = Path.Combine(pasta, "saida" + Ext(moderno));
        var tipo = moderno.DocType();
        var marca = "legado-" + tipo;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela(app);
                var perguntas = new List<(string Filtro, FormatoDocumento Sugerido)>();
                var salvos = new List<string>();
                j.Salvo += c => salvos.Add(c);
                j.EscolherDestinoParaTeste = (filtro, sugerido) => { perguntas.Add((filtro, sugerido)); return new DestinoSalvar(saida, moderno); };
                j.Show();
                j.Activate();
                await j.AbrirAsync(original);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));

                Assert.True(j.ModoCompatibilidade);
                Assert.Equal(moderno, j.SessaoAtual!.Formato);
                Assert.Contains("[Modo de compatibilidade]", j.Title);
                Assert.StartsWith(Path.GetFileName(original), j.Title);
                Assert.StartsWith("Este documento está no formato antigo", j.AvisosDeAbertura[0]);

                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(tipo, marca))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor(SalvarPeloEditor))).Trim('"'));
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Ctrl+S nao gravou a copia moderna");

                // O Ctrl+S virou "Salvar como": o seletor foi chamado UMA vez, com as linhas do tipo e o moderno pre-selecionado.
                var p = Assert.Single(perguntas);
                Assert.Equal(moderno, p.Sugerido);
                Assert.DoesNotContain("97-2003", p.Filtro);   // o Salvar como nunca oferece o formato antigo
                Assert.Equal([saida], salvos);
                Assert.Null(j.UltimoErro);
                // A sessao apontou para o arquivo novo e saiu do Modo de compatibilidade.
                Assert.False(j.ModoCompatibilidade);
                Assert.Equal(saida, j.SessaoAtual!.ArquivoOrigem);
                Assert.DoesNotContain("[Modo de compatibilidade]", j.Title);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaAntes, Sha(original));   // o original legado nao foi tocado
            ValidadorDePacote.Validar(saida, moderno);
            Assert.Contains(marca, Texto(saida, tipo));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // .doc com destino legado escolhido a mao: a guarda recusa e o original fica intacto
    public void Legado_SalvarComExtensaoLegada_Recusado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-legado-recusa-");
        var original = Path.Combine(pasta, "amostra.doc");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "legado", "amostra.doc"), original);
        var shaAntes = Sha(original);
        var legado = Path.Combine(pasta, "x.doc");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.EscolherDestinoParaTeste = (_, _) => new DestinoSalvar(legado, FormatoDocumento.Docx);
                j.Show();
                j.Activate();
                await j.AbrirAsync(original);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType("word", "nao-grava"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor(SalvarPeloEditor))).Trim('"'));
                Assert.True(await Aguardar(() => j.UltimoErro is not null), "a guarda nao recusou o destino .doc");
                Assert.Equal(JanelaEditor.MensagemExtensaoNaoCombina(legado, FormatoDocumento.Docx), j.UltimoErro);
                Assert.True(j.ModoCompatibilidade);   // nada foi salvo: continua em compatibilidade e modificado
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.False(File.Exists(legado));
            Assert.Equal(shaAntes, Sha(original));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // M3 da revisao final do Plano 6: SalvarAsync (API do host) num .doc em compatibilidade recusa com a mensagem pt-BR, ANTES de
                    // pedir o salvamento ao editor — antes so' o Novo era recusado aqui
    public void Legado_SalvarAsync_RecusaComMensagem()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-legado-salvarasync-");
        var original = Path.Combine(pasta, "amostra.doc");
        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "legado", "amostra.doc"), original);
        var shaAntes = Sha(original);
        string? mensagem = null;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                try
                {
                    j.Show();
                    await j.AbrirAsync(original);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType("word", "nao-grava"))).Trim('"'));
                    mensagem = (await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarAsync())).Message;
                    Assert.True(j.ModoCompatibilidade);
                }
                finally { j.Close(); }
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(JanelaEditor.MensagemCompatibilidadeExigeSalvarComo, mensagem);
            Assert.Equal(shaAntes, Sha(original));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Modelo .dotx sintetico (WordprocessingDocumentType.Template) com um texto de marca.</summary>
    private static string Dotx(string pasta, string texto)
    {
        var docx = FabricaDocumentos.Docx(pasta, texto);
        var dotx = Path.Combine(pasta, "modelo.dotx");
        File.Move(docx, dotx);
        using (var d = WordprocessingDocument.Open(dotx, true)) d.ChangeDocumentType(WordprocessingDocumentType.Template);
        return dotx;
    }

    [SkippableFact] // .dotx aberto como NOVO a partir do modelo: titulo Documento, Novo, salvar pergunta destino .docx
    public void Modelo_NovoDoModelo_AbreComoDocumentoNovo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-modelo-");
        var dotx = Dotx(pasta, "Texto do modelo");
        var shaAntes = Sha(dotx);
        var saida = Path.Combine(pasta, "contrato.docx");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var perguntas = new List<FormatoDocumento>();
                var salvos = new List<string>();
                j.Salvo += c => salvos.Add(c);
                j.EscolherDestinoParaTeste = (_, sugerido) => { perguntas.Add(sugerido); return new DestinoSalvar(saida, FormatoDocumento.Docx); };
                j.Show();
                j.Activate();
                await j.NovoDoModeloAsync(dotx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.True(j.SessaoAtual!.Novo);
                Assert.Equal(FormatoDocumento.Docx, j.SessaoAtual.Formato);
                Assert.False(j.ModoCompatibilidade);
                Assert.StartsWith("Documento", j.Title);
                Assert.DoesNotContain("modelo", j.Title, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Texto do modelo", await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType("word", " usado-do-modelo"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor(SalvarPeloEditor))).Trim('"'));
                Assert.True(await Aguardar(() => salvos.Count > 0));
                Assert.Equal([FormatoDocumento.Docx], perguntas);
                Assert.Equal([saida], salvos);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaAntes, Sha(dotx));   // o modelo nao e' alterado
            var texto = FabricaDocumentos.TextoDocx(saida);
            Assert.Contains("Texto do modelo", texto);
            Assert.Contains("usado-do-modelo", texto);
            // E o modelo aberto pelo dialogo Abrir (AbrirAsync) edita o PROPRIO modelo, no mesmo formato.
            Assert.Equal(FormatoDocumento.Dotx, FormatoDocumentoExtensoes.DaExtensao(dotx));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    private static readonly byte[] VbaMarcador = [.. "VBA-FALSO-MOFFICE-PLANO6"u8.ToArray(), .. Enumerable.Range(0, 256).Select(i => (byte)i)];

    /// <summary>.docm sintetico: MacroEnabledDocument + VbaProjectPart com bytes marcadores (nao e' um projeto VBA valido).</summary>
    private static string Docm(string pasta, string texto)
    {
        var docx = FabricaDocumentos.Docx(pasta, texto);
        var docm = Path.Combine(pasta, "macro.docm");
        File.Move(docx, docm);
        using (var d = WordprocessingDocument.Open(docm, true))
        {
            d.ChangeDocumentType(WordprocessingDocumentType.MacroEnabledDocument);
            var vba = d.MainDocumentPart!.AddNewPart<VbaProjectPart>();
            using var ms = new MemoryStream(VbaMarcador);
            vba.FeedData(ms);
        }
        return docm;
    }

    private static byte[]? VbaDoPacote(string pacote)
    {
        using var z = ZipFile.OpenRead(pacote);
        var e = z.Entries.FirstOrDefault(x => x.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
        if (e is null) return null;
        using var s = e.Open(); using var ms = new MemoryStream(); s.CopyTo(ms);
        return ms.ToArray();
    }

    [SkippableFact] // .docm com vbaProject.bin marcador: abre, faixa de macros, salva .docm com o vbaProject.bin byte a byte
    public void Macro_SalvaNoMesmoFormato_PreservandoVba()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-macro-");
        var docm = Docm(pasta, "Texto com macro");
        Assert.Equal(VbaMarcador, VbaDoPacote(docm));   // linha de base: a fixture tem o VBA
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var salvos = new List<string>();
                j.Salvo += c => salvos.Add(c);
                j.EscolherDestinoParaTeste = (_, _) => throw new InvalidOperationException("o .docm existente nao devia perguntar o destino");
                j.Show();
                j.Activate();
                await j.AbrirAsync(docm);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.True(j.MacrosPreservadas);
                Assert.Equal(JanelaEditor.MensagemMacros, j.AvisosDeAbertura[0]);
                Assert.Equal(FormatoDocumento.Docm, j.SessaoAtual!.Formato);
                Assert.False(j.ModoCompatibilidade);
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirPorDocType("word", " editado-com-macro"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor(SalvarPeloEditor))).Trim('"'));
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Ctrl+S nao gravou o .docm");
                Assert.Equal([docm], salvos);
                Assert.Null(j.UltimoErro);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(VbaMarcador, VbaDoPacote(docm));   // o VBA continua IGUAL, byte a byte
            using (var z = ZipFile.OpenRead(docm))
            {
                using var r = new StreamReader(z.GetEntry("[Content_Types].xml")!.Open());
                Assert.Contains("macroEnabled", r.ReadToEnd());
            }
            using var d = WordprocessingDocument.Open(docm, false);
            Assert.Contains("editado-com-macro", d.MainDocumentPart!.Document.Body!.InnerText);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // .ppsx com AbrirEApresentarAsync: a ponte pede tela cheia (window.__moffice.telaCheia == true)
    public void Ppsx_AbrirEApresentar_EntraEmTelaCheia()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-ppsx-");
        var pptx = FabricaDocumentos.Pptx(pasta, "Slide de apresentacao");
        var ppsx = Path.Combine(pasta, "show.ppsx");
        File.Move(pptx, ppsx);
        using (var d = PresentationDocument.Open(ppsx, true)) d.ChangeDocumentType(PresentationDocumentType.Slideshow);
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mslides");
                try
                {
                    j.Left = 120; j.Top = 90; j.Width = 1100; j.Height = 760;
                    j.Show();
                    j.Activate();
                    await j.AbrirEApresentarAsync(ppsx);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    Assert.True(await Aguardar(() => j.EmTelaCheia, 200), "o .ppsx nao entrou em apresentacao (tela cheia)");
                    Assert.Equal("true", (await j.ExecutarJsAsync("String(window.__moffice.telaCheia === true)")).Trim('"'));
                    Assert.Equal(FormatoDocumento.Ppsx, j.SessaoAtual!.Formato);
                }
                finally
                {
                    if (j.EmTelaCheia) j.DefinirTelaCheia(false);   // nunca deixar a tela presa, mesmo com o teste falhando
                    j.Close();
                }
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // controle: o MESMO .ppsx aberto pelo Abrir comum (AbrirAsync) NAO apresenta — abre para edicao
    public void Ppsx_AbrirComum_NaoApresenta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-ppsx-edicao-");
        var pptx = FabricaDocumentos.Pptx(pasta, "Slide de edicao");
        var ppsx = Path.Combine(pasta, "show.ppsx");
        File.Move(pptx, ppsx);
        using (var d = PresentationDocument.Open(ppsx, true)) d.ChangeDocumentType(PresentationDocumentType.Slideshow);
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mslides");
                try
                {
                    j.Show();
                    j.Activate();
                    await j.AbrirAsync(ppsx);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    await Task.Delay(1500);
                    Assert.False(j.EmTelaCheia);
                    Assert.Equal(FormatoDocumento.Ppsx, j.SessaoAtual!.Formato);
                }
                finally { j.Close(); }
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
