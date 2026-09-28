using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 5 (S4) — figuras do documento VISIVEIS no editor, colar imagem e inserir imagem de arquivo.
///
/// A pergunta que ficou aberta no Marco 0 (relatorio da Task 5A, secao 5): o sdkjs guarda a pasta do
/// <c>DesktopOfflineAppDocumentEndLoad</c> em <c>AscCommon.Ss.tia</c> e prefixa <c>file://</c> nela; o
/// Chromium RECUSA subrecurso <c>file://</c> vindo de uma pagina <c>https://</c>, entao a midia do documento
/// nunca chegava a tela. Aqui isso e' provado dos dois lados: a requisicao da midia SAI (e volta 200 de
/// <c>moffice-doc.local</c>) e a figura APARECE (pixels vermelhos na captura da janela).
/// </summary>
[Collection(JanelaReal.Nome)]
public class MidiaTests
{
    private readonly ITestOutputHelper _saida;
    public MidiaTests(ITestOutputHelper saida) { _saida = saida; }

    /// <summary>Quadrado vermelho de 64x64 = 4096 pixels no documento. 500 na captura e' folgado: sobra
    /// margem para o zoom do editor e para a reamostragem da captura.</summary>
    private const int MinimoPixelsVermelhos = 500;

    [SkippableFact] // janela real: um DOCX com figura abre, a midia e' pedida por moffice-doc.local (200) e a figura aparece
    public void AbreDocxComImagem_FiguraVisivel()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-midia-").FullName;
        var docx = FabricaDocumentos.DocxComImagem(pasta, "Figura abaixo");

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor { InstrumentarRede = true });
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            // A figura e' desenhada depois do "documento pronto" (a imagem e' carregada por um <img> e so
            // entra no desenho quando chega); esperar ate a midia ser servida, e nao um delay as cegas.
            var servidas = await EsperarMidias(j, 1, 100);
            _saida.WriteLine($"midias servidas por {Midia.BaseUrlDocumento}/media/: {servidas}");
            _saida.WriteLine("base das midias no editor: " + await j.ExecutarJsAsync(BaseNoEditor));
            Assert.True(servidas >= 1,
                "nenhuma midia foi pedida por moffice-doc.local — a base file:// continua (ver Task 5A, secao 5)");

            byte[] png = [];
            var vermelhos = 0;
            for (var i = 0; i < 25 && vermelhos < MinimoPixelsVermelhos; i++)
            {
                await Task.Delay(200);
                png = await j.CapturarPngAsync();
                vermelhos = Pixels.ContarVermelhos(png);
            }
            _saida.WriteLine($"pixels vermelhos na captura: {vermelhos}");
            Assert.True(vermelhos >= MinimoPixelsVermelhos,
                $"a figura vermelha nao aparece na captura (pixels vermelhos: {vermelhos})");
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    [SkippableFact] // janela real: colar uma imagem (data URL) pela ponte grava em media/ e a figura e' inserida e aparece
    public void ColarImagem_GravaEmMediaEInsere()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-colar-").FullName;
        // Documento SEM figura nenhuma: tudo de vermelho que aparecer na tela veio da colagem.
        var docx = FabricaDocumentos.Docx(pasta, "Documento sem figura");
        var vermelho = FabricaDocumentos.PngSolido(64, 255, 0, 0);
        var dataUrl = "data:image/png;base64," + Convert.ToBase64String(vermelho);
        var salvo = Path.Combine(pasta, "com-figura.docx");
        Assert.Empty(FabricaDocumentos.ImagensDocx(docx));

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor { InstrumentarRede = true });
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            var media = Path.Combine(j.SessaoAtual!.PastaRecovery, "media");
            Assert.Empty(Directory.Exists(media) ? Directory.GetFiles(media) : []);

            // Caminho REAL do app oficial para colar/inserir imagem (sdkjs/word/sdk-all.js):
            //   Asc.asc_docs_api.prototype.AddImageUrl = nQc = urls.map(u =>
            //       AscCommon.Ss.RS(window.AscDesktopEditor.LocalFileGetImageUrl(u)))  ->  this.RHa(urls)
            // ou seja: a ponte grava o arquivo em media/ e devolve o NOME; o sdkjs monta a URL e insere.
            var r = (await j.ExecutarJsAsync(ColarImagem(dataUrl))).Trim('"');
            _saida.WriteLine("AddImageUrl: " + r);
            Assert.StartsWith("ok", r);

            var servidas = await EsperarMidias(j, 1, 100);
            _saida.WriteLine($"midias servidas: {servidas}");

            var arquivos = string.Join(", ", Directory.GetFiles(media).Select(Path.GetFileName));
            _saida.WriteLine("media/: " + arquivos);
            Assert.Equal("image1.png", arquivos);

            var vermelhos = 0;
            for (var i = 0; i < 25 && vermelhos < MinimoPixelsVermelhos; i++)
            {
                await Task.Delay(200);
                vermelhos = Pixels.ContarVermelhos(await j.CapturarPngAsync());
            }
            _saida.WriteLine($"pixels vermelhos depois de colar: {vermelhos}");
            Assert.True(vermelhos >= MinimoPixelsVermelhos,
                $"a figura colada nao aparece na captura (pixels vermelhos: {vermelhos})");

            // A figura tem de SOBREVIVER ao salvamento: o editor serializa a midia pelo caminho inverso
            // (`fe.G6` tira o prefixo `tia + "/media/"` da URL e devolve o nome), entao trocar a base tambem
            // muda o que e' gravado — se a troca quebrasse esse lado, o DOCX salvo sairia sem imagem nenhuma.
            await j.SalvarComoAsync(salvo, FormatoDocumento.Docx);
            Assert.Null(j.UltimoErro);
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));

        var imagens = FabricaDocumentos.ImagensDocx(salvo);
        _saida.WriteLine($"imagens no DOCX salvo: {imagens.Count} ({string.Join(", ", imagens.Select(i => i.Length + " bytes"))})");
        Assert.Single(imagens);
        Assert.Equal(vermelho, imagens[0]);
    }

    [SkippableFact] // janela real: "inserir imagem do arquivo" (OpenFilenameDialog -> LocalFileGetImageUrl) insere a figura
    public void InserirImagemDeArquivo_PassaPeloDialogoEInsere()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-inserir-").FullName;
        var docx = FabricaDocumentos.Docx(pasta, "Documento sem figura");
        var imagem = Path.Combine(pasta, "vermelho.png");
        File.WriteAllBytes(imagem, FabricaDocumentos.PngSolido(64, 255, 0, 0));

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor { InstrumentarRede = true });
            // Costura do teste: no app o arquivo vem do OpenFileDialog do Windows (modal — impossivel num
            // teste automatizado). O RESTO do caminho e' o mesmo do produto: o editor chama
            // OpenFilenameDialog("images", false, cb), o host devolve o CAMINHO escolhido pelo callback, e o
            // sdkjs chama LocalFileGetImageUrl(caminho) por conta propria.
            j.EscolherImagemParaTeste = () => imagem;
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));

            var r = (await j.ExecutarJsAsync(InserirImagemDoDialogo)).Trim('"');
            _saida.WriteLine("asc_addImage: " + r);
            Assert.StartsWith("ok", r);

            var media = Path.Combine(j.SessaoAtual!.PastaRecovery, "media");
            var gravou = "";
            for (var i = 0; i < 100 && gravou.Length == 0; i++)
            {
                await Task.Delay(200);
                if (Directory.Exists(media)) gravou = string.Join(", ", Directory.GetFiles(media).Select(Path.GetFileName));
            }
            _saida.WriteLine("media/: " + gravou);
            Assert.Equal("image1.png", gravou);

            var vermelhos = 0;
            for (var i = 0; i < 25 && vermelhos < MinimoPixelsVermelhos; i++)
            {
                await Task.Delay(200);
                vermelhos = Pixels.ContarVermelhos(await j.CapturarPngAsync());
            }
            _saida.WriteLine($"pixels vermelhos depois de inserir do arquivo: {vermelhos}");
            Assert.True(vermelhos >= MinimoPixelsVermelhos,
                $"a figura inserida do arquivo nao aparece na captura (pixels vermelhos: {vermelhos})");
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    public static TheoryData<FormatoDocumento> Editores => [FormatoDocumento.Docx, FormatoDocumento.Xlsx, FormatoDocumento.Pptx];

    [SkippableTheory] // janela real: a troca da base tem de achar a propriedade nos TRES editores
    [MemberData(nameof(Editores))]
    public void TrocaDaBase_AchaAPropriedadeNosTresEditores(FormatoDocumento fmt)
    {
        // Cada editor e' um bundle minificado SEPARADO e o nome da propriedade muda em cada um
        // (word: AscCommon.Ss.tia | slide: AscCommon.ru.Eda | cell: AscCommon.Cw.$pa). A primeira versao
        // desta troca procurava o nome do editor de TEXTO e, por isso, nao fazia nada no PPTX/XLSX: o
        // documento abria, nada falhava, e as figuras do slide continuavam invisiveis — em silencio.
        // Por isso a busca e' pelo VALOR (o `file://` + pasta que o EndLoad acabou de gravar) e por isso
        // este teste existe para os tres.
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory($"moffice-base-{fmt}-").FullName;
        var arquivo = fmt switch
        {
            FormatoDocumento.Xlsx => FabricaDocumentos.Xlsx(pasta, "Original mOffice"),
            FormatoDocumento.Pptx => FabricaDocumentos.Pptx(pasta, "Original mOffice"),
            _ => FabricaDocumentos.Docx(pasta, "Original mOffice"),
        };

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(arquivo);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            _saida.WriteLine($"[{fmt}] troca da base: {j.TrocaDaBaseDeMidia}");
            Assert.StartsWith("trocado ", j.TrocaDaBaseDeMidia);
            Assert.EndsWith(Midia.BaseUrlDocumento, j.TrocaDaBaseDeMidia);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    // -----------------------------------------------------------------------------------------------------
    // Controle negativo (achado maior da revisao 1): a troca da base FALHANDO tem de ficar OBSERVAVEL — o
    // texto continua editavel (o documento nao e' bloqueado), mas MidiasIndisponiveis/UltimoErro dizem que as
    // figuras nao vao aparecer, e ZERO midia e' servida. Sem DesligarTrocaDeBaseParaTeste, este mesmo cenario
    // so' aconteceria com um motor de verdade num bundle diferente do esperado — impossivel de fabricar num
    // teste sem essa costura.
    // -----------------------------------------------------------------------------------------------------

    [SkippableFact] // janela real: com a troca desligada pela costura de teste, o estado de falha fica visivel
    public void TrocaDaBaseFalha_MidiasIndisponiveisEUltimoErroFicamObservaveis_ZeroMidiaServida()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-midia-neg-").FullName;
        var docx = FabricaDocumentos.DocxComImagem(pasta, "Figura abaixo");

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor { InstrumentarRede = true, DesligarTrocaDeBaseParaTeste = true });
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            _saida.WriteLine("troca da base (desligada para teste): " + j.TrocaDaBaseDeMidia);
            Assert.DoesNotContain("trocado ", j.TrocaDaBaseDeMidia);

            // Tempo para a midia SER pedida se a troca tivesse funcionado (nao deve, aqui) — o mesmo padrao
            // de espera dos testes positivos, so' que confirmando que o numero fica em zero.
            var servidas = 0;
            for (var i = 0; i < 15; i++) { await Task.Delay(200); servidas = await j.ContarMidiasServidasAsync(); }
            _saida.WriteLine($"midias servidas com a troca desligada: {servidas}");
            Assert.Equal(0, servidas);

            Assert.True(j.MidiasIndisponiveis);
            Assert.False(string.IsNullOrEmpty(j.UltimoErro));
            _saida.WriteLine("UltimoErro: " + j.UltimoErro);
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    [SkippableFact] // controle POSITIVO da costura: com a troca LIGADA (o padrao), o mesmo documento continua verde
    public void TrocaDaBaseFunciona_ComACosturaDesligadaPorPadrao_MidiasIndisponiveisFicaFalso()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-midia-pos-").FullName;
        var docx = FabricaDocumentos.DocxComImagem(pasta, "Figura abaixo");

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor { InstrumentarRede = true });
            Assert.False(j.DesligarTrocaDeBaseParaTeste); // padrao: a costura nao interfere fora do teste negativo
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            var servidas = await EsperarMidias(j, 1, 100);
            _saida.WriteLine($"midias servidas (costura desligada): {servidas}");
            Assert.True(servidas >= 1);
            Assert.False(j.MidiasIndisponiveis);
            Assert.Null(j.UltimoErro);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    [SkippableFact] // janela real (fix 2 da revisao da Task 5, critico): reabrir na MESMA janela zera o estado de midia
    // do documento anterior — um B saudavel depois de um A com a troca falhada nao pode continuar dizendo
    // "MidiasIndisponiveis" nem carregar o UltimoErro velho. Mutacao: sem o reset no comeco de AbrirAsync,
    // o Assert.False(MidiasIndisponiveis) do segundo documento fica vermelho.
    public void ReabrirNaMesmaJanela_AposFalhaDaTroca_EstadoDeMidiaEZeradoEAFiguraAparece()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        // DocxComImagem grava sempre "sintetico-imagem.docx": uma pasta por documento para B nao sobrescrever A.
        var pastaA = Directory.CreateTempSubdirectory("moffice-midia-reabre-a-").FullName;
        var pastaB = Directory.CreateTempSubdirectory("moffice-midia-reabre-b-").FullName;
        var docxA = FabricaDocumentos.DocxComImagem(pastaA, "Figura A");
        var docxB = FabricaDocumentos.DocxComImagem(pastaB, "Figura B");

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor { InstrumentarRede = true, DesligarTrocaDeBaseParaTeste = true });
            j.Show();

            // A: troca desligada pela costura -> estado de falha, como no controle negativo.
            await j.AbrirAsync(docxA);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            Assert.True(j.MidiasIndisponiveis);
            Assert.False(string.IsNullOrEmpty(j.UltimoErro));
            Assert.Equal(0, await j.ContarMidiasServidasAsync());
            _saida.WriteLine("A (troca desligada): " + j.TrocaDaBaseDeMidia + " | UltimoErro: " + j.UltimoErro);

            // B: costura desligada, MESMA janela. NavegarAsync so' volta no NavigationCompleted da pagina nova,
            // entao o "pronto" lido abaixo e' o do documento B, nao um resto do A.
            j.DesligarTrocaDeBaseParaTeste = false;
            await j.AbrirAsync(docxB);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            var servidas = await EsperarMidias(j, 1, 100);
            _saida.WriteLine($"B (reaberto): troca da base: {j.TrocaDaBaseDeMidia} | midias servidas: {servidas}");
            Assert.True(servidas >= 1, "a midia de B nao foi servida por moffice-doc.local");
            Assert.False(j.MidiasIndisponiveis);
            Assert.Null(j.UltimoErro);
            Assert.StartsWith("trocado ", j.TrocaDaBaseDeMidia);

            var vermelhos = 0;
            for (var i = 0; i < 25 && vermelhos < MinimoPixelsVermelhos; i++)
            {
                await Task.Delay(200);
                vermelhos = Pixels.ContarVermelhos(await j.CapturarPngAsync());
            }
            _saida.WriteLine($"pixels vermelhos na captura de B: {vermelhos}");
            Assert.True(vermelhos >= MinimoPixelsVermelhos,
                $"a figura vermelha de B nao aparece na captura (pixels vermelhos: {vermelhos})");
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(4));
    }

    // -----------------------------------------------------------------------------------------------------
    // Unidades (sem janela)
    // -----------------------------------------------------------------------------------------------------

    [Fact] // image1.png, image2.png, ... dentro de <recovery>/media, como o GetLocalImageUrlLocal do app oficial
    public async Task GravarImagem_NumeraSequencial()
    {
        var sessao = SessaoDeTeste(out var pasta);
        var png = FabricaDocumentos.PngSolido(8, 255, 0, 0);
        Assert.Equal("image1.png", await Midia.GravarImagemAsync(sessao, png, "png"));
        Assert.Equal("image2.png", await Midia.GravarImagemAsync(sessao, FabricaDocumentos.PngSolido(8, 0, 255, 0), "png"));
        Assert.Equal("image3.png", await Midia.GravarImagemAsync(sessao, FabricaDocumentos.PngSolido(8, 0, 0, 255), ".PNG"));
        var media = Path.Combine(pasta, "media");
        Assert.Equal(["image1.png", "image2.png", "image3.png"], Directory.GetFiles(media).Select(Path.GetFileName).Order());
        Assert.Equal(png, await File.ReadAllBytesAsync(Path.Combine(media, "image1.png")));
    }

    [Fact] // a numeracao continua DEPOIS da midia que o x2t ja escreveu (o app oficial conta os arquivos da pasta)
    public async Task GravarImagem_ContinuaDepoisDaMidiaDoDocumento()
    {
        var sessao = SessaoDeTeste(out var pasta);
        var media = Directory.CreateDirectory(Path.Combine(pasta, "media")).FullName;
        await File.WriteAllBytesAsync(Path.Combine(media, "image1.png"), FabricaDocumentos.PngSolido(4, 1, 2, 3));
        await File.WriteAllBytesAsync(Path.Combine(media, "image2.jpg"), FabricaDocumentos.PngSolido(4, 1, 2, 3));
        Assert.Equal("image3.png", await Midia.GravarImagemAsync(sessao, FabricaDocumentos.PngSolido(8, 255, 0, 0), "png"));
    }

    [Fact] // conteudo que nao e' imagem NAO entra em media/ (mesma recusa do GetLocalImageUrlLocal: "error")
    public async Task GravarImagem_RecusaConteudoQueNaoEImagem()
    {
        var sessao = SessaoDeTeste(out var pasta);
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            Midia.GravarImagemAsync(sessao, System.Text.Encoding.UTF8.GetBytes("<html>nao sou imagem</html>"), "png"));
        var media = Path.Combine(pasta, "media");
        Assert.Empty(Directory.Exists(media) ? Directory.GetFiles(media) : []);
    }

    [Fact] // a extensao vem do CONTEUDO, nao do que o chamador disse (nada de nome de arquivo vindo da pagina)
    public async Task GravarImagem_ExtensaoSaiDoConteudo()
    {
        var sessao = SessaoDeTeste(out _);
        Assert.Equal("image1.png", await Midia.GravarImagemAsync(sessao, FabricaDocumentos.PngSolido(8, 255, 0, 0), "..\\..\\evil.exe"));
    }

    [Fact] // minor 4 da revisao 1: a trava agora e' POR SESSAO (pasta de recovery) — duas chamadas
    // CONCORRENTES na MESMA sessao ainda tem de serializar (sem isso as duas leriam a pasta no mesmo
    // instante, escolheriam o MESMO indice, e a segunda gravacao sobrescreveria a primeira)
    public async Task GravarImagem_DuasChamadasConcorrentesNaMesmaSessao_NaoColidem()
    {
        var sessao = SessaoDeTeste(out var pasta);
        var a = Midia.GravarImagemAsync(sessao, FabricaDocumentos.PngSolido(8, 255, 0, 0), "png");
        var b = Midia.GravarImagemAsync(sessao, FabricaDocumentos.PngSolido(8, 0, 255, 0), "png");

        var nomes = await Task.WhenAll(a, b);

        Assert.Equal(["image1.png", "image2.png"], nomes.Order());
        var media = Path.Combine(pasta, "media");
        Assert.Equal(["image1.png", "image2.png"], Directory.GetFiles(media).Select(Path.GetFileName).Order());
    }

    [Fact] // sessoes DIFERENTES (pastas de recovery diferentes) tem travas diferentes — cada uma numera a
    // partir de 1, sem esperar a outra
    public async Task GravarImagem_DuasSessoesDiferentes_NaoCompartilhamNumeracao()
    {
        var sessaoA = SessaoDeTeste(out var pastaA);
        var sessaoB = SessaoDeTeste(out var pastaB);

        var nomeA = await Midia.GravarImagemAsync(sessaoA, FabricaDocumentos.PngSolido(8, 255, 0, 0), "png");
        var nomeB = await Midia.GravarImagemAsync(sessaoB, FabricaDocumentos.PngSolido(8, 0, 255, 0), "png");

        Assert.Equal("image1.png", nomeA);
        Assert.Equal("image1.png", nomeB);
        Assert.NotEqual(pastaA, pastaB);
    }

    [Fact] // so' a pasta de midia do documento e' "midia do documento"
    public void EhMidiaDoDocumento_SoAsUrlsDaPastaDeMidia()
    {
        Assert.True(Midia.EhMidiaDoDocumento("https://moffice-doc.local/media/image1.png"));
        Assert.False(Midia.EhMidiaDoDocumento("https://moffice-doc.local/Editor.bin"));
        Assert.False(Midia.EhMidiaDoDocumento("https://moffice.local/web-apps/apps/api/documents/index.html"));
        Assert.False(Midia.EhMidiaDoDocumento("file:///C:/temp/media/image1.png"));
    }

    // -----------------------------------------------------------------------------------------------------

    private static SessaoDocumento SessaoDeTeste(out string pasta)
    {
        pasta = Directory.CreateTempSubdirectory("moffice-midia-u-").FullName;
        return new SessaoDocumento
        {
            ArquivoOrigem = Path.Combine(pasta, "doc.docx"),
            Formato = FormatoDocumento.Docx,
            PastaRecovery = pasta,
        };
    }

    private static async Task<int> EsperarMidias(JanelaEditor j, int minimo, int tentativas)
    {
        var n = 0;
        for (var i = 0; i < tentativas && n < minimo; i++)
        {
            await Task.Delay(200);
            n = await j.ContarMidiasServidasAsync();
        }
        return n;
    }

    private const string BaseNoEditor = """
        (function () {
          try {
            var w = (typeof AscCommon !== "undefined") ? window : window.frames[0];
            return String(w.AscCommon && w.AscCommon.Ss ? w.AscCommon.Ss.tia : "sem AscCommon.Ss");
          } catch (e) { return "erro: " + e; }
        })()
        """;

    private static string ColarImagem(string dataUrl) => """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            if (typeof api.AddImageUrl !== "function") return "sem AddImageUrl";
            api.AddImageUrl([DADOS]);
            return "ok AddImageUrl";
          } catch (e) { return "erro: " + e; }
        })()
        """.Replace("DADOS", System.Text.Json.JsonSerializer.Serialize(dataUrl));

    /// <summary>Caminho do menu "Inserir imagem do arquivo" do web-apps: `asc_addImage` (= `qUa` no
    /// sdk-all.js), que chama `AscDesktopEditor.OpenFilenameDialog("images", false, cb)` e, no callback,
    /// `LocalFileGetImageUrl(caminho)` + `editor.RHa([...])`.</summary>
    private const string InserirImagemDoDialogo = """
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            if (typeof api.asc_addImage !== "function") return "sem asc_addImage";
            api.asc_addImage();
            return "ok asc_addImage";
          } catch (e) { return "erro: " + e; }
        })()
        """;
}
