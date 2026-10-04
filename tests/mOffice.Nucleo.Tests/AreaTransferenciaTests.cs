using System.Runtime.InteropServices;
using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 6 do Plano 3 — figuras de uma colagem do Word (<c>%TEMP%\msohtmlclip*\</c>). Regra pura, sem janela.
/// </summary>
public class AreaTransferenciaTests
{
    private static readonly string Temp = Path.Combine(Path.GetTempPath(), "mOffice-testes", "temp-falso");

    [Theory] // so' imagens dentro de <temp>\msohtmlclip*\ contam como clipe do Word
    [InlineData(@"msohtmlclip1\01\clip_image002.png", true)]
    [InlineData(@"MSOHTMLCLIP\01\clip_image004.jpg", true)]
    [InlineData(@"msohtmlclip1\01\clip_filelist.xml", false)]
    [InlineData(@"outra\01\clip_image002.png", false)]
    [InlineData(@"msohtmlclip1\..\..\segredo.png", false)]
    [InlineData(@"clip_image002.png", false)]
    [InlineData(@"msohtmlclip1\01\clip_image003.svg", false)] // SVG fica fora de proposito (pode trazer <script>)
    public void EhClipeDoWord_Contencao(string relativo, bool esperado)
        => Assert.Equal(esperado, PonteDesktop.EhClipeDoWord(Path.Combine(Temp, relativo), Temp));

    [Fact] // a forma file:/// que o HTML do Word traz (com %20) tambem e' reconhecida
    public void EhClipeDoWord_UrlFile()
    {
        var caminho = Path.Combine(Temp, "msohtmlclip1", "01", "clip image002.png");
        Assert.True(PonteDesktop.EhClipeDoWord(new Uri(caminho).AbsoluteUri, Temp));
    }

    [Fact] // prefixo de TEXTO igual nao e' contencao: "<temp>-irma\msohtmlclip1\..." comeca com o texto de <temp>
    public void EhClipeDoWord_PastaIrmaComMesmoPrefixo_Recusa()
    {
        var irma = Temp + "-irma";
        Assert.False(PonteDesktop.EhClipeDoWord(Path.Combine(irma, "msohtmlclip1", "01", "clip_image002.png"), Temp));
        Assert.False(PonteDesktop.EhClipeDoWord("https://exemplo.invalid/msohtmlclip1/01/clip_image002.png", Temp));
        Assert.False(PonteDesktop.EhClipeDoWord(@"msohtmlclip1\01\clip_image002.png", Temp)); // relativo: nunca
    }

    [Fact] // um junction DENTRO de msohtmlclip* apontando para fora nao vale como clipe (a leitura sairia da pasta)
    public void EhClipeDoWord_SegmentoQueELink_Recusa()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-clip-link-").FullName;
        var fora = Directory.CreateTempSubdirectory("moffice-clip-fora-").FullName;
        try
        {
            File.WriteAllBytes(Path.Combine(fora, "clip_image002.png"), FabricaDocumentos.PngSolido(8, 255, 0, 0));
            var clip = Path.Combine(raiz, "msohtmlclip1");
            Directory.CreateDirectory(clip);
            // Controle positivo: pasta de verdade com o mesmo arquivo passa.
            var real = Path.Combine(clip, "02");
            Directory.CreateDirectory(real);
            File.Copy(Path.Combine(fora, "clip_image002.png"), Path.Combine(real, "clip_image002.png"));
            Assert.True(PonteDesktop.EhClipeDoWord(Path.Combine(real, "clip_image002.png"), raiz));

            var link = Path.Combine(clip, "01");
            try { Directory.CreateSymbolicLink(link, fora); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                // Sem privilegio de symlink: junction (mklink /J) nao exige privilegio.
                var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe", $"/c mklink /J \"{link}\" \"{fora}\"")
                { CreateNoWindow = true, UseShellExecute = false })!;
                p.WaitForExit();
            }
            Assert.True(File.GetAttributes(link).HasFlag(FileAttributes.ReparsePoint), "o link de teste nao foi criado");
            Assert.True(File.Exists(Path.Combine(link, "clip_image002.png")));
            Assert.False(PonteDesktop.EhClipeDoWord(Path.Combine(link, "clip_image002.png"), raiz));
        }
        finally
        {
            var link = Path.Combine(raiz, "msohtmlclip1", "01");
            if (Directory.Exists(link)) Directory.Delete(link);   // remove o LINK, nunca o alvo
            Directory.Delete(raiz, true);
            Directory.Delete(fora, true);
        }
    }

    [SkippableFact] // 8.3 x nome longo: a MESMA pasta escrita das duas formas continua contida (e a de fora, nao)
    public void EhClipeDoWord_NomeCurto83_ENomeLongo()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-clip-nome-bem-longo-").FullName;
        try
        {
            var clipe = Path.Combine(raiz, "msohtmlclip1", "01", "clip_image002.png");
            Directory.CreateDirectory(Path.GetDirectoryName(clipe)!);
            File.WriteAllBytes(clipe, FabricaDocumentos.PngSolido(8, 255, 0, 0));
            var fora = Path.Combine(raiz, "outra-pasta-longa", "01", "clip_image002.png");
            Directory.CreateDirectory(Path.GetDirectoryName(fora)!);
            File.WriteAllBytes(fora, FabricaDocumentos.PngSolido(8, 255, 0, 0));

            var raizCurta = NomeCurto(raiz);
            var clipeCurto = NomeCurto(clipe);
            Skip.If(raizCurta == raiz || clipeCurto == clipe,
                "nomes 8.3 desligados neste volume (fsutil 8dot3name): não há forma curta para testar");

            // A expansao em si (GetLongPathNameW), inclusive com cauda que ainda nao existe.
            Assert.Equal(clipe, PonteDesktop.FormaLonga(clipeCurto), ignoreCase: true);
            Assert.Equal(Path.Combine(raiz, "ainda-nao-existe", "x.png"),
                PonteDesktop.FormaLonga(Path.Combine(raizCurta, "ainda-nao-existe", "x.png")), ignoreCase: true);
            // Contencao: MEDIDO que o .NET 10 ja' expande nomes com '~' dentro de Path.GetFullPath/GetRelativePath
            // (sem FormaLonga, e ate' sem GetFullPath, as linhas abaixo continuam verdes) — FormaLonga deixa isso
            // explicito em vez de depender de um detalhe do runtime; estas linhas guardam o COMPORTAMENTO.
            Assert.True(PonteDesktop.EhClipeDoWord(clipe, raiz));                            // controle
            Assert.True(PonteDesktop.EhClipeDoWord(clipe, raizCurta));                       // raiz curta, candidato longo
            Assert.True(PonteDesktop.EhClipeDoWord(clipeCurto, raiz));                       // candidato todo curto (MSOHTM~1\...)
            Assert.True(PonteDesktop.EhClipeDoWord(new Uri(clipeCurto).AbsoluteUri, raizCurta));
            Assert.False(PonteDesktop.EhClipeDoWord(NomeCurto(fora), raizCurta));            // fora de msohtmlclip*, em 8.3 tambem
        }
        finally { Directory.Delete(raiz, true); }
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetShortPathNameW(string longo, char[]? curto, uint tamanho);

    private static string NomeCurto(string caminho)
    {
        var n = GetShortPathNameW(caminho, null, 0);
        if (n == 0) return caminho;
        var buf = new char[n];
        var m = GetShortPathNameW(caminho, buf, n);
        return m == 0 || m >= n ? caminho : new string(buf, 0, (int)m);
    }

    [Fact] // um clipe real do Word vira midia do documento (nome em media/), sem dialogo nenhum
    public void ImagemUrl_ClipeDoWord_Grava()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-clip-").FullName;
        var clipe = Path.Combine(pasta, "msohtmlclip1", "01", "clip_image002.png");
        Directory.CreateDirectory(Path.GetDirectoryName(clipe)!);
        File.WriteAllBytes(clipe, FabricaDocumentos.PngSolido(8, 255, 0, 0));
        var s = new SessaoDocumento { ArquivoOrigem = "x.docx", Formato = FormatoDocumento.Docx, PastaRecovery = Directory.CreateTempSubdirectory("moffice-rec-").FullName };
        var ponte = new PonteDesktop(s) { PastaTempDoSistema = pasta };
        var nome = ponte.ImagemUrl(new Uri(clipe).AbsoluteUri);
        Assert.StartsWith("image", nome);
        Assert.True(File.Exists(Path.Combine(s.PastaRecovery, "media", nome)));
        // Nao consome nada: o Word cita a MESMA figura de novo numa segunda colagem e ela continua valendo.
        Assert.Equal(nome, ponte.ImagemUrl(clipe));
        Directory.Delete(pasta, true);
        Directory.Delete(s.PastaRecovery, true);
    }

    [Fact] // controle negativo: o MESMO arquivo fora de msohtmlclip* continua recusado (a regra nao abriu o %TEMP% todo)
    public void ImagemUrl_ImagemNoTempForaDoClipe_Recusa()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-clip-").FullName;
        var solto = Path.Combine(pasta, "outra", "01", "clip_image002.png");
        Directory.CreateDirectory(Path.GetDirectoryName(solto)!);
        File.WriteAllBytes(solto, FabricaDocumentos.PngSolido(8, 255, 0, 0));
        var s = new SessaoDocumento { ArquivoOrigem = "x.docx", Formato = FormatoDocumento.Docx, PastaRecovery = Directory.CreateTempSubdirectory("moffice-rec-").FullName };
        var ponte = new PonteDesktop(s) { PastaTempDoSistema = pasta };
        Assert.Equal("error", ponte.ImagemUrl(new Uri(solto).AbsoluteUri));
        Assert.False(Directory.Exists(Path.Combine(s.PastaRecovery, "media")));
        Directory.Delete(pasta, true);
        Directory.Delete(s.PastaRecovery, true);
    }
}

/// <summary>
/// Task 6 do Plano 3 — janela real: botoes Copiar/Colar do editor e arquivos soltos do Explorer.
///
/// A area de transferencia do Windows e' COMPARTILHADA com o usuario: todo teste que a escreve tira uma copia
/// antes (<see cref="CopiaDaAreaDeTransferencia"/>) e a devolve num <c>finally</c>, na mesma thread STA.
/// </summary>
[Collection(JanelaReal.Nome)]
public class AreaTransferenciaJanelaTests
{
    private readonly ITestOutputHelper _saida;
    public AreaTransferenciaJanelaTests(ITestOutputHelper saida) { _saida = saida; }

    private const int MinimoPixelsVermelhos = 500;

    [SkippableFact] // botao Colar: texto posto no clipboard do Windows entra no documento
    public void BotaoColar_ColaTextoDoClipboard()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(() => ComJanela(async j =>
        {
            var copia = CopiaDaAreaDeTransferencia.Tirar();
            try
            {
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                // O host so' manda o atalho com a janela ATIVA (o clique no botao a ativa; aqui, Activate()).
                await Ativar(j);
                CopiaDaAreaDeTransferencia.Tentar(() => System.Windows.Clipboard.SetText("colado-555"));

                // O que o botao Colar do web-apps chama: api.Paste() (sdkjs: asc_desktop_copypaste(api, "Paste")
                // -> foco na area de texto do editor -> AscDesktopEditor.Paste()).
                var r = (await j.ExecutarJsAsync(EditorReal.NoEditor("return String(api.Paste()) + ' ' + w.__moffice.ultimaAreaTransferencia;"))).Trim('"');
                _saida.WriteLine("api.Paste(): " + r);
                Assert.Equal("true paste:host", r);

                var texto = await EsperarTexto(j, "colado-555");
                _saida.WriteLine("texto do editor: " + texto);
                Assert.Contains("colado-555", texto);
                Assert.Empty(j.Guarda.Bloqueadas);
            }
            finally { copia.Restaurar(); }
        }), TimeSpan.FromMinutes(3));
    }

    [SkippableFact] // botao Copiar (clique REAL no botao da barra): texto selecionado vai para o clipboard do Windows
    public void BotaoCopiar_ColocaNoClipboard()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(() => ComJanela(async j =>
        {
            var copia = CopiaDaAreaDeTransferencia.Tirar();
            try
            {
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("copiado-444"))).Trim('"'));
                CopiaDaAreaDeTransferencia.Tentar(() => System.Windows.Clipboard.SetText("antes-do-copiar"));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_EditSelectAll(); return 'ok';"))).Trim('"'));

                // Clique de verdade no botao Copiar da barra (DevTools: Input.dispatchMouseEvent) — e' o gesto do
                // usuario que da' ativacao e foco ao execCommand("copy") dentro de AscDesktopEditor.Copy().
                await ClicarNoBotaoDaBarra(j, "id-toolbar-btn-copy");
                var noClipboard = "";
                for (var i = 0; i < 50 && !noClipboard.Contains("copiado-444"); i++)
                {
                    await Task.Delay(100);
                    noClipboard = CopiaDaAreaDeTransferencia.Tentar(System.Windows.Clipboard.GetText);
                }
                var caminho = (await j.ExecutarJsAsync("String(window.frames[0].__moffice.ultimaAreaTransferencia)")).Trim('"');
                _saida.WriteLine($"caminho: {caminho} | clipboard: {noClipboard}");
                Assert.Contains("copiado-444", noClipboard);
                Assert.Equal("copy:execCommand", caminho);
            }
            finally { copia.Restaurar(); }
        }), TimeSpan.FromMinutes(3));
    }

    [SkippableFact] // Recortar SEM foco do documento: vai pelo host e o texto nunca se perde (fica no documento ou no clipboard)
    public void Recortar_SemFocoDoDocumento_NaoPerdeOTexto()
    {
        // Medido na S7: sem foco do documento o execCommand("cut") devolve true, o sdkjs APAGA a selecao e nada
        // chega a' area de transferencia. A guarda (document.hasFocus() em areaDeTransferencia) manda o pedido ao
        // host; o host so' manda o Ctrl+X com a janela ativa. Qualquer que seja o caso, o marcador sobrevive.
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(() => ComJanela(async j =>
        {
            var copia = CopiaDaAreaDeTransferencia.Tirar();
            try
            {
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("recorte-777"))).Trim('"'));
                CopiaDaAreaDeTransferencia.Tentar(() => System.Windows.Clipboard.SetText("antes-do-recortar"));

                var r = (await j.ExecutarJsAsync(EditorReal.NoEditor("""
                    var foco = w.document.hasFocus();
                    api.asc_EditSelectAll();
                    var x = api.Cut();
                    return foco + " " + String(x) + " " + w.__moffice.ultimaAreaTransferencia;
                    """))).Trim('"');
                _saida.WriteLine("api.Cut() sem foco: " + r);
                Assert.Equal("false true cut:host", r);   // premissa (sem foco do documento) + caminho do host

                // O host pode ou nao mandar o Ctrl+X (so' com a janela ativa): esperar o efeito assentar e conferir
                // que o marcador esta' em ALGUM lugar.
                var noDocumento = "";
                var noClipboard = "";
                for (var i = 0; i < 15; i++)
                {
                    await Task.Delay(200);
                    noClipboard = CopiaDaAreaDeTransferencia.Tentar(System.Windows.Clipboard.GetText);
                    noDocumento = (await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor)).Trim('"');
                    if (noClipboard.Contains("recorte-777")) break;
                }
                _saida.WriteLine($"janela ativa: {j.IsActive} | documento: [{noDocumento}] | clipboard: [{noClipboard}]");
                Assert.True(noDocumento.Contains("recorte-777") || noClipboard.Contains("recorte-777"),
                    "o texto recortado se perdeu: nem no documento nem na area de transferencia");
            }
            finally { copia.Restaurar(); }
        }), TimeSpan.FromMinutes(3));
    }

    public static TheoryData<FormatoDocumento> Editores => [FormatoDocumento.Docx, FormatoDocumento.Xlsx, FormatoDocumento.Pptx];

    [SkippableTheory] // soltar imagem (pelo caminho do host, simulando o que o postMessage entrega) insere a figura nos tres editores
    [MemberData(nameof(Editores))]
    public void SoltarImagem_InsereNosTresEditores(FormatoDocumento fmt)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory($"moffice-soltar-{fmt}-").FullName;
        try
        {
            var imagem = Path.Combine(pasta, "vermelho.png");
            File.WriteAllBytes(imagem, FabricaDocumentos.PngSolido(64, 255, 0, 0));

            JanelaSta.Executar(() => ComJanela(async j =>
            {
                await j.NovoAsync(fmt);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var media = Path.Combine(j.SessaoAtual!.PastaRecovery, "media");
                var antes = Directory.Exists(media) ? Directory.GetFiles(media).Length : 0;
                Assert.True(Pixels.ContarVermelhos(await j.CapturarPngAsync()) < MinimoPixelsVermelhos, "ja' havia vermelho na tela antes de soltar");

                await j.ArquivosSoltos([imagem]);

                var depois = antes;
                for (var i = 0; i < 100 && depois == antes; i++)
                {
                    await Task.Delay(100);
                    depois = Directory.Exists(media) ? Directory.GetFiles(media).Length : 0;
                }
                _saida.WriteLine($"[{fmt}] media/: {antes} -> {depois} ({string.Join(", ", Directory.GetFiles(media).Select(Path.GetFileName))})");
                Assert.Equal(antes + 1, depois);

                var vermelhos = 0;
                for (var i = 0; i < 40 && vermelhos < MinimoPixelsVermelhos; i++)
                {
                    await Task.Delay(200);
                    vermelhos = Pixels.ContarVermelhos(await j.CapturarPngAsync());
                }
                _saida.WriteLine($"[{fmt}] pixels vermelhos depois de soltar: {vermelhos}");
                Assert.True(vermelhos >= MinimoPixelsVermelhos, $"a figura solta nao aparece ({fmt}, pixels vermelhos: {vermelhos})");
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                // Consentimento de UM uso: o mesmo caminho pedido de novo pela pagina, sem novo soltar, e' recusado.
                Assert.Equal("error", await PedirPelaPagina(j, imagem));
                Assert.Empty(j.Guarda.Bloqueadas);
            }), TimeSpan.FromMinutes(3));
        }
        finally { ApagarPasta(pasta); }
    }

    [SkippableFact] // soltar imagem que o editor NAO insere: o consentimento dado para ela e' revogado (nao sobra para a pagina)
    public void SoltarImagem_QueNaoEntra_RevogaOConsentimento()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-soltar-revoga-").FullName;
        try
        {
            var imagem = Path.Combine(pasta, "vermelho.png");
            File.WriteAllBytes(imagem, FabricaDocumentos.PngSolido(64, 255, 0, 0));

            JanelaSta.Executar(() => ComJanela(async j =>
            {
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                // O editor "nao pede o dialogo": asc_addImage trocado por uma funcao vazia — o resultado da insercao
                // e' "dialogo nao pedido", nunca "ok".
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_addImage = function () { }; return 'ok';"))).Trim('"'));

                await j.ArquivosSoltos([imagem]);

                var media = Path.Combine(j.SessaoAtual!.PastaRecovery, "media");
                Assert.False(Directory.Exists(media) && Directory.GetFiles(media).Length > 0, "a imagem nao deveria ter entrado");
                // Sem a revogacao, a pagina poderia pedir o arquivo depois por conta propria e ele seria copiado.
                Assert.Equal("error", await PedirPelaPagina(j, imagem));
                Assert.False(Directory.Exists(media) && Directory.GetFiles(media).Length > 0, "a pagina copiou o arquivo sem um novo soltar");
            }), TimeSpan.FromMinutes(3));
        }
        finally { ApagarPasta(pasta); }
    }

    [SkippableFact] // soltar um documento nao navega a janela para fora do editor e pede abertura
    public void SoltarDocumento_PedeAbertura()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-soltar-doc-").FullName;
        try
        {
            var docx = FabricaDocumentos.Docx(pasta, "Documento solto");
            var txt = Path.Combine(pasta, "nota.txt");
            File.WriteAllText(txt, "abre pelo dialogo de importacao");
            var desconhecido = Path.Combine(pasta, "nota.xyz");
            File.WriteAllText(desconhecido, "nao abre");

            JanelaSta.Executar(() => ComJanela(async j =>
            {
                var pedidos = new List<string>();
                j.AbrirPedido += pedidos.Add;
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var url = (await j.ExecutarJsAsync("location.href")).Trim('"');

                await j.ArquivosSoltos([docx, txt, desconhecido]);
                Assert.Equal([docx, txt], pedidos);                     // o .txt agora abre (Plano 6, Task 3); formato desconhecido nao vira pedido

                // O ouvinte de CAPTURA esta' instalado no quadro do editor: um drop de arquivo nao chega ao ondrop do
                // sdkjs (area do documento) e e' cancelado mesmo FORA da area do sdkjs (o <body>, onde nenhum tratador
                // do editor chamaria preventDefault) — o navegador nao abre o arquivo por cima do editor.
                var r = (await j.ExecutarJsAsync("""
                    (function () {
                      var w = window.frames[0];
                      function soltar(alvo) {
                        var dt = new w.DataTransfer();
                        dt.items.add(new w.File(["x"], "x.png", { type: "image/png" }));
                        var ev = new w.DragEvent("drop", { dataTransfer: dt, bubbles: true, cancelable: true });
                        var chegou = false;
                        alvo.addEventListener("drop", function () { chegou = true; });
                        alvo.dispatchEvent(ev);
                        return ev.defaultPrevented + " " + chegou;
                      }
                      var sdk = w.document.getElementById("editor_sdk");
                      if (!sdk) return "sem editor_sdk";
                      return soltar(sdk) + " | " + soltar(w.document.body) + " | " + (w.__moffice.soltos || 0);
                    })()
                    """)).Trim('"');
                _saida.WriteLine("drop sintetico no quadro: " + r);
                Assert.Equal("true false | true false | 2", r);

                await Task.Delay(500);
                Assert.Equal("true", (await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"'));
                Assert.Equal(url, (await j.ExecutarJsAsync("location.href")).Trim('"'));
                Assert.Equal([docx, txt], pedidos);                     // o File montado em JS nao traz caminho: nada novo
            }), TimeSpan.FromMinutes(3));
        }
        finally { ApagarPasta(pasta); }
    }

    [SkippableFact] // cadeia inteira sem o arrastar do sistema: File com CAMINHO REAL -> ouvinte -> mensagem do quadro -> host
    public void SoltarComCaminhoReal_PassaPeloOuvinteEPelaMensagemDoQuadro()
    {
        // Um arrastar do Explorer nao se automatiza; o que ele entrega a' pagina e' um File com caminho de disco.
        // DOM.setFileInputFiles (DevTools) produz exatamente isso num <input type=file> do quadro; o drop
        // sintetico com esses File passa pelo MESMO ouvinte de captura, pelo postMessageWithAdditionalObjects e
        // pelo WebMessageReceived do quadro (origem conferida) ate' ArquivosSoltos.
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-soltar-cadeia-").FullName;
        try
        {
            var imagem = Path.Combine(pasta, "vermelho.png");
            File.WriteAllBytes(imagem, FabricaDocumentos.PngSolido(64, 255, 0, 0));
            var docx = FabricaDocumentos.Docx(pasta, "Outro documento");

            JanelaSta.Executar(() => ComJanela(async j =>
            {
                var pedidos = new List<string>();
                j.AbrirPedido += pedidos.Add;
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("ok", (await j.ExecutarJsAsync("""
                    (function () {
                      var w = window.frames[0], i = w.document.createElement("input");
                      i.type = "file"; i.multiple = true; i.id = "teste-soltar"; i.style.display = "none";
                      w.document.body.appendChild(i);
                      return "ok";
                    })()
                    """)).Trim('"'));
                var wv = await j.ProntoAsync();
                var arvore = await wv.CallDevToolsProtocolMethodAsync("DOM.getDocument", """{"depth":-1,"pierce":true}""");
                int? docDoQuadro = null;
                using (var d = System.Text.Json.JsonDocument.Parse(arvore, new System.Text.Json.JsonDocumentOptions { MaxDepth = 4096 }))
                {
                    void Andar(System.Text.Json.JsonElement n)
                    {
                        if (docDoQuadro is null && n.TryGetProperty("nodeName", out var nome) && nome.GetString() == "IFRAME"
                            && n.TryGetProperty("contentDocument", out var cd))
                            docDoQuadro = cd.GetProperty("nodeId").GetInt32();
                        if (n.TryGetProperty("children", out var filhos)) foreach (var f in filhos.EnumerateArray()) Andar(f);
                    }
                    Andar(d.RootElement.GetProperty("root"));
                }
                Assert.NotNull(docDoQuadro);
                var achado = await wv.CallDevToolsProtocolMethodAsync("DOM.querySelector",
                    System.Text.Json.JsonSerializer.Serialize(new { nodeId = docDoQuadro, selector = "#teste-soltar" }));
                var noInput = System.Text.Json.JsonDocument.Parse(achado).RootElement.GetProperty("nodeId").GetInt32();
                await wv.CallDevToolsProtocolMethodAsync("DOM.setFileInputFiles",
                    System.Text.Json.JsonSerializer.Serialize(new { files = new[] { imagem, docx }, nodeId = noInput }));

                var r = (await j.ExecutarJsAsync("""
                    (function () {
                      var w = window.frames[0], i = w.document.getElementById("teste-soltar");
                      var dt = new w.DataTransfer();
                      for (var k = 0; k < i.files.length; k++) dt.items.add(i.files[k]);
                      var ev = new w.DragEvent("drop", { dataTransfer: dt, bubbles: true, cancelable: true });
                      w.document.getElementById("editor_sdk").dispatchEvent(ev);
                      return i.files.length + " " + ev.defaultPrevented;
                    })()
                    """)).Trim('"');
                _saida.WriteLine("drop com File de caminho real: " + r);
                Assert.Equal("2 true", r);

                var media = Path.Combine(j.SessaoAtual!.PastaRecovery, "media");
                var arquivos = 0;
                for (var k = 0; k < 100 && (arquivos == 0 || pedidos.Count == 0); k++)
                {
                    await Task.Delay(100);
                    arquivos = Directory.Exists(media) ? Directory.GetFiles(media).Length : 0;
                }
                _saida.WriteLine($"media/: {arquivos} | pedidos: {string.Join(", ", pedidos)}");
                Assert.Equal(1, arquivos);                  // a imagem entrou no documento
                Assert.Equal([docx], pedidos);              // o documento virou pedido de abertura
                Assert.Equal("true", (await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"'));
            }), TimeSpan.FromMinutes(3));
        }
        finally { ApagarPasta(pasta); }
    }

    /// <summary>Janela de teste (EditorReal.NovaJanela) que SEMPRE fecha, passe ou falhe o corpo.</summary>
    private static async Task ComJanela(Func<JanelaEditor, Task> corpo)
    {
        var j = EditorReal.NovaJanela();
        try
        {
            j.Show();
            await corpo(j);
        }
        finally
        {
            try { j.Close(); } catch (InvalidOperationException) { /* ja' fechando */ }
        }
    }

    private static void ApagarPasta(string pasta)
    {
        try { if (Directory.Exists(pasta)) Directory.Delete(pasta, true); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* melhor esforco: arquivo ainda preso */ }
    }

    /// <summary>A PAGINA pedindo um caminho a LocalFileGetImageUrl por conta propria (sem soltar nem dialogo).</summary>
    private static async Task<string> PedirPelaPagina(JanelaEditor j, string caminho) =>
        (await j.ExecutarJsAsync($"window.frames[0].AscDesktopEditor.LocalFileGetImageUrl({System.Text.Json.JsonSerializer.Serialize(caminho)})")).Trim('"');

    /// <summary>Ativa a janela de teste (o clique do usuario faria isso). Sem ativacao o teste nao prova nada.</summary>
    private static async Task Ativar(JanelaEditor j)
    {
        for (var i = 0; i < 20 && !j.IsActive; i++)
        {
            j.Activate();
            await Task.Delay(100);
        }
        Assert.True(j.IsActive, "a janela de teste nao ficou ativa (Activate recusado pelo Windows?)");
    }

    // -------------------------------------------------------------------------------------------------------

    private static async Task<string> EsperarTexto(JanelaEditor j, string trecho)
    {
        var texto = "";
        for (var i = 0; i < 50 && !texto.Contains(trecho); i++)
        {
            await Task.Delay(200);
            texto = (await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor)).Trim('"');
        }
        return texto;
    }

    /// <summary>Clique de mouse de verdade (DevTools Input.dispatchMouseEvent) no centro de um botao da barra do
    /// web-apps, que vive no quadro do editor: coordenadas = retangulo do quadro + retangulo do botao.</summary>
    private static async Task ClicarNoBotaoDaBarra(JanelaEditor j, string id)
    {
        var json = (await j.ExecutarJsAsync($$"""
            (function () {
              var f = document.querySelector("iframe");
              var b = window.frames[0].document.getElementById({{System.Text.Json.JsonSerializer.Serialize(id)}});
              if (!f || !b) return "";
              var q = f.getBoundingClientRect(), r = b.getBoundingClientRect();
              return JSON.stringify({ x: q.left + r.left + r.width / 2, y: q.top + r.top + r.height / 2, w: r.width });
            })()
            """));
        var texto = System.Text.Json.JsonSerializer.Deserialize<string>(json) ?? "";
        Assert.False(string.IsNullOrEmpty(texto), "botao " + id + " nao encontrado na barra");
        using var doc = System.Text.Json.JsonDocument.Parse(texto);
        var x = doc.RootElement.GetProperty("x").GetDouble();
        var y = doc.RootElement.GetProperty("y").GetDouble();
        Assert.True(doc.RootElement.GetProperty("w").GetDouble() > 0, "botao " + id + " invisivel");
        var wv = await j.ProntoAsync();
        foreach (var tipo in new[] { "mouseMoved", "mousePressed", "mouseReleased" })
            await wv.CallDevToolsProtocolMethodAsync("Input.dispatchMouseEvent", System.Text.Json.JsonSerializer.Serialize(new
            {
                type = tipo, x, y, button = tipo == "mouseMoved" ? "none" : "left", clickCount = tipo == "mouseMoved" ? 0 : 1,
            }));
    }
}

/// <summary>
/// Copia da area de transferencia do Windows, para devolver ao usuario depois do teste (melhor esforco: cada
/// formato disponivel e' copiado por valor; um formato que nao se deixa ler fica de fora). Tem de ser usada numa
/// thread STA — a mesma de <see cref="JanelaSta.Executar"/>.
/// </summary>
internal sealed class CopiaDaAreaDeTransferencia
{
    private readonly System.Windows.DataObject? _copia;
    private CopiaDaAreaDeTransferencia(System.Windows.DataObject? copia) { _copia = copia; }

    public static CopiaDaAreaDeTransferencia Tirar()
    {
        var atual = Tentar(System.Windows.Clipboard.GetDataObject);
        if (atual is null) return new(null);
        var copia = new System.Windows.DataObject();
        var formatos = 0;
        foreach (var formato in Tentar(() => atual.GetFormats(false)) ?? [])
        {
            try
            {
                var dado = atual.GetData(formato, false);
                if (dado is null) continue;
                copia.SetData(formato, dado, false);
                formatos++;
            }
            catch (Exception e) when (e is COMException or InvalidOperationException or OutOfMemoryException or ArgumentException) { }
        }
        return new(formatos == 0 ? null : copia);
    }

    public void Restaurar()
    {
        if (_copia is null) Tentar(System.Windows.Clipboard.Clear);
        else Tentar(() => System.Windows.Clipboard.SetDataObject(_copia, true));
    }

    /// <summary>A area de transferencia pode estar aberta por outro programa naquele instante (CLIPBRD_E_CANT_OPEN):
    /// tenta de novo algumas vezes antes de desistir.</summary>
    public static T Tentar<T>(Func<T> acao)
    {
        for (var i = 0; ; i++)
        {
            try { return acao(); }
            catch (Exception e) when (e is COMException or System.Runtime.InteropServices.ExternalException && i < 10) { Thread.Sleep(50); }
        }
    }

    public static void Tentar(Action acao) => Tentar(() => { acao(); return 0; });
}
