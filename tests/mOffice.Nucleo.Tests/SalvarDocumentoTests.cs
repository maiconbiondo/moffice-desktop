using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Round-trip completo pelo EDITOR (nao pelo x2t direto, que a Task 2 ja cobre): abre um DOCX, edita pela API
/// publica do editor, pede o salvamento pelo mesmo caminho do app oficial (asc_Save -> DesktopOfflineApp
/// DocumentStartSave -> AscDesktopEditor.LocalFileSave) e confere o arquivo regravado.
/// </summary>
[Collection(JanelaReal.Nome)]
public class SalvarDocumentoTests
{
    private readonly ITestOutputHelper _saida;
    public SalvarDocumentoTests(ITestOutputHelper saida) { _saida = saida; }

    [SkippableFact] // janela real: abre DOCX, insere texto pela API do editor, salva, e o DOCX regravado contem o texto novo E o antigo
    public void EditaESalva_DocxContemTextoNovoEAntigo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-salvar-").FullName;
        var docx = FabricaDocumentos.Docx(pasta, "Texto original");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));

            // 1) o documento abriu com o conteudo antigo (linha de base da assercao final)
            Assert.Contains("Texto original", await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor));

            // 2) edicao pela API publica do editor (ver a nota em InserirTexto: neste build e' asc_AddText).
            var inserido = (await j.ExecutarJsAsync(EditorReal.InserirTexto)).Trim('"');
            _saida.WriteLine("insercao: " + inserido);
            Assert.StartsWith("ok", inserido);

            // 3) a edicao chegou ao MODELO do editor (e nao so ao canvas): o texto novo sai pela API de leitura.
            // SEM Task.Delay nenhum entre inserir e verificar (achado importante 2 da revisao 2026-09-17, fix
            // review 1: a mutacao "remove o WaitAsync" so falha o teste SE nao houver folga nenhuma para o
            // lote automatico do editor, medido em ~1,7s depois da edicao, chegar antes do SalvarAsync). Uma
            // unica leitura basta: asc_AddText muda o modelo SINCRONAMENTE, dentro da mesma chamada de script.
            Assert.Equal("true", (await j.ExecutarJsAsync(EditorReal.TemTextoNovo)).Trim('"'));

            // 4) salvar IMEDIATAMENTE, sem Task.Delay nenhum ate aqui (achado importante 2): a UNICA coisa que
            // pode colocar o texto novo no disco e' o proprio asc_Save descarregando os changes pendentes ANTES
            // do LocalFileSave chegar (SalvarComoAsync espera `tcs.Task.WaitAsync(EsperaSalvar)`, que so
            // completa depois do LocalFileSave — e o sdkjs so manda LocalFileSave depois de descarregar). Se
            // essa ordem (flush -> LocalFileSave -> converte) quebrasse, este teste passaria a depender de
            // sorte com o lote automatico do editor — e' exatamente o que a mutacao "remove o WaitAsync" prova
            // (ver o relatorio da Task 6, revisao 2026-09-17 fix review 1: sem esta folga eliminada, a mutacao
            // nao derrubava o teste, porque o lote automatico ainda tinha tempo de rodar antes do SalvarAsync).
            var changes = j.SessaoAtual!.ArquivoChanges;
            _saida.WriteLine($"changes0.json antes de salvar: {File.Exists(changes)}");
            // O evento que a casca usa para registrar os RECENTES (Task 6, revisao 1). Assinado ANTES do
            // salvamento: e' o unico aviso de "gravou" que o mOffice.Comum recebe, e ele tem de valer para
            // TODO caminho de salvamento — este teste cobre o SalvarAsync do host; o Ctrl+S do editor e o
            // "Salvar" do dialogo de fechamento passam pelo MESMO ConverterESinalizarAsync.
            var salvos = new List<string>();
            j.Salvo += c => salvos.Add(c);
            await j.SalvarAsync();
            Assert.Null(j.UltimoErro);

            // 4b) o Salvo disparou UMA vez, com o caminho realmente gravado.
            Assert.Equal([docx], salvos);

            // 5) a edicao passou pela PONTE (changes0.json). Sem changes, o x2t converteria so o Editor.bin
            //    original e o texto novo nao poderia aparecer no arquivo salvo.
            Assert.True(File.Exists(changes), "nenhum change gravado — a edição não chegou na ponte");
            _saida.WriteLine($"changes0.json: {new FileInfo(changes).Length} bytes, índice {j.SessaoAtual!.IndiceChanges}");

            // 6) o DOCX regravado tem o texto NOVO e o ANTIGO (round-trip real, nao truncamento)
            var texto = FabricaDocumentos.TextoDocx(docx);
            _saida.WriteLine("texto salvo: " + texto);
            Assert.Contains("mOffice PoC salvou", texto);
            Assert.Contains("Texto original", texto);

            // 7) o contador de changes do JS concorda com o do C# (achado da revisao final 2026-09-17).
            //    Quem responde LocalFileGetCurrentChangesIndex ao editor e' o JS (window.__moffice.indice);
            //    quem grava o changes0.json e' o C# (SessaoDocumento.IndiceChanges). Se os dois discordarem,
            //    o proximo lote de changes pede delecao a partir do lugar errado e o editor perde edicao —
            //    em silencio, que e' o modo de falha desta pilha.
            Assert.Equal(j.SessaoAtual!.IndiceChanges.ToString(), (await j.ExecutarJsAsync(IndiceNoQuadroJs)).Trim('"'));

            // 7b) e concorda TAMBEM no caso que divergia: deleteIndex < 0 ("sem delecao, so' apenda a partir
            //     do indice atual"). O JS fazia `deleteIndex + count` sempre, ou seja `-1 + count`, enquanto o
            //     C# soma `count` ao indice corrente. Fica por ULTIMO, depois de todas as asserções sobre o
            //     arquivo salvo, porque apenda um change de verdade a' sessao.
            var antes = j.SessaoAtual!.IndiceChanges;
            await j.ExecutarJsAsync(ApendarChangeNoQuadroJs);
            _saida.WriteLine($"indice apos LocalFileSaveChanges(-1, 1): C# {j.SessaoAtual!.IndiceChanges}, antes {antes}");
            Assert.Equal(antes + 1, j.SessaoAtual!.IndiceChanges);
            Assert.Equal(j.SessaoAtual!.IndiceChanges.ToString(), (await j.ExecutarJsAsync(IndiceNoQuadroJs)).Trim('"'));

            // 9) EXPORTAR para PDF nao e' salvar (achado da revisao 2): grava o arquivo, dispara Exportado
            //    (nunca Salvo) e deixa a SESSAO onde estava — titulo, arquivo de origem e o estado de
            //    "modificado" do documento seguem sendo os do DOCX. Sem isto, um "salvar como PDF" fazia o
            //    proximo Ctrl+S tentar regravar o PDF em vez do documento do usuario.
            var pdf = Path.Combine(pasta, "exportado.pdf");
            var exportados = new List<string>();
            j.Exportado += c => exportados.Add(c);
            salvos.Clear();

            // 9a) EDITAR de novo ANTES de exportar (achado CRITICO da revisao 2 do branch). Sem esta edicao o
            //     documento ja' estava salvo e a exportacao nao teria nada a preservar — o teste passaria sem
            //     exercitar o defeito. O `asc_Save` que a exportacao dispara manda
            //     `onDocumentModifiedChanged(false)` ANTES do `LocalFileSave` (protocolo do Marco 0, passo 6
            //     antes do 7), entao o editor apaga o sinal de "tem alteracao pendente" e o host grava um PDF
            //     — que NAO guarda o documento editavel. Sem a restauracao do nucleo, o documento fica
            //     parecendo salvo: titulo sem asterisco, botao Salvar desligado e, o pior, o proximo Close()
            //     nao pergunta nada e apaga a pasta de recuperacao. As alteracoes somem.
            Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTexto)).Trim('"'));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
            Assert.True(j.Modificado, "linha de base: o documento tem de estar modificado antes de exportar");
            var origemAntes = j.SessaoAtual!.ArquivoOrigem;
            var formatoAntes = j.SessaoAtual!.Formato;
            var tituloAntes = j.Title;
            Assert.Contains("*", tituloAntes);                      // linha de base do asterisco
            var modificadoDepois = new List<bool>();
            j.ModificadoAlterado += m => modificadoDepois.Add(m);

            await j.SalvarComoAsync(pdf, FormatoDocumento.Pdf);
            Assert.Null(j.UltimoErro);
            Assert.True(File.Exists(pdf), "o PDF nao foi gravado");
            Assert.Equal("%PDF-", File.ReadAllText(pdf)[..5]);
            Assert.Equal([pdf], exportados);
            Assert.Empty(salvos);                                   // exportar NAO e' salvar
            Assert.Equal(origemAntes, j.SessaoAtual!.ArquivoOrigem); // a sessao continua no DOCX
            Assert.Equal(formatoAntes, j.SessaoAtual!.Formato);

            // 9b) e a ALTERACAO PENDENTE sobrevive a' exportacao: o documento editavel nao foi gravado em
            //     lugar nenhum, entao continua "modificado" para todos os efeitos — inclusive o asterisco do
            //     titulo e o aviso que a casca usa para (re)habilitar o botao Salvar. Mutacao que prova o
            //     teste: tirar o RestaurarModificadoSeHouveEdicao() do ramo de exportacao derruba estas tres.
            Assert.True(j.Modificado, "a alteracao pendente sumiu depois de exportar para PDF");
            Assert.Equal(tituloAntes, j.Title);                     // com o asterisco de volta
            Assert.Contains(true, modificadoDepois);                // a casca foi avisada de que continua pendente

            // 10) nada saiu para fora das origens permitidas durante abrir + editar + salvar + exportar
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    // O que o EDITOR le' quando pergunta o indice de changes — a funcao do shim, no quadro do editor, e nao
    // a variavel de estado: e' esse valor que o sdkjs usa.
    private const string IndiceNoQuadroJs = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w || !w.AscDesktopEditor) return "sem quadro";
            return String(w.AscDesktopEditor.LocalFileGetCurrentChangesIndex());
          } catch (e) { return "erro: " + e; }
        })()
        """;

    // Chama o mesmo caminho do editor com deleteIndex = -1 (apendar sem deletar), como o desktop-sdk faz.
    private const string ApendarChangeNoQuadroJs = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w || !w.AscDesktopEditor) return "sem quadro";
            w.AscDesktopEditor.LocalFileSaveChanges("moffice-teste-indice", -1, 1);
            return "ok";
          } catch (e) { return "erro: " + e; }
        })()
        """;

    [SkippableFact] // janela real: Salvar como ODT a partir de um DOCX aberto
    public void SalvarComo_Odt_ContemTexto()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-salvarcomo-").FullName;
        var docx = FabricaDocumentos.Docx(pasta, "Texto para ODT");
        var odt = Path.Combine(pasta, "saida.odt");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));

            await j.SalvarComoAsync(odt, FormatoDocumento.Odt);
            Assert.Null(j.UltimoErro);

            Assert.True(File.Exists(odt), "o ODT não foi gravado");
            Assert.Contains("Texto para ODT", FabricaDocumentos.TextoOdf(odt));
            // "Salvar como" nao pode mexer no arquivo de origem
            Assert.Contains("Texto para ODT", FabricaDocumentos.TextoDocx(docx));
            // achado importante 1 da revisao 2026-09-17 (fix review 1): a sessao passa a apontar para o
            // arquivo NOVO (mesma decisao dos editores reais) — um SalvarAsync() seguinte deve gravar no ODT,
            // nao voltar a gravar no DOCX de origem.
            Assert.Equal(odt, j.SessaoAtual!.ArquivoOrigem);
            Assert.Equal(FormatoDocumento.Odt, j.SessaoAtual!.Formato);

            // AtualizarSessaoNoEditorAsync (achado 2 da revisao 2026-09-17, fix review 2): depois do "salvar
            // como" ela tem de (a) limpar o cache window.frames[0].__moffice.sessao — ponte-desktop.js so o lê
            // de novo (via SessaoJson(), que ja reflete a SessaoDocumento repontada acima) quando esse cache
            // esta nulo — e (b) chamar SetDocumentName, que em ponte-desktop.js escreve
            // `document.title = nome` no PROPRIO quadro do editor.
            Assert.Equal("null", (await j.ExecutarJsAsync(SessaoNoQuadroJs)).Trim('"'));
            Assert.Contains("saida.odt", (await j.ExecutarJsAsync(TituloNoQuadroJs)).Trim('"'));

            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    // window.frames[0].__moffice.sessao (cache lido por sessao() em ponte-desktop.js): String(null) => "null".
    private const string SessaoNoQuadroJs = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w || !w.__moffice) return "sem quadro";
            return String(w.__moffice.sessao);
          } catch (e) { return "erro: " + e; }
        })()
        """;

    // document.title do QUADRO do editor: e' o que SetDocumentName (ponte-desktop.js) escreve, nao o do
    // documento de nivel superior.
    private const string TituloNoQuadroJs = """
        (function () {
          try {
            var w = window.frames[0];
            if (!w) return "sem quadro";
            return String(w.document.title);
          } catch (e) { return "erro: " + e; }
        })()
        """;

    // -----------------------------------------------------------------------------------------------------
    // Achado 1 da revisao 2026-09-17 (fix review 2): um SalvarComoAsync que substitui outro ainda esperando
    // LocalFileSave tem de virar um InvalidOperationException com mensagem pt-BR (nao um TaskCanceledException
    // cru), e UltimoErro tem de refletir isso.
    //
    // SalvarComoAsync inteiro NAO da para exercitar sem janela: pede asc_Save de verdade via ExecutarJsAsync
    // (WebView2), e a primeira chamada completaria (asc_Save e' sincrono no sdkjs, LocalFileSave chega dentro
    // da mesma chamada de script) antes da segunda ter chance de cancelar o TCS dela — nao ha como produzir o
    // cancelamento disparando duas chamadas reais back-to-back. Por isso o teste dirige o cancelamento por um
    // seam interno (JanelaEditor.EsperarLocalFileSaveAsync, internal + InternalsVisibleTo para este projeto de
    // teste), que isola exatamente o `await tcs.Task.WaitAsync(...)` e o mapeamento de excecao ao redor dele,
    // sem precisar de WebView nenhuma.
    // -----------------------------------------------------------------------------------------------------
    [Fact]
    public void EsperarLocalFileSaveAsync_TcsCancelado_VaraInvalidOperationExceptionPtBr()
    {
        // new JanelaEditor() (InitializeComponent/DependencyObject) segue a mesma convencao do resto do
        // arquivo: precisa de thread STA com Dispatcher, mesmo sem Show()/AbrirAsync() nenhum — nao abre
        // janela real nem WebView2 nenhuma, entao nao precisa do PularSeNaoDaParaAbrirJanela (ONLYOFFICE/RDP).
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            var tcs = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            tcs.TrySetCanceled();

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => j.EsperarLocalFileSaveAsync(tcs));

            Assert.Equal("O pedido de salvamento foi substituído por outro antes de o editor responder.", ex.Message);
            Assert.Equal(ex.Message, j.UltimoErro);
        }, TimeSpan.FromSeconds(30));
    }
}
