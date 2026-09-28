using mOffice.Nucleo.Recuperacao;

namespace mOffice.Nucleo.Tests;

/// <summary>Revisao final do Plano 3 na janela real: C1 (encerramento do Windows com alteracao nao salva), I2 (original
/// alterado depois da queda) e M5 (acoes do host pedidas pela pagina de um documento anterior).</summary>
[Collection(JanelaReal.Nome)]
public class RevisaoFinalPlano3JanelaTests
{
    /// <summary>Fecha a janela pelo MESMO caminho do encerramento do WPF (Application.Shutdown, que e' o que o
    /// WM_QUERYENDSESSION de um logoff/reinicio do Windows dispara): <c>Window.InternalClose(shutdown: true,
    /// ignoreCancel: true)</c> — o Closing roda, mas o <c>e.Cancel</c> dele e' IGNORADO. Nao ha' como disparar um
    /// logoff de verdade num teste; este e' o metodo interno que o proprio WPF chama para cada janela.</summary>
    private static void FecharComoOEncerramentoDoWindows(System.Windows.Window j)
    {
        var fechar = typeof(System.Windows.Window).GetMethod("InternalClose",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, [typeof(bool), typeof(bool)]);
        Assert.NotNull(fechar);
        fechar!.Invoke(j, [true, true]);
    }

    private static void Limpar(params string[] pastas)
    {
        foreach (var p in pastas)
            try { if (!string.IsNullOrEmpty(p) && Directory.Exists(p)) Directory.Delete(p, true); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    /// <summary>Abre <paramref name="docx"/>, espera o editor, digita <paramref name="texto"/> e descarrega os changes no
    /// disco (a pasta passa a ter o que a recuperacao lista). Devolve a pasta de recuperacao.</summary>
    private static async Task<string> AbrirEditar(JanelaEditor j, string docx, string texto)
    {
        await j.AbrirAsync(docx);
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
        await j.ExecutarJsAsync(EditorReal.InserirTextoWord(texto));
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
        await j.DescarregarChangesAsync();
        return j.SessaoAtual!.PastaRecovery;
    }

    [SkippableFact] // C1 (critico): logoff/reinicio do Windows com alteracao nao salva — o WPF fecha IGNORANDO o e.Cancel,
                    // a pergunta "Salvar?" nunca roda; a pasta de recuperacao tem de FICAR (e aparecer como recuperavel),
                    // e a pergunta enfileirada nao pode rodar depois numa janela fechada. Controles: o mesmo encerramento
                    // SEM alteracao apaga a pasta; e "Nao salvar" (decisao do usuario) continua apagando.
    public void C1_EncerramentoDoWindows_ComAlteracao_MantemAPastaRecuperavel()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-c1-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var antes = File.ReadAllBytes(docx);
        string recovery = "", semAlteracao = "", naoSalvar = "";
        var perguntou = false; var fechou = false; var log = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.NaoSalvar); };
                j.Fechado += () => fechou = true;
                j.Show();
                log = j.ArquivoDeLog;
                recovery = await AbrirEditar(j, docx, " antes-do-logoff");
                Assert.True(j.Modificado);
                FecharComoOEncerramentoDoWindows(j);
                for (var i = 0; i < 50 && !fechou; i++) await Task.Delay(100);
                // A decisao que o AoFechar enfileirou (BeginInvoke) roda aqui, com o Dispatcher do teste ainda vivo — num
                // encerramento de verdade ela nem chega a rodar. Nos dois casos, nao pode perguntar nada.
                await Task.Delay(1500);
            }, TimeSpan.FromSeconds(180));
            Assert.True(fechou, "a janela nao fechou pelo caminho do encerramento");
            Assert.False(perguntou, "a pergunta enfileirada rodou depois de a janela fechar");
            Assert.True(Directory.Exists(recovery), "o encerramento do Windows apagou a unica copia do trabalho nao salvo");
            Assert.True(TravaSessao.Livre(recovery));
            Assert.Contains(recovery, Recuperaveis.Listar(SessaoDocumento.RaizRecovery, "mtexto").Select(d => d.Pasta));
            Assert.Equal(antes, File.ReadAllBytes(docx));   // e o arquivo do usuario nao foi tocado
            Assert.Contains("FECHAR sem a decisao do usuario", File.ReadAllText(log));

            // Controle 1: o MESMO encerramento, sem alteracao nenhuma — nada a preservar, a pasta sai.
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                semAlteracao = j.SessaoAtual!.PastaRecovery;
                FecharComoOEncerramentoDoWindows(j);
                for (var i = 0; i < 50 && Directory.Exists(semAlteracao); i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            Assert.False(Directory.Exists(semAlteracao));

            // Controle 2: com alteracao, fechamento normal e "Nao salvar" (decisao informada) — a pasta sai.
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.Show();   // a pergunta de teste responde NaoSalvar
                naoSalvar = await AbrirEditar(j, docx, " descartado");
                j.Close();
                for (var i = 0; i < 50 && Directory.Exists(naoSalvar); i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            Assert.False(Directory.Exists(naoSalvar));
            Assert.Equal(antes, File.ReadAllBytes(docx));
        }
        finally { Limpar(recovery, semAlteracao, naoSalvar, pasta); }
    }

    /// <summary>Abre o docx, edita e "cai" (fecha deixando a pasta, com a trava solta). Devolve a pasta de recuperacao.</summary>
    private static string EditarECair(string docx, string texto)
    {
        var recovery = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
            recovery = await AbrirEditar(j, docx, texto);
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.True(Recuperaveis.EhRecuperavel(recovery));
        return recovery;
    }

    [SkippableFact] // I2: o usuario reabriu e salvou o original DEPOIS da queda; a sessao recuperada nao grava por cima dele —
                    // o "Salvar" da API e' recusado e o Ctrl+S vira "Salvar como" com "<nome> (recuperado)"; o original
                    // fica byte a byte como o usuario o deixou
    public void I2_OriginalAlteradoDepoisDaQueda_SalvarViraSalvarComo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-i2-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var recovery = "";
        try
        {
            recovery = EditarECair(docx, " recuperado-i2");
            // Depois da queda o usuario abriu o original em outro lugar e salvou uma versao nova.
            var outra = Directory.CreateTempSubdirectory("mOffice-testes-i2-nova-").FullName;
            File.Copy(mOffice.Conversor.Tests.FabricaDocumentos.Docx(outra, "versao-mais-nova-do-usuario"), docx, true);
            File.SetLastWriteTimeUtc(docx, DateTime.UtcNow.AddMinutes(1));
            Directory.Delete(outra, true);
            var originalNovo = File.ReadAllBytes(docx);
            var copia = Path.Combine(pasta, Path.GetFileNameWithoutExtension(docx) + " (recuperado).docx");
            var pedidos = 0; string? nomeSugerido = null; var salvos = new List<string>();
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.EscolherDestinoParaTeste = (_, sugerido) => { pedidos++; nomeSugerido = JanelaEditor.NomeSugerido(j.SessaoAtual!); return new DestinoSalvar(copia, sugerido); };
                j.Salvo += c => salvos.Add(c);
                j.Show();
                await j.RecuperarAsync(recovery);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.True(j.OriginalAlteradoDepoisDaQueda);
                Assert.True(j.SessaoAtual!.Desvinculada);
                // "Salvar" pela API do host (grava no original): recusado, sem tocar no arquivo.
                var e = await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarAsync());
                Assert.Equal(JanelaEditor.MensagemOriginalAlteradoRecusado, e.Message);
                Assert.True(j.Modificado);
                Assert.Equal(originalNovo, File.ReadAllBytes(docx));
                Assert.Equal(0, pedidos);
                // Ctrl+S do editor: vira "Salvar como", com o nome sugerido que nao sobrescreve o original.
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"))).Trim('"'));
                for (var i = 0; i < 150 && salvos.Count < 1; i++) await Task.Delay(100);
                Assert.Equal(1, pedidos);
                Assert.Equal(Path.GetFileNameWithoutExtension(docx) + " (recuperado)", nomeSugerido);
                Assert.Equal([copia], salvos);
                Assert.False(j.SessaoAtual!.Desvinculada);
                Assert.False(j.Modificado);
                // O sessao.json agora aponta para a copia, com o carimbo DELA (uma nova queda nao fica desvinculada).
                var m = MetadadosSessao.Ler(recovery)!;
                Assert.Equal(copia, m.Origem);
                Assert.False(m.OrigemMudou());
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Equal(originalNovo, File.ReadAllBytes(docx));
            var texto = mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(copia);
            Assert.Contains("recuperado-i2", texto);
            Assert.Contains("Texto base", texto);
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // I2 (controle): original INTACTO desde a queda — a sessao recuperada salva normalmente no original
    public void I2_OriginalIntacto_SalvarGravaNoOriginal()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-i2c-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var recovery = "";
        try
        {
            recovery = EditarECair(docx, " recuperado-i2c");
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.EscolherDestinoParaTeste = (_, _) => throw new InvalidOperationException("nao devia perguntar o destino");
                j.Show();
                await j.RecuperarAsync(recovery);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.False(j.OriginalAlteradoDepoisDaQueda);
                Assert.True((await j.SalvarAsync()).Gravou);
                Assert.Null(j.UltimoErro);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Contains("recuperado-i2c", mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(docx));
            Assert.False(Directory.Exists(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // M5: Print/Colar/Copiar/Recortar vindos da ponte de um documento ANTERIOR (a pagina velha vive ate' o
                    // NavigationCompleted da nova) sao ignorados; controle: o mesmo Print pela ponte ATUAL gera o PDF
    public void M5_AcoesDaPaginaAnterior_SaoIgnoradas()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-m5-").FullName;
        var a = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "documento A");
        var pastaB = Directory.CreateDirectory(Path.Combine(pasta, "b")).FullName;
        var b = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pastaB, "documento B");
        var log = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.Show();
                log = j.ArquivoDeLog;
                await j.AbrirAsync(a);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                var ponteDeA = j.PonteParaTeste!;
                await j.AbrirAsync(b);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.NotSame(ponteDeA, j.PonteParaTeste);
                foreach (var tipo in new[] { "imprimir", "colar", "copiar", "recortar" })
                    ponteDeA.Avisar(tipo, tipo == "imprimir" ? "{}" : "");
                await Task.Delay(500);
                Assert.Null(j.ImpressaoEmCursoParaTeste);
                // Controle: o mesmo Print pela ponte do documento ATUAL chega ao host.
                var pdf = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                j.ImprimirPedido += (_, _, arquivo) => pdf.TrySetResult(arquivo);
                j.PonteParaTeste!.Avisar("imprimir", "{}");
                Assert.NotNull(j.ImpressaoEmCursoParaTeste);
                var arquivo = await pdf.Task.WaitAsync(TimeSpan.FromSeconds(60));
                mOffice.Nucleo.Impressao.PastaImpressao.Apagar(arquivo);
                j.Close();
            }, TimeSpan.FromSeconds(240));
            var texto = File.ReadAllText(log);
            foreach (var tipo in new[] { "imprimir", "colar", "copiar", "recortar" })
                Assert.Contains("PONTE aviso " + tipo + " ignorado: veio da pagina de um documento anterior", texto);
        }
        finally { Limpar(pasta); }
    }

    // --- Fix round 2 ------------------------------------------------------------------------------------------

    [SkippableFact] // C1 (fix round 2, item 4): editar -> salvar -> editar -> desfazer ate' o estado salvo -> fechar: o documento
                    // esta' inteiramente salvo, entao a pasta SAI (antes ficava e o documento reaparecia em "Documentos
                    // recuperados", com os changes de antes do salvamento)
    public void C1_EditarSalvarEditarDesfazer_FecharApagaAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-c1u-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = ""; var perguntou = false; var fechou = false;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.Cancelar); };
                j.Fechado += () => fechou = true;
                j.Show();
                recovery = await AbrirEditar(j, docx, " salvo-1");
                Assert.True((await j.SalvarAsync()).Gravou);
                Assert.False(j.Modificado);
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" desfeito-2"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.Undo(); return \"ok\";"))).Trim('"'));
                Assert.Equal("false", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "false", 100));
                Assert.False(j.Modificado);
                await j.DescarregarChangesAsync();
                Assert.True(new FileInfo(j.SessaoAtual!.ArquivoChanges).Length > 0, "o cenario precisa de changes no arquivo");
                j.Close();
                for (var i = 0; i < 50 && !fechou; i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            Assert.True(fechou);
            Assert.False(perguntou);
            Assert.False(Directory.Exists(recovery), "a pasta de um documento inteiramente salvo sobrou (reaparece como recuperavel)");
            Assert.Contains("salvo-1", mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(docx));
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // C1 (fix round 2, item 4): encerramento do Windows NO MEIO de um salvamento — o Ctrl+S de um documento novo
                    // ja' zerou o Modificado e o dialogo "Salvar como" esta' aberto: a pasta FICA e e' recuperavel
    public void C1_EncerramentoDuranteOSalvar_MantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        string recovery = ""; bool? modificadoNoDialogo = null; var fechou = false;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.Fechado += () => fechou = true;
                j.EscolherDestinoParaTeste = (_, _) =>
                {
                    modificadoNoDialogo = j.Modificado;
                    FecharComoOEncerramentoDoWindows(j);
                    return null;
                };
                j.Show();
                await j.NovoAsync(mOffice.Conversor.FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                recovery = j.SessaoAtual!.PastaRecovery;
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord("nunca salvo"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                await j.DescarregarChangesAsync();
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"))).Trim('"'));
                for (var i = 0; i < 100 && !fechou; i++) await Task.Delay(100);
                await Task.Delay(500);
            }, TimeSpan.FromSeconds(180));
            Assert.True(fechou, "a janela nao fechou");
            Assert.False(modificadoNoDialogo, "o cenario exige o Modificado ja' zerado pelo asc_Save");
            Assert.True(Directory.Exists(recovery), "o encerramento no meio do salvar apagou o documento novo nunca salvo");
            Assert.True(Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery); }
    }

    [SkippableFact] // itens 1/2 (fix round 2) de ponta a ponta: o changes0.json preso por outro programa sem compartilhar escrita
                    // -> o lote do editor NAO grava: o indice do JS e o do host NAO andam, o quadro congela, a janela fecha
                    // sem perguntar com a mensagem pt-BR, a pasta fica (arquivo byte a byte) e volta a ser recuperavel
    public void Alteracoes_ArquivoPreso_FalhaFechadoSemDivergirContadores()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-preso-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = "", resultado = "", mensagem = ""; string? mensagemDeRecuperacao = null; var perguntou = false; var fechou = false;
        byte[] antes = [];
        var indiceHostAntes = -1; var indiceHostDepois = -2;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.NaoSalvar); };
                j.AlteracoesNaoGravadas += (_, m) => mensagem = m;
                j.RecuperacaoFalhou += (_, m) => mensagemDeRecuperacao = m;
                j.Fechado += () => fechou = true;
                j.Show();
                recovery = await AbrirEditar(j, docx, " antes-de-prender");
                var arquivo = j.SessaoAtual!.ArquivoChanges;
                antes = File.ReadAllBytes(arquivo);
                indiceHostAntes = j.SessaoAtual.IndiceChanges;
                using (new FileStream(arquivo, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // O lote chega pelo MESMO caminho do editor (LocalFileSaveChanges do quadro), e o script devolve os dois
                    // indices do JS e o congelamento, lidos na mesma volta.
                    resultado = (await j.ExecutarJsAsync("""
                        (function () {
                          var w = window.frames[0], d = w.AscDesktopEditor;
                          var antes = d.LocalFileGetCurrentChangesIndex();
                          d.LocalFileSaveChanges("lote-preso", -1, 1);
                          return antes + "|" + d.LocalFileGetCurrentChangesIndex() + "|" + !!w.__moffice.congelado;
                        })()
                        """)).Trim('"');
                    indiceHostDepois = j.SessaoAtual.IndiceChanges;
                    for (var i = 0; i < 100 && !fechou; i++) await Task.Delay(100);
                }
            }, TimeSpan.FromSeconds(180));
            var partes = resultado.Split('|');
            Assert.Equal(3, partes.Length);
            Assert.Equal(partes[0], partes[1]);                // o indice do JS nao andou
            Assert.Equal("true", partes[2]);                   // o quadro congelou
            Assert.Equal(indiceHostAntes, indiceHostDepois);  // nem o do host
            Assert.True(fechou, "a janela nao fechou depois da falha");
            Assert.False(perguntou);
            // Fix round 3: evento e mensagem PROPRIOS da sessao viva (sem o "use Descartar" da recuperacao).
            Assert.Equal(JanelaEditor.MensagemFalhaAoGravarAlteracoes("sintetico.docx", "mOffice", recuperavel: true), mensagem);
            Assert.Null(mensagemDeRecuperacao);
            Assert.Equal(antes, File.ReadAllBytes(Path.Combine(recovery, "changes", "changes0.json")));
            Assert.True(Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    // --- Fix round 3 ------------------------------------------------------------------------------------------

    [SkippableTheory] // item 3a: a chamada SalvarChanges LANCA no JS (falha do proprio objeto de host): o JS avisa o host
                      // ("falhaGravacao") e a sessao falha fechada igual ao retorno "falhou" — indice do JS parado, quadro
                      // congelado, janela fechada sem perguntar, pasta mantida. Fix round 4: SEM changes gravados antes
                      // (`editarAntes` falso) a pasta nao tem o que recuperar e a mensagem NAO manda usar Recuperar.
    [InlineData(true)]
    [InlineData(false)]
    public void Alteracoes_ChamadaAoHostLanca_FalhaFechado(bool editarAntes)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-lanca-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = "", resultado = ""; string? mensagem = null; var perguntou = false; var fechou = false;
        byte[] antes = [];
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.NaoSalvar); };
                j.AlteracoesNaoGravadas += (_, m) => mensagem = m;
                j.Fechado += () => fechou = true;
                j.Show();
                if (editarAntes) recovery = await AbrirEditar(j, docx, " antes-de-lancar");
                else
                {
                    await j.AbrirAsync(docx);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    recovery = j.SessaoAtual!.PastaRecovery;
                }
                antes = File.Exists(j.SessaoAtual!.ArquivoChanges) ? File.ReadAllBytes(j.SessaoAtual.ArquivoChanges) : [];
                j.PonteParaTeste!.LancarEmSalvarChangesParaTeste = true;
                resultado = (await j.ExecutarJsAsync("""
                    (function () {
                      var w = window.frames[0], d = w.AscDesktopEditor;
                      var antes = d.LocalFileGetCurrentChangesIndex();
                      d.LocalFileSaveChanges("lote-perdido", -1, 1);
                      return antes + "|" + d.LocalFileGetCurrentChangesIndex() + "|" + !!w.__moffice.congelado;
                    })()
                    """)).Trim('"');
                for (var i = 0; i < 100 && !fechou; i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            var partes = resultado.Split('|');
            Assert.Equal(3, partes.Length);
            Assert.Equal(partes[0], partes[1]);
            Assert.Equal("true", partes[2]);
            Assert.True(fechou, "a janela nao fechou: o host nao soube que o lote se perdeu");
            Assert.False(perguntou);
            Assert.Equal(JanelaEditor.MensagemFalhaAoGravarAlteracoes("sintetico.docx", "mOffice", recuperavel: editarAntes), mensagem);
            var arquivoChanges = Path.Combine(recovery, "changes", "changes0.json");
            Assert.Equal(antes, File.Exists(arquivoChanges) ? File.ReadAllBytes(arquivoChanges) : []);
            Assert.Equal(editarAntes, Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // item 3b: a descarga (impressao) com o changes0.json preso: o lote do autosave falha, a sessao congela e
                    // DescarregarChangesAsync falha NA HORA com a mensagem neutra (nao espera os 10 s)
    public void Descarga_SessaoCongelaNoMeio_FalhaNaHora()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-descarga-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = ""; Exception? erro = null; long ms = -1;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.Show();
                recovery = await AbrirEditar(j, docx, " primeiro");
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" pendente"));
                using (new FileStream(j.SessaoAtual!.ArquivoChanges, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var relogio = System.Diagnostics.Stopwatch.StartNew();
                    try { await j.DescarregarChangesAsync(); } catch (Exception ex) { erro = ex; }
                    ms = relogio.ElapsedMilliseconds;
                }
                for (var i = 0; i < 50 && j.IsVisible; i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            Assert.NotNull(erro);
            Assert.Equal(JanelaEditor.MensagemSessaoCongelada, Assert.IsType<InvalidOperationException>(erro).Message);
            Assert.DoesNotContain("a recuperação falhou", JanelaEditor.MensagemSessaoCongelada);
            Assert.True(ms < 5000, $"a descarga esperou {ms} ms (devia falhar na hora)");
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableTheory] // item 2: Ctrl+S e fechar LOGO EM SEGUIDA — o Closed mantem a pasta (salvamento em curso), o salvamento
                      // termina com a janela fechada, grava o arquivo, e a pasta SAI: nada aparece como recuperavel. Com
                      // `falharEndSave` o EndSave ao editor LANCA (WebView2 ja' descartada): o registro da gravacao (Salvo,
                      // carimbo, limpeza da pasta) tem de rodar do mesmo jeito.
    [InlineData(false)]
    [InlineData(true)]
    public void Salvar_EFecharLogoEmSeguida_ArquivoGravadoEPastaApagada(bool falharEndSave)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-salvafecha-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = ""; var salvos = new List<string>(); var fechou = false; var log = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.Salvo += c => salvos.Add(c);
                j.Fechado += () => fechou = true;
                j.Show();
                log = j.ArquivoDeLog;
                recovery = await AbrirEditar(j, docx, " salvo-e-fechado");
                j.FalharEndSaveParaTeste = falharEndSave;
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"))).Trim('"'));
                j.Close();
                Assert.True(fechou);
                for (var i = 0; i < 300 && (salvos.Count == 0 || Directory.Exists(recovery)); i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            Assert.Equal([docx], salvos);
            Assert.Contains("salvo-e-fechado", mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(docx));
            Assert.False(Directory.Exists(recovery), "a pasta de um documento salvo sobrou depois do fechamento");
            Assert.DoesNotContain(recovery, Recuperaveis.Listar(SessaoDocumento.RaizRecovery, "mtexto").Select(d => d.Pasta));
            Assert.Contains("RECUPERACAO janela fechada durante o salvamento, que deu certo", File.ReadAllText(log));
        }
        finally { Limpar(recovery, pasta); }
    }

    // --- Fix round 4 ------------------------------------------------------------------------------------------

    [SkippableFact] // item 2: a ponte grava DEPOIS de o x2t ler a pasta (costura entre a conversao e o EndSave) — um change
                    // desfeito e outro do MESMO tamanho, que o tamanho do changes0.json nao via — com a janela ja' fechada:
                    // a pasta FICA (o arquivo gravado nao tem esse change)
    public void Salvar_EFechar_GravacaoDepoisDaLeitura_MantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-geracao-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = ""; var salvos = new List<string>(); long tamanhoAntes = -1, tamanhoDepois = -2;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.Salvo += c => salvos.Add(c);
                j.Show();
                recovery = await AbrirEditar(j, docx, " geracao");
                var s = j.SessaoAtual!;
                var ponte = j.PonteParaTeste!;
                j.DepoisDaConversaoParaTeste = () =>
                {
                    // Troca o ULTIMO change por outro do mesmo tamanho (desfazer + digitar), pela ponte da sessao.
                    var texto = File.ReadAllText(s.ArquivoChanges);
                    var ultimo = texto.TrimEnd(',').Split("\",\"").Last().Trim('"');
                    tamanhoAntes = new FileInfo(s.ArquivoChanges).Length;
                    Assert.Equal("ok", ponte.SalvarChanges(new string('x', ultimo.Length), s.IndiceChanges - 1, 1));
                    tamanhoDepois = new FileInfo(s.ArquivoChanges).Length;
                    return Task.CompletedTask;
                };
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"))).Trim('"'));
                j.Close();
                for (var i = 0; i < 300 && salvos.Count == 0; i++) await Task.Delay(100);
                await Task.Delay(500);
            }, TimeSpan.FromSeconds(180));
            Assert.Equal([docx], salvos);
            Assert.Equal(tamanhoAntes, tamanhoDepois);          // o cenario que o tamanho nao via
            Assert.True(Directory.Exists(recovery), "a pasta saiu com um change gravado depois da leitura do x2t");
            Assert.True(Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // item 1: um <doc>.moffice-tmp VELHO (e somente-leitura) de um salvamento que morreu no meio nao impede o
                    // proximo: sai antes da conversao, o arquivo e' gravado e nenhum temporario sobra
    public void Salvar_ComTemporarioVelhoSomenteLeitura_GravaENaoSobraTemporario()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-tmpvelho-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var tmp = docx + ".moffice-tmp";
        File.WriteAllText(tmp, "lixo de um salvamento que morreu");
        File.SetAttributes(tmp, FileAttributes.ReadOnly);
        string recovery = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.Show();
                recovery = await AbrirEditar(j, docx, " depois-do-tmp");
                Assert.True((await j.SalvarAsync()).Gravou);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            Assert.Contains("depois-do-tmp", mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(docx));
            Assert.False(File.Exists(tmp));
        }
        finally
        {
            try { if (File.Exists(tmp)) File.SetAttributes(tmp, FileAttributes.Normal); } catch (IOException) { }
            Limpar(recovery, pasta);
        }
    }

    // --- Fix round 5 ------------------------------------------------------------------------------------------

    [SkippableTheory] // o usuario DIGITA enquanto o salvamento converte (janela aberta; costura pausa depois de o x2t gravar o
                      // arquivo): no fim o documento continua MODIFICADO, fechar PERGUNTA, e o arquivo salvo NAO tem o
                      // marcador. "digitando" = edicao real no editor; "gravacaoDireta" = so' a ponte grava (a geracao de
                      // gravacao, sem aviso de modificado); controle "semEdicao" = nada muda, o documento fica salvo.
    [InlineData("digitando")]
    [InlineData("gravacaoDireta")]
    [InlineData("semEdicao")]
    public void Salvar_EdicaoDuranteOSalvamento_ContinuaModificado(string modo)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("mOffice-testes-durante-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        string recovery = ""; bool? modificadoNoFim = null; var perguntou = false; var fechou = false;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.NaoSalvar); };
                j.Fechado += () => fechou = true;
                j.Show();
                recovery = await AbrirEditar(j, docx, " salvo-antes");
                var s = j.SessaoAtual!;
                var ponte = j.PonteParaTeste!;
                j.DepoisDaConversaoParaTeste = async () =>
                {
                    if (modo == "digitando")
                    {
                        await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" marcador-durante"));
                        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                    }
                    else if (modo == "gravacaoDireta")
                        Assert.Equal("ok", ponte.SalvarChanges("marcador-durante", -1, 1));
                };
                Assert.True((await j.SalvarAsync()).Gravou);
                await Task.Delay(300);   // avisos do editor enfileirados no Dispatcher
                modificadoNoFim = j.Modificado;
                j.Close();
                for (var i = 0; i < 50 && !fechou; i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            var texto = mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(docx);
            Assert.Contains("salvo-antes", texto);
            Assert.DoesNotContain("marcador-durante", texto);
            var editou = modo != "semEdicao";
            Assert.Equal(editou, modificadoNoFim);
            Assert.Equal(editou, perguntou);
            Assert.True(fechou);
        }
        finally { Limpar(recovery, pasta); }
    }
}
