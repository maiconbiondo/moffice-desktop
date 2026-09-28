using mOffice.Conversor;
using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class FecharComAlteracoesTests
{
    private static JanelaEditor Abrir(Func<JanelaEditor, Task<DecisaoFechar>> pergunta)
    {
        // PrepararParaTeste primeiro (reserva sem modal — nunca um MessageBox real preso na thread STA),
        // e SO' DEPOIS a pergunta deste teste, que e' o que cada caso quer exercitar.
        var j = EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { NomeApp = "mTexto" });
        j.PerguntarAoFechar = pergunta;
        j.Show();
        return j;
    }

    [SkippableFact] // sem alteracoes: fecha sem perguntar; evento Fechado dispara; recovery apagada
    public void Fechar_SemAlteracoes_NaoPergunta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        JanelaSta.Executar(async () =>
        {
            var perguntou = false; var fechou = false;
            var j = Abrir(_ => { perguntou = true; return Task.FromResult(DecisaoFechar.Cancelar); });
            j.Fechado += () => fechou = true;
            await j.AbrirAsync(doc);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            var recovery = j.SessaoAtual!.PastaRecovery;
            j.Close();
            Assert.False(perguntou); Assert.True(fechou); Assert.False(Directory.Exists(recovery));
        });
    }

    [SkippableFact] // com alteracoes: Cancelar mantem a janela aberta; NaoSalvar fecha sem gravar
    public void Fechar_ComAlteracoes_CancelarSegura_NaoSalvarFecha()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        var antes = File.GetLastWriteTimeUtc(doc);
        JanelaSta.Executar(async () =>
        {
            var decisao = DecisaoFechar.Cancelar; var vezes = 0;
            var j = Abrir(_ => { vezes++; return Task.FromResult(decisao); });
            await j.AbrirAsync(doc);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord("mudou"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
            j.Close();
            // O AoFechar decide de forma assincrona (e.Cancel = true + await da pergunta), entao a decisao
            // Cancelar so' chega ao contador depois de o Dispatcher rodar a continuacao.
            for (var i = 0; i < 50 && vezes == 0; i++) await Task.Delay(100);
            Assert.Equal(1, vezes); Assert.True(j.IsVisible);
            decisao = DecisaoFechar.NaoSalvar;
            j.Close();
            for (var i = 0; i < 50 && j.IsVisible; i++) await Task.Delay(100);
            Assert.Equal(2, vezes); Assert.False(j.IsVisible);
        });
        Assert.Equal(antes, File.GetLastWriteTimeUtc(doc));
    }

    [SkippableFact] // com alteracoes: Salvar grava no arquivo de origem e fecha
    public void Fechar_ComAlteracoes_SalvarGravaEFecha()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        JanelaSta.Executar(async () =>
        {
            var j = Abrir(_ => Task.FromResult(DecisaoFechar.Salvar));
            await j.AbrirAsync(doc);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord("salvo ao fechar"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
            j.Close();
            for (var i = 0; i < 150 && j.IsVisible; i++) await Task.Delay(200);
            Assert.False(j.IsVisible);
        // 180s (achado 13 da revisao 1): o fechamento com "Salvar" espera o asc_Save + o x2t de volta para
        // DOCX depois da abertura — o padrao de 90s de JanelaSta.Executar e' apertado numa maquina fria.
        }, TimeSpan.FromSeconds(180));
        Assert.Contains("salvo ao fechar", FabricaDocumentos.TextoDocx(doc));
    }

    // -------------------------------------------------------------------------------------------------
    // Achados da revisão final do Plano 2. O diálogo de destino (`EscolherDestino`) é modal: num teste de
    // janela real ele travaria a thread STA, então estes testes substituem só a ESCOLHA, pela costura
    // `EscolherDestinoParaTeste` — filtro, gravação, eventos e caminho de fechamento continuam os do produto.
    // -------------------------------------------------------------------------------------------------

    [SkippableFact] // CRÍTICO: documento NOVO + "Salvar" ao fechar grava um documento de verdade, fecha e dispara Salvo;
                    // e o filtro oferecido nesse caminho NÃO tem PDF (escolher PDF exportaria e perderia o editável)
    public void Fechar_DocumentoNovo_Salvar_GravaDocumentoENuncaOferecePdf()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(DadosDeTeste.RaizDeTestes, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var destino = Path.Combine(pasta, "novo-ao-fechar.docx");
        var filtros = new List<string>();
        var salvos = new List<string>();
        var exportados = new List<string>();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Abrir(_ => Task.FromResult(DecisaoFechar.Salvar));
                j.EscolherDestinoParaTeste = (filtro, _) => { filtros.Add(filtro); return new DestinoSalvar(destino, FormatoDocumento.Docx); };
                j.Salvo += c => salvos.Add(c);
                j.Exportado += c => exportados.Add(c);
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord("novo salvo ao fechar"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                j.Close();
                for (var i = 0; i < 150 && j.IsVisible; i++) await Task.Delay(200);
                Assert.False(j.IsVisible);
                Assert.Null(j.UltimoErro);
            // 180s: abre o modelo vazio, edita, pede asc_Save e roda o x2t de volta para DOCX.
            }, TimeSpan.FromSeconds(180));

            // O documento EDITÁVEL existe e tem o texto (era exatamente isto que a exportação para PDF perdia).
            Assert.True(File.Exists(destino));
            Assert.Contains("novo salvo ao fechar", FabricaDocumentos.TextoDocx(destino));
            Assert.Equal(new[] { destino }, salvos);   // "salvar", não "exportar"
            Assert.Empty(exportados);

            // O segundo caso do achado, provado POR CONSTRUÇÃO: escolher PDF aqui é impossível porque o
            // diálogo deste caminho não oferece PDF nenhum.
            var filtro = Assert.Single(filtros);
            Assert.DoesNotContain("pdf", filtro, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("*.docx", filtro);
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [SkippableFact] // a gravação do "Salvar" ao fechar falha: a janela CONTINUA aberta, ErroAoFechar dispara e explica
    public void Fechar_SalvarFalha_JanelaContinuaAbertaEAvisa()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(DadosDeTeste.RaizDeTestes, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        // Destino impossível de gravar, de forma determinística: a "pasta" de destino é um ARQUIVO, então o
        // Directory.CreateDirectory de ConverterESinalizarAsync estoura com IOException.
        var bloqueio = Path.Combine(pasta, "bloqueio"); File.WriteAllText(bloqueio, "nao sou pasta");
        var destino = Path.Combine(bloqueio, "x.docx");
        var erros = new List<string>();
        var abortos = 0;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Abrir(_ => Task.FromResult(DecisaoFechar.Salvar));
                j.EscolherDestinoParaTeste = (_, _) => new DestinoSalvar(destino, FormatoDocumento.Docx);
                j.ErroAoFechar += m => erros.Add(m);
                j.FechamentoAbortado += () => abortos++; // Plano 5 (fix round 1): o ramo Falhar tambem avisa o abandono
                await j.NovoAsync(FormatoDocumento.Docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord("nao vai gravar"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                j.Close();
                for (var i = 0; i < 150 && erros.Count == 0; i++) await Task.Delay(200);
                // Não fecha com a gravação falhada: seria perder o documento do usuário em silêncio.
                Assert.True(j.IsVisible);
                Assert.NotNull(j.UltimoErro);
                // CRÍTICO (revisão 2): a alteração pendente tem de SOBREVIVER à gravação que falhou. O
                // `asc_Save` já mandou `onDocumentModifiedChanged(false)` antes do `LocalFileSave`
                // (protocolo do Marco 0, passo 6 antes do 7), então sem a restauração do núcleo o documento
                // fica marcado como salvo sem ter sido gravado — e ESTE segundo `Close()` não perguntaria
                // nada, fecharia direto e o `Closed` apagaria a recuperação.
                //
                // MEDIDO (revisão 2): esta asserção passa MESMO sem o `RestaurarModificadoSeHouveEdicao()`
                // do caminho de falha — porque, ao receber `EndSave(2)`, o próprio editor volta a mandar
                // `onDocumentModifiedChanged(true)` e o flag se recupera sozinho. Ou seja: aqui a restauração
                // do núcleo é defesa em profundidade, não a única coisa segurando o estado (só três lugares
                // escrevem `Modificado`: a restauração, o salvamento bem-sucedido e a ponte do editor). Quem
                // prova a restauração de fato é o passo 9b de `SalvarDocumentoTests`, no ramo de EXPORTAÇÃO,
                // onde o `EndSave(0)` não faz o editor voltar atrás e a mutação derruba o teste. Esta
                // asserção fica assim mesmo: é o ESTADO FINAL que importa para o usuário, e ela travaria uma
                // regressão futura em qualquer um dos dois lados.
                Assert.True(j.Modificado, "a alteração pendente sumiu depois de a gravação falhar");
                j.PerguntarAoFechar = _ => Task.FromResult(DecisaoFechar.NaoSalvar);
                j.Close();
                for (var i = 0; i < 50 && j.IsVisible; i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            var erro = Assert.Single(erros);
            Assert.StartsWith("Não foi possível salvar:", erro);
            Assert.Equal(1, abortos);
            Assert.False(File.Exists(destino));
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [SkippableFact] // a pergunta de fechar estoura E a reserva também: a janela continua aberta e diz por quê
                    // (sem a reserva, uma pergunta que falha sempre tornava a janela impossível de fechar)
    public void Fechar_PerguntaEReservaEstouram_JanelaContinuaAbertaEAvisa()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(DadosDeTeste.RaizDeTestes, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        var erros = new List<string>();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Abrir(_ => throw new InvalidOperationException("diálogo quebrado"));
                // A reserva de verdade é um MessageBox — travaria a thread STA do teste. Aqui ela também
                // falha, que é justamente o caminho "nem a reserva respondeu" sendo exercitado.
                j.PerguntaDeReserva = _ => throw new InvalidOperationException("reserva quebrada");
                j.ErroAoFechar += m => erros.Add(m);
                await j.AbrirAsync(doc);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord("mudou"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                j.Close();
                for (var i = 0; i < 50 && erros.Count == 0; i++) await Task.Delay(100);
                Assert.True(j.IsVisible);   // trata como Cancelar: nunca descarta a alteração do usuário
                // E continua fechável pelo caminho normal — a falha não deixou a janela travada.
                j.PerguntarAoFechar = _ => Task.FromResult(DecisaoFechar.NaoSalvar);
                j.Close();
                for (var i = 0; i < 50 && j.IsVisible; i++) await Task.Delay(100);
                Assert.False(j.IsVisible);
            });
            var erro = Assert.Single(erros);
            Assert.Contains("reserva quebrada", erro);
            Assert.StartsWith("Não foi possível perguntar", erro);
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [SkippableFact] // a pergunta de fechar estoura mas a RESERVA responde: fecha normalmente, e ErroAoFechar
                    // NAO dispara — a reserva existe para o usuário não ficar preso, não para virar aviso de erro
    public void Fechar_PerguntaEstouraMasReservaResponde_FechaSemAvisarErro()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(DadosDeTeste.RaizDeTestes, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        var antes = File.GetLastWriteTimeUtc(doc);
        var erros = new List<string>();
        var reservas = 0;
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Abrir(_ => throw new InvalidOperationException("diálogo quebrado"));
                // A reserva de verdade é o MessageBox padrão — modal, travaria a thread STA do teste. Aqui
                // ela responde NaoSalvar, que é o caminho de SUCESSO da reserva.
                j.PerguntaDeReserva = _ => { reservas++; return Task.FromResult(DecisaoFechar.NaoSalvar); };
                j.ErroAoFechar += m => erros.Add(m);
                await j.AbrirAsync(doc);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord("mudou"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
                j.Close();
                for (var i = 0; i < 50 && j.IsVisible; i++) await Task.Delay(100);
                Assert.False(j.IsVisible);      // a reserva respondeu: a janela fecha
            });
            Assert.Equal(1, reservas);
            Assert.Empty(erros);                // reserva que FUNCIONA não é erro para o usuário
            // DENTRO do try: o `finally` apaga a pasta, e `File.GetLastWriteTimeUtc` de arquivo inexistente
            // não lança — devolve 1601-01-01, que faria esta asserção falhar por um motivo que nada tem a
            // ver com o que ela testa.
            Assert.Equal(antes, File.GetLastWriteTimeUtc(doc));  // NaoSalvar: o arquivo não foi tocado
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [SkippableFact] // ModificadoAlterado dispara true na edicao e false depois de salvar
    public void ModificadoAlterado_DisparaNaEdicaoENoSalvar()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        var doc = FabricaDocumentos.Docx(pasta, "abc");
        JanelaSta.Executar(async () =>
        {
            var j = Abrir(_ => Task.FromResult(DecisaoFechar.NaoSalvar));
            var valores = new List<bool>();
            j.ModificadoAlterado += v => valores.Add(v);
            await j.AbrirAsync(doc);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord("x"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));
            await j.SalvarAsync();
            Assert.Equal(new[] { true, false }, valores);
            j.Close();
        });
    }

    [SkippableFact] // Plano 5, Task 2 (fix round 1): todo fechamento ABANDONADO avisa (FechamentoAbortado) — "Salvar" com o
                    // "Salvar como" cancelado e "Cancelar" na pergunta; o fechamento que fecha, nao
    public void Fechar_Abandonado_DisparaFechamentoAbortado()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var decisao = DecisaoFechar.Salvar; var escolhas = 0; var abortos = 0;
            var j = Abrir(_ => Task.FromResult(decisao));
            j.EscolherDestinoParaTeste = (_, _) => { escolhas++; return null; }; // o usuario cancela o "Salvar como"
            j.FechamentoAbortado += () => abortos++;
            await j.NovoAsync(FormatoDocumento.Docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord("abandonado"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));

            j.Close(); // "Salvar" -> dialogo de destino cancelado
            for (var i = 0; i < 50 && abortos == 0; i++) await Task.Delay(100);
            Assert.Equal(1, escolhas); Assert.Equal(1, abortos); Assert.True(j.IsVisible);

            decisao = DecisaoFechar.Cancelar;
            j.Close(); // "Cancelar" na pergunta
            for (var i = 0; i < 50 && abortos == 1; i++) await Task.Delay(100);
            Assert.Equal(2, abortos); Assert.True(j.IsVisible);

            decisao = DecisaoFechar.NaoSalvar;
            j.Close(); // fecha de verdade: nao e' abandono
            for (var i = 0; i < 50 && j.IsVisible; i++) await Task.Delay(100);
            Assert.False(j.IsVisible);
            Assert.Equal(2, abortos);
        }, TimeSpan.FromSeconds(180));
    }
}
