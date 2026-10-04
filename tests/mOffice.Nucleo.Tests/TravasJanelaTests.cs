using System.IO.Compression;
using System.Security.Cryptography;
using mOffice.Conversor;
using mOffice.Conversor.Tests;
using mOffice.Nucleo.Travas;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 6, Task 5 (SO-02/MvP-05, SO-14 parte 1, MvP-12 minimo) pela JANELA real: o documento aberto trava o original e cria o
/// <c>~$</c> com o autor; arquivo em uso/somente leitura abre somente leitura (restricao ReadOnly no editor, Ctrl+S vira Salvar como);
/// o salvar por cima confere o carimbo e pergunta no conflito; documento assinado avisa. O "outro usuario" e' sempre um FileStream
/// deste processo — nada de rede, nada de Office real. Em sessao remota (RDP) estes testes sao ignorados (EditorReal).
/// </summary>
[Collection(JanelaReal.Nome)]
public class TravasJanelaTests
{
    private const string Autor = "Dra. Teste Trava";
    private const string SalvarPeloEditor = "api.asc_Save(false, false, true); return \"ok\";";

    /// <summary>Cabecalho do web-apps de texto em "Somente leitura" (Common.Views.Header.readOnly, que poe "(Somente leitura)" no titulo).</summary>
    private static readonly string CabecalhoSomenteLeitura = EditorReal.NoEditor(
        "var h = w.DE.getController('Viewport').getView('Common.Views.Header'); return String(!!(h && h.readOnly));");

    private static string NovaPasta(string prefixo) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, prefixo + Guid.NewGuid().ToString("N"))).FullName;

    private static string Sha(string arquivo)
    {
        using var f = new FileStream(arquivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Convert.ToHexString(SHA256.HashData(f));
    }

    private static string Docx(string pasta, string nome, string texto)
    {
        var gerado = FabricaDocumentos.Docx(pasta, texto);
        var destino = Path.Combine(pasta, nome);
        File.Move(gerado, destino);
        return destino;
    }

    private static bool PodeGravar(string arquivo)
    {
        try { new FileStream(arquivo, FileMode.Open, FileAccess.Write, FileShare.ReadWrite).Dispose(); return true; }
        catch (IOException) { return false; }
    }

    private static async Task<bool> Aguardar(Func<bool> condicao, int tentativas = 150)
    {
        for (var i = 0; i < tentativas && !condicao(); i++) await Task.Delay(100);
        return condicao();
    }

    private static JanelaEditor Janela() =>
        EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { Autor = new Autor("uteste", Autor) });

    private static async Task AbrirEEsperar(JanelaEditor j, string arquivo)
    {
        j.Show();
        j.Activate();
        await j.AbrirAsync(arquivo);
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
    }

    private static async Task EditarESalvarPeloEditor(JanelaEditor j, string marca)
    {
        Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(marca))).Trim('"'));
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
        Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor(SalvarPeloEditor))).Trim('"'));
    }

    [SkippableFact] // documento aberto: outro processo nao consegue gravar no original e o ~$ tem o nome do autor; o Ctrl+S grava COM a trava
                    // (Substituir passa) e o arquivo novo continua travado; fechar libera e apaga o ~$
    public void Aberto_TravaOriginalECriaDono()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-trava-");
        var original = Docx(pasta, "Contrato social.docx", "texto original");
        var dono = ArquivoDeDono.CaminhoPara(original);
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                var salvos = new List<string>();
                j.Salvo += salvos.Add;
                await AbrirEEsperar(j, original);
                Assert.False(PodeGravar(original));                       // o "Word do colega" nao grava
                Assert.True(File.GetAttributes(dono).HasFlag(FileAttributes.Hidden));
                Assert.Equal(Autor, ArquivoDeDono.LerNome(dono));
                Assert.Equal(EstadoAcesso.EmUso, AcessoAoArquivo.Sondar(original));   // outra janela/pessoa veria "em uso"
                Assert.False(j.SomenteLeitura);

                await EditarESalvarPeloEditor(j, "gravado com a trava");
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Ctrl+S nao gravou: " + j.UltimoErro);
                Assert.Null(j.UltimoErro);
                Assert.Contains("gravado com a trava", FabricaDocumentos.TextoDocx(Copia(original, pasta)));
                Assert.False(PodeGravar(original));                       // continua travado (o arquivo NOVO)
                Assert.Equal(Autor, ArquivoDeDono.LerNome(dono));
                Assert.Empty(Directory.GetFiles(pasta, "*.moffice-bak*"));
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.True(PodeGravar(original));
            Assert.False(File.Exists(dono));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Le' o docx travado por uma copia (o OpenXml SDK abre com FileShare.Read, que a trava recusaria).</summary>
    private static string Copia(string arquivo, string pasta)
    {
        var c = Path.Combine(pasta, "leitura-" + Guid.NewGuid().ToString("N") + ".docx");
        using (var de = new FileStream(arquivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (var para = File.Create(c)) de.CopyTo(para);
        return c;
    }

    [SkippableFact] // arquivo com atributo somente leitura: abre sem perguntar, faixa e titulo de somente leitura, cabecalho do editor em
                    // "Somente leitura" (restricao ReadOnly), Ctrl+S vira Salvar como, gravar no original e' recusado; o Salvar como trava o novo
    public void AtributoSomenteLeitura_AbreSomenteLeitura()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-somente-leitura-");
        var original = Docx(pasta, "Parecer.docx", "texto protegido");
        File.SetAttributes(original, FileAttributes.ReadOnly);
        var shaAntes = Sha(original);
        var copia = Path.Combine(pasta, "Parecer revisado.docx");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                var perguntasEmUso = 0;
                j.PerguntarEmUso = (_, _, _) => { perguntasEmUso++; return Task.FromResult(true); };
                var destinos = new List<FormatoDocumento>();
                j.EscolherDestinoParaTeste = (_, sugerido) => { destinos.Add(sugerido); return new DestinoSalvar(copia, FormatoDocumento.Docx); };
                var salvos = new List<string>();
                j.Salvo += salvos.Add;
                await AbrirEEsperar(j, original);
                Assert.Equal(0, perguntasEmUso);
                Assert.True(j.SomenteLeitura);
                Assert.True(j.SessaoAtual!.ExigeSalvarComo);
                Assert.Contains("(Somente leitura)", j.Title);
                Assert.Equal(JanelaEditor.MensagemSomenteLeitura(EstadoAcesso.SomenteLeituraAtributo), j.AvisosDeAbertura[0]);
                Assert.Equal("true", await EditorReal.Esperar(j, CabecalhoSomenteLeitura, "true", 100));
                Assert.False(File.Exists(ArquivoDeDono.CaminhoPara(original)));   // somente leitura nao cria ~$

                // A API do host tentando gravar no original: recusado, nada tocado.
                var recusa = await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarAsync());
                Assert.Equal(JanelaEditor.MensagemSomenteLeituraRecusado, recusa.Message);

                // Ctrl+S: o editor continua editavel e o salvar pergunta o destino.
                await EditarESalvarPeloEditor(j, "alteracao em copia");
                Assert.True(await Aguardar(() => salvos.Count > 0), "o Ctrl+S nao gravou a copia: " + j.UltimoErro);
                Assert.Equal([FormatoDocumento.Docx], destinos);
                Assert.Equal([copia], salvos);
                Assert.False(j.SomenteLeitura);                                    // o arquivo novo e' do usuario
                Assert.DoesNotContain("(Somente leitura)", j.Title);
                Assert.False(PodeGravar(copia));                                   // e fica travado
                Assert.Equal(Autor, ArquivoDeDono.LerNome(ArquivoDeDono.CaminhoPara(copia)));
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaAntes, Sha(original));
            Assert.Contains("alteracao em copia", FabricaDocumentos.TextoDocx(copia));
        }
        finally
        {
            try { File.SetAttributes(original, FileAttributes.Normal); } catch { }
            try { Directory.Delete(pasta, true); } catch { }
        }
    }

    [SkippableFact] // arquivo aberto para escrita por outro FileStream (o "Word do colega") e com o ~$ dele: PerguntarEmUso recebe o nome do
                    // arquivo e o dono; true abre somente leitura a partir da copia; o original nao e' tocado nem travado por nos
    public void EmUso_PerguntaEAbreSomenteLeitura()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-em-uso-");
        var original = Docx(pasta, "Peticao inicial.docx", "texto do colega");
        var dono = ArquivoDeDono.CaminhoPara(original);
        File.WriteAllBytes(dono, ArquivoDeDono.Montar("Dr. Colega", AppDoDono.Word));
        var shaAntes = Sha(original);
        try
        {
            using (var colega = new FileStream(original, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            using (var donoDoColega = new FileStream(dono, FileMode.Open, FileAccess.ReadWrite, FileShare.Read | FileShare.Delete))
            {
                JanelaSta.Executar(async () =>
                {
                    var j = Janela();
                    var perguntas = new List<(string Arquivo, string? Dono)>();
                    j.PerguntarEmUso = (_, a, d) => { perguntas.Add((a, d)); return Task.FromResult(true); };
                    await AbrirEEsperar(j, original);
                    Assert.Equal([("Peticao inicial.docx", (string?)"Dr. Colega")], perguntas);
                    Assert.True(j.SomenteLeitura);
                    Assert.Equal(EstadoAcesso.EmUso, j.SessaoAtual!.MotivoSomenteLeitura);
                    Assert.Equal(original, j.SessaoAtual.ArquivoOrigem);
                    Assert.Equal(JanelaEditor.MensagemSomenteLeitura(EstadoAcesso.EmUso, "Dr. Colega"), j.AvisosDeAbertura[0]);
                    Assert.Contains("em uso por Dr. Colega", j.AvisosDeAbertura[0]);
                    Assert.Equal("true", await EditorReal.Esperar(j, CabecalhoSomenteLeitura, "true", 100));
                    Assert.Null(j.SessaoAtual.TravaDoOriginal);
                    // a copia de leitura vive na pasta da sessao
                    Assert.StartsWith(j.SessaoAtual.PastaRecovery, j.SessaoAtual.CopiaDeLeitura);
                    j.Close();
                }, TimeSpan.FromSeconds(240));
                Assert.Equal("Dr. Colega", ArquivoDeDono.LerNome(dono));   // o ~$ do colega nao foi mexido
            }
            Assert.Equal(shaAntes, Sha(original));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // EmUso e o usuario cancela: a abertura lanca AberturaEmUsoCanceladaException, nada abre e o original nao e' tocado
    public void EmUso_Cancelar_NaoAbre()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-em-uso-cancelar-");
        var original = Docx(pasta, "Peticao.docx", "texto do colega");
        var shaAntes = Sha(original);
        try
        {
            using (new FileStream(original, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                JanelaSta.Executar(async () =>
                {
                    var j = Janela();
                    var perguntas = new List<(string, string?)>();
                    j.PerguntarEmUso = (_, a, d) => { perguntas.Add((a, d)); return Task.FromResult(false); };
                    j.Show();
                    await Assert.ThrowsAsync<JanelaEditor.AberturaEmUsoCanceladaException>(() => j.AbrirAsync(original));
                    Assert.Equal([("Peticao.docx", (string?)null)], perguntas);   // sem ~$: dono desconhecido
                    Assert.Null(j.SessaoAtual);
                    j.Close();
                }, TimeSpan.FromSeconds(120));
            }
            Assert.Equal(shaAntes, Sha(original));
            Assert.False(File.Exists(ArquivoDeDono.CaminhoPara(original)));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Simula a trava indisponivel (rede que nao segura handle, ou o original gravado antes de a trava existir) e grava o
    /// arquivo "do outro" por cima do original.</summary>
    private static string GravarPorFora(JanelaEditor j, string original, string pasta)
    {
        j.SessaoAtual!.TravaDoOriginal!.Soltar();
        var outro = Docx(NovaPastaDentro(pasta), "outro.docx", "versao do outro");
        File.Copy(outro, original, overwrite: true);
        File.SetLastWriteTimeUtc(original, DateTime.UtcNow.AddMinutes(2));
        return Sha(original);
    }

    private static string NovaPastaDentro(string pasta) => Directory.CreateDirectory(Path.Combine(pasta, Guid.NewGuid().ToString("N"))).FullName;

    [SkippableFact] // o original muda por fora depois da abertura: o Salvar chama PerguntarConflito ANTES de gravar; Cancelar nao grava
    public void Conflito_PerguntaAntesDeGravar_CancelarPreservaOArquivoDoOutro()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-conflito-");
        var original = Docx(pasta, "Contrato.docx", "texto original");
        string shaDoOutro = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                var conflitos = 0;
                j.PerguntarConflito = _ => { conflitos++; return Task.FromResult(DecisaoConflito.Cancelar); };
                await AbrirEEsperar(j, original);
                shaDoOutro = GravarPorFora(j, original, pasta);
                await EditarESalvarPeloEditor(j, "nao pode gravar");
                Assert.True(await Aguardar(() => j.UltimoErro is not null), "o salvar nao parou no conflito");
                Assert.Equal(1, conflitos);
                Assert.StartsWith(JanelaEditor.MensagemConflito, j.UltimoErro);
                Assert.True(await Aguardar(() => j.Modificado), "nada foi gravado: o documento tem de continuar modificado");
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaDoOutro, Sha(original));
            Assert.Contains("versao do outro", FabricaDocumentos.TextoDocx(original));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // Substituir grava por cima e renova o carimbo e a trava (um 2o salvar nao pergunta de novo)
    public void Conflito_Substituir_GravaERenovaCarimbo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-conflito-substituir-");
        var original = Docx(pasta, "Contrato.docx", "texto original");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                var conflitos = 0;
                j.PerguntarConflito = _ => { conflitos++; return Task.FromResult(DecisaoConflito.Substituir); };
                var salvos = new List<string>();
                j.Salvo += salvos.Add;
                await AbrirEEsperar(j, original);
                GravarPorFora(j, original, pasta);
                await EditarESalvarPeloEditor(j, "minha versao");
                Assert.True(await Aguardar(() => salvos.Count == 1), "o Substituir nao gravou: " + j.UltimoErro);
                Assert.Equal(1, conflitos);
                Assert.False(PodeGravar(original));            // a trava voltou (sobre o arquivo gravado)
                await EditarESalvarPeloEditor(j, " e mais");
                Assert.True(await Aguardar(() => salvos.Count == 2), "o 2o salvar nao gravou: " + j.UltimoErro);
                Assert.Equal(1, conflitos);                    // carimbo renovado: nao perguntou de novo
                j.Close();
            }, TimeSpan.FromSeconds(240));
            var texto = FabricaDocumentos.TextoDocx(original);
            Assert.Contains("minha versao", texto);
            Assert.DoesNotContain("versao do outro", texto);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // "Salvar uma copia" no conflito: nada grava no original e o "Salvar como" abre em seguida (grava a copia)
    public void Conflito_SalvarCopia_AbreSalvarComo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-conflito-copia-");
        var original = Docx(pasta, "Contrato.docx", "texto original");
        var copia = Path.Combine(pasta, "Contrato (minha copia).docx");
        string shaDoOutro = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                j.PerguntarConflito = _ => Task.FromResult(DecisaoConflito.SalvarCopia);
                j.EscolherDestinoParaTeste = (_, _) => new DestinoSalvar(copia, FormatoDocumento.Docx);
                var salvos = new List<string>();
                j.Salvo += salvos.Add;
                await AbrirEEsperar(j, original);
                shaDoOutro = GravarPorFora(j, original, pasta);
                await EditarESalvarPeloEditor(j, "minha copia");
                Assert.True(await Aguardar(() => salvos.Count == 1, 300), "a copia nao foi gravada: " + j.UltimoErro);
                Assert.Equal([copia], salvos);
                Assert.Equal(copia, j.SessaoAtual!.ArquivoOrigem);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaDoOutro, Sha(original));
            Assert.Contains("minha copia", FabricaDocumentos.TextoDocx(copia));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // docx com parte _xmlsignatures/ (sintetica) mostra o aviso de assinatura nos AvisosDeAbertura
    public void Assinado_AvisaAoAbrir()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-assinado-");
        var original = Docx(pasta, "Procuracao.docx", "documento assinado");
        using (var z = ZipFile.Open(original, ZipArchiveMode.Update))
        using (var w = new StreamWriter(z.CreateEntry("_xmlsignatures/sig1.xml").Open()))
            w.Write("<?xml version=\"1.0\"?><Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\"/>");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                await AbrirEEsperar(j, original);
                Assert.True(j.AssinaturaDigital);
                Assert.Equal(JanelaEditor.AvisoAssinaturaDigital, j.AvisosDeAbertura[0]);
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // sobra de queda: um ~$ de "Outro" sem ninguem segurando nao pergunta nada; a abertura toma o ~$ com o nosso nome
    public void Queda_DonoVelho_NaoBloqueia()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-dono-velho-");
        var original = Docx(pasta, "Contrato.docx", "texto");
        var dono = ArquivoDeDono.CaminhoPara(original);
        File.WriteAllBytes(dono, ArquivoDeDono.Montar("Outro", AppDoDono.Word));
        File.SetAttributes(dono, FileAttributes.Hidden);
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                var perguntas = 0;
                j.PerguntarEmUso = (_, _, _) => { perguntas++; return Task.FromResult(false); };
                await AbrirEEsperar(j, original);
                Assert.Equal(0, perguntas);
                Assert.False(j.SomenteLeitura);
                Assert.Equal(Autor, ArquivoDeDono.LerNome(dono));
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.False(File.Exists(dono));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // caminhos que apagam (backlog 56): fechar MANTENDO a pasta de recuperacao (o estado de uma queda/logoff) solta o original
                    // e apaga o ~$ — a pasta com o trabalho continua la'
    public void FecharMantendoRecuperacao_SoltaOOriginalEMantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-trava-recuperacao-");
        var original = Docx(pasta, "Contrato.docx", "texto");
        string recuperacao = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                await AbrirEEsperar(j, original);
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("trabalho nao salvo"))).Trim('"'));
                Assert.True(await Aguardar(() => j.Modificado));
                recuperacao = j.SessaoAtual!.PastaRecovery;
                j.ManterRecoveryAoFecharParaTeste = true;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.True(PodeGravar(original));
            Assert.False(File.Exists(ArquivoDeDono.CaminhoPara(original)));
            Assert.True(Directory.Exists(recuperacao));
        }
        finally
        {
            try { Directory.Delete(pasta, true); } catch { }
            try { if (recuperacao.Length > 0) Directory.Delete(recuperacao, true); } catch { }
        }
    }

    // --- Fix 1 da revisao da Task 5 ---------------------------------------------------------------------------------------------

    [SkippableFact] // IMPORTANTE: sessao RECUPERADA de um original Livre (nao desvinculado) trava o original e cria o ~$ como uma abertura
                    // normal, e o carimbo e' tirado na recuperacao: gravacao por fora (com a trava indisponivel) -> o Ctrl+S pergunta o conflito
    public void Recuperada_TravaEConfereConflito()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-trava-recuperada-");
        var original = Docx(pasta, "Contrato.docx", "texto original");
        var dono = ArquivoDeDono.CaminhoPara(original);
        string recuperacao = "";
        string shaDoOutro = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                await AbrirEEsperar(j, original);
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord("trabalho recuperado"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                await j.DescarregarChangesAsync();
                recuperacao = j.SessaoAtual!.PastaRecovery;
                j.ManterRecoveryAoFecharParaTeste = true;   // "queda": a pasta fica, a trava do original sai
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.True(PodeGravar(original));
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                var conflitos = 0;
                j.PerguntarConflito = _ => { conflitos++; return Task.FromResult(DecisaoConflito.Cancelar); };
                j.Show();
                await j.RecuperarAsync(recuperacao);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.True(j.SessaoAtual!.Recuperada);
                Assert.False(j.SessaoAtual.Desvinculada);
                Assert.False(PodeGravar(original));                       // travado como numa abertura normal
                Assert.Equal(Autor, ArquivoDeDono.LerNome(dono));
                Assert.NotNull(j.SessaoAtual.Carimbo);
                shaDoOutro = GravarPorFora(j, original, pasta);
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor(SalvarPeloEditor))).Trim('"'));
                Assert.True(await Aguardar(() => j.UltimoErro is not null), "o salvar da sessao recuperada nao parou no conflito");
                Assert.Equal(1, conflitos);
                Assert.StartsWith(JanelaEditor.MensagemConflito, j.UltimoErro);
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaDoOutro, Sha(original));
        }
        finally
        {
            try { Directory.Delete(pasta, true); } catch { }
            try { if (recuperacao.Length > 0 && Directory.Exists(recuperacao)) Directory.Delete(recuperacao, true); } catch { }
        }
    }

    [SkippableFact] // MINOR 3: sem conseguir reservar o original, a faixa avisa (e o carimbo continua: o conflito ainda e' detectado)
    public void SemTrava_AvisaNaFaixaEContinuaConferindo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-sem-trava-");
        var original = Docx(pasta, "Contrato.docx", "texto original");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                j.FalharTravaParaTeste = true;
                await AbrirEEsperar(j, original);
                Assert.Null(j.SessaoAtual!.TravaDoOriginal);
                Assert.Equal(JanelaEditor.AvisoSemReserva, j.AvisosDeAbertura[0]);
                Assert.NotNull(j.SessaoAtual.Carimbo);
                Assert.True(PodeGravar(original));
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // revisao final do Plano 6 (minor da Task 5): o Retomar falha no aposTroca E de novo em TravarArquivoGravado — a faixa
                    // "sem reserva" aparece UMA vez por salvamento (antes, duas)
    public void RetomarFalhaNoSalvar_FaixaSemReservaUmaVez()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-retomar-falha-");
        var original = Docx(pasta, "Contrato.docx", "texto original");
        var avisos = new List<string>();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = Janela();
                j.AvisoAoUsuario += a => avisos.Add(a);
                await AbrirEEsperar(j, original);
                var trava = j.SessaoAtual!.TravaDoOriginal;
                Assert.NotNull(trava);
                trava.FalharRetomarParaTeste = true;
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" editado"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                Assert.True((await j.SalvarAsync()).Gravou);
                await Task.Delay(500);
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(1, avisos.Count(a => a == JanelaEditor.AvisoSemReserva));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
