using mOffice.Conversor;
using mOffice.Nucleo;
using mOffice.Nucleo.Recuperacao;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>Recuperacao apos queda (Task 2 do Plano 3) — as pecas puras: metadados da sessao, trava da pasta,
/// leitura do changes0.json no formato do editor e o truncamento RELATIVO a esta sessao.</summary>
public class RecuperacaoTests
{
    private static string Pasta() { var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }

    [Fact] // metadados vao e voltam; arquivo corrompido ou ausente = null
    public void Metadados_GravarELer()
    {
        var p = Pasta();
        new MetadadosSessao(@"C:\x\peticao.docx", (int)FormatoDocumento.Docx, "peticao.docx", false, "mtexto", 2).Gravar(p);
        Assert.Equal(new MetadadosSessao(@"C:\x\peticao.docx", 65, "peticao.docx", false, "mtexto", 2), MetadadosSessao.Ler(p));
        File.WriteAllText(Path.Combine(p, "sessao.json"), "{quebrado");
        Assert.Null(MetadadosSessao.Ler(p));
        Assert.Null(MetadadosSessao.Ler(Pasta()));
    }

    [Fact] // changes0.json no formato do editor ("c",) vira array JS com a contagem certa
    public void Changes_ViramArrayComContagem()
    {
        var p = Pasta(); Directory.CreateDirectory(Path.Combine(p, "changes"));
        var a = Path.Combine(p, "changes", "changes0.json");
        Assert.Equal(new LeituraChanges("[]", 0, false, 0), ChangesGravados.Ler(a));
        File.WriteAllText(a, "\"12;AAA\",\"13;BBB\",");
        Assert.Equal(new LeituraChanges("[\"12;AAA\",\"13;BBB\"]", 2, false, 0), ChangesGravados.Ler(a));
    }

    private static string ArquivoChanges(byte[] conteudo)
    {
        var p = Pasta(); Directory.CreateDirectory(Path.Combine(p, "changes"));
        var a = Path.Combine(p, "changes", "changes0.json");
        File.WriteAllBytes(a, conteudo);
        return a;
    }

    private static byte[] Utf8(string s) => System.Text.Encoding.UTF8.GetBytes(s);

    [Fact] // queda no meio da gravacao: cauda de zeros (o que o NTFS deixa num arquivo estendido e nao escrito)
    public void Changes_CaudaDeZeros_MantemOPrefixoValido()
    {
        var a = ArquivoChanges([.. Utf8("\"12;AAA\",\"13;BBB\","), .. new byte[37]]);
        var antes = File.ReadAllBytes(a);
        Assert.Equal(new LeituraChanges("[\"12;AAA\",\"13;BBB\"]", 2, true, 37), ChangesGravados.Ler(a));
        Assert.Equal(antes, File.ReadAllBytes(a)); // Ler e' PURA: nao conserta nada
    }

    [Fact] // queda no meio de um change: a string pela metade (e o change sem a virgula final) nao contam
    public void Changes_CaudaCortadaNoMeioDaString_MantemOPrefixoValido()
    {
        var a = ArquivoChanges(Utf8("\"12;AAA\",\"13;BBB\",\"14;CC"));
        Assert.Equal(new LeituraChanges("[\"12;AAA\",\"13;BBB\"]", 2, true, 6), ChangesGravados.Ler(a));
        var b = ArquivoChanges(Utf8("\"12;AAA\",\"13;BBB\""));
        Assert.Equal(new LeituraChanges("[\"12;AAA\"]", 1, true, 8), ChangesGravados.Ler(b));
    }

    [Fact] // lixo do comeco ao fim: nada aproveitavel, mas MARCADO como descartado (nao "vazio" silencioso)
    public void Changes_ArquivoTodoLixo_ZeroChangesMarcadoComoDescartado()
    {
        var a = ArquivoChanges(Utf8("{isto nao e' um changes0.json]"));
        var l = ChangesGravados.Ler(a);
        Assert.Equal(("[]", 0, true), (l.ArrayJs, l.Quantidade, l.Descartou));
        Assert.Equal(new FileInfo(a).Length, l.BytesDescartados);
    }

    [Fact] // Reparar: copia o original para .corrompido (sem sobrescrever copia anterior) e regrava so' o prefixo
    public void Changes_Reparar_GuardaOOriginalERegravaOPrefixo()
    {
        var original = Utf8("\"12;AAA\",\"13;BBB\",\"14;CC\0\0\0");
        var a = ArquivoChanges(original);
        File.WriteAllText(a + ".corrompido", "copia de uma recuperacao anterior");
        var r = ChangesGravados.Reparar(a);
        Assert.True(r.Reparou);
        Assert.Equal(2, r.Quantidade);
        Assert.Equal(9, r.BytesDescartados);
        Assert.Equal(a + ".corrompido.1", r.Copia);
        Assert.Equal("copia de uma recuperacao anterior", File.ReadAllText(a + ".corrompido")); // intocada
        Assert.Equal(original, File.ReadAllBytes(r.Copia!));
        Assert.Equal("\"12;AAA\",\"13;BBB\",", File.ReadAllText(a));
        Assert.Equal(new LeituraChanges("[\"12;AAA\",\"13;BBB\"]", 2, false, 0), ChangesGravados.Ler(a));
        Assert.False(File.Exists(a + ".tmp"));
        // Arquivo sao: Reparar nao mexe em nada (nem cria copia).
        var r2 = ChangesGravados.Reparar(a);
        Assert.Equal(new ReparoChanges(false, 2, 0, null), r2);
        Assert.False(File.Exists(a + ".corrompido.2"));
    }

    [Fact] // trava: enquanto a sessao vive, a pasta nao esta livre; depois de liberar, esta
    public void Trava_BloqueiaEnquantoViva()
    {
        var p = Pasta();
        using (var t = TravaSessao.Adquirir(p)) { Assert.NotNull(t); Assert.False(TravaSessao.Livre(p)); Assert.Null(TravaSessao.Adquirir(p)); }
        Assert.True(TravaSessao.Livre(p));
    }

    [Fact] // Livre so' sonda: pasta sem rec.lock (ou inexistente) esta' livre e continua sem rec.lock
    public void Trava_Livre_NaoCriaNada()
    {
        var p = Pasta();
        Assert.True(TravaSessao.Livre(p));
        Assert.Empty(Directory.EnumerateFileSystemEntries(p));
        var inexistente = Path.Combine(p, "nao-existe");
        Assert.True(TravaSessao.Livre(inexistente));
        Assert.False(Directory.Exists(inexistente));
        Assert.Null(TravaSessao.Adquirir(inexistente)); // Adquirir tambem nao cria a pasta
        Assert.False(Directory.Exists(inexistente));
    }

    [Fact] // ponte congelada (reaplicacao de uma sessao recuperada falhou): LocalFileSaveChanges nao grava nada
    public void SalvarChanges_PonteCongelada_NaoTocaNoArquivo()
    {
        var p = Pasta(); Directory.CreateDirectory(Path.Combine(p, "changes"));
        var s = new SessaoDocumento { ArquivoOrigem = "x.docx", Formato = FormatoDocumento.Docx, PastaRecovery = p, ChangesAbertos = 2, Recuperada = true };
        File.WriteAllText(s.ArquivoChanges, "\"r1\",\"r2\",");
        var antes = File.ReadAllBytes(s.ArquivoChanges);
        var ponte = new PonteDesktop(s);
        ponte.Congelar();
        ponte.SalvarChanges("n1", 0, 1);    // apendar
        ponte.SalvarChanges("n2", -1, 1);   // apendar sem delecao
        ponte.SalvarChanges("", 0, 0);      // "desfazer tudo" desta sessao
        Assert.Equal(antes, File.ReadAllBytes(s.ArquivoChanges));
        Assert.Equal(0, s.IndiceChanges);
    }

    [Fact] // truncamento relativo a esta sessao: com 2 changes recuperados, deleteIndex 1 mantem os 2 + 1 desta sessao
    public void SalvarChanges_TruncaDepoisDosRecuperados()
    {
        var p = Pasta(); Directory.CreateDirectory(Path.Combine(p, "changes"));
        var s = new SessaoDocumento { ArquivoOrigem = "x.docx", Formato = FormatoDocumento.Docx, PastaRecovery = p, ChangesAbertos = 2, Recuperada = true };
        File.WriteAllText(s.ArquivoChanges, "\"r1\",\"r2\",");
        var ponte = new PonteDesktop(s);
        ponte.SalvarChanges("n1\",\"n2", 0, 2);   // sessao nova grava 2
        ponte.SalvarChanges("n3", 1, 1);          // desfaz a partir do 2o desta sessao e grava 1
        Assert.Equal(new LeituraChanges("[\"r1\",\"r2\",\"n1\",\"n3\"]", 4, false, 0), ChangesGravados.Ler(s.ArquivoChanges));
    }
}

/// <summary>Recuperacao — a casca pura (Task 3 do Plano 3): quais pastas sao recuperaveis, o descarte, e a
/// varredura de 30 dias que nunca apaga trabalho pendente nem sessao viva. Cada teste trabalha numa raiz
/// propria, apagada no fim.</summary>
public sealed class RecuperaveisTests : IDisposable
{
    private readonly string _raiz = Directory.CreateTempSubdirectory("mOffice-testes-recuperaveis-").FullName;

    public void Dispose()
    {
        try { Directory.Delete(_raiz, true); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    private static string PastaSessao(string raiz, string app, int changes, bool novo = false)
    {
        var p = Path.Combine(raiz, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(Path.Combine(p, "changes"));
        File.WriteAllBytes(Path.Combine(p, "Editor.bin"), [1, 2, 3]);
        File.WriteAllText(Path.Combine(p, "changes", "changes0.json"), string.Concat(Enumerable.Repeat("\"c\",", changes)));
        new MetadadosSessao(@"C:\x\doc.docx", 65, novo ? "Documento" : "doc.docx", novo, app, 1).Gravar(p);
        return p;
    }

    /// <summary>Tudo dentro da pasta (e ela propria) com a data de escrita <paramref name="dias"/> atras.</summary>
    private static void Envelhecer(string p, int dias)
    {
        var quando = DateTime.UtcNow.AddDays(-dias);
        foreach (var f in Directory.EnumerateFileSystemEntries(p, "*", SearchOption.AllDirectories).Append(p))
            if (File.Exists(f)) File.SetLastWriteTimeUtc(f, quando); else Directory.SetLastWriteTimeUtc(f, quando);
    }

    [Fact] // so' entram sessoes do app pedido (sem diferenciar maiusculas), com changes e trava livre
    public void Listar_FiltraPorAppChangesETrava()
    {
        var boa = PastaSessao(_raiz, "mtexto", 2);
        PastaSessao(_raiz, "mtexto", 0);            // abriu e nao editou
        PastaSessao(_raiz, "mplanilha", 3);         // outro app
        var viva = PastaSessao(_raiz, "mtexto", 1);
        using var trava = TravaSessao.Adquirir(viva);  // sessao viva de outra janela
        Assert.Equal([boa], Recuperaveis.Listar(_raiz, "mtexto").Select(d => d.Pasta));
        Assert.Equal([boa], Recuperaveis.Listar(_raiz, "mTexto").Select(d => d.Pasta));
    }

    [Fact] // campos do item vem do sessao.json e do changes0.json; do mais recente para o mais antigo
    public void Listar_MontaOItemEOrdenaDoMaisRecente()
    {
        var antiga = PastaSessao(_raiz, "mtexto", 1);
        var nova = PastaSessao(_raiz, "mtexto", 1, novo: true);
        var hAntiga = DateTime.UtcNow.AddHours(-5); var hNova = DateTime.UtcNow.AddMinutes(-3);
        File.SetLastWriteTimeUtc(Path.Combine(antiga, "changes", "changes0.json"), hAntiga);
        File.SetLastWriteTimeUtc(Path.Combine(nova, "changes", "changes0.json"), hNova);
        var lista = Recuperaveis.Listar(_raiz, "mtexto");
        Assert.Equal([nova, antiga], lista.Select(d => d.Pasta));
        Assert.Equal(new DocumentoRecuperavel(antiga, "doc.docx", @"C:\x\doc.docx", false, hAntiga), lista[1]);
        Assert.Equal(("Documento", true), (lista[0].Titulo, lista[0].Novo));
        Assert.Equal(hNova.ToLocalTime(), lista[0].UltimaAlteracaoLocal);
    }

    [Fact] // cada condicao sozinha tira a pasta da lista: sem sessao.json legivel, sem Editor.bin, changes so' lixo
    public void EhRecuperavel_ExigeMetadadosBinEChanges()
    {
        var semMeta = PastaSessao(_raiz, "mtexto", 1); File.WriteAllText(Path.Combine(semMeta, "sessao.json"), "{quebrado");
        var semBin = PastaSessao(_raiz, "mtexto", 1); File.Delete(Path.Combine(semBin, "Editor.bin"));
        var lixo = PastaSessao(_raiz, "mtexto", 0); File.WriteAllText(Path.Combine(lixo, "changes", "changes0.json"), "\0\0\0");
        var boa = PastaSessao(_raiz, "mtexto", 1);   // controle: a mesma forma, inteira, entra
        Assert.All(new[] { semMeta, semBin, lixo }, p => Assert.False(Recuperaveis.EhRecuperavel(p)));
        Assert.True(Recuperaveis.EhRecuperavel(boa));
        Assert.Equal([boa], Recuperaveis.Listar(_raiz, "mtexto").Select(d => d.Pasta));
        Assert.Empty(Recuperaveis.Listar(Path.Combine(_raiz, "nao-existe"), "mtexto"));
    }

    [Fact] // a varredura de 30 dias nunca apaga um documento recuperavel, mas continua apagando sessao sem changes
    public void Varredura_PreservaRecuperaveis()
    {
        var pendente = PastaSessao(_raiz, "mtexto", 1);
        var vazia = PastaSessao(_raiz, "mtexto", 0);
        Envelhecer(pendente, 90); Envelhecer(vazia, 90);
        Assert.Equal(1, SessaoDocumento.VarrerRecoveryAntigo(_raiz, TimeSpan.FromDays(30)));
        Assert.True(Directory.Exists(pendente)); Assert.False(Directory.Exists(vazia));
    }

    [Fact] // sessao VIVA (trava presa) e velha: a varredura nao toca em nada dela — nem parcialmente. Controle: a
           // mesma pasta, sem a trava presa, sai (prova que o envelhecimento com rec.lock dentro funciona)
    public void Varredura_NaoTocaEmSessaoViva()
    {
        var viva = PastaSessao(_raiz, "mtexto", 0);      // sem changes: so' a trava a protege
        var solta = PastaSessao(_raiz, "mtexto", 0);
        // rec.lock criado ANTES de envelhecer: criar o arquivo depois deixaria a pasta "fresca" e o teste
        // passaria pela idade, sem exercitar a checagem da trava.
        TravaSessao.Adquirir(viva)!.Dispose(); TravaSessao.Adquirir(solta)!.Dispose();
        Envelhecer(viva, 90); Envelhecer(solta, 90);
        using (TravaSessao.Adquirir(viva))
            Assert.Equal(1, SessaoDocumento.VarrerRecoveryAntigo(_raiz, TimeSpan.FromDays(30)));
        Assert.False(Directory.Exists(solta));
        Assert.True(File.Exists(Path.Combine(viva, "Editor.bin")), "a varredura apagou o conteudo de uma sessao viva");
        Assert.True(File.Exists(Path.Combine(viva, "sessao.json")));
    }

    [Fact] // descartar apaga a pasta recuperavel mas nunca uma sessao viva (nem uma pasta que nao e' recuperavel)
    public void Descartar_NaoApagaSessaoViva()
    {
        var viva = PastaSessao(_raiz, "mtexto", 1);
        using (TravaSessao.Adquirir(viva)) Recuperaveis.Descartar(viva);
        Assert.True(Directory.Exists(viva)); Assert.True(File.Exists(Path.Combine(viva, "Editor.bin")));
        Recuperaveis.Descartar(viva);
        Assert.False(Directory.Exists(viva));

        var semChanges = PastaSessao(_raiz, "mtexto", 0);
        Recuperaveis.Descartar(semChanges);
        Assert.True(File.Exists(Path.Combine(semChanges, "Editor.bin")));
        Recuperaveis.Descartar(Path.Combine(_raiz, "nao-existe")); // nao lanca
    }

    [Fact] // pasta que e' JUNCAO para outro lugar: nem lista nem descarta (o conteudo do outro lado fica intacto)
    public void Descartar_JuncaoNaoApagaOOutroLado()
    {
        var alvo = Directory.CreateTempSubdirectory("mOffice-testes-juncao-alvo-").FullName;
        try
        {
            var conteudo = PastaSessao(alvo, "mtexto", 1);
            var juncao = Path.Combine(_raiz, "juncao");
            // mklink /J nao exige privilegio (ao contrario do link simbolico de diretorio).
            using (var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe", $"/c mklink /J \"{juncao}\" \"{conteudo}\"") { CreateNoWindow = true, UseShellExecute = false }))
                p!.WaitForExit();
            Assert.True(new DirectoryInfo(juncao).Attributes.HasFlag(FileAttributes.ReparsePoint), "a juncao nao foi criada");
            Assert.True(Recuperaveis.EhRecuperavel(juncao)); // controle: pelo link o conteudo PARECE recuperavel
            Assert.Empty(Recuperaveis.Listar(_raiz, "mtexto"));
            Recuperaveis.Descartar(juncao);
            Assert.True(File.Exists(Path.Combine(conteudo, "Editor.bin")));
            Assert.True(File.Exists(Path.Combine(conteudo, "changes", "changes0.json")));
            Directory.Delete(juncao); // so' o link
        }
        finally { try { Directory.Delete(alvo, true); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { } }
    }

    [Fact] // abertura NORMAL sem conseguir a trava: sessao.json NAO e' gravado (pasta viva sem trava nao pode
           // parecer "caida" para outra instancia) e o log diz por que. Controle: com a trava, grava.
    public void MarcarRecuperavel_SemTrava_NaoGravaSessaoJson()
    {
        var viva = Path.Combine(_raiz, "viva"); Directory.CreateDirectory(viva);
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = viva };
        var m = new MetadadosSessao(@"C:\x\doc.docx", 65, "doc.docx", false, "mtexto", 1);
        var log = new List<string>();
        using (TravaSessao.Adquirir(viva)) JanelaEditor.MarcarRecuperavel(s, m, log.Add);
        Assert.Null(s.Trava); Assert.Null(s.Metadados);
        Assert.False(File.Exists(Path.Combine(viva, MetadadosSessao.NomeArquivo)));
        Assert.Contains(log, l => l.Contains("sem trava"));

        var ok = Path.Combine(_raiz, "ok"); Directory.CreateDirectory(ok);
        var s2 = new SessaoDocumento { ArquivoOrigem = @"C:\x\doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = ok };
        JanelaEditor.MarcarRecuperavel(s2, m, log.Add);
        try
        {
            Assert.NotNull(s2.Trava); Assert.Equal(m, s2.Metadados); Assert.Equal(m, MetadadosSessao.Ler(ok));
            Assert.False(TravaSessao.Livre(ok));
        }
        finally { s2.LiberarTrava(); }
    }
}

/// <summary>Recuperacao apos queda na janela real (Task 2 do Plano 3): editar, descarregar os changes no disco,
/// "cair" (fechar sem apagar a pasta, soltando a trava como o SO faria) e reabrir pela pasta.</summary>
[Collection(JanelaReal.Nome)]
public class RecuperacaoJanelaTests
{
    private readonly ITestOutputHelper _saida;
    public RecuperacaoJanelaTests(ITestOutputHelper saida) { _saida = saida; }

    /// <summary>Abre um .docx novo, digita " recuperado-321", descarrega os changes no disco e "cai" (fecha
    /// deixando a pasta, com a trava solta). Devolve a pasta do teste e a pasta de recuperacao.</summary>
    private static (string Pasta, string Recovery) EditarECair()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-queda-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var docxAntes = File.GetLastWriteTimeUtc(docx);
        string recovery = "";
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            recovery = j.SessaoAtual!.PastaRecovery;
            // Enquanto a janela vive: metadados gravados (IdUsuario 1) e a pasta travada.
            Assert.Equal(1, MetadadosSessao.Ler(recovery)?.IdUsuario);
            Assert.False(TravaSessao.Livre(recovery));
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" recuperado-321"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
            // Forca o editor a descarregar os changes no disco pela API do host (Task 4: o mesmo caminho do
            // autosave, sem salvar o arquivo — antes o teste chamava `asc_Save(true)` direto no quadro).
            await j.DescarregarChangesAsync();
            Assert.True(new FileInfo(j.SessaoAtual.ArquivoChanges).Length > 0, "o editor nao gravou changes");
            Assert.True(j.Modificado); // descarregar nao e' salvar
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.True(Directory.Exists(recovery));
        // A descarga nao virou um "Salvar" do documento: o arquivo do usuario ficou intocado.
        Assert.Equal(docxAntes, File.GetLastWriteTimeUtc(docx));
        Assert.True(TravaSessao.Livre(recovery));
        return (pasta, recovery);
    }

    [SkippableFact] // editar, gravar os changes, "cair" (fechar sem apagar), reabrir pela pasta: o texto digitado volta
    public void Queda_Reabrir_TextoVolta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, recovery) = EditarECair();
        Assert.True(TravaSessao.Livre(recovery));
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            var tamanhoAntes = new FileInfo(Path.Combine(recovery, "changes", "changes0.json")).Length;
            await j.RecuperarAsync(recovery);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            Assert.Null(j.UltimoErro);
            // Estado do EDITOR (no quadro), antes de qualquer edicao nova: o host entregou a contagem e o
            // "nao salvo" — e' disso que o l5 do sdkjs tira o "modificado" de um documento recuperado.
            Assert.Equal("true", (await j.ExecutarJsAsync("String(window.frames[0].AscDesktopEditor.LocalFileGetOpenChangesCount() > 0)")).Trim('"'));
            Assert.Equal("true", (await j.ExecutarJsAsync("String(window.frames[0].AscDesktopEditor.LocalFileGetSaved() === false)")).Trim('"'));
            // O editor NAO reenviou os changes aplicados como se fossem desta sessao: depois de uma janela maior
            // que o autosave dele (~2 s medidos), o arquivo tem o mesmo tamanho de antes da reabertura.
            await Task.Delay(3000);
            Assert.Equal(tamanhoAntes, new FileInfo(Path.Combine(recovery, "changes", "changes0.json")).Length);
            // O texto voltou UMA vez (aplicar em dobro duplicaria a insercao).
            var texto = await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor);
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(texto, "recuperado-321"));
            Assert.True(j.Modificado);
            Assert.True(j.SessaoAtual!.Recuperada);
            // CheckUserId sobe a cada reabertura da pasta: a sessao que caiu era 1, esta e' 2 (e o sessao.json
            // ja' diz 2, para uma proxima queda subir para 3).
            Assert.Equal("\"2\"", await j.ExecutarJsAsync("window.frames[0].AscDesktopEditor.CheckUserId()"));
            Assert.Equal(2, MetadadosSessao.Ler(recovery)?.IdUsuario);
            Assert.False(TravaSessao.Livre(recovery));
            // Salvar depois de recuperar grava o documento COM os changes recuperados (o x2t le' o Editor.bin +
            // changes0.json da pasta), e o "salvar como" regrava o sessao.json apontando para o arquivo novo.
            var salvo = Path.Combine(pasta, "salvo.docx");
            Assert.True((await j.SalvarComoAsync(salvo, FormatoDocumento.Docx)).Gravou);
            Assert.Contains("recuperado-321", mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(salvo));
            Assert.False(j.Modificado);
            Assert.Equal(salvo, MetadadosSessao.Ler(recovery)?.Origem);
            Assert.Equal(2, MetadadosSessao.Ler(recovery)?.IdUsuario);
            j.Close();
        }, TimeSpan.FromSeconds(180));
        Assert.False(Directory.Exists(recovery)); // fechamento limpo depois da recuperacao apaga a pasta
    }

    [SkippableFact] // pasta sem sessao.json nao e' recuperavel: erro pt-BR, sem abrir nada
    public void Recuperar_PastaSemMetadados_Lanca()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-queda-vazia-").FullName;
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            var e = await Assert.ThrowsAsync<InvalidOperationException>(() => j.RecuperarAsync(pasta));
            Assert.Contains("sessao.json", e.Message);
            Assert.Null(j.SessaoAtual);
            j.Close();
        });
    }

    [SkippableFact] // FALHAR FECHADO: a reaplicacao lanca -> evento com a mensagem, janela fechada sem perguntar,
                    // pasta intacta (changes0.json byte a byte) e trava livre para uma nova tentativa
    public void Recuperar_ReaplicacaoFalha_FechaEMantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (_, recovery) = EditarECair();
        var arquivoChanges = Path.Combine(recovery, "changes", "changes0.json");
        var antes = File.ReadAllBytes(arquivoChanges);
        string? mensagem = null; JanelaEditor? quem = null; var fechou = false; var perguntou = false;
        Task? descarga = null, pdf = null; var imediato = false;
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto");
            j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.Cancelar); };
            j.FalharReaplicacaoParaTeste = true;
            // Sessao congelada (Task 4): descarregar os changes / gerar o PDF de impressao falham NA HORA, com
            // mensagem pt-BR — nunca esperam os 10 s de um LocalFileSaveChanges que a ponte congelada nao grava.
            j.RecuperacaoFalhou += (janela, m) =>
            {
                quem = janela; mensagem = m;
                descarga = j.DescarregarChangesAsync();
                pdf = j.GerarPdfDeImpressaoAsync(mOffice.Nucleo.Impressao.OpcoesImpressao.Padrao);
                imediato = descarga.IsCompleted && pdf.IsCompleted;
            };
            j.Fechado += () => fechou = true;
            j.Show();
            await j.RecuperarAsync(recovery);
            for (var i = 0; i < 300 && !fechou; i++) await Task.Delay(100);
            Assert.True(fechou, "a janela nao fechou depois da falha da reaplicacao");
            Assert.Same(j, quem);
            Assert.Equal(mensagem, j.FalhaDaRecuperacao);
        }, TimeSpan.FromSeconds(180));
        Assert.False(perguntou); // sem a pergunta de salvar: o estado do editor nao e' confiavel
        Assert.True(imediato, "descarga/PDF numa sessao congelada tem de falhar na hora, sem esperar");
        foreach (var t in new[] { descarga, pdf })
        {
            Assert.NotNull(t);
            Assert.True(t!.IsFaulted);
            Assert.Contains("congelada", Assert.IsType<InvalidOperationException>(t.Exception!.InnerException).Message);
        }
        Assert.NotNull(mensagem);
        Assert.StartsWith("Não foi possível reaplicar as alterações recuperadas de \"sintetico.docx\"", mensagem);
        Assert.True(Directory.Exists(recovery));
        Assert.Equal(antes, File.ReadAllBytes(arquivoChanges));
        Assert.True(TravaSessao.Livre(recovery));
        Assert.NotNull(MetadadosSessao.Ler(recovery)); // continua recuperavel
        Directory.Delete(recovery, true);
    }

    [SkippableFact] // changes0.json que nao da' para ler (travado por outro processo): a excecao documentada,
                    // com a causa dentro, e a trava da pasta solta
    public void Recuperar_ChangesIlegivel_LancaComCausaESoltaATrava()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-queda-travada-").FullName;
        new MetadadosSessao(Path.Combine(pasta, "x.docx"), (int)FormatoDocumento.Docx, "x.docx", false, "mtexto", 1).Gravar(pasta);
        File.WriteAllBytes(Path.Combine(pasta, "Editor.bin"), [1, 2, 3]);
        Directory.CreateDirectory(Path.Combine(pasta, "changes"));
        var arquivoChanges = Path.Combine(pasta, "changes", "changes0.json");
        File.WriteAllText(arquivoChanges, "\"1;A\",");
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto"); j.Show();
            using (new FileStream(arquivoChanges, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                var e = await Assert.ThrowsAsync<InvalidOperationException>(() => j.RecuperarAsync(pasta));
                Assert.Contains("changes0.json", e.Message);
                Assert.IsAssignableFrom<IOException>(e.InnerException);
            }
            Assert.True(TravaSessao.Livre(pasta));
            Assert.Null(j.SessaoAtual);
            j.Close();
        });
    }

    // --- Fix round 1 da revisao da Task 3: RETENCAO ate' o conteudo recuperado aparecer no editor ------------

    /// <summary>Reabre <paramref name="recovery"/> com a janela preparada por <paramref name="preparar"/> e fecha
    /// como o gerente de janelas fecharia. A pergunta de fechar responde "Nao salvar" — o pior caso: se ela for
    /// feita numa sessao que nao confirmou, a pasta some. Devolve (perguntou, fechou).</summary>
    private static (bool Perguntou, bool Fechou) RecuperarEFechar(string recovery, Action<JanelaEditor> preparar, Func<JanelaEditor, Task> depoisDeRecuperar, bool esperaLancar)
    {
        var perguntou = false; var fechou = false;
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela("mtexto");
            j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.NaoSalvar); };
            j.Fechado += () => fechou = true;
            preparar(j);
            j.Show();
            try
            {
                if (esperaLancar) await Assert.ThrowsAsync<InvalidOperationException>(() => j.RecuperarAsync(recovery));
                else await j.RecuperarAsync(recovery);
                await depoisDeRecuperar(j);
            }
            finally
            {
                // Fecha mesmo com uma assercao falhando: uma janela viva segura o rec.lock e o Limpar do teste
                // deixaria a pasta para tras (so' com o rec.lock) — visto numa rodada de mutacao.
                j.Close();
                for (var i = 0; i < 100 && !fechou; i++) await Task.Delay(100);
            }
        }, TimeSpan.FromSeconds(180));
        return (perguntou, fechou);
    }

    /// <summary>Apaga as pastas do teste mesmo quando uma assercao falha (senao, com uma mutacao, a pasta de
    /// recuperacao ficaria recuperavel na raiz de dados dos testes).</summary>
    private static void Limpar(params string[] pastas)
    {
        foreach (var p in pastas)
            try { if (Directory.Exists(p)) Directory.Delete(p, true); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    [SkippableFact] // (A) a navegacao falha DEPOIS de a sessao recuperada estar na janela; o gerente fecha a
                    // janela (catch de CriarAsync): sem pergunta, pasta e changes0.json intactos, trava livre
    public void Recuperar_NavegacaoFalha_FechaSemPerguntarEMantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, recovery) = EditarECair();
        try
        {
            var arquivoChanges = Path.Combine(recovery, "changes", "changes0.json");
            var antes = File.ReadAllBytes(arquivoChanges);
            var (perguntou, fechou) = RecuperarEFechar(recovery, j => j.FalharNavegacaoParaTeste = true, j =>
            {
                Assert.True(j.SessaoAtual is { Recuperada: true, Modificado: true }); // sem a regra, o X perguntaria
                Assert.False(j.RecuperacaoConfirmada);
                return Task.CompletedTask;
            }, esperaLancar: true);
            Assert.True(fechou); Assert.False(perguntou);
            Assert.Equal(antes, File.ReadAllBytes(arquivoChanges));
            Assert.True(File.Exists(Path.Combine(recovery, "Editor.bin")));
            Assert.True(TravaSessao.Livre(recovery));
            Assert.True(Recuperaveis.EhRecuperavel(recovery)); // continua na lista da boas-vindas
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // (B) navegacao ok, mas o editor trava em "Carregando" (EndLoad nunca chega); o usuario fecha:
                    // sem pergunta, pasta e changes0.json intactos, trava livre
    public void Recuperar_EditorNaoCarrega_FechaSemPerguntarEMantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, recovery) = EditarECair();
        try
        {
            var arquivoChanges = Path.Combine(recovery, "changes", "changes0.json");
            var antes = File.ReadAllBytes(arquivoChanges);
            var (perguntou, fechou) = RecuperarEFechar(recovery, j => j.SuprimirEndLoadParaTeste = true, async j =>
            {
                // Tempo de sobra para o LocalStartOpen chegar (e ser ignorado pela costura).
                await Task.Delay(5000);
                Assert.Equal("false", (await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"'));
                Assert.True(j.SessaoAtual is { Recuperada: true, Modificado: true });
                Assert.False(j.RecuperacaoConfirmada);
            }, esperaLancar: false);
            Assert.True(fechou); Assert.False(perguntou);
            Assert.Equal(antes, File.ReadAllBytes(arquivoChanges));
            Assert.True(TravaSessao.Livre(recovery));
            Assert.True(Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // controle de (A)/(B): a MESMA recuperacao, com o conteudo na tela, pergunta ao fechar — e
                    // "Nao salvar" (escolha informada) apaga a pasta, como qualquer documento
    public void Recuperar_Confirmada_PerguntaAoFechar()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, recovery) = EditarECair();
        try
        {
            var (perguntou, fechou) = RecuperarEFechar(recovery, _ => { }, async j =>
            {
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                for (var i = 0; i < 50 && !j.RecuperacaoConfirmada; i++) await Task.Delay(100);
                Assert.True(j.RecuperacaoConfirmada);
            }, esperaLancar: false);
            Assert.True(fechou); Assert.True(perguntou);
            Assert.False(Directory.Exists(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    // --- Fix round 2: a confirmacao exige EndLoad ok E o "pronto" do editor da MESMA sessao ------------------

    [SkippableFact] // (C) EndLoad voltou "ok" mas o editor nunca avisou "pronto" (no editor de texto os changes so'
                    // sao ENFILEIRADOS no EndLoad e aplicados depois — uma falha ali nao passa pelo EndLoad):
                    // fechar nao pergunta e a pasta fica
    public void Recuperar_EndLoadOkSemPronto_FechaSemPerguntarEMantemAPasta()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, recovery) = EditarECair();
        try
        {
            var arquivoChanges = Path.Combine(recovery, "changes", "changes0.json");
            var antes = File.ReadAllBytes(arquivoChanges);
            var (perguntou, fechou) = RecuperarEFechar(recovery, j => j.IgnorarProntoParaTeste = true, async j =>
            {
                // O editor carregou de verdade (espelho do JS) e o sinal 1 chegou; so' o "pronto" nao contou.
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                for (var i = 0; i < 50 && !j.EndLoadRecuperadoOk; i++) await Task.Delay(100);
                Assert.True(j.EndLoadRecuperadoOk);
                Assert.False(j.RecuperacaoConfirmada);
            }, esperaLancar: false);
            Assert.True(fechou); Assert.False(perguntou);
            Assert.Equal(antes, File.ReadAllBytes(arquivoChanges));
            Assert.True(TravaSessao.Livre(recovery));
            Assert.True(Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    [SkippableFact] // troca de sessao na MESMA janela com a recuperacao ainda pendente: a pasta recuperada fica
                    // (so' a trava e' solta) e continua recuperavel; o documento novo abre normalmente
    public void Recuperar_TrocaDeSessaoPendente_MantemAPastaAnterior()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var (pasta, recovery) = EditarECair();
        try
        {
            var outro = mOffice.Conversor.Tests.FabricaDocumentos.Docx(Directory.CreateDirectory(Path.Combine(pasta, "outro")).FullName, "Outro");
            var perguntou = false;
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.PerguntarAoFechar = _ => { perguntou = true; return Task.FromResult(DecisaoFechar.NaoSalvar); };
                j.SuprimirEndLoadParaTeste = true;
                j.Show();
                try
                {
                    await j.RecuperarAsync(recovery);
                    Assert.False(j.RecuperacaoConfirmada);
                    j.SuprimirEndLoadParaTeste = false;
                    await j.AbrirAsync(outro);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    Assert.False(j.SessaoAtual!.Recuperada);
                    Assert.True(Directory.Exists(recovery));
                    Assert.True(TravaSessao.Livre(recovery));
                    Assert.True(Recuperaveis.EhRecuperavel(recovery));
                }
                finally { j.Close(); }
            }, TimeSpan.FromSeconds(180));
            Assert.False(perguntou);
            Assert.True(Recuperaveis.EhRecuperavel(recovery));
        }
        finally { Limpar(recovery, pasta); }
    }

    // --- Task 8 do Plano 3: pergunta 6 da spec §14 — recuperacao com dezenas de milhares de changes ----------

    /// <summary>Changes minimos que a medida exige (spec §14 pergunta 6: "dezenas de milhares").</summary>
    private const int ChangesDaMedida = 20_000;

    /// <summary>Insercoes de 1 caractere por script (lote). Um script so' — o editor processa o lote inteiro sem
    /// voltar ao host; entre lotes a descarga grava os changes e a contagem decide se precisa de mais. Medido na
    /// primeira rodada (Task 8): cada asc_AddText de 1 caractere neste documento vira ~35 changes (2.000
    /// insercoes = 70.297 changes), entao lotes pequenos param perto do minimo em vez de passar dele 3,5 vezes.</summary>
    private const int InsercoesPorLote = 250;

    /// <summary>Marcador unico digitado DEPOIS de todos os lotes: a recuperacao so' vale se ele voltar (o ultimo
    /// change aplicado, nao so' os primeiros).</summary>
    private const string FimDaMedida = "fim-medida-777";

    [SkippableFact, Trait("Categoria", "Medida")]
    // Spec §14 pergunta 6 (Task 8 do Plano 3): reabrir uma pasta com >= 20.000 changes leva no maximo o DOBRO
    // da abertura normal do mesmo documento. Media de 3 aberturas normais x media de 3 recuperacoes (cada uma de
    // uma COPIA da pasta que "caiu", porque recuperar muda o sessao.json). Tempo = do pedido (AbrirAsync /
    // RecuperarAsync, com a WebView2 da janela ja' pronta) ate' o editor avisar "pronto"
    // (asc_onDocumentContentReady) — no editor de texto os changes recuperados sao aplicados ANTES desse aviso
    // (Task 3: o EndLoad so' os enfileira). Uma abertura de aquecimento, nao contada, vem antes das medidas.
    // OPT-IN (fix round 1 da revisao da Task 8), na convencao de AmostrasEMedidasTests: a asserção e' uma RAZAO
    // entre tempos (folga medida 1,38x contra o limiar de 2x) e nao pode deixar a suite normal instavel numa
    // maquina carregada. Rode com MOFFICE_MEDIDAS=1.
    public void Medida_RecuperarMilharesDeChanges_AteODobroDaAbertura()
    {
        Skip.If(Environment.GetEnvironmentVariable("MOFFICE_MEDIDAS") != "1",
                "medida de tempo da recuperacao com milhares de changes: rode com MOFFICE_MEDIDAS=1");
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-medida-recuperacao-").FullName;
        string? recovery = null;
        try
        {
            var docx = DocxLongo(pasta, paragrafos: 240, palavrasPorParagrafo: 50);

            // 1) Abertura normal: aquecimento + 3 medidas.
            var aberturas = new List<double>();
            var paginas = "?";
            for (var rodada = 0; rodada <= 3; rodada++)
            {
                var ms = 0.0;
                JanelaSta.Executar(async () =>
                {
                    var j = EditorReal.NovaJanela("mtexto"); j.Show();
                    try
                    {
                        await j.ProntoAsync();
                        ms = await MedirAtePronto(j, () => j.AbrirAsync(docx));
                        Assert.Null(j.UltimoErro);
                        if (rodada == 0) paginas = await PaginasDepoisDoLayout(j);
                    }
                    finally { j.Close(); }
                }, TimeSpan.FromMinutes(3));
                if (rodada > 0) aberturas.Add(ms);
                _saida.WriteLine($"MEDIDA abertura normal {(rodada == 0 ? "(aquecimento, nao conta)" : "#" + rodada)}: {ms:F0} ms");
            }
            _saida.WriteLine($"MEDIDA documento: {paginas} paginas (240 paragrafos x 50 palavras)");

            // 2) Gerar >= 20.000 changes digitando 1 caractere por chamada, em lotes; depois "cair".
            var insercoes = 0; var changes = 0; var geracao = System.Diagnostics.Stopwatch.StartNew();
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
                try
                {
                    await j.AbrirAsync(docx);
                    Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                    recovery = j.SessaoAtual!.PastaRecovery;
                    for (var lote = 0; lote < 40 && changes < ChangesDaMedida; lote++)
                    {
                        var r = (await j.ExecutarJsAsync(LoteDeInsercoes(InsercoesPorLote))).Trim('"');
                        Assert.Equal("ok " + InsercoesPorLote, r);
                        insercoes += InsercoesPorLote;
                        await j.DescarregarChangesAsync();
                        changes = ChangesGravados.Ler(j.SessaoAtual.ArquivoChanges).Quantidade;
                        _saida.WriteLine($"MEDIDA lote {lote + 1}: {insercoes} insercoes, {changes} changes no disco");
                    }
                    Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" " + FimDaMedida))).Trim('"'));
                    insercoes++;
                    await j.DescarregarChangesAsync();
                    changes = ChangesGravados.Ler(j.SessaoAtual.ArquivoChanges).Quantidade;
                }
                finally { j.Close(); }
            }, TimeSpan.FromMinutes(20));
            geracao.Stop();
            _saida.WriteLine($"MEDIDA geracao: {insercoes} insercoes -> {changes} changes em {geracao.Elapsed.TotalSeconds:F0} s; " +
                             $"changes0.json = {new FileInfo(Path.Combine(recovery!, "changes", "changes0.json")).Length / 1024} KiB");
            Assert.True(changes >= ChangesDaMedida, $"so' {changes} changes gerados (minimo {ChangesDaMedida})");
            Assert.True(changes >= insercoes, $"{insercoes} insercoes viraram so' {changes} changes (cada insercao deveria virar ao menos 1)");
            Assert.True(TravaSessao.Livre(recovery!));

            // A pasta que "caiu" vira a matriz; cada rodada recupera uma COPIA dela.
            var matriz = Path.Combine(pasta, "matriz");
            CopiarPasta(recovery!, matriz);
            Directory.Delete(recovery!, true);

            // 3) Recuperacao: 3 medidas.
            var recuperacoes = new List<double>();
            for (var rodada = 1; rodada <= 3; rodada++)
            {
                var copia = Path.Combine(pasta, "recuperacao-" + rodada);
                CopiarPasta(matriz, copia);
                var ms = 0.0;
                JanelaSta.Executar(async () =>
                {
                    var j = EditorReal.NovaJanela("mtexto"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
                    try
                    {
                        await j.ProntoAsync();
                        ms = await MedirAtePronto(j, () => j.RecuperarAsync(copia));
                        Assert.Null(j.UltimoErro);
                        Assert.Null(j.FalhaDaRecuperacao);
                        Assert.True(j.SessaoAtual!.Recuperada);
                        Assert.Equal(changes, j.SessaoAtual.ChangesAbertos);
                        // O ULTIMO change voltou (e uma vez so'): a medida nao conta uma reaplicacao pela metade.
                        var texto = await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor);
                        Assert.Single(System.Text.RegularExpressions.Regex.Matches(texto, FimDaMedida));
                    }
                    finally { j.Close(); }
                }, TimeSpan.FromMinutes(5));
                recuperacoes.Add(ms);
                _saida.WriteLine($"MEDIDA recuperacao #{rodada} ({changes} changes): {ms:F0} ms");
                Limpar(copia);
            }

            var mediaAbertura = aberturas.Average();
            var mediaRecuperacao = recuperacoes.Average();
            var razao = mediaRecuperacao / mediaAbertura;
            _saida.WriteLine($"MEDIDA RESUMO: abertura normal media {mediaAbertura:F0} ms ({string.Join(" / ", aberturas.Select(a => a.ToString("F0")))}); " +
                             $"recuperacao com {changes} changes media {mediaRecuperacao:F0} ms ({string.Join(" / ", recuperacoes.Select(a => a.ToString("F0")))}); " +
                             $"razao {razao:F2}x (limiar 2,00x)");
            Assert.True(razao <= 2.0, $"recuperacao {mediaRecuperacao:F0} ms > 2x a abertura normal {mediaAbertura:F0} ms (razao {razao:F2})");
        }
        finally { Limpar(pasta); if (recovery is not null) Limpar(recovery); }
    }

    /// <summary>Do pedido ate' o "pronto" do editor, consultando a cada 50 ms (a espera de <see cref="EditorReal.Esperar"/>
    /// tem passo de 200 ms — grosso demais para uma razao entre dois tempos de poucos segundos).</summary>
    private static async Task<double> MedirAtePronto(JanelaEditor j, Func<Task> abrir)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        await abrir();
        for (var i = 0; i < 2400; i++)
        {
            if ((await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"') == "true") return sw.Elapsed.TotalMilliseconds;
            await Task.Delay(50);
        }
        throw new Xunit.Sdk.XunitException("o editor nao ficou pronto em 120 s");
    }

    /// <summary>Numero de paginas depois de o layout terminar: o editor avisa "pronto" antes de paginar o documento
    /// inteiro (logo depois do "pronto" o getCountPages ainda diz 1), entao a contagem e' lida ate' ficar estavel
    /// por 1 s (no maximo 20 s). So' informativo — confirma o "~20 paginas" da medida.</summary>
    private static async Task<string> PaginasDepoisDoLayout(JanelaEditor j)
    {
        var ultimo = ""; var estavel = 0;
        for (var i = 0; i < 80 && estavel < 4; i++)
        {
            await Task.Delay(250);
            var atual = (await j.ExecutarJsAsync(EditorReal.NoEditor(
                "return String(typeof api.getCountPages === 'function' ? api.getCountPages() : -1);"))).Trim('"');
            estavel = atual == ultimo ? estavel + 1 : 0;
            ultimo = atual;
        }
        return ultimo;
    }

    /// <summary>Um lote de <paramref name="n"/> insercoes de UM caractere cada (asc_AddText, o mesmo caminho de
    /// <see cref="EditorReal.InserirTextoWord"/>), num unico script.</summary>
    private static string LoteDeInsercoes(int n) => EditorReal.NoEditor($$"""
            if (typeof api.asc_AddText !== "function") return "sem asc_AddText";
            if (typeof api.asc_RemoveSelection === "function") api.asc_RemoveSelection();
            var letras = "abcdefghij";
            for (var i = 0; i < {{n}}; i++) api.asc_AddText(letras.charAt(i % 10));
            return "ok {{n}}";
        """);

    /// <summary>DOCX sintetico de ~20 paginas: <paramref name="paragrafos"/> paragrafos de texto repetido. Varios
    /// paragrafos (e nao um so', como <see cref="mOffice.Conversor.Tests.FabricaDocumentos.Docx"/>): as
    /// insercoes caem no primeiro, e um paragrafo unico de 20 paginas faria cada insercao refazer o layout dele
    /// inteiro — mediria o paragrafo gigante, nao a recuperacao.</summary>
    private static string DocxLongo(string pasta, int paragrafos, int palavrasPorParagrafo)
    {
        var caminho = Path.Combine(pasta, "longo.docx");
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(caminho, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var corpo = new DocumentFormat.OpenXml.Wordprocessing.Body();
        for (var p = 0; p < paragrafos; p++)
        {
            var texto = string.Join(" ", Enumerable.Range(0, palavrasPorParagrafo).Select(w => $"palavra{p * palavrasPorParagrafo + w}"));
            corpo.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(texto))));
        }
        main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(corpo);
        main.Document.Save();
        return caminho;
    }

    private static void CopiarPasta(string origem, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var d in Directory.GetDirectories(origem, "*", SearchOption.AllDirectories))
            Directory.CreateDirectory(Path.Combine(destino, Path.GetRelativePath(origem, d)));
        foreach (var f in Directory.GetFiles(origem, "*", SearchOption.AllDirectories))
            File.Copy(f, Path.Combine(destino, Path.GetRelativePath(origem, f)), overwrite: true);
    }
}
