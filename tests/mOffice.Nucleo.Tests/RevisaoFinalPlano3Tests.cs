using System.Text.Json.Nodes;
using mOffice.Conversor;
using mOffice.Nucleo.Ortografia;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Nucleo.Tests;

/// <summary>Revisao final do Plano 3 — as pecas puras (sem janela): M1 (reparo + leitura numa leitura so', que
/// lanca), M2/M7 (gravacao atomica e no disco), M4 (senha do XML do editor), M5 (teto do corretor), M6 (o
/// "Nao salvar" apaga o sessao.json primeiro) e I2 (carimbo do original no sessao.json).</summary>
public sealed class RevisaoFinalPlano3Tests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("mOffice-testes-revfinal-").FullName;

    public void Dispose()
    {
        try { Directory.Delete(_pasta, true); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    private string PastaSessao(int changes)
    {
        var p = Path.Combine(_pasta, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(Path.Combine(p, "changes"));
        File.WriteAllBytes(Path.Combine(p, "Editor.bin"), [1, 2, 3]);
        File.WriteAllText(Path.Combine(p, "changes", "changes0.json"), string.Concat(Enumerable.Repeat("\"c\",", changes)));
        new MetadadosSessao(@"C:\x\doc.docx", 65, "doc.docx", false, "mtexto", 1).Gravar(p);
        return p;
    }

    // --- M1 -----------------------------------------------------------------------------------------------

    [Fact] // M1: changes0.json que nao da' para ler LANCA (nunca "zero changes" em silencio, como a Ler pura faz)
    public void RepararELer_ArquivoIlegivel_Lanca()
    {
        var arq = Path.Combine(_pasta, "changes0.json");
        File.WriteAllText(arq, "\"c1\",\"c2\",");
        using (new FileStream(arq, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            Assert.ThrowsAny<IOException>(() => ChangesGravados.RepararELer(arq));
            // Contraste: a leitura PURA (usada por quem so' sonda pastas alheias) engole e diz "zero" — era isso que a
            // recuperacao usava depois do reparo.
            Assert.Equal(0, ChangesGravados.Ler(arq).Quantidade);
        }
        // Controle: o mesmo arquivo, legivel.
        var (reparo, leitura) = ChangesGravados.RepararELer(arq);
        Assert.False(reparo.Reparou);
        Assert.Equal(new LeituraChanges("[\"c1\",\"c2\"]", 2, false, 0), leitura);
    }

    [Fact] // M1: com cauda rasgada, a leitura devolvida e' o que ficou NO DISCO depois do reparo (uma leitura so')
    public void RepararELer_CaudaRasgada_DevolveOQueFicouNoDisco()
    {
        var arq = Path.Combine(_pasta, "changes0.json");
        File.WriteAllText(arq, "\"c1\",\"c2\",\"c3-pela-met");
        var (reparo, leitura) = ChangesGravados.RepararELer(arq);
        Assert.True(reparo.Reparou);
        Assert.Equal(2, reparo.Quantidade);
        Assert.Equal(ChangesGravados.Ler(arq), leitura);
        Assert.Equal(new LeituraChanges("[\"c1\",\"c2\"]", 2, false, 0), leitura);
        Assert.Equal((new ReparoChanges(false, 0, 0, null), new LeituraChanges("[]", 0, false, 0)),
            ChangesGravados.RepararELer(Path.Combine(_pasta, "nao-existe.json")));
    }

    // --- M2 / M7 ------------------------------------------------------------------------------------------

    [Fact] // M2: o truncamento do changes0.json (deleteIndex) TROCA o arquivo inteiro (.tmp + troca de nome), nunca
           // reescreve o original no lugar — o NOME passa a apontar para OUTRO arquivo do NTFS (indice diferente). E' o
           // que garante que uma queda no meio deixa o changes0.json antigo inteiro, e nao truncado/rasgado.
    public void SalvarChanges_Truncamento_TrocaOArquivoInteiro()
    {
        var p = Path.Combine(_pasta, "sessao"); Directory.CreateDirectory(p);
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = p };
        var ponte = new PonteDesktop(s);
        ponte.SalvarChanges("c1\",\"c2\",\"c3", -1, 3);
        Assert.Equal("\"c1\",\"c2\",\"c3\",", File.ReadAllText(s.ArquivoChanges));
        var indiceAntes = IndiceNtfs(s.ArquivoChanges);
        ponte.SalvarChanges("", 2, 0);   // so' desfaz ate' o 2o change (sem apendar: o arquivo final e' so' a troca)
        Assert.Equal("\"c1\",\"c2\",", File.ReadAllText(s.ArquivoChanges));
        Assert.NotEqual(indiceAntes, IndiceNtfs(s.ArquivoChanges));
        Assert.False(File.Exists(s.ArquivoChanges + ".tmp"));
        // Controle do instrumento: apendar (o caminho normal, que NAO troca o arquivo) mantem o mesmo indice.
        var indiceTrocado = IndiceNtfs(s.ArquivoChanges);
        ponte.SalvarChanges("n3", -1, 1);
        Assert.Equal("\"c1\",\"c2\",\"n3\",", File.ReadAllText(s.ArquivoChanges));
        Assert.Equal(indiceTrocado, IndiceNtfs(s.ArquivoChanges));
    }

    [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetFileInformationByHandle(Microsoft.Win32.SafeHandles.SafeFileHandle arquivo, out InfoArquivo info);

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct InfoArquivo
    {
        public uint Atributos;
        public System.Runtime.InteropServices.ComTypes.FILETIME Criacao, Acesso, Escrita;
        public uint Volume, TamanhoAlto, TamanhoBaixo, Links, IndiceAlto, IndiceBaixo;
    }

    /// <summary>Identidade do arquivo no NTFS (volume + indice): muda quando o NOME passa a apontar para outro arquivo.</summary>
    private static (uint, uint, uint) IndiceNtfs(string arquivo)
    {
        using var h = File.OpenHandle(arquivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        Assert.True(GetFileInformationByHandle(h, out var i));
        return (i.Volume, i.IndiceAlto, i.IndiceBaixo);
    }

    [Fact] // M7 (e M2): o flush para o disco antes da troca nao e' observavel por teste de comportamento (so' uma queda
           // de energia o mostraria). Guarda estrutural: a gravacao atomica faz Flush(true) ANTES do Move, e os tres
           // arquivos que a recuperacao le' depois de uma queda passam por ela.
    public void GravacaoAtomica_FlushAntesDaTroca_EUsadaPelosArquivosDaRecuperacao()
    {
        var src = Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo");
        var atomica = File.ReadAllText(Path.Combine(src, "Recuperacao", "GravacaoAtomica.cs"));
        var flush = atomica.IndexOf("fs.Flush(flushToDisk: true);", StringComparison.Ordinal);
        var troca = atomica.IndexOf("File.Move(tmp, destino, true)", StringComparison.Ordinal);
        Assert.True(flush > 0 && troca > flush, "GravacaoAtomica tem de fazer Flush(true) antes do File.Move");
        Assert.Contains("GravacaoAtomica.Gravar(", File.ReadAllText(Path.Combine(src, "Recuperacao", "MetadadosSessao.cs")));
        Assert.Contains("GravacaoAtomica.Gravar(", File.ReadAllText(Path.Combine(src, "Recuperacao", "ChangesGravados.cs")));
        Assert.Contains("GravacaoAtomica.Truncar(", File.ReadAllText(Path.Combine(src, "PonteDesktop.cs")));   // Truncar comeca pela Gravar
        // Comportamento: Gravar troca por cima e nao deixa o .tmp.
        var p = Path.Combine(_pasta, "m"); var m = new MetadadosSessao(@"C:\x\a.docx", 65, "a.docx", false, "mtexto", 1);
        m.Gravar(p); (m with { IdUsuario = 2 }).Gravar(p);
        Assert.Equal(2, MetadadosSessao.Ler(p)!.IdUsuario);
        Assert.False(File.Exists(Path.Combine(p, MetadadosSessao.NomeArquivo + ".tmp")));
    }

    [Fact] // M2: a troca de nome nao substitui um arquivo aberto por outro processo (antivirus lendo o que acabou de ser
           // gravado): um leitor que solta o arquivo logo depois nao derruba a gravacao (ela tenta de novo); um que nao
           // solta faz a gravacao falhar com o arquivo ANTIGO inteiro e sem .tmp sobrando
    public async Task GravacaoAtomica_LeitorMomentaneo_TentaDeNovo_LeitorPreso_FalhaSemEstragar()
    {
        var arq = Path.Combine(_pasta, "changes0.json");
        File.WriteAllText(arq, "\"antigo\",");
        var leitor = new FileStream(arq, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var solta = Task.Run(async () => { await Task.Delay(120); leitor.Dispose(); });
        GravacaoAtomica.Gravar(arq, "\"novo\","u8);
        await solta;
        Assert.Equal("\"novo\",", File.ReadAllText(arq));

        using (new FileStream(arq, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            Assert.ThrowsAny<Exception>(() => GravacaoAtomica.Gravar(arq, "\"outro\","u8));
        Assert.Equal("\"novo\",", File.ReadAllText(arq));
        Assert.False(File.Exists(arq + ".tmp"));
    }

    // --- Fix round 2, itens 1 e 2: contadores so' andam com o arquivo em dia -------------------------------

    /// <summary>Sessao com "c1","c2","c3" gravados pela ponte (IndiceChanges = 3).</summary>
    private (SessaoDocumento S, PonteDesktop Ponte, List<string> Falhas, Func<int> Salvos) SessaoComTresChanges()
    {
        var p = Path.Combine(_pasta, Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p);
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = p };
        var falhas = new List<string>();
        var ponte = new PonteDesktop(s, aoFalharGravacao: falhas.Add);
        Assert.Equal("ok", ponte.SalvarChanges("c1\",\"c2\",\"c3", -1, 3));
        var salvos = 0;
        ponte.ChangesSalvos += () => salvos++;
        return (s, ponte, falhas, () => salvos);
    }

    [Fact] // item 1: truncamento com um leitor que compartilha escrita/delecao (a troca atomica falha): corta NO LUGAR,
           // conteudo certo, indice e ChangesSalvos avancam
    public void SalvarChanges_TruncamentoComLeitorCompartilhando_CortaNoLugar()
    {
        var (s, ponte, falhas, salvos) = SessaoComTresChanges();
        using (var leitor = new FileStream(s.ArquivoChanges, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            Assert.Equal("ok", ponte.SalvarChanges("n3", 2, 1));
        Assert.Equal("\"c1\",\"c2\",\"n3\",", File.ReadAllText(s.ArquivoChanges));
        Assert.Equal(3, s.IndiceChanges);
        Assert.Equal(1, salvos());
        Assert.Empty(falhas);
        Assert.False(ponte.Congelada);
        Assert.False(File.Exists(s.ArquivoChanges + ".tmp"));
    }

    [Fact] // item 1: truncamento com o arquivo preso por quem NAO compartilha escrita (troca e corte no lugar impossiveis):
           // falha FECHADA — "falhou", ponte congelada, indice parado, sem ChangesSalvos, arquivo byte a byte, host avisado;
           // e nada mais grava depois ("congelada")
    public void SalvarChanges_TruncamentoComArquivoPreso_FalhaFechadoSemDivergir()
    {
        var (s, ponte, falhas, salvos) = SessaoComTresChanges();
        var antes = File.ReadAllBytes(s.ArquivoChanges);
        using (new FileStream(s.ArquivoChanges, FileMode.Open, FileAccess.Read, FileShare.Read))
            Assert.Equal("falhou", ponte.SalvarChanges("n3", 2, 1));
        Assert.Equal(antes, File.ReadAllBytes(s.ArquivoChanges));
        Assert.Equal(3, s.IndiceChanges);
        Assert.Equal(0, salvos());
        Assert.Single(falhas);
        Assert.True(ponte.Congelada);
        Assert.Equal("congelada", ponte.SalvarChanges("n4", -1, 1));
        Assert.Equal(antes, File.ReadAllBytes(s.ArquivoChanges));
        Assert.Equal(3, s.IndiceChanges);
    }

    [Fact] // item 2: apendice com o arquivo preso por um instante (leitura com FileShare.Read, como o File.ReadAllBytes de
           // antes na boas-vindas) tenta de novo e grava; com um leitor que compartilha escrita, grava direto (controle)
    public async Task SalvarChanges_Apendice_TentaDeNovoComArquivoPresoPorUmInstante()
    {
        var (s, ponte, falhas, salvos) = SessaoComTresChanges();
        var leitor = new FileStream(s.ArquivoChanges, FileMode.Open, FileAccess.Read, FileShare.Read);
        var solta = Task.Run(async () => { await Task.Delay(120); leitor.Dispose(); });
        Assert.Equal("ok", ponte.SalvarChanges("n4", -1, 1));
        await solta;
        using (new FileStream(s.ArquivoChanges, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            Assert.Equal("ok", ponte.SalvarChanges("n5", -1, 1));
        Assert.Equal("\"c1\",\"c2\",\"c3\",\"n4\",\"n5\",", File.ReadAllText(s.ArquivoChanges));
        Assert.Equal(5, s.IndiceChanges);
        Assert.Equal(2, salvos());
        Assert.Empty(falhas);
    }

    [Fact] // item 2: apendice com o arquivo preso de vez: falha fechada, sem meio lote no arquivo e sem o indice andar
    public void SalvarChanges_ApendiceComArquivoPreso_FalhaFechado()
    {
        var (s, ponte, falhas, salvos) = SessaoComTresChanges();
        var antes = File.ReadAllBytes(s.ArquivoChanges);
        using (new FileStream(s.ArquivoChanges, FileMode.Open, FileAccess.Read, FileShare.Read))
            Assert.Equal("falhou", ponte.SalvarChanges("n4", -1, 1));
        Assert.Equal(antes, File.ReadAllBytes(s.ArquivoChanges));
        Assert.Equal(3, s.IndiceChanges);
        Assert.Equal(0, salvos());
        Assert.Single(falhas);
    }

    [Fact] // item 2: a leitura (boas-vindas, outra instancia) nao atrapalha quem esta' gravando — le' com o arquivo aberto
           // para escrita pela sessao
    public void Ler_ComArquivoAbertoParaEscrita_LeENaoAtrapalha()
    {
        var arq = Path.Combine(_pasta, "changes0.json");
        File.WriteAllText(arq, "\"c1\",\"c2\",");
        using var gravando = new FileStream(arq, FileMode.Append, FileAccess.Write, FileShare.Read | FileShare.Delete);
        Assert.Equal(2, ChangesGravados.Ler(arq).Quantidade);
        Assert.Equal(2, ChangesGravados.RepararELer(arq).Leitura.Quantidade);
    }

    [Fact] // item 2: a pasta de uma sessao VIVA (trava presa) nunca e' lida por Listar/EhRecuperavel — nem sessao.json nem
           // changes0.json. Controle: a pasta caida (trava livre) e' lida e listada.
    public void Listar_NuncaLeAPastaDeSessaoViva()
    {
        var viva = PastaSessao(changes: 2);
        var caida = PastaSessao(changes: 2);
        var lidos = new List<string>();
        void Anotar(string c) { if (c.StartsWith(_pasta, StringComparison.OrdinalIgnoreCase)) lock (lidos) lidos.Add(c); }
        using var trava = TravaSessao.Adquirir(viva);
        Assert.NotNull(trava);
        GravacaoAtomica.AoLerParaTeste += Anotar;
        try
        {
            Assert.Equal([caida], Recuperaveis.Listar(_pasta, "mtexto").Select(d => d.Pasta));
            Assert.False(Recuperaveis.EhRecuperavel(viva));
        }
        finally { GravacaoAtomica.AoLerParaTeste -= Anotar; }
        lock (lidos)
        {
            Assert.DoesNotContain(lidos, c => c.StartsWith(viva, StringComparison.OrdinalIgnoreCase));
            Assert.Contains(lidos, c => c.StartsWith(caida, StringComparison.OrdinalIgnoreCase));   // controle do instrumento
        }
    }

    // --- M4 -----------------------------------------------------------------------------------------------

    [Theory] // M4: senha feita so' de espacos (e espacos nas pontas) e' senha — nao vira ""
    [InlineData("<m_sPassword>   </m_sPassword>", "   ")]
    [InlineData("<m_sPassword> a b </m_sPassword>", " a b ")]
    [InlineData("<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>\n<m_sPassword>\t</m_sPassword>", "\t")]
    public void Senha_EspacosSaoPreservados(string xml, string esperado) => Assert.Equal(esperado, OpcoesAvancadas.Senha(xml));

    [Fact] // M4: o leitor e' explicito — DTD proibido e sem resolver externo (o XElement.Parse de antes usava DtdProcessing.Parse)
    public void Senha_LeitorProibeDtd()
    {
        Assert.Equal(System.Xml.DtdProcessing.Prohibit, OpcoesAvancadas.Configuracao.DtdProcessing);
        Assert.False(OpcoesAvancadas.Configuracao.IgnoreWhitespace);
        Assert.Null(OpcoesAvancadas.Senha("<!DOCTYPE x [<!ENTITY e \"y\">]><m_sPassword>&e;</m_sPassword>"));
        Assert.Null(OpcoesAvancadas.Senha("<m_sPassword>&e;</m_sPassword>"));   // entidade nao declarada
        Assert.Equal("abc", OpcoesAvancadas.Senha("<m_sPassword>abc</m_sPassword>"));   // controle
    }

    // --- M5 (corretor) ------------------------------------------------------------------------------------

    private static string PedidoSpell(int n, string tipo = "spell") =>
        new JsonObject
        {
            ["type"] = tipo, ["ParagraphId"] = "1", ["RecalcId"] = 1,
            ["usrWords"] = new JsonArray(Enumerable.Range(0, n).Select(i => (JsonNode?)JsonValue.Create("p" + i)).ToArray()),
            ["usrLang"] = new JsonArray(Enumerable.Range(0, n).Select(_ => (JsonNode?)JsonValue.Create(1046)).ToArray()),
        }.ToJsonString();

    [Fact] // M5 refeito no fix round 2: "spell" NAO tem teto de palavras e e' SEMPRE respondido (o editor de planilha manda
           // todas as celulas de texto de uma coluna num pedido so'; sem resposta, o corretor dele fica preso esperando)
    public void Ortografia_SpellGrande_SempreResponde()
    {
        var dic = new DicionarioUsuario(Path.Combine(_pasta, "dic.txt"));
        var chamadas = 0;
        var r = PedidoOrtografia.Responder(PedidoSpell(10_000), (_, _) => { chamadas++; return true; }, (_, _) => [], dic);
        Assert.NotNull(r);
        Assert.Equal(10_000, chamadas);
        Assert.Equal(10_000, JsonNode.Parse(r!)!["usrCorrect"]!.AsArray().Count);
    }

    [Fact] // M5 (fix round 2): "suggest" (a operacao cara) tem teto de palavras COM sugestao, mas tambem e' sempre respondido —
           // acima do teto as palavras recebem lista vazia, com log
    public void Ortografia_SuggestGrande_RespondeComTetoDeSugestoes()
    {
        var dic = new DicionarioUsuario(Path.Combine(_pasta, "dic.txt"));
        var chamadas = 0; var log = new List<string>();
        var max = PedidoOrtografia.MaximoPalavrasComSugestao;
        var r = PedidoOrtografia.Responder(PedidoSpell(max + 50, "suggest"), (_, _) => true, (_, _) => { chamadas++; return ["x"]; }, dic, log.Add);
        Assert.NotNull(r);
        Assert.Equal(max, chamadas);
        var listas = JsonNode.Parse(r!)!["usrSuggest"]!.AsArray();
        Assert.Equal(max + 50, listas.Count);
        Assert.Single(listas[max - 1]!.AsArray());
        Assert.Empty(listas[max]!.AsArray());
        Assert.Contains(log, l => l.Contains("suggest") && l.Contains((max + 50).ToString()));
    }

    [Fact] // M5: texto do pedido acima do teto nem e' parseado
    public void Ortografia_TetoDeCaracteres()
    {
        var log = new List<string>();
        var enorme = "{\"type\":\"spell\",\"x\":\"" + new string('a', PedidoOrtografia.MaximoCaracteresPorPedido) + "\"}";
        Assert.Null(PedidoOrtografia.Responder(enorme, (_, _) => true, (_, _) => [], new DicionarioUsuario(Path.Combine(_pasta, "d.txt")), log.Add));
        Assert.Contains(log, l => l.Contains("caracteres"));
    }

    // --- M6 -----------------------------------------------------------------------------------------------

    [Fact] // M6: o apagar do fechamento limpo ("Nao salvar") tira o sessao.json PRIMEIRO — interrompido logo depois
           // (costura: queda simulada antes do apagar recursivo), a pasta que sobra ja' nao e' recuperavel, e o documento
           // descartado nao reaparece como "recuperado". Controle: sem interrupcao, some tudo.
    public void ApagarRecovery_Interrompido_NaoDeixaPastaRecuperavel()
    {
        var p = PastaSessao(changes: 2);
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = p };
        s.Trava = TravaSessao.Adquirir(p);
        Assert.NotNull(s.Trava);
        s.AntesDoApagarRecursivoParaTeste = () => throw new IOException("queda simulada no meio do apagar");
        s.ApagarRecovery();
        Assert.True(File.Exists(Path.Combine(p, "Editor.bin")), "a costura devia ter interrompido o apagar");
        Assert.True(File.Exists(Path.Combine(p, "changes", "changes0.json")));
        Assert.False(File.Exists(Path.Combine(p, MetadadosSessao.NomeArquivo)));
        Assert.False(Recuperaveis.EhRecuperavel(p));
        Assert.Empty(Recuperaveis.Listar(_pasta, "mtexto"));

        var q = PastaSessao(changes: 2);
        new SessaoDocumento { ArquivoOrigem = @"C:\x\doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = q }.ApagarRecovery();
        Assert.False(Directory.Exists(q));
    }

    // --- I2 (pecas puras) ---------------------------------------------------------------------------------

    [Fact] // I2: carimbo do original no sessao.json — igual = nao mudou; data, tamanho, sumico ou carimbo ausente = mudou
    public void Metadados_OrigemMudou()
    {
        var original = Path.Combine(_pasta, "peticao.docx");
        File.WriteAllBytes(original, [1, 2, 3, 4]);
        var m = new MetadadosSessao(original, 65, "peticao.docx", false, "mtexto", 1).ComCarimboDe(original);
        Assert.NotNull(m.OrigemEscritaUtc); Assert.Equal(4, m.OrigemTamanho);
        var p = Path.Combine(_pasta, "s"); m.Gravar(p);
        var lido = MetadadosSessao.Ler(p)!;
        Assert.Equal(m, lido);                              // o carimbo vai e volta pelo JSON sem perder precisao
        Assert.False(lido.OrigemMudou());

        File.SetLastWriteTimeUtc(original, m.OrigemEscritaUtc!.Value.AddSeconds(5));
        Assert.True(lido.OrigemMudou());                    // so' a data
        File.SetLastWriteTimeUtc(original, m.OrigemEscritaUtc!.Value);
        Assert.False(lido.OrigemMudou());
        File.WriteAllBytes(original, [1, 2, 3, 4, 5]); File.SetLastWriteTimeUtc(original, m.OrigemEscritaUtc!.Value);
        Assert.True(lido.OrigemMudou());                    // so' o tamanho
        File.Delete(original);
        Assert.True(lido.OrigemMudou());                    // sumiu

        // sessao.json de antes desta versao (sem carimbo): legivel, e falha FECHADA (conta como "mudou").
        var q = Path.Combine(_pasta, "antigo"); Directory.CreateDirectory(q);
        File.WriteAllText(Path.Combine(q, MetadadosSessao.NomeArquivo),
            """{"Origem":"C:\\x\\a.docx","Formato":65,"Titulo":"a.docx","Novo":false,"App":"mtexto","IdUsuario":3,"Protegido":true}""");
        var antigo = MetadadosSessao.Ler(q)!;
        Assert.Null(antigo.OrigemEscritaUtc); Assert.Null(antigo.OrigemTamanho);
        Assert.True(antigo.Protegido);
        Assert.True(antigo.OrigemMudou());
        // Documento NOVO nao tem original: nunca "mudou".
        Assert.False((antigo with { Novo = true }).OrigemMudou());
    }

    [Fact] // I2: o "Salvar como" de uma sessao desvinculada sugere "<nome> (recuperado)" — aceitar o nome sugerido nunca
           // sobrescreve o original alterado. Controles: sessao normal e documento novo.
    public void NomeSugerido_DesvinculadaGanhaSufixo()
    {
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\peticao inicial.docx", Formato = FormatoDocumento.Docx, PastaRecovery = _pasta };
        Assert.Equal("peticao inicial", JanelaEditor.NomeSugerido(s));
        s.Desvinculada = true;
        Assert.Equal("peticao inicial (recuperado)", JanelaEditor.NomeSugerido(s));
        var novo = new SessaoDocumento { ArquivoOrigem = @"C:\x\new.docx", Formato = FormatoDocumento.Docx, PastaRecovery = _pasta, Novo = true };
        Assert.Equal("Documento", JanelaEditor.NomeSugerido(novo));
    }
}
