using mOffice.Comum;

namespace mOffice.Comum.Tests;

public class ConfiguracoesTests
{
    private static string Arquivo() => Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "config.json");

    [Fact] public void Padroes_ClaroESemGrampo() { var c = new Configuracoes(Arquivo()); Assert.Equal(ModoTema.Claro, c.Tema); Assert.False(c.Grampo); }

    [Fact] // set grava e um segundo objeto le o mesmo arquivo
    public void Set_PersisteNoDisco()
    {
        var a = Arquivo(); var c = new Configuracoes(a);
        c.Tema = ModoTema.Escuro; c.Grampo = true;
        var d = new Configuracoes(a);
        Assert.Equal(ModoTema.Escuro, d.Tema); Assert.True(d.Grampo);
    }

    private static string[] Ilegiveis(string a) => Directory.GetFiles(Path.GetDirectoryName(a)!, "config.json.ilegivel-*");

    [Fact] // M2 da revisao final do Plano 6: arquivo ilegivel = padroes na leitura, sem lancar; o 1o set POE O ILEGIVEL DE LADO (bytes preservados
           // em config.json.ilegivel-<data>) e recomeca dos padroes — antes o config.json corrompido nunca mais era gravado
    public void Corrompido_UsaPadroes_PoeDeLadoERecomeca()
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, "{nao e json");
        var antes = File.ReadAllBytes(a);
        var c = new Configuracoes(a);
        Assert.Equal(ModoTema.Claro, c.Tema);
        c.Tema = ModoTema.Escuro;
        c.NomeUsuario = "Maria";
        var guardado = Assert.Single(Ilegiveis(a));
        Assert.Matches(@"config\.json\.ilegivel-\d{14}$", guardado);
        Assert.Equal(antes, File.ReadAllBytes(guardado));             // bytes intactos, de lado
        var d = new Configuracoes(a);                                 // o config.json voltou a ser gravado, dos padroes + as mudancas
        Assert.Equal(ModoTema.Escuro, d.Tema); Assert.Equal("Maria", d.NomeUsuario); Assert.False(d.Grampo);
    }

    [Theory] // truncado no meio, vazio (0 byte), so' espacos, lixo binario, "null": nunca se escreve por cima — os bytes vao de lado
    [InlineData("{\"Tema\":1,\"NomeUsu")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\0\0\0\0")]
    [InlineData("null")]
    public void Ilegivel_NuncaEhSobrescrito_VaiDeLado(string conteudo)
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, conteudo);
        var antes = File.ReadAllBytes(a);
        new Configuracoes(a).FormatoOdf = true;
        Assert.Equal(antes, File.ReadAllBytes(Assert.Single(Ilegiveis(a))));
        Assert.True(new Configuracoes(a).FormatoOdf);
    }

    [Fact] // preso por outro processo (inacessivel, possivelmente passageiro): nao grava, nao poe de lado, o valor vale na memoria
    public void Preso_NaoGrava_ValorNaMemoria()
    {
        var a = Arquivo(); var c = new Configuracoes(a); c.NomeUsuario = "Ana";
        var antes = File.ReadAllBytes(a);
        using (new FileStream(a, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            c.Tema = ModoTema.Escuro;
            Assert.Equal(ModoTema.Escuro, c.Tema);
        }
        Assert.Equal(antes, File.ReadAllBytes(a));
        Assert.Empty(Ilegiveis(a));
    }

    [Fact] // I2 da revisao final do Plano 6: config.json SOMENTE LEITURA — o set nao lanca (o app caia), loga, e o valor fica na memoria e
           // prevalece sobre o arquivo antigo ate' uma gravacao dar certo
    public void SomenteLeitura_SetNaoLanca_ValorPrevalece()
    {
        var a = Arquivo(); var c = new Configuracoes(a); c.NomeUsuario = "Ana";
        var antes = File.ReadAllBytes(a);
        var logs = new List<string>(); var logAntes = Configuracoes.Log; Configuracoes.Log = l => { lock (logs) logs.Add(l); };
        File.SetAttributes(a, FileAttributes.ReadOnly);
        try
        {
            c.Tema = ModoTema.Escuro;                        // nao lanca
            c.TemaEditorId = "theme-night";
            Assert.Equal(ModoTema.Escuro, c.Tema);           // vale neste processo
            Assert.Equal("theme-night", c.TemaEditorId);
            Assert.Equal("Ana", c.NomeUsuario);
            Assert.Equal(antes, File.ReadAllBytes(a));       // arquivo intacto
            lock (logs) Assert.Contains(logs, l => l.StartsWith("CONFIG config.json nao gravado", StringComparison.Ordinal));
        }
        finally { File.SetAttributes(a, FileAttributes.Normal); Configuracoes.Log = logAntes; }
        c.Grampo = true;                                     // gravou: tudo o que estava na memoria chega ao disco
        var d = new Configuracoes(a);
        Assert.Equal(ModoTema.Escuro, d.Tema); Assert.Equal("theme-night", d.TemaEditorId); Assert.True(d.Grampo); Assert.Equal("Ana", d.NomeUsuario);
    }

    [Fact] // I2: posicao-<app>.json somente leitura — o fechamento da janela nao lanca
    public void Posicao_SomenteLeitura_NaoLanca()
    {
        var a = Arquivo(); var c = new Configuracoes(a);
        c.DefinirPosicao("mtexto", new PosicaoSalva(1, 2, 3, 4, false));
        var arq = Path.Combine(Path.GetDirectoryName(a)!, "posicao-mtexto.json");
        File.SetAttributes(arq, FileAttributes.ReadOnly);
        try { c.DefinirPosicao("mtexto", new PosicaoSalva(5, 6, 7, 8, true)); }
        finally { File.SetAttributes(arq, FileAttributes.Normal); }
        Assert.Equal(new PosicaoSalva(1, 2, 3, 4, false), c.Posicao("mtexto"));
    }

    [Fact] // dois escritores (instancias/threads) em campos diferentes, N vezes: os dois campos sobrevivem e nada vira padrao
    public async Task EscritaConcorrente_CamposDiferentes_Sobrevivem()
    {
        var a = Arquivo();
        var semente = new Configuracoes(a) { PastaModelos = @"D:\modelos", TemaEditorId = "theme-dark" };
        const int N = 60;
        var t1 = Task.Run(() => { var c = new Configuracoes(a); for (var i = 1; i <= N; i++) c.NomeUsuario = "autor" + i; });
        var t2 = Task.Run(() => { var c = new Configuracoes(a); for (var i = 1; i <= N; i++) c.PastaPadraoSalvar = @"C:\pasta" + i; });
        var t3 = Task.Run(() => { var c = new Configuracoes(a); for (var i = 1; i <= N; i++) c.Tema = i % 2 == 0 ? ModoTema.Escuro : ModoTema.Claro; });
        await Task.WhenAll(t1, t2, t3);
        var d = new Configuracoes(a);
        Assert.Equal("autor" + N, d.NomeUsuario);
        Assert.Equal(@"C:\pasta" + N, d.PastaPadraoSalvar);
        Assert.Equal(ModoTema.Escuro, d.Tema);
        Assert.Equal(@"D:\modelos", d.PastaModelos);      // o que ninguem tocou nao se perde
        Assert.Equal("theme-dark", d.TemaEditorId);
    }

    [Fact] // gravacao atomica: nenhum temporario sobra na pasta
    public void Gravar_NaoDeixaTemporarios()
    {
        var a = Arquivo(); var c = new Configuracoes(a);
        for (var i = 0; i < 20; i++) c.NomeUsuario = "n" + i;
        c.DefinirPosicao("mtexto", new PosicaoSalva(1, 2, 3, 4, false));
        var arquivos = Directory.GetFiles(Path.GetDirectoryName(a)!).Select(Path.GetFileName).OrderBy(x => x).ToList();
        Assert.Equal(["config.json", "posicao-mtexto.json"], arquivos);
    }

    [Fact] // SO-07: a posicao vai para um arquivo POR APP (nunca no config.json compartilhado); config antigo com Posicoes dentro continua lendo (campo ignorado)
    public void Posicao_ArquivoPorApp_EConfigAntigoCompativel()
    {
        var a = Arquivo(); var c = new Configuracoes(a);
        c.Tema = ModoTema.Escuro;
        var configAntes = File.ReadAllBytes(a);
        c.DefinirPosicao("mtexto", new PosicaoSalva(10, 20, 810, 620, true));
        Assert.Equal(configAntes, File.ReadAllBytes(a));
        Assert.True(File.Exists(Path.Combine(Path.GetDirectoryName(a)!, "posicao-mtexto.json")));
        Assert.Null(c.Posicao("mslides"));
        Assert.Equal(new PosicaoSalva(10, 20, 810, 620, true), new Configuracoes(a).Posicao("mtexto"));

        var b = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(b)!);
        File.WriteAllText(b, """{"Tema":2,"Posicoes":{"mtexto":{"Esquerda":1,"Topo":2,"Direita":3,"Fundo":4,"Maximizada":false}}}""");
        var velho = new Configuracoes(b);
        Assert.Equal(ModoTema.Escuro, velho.Tema);
        Assert.Null(velho.Posicao("mtexto"));                       // ignorado (o proximo fechamento grava o novo arquivo)
    }

    [Fact] // posicao ilegivel: Posicao devolve null e DefinirPosicao a substitui por inteiro (arquivo so' dela); o config nao e' tocado
    public void Posicao_Ilegivel_NaoAfetaOConfig()
    {
        var a = Arquivo(); var c = new Configuracoes(a); c.NomeUsuario = "Ana";
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(a)!, "posicao-mtexto.json"), "{trunc");
        Assert.Null(c.Posicao("mtexto"));
        c.DefinirPosicao("mtexto", new PosicaoSalva(1, 2, 3, 4, false));
        Assert.Equal(new PosicaoSalva(1, 2, 3, 4, false), c.Posicao("mtexto"));
        Assert.Equal("Ana", c.NomeUsuario);
    }

    [Fact] // config antiga sem o campo Grampo continua lendo (sem migracao)
    public void CampoAusente_CaiNoPadrao()
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, "{\"Tema\":2}");
        var c = new Configuracoes(a);
        Assert.Equal(ModoTema.Escuro, c.Tema); Assert.False(c.Grampo);
    }

    [Fact] // Plano 6, Task 4: os campos novos nascem nulos/false, gravam e outro objeto le
    public void NovosCampos_PadroesEPersistencia()
    {
        var a = Arquivo(); var c = new Configuracoes(a);
        Assert.Null(c.NomeUsuario); Assert.Null(c.PastaPadraoSalvar); Assert.False(c.FormatoOdf); Assert.Null(c.PastaModelos); Assert.Null(c.TemaEditorId);
        c.NomeUsuario = "Dra. Ana"; c.PastaPadraoSalvar = @"D:\Docs"; c.FormatoOdf = true; c.PastaModelos = @"D:\Mod"; c.TemaEditorId = "theme-night";
        var d = new Configuracoes(a);
        Assert.Equal("Dra. Ana", d.NomeUsuario); Assert.Equal(@"D:\Docs", d.PastaPadraoSalvar); Assert.True(d.FormatoOdf);
        Assert.Equal(@"D:\Mod", d.PastaModelos); Assert.Equal("theme-night", d.TemaEditorId);
        c.NomeUsuario = "   "; Assert.Null(new Configuracoes(a).NomeUsuario); // em branco = voltar ao padrao
    }

    [Fact] // config.json de antes do Plano 6 continua lendo; os campos novos caem nos padroes
    public void ConfigAntigaSemCamposNovos_CaiNosPadroes()
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!); File.WriteAllText(a, "{\"Tema\":1,\"Grampo\":false}");
        var c = new Configuracoes(a);
        Assert.Equal(ModoTema.Claro, c.Tema); Assert.Null(c.NomeUsuario); Assert.False(c.FormatoOdf); Assert.Null(c.TemaEditorId);
    }

    [Fact] // sem valor, a pasta de modelos e' a do Office
    public void PastaModelosEfetiva_PadraoTemplatesDoOffice()
    {
        var c = new Configuracoes(Arquivo());
        Assert.Equal(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates"), c.PastaModelosEfetiva);
        c.PastaModelos = @"D:\Mod"; Assert.Equal(@"D:\Mod", c.PastaModelosEfetiva);
    }

    [Fact] // atomicidade: um leitor SEM mutex nunca ve o arquivo truncado/vazio enquanto outro o reescreve (o que apagava as configuracoes no logoff)
    public async Task Escrita_EhAtomica_LeitorNuncaVeMeio()
    {
        var a = Arquivo(); Directory.CreateDirectory(Path.GetDirectoryName(a)!);
        var payload = new string('x', 400_000);
        ArquivoAtomico.Escrever(a, payload);
        var parar = false; var vistos = new List<int>();
        var leitor = Task.Run(() =>
        {
            while (!Volatile.Read(ref parar))
            {
                try { vistos.Add(ArquivoAtomico.Ler(a).Length); }
                catch (IOException) { vistos.Add(-1); }
            }
        });
        for (var i = 0; i < 150; i++) ArquivoAtomico.Escrever(a, payload);
        Volatile.Write(ref parar, true); await leitor;
        Assert.NotEmpty(vistos);
        Assert.All(vistos, n => Assert.Equal(payload.Length, n));
    }
}
