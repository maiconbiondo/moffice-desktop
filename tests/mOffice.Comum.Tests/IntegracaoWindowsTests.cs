using System.Text.RegularExpressions;
using System.Windows.Shell;
using mOffice.Comum;
using mOffice.Conversor;

namespace mOffice.Comum.Tests;

public class IntegracaoWindowsTests
{
    private static string Raiz => AplicativoOfficeTests.RaizRepo();
    private static string Pasta() { var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(p); return p; }

    [Fact] // Jump List: recentes do Windows ligados; os fixados numa categoria "Fixados", na ordem da lista
    public void JumpList_FixadosNaCategoriaPropria()
    {
        var itens = new[] { new ItemRecente(@"C:\x\a.docx", true, DateTimeOffset.Now), new ItemRecente(@"C:\x\b.docx", false, DateTimeOffset.Now) };
        var jl = IntegracaoWindows.MontarJumpList(itens, @"C:\Program Files\mOffice\mTexto.exe");
        Assert.True(jl.ShowRecentCategory);
        Assert.False(jl.ShowFrequentCategory);
        var fixados = jl.JumpItems.OfType<JumpPath>().Where(p => p.CustomCategory == "Fixados").Select(p => p.Path).ToList();
        Assert.Equal([@"C:\x\a.docx"], fixados);
        Assert.Single(jl.JumpItems); // o nao-fixado fica so' nos Recentes do Windows
    }

    [Fact] // IMPORTANT 2: fixado em UNC (share morto) NAO entra na Jump List — o Apply resolve o item pelo shell na thread da tela e travava ~14 s (medido)
    public void JumpList_FixadoEmRedeOuIndisponivel_FicaDeFora_ESemEspera()
    {
        var itens = new[]
        {
            new ItemRecente(@"\\10.255.255.1\x\a.docx", true, DateTimeOffset.Now),
            new ItemRecente(@"C:\x\ok.docx", true, DateTimeOffset.Now),
            new ItemRecente(@"\\?\UNC\10.255.255.1\x\b.docx", true, DateTimeOffset.Now),
        };
        var t0 = System.Diagnostics.Stopwatch.StartNew();
        var jl = IntegracaoWindows.MontarJumpList(itens, @"C:\Program Files\mOffice\mTexto.exe");
        Assert.True(t0.ElapsedMilliseconds < 1000, $"montar levou {t0.ElapsedMilliseconds} ms");
        Assert.Equal([@"C:\x\ok.docx"], jl.JumpItems.OfType<JumpPath>().Select(p => p.Path).ToList());
    }

    // ---- Verificador de recente -----------------------------------------------------------------------------

    [Fact] // caminho de rede que nao responde no limite: SemResposta (nunca "nao existe"), e a interface nao espera mais que o limite
    public async Task Verificar_SemResposta_NoLimite()
    {
        var t0 = System.Diagnostics.Stopwatch.StartNew();
        var r = await VerificadorDeRecente.VerificarAsync(@"\\srv\p\a.docx", TimeSpan.FromMilliseconds(200), _ => { Thread.Sleep(2000); return true; });
        Assert.Equal(EstadoDoRecente.SemResposta, r);
        Assert.True(t0.ElapsedMilliseconds < 1500, $"esperou {t0.ElapsedMilliseconds} ms: passou do limite");
    }

    [Fact]
    public async Task Verificar_NaoExiste() =>
        Assert.Equal(EstadoDoRecente.NaoExiste, await VerificadorDeRecente.VerificarAsync(@"C:\x\a.docx", TimeSpan.FromSeconds(2), _ => false));

    [Fact]
    public async Task Verificar_Existe() =>
        Assert.Equal(EstadoDoRecente.Existe, await VerificadorDeRecente.VerificarAsync(@"C:\x\a.docx", TimeSpan.FromSeconds(2), _ => true));

    [Fact] // erro de E/S ao sondar (rede caiu no meio) e' "sem resposta", nunca "nao existe" (nao se tira o recente por isso)
    public async Task Verificar_ErroDeES_SemResposta() =>
        Assert.Equal(EstadoDoRecente.SemResposta, await VerificadorDeRecente.VerificarAsync(@"\\srv\p\a.docx", TimeSpan.FromSeconds(2), _ => throw new IOException("rede")));

    [Fact] // a sonda real: arquivo que existe / que nao existe
    public async Task Verificar_ArquivoReal()
    {
        var p = Pasta(); var a = Path.Combine(p, "a.docx"); File.WriteAllText(a, "x");
        Assert.Equal(EstadoDoRecente.Existe, await VerificadorDeRecente.VerificarAsync(a, TimeSpan.FromSeconds(2)));
        Assert.Equal(EstadoDoRecente.NaoExiste, await VerificadorDeRecente.VerificarAsync(Path.Combine(p, "b.docx"), TimeSpan.FromSeconds(2)));
    }

    // ---- Pastas do dialogo Abrir ---------------------------------------------------------------------------

    [Fact] // ate' 5 pastas DISTINTAS, so' local ou UNC bem formado, na ordem dos recentes, sem tocar o disco
    public void PastasDosRecentes_CincoDistintasBemFormadas()
    {
        var caminhos = new[]
        {
            @"C:\a\1.docx", @"c:\A\2.docx",           // mesma pasta (sem diferenciar maiusculas)
            @"\\srv\comp\b\3.docx", @"\\srv\3.docx",  // UNC bem formado / sem compartilhamento
            @"relativo\4.docx", @"\\?\C:\x\5.docx", "",
            @"D:\c\6.docx", @"E:\d\7.docx", @"F:\e\8.docx", @"G:\f\9.docx", @"H:\g\10.docx",
        };
        var pastas = AplicativoOffice.PastasDosRecentes(caminhos);
        Assert.Equal([@"C:\a", @"\\srv\comp\b", @"D:\c", @"E:\d", @"F:\e"], pastas);
    }

    // ---- Imprimir (verbo do Explorer) ----------------------------------------------------------------------

    [Fact] // o PDF existe quando o envio e' chamado, e foi apagado no fim; 0 = ok
    public async Task Imprimir_PdfExisteNoEnvio_ApagadoNoFim()
    {
        var p = Pasta(); var doc = Path.Combine(p, "a.docx"); File.WriteAllText(doc, "x");
        string? pdfNoEnvio = null; var existiaNoEnvio = false; string? appRecebido = null;
        var codigo = await ImpressaoDireta.ImprimirAsync(doc, "mtexto",
            enviar: (pdf, app) => { pdfNoEnvio = pdf; appRecebido = app; existiaNoEnvio = File.Exists(pdf); return Task.CompletedTask; },
            converter: (origem, destino) => { Assert.Equal(doc, origem); File.WriteAllText(destino, "%PDF"); return Task.CompletedTask; },
            pastaImpressao: Path.Combine(p, "imp"));
        Assert.Equal(0, codigo);
        Assert.True(existiaNoEnvio); Assert.Equal("mtexto", appRecebido);
        Assert.False(File.Exists(pdfNoEnvio!));
    }

    [Fact] // envio que falha: 20 e o PDF sai do mesmo jeito (nao fica lixo)
    public async Task Imprimir_EnvioFalha_Codigo20_ApagaPdf()
    {
        var p = Pasta(); var doc = Path.Combine(p, "a.docx"); File.WriteAllText(doc, "x");
        string? pdfNoEnvio = null; var log = "";
        var codigo = await ImpressaoDireta.ImprimirAsync(doc, "mtexto",
            enviar: (pdf, _) => { pdfNoEnvio = pdf; throw new InvalidOperationException("sem impressora"); },
            converter: (_, destino) => { File.WriteAllText(destino, "%PDF"); return Task.CompletedTask; },
            pastaImpressao: Path.Combine(p, "imp"), log: l => log += l);
        Assert.Equal(20, codigo);
        Assert.False(File.Exists(pdfNoEnvio!));
        Assert.Contains("sem impressora", log);
    }

    [Fact] // conversao que falha (arquivo protegido/corrompido): 20, nunca chama o envio
    public async Task Imprimir_ConversaoFalha_Codigo20_SemEnvio()
    {
        var p = Pasta(); var doc = Path.Combine(p, "a.docx"); File.WriteAllText(doc, "x"); var enviou = false;
        var codigo = await ImpressaoDireta.ImprimirAsync(doc, "mtexto", enviar: (_, _) => { enviou = true; return Task.CompletedTask; },
            converter: (_, _) => throw new InvalidOperationException("x2t falhou"), pastaImpressao: Path.Combine(p, "imp"));
        Assert.Equal(20, codigo); Assert.False(enviou);
        Assert.Empty(Directory.GetFiles(Path.Combine(p, "imp"), "*.pdf"));
    }

    [Fact] // menor: falha ao imprimir AVISA em pt-BR antes de sair com 20 (o Office pergunta; nao falha calado) e o sucesso nao avisa
    public async Task Imprimir_Falha_Avisa_Sucesso_NaoAvisa()
    {
        var p = Pasta(); var doc = Path.Combine(p, "Contrato.docx"); File.WriteAllText(doc, "x");
        var avisos = new List<string>();
        Task Ok(string a, string b) => Task.CompletedTask;
        Task Converte(string o, string d) { File.WriteAllText(d, "%PDF"); return Task.CompletedTask; }
        Assert.Equal(0, await ImpressaoDireta.ImprimirAsync(doc, "mtexto", Ok, Converte, Path.Combine(p, "imp"), avisar: avisos.Add));
        Assert.Empty(avisos);
        Assert.Equal(20, await ImpressaoDireta.ImprimirAsync(doc, "mtexto", (_, _) => throw new InvalidOperationException("A impressora \"X\" não está disponível."), Converte, Path.Combine(p, "imp"), avisar: avisos.Add));
        Assert.Equal(20, await ImpressaoDireta.ImprimirAsync(doc, "mtexto", Ok, (_, _) => throw new ErroConversao("falha", 1, ""), Path.Combine(p, "imp"), avisar: avisos.Add));
        Assert.Equal(20, await ImpressaoDireta.ImprimirAsync(Path.Combine(p, "sumiu.docx"), "mtexto", Ok, Converte, Path.Combine(p, "imp"), avisar: avisos.Add));
        Assert.Equal(3, avisos.Count);
        Assert.Contains("Contrato.docx", avisos[0]); Assert.Contains("não está disponível", avisos[0]);
        Assert.Contains("senha", avisos[1]);          // conversao falhou: protegido, em uso ou danificado
        Assert.Contains("sumiu.docx", avisos[2]); Assert.Contains("não foi encontrado", avisos[2]);
        Assert.All(avisos, a => Assert.StartsWith("Não foi possível imprimir", a));
    }

    [Fact] // arquivo que nao existe: 20 sem converter nada
    public async Task Imprimir_ArquivoAusente_Codigo20() =>
        Assert.Equal(20, await ImpressaoDireta.ImprimirAsync(Path.Combine(Pasta(), "nao.docx"), "mtexto", enviar: (_, _) => Task.CompletedTask));

    // ---- Icones e registro gerados -------------------------------------------------------------------------

    private static readonly string[] Treze =
    [
        "mtexto-documento", "mtexto-modelo", "mtexto-macro", "mtexto-legado",
        "mplanilha-documento", "mplanilha-modelo", "mplanilha-macro", "mplanilha-legado",
        "mslides-documento", "mslides-modelo", "mslides-macro", "mslides-legado", "mslides-apresentacao",
    ];

    [Fact] // os 13 .ico existem, com 6 quadros PNG 16/24/32/48/64/256 cada, e o csproj os copia para a saida
    public void IconesPorTipo_Treze_SeisQuadrosCada()
    {
        var pasta = Path.Combine(Raiz, "src", "mOffice.Comum", "Icones", "Documentos");
        Assert.Equal(Treze.Order(), Directory.GetFiles(pasta, "*.ico").Select(f => Path.GetFileNameWithoutExtension(f)).Order());
        foreach (var nome in Treze)
        {
            var b = File.ReadAllBytes(Path.Combine(pasta, nome + ".ico"));
            Assert.Equal(1, BitConverter.ToUInt16(b, 2));   // tipo = icone
            Assert.Equal(6, BitConverter.ToUInt16(b, 4));
            var tamanhos = Enumerable.Range(0, 6).Select(i => b[6 + 16 * i] == 0 ? 256 : (int)b[6 + 16 * i]).Order().ToArray();
            Assert.Equal([16, 24, 32, 48, 64, 256], tamanhos);
            for (var i = 0; i < 6; i++)   // cada quadro e' um PNG
            {
                var off = BitConverter.ToInt32(b, 6 + 16 * i + 12);
                Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, b[off..(off + 4)]);
            }
        }
        var csproj = File.ReadAllText(Path.Combine(Raiz, "src", "mOffice.Comum", "mOffice.Comum.csproj"));
        Assert.Contains(@"Include=""Icones\Documentos\*.ico""", csproj);
        Assert.Contains("CopyToOutputDirectory=\"PreserveNewest\"", csproj);
        Assert.Contains("CopyToPublishDirectory=\"PreserveNewest\"", csproj);
    }

    [Fact] // todo icone que o registro gerado usa existe em disco (nada de DefaultIcon apontando para arquivo que nao veio) e cada papel tem o seu
    public void IconesGerados_ExistemEPapeisCertos()
    {
        var pasta = Path.Combine(Raiz, "src", "mOffice.Comum", "Icones", "Documentos");
        foreach (var i in TabelaFormatos.Todas.Where(i => i.ProgId is not null))
            Assert.True(File.Exists(Path.Combine(pasta, GeradorAssociacoes.ArquivoIcone(i))), $"{i.Extensao}: falta {GeradorAssociacoes.ArquivoIcone(i)}");
        Assert.Equal("apresentacao", GeradorAssociacoes.Papel(TabelaFormatos.DaExtensao(".ppsx")!));
        Assert.Equal("modelo", GeradorAssociacoes.Papel(TabelaFormatos.DaExtensao(".dotx")!));
        Assert.Equal("macro", GeradorAssociacoes.Papel(TabelaFormatos.DaExtensao(".docm")!));
        Assert.Equal("legado", GeradorAssociacoes.Papel(TabelaFormatos.DaExtensao(".doc")!));
        Assert.Equal("documento", GeradorAssociacoes.Papel(TabelaFormatos.DaExtensao(".docx")!));
        Assert.Equal("mslides-apresentacao.ico", GeradorAssociacoes.ArquivoIcone(TabelaFormatos.DaExtensao(".ppsx")!));
    }

    [Fact] // ShellNew: so' .docx/.xlsx/.pptx, com o modelo vazio do pacote; o valor sai na desinstalacao e a chave so' se vazia (compartilhada com o Office)
    public void ShellNew_TresExtensoes_SemUninsdeletekey()
    {
        var reg = File.ReadAllLines(Path.Combine(Raiz, "tools", "installer", "associacoes.iss")).Where(l => l.Contains("\\ShellNew\"")).ToList();
        Assert.Equal(6, reg.Count);
        foreach (var ext in new[] { ".docx", ".xlsx", ".pptx" })
        {
            Assert.Contains(reg, l => l.Contains($"Software\\Classes\\{ext}\\ShellNew\"") && l.Contains("ValueName: \"FileName\"")
                && l.Contains($"ValueData: \"{{app}}\\Modelos\\pt-BR\\new{ext}\"") && l.Contains("Flags: uninsdeletevalue;") && l.EndsWith("Tasks: associar"));
            Assert.Contains(reg, l => l.Contains($"Software\\Classes\\{ext}\\ShellNew\"") && l.Contains("Flags: uninsdeletekeyifempty;"));
        }
        Assert.DoesNotContain(reg, l => Regex.IsMatch(l, @"uninsdeletekey(?!ifempty)"));
        Assert.DoesNotContain(reg, l => l.Contains(".docm") || l.Contains(".dotx") || l.Contains(".odt"));
    }

    [Fact] // o verbo Imprimir existe em TODO ProgID gerado (contagem = linhas da tabela com ProgID), sempre preso a task associar
    public void VerboImprimir_EmTodoProgId()
    {
        var linhas = File.ReadAllLines(Path.Combine(Raiz, "tools", "installer", "associacoes.iss")).Where(l => l.Contains("\\shell\\print\\command\"")).ToList();
        Assert.Equal(TabelaFormatos.Todas.Count(i => i.ProgId is not null), linhas.Count);
        Assert.All(linhas, l => { Assert.Contains("/imprimir", l); Assert.EndsWith("Tasks: associar", l); });
    }
}
