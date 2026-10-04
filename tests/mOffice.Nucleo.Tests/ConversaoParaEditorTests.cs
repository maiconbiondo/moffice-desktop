using System.Reflection;
using System.Text.Json;
using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 8 — as regras sem janela: consentimento de conversao/leitura na ponte (o caminho tem de vir do dialogo do host),
/// vinculo externo de planilha, a midia do convertido entrando no media/ do documento aberto e o id da pasta da conversao.</summary>
public class ConversaoParaEditorTests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("moffice-t8-pura-").FullName;
    private readonly List<(string Tipo, string Param)> _avisos = [];
    private readonly SessaoDocumento _sessao;
    private readonly PonteDesktop _ponte;

    public ConversaoParaEditorTests()
    {
        Directory.CreateDirectory(Path.Combine(_pasta, "docs"));
        _sessao = new SessaoDocumento
        {
            ArquivoOrigem = Path.Combine(_pasta, "docs", "aberto.xlsx"), Formato = FormatoDocumento.Xlsx,
            PastaRecovery = Directory.CreateDirectory(Path.Combine(_pasta, "recovery")).FullName,
        };
        _ponte = new PonteDesktop(_sessao, (t, p) => _avisos.Add((t, p)));
    }

    public void Dispose() { try { Directory.Delete(_pasta, true); } catch { } GC.SuppressFinalize(this); }

    private string Arquivo(string nome)
    {
        var c = Path.Combine(_pasta, "docs", nome);
        File.WriteAllBytes(c, [1, 2, 3]);
        return c;
    }

    [Fact] // caminho que nao veio do dialogo do host: recusado, nenhum aviso (o x2t nao roda)
    public void PedirConversao_SemConsentimento_Recusa()
    {
        var c = Arquivo("outro.docx");
        Assert.Equal("", _ponte.PedirConversao(c, 8193, "converter"));
        Assert.Equal("", _ponte.PedirConversao(c, 8192, "comparar"));
        Assert.Empty(_avisos);
    }

    [Fact] // consentimento de UM uso: o segundo pedido do mesmo caminho e' recusado
    public void PedirConversao_ComConsentimento_UmUso()
    {
        var c = Arquivo("outro.docx");
        _ponte.PermitirConversao(c);
        var id = _ponte.PedirConversao(c, 8193, "converter");
        Assert.Matches("^[0-9a-f]{32}$", id);
        var (tipo, param) = Assert.Single(_avisos);
        Assert.Equal("converterArquivo", tipo);
        using var d = JsonDocument.Parse(param);
        Assert.Equal((id, c, 8193, "converter"), (d.RootElement.GetProperty("id").GetString(), d.RootElement.GetProperty("caminho").GetString(),
            d.RootElement.GetProperty("formato").GetInt32(), d.RootElement.GetProperty("uso").GetString()));
        Assert.Equal("", _ponte.PedirConversao(c, 8193, "converter"));
    }

    [Fact] // uso fora de converter/comparar/combinar: recusado (e o consentimento nao e' gasto)
    public void PedirConversao_UsoDesconhecido_Recusa()
    {
        var c = Arquivo("outro.docx");
        _ponte.PermitirConversao(c);
        Assert.Equal("", _ponte.PedirConversao(c, 8193, "apagar"));
        Assert.NotEqual("", _ponte.PedirConversao(c, 8193, "comparar"));
    }

    [Fact] // conversao e leitura sao consentimentos SEPARADOS
    public void Consentimentos_Separados()
    {
        var c = Arquivo("dados.csv");
        _ponte.PermitirLeitura(c);
        Assert.Equal("", _ponte.PedirConversao(c, 8194, "comparar"));
        var id = _ponte.PedirLeitura(c);
        Assert.Matches("^[0-9a-f]{32}$", id);
        Assert.Equal("", _ponte.PedirLeitura(c));
        Assert.Equal("lerArquivo", _avisos.Single().Tipo);
        var d = Arquivo("doc.docx");
        _ponte.PermitirConversao(d);
        Assert.Equal("", _ponte.PedirLeitura(d));
    }

    /// <summary>Fix 1 da revisao da Task 8 (seguranca): caminhos de DISPOSITIVO/namespace e nao totalmente qualificados. Recusados ANTES de qualquer
    /// GetFullPath/File.Exists (nenhuma E/S), com aviso "vinculoRecusado" (faixa).</summary>
    public static TheoryData<string> CaminhosDeDispositivo => new()
    {
        @"\\?\C:\x\custos.xlsx", @"\\.\C:\x\custos.xlsx", "//?/C:/x/custos.xlsx", "//./C:/x/custos.xlsx", @"\\?\UNC\srv\p\custos.xlsx",
        @"\\.\pipe\custos.xlsx", "file:////?/C:/x/custos.xlsx", "file://./C:/x/custos.xlsx",
        @"C:custos.xlsx", @"\custos.xlsx", "custos?.xlsx", @"..\cus*tos.xlsx",
    };

    [Theory] // vinculo externo: dispositivo/namespace ou relativo -> recusado sem E/S nenhuma e sem pergunta
    [MemberData(nameof(CaminhosDeDispositivo))]
    public void PedirConversao_Vinculo_DispositivoOuRelativo_Recusado(string caminho)
    {
        Assert.Equal("", _ponte.PedirConversao(caminho, 8194, "converter"));
        Assert.Equal("", _ponte.AbrirReferenciaExterna(caminho));
        Assert.All(_avisos, a => Assert.Equal("vinculoRecusado", a.Tipo));
        Assert.Equal(2, _avisos.Count);
    }

    [Fact] // o arquivo de dispositivo recusado mesmo EXISTINDO (\\?\ + caminho real): a regra e' pela grafia, antes do disco
    public void PedirConversao_Vinculo_DispositivoDeArquivoExistente_Recusado()
    {
        var xlsx = Arquivo("origem.xlsx");
        Assert.Equal("", _ponte.PedirConversao(@"\\?\" + xlsx, 8194, "converter"));
        Assert.Equal("vinculoRecusado", _avisos.Single().Tipo);
    }

    private (string Id, string Caminho, bool Confirmar) UltimoPedido(string tipo)
    {
        var (t, p) = _avisos[^1];
        Assert.Equal(tipo, t);
        using var d = JsonDocument.Parse(p);
        return (d.RootElement.GetProperty("id").GetString()!, d.RootElement.GetProperty("caminho").GetString()!, d.RootElement.GetProperty("confirmar").GetBoolean());
    }

    [Fact] // UNC (e unidade local/mapeada): PEDE confirmacao ao host, sem tocar no caminho (um servidor que nao existe nao trava nada)
    public void PedirConversao_Vinculo_UncPedeConfirmacaoSemEs()
    {
        const string unc = @"\\moffice-servidor-inexistente.invalid\juridico\custos.xlsx";
        var relogio = System.Diagnostics.Stopwatch.StartNew();
        var id = _ponte.PedirConversao(unc, 8194, "converter");
        relogio.Stop();
        Assert.Matches("^[0-9a-f]{32}$", id);
        Assert.Equal((id, unc, true), UltimoPedido("converterArquivo"));
        Assert.True(relogio.ElapsedMilliseconds < 500, $"levou {relogio.ElapsedMilliseconds} ms: houve E/S no caminho UNC");
        // Local que NAO existe: tambem pergunta (a existencia so' e' vista DEPOIS de o usuario permitir).
        var inexistente = Path.Combine(_pasta, "nao-existe.xlsx");
        Assert.NotEqual("", _ponte.PedirConversao(inexistente, 8194, "converter"));
        Assert.True(UltimoPedido("converterArquivo").Confirmar);
        Assert.NotEqual("", _ponte.PedirConversao(new Uri(Arquivo("origem.xlsx")).AbsoluteUri, 8194, "converter"));
        Assert.Equal(Path.Combine(_pasta, "docs", "origem.xlsx"), UltimoPedido("converterArquivo").Caminho);
        // Fora da regra do vinculo: so' com o dialogo do host.
        Assert.Equal("", _ponte.PedirConversao(unc, 8193, "converter"));
        Assert.Equal("", _ponte.PedirConversao(unc, 8194, "comparar"));
        Assert.Equal("", _ponte.PedirConversao(Arquivo("texto.docx"), 8194, "converter"));
        Assert.Equal("", _ponte.PedirConversao("https://exemplo.invalid/a.xlsx", 8194, "converter"));
    }

    [Fact] // Fix 1b: vinculo relativo numa sessao SALVA resolve (so' texto) pela pasta do original e PERGUNTA mostrando o caminho resolvido
    public void Vinculo_Relativo_SessaoSalva_ResolvePelaPastaEPergunta()
    {
        var docs = Path.Combine(_pasta, "docs");
        Assert.NotEqual("", _ponte.PedirConversao("custos.xlsx", 8194, "converter"));
        Assert.Equal((Path.Combine(docs, "custos.xlsx"), true), (UltimoPedido("converterArquivo").Caminho, UltimoPedido("converterArquivo").Confirmar));
        Assert.NotEqual("", _ponte.PedirConversao(@"..\pasta\custos.xlsx", 8194, "converter"));
        Assert.Equal(Path.Combine(_pasta, "pasta", "custos.xlsx"), UltimoPedido("converterArquivo").Caminho);
        // ".." que sobe alem da raiz so' resolve: a pergunta mostra o caminho REAL.
        var raiz = Path.GetPathRoot(docs)!;
        Assert.NotEqual("", _ponte.AbrirReferenciaExterna(@"..\..\..\..\..\..\..\..\..\..\..\..\custos.xlsx"));
        Assert.Equal((Path.Combine(raiz, "custos.xlsx"), true), (UltimoPedido("abrirReferenciaExterna").Caminho, UltimoPedido("abrirReferenciaExterna").Confirmar));
        // A decisao vale pelo caminho RESOLVIDO: negar o absoluto vale para o relativo que da' nele.
        _ponte.RegistrarDecisaoDeVinculo(Path.Combine(docs, "custos.xlsx"), false);
        var antes = _avisos.Count;
        Assert.Equal("", _ponte.PedirConversao("custos.xlsx", 8194, "converter"));
        Assert.Equal(antes, _avisos.Count);
        Assert.DoesNotContain(_avisos, a => a.Tipo == "vinculoRecusado");
    }

    [Fact] // Fix 1b: documento NOVO nao tem pasta -> relativo recusado (faixa), sem E/S
    public void Vinculo_Relativo_SessaoNova_Recusado()
    {
        var avisos = new List<(string, string)>();
        var nova = new PonteDesktop(new SessaoDocumento
        {
            ArquivoOrigem = Path.Combine(_pasta, "docs", "modelo.xlsx"), Formato = FormatoDocumento.Xlsx, PastaRecovery = _sessao.PastaRecovery, Novo = true,
        }, (t, p) => avisos.Add((t, p)));
        Assert.Equal("", nova.PedirConversao("custos.xlsx", 8194, "converter"));
        Assert.Equal("", nova.AbrirReferenciaExterna(@"..\pasta\custos.xlsx"));
        Assert.Equal(["vinculoRecusado", "vinculoRecusado"], avisos.Select(a => a.Item1));
    }

    [Theory] // Fix 1b: relativo cuja pasta-base e' de DISPOSITIVO -> o resultado e' de dispositivo -> recusado
    [InlineData(@"\\?\C:\docs\aberto.xlsx")]
    [InlineData(@"\\.\C:\docs\aberto.xlsx")]
    public void Vinculo_Relativo_ResolvendoParaDispositivo_Recusado(string original)
    {
        var avisos = new List<(string, string)>();
        var p = new PonteDesktop(new SessaoDocumento { ArquivoOrigem = original, Formato = FormatoDocumento.Xlsx, PastaRecovery = _sessao.PastaRecovery },
            (t, x) => avisos.Add((t, x)));
        Assert.Equal("", p.PedirConversao("custos.xlsx", 8194, "converter"));
        Assert.Equal("", p.AbrirReferenciaExterna("custos.xlsx"));
        Assert.Equal(["vinculoRecusado", "vinculoRecusado"], avisos.Select(a => a.Item1));
    }

    [Theory] // resolucao pura (so' texto)
    [InlineData("custos.xlsx", @"C:\docs", @"C:\docs\custos.xlsx")]
    [InlineData(@"..\pasta\custos.xlsx", @"\\srv\share\docs", @"\\srv\share\pasta\custos.xlsx")]
    [InlineData("custos.xlsx", null, "custos.xlsx")]                         // sem pasta: continua relativo (CaminhoBruto recusa)
    [InlineData(@"\custos.xlsx", @"C:\docs", @"\custos.xlsx")]               // enraizado pela metade: nao e' "relativo do Excel"
    [InlineData(@"C:\x\custos.xlsx", @"C:\docs", @"C:\x\custos.xlsx")]
    [InlineData("custos?.xlsx", @"C:\docs", "custos?.xlsx")]
    public void ResolverRelativo_SoTexto(string referencia, string? pasta, string esperado) =>
        Assert.Equal(esperado, PonteDesktop.ResolverRelativo(referencia, pasta));

    [Fact] // a decisao do usuario vale por caminho EXATO na sessao: negado -> recusa sem perguntar; permitido -> nao pergunta de novo
    public void PedirConversao_Vinculo_DecisaoLembrada()
    {
        const string negado = @"\\srv\p\salarios.xlsx", permitido = @"\\srv\p\custos.xlsx";
        _ponte.RegistrarDecisaoDeVinculo(negado, false);
        _ponte.RegistrarDecisaoDeVinculo(permitido, true);
        Assert.Equal("", _ponte.PedirConversao(negado, 8194, "converter"));
        Assert.Equal("", _ponte.AbrirReferenciaExterna(negado));
        Assert.Empty(_avisos);
        Assert.NotEqual("", _ponte.PedirConversao(permitido, 8194, "converter"));
        Assert.False(UltimoPedido("converterArquivo").Confirmar);
        Assert.NotEqual("", _ponte.AbrirReferenciaExterna(permitido));
        Assert.False(UltimoPedido("abrirReferenciaExterna").Confirmar);
        // Outro caminho no mesmo servidor: pergunta.
        Assert.NotEqual("", _ponte.PedirConversao(@"\\srv\p\outra.xlsx", 8194, "converter"));
        Assert.True(UltimoPedido("converterArquivo").Confirmar);
    }

    [Fact] // "Abrir origem" do vinculo: planilha -> pedido com confirmacao (a existencia e' vista pelo host DEPOIS do sim); nao-planilha -> nada
    public void AbrirReferenciaExterna_SoPlanilhaExistente()
    {
        const string unc = @"\\moffice-servidor-inexistente.invalid\juridico\origem.xlsx";
        var id = _ponte.AbrirReferenciaExterna(unc);
        Assert.Matches("^[0-9a-f]{32}$", id);
        Assert.Equal((id, unc, true), UltimoPedido("abrirReferenciaExterna"));
        Assert.Equal("", _ponte.AbrirReferenciaExterna(Arquivo("texto.docx")));
        Assert.Single(_avisos);
    }

    [Theory] // classificacao pura (sem E/S): o que e' recusado e o que vira pergunta
    [InlineData(@"C:\x\a.xlsx", @"C:\x\a.xlsx")]
    [InlineData(@"C:\x\..\y\a.xlsx", @"C:\y\a.xlsx")]
    [InlineData(@"\\srv\p\a.xlsx", @"\\srv\p\a.xlsx")]
    [InlineData("file:///C:/x/a%20b.xlsx", @"C:\x\a b.xlsx")]
    [InlineData("file://srv/p/a.xlsx", @"\\srv\p\a.xlsx")]
    [InlineData(@"\\?\C:\x\a.xlsx", null)]
    [InlineData("//./C:/x/a.xlsx", null)]
    [InlineData("a.xlsx", null)]
    [InlineData("https://srv/a.xlsx", null)]
    public void ClassificarVinculo(string referencia, string? esperado) =>
        Assert.Equal(esperado, PonteDesktop.CaminhoDeVinculo(referencia));

    [Fact] // a pagina NAO alcanca os consentimentos: o objeto de host e' COM-visivel (todo membro publico e' chamavel pelo JS)
    public void Consentimentos_NaoSaoPublicos()
    {
        foreach (var nome in new[] { "PermitirConversao", "PermitirLeitura", "PermitirArquivoEscolhido", "RegistrarDecisaoDeVinculo", "DecisaoDeVinculo" })
            Assert.Null(typeof(PonteDesktop).GetMethod(nome, BindingFlags.Public | BindingFlags.Instance));
        foreach (var nome in new[] { "PedirConversao", "PedirLeitura", "LiberarConversao", "AbrirReferenciaExterna" })
            Assert.NotNull(typeof(PonteDesktop).GetMethod(nome, BindingFlags.Public | BindingFlags.Instance));
    }

    [Fact] // a midia do convertido entra no media/ do documento com nome novo; emf e svg do mesmo nome-base continuam juntos
    public async Task ImportarMidiaConvertida_NomesNovosEPares()
    {
        var media = Directory.CreateDirectory(Path.Combine(_sessao.PastaRecovery, "media")).FullName;
        File.WriteAllBytes(Path.Combine(media, "image1.png"), [9]);   // midia do proprio documento: nao pode ser sobrescrita
        var conv = Directory.CreateDirectory(Path.Combine(_pasta, "conv", "media")).FullName;
        File.WriteAllBytes(Path.Combine(conv, "image1.png"), [1]);
        File.WriteAllBytes(Path.Combine(conv, "image2.emf"), [2]);
        File.WriteAllBytes(Path.Combine(conv, "image2.svg"), [3]);
        File.WriteAllBytes(Path.Combine(conv, "estranho.e%x"), [4]);  // extensao que nao e' alfanumerica: fica de fora

        var mapa = await Midia.ImportarMidiaConvertidaAsync(_sessao, conv);

        Assert.Equal(3, mapa.Count);
        Assert.Equal([9], File.ReadAllBytes(Path.Combine(media, "image1.png")));
        Assert.Equal([1], File.ReadAllBytes(Path.Combine(media, mapa["media/image1.png"])));
        Assert.Equal(Path.GetFileNameWithoutExtension(mapa["media/image2.emf"]), Path.GetFileNameWithoutExtension(mapa["media/image2.svg"]));
        Assert.EndsWith(".svg", mapa["media/image2.svg"]);
        Assert.Equal([3], File.ReadAllBytes(Path.Combine(media, mapa["media/image2.svg"])));
        Assert.DoesNotContain(mapa.Values, v => v == "image1.png");
    }

    [Theory] // o id vem da pagina no close(): so' letras e digitos viram pasta
    [InlineData("..")] [InlineData(@"..\..\x")] [InlineData("a/b")] [InlineData("")]
    public void PastaDaConversao_IdInvalido(string id) =>
        Assert.Throws<ArgumentException>(() => ConversaoParaEditor.PastaDaConversao(_sessao.PastaRecovery, id));

    [Fact] // Liberar com id da pagina nunca apaga fora de convertidos\
    public void Liberar_IdInvalido_NaoApagaNada()
    {
        ConversaoParaEditor.Liberar(_sessao.PastaRecovery, "..");
        Assert.True(Directory.Exists(_sessao.PastaRecovery));
    }

    [Theory] // aviso "abrirArquivo": {filtro, multi}
    [InlineData("""{"filtro":"word","multi":true}""", "word", true)]
    [InlineData("""{"filtro":"images","multi":false}""", "images", false)]
    [InlineData("images", "images", false)]
    public void LerPedidoDeDialogo(string param, string filtro, bool multi) =>
        Assert.Equal((filtro, multi), JanelaEditor.LerPedidoDeDialogo(param));
}
