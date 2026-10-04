using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Media;
using mOffice.Comum.Atualizacao;
using mOffice.Comum.Dialogos;
using mOffice.Nucleo.Tests;
using Button = System.Windows.Controls.Button;

namespace mOffice.Comum.Tests;

/// <summary>Fonte de releases FALSA: nenhum teste deste arquivo toca a rede (o repositorio publico nem existe
/// ainda). Conta as chamadas para provar que abrir o dialogo nao consulta nada.</summary>
internal sealed class FonteDeReleasesFalsa : IFonteDeReleases
{
    private readonly ReleaseMaisRecente? _release;
    private readonly Exception? _erro;
    public int Chamadas { get; private set; }

    public FonteDeReleasesFalsa(ReleaseMaisRecente? release, Exception? erro = null) { _release = release; _erro = erro; }

    public Task<ReleaseMaisRecente?> ObterUltimaAsync(CancellationToken ct)
    {
        Chamadas++;
        return _erro is not null ? Task.FromException<ReleaseMaisRecente?>(_erro) : Task.FromResult(_release);
    }
}

internal static class ReleasesDeTeste
{
    public const string ShaValido = "058cf405f778fc15284646e9d7ad8377171681366424a448b588a17bb4a1c813";

    public static string TagAtual => "v" + ServicoAtualizacao.VersaoAtualTexto();

    public static string VersaoNova
    {
        get { var v = ServicoAtualizacao.VersaoAtual(); return $"{v.Major}.{v.Minor}.{v.Build + 1}"; }
    }

    /// <summary>Versao nova com asset em <c>example.invalid</c>: fora da lista de hosts, o download e' RECUSADO
    /// antes de qualquer I/O — nenhum teste chega perto da rede.</summary>
    public static ReleaseMaisRecente Nova(string? corpo = null, string? url = "https://example.invalid/mOffice-Setup.exe",
        string? nome = "mOffice-Setup.exe") =>
        new("v" + VersaoNova, corpo ?? $"Notas da versão.\n\nSHA256: {ShaValido}", nome, url, 123456);

    public static ReleaseMaisRecente Atual() => new(TagAtual, "notas", "x.exe", "https://github.com/x.exe", 1);
}

// =========================================================================================================
// ServicoAtualizacao (porte de UpdateServiceTests do mPDF)
// =========================================================================================================

public class ServicoAtualizacaoTests
{
    // ---- verificar (fonte falsa) ------------------------------------------------------------------------

    [Fact]
    public async Task Verificar_TagMaisNova_Disponivel()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(ReleasesDeTeste.Nova()));
        var r = await s.VerificarAsync();
        Assert.Equal(StatusVerificacao.Disponivel, r.Status);
        Assert.Equal("v" + ReleasesDeTeste.VersaoNova, r.Info!.TagVersao);
        Assert.Equal(ReleasesDeTeste.ShaValido, r.Info.Sha256Esperado);
        Assert.Equal("mOffice-Setup.exe", r.Info.NomeAsset);
    }

    [Fact]
    public async Task Verificar_TagIgual_Atualizado()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(ReleasesDeTeste.Atual()));
        Assert.Equal(StatusVerificacao.Atualizado, (await s.VerificarAsync()).Status);
    }

    [Fact]
    public async Task Verificar_TagMaisVelha_Atualizado()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(new ReleaseMaisRecente("v0.0.1", "x", "a.exe", "https://github.com/a.exe", 1)));
        Assert.Equal(StatusVerificacao.Atualizado, (await s.VerificarAsync()).Status);
    }

    [Fact] // erro de rede: mensagem pt-BR SEM o detalhe tecnico da excecao
    public async Task Verificar_FalhaDeRede_ErroSemRede_SemDetalheTecnico()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null, new HttpRequestException("No such host is known (xyz.detalhe)")));
        var r = await s.VerificarAsync();
        Assert.Equal(StatusVerificacao.Erro, r.Status);
        Assert.Equal(TipoErroAtualizacao.SemRede, r.TipoErro);
        Assert.DoesNotContain("xyz", r.MensagemErro);
        Assert.Contains("internet", r.MensagemErro);
    }

    [Theory] // 403 e 429 sao limite de taxa (mensagem propria), nao "confira a internet"
    [InlineData(HttpStatusCode.Forbidden)]
    [InlineData(HttpStatusCode.TooManyRequests)]
    public async Task Verificar_403e429_LimiteTaxa(HttpStatusCode status)
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null, new HttpRequestException("rate", null, status)));
        var r = await s.VerificarAsync();
        Assert.Equal(TipoErroAtualizacao.LimiteTaxa, r.TipoErro);
        Assert.Equal(ServicoAtualizacao.MensagemLimiteTaxa, r.MensagemErro);
    }

    // ---- M-1 (revisao final): varredura dos downloads abandonados ------------------------------------------

    private static string PastaDeDownload(string raiz, string nome, TimeSpan idade, bool comArquivo = true)
    {
        var p = Directory.CreateDirectory(Path.Combine(raiz, nome)).FullName;
        if (comArquivo) File.WriteAllText(Path.Combine(p, "mOffice-Setup.exe"), "instalador");
        var quando = DateTime.UtcNow - idade;
        Directory.SetCreationTimeUtc(p, quando); Directory.SetLastWriteTimeUtc(p, quando);
        return p;
    }

    [Fact] // apaga so' update-* com mais de 1 dia; poupa o recente, o "em uso", o que nao e' update-* e o com subpasta
    public void VarrerDownloadsAntigos_SoOsVelhos()
    {
        var raiz = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "moffice-varredura-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            var velha = PastaDeDownload(raiz, "update-velha", TimeSpan.FromDays(3));
            var vazia = PastaDeDownload(raiz, "update-vazia", TimeSpan.FromDays(3), comArquivo: false);
            var recente = PastaDeDownload(raiz, "update-recente", TimeSpan.FromHours(2));
            var emUso = PastaDeDownload(raiz, "update-emuso", TimeSpan.FromDays(3));
            var outra = PastaDeDownload(raiz, "x2t", TimeSpan.FromDays(3));
            var comSub = PastaDeDownload(raiz, "update-comsub", TimeSpan.FromDays(3));
            Directory.CreateDirectory(Path.Combine(comSub, "sub"));
            Directory.SetLastWriteTimeUtc(comSub, DateTime.UtcNow.AddDays(-3));

            Assert.Equal(2, ServicoAtualizacao.VarrerDownloadsAntigos(raiz, TimeSpan.FromDays(1), emUso));

            Assert.False(Directory.Exists(velha)); Assert.False(Directory.Exists(vazia));
            Assert.True(Directory.Exists(recente)); Assert.True(Directory.Exists(emUso));
            Assert.True(Directory.Exists(outra)); Assert.True(Directory.Exists(comSub));
        }
        finally { Directory.Delete(raiz, recursive: true); }
    }

    [Fact] // nunca segue link: um update-* que e' JUNCTION para outra pasta nao e' tocado, nem o que ha' do outro lado
    public void VarrerDownloadsAntigos_NaoSegueJunction()
    {
        var raiz = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "moffice-varredura-" + Guid.NewGuid().ToString("N"))).FullName;
        var alvo = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "moffice-alvo-" + Guid.NewGuid().ToString("N"))).FullName;
        var link = Path.Combine(raiz, "update-link");
        try
        {
            File.WriteAllText(Path.Combine(alvo, "documento-importante.docx"), "nao apague");
            var mk = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("cmd.exe", $"/c mklink /J \"{link}\" \"{alvo}\"")
                { CreateNoWindow = true, UseShellExecute = false })!;
            mk.WaitForExit();
            Assert.True(Directory.Exists(link) && new DirectoryInfo(link).Attributes.HasFlag(FileAttributes.ReparsePoint), "junction nao criada");
            Directory.SetLastWriteTimeUtc(alvo, DateTime.UtcNow.AddDays(-3)); Directory.SetCreationTimeUtc(alvo, DateTime.UtcNow.AddDays(-3));

            Assert.Equal(0, ServicoAtualizacao.VarrerDownloadsAntigos(raiz, TimeSpan.FromDays(1)));
            Assert.True(Directory.Exists(link));
            Assert.True(File.Exists(Path.Combine(alvo, "documento-importante.docx")));
        }
        finally
        {
            if (Directory.Exists(link)) Directory.Delete(link); // remove so' o link
            Directory.Delete(raiz, recursive: true);
            Directory.Delete(alvo, recursive: true);
        }
    }

    [Fact] // raiz inexistente: nada, sem lancar
    public void VarrerDownloadsAntigos_RaizInexistente() =>
        Assert.Equal(0, ServicoAtualizacao.VarrerDownloadsAntigos(Path.Combine(Path.GetTempPath(), "nao-existe-" + Guid.NewGuid().ToString("N")), TimeSpan.FromDays(1)));

    [Fact]
    public async Task Verificar_ReleaseNula_Erro()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null));
        Assert.Equal(StatusVerificacao.Erro, (await s.VerificarAsync()).Status);
    }

    [Fact]
    public async Task Verificar_SemExe_ErroSemAsset()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(ReleasesDeTeste.Nova(url: null, nome: null)));
        Assert.Equal(TipoErroAtualizacao.SemAsset, (await s.VerificarAsync()).TipoErro);
    }

    [Fact]
    public async Task Verificar_CorpoSemHash_ErroSemHash()
    {
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(ReleasesDeTeste.Nova(corpo: "notas sem hash")));
        Assert.Equal(TipoErroAtualizacao.SemHash, (await s.VerificarAsync()).TipoErro);
    }

    // ---- versao -----------------------------------------------------------------------------------------

    [Theory]
    [InlineData("1.10.0", "1.9.0", true)]  // numerico: string ordenaria ao contrario
    [InlineData("1.9.0", "1.10.0", false)]
    [InlineData("2.0.0", "1.99.99", true)]
    [InlineData("1.0.0", "1.0.0", false)]
    [InlineData("0.99.0", "1.0.0", false)]
    [InlineData("1.0.0", "0.99.0", true)]  // o RC 0.99.0 atualiza para a 1.0.0 (criterio 7)
    [InlineData("v1.0.1", "1.0.0", true)]  // prefixo "v" da tag
    public void EhMaisNova_Numerica(string tag, string atual, bool esperado) =>
        Assert.Equal(esperado, ServicoAtualizacao.EhMaisNova(tag, atual));

    [Fact] // licao do mPDF: 1.0.0 (Revision -1) e 1.0.0.0 (Revision 0) sao a MESMA versao
    public void EhMaisNova_IgnoraRevision()
    {
        Assert.False(ServicoAtualizacao.EhMaisNova("1.0.0", "1.0.0.0"));
        Assert.False(ServicoAtualizacao.EhMaisNova("1.0.0.0", "1.0.0"));
    }

    [Fact]
    public void EhMaisNova_TagInvalida_False() => Assert.False(ServicoAtualizacao.EhMaisNova("não-é-versão", "1.0.0"));

    [Fact] // a versao atual e' a do produto (src/Directory.Build.props via AplicativoOffice.Versao), com 3 componentes
    public void VersaoAtual_TresComponentes_IgualAoProduto()
    {
        Assert.Equal(AplicativoOffice.Versao, ServicoAtualizacao.VersaoAtualTexto());
        Assert.Matches(@"^\d+\.\d+\.\d+$", ServicoAtualizacao.VersaoAtualTexto());
        Assert.Equal(-1, ServicoAtualizacao.VersaoAtual().Revision);
    }

    [Theory]
    [InlineData("v1.0.1", "1.0.1")]
    [InlineData("1.0.1", "1.0.1")]
    public void VersaoDaTag_TiraOV(string tag, string esperado) => Assert.Equal(esperado, ServicoAtualizacao.VersaoDaTag(tag));

    // ---- hash no corpo ----------------------------------------------------------------------------------

    [Fact] // a ULTIMA linha vence (mutacao "primeira linha" tem de reprovar aqui)
    public void ExtrairSha256_UltimaLinhaVence()
    {
        var corpo = $"SHA256: {new string('a', 64)}\n\nnotas\r\nSHA256: {ReleasesDeTeste.ShaValido}";
        Assert.Equal(ReleasesDeTeste.ShaValido, ServicoAtualizacao.ExtrairSha256(corpo));
    }

    [Fact]
    public void ExtrairSha256_Maiusculas_ViramMinusculas() =>
        Assert.Equal(ReleasesDeTeste.ShaValido, ServicoAtualizacao.ExtrairSha256("SHA256: " + ReleasesDeTeste.ShaValido.ToUpperInvariant()));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("notas sem nada")]
    [InlineData("SHA256: abc123")] // comprimento errado
    public void ExtrairSha256_SemLinhaValida_Null(string? corpo) => Assert.Null(ServicoAtualizacao.ExtrairSha256(corpo));

    [Fact] // ancora de linha inteira: texto depois do hash invalida a linha
    public void ExtrairSha256_TextoDepoisNaMesmaLinha_Null() =>
        Assert.Null(ServicoAtualizacao.ExtrairSha256($"SHA256: {ReleasesDeTeste.ShaValido} (verificado)"));

    // ---- verificacao do arquivo baixado (arquivo REAL em disco) ------------------------------------------

    [Fact]
    public void VerificarEFinalizar_Adulterado_RecusaEApaga()
    {
        var p = ArquivoTemp("conteúdo adulterado");
        try
        {
            var r = ServicoAtualizacao.VerificarEFinalizar(p, ReleasesDeTeste.ShaValido);
            Assert.Equal(StatusDownload.Recusado, r.Status);
            Assert.Null(r.Arquivo);
            Assert.False(File.Exists(p));
        }
        finally { if (File.Exists(p)) File.Delete(p); }
    }

    [Fact]
    public void VerificarEFinalizar_HashConfere_ArquivoVerificado_SemDiferenciarMaiusculas()
    {
        var p = ArquivoTemp("instalador simulado");
        try
        {
            var r = ServicoAtualizacao.VerificarEFinalizar(p, Sha(p).ToUpperInvariant());
            Assert.Equal(StatusDownload.Verificado, r.Status);
            Assert.Equal(p, r.Arquivo!.Caminho);
            Assert.True(File.Exists(p));
        }
        finally { File.Delete(p); }
    }

    // ---- download: nome e host do asset (dado nao confiavel), sem rede -----------------------------------

    [Theory]
    [InlineData(@"C:\evil\x.exe")]
    [InlineData(@"\\host\x.exe")]
    [InlineData(@"..\..\x.exe")]
    [InlineData("a/b.exe")]
    public async Task Baixar_NomeMalicioso_RecusaSemTocarNoDisco(string nome)
    {
        var antes = PastasDeUpdate();
        var info = new InfoAtualizacao("v9.9.9", null, "https://github.com/x/x/releases/download/v9.9.9/x.exe", nome, 10, ReleasesDeTeste.ShaValido);
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null));
        var r = await s.BaixarEVerificarAsync(info, null, CancellationToken.None);
        Assert.Equal(StatusDownload.Recusado, r.Status);
        Assert.Equal(antes, PastasDeUpdate());
    }

    [Fact]
    public async Task Baixar_HostForaDaLista_RecusaSemTocarNoDisco()
    {
        var antes = PastasDeUpdate();
        var info = new InfoAtualizacao("v9.9.9", null, "https://evil.example.com/x.exe", "x.exe", 10, ReleasesDeTeste.ShaValido);
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null));
        Assert.Equal(StatusDownload.Recusado, (await s.BaixarEVerificarAsync(info, null, CancellationToken.None)).Status);
        Assert.Equal(antes, PastasDeUpdate());
    }

    [Fact] // token JA' cancelado: propaga o cancelamento (nunca vira "recusado") e nao deixa pasta para tras
    public async Task Baixar_TokenCancelado_Propaga_ELimpaAPasta()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var antes = PastasDeUpdate();
        var info = new InfoAtualizacao("v9.9.9", null, "https://github.com/x/x/releases/download/v9.9.9/x.exe", "x.exe", 10, ReleasesDeTeste.ShaValido);
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => s.BaixarEVerificarAsync(info, null, cts.Token));
        Assert.Equal(antes, PastasDeUpdate());
    }

    [Theory]
    [InlineData("mOffice-Setup-1.0.0.exe")]
    [InlineData("a.exe")]
    public void SanearNome_Valido_Igual(string nome) => Assert.Equal(nome, ServicoAtualizacao.SanearNomeDoAsset(nome));

    [Theory]
    [InlineData(@"C:\evil\x.exe")]
    [InlineData(@"\\host\x.exe")]
    [InlineData(@"..\..\x.exe")]
    [InlineData("a/b.exe")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void SanearNome_Malicioso_Null(string? nome) => Assert.Null(ServicoAtualizacao.SanearNomeDoAsset(nome));

    [Theory]
    [InlineData("https://github.com/maiconbiondo/moffice-desktop/releases/download/v1.0.0/x.exe")]
    [InlineData("https://objects.githubusercontent.com/a")]
    [InlineData("https://api.github.com/x")]
    [InlineData("https://sub.objects.githubusercontent.com/a")]
    public void HostPermitido_GitHub(string url) => Assert.True(ServicoAtualizacao.HostPermitido(url));

    [Theory]
    [InlineData("https://evil.example.com/x.exe")]
    [InlineData("http://github.com/x.exe")]           // sem HTTPS
    [InlineData("https://evilgithub.com/x.exe")]      // substring sem fronteira de dominio
    [InlineData("https://github.com.evil.com/x.exe")] // prefixo, nao sufixo
    [InlineData("nao e url")]
    [InlineData(null)]
    public void HostPermitido_Recusa(string? url) => Assert.False(ServicoAtualizacao.HostPermitido(url));

    // ---- contrato do repositorio e do JSON da API ----------------------------------------------------------

    [Fact]
    public void Repositorio_EAPiDeUltimaRelease()
    {
        Assert.Equal("maiconbiondo/moffice-desktop", ServicoAtualizacao.Repositorio);
        Assert.Equal("https://api.github.com/repos/maiconbiondo/moffice-desktop/releases/latest", ServicoAtualizacao.UrlReleases);
    }

    [Fact] // o PRIMEIRO asset .exe vence; o resto do JSON e' ignorado
    public void LerRelease_PrimeiroExe()
    {
        using var doc = JsonDocument.Parse("""
            {"tag_name":"v1.0.1","body":"notas\nSHA256: 058cf405f778fc15284646e9d7ad8377171681366424a448b588a17bb4a1c813",
             "assets":[{"name":"LICENCAS.txt","browser_download_url":"https://github.com/l.txt","size":5},
                       {"name":"mOffice-Setup-1.0.1.exe","browser_download_url":"https://github.com/s.exe","size":1400000000},
                       {"name":"outro.exe","browser_download_url":"https://github.com/o.exe","size":1}]}
            """);
        var r = FonteGitHub.LerRelease(doc.RootElement);
        Assert.Equal("v1.0.1", r.Tag);
        Assert.Equal("mOffice-Setup-1.0.1.exe", r.NomeAsset);
        Assert.Equal("https://github.com/s.exe", r.UrlAsset);
        Assert.Equal(1400000000, r.TamanhoAsset);
    }

    [Fact]
    public void LerRelease_SemExe_AssetNulo()
    {
        using var doc = JsonDocument.Parse("""{"tag_name":"v1.0.1","body":null,"assets":[]}""");
        var r = FonteGitHub.LerRelease(doc.RootElement);
        Assert.Null(r.NomeAsset); Assert.Null(r.UrlAsset); Assert.Null(r.Corpo);
    }

    // ---- fix round 1: limite do download, redirecionamentos e reconferencia ---------------------------------

    [Fact] // I-1: 60 min (o instalador tem ~1,4 GB); M-1: redirecionamento automatico desligado
    public void ClienteDeDownload_LimiteESemRedirecionamentoAutomatico()
    {
        Assert.Equal(TimeSpan.FromMinutes(60), ServicoAtualizacao.LimiteDoDownload);
        using var h = ServicoAtualizacao.CriarHandlerDeDownload();
        Assert.False(h.AllowAutoRedirect);
        Assert.Equal(5, ServicoAtualizacao.MaxRedirecionamentos);
    }

    [Fact] // M-1: o destino atual dos assets do GitHub
    public void HostPermitido_ReleaseAssets() =>
        Assert.True(ServicoAtualizacao.HostPermitido("https://release-assets.githubusercontent.com/github-production-release-asset/1/x"));

    [Fact] // M-1: cada salto passa pela lista (relativo resolvido contra o atual)
    public void ProximoSalto_PermitidosENegados()
    {
        var atual = new Uri("https://github.com/maiconbiondo/moffice-desktop/releases/download/v1.0.1/x.exe");
        Assert.Equal("https://release-assets.githubusercontent.com/a?b=1",
            ServicoAtualizacao.ProximoSalto(atual, new Uri("https://release-assets.githubusercontent.com/a?b=1"))!.AbsoluteUri);
        Assert.Equal("https://github.com/outro/x.exe", ServicoAtualizacao.ProximoSalto(atual, new Uri("/outro/x.exe", UriKind.Relative))!.AbsoluteUri);
        Assert.Null(ServicoAtualizacao.ProximoSalto(atual, new Uri("https://evil.example.com/x.exe")));
        Assert.Null(ServicoAtualizacao.ProximoSalto(atual, new Uri("http://release-assets.githubusercontent.com/a"))); // sem HTTPS
        Assert.Null(ServicoAtualizacao.ProximoSalto(atual, new Uri("https://release-assets.githubusercontent.com.evil.com/a")));
        Assert.Null(ServicoAtualizacao.ProximoSalto(atual, null));
    }

    /// <summary>Servidor FALSO em memoria (nenhum byte sai da maquina): responde por URL e registra o que foi pedido.</summary>
    private sealed class ServidorFalso(Func<Uri, HttpResponseMessage> responder) : HttpMessageHandler
    {
        public List<string> Pedidos { get; } = [];
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage r, CancellationToken ct)
        {
            Pedidos.Add(r.RequestUri!.AbsoluteUri);
            return Task.FromResult(responder(r.RequestUri));
        }
    }

    private static HttpResponseMessage Redireciona(string para) =>
        new(HttpStatusCode.Found) { Headers = { Location = new Uri(para, UriKind.RelativeOrAbsolute) } };

    [Fact] // M-1: github.com -> release-assets (permitido) -> 200: baixa e verifica
    public async Task Baixar_RedirecionamentoPermitido_Verifica()
    {
        var conteudo = System.Text.Encoding.UTF8.GetBytes("instalador pelo redirecionamento");
        var sha = Convert.ToHexString(SHA256.HashData(conteudo)).ToLowerInvariant();
        var srv = new ServidorFalso(u => u.Host == "github.com"
            ? Redireciona("https://release-assets.githubusercontent.com/asset/1")
            : new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(conteudo) });
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null), new HttpClient(srv));
        var info = new InfoAtualizacao("v9.9.9", null, "https://github.com/x/x/releases/download/v9.9.9/x.exe", "x.exe", conteudo.Length, sha);
        var r = await s.BaixarEVerificarAsync(info, null, CancellationToken.None);
        try
        {
            Assert.Equal(StatusDownload.Verificado, r.Status);
            Assert.Equal(2, srv.Pedidos.Count);
            Assert.Equal(conteudo, File.ReadAllBytes(r.Arquivo!.Caminho));
        }
        finally { if (r.Arquivo is { } a) { File.Delete(a.Caminho); Directory.Delete(Path.GetDirectoryName(a.Caminho)!); } }
    }

    [Fact] // M-1: redirecionamento para fora da lista: recusa SEM pedir o destino, e nao deixa pasta
    public async Task Baixar_RedirecionamentoNegado_NuncaPedeODestino()
    {
        var antes = PastasDeUpdate();
        var srv = new ServidorFalso(u => u.Host == "github.com"
            ? Redireciona("https://evil.example.com/x.exe")
            : new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) });
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null), new HttpClient(srv));
        var info = new InfoAtualizacao("v9.9.9", null, "https://github.com/x/x.exe", "x.exe", 3, ReleasesDeTeste.ShaValido);
        var r = await s.BaixarEVerificarAsync(info, null, CancellationToken.None);
        Assert.Equal(StatusDownload.Recusado, r.Status);
        Assert.DoesNotContain(srv.Pedidos, p => p.Contains("evil", StringComparison.Ordinal));
        Assert.Equal(antes, PastasDeUpdate());
    }

    [Fact] // M-1: laco de redirecionamentos: no maximo MaxRedirecionamentos saltos, depois recusa
    public async Task Baixar_SaltosDemais_Recusa()
    {
        var srv = new ServidorFalso(u => Redireciona("https://github.com/salto/" + Guid.NewGuid().ToString("N")));
        using var s = new ServicoAtualizacao(new FonteDeReleasesFalsa(null), new HttpClient(srv));
        var info = new InfoAtualizacao("v9.9.9", null, "https://github.com/x/x.exe", "x.exe", 3, ReleasesDeTeste.ShaValido);
        var r = await s.BaixarEVerificarAsync(info, null, CancellationToken.None);
        Assert.Equal(StatusDownload.Recusado, r.Status);
        Assert.Equal(ServicoAtualizacao.MaxRedirecionamentos + 1, srv.Pedidos.Count);
    }

    /// <summary>Corpo que manda uns bytes e depois PARA (servidor travado): so' o cancelamento o destrava.</summary>
    private sealed class FluxoTravado : Stream
    {
        private bool _mandou;
        public override bool CanRead => true; public override bool CanSeek => false; public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] b, int o, int c) => throw new NotSupportedException("so' assincrono");
        public override long Seek(long o, SeekOrigin s) => throw new NotSupportedException();
        public override void SetLength(long v) => throw new NotSupportedException();
        public override void Write(byte[] b, int o, int c) => throw new NotSupportedException();
        public override async ValueTask<int> ReadAsync(Memory<byte> destino, CancellationToken ct = default)
        {
            if (!_mandou) { _mandou = true; destino.Span[0] = 42; return 1; }
            await Task.Delay(Timeout.Infinite, ct);
            return 0;
        }
    }

    private static ServicoAtualizacao ServicoTravado(TimeSpan limite) =>
        new(new FonteDeReleasesFalsa(null),
            new HttpClient(new ServidorFalso(_ => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(new FluxoTravado()) }))
            { Timeout = Timeout.InfiniteTimeSpan },
            limite);

    private static InfoAtualizacao InfoGithub() =>
        new("v9.9.9", null, "https://github.com/x/x/releases/download/v9.9.9/x.exe", "x.exe", 1000, ReleasesDeTeste.ShaValido);

    [Fact] // M-2: o corpo que trava no meio NAO trava o download: o limite do download inteiro estoura, recusa e limpa
    public async Task Baixar_CorpoTravado_LimiteEstouraERecusa()
    {
        var antes = PastasDeUpdate();
        using var s = ServicoTravado(TimeSpan.FromMilliseconds(300));
        var relogio = System.Diagnostics.Stopwatch.StartNew();
        var r = await s.BaixarEVerificarAsync(InfoGithub(), null, CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(30));
        Assert.Equal(StatusDownload.Recusado, r.Status);
        Assert.Equal("Não foi possível baixar a atualização. Confira a conexão com a internet.", r.MensagemErro);
        Assert.True(relogio.Elapsed < TimeSpan.FromSeconds(10), $"demorou {relogio.Elapsed}");
        Assert.Equal(antes, PastasDeUpdate());
    }

    [Fact] // M-2: com o limite ligado, o cancelamento do USUARIO continua sendo cancelamento (nao "falhou")
    public async Task Baixar_CorpoTravado_CancelamentoDoUsuarioPropaga()
    {
        using var s = ServicoTravado(TimeSpan.FromMinutes(5));
        using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(300));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => s.BaixarEVerificarAsync(InfoGithub(), null, cts.Token).WaitAsync(TimeSpan.FromSeconds(30)));
    }

    [Fact] // M-1: enquanto a trava vive ninguem escreve, renomeia nem apaga o instalador; solta, volta ao normal
    public void ReconferirETravar_BloqueiaEscritaEApagamento()
    {
        var p = ArquivoTemp("instalador");
        try
        {
            var a = ServicoAtualizacao.VerificarEFinalizar(p, Sha(p)).Arquivo!;
            using (var trava = a.ReconferirETravar())
            {
                Assert.NotNull(trava);
                Assert.ThrowsAny<IOException>(() => File.WriteAllText(p, "trocado"));
                Assert.ThrowsAny<IOException>(() => File.Delete(p));
                Assert.ThrowsAny<IOException>(() => File.Move(p, p + ".movido"));
                Assert.Equal("instalador", File.ReadAllText(p)); // leitura continua liberada (o Windows executa lendo)
            }
            File.WriteAllText(p, "agora pode");
        }
        finally { if (File.Exists(p)) File.Delete(p); }
    }

    [Fact] // I-4: Reconferir confere o arquivo intacto e recusa (apagando) o alterado
    public void Reconferir_IntactoEAlterado()
    {
        var p = ArquivoTemp("instalador");
        try
        {
            var a = ServicoAtualizacao.VerificarEFinalizar(p, Sha(p)).Arquivo!;
            Assert.Equal(Sha(p), a.Sha256);
            Assert.True(a.Reconferir());
            Assert.True(File.Exists(p));
            File.AppendAllText(p, "!");
            Assert.False(a.Reconferir());
            Assert.False(File.Exists(p));
        }
        finally { if (File.Exists(p)) File.Delete(p); }
    }

    internal static string ArquivoTemp(string conteudo)
    {
        var p = Path.Combine(Path.GetTempPath(), $"moffice-atualizacao-teste-{Guid.NewGuid():N}.exe");
        File.WriteAllText(p, conteudo);
        return p;
    }

    internal static string Sha(string p) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(p))).ToLowerInvariant();

    private static HashSet<string> PastasDeUpdate() =>
        Directory.Exists(ServicoAtualizacao.PastaDeDownloads)
            ? Directory.GetDirectories(ServicoAtualizacao.PastaDeDownloads, "update-*").ToHashSet()
            : [];
}

// =========================================================================================================
// Verbos do RMM (porte de SilentUpdateRunnerTests)
// =========================================================================================================

public class ExecutorDeAtualizacaoTests
{
    private sealed class Lancador { public string? Caminho; public int Vezes; public void Lancar(ServicoAtualizacao.ArquivoVerificado a) { Caminho = a.Caminho; Vezes++; } }

    [Theory]
    [InlineData(new[] { "/checkupdate" }, "/checkupdate")]
    [InlineData(new[] { "/CHECKUPDATE" }, "/checkupdate")]
    [InlineData(new[] { "/update" }, "/update")]
    [InlineData(new[] { "/Update", "x" }, "/update")]
    [InlineData(new[] { "C:\\a.docx" }, null)]
    [InlineData(new[] { "C:\\a.docx", "/update" }, null)] // so' o PRIMEIRO argumento e' verbo (como no mPDF)
    [InlineData(new string[0], null)]
    public void VerboDe(string[] args, string? esperado) => Assert.Equal(esperado, ExecutorDeAtualizacao.VerboDe(args));

    [Fact]
    public async Task Verificar_JaNaUltima()
    {
        var e = new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(ReleasesDeTeste.Atual()));
        Assert.Equal(ResultadoAtualizacaoCli.JaNaUltima, (await e.VerificarAsync(CancellationToken.None)).Resultado);
    }

    [Fact] // a versao disponivel sai SEM o "v" da tag
    public async Task Verificar_Disponivel()
    {
        var s = await new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(ReleasesDeTeste.Nova())).ExecutarAsync("/checkupdate", CancellationToken.None);
        Assert.Equal(ResultadoAtualizacaoCli.AtualizacaoDisponivel, s.Resultado);
        Assert.Equal(ReleasesDeTeste.VersaoNova, s.VersaoDisponivel);
    }

    [Fact]
    public async Task Verificar_Erro()
    {
        var e = new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(null, new HttpRequestException("dns")));
        Assert.Equal(ResultadoAtualizacaoCli.Erro, (await e.VerificarAsync(CancellationToken.None)).Resultado);
    }

    [Fact]
    public async Task Atualizar_JaNaUltima_NaoLanca()
    {
        var l = new Lancador();
        var s = await new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(ReleasesDeTeste.Atual()), l.Lancar).ExecutarAsync("/update", CancellationToken.None);
        Assert.Equal(ResultadoAtualizacaoCli.JaNaUltima, s.Resultado);
        Assert.Equal(0, l.Vezes);
    }

    [Fact] // download recusado (host fora da lista, sem rede) -> erro e NADA e' lancado
    public async Task Atualizar_DownloadRecusado_Erro_NaoLanca()
    {
        var l = new Lancador();
        var s = await new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(ReleasesDeTeste.Nova()), l.Lancar).AtualizarAsync(CancellationToken.None);
        Assert.Equal(ResultadoAtualizacaoCli.Erro, s.Resultado);
        Assert.Equal(0, l.Vezes);
    }

    [Fact] // com um ArquivoVerificado REAL (obtido pelo caminho legitimo) o lancador recebe o MESMO caminho
    public void InstalarVerificado_Lanca()
    {
        var p = ServicoAtualizacaoTests.ArquivoTemp("instalador");
        try
        {
            var a = ServicoAtualizacao.VerificarEFinalizar(p, ServicoAtualizacaoTests.Sha(p)).Arquivo!;
            var l = new Lancador();
            var s = new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(null), l.Lancar).InstalarVerificado(a, "1.0.0", "1.0.1");
            Assert.Equal(ResultadoAtualizacaoCli.AtualizacaoIniciada, s.Resultado);
            Assert.Equal(1, l.Vezes);
            Assert.Equal(p, l.Caminho);
        }
        finally { File.Delete(p); }
    }

    [Fact] // I-4 no /update: arquivo trocado entre verificar e lancar -> erro, nada lancado
    public void InstalarVerificado_Adulterado_NaoLanca()
    {
        var p = ServicoAtualizacaoTests.ArquivoTemp("instalador");
        try
        {
            var a = ServicoAtualizacao.VerificarEFinalizar(p, ServicoAtualizacaoTests.Sha(p)).Arquivo!;
            File.WriteAllText(p, "trocado");
            var l = new Lancador();
            var s = new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(null), l.Lancar).InstalarVerificado(a, "1.0.0", "1.0.1");
            Assert.Equal(ResultadoAtualizacaoCli.Erro, s.Resultado);
            Assert.Equal(0, l.Vezes);
        }
        finally { if (File.Exists(p)) File.Delete(p); }
    }

    [Fact] // M-3 no /update: UAC recusado -> "Atualizacao cancelada" (ASCII na linha)
    public void InstalarVerificado_UacRecusado_Cancelada()
    {
        var p = ServicoAtualizacaoTests.ArquivoTemp("instalador");
        try
        {
            var a = ServicoAtualizacao.VerificarEFinalizar(p, ServicoAtualizacaoTests.Sha(p)).Arquivo!;
            var s = new ExecutorDeAtualizacao(() => new FonteDeReleasesFalsa(null), _ => throw new System.ComponentModel.Win32Exception(1223))
                .InstalarVerificado(a, "1.0.0", "1.0.1");
            Assert.Equal(ResultadoAtualizacaoCli.Erro, s.Resultado);
            Assert.EndsWith("erro=Atualizacao cancelada.", ExecutorDeAtualizacao.FormatarStatus("/update", s));
        }
        finally { File.Delete(p); }
    }

    [Theory] // contrato do RMM: 0 = atualizado, 10 = ha' versao nova / instalacao iniciada, 20 = erro
    [InlineData(ResultadoAtualizacaoCli.JaNaUltima, 0)]
    [InlineData(ResultadoAtualizacaoCli.AtualizacaoDisponivel, 10)]
    [InlineData(ResultadoAtualizacaoCli.AtualizacaoIniciada, 10)]
    [InlineData(ResultadoAtualizacaoCli.Erro, 20)]
    public void MapearExitCode(ResultadoAtualizacaoCli r, int esperado) => Assert.Equal(esperado, ExecutorDeAtualizacao.MapearExitCode(r));

    [Fact]
    public void FormatarStatus_Disponivel()
    {
        var linha = ExecutorDeAtualizacao.FormatarStatus("/checkupdate",
            new StatusAtualizacaoCli(ResultadoAtualizacaoCli.AtualizacaoDisponivel, "1.0.0", "1.0.1", null));
        Assert.Equal("mOffice /checkupdate: instalada=1.0.0 disponivel=1.0.1 atualizar=sim", linha);
    }

    [Fact]
    public void FormatarStatus_JaNaUltima_EIniciada()
    {
        Assert.Equal("mOffice /checkupdate: instalada=1.0.0 disponivel=1.0.0 atualizar=nao",
            ExecutorDeAtualizacao.FormatarStatus("/checkupdate", new(ResultadoAtualizacaoCli.JaNaUltima, "1.0.0", "1.0.0", null)));
        Assert.Equal("mOffice /update: instalada=1.0.0 disponivel=1.0.1 acao=instalando",
            ExecutorDeAtualizacao.FormatarStatus("/update", new(ResultadoAtualizacaoCli.AtualizacaoIniciada, "1.0.0", "1.0.1", null)));
    }

    [Fact] // a mensagem de erro e' pt-BR com acento; a linha continua ASCII
    public void FormatarStatus_Erro_Ascii()
    {
        var linha = ExecutorDeAtualizacao.FormatarStatus("/update",
            new(ResultadoAtualizacaoCli.Erro, "1.0.0", null, "Não foi possível verificar — conexão\r\nfalhou ✗"));
        Assert.All(linha, c => Assert.True(c is >= ' ' and <= '~', $"caractere fora do ASCII imprimivel: U+{(int)c:X4}"));
        Assert.StartsWith("mOffice /update: instalada=1.0.0 erro=Nao foi possivel verificar", linha);
    }
}

// =========================================================================================================
// Instalador silencioso, mutex de presenca e a atualizacao armada no encerramento
// =========================================================================================================

public class InstaladorSilenciosoTests
{
    [Fact] // RMM: invisivel; botao: com a janela de progresso (revisao final, I-1)
    public void Argumentos_RmmInvisivel_BotaoComProgresso()
    {
        Assert.Equal("/VERYSILENT /SUPPRESSMSGBOXES /NORESTART", InstaladorSilencioso.ArgumentosSilenciosos);
        Assert.Equal("/SILENT /SUPPRESSMSGBOXES /NORESTART", InstaladorSilencioso.ArgumentosComProgresso);
    }

    [Fact] // botao das Configuracoes: /SILENT (progresso), reabre o app (sem /NORELAUNCH), mesmo caminho, shell-execute (UAC)
    public void StartInfo_ComReabertura()
    {
        var psi = InstaladorSilencioso.MontarStartInfo(@"C:\Temp\mOffice-Setup-9.9.9.exe", reabrir: true, app: "mtexto");
        Assert.Equal(@"C:\Temp\mOffice-Setup-9.9.9.exe", psi.FileName);
        Assert.Equal("/SILENT /SUPPRESSMSGBOXES /NORESTART /REABRIR=mtexto", psi.Arguments);
        Assert.DoesNotContain("/VERYSILENT", psi.Arguments);
        Assert.DoesNotContain("/NORELAUNCH", psi.Arguments);
        Assert.True(psi.UseShellExecute);
    }

    [Fact] // RMM: acrescenta /NORELAUNCH
    public void StartInfo_SemReabertura()
    {
        var psi = InstaladorSilencioso.MontarStartInfo(@"C:\Temp\mOffice-Setup-9.9.9.exe", reabrir: false);
        Assert.Equal("/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /NORELAUNCH", psi.Arguments);
        Assert.DoesNotContain("/REABRIR=", psi.Arguments);
        Assert.True(psi.UseShellExecute);
    }

    [Fact] // a sobrecarga de producao (ArquivoVerificado) reabre por padrao
    public void StartInfo_ArquivoVerificado_ReabrePorPadrao()
    {
        var p = ServicoAtualizacaoTests.ArquivoTemp("x");
        try
        {
            var a = ServicoAtualizacao.VerificarEFinalizar(p, ServicoAtualizacaoTests.Sha(p)).Arquivo!;
            Assert.Equal(InstaladorSilencioso.ArgumentosComProgresso, InstaladorSilencioso.MontarStartInfo(a).Arguments);
            Assert.EndsWith(InstaladorSilencioso.ParametroNaoReabrir, InstaladorSilencioso.MontarStartInfo(a, reabrir: false).Arguments);
        }
        finally { File.Delete(p); }
    }

    [Fact] // nome fixo da suite, Global\, 8 hex; diferente dos nomes de instancia unica por app
    public void MutexDeAtualizacao_Formato()
    {
        Assert.Matches(@"^Global\\mOffice-Update-Mutex-[0-9A-F]{8}$", NomesDeInstancia.MutexDeAtualizacao);
        foreach (var app in new[] { "mtexto", "mplanilha", "mslides" })
            Assert.NotEqual(NomesDeInstancia.MutexDeAtualizacao, NomesDeInstancia.Mutex(app));
    }

    [Fact] // enquanto o handle vive, o objeto nomeado EXISTE (e' o que o AppMutex= do instalador ve')
    public void AdquirirMutex_ObjetoNomeadoPassaAExistir()
    {
        using var segurado = InstaladorSilencioso.AdquirirMutexDoApp();
        using var sonda = new Mutex(initiallyOwned: false, NomesDeInstancia.MutexDeAtualizacao, out var criadoAgora);
        Assert.False(criadoAgora);
    }

    // ---- AtualizacaoAoEncerrar ----------------------------------------------------------------------------

    private static ServicoAtualizacao.ArquivoVerificado Verificado(out string caminho)
    {
        caminho = ServicoAtualizacaoTests.ArquivoTemp("instalador");
        return ServicoAtualizacao.VerificarEFinalizar(caminho, ServicoAtualizacaoTests.Sha(caminho)).Arquivo!;
    }

    [Fact] // armada: o encerramento lanca UMA vez, com o caminho verificado
    public void Armada_EncerramentoLancaUmaVez()
    {
        var a = Verificado(out var p);
        try
        {
            var lancados = new List<string>();
            var at = new AtualizacaoAoEncerrar(x => lancados.Add(x.Caminho));
            at.Armar(a);
            Assert.True(at.Armada);
            Assert.Null(at.IniciarSeArmada());
            Assert.Null(at.IniciarSeArmada());
            Assert.Equal(new[] { p }, lancados);
            Assert.False(at.Armada);
        }
        finally { File.Delete(p); }
    }

    [Fact] // sem armar (encerramento comum do app): nada e' lancado
    public void NaoArmada_NadaLanca()
    {
        var n = 0;
        Assert.Null(new AtualizacaoAoEncerrar(_ => n++).IniciarSeArmada());
        Assert.Equal(0, n);
    }

    [Fact] // "Cancelar" num documento desarma: o encerramento seguinte nao instala
    public void Desarmada_NaoLanca()
    {
        var a = Verificado(out var p);
        try
        {
            var n = 0;
            var at = new AtualizacaoAoEncerrar(_ => n++);
            at.Armar(a);
            Assert.True(at.Desarmar());
            Assert.False(at.Desarmar());
            Assert.Null(at.IniciarSeArmada());
            Assert.Equal(0, n);
        }
        finally { File.Delete(p); }
    }

    [Fact] // o lancamento falhou: mensagem pt-BR com o caminho (o app nao encerra), o detalhe vai so' para o log
    public void FalhaAoLancar_MensagemComCaminho()
    {
        var a = Verificado(out var p);
        try
        {
            Exception? logada = null;
            var at = new AtualizacaoAoEncerrar(_ => throw new System.ComponentModel.Win32Exception(2, "The system cannot find the file specified"), ex => logada = ex);
            at.Armar(a);
            var msg = at.IniciarSeArmada();
            Assert.Equal(AtualizacaoAoEncerrar.MensagemFalhaAoIniciar(p), msg);
            Assert.DoesNotContain("cannot find", msg);
            Assert.NotNull(logada);
        }
        finally { File.Delete(p); }
    }

    // ---- fix round 1 ----------------------------------------------------------------------------------

    [Fact] // M-3: UAC recusado (ERROR_CANCELLED 1223) -> "Atualização cancelada.", nao o erro generico
    public void UacRecusado_AtualizacaoCancelada()
    {
        var a = Verificado(out var p);
        try
        {
            var at = new AtualizacaoAoEncerrar(_ => throw new System.ComponentModel.Win32Exception(1223, "The operation was canceled by the user"));
            at.Armar(a);
            Assert.Equal("Atualização cancelada.", at.IniciarSeArmada());
        }
        finally { File.Delete(p); }
    }

    [Fact] // I-4: o instalador TROCADO no disco depois da verificacao e' recusado e apagado; nada e' lancado
    public void AdulteradoDepoisDeArmar_RecusaApagaENaoLanca()
    {
        var a = Verificado(out var p);
        try
        {
            var lancados = 0; var soltou = 0;
            var at = new AtualizacaoAoEncerrar(_ => lancados++, soltarPresenca: () => soltou++);
            at.Armar(a);
            File.WriteAllText(p, "instalador trocado por outro processo");
            Assert.Equal(ServicoAtualizacao.MensagemAlteradoDepois, at.IniciarSeArmada());
            Assert.Equal(0, lancados);
            Assert.Equal(0, soltou);          // a presenca nem chega a ser solta
            Assert.False(File.Exists(p));
            Assert.False(at.Armada);
        }
        finally { if (File.Exists(p)) File.Delete(p); }
    }

    [Fact] // I-4: arquivo apagado depois de armar: tambem recusa (nunca lanca um caminho que nao confere)
    public void ApagadoDepoisDeArmar_Recusa()
    {
        var a = Verificado(out var p);
        File.Delete(p);
        var lancados = 0;
        var at = new AtualizacaoAoEncerrar(_ => lancados++);
        at.Armar(a);
        Assert.Equal(ServicoAtualizacao.MensagemAlteradoDepois, at.IniciarSeArmada());
        Assert.Equal(0, lancados);
    }

    [Fact] // I-5: a presenca (mutex do AppMutex=) sai ANTES do lancamento; sucesso nao a retoma
    public void Presenca_SoltaAntesDeLancar_SemRetomarNoSucesso()
    {
        var a = Verificado(out var p);
        try
        {
            var ordem = new List<string>();
            var at = new AtualizacaoAoEncerrar(_ => ordem.Add("lancar"), soltarPresenca: () => ordem.Add("soltar"), retomarPresenca: () => ordem.Add("retomar"));
            at.Armar(a);
            Assert.Null(at.IniciarSeArmada());
            Assert.Equal(new[] { "soltar", "lancar" }, ordem);
        }
        finally { File.Delete(p); }
    }

    [Fact] // M-1: DURANTE o lancamento (o Process.Start, a pergunta do UAC) o instalador nao pode ser trocado nem apagado
    public void DuranteOLancamento_ArquivoTravado()
    {
        var a = Verificado(out var p);
        try
        {
            var provou = false;
            var at = new AtualizacaoAoEncerrar(x =>
            {
                Assert.ThrowsAny<IOException>(() => File.WriteAllText(x.Caminho, "trocado durante o UAC"));
                Assert.ThrowsAny<IOException>(() => File.Delete(x.Caminho));
                provou = true;
            });
            at.Armar(a);
            Assert.Null(at.IniciarSeArmada());
            Assert.True(provou);
            File.Delete(p); // depois que o lancamento voltou, a trava foi solta
            Assert.False(File.Exists(p));
        }
        finally { if (File.Exists(p)) File.Delete(p); }
    }

    [Fact] // M-3: outro app da suite aberto no ENCERRAMENTO: desarma, nao lanca e diz qual fechar
    public void OutroAppAbertoNoEncerramento_DesarmaENaoLanca()
    {
        var a = Verificado(out var p);
        try
        {
            var lancados = 0; var soltou = 0;
            var at = new AtualizacaoAoEncerrar(_ => lancados++, soltarPresenca: () => soltou++, outrosAppsAbertos: () => ["mPlanilha"]);
            at.Armar(a);
            Assert.Equal(OutrosAppsDaSuite.MensagemFecharOutros(["mPlanilha"]), at.IniciarSeArmada());
            Assert.Equal(0, lancados); Assert.Equal(0, soltou);
            Assert.False(at.Armada);
            Assert.True(File.Exists(p)); // o instalador verificado fica onde esta'
        }
        finally { File.Delete(p); }
    }

    [Fact] // I-5: o lancamento falhou (o app continua aberto): a presenca e' retomada
    public void Presenca_RetomadaSeOLancamentoFalhar()
    {
        var a = Verificado(out var p);
        try
        {
            var ordem = new List<string>();
            var at = new AtualizacaoAoEncerrar(_ => { ordem.Add("lancar"); throw new InvalidOperationException("x"); },
                soltarPresenca: () => ordem.Add("soltar"), retomarPresenca: () => ordem.Add("retomar"));
            at.Armar(a);
            Assert.NotNull(at.IniciarSeArmada());
            Assert.Equal(new[] { "soltar", "lancar", "retomar" }, ordem);
        }
        finally { File.Delete(p); }
    }

    [Fact] // I-2 (lado da casca): o fechamento abandonado desarma -> o encerramento seguinte nao instala nada
    public void FechamentoAbandonado_Desarma_EncerramentoNaoLanca()
    {
        var a = Verificado(out var p);
        try
        {
            var lancados = 0;
            var at = new AtualizacaoAoEncerrar(_ => lancados++);
            at.Armar(a);
            // "Salvar" -> "Salvar como" cancelado: a JanelaEditor dispara FechamentoAbortado, que a casca liga ao desarme
            // (fiacao conferida em AtualizacaoRoteamentoTests; o disparo, em FecharComAlteracoesTests no nucleo).
            Action fechamentoAbortado = () => at.Desarmar();
            fechamentoAbortado();
            // Mais tarde o usuario fecha a ultima janela por conta propria: o app encerra SEM instalar.
            Assert.Null(at.IniciarSeArmada());
            Assert.Equal(0, lancados);
        }
        finally { File.Delete(p); }
    }

    // ---- I-3: outros apps da suite abertos ------------------------------------------------------------

    [Fact] // detecta pelo mutex de instancia unica de cada app (nomes de TESTE, nunca os reais), fora o proprio app
    public void OutrosApps_DetectaPeloMutexDeInstancia()
    {
        var sufixo = Guid.NewGuid().ToString("N");
        string Nome(PerfilApp perfil) => $"mOffice-teste-{perfil.Nome}-{sufixo}";
        Assert.Empty(OutrosAppsDaSuite.Abertos(PerfilApp.Texto, Nome));
        using (new Mutex(false, Nome(PerfilApp.Slides)))
        using (new Mutex(false, Nome(PerfilApp.Texto))) // o PROPRIO app nunca conta
        {
            Assert.Equal(new[] { "mSlides" }, OutrosAppsDaSuite.Abertos(PerfilApp.Texto, Nome));
            Assert.Equal(new[] { "mTexto", "mSlides" }, OutrosAppsDaSuite.Abertos(PerfilApp.Planilha, Nome));
        }
        Assert.Empty(OutrosAppsDaSuite.Abertos(PerfilApp.Texto, Nome));
    }

    [Fact]
    public void OutrosApps_Mensagem()
    {
        Assert.Equal("Feche o mPlanilha antes de atualizar (documentos abertos nele não seriam salvos).",
            OutrosAppsDaSuite.MensagemFecharOutros(["mPlanilha"]));
        Assert.Equal("Feche o mPlanilha e o mSlides antes de atualizar (documentos abertos neles não seriam salvos).",
            OutrosAppsDaSuite.MensagemFecharOutros(["mPlanilha", "mSlides"]));
    }
}

// =========================================================================================================
// Botao "Verificar atualizações" (VM do dialogo de Configuracoes), sem rede
// =========================================================================================================

[Collection("janela real")]
public class ConfiguracoesAtualizacaoTests
{
    private static Configuracoes Cfg() => new(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "config.json"));

    private static DialogoConfiguracoesViewModel Vm(FonteDeReleasesFalsa fonte, Func<string, bool> confirmar, List<string>? perguntas = null,
        IReadOnlyList<string>? outros = null) =>
        new(Cfg(), new TemaService(new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() }),
            () => fonte, p => { perguntas?.Add(p); return confirmar(p); }, () => outros ?? [], "mTexto");

    [Fact] // I-3: outro app da suite aberto -> nem pergunta, nem baixa; diz qual fechar
    public void Disponivel_OutroAppAberto_NaoOfereceInstalar()
    {
        JanelaSta.Executar(async () =>
        {
            var perguntas = new List<string>();
            var vm = Vm(new FonteDeReleasesFalsa(ReleasesDeTeste.Nova()), _ => true, perguntas, ["mPlanilha"]);
            await vm.VerificarAtualizacoesCommand.ExecuteAsync(null);
            Assert.Empty(perguntas);
            Assert.Equal(EstadoAtualizacao.Disponivel, vm.EstadoAtualizacao);
            Assert.Contains("Feche o mPlanilha antes de atualizar (documentos abertos nele não seriam salvos).", vm.MensagemAtualizacao);
            Assert.Null(vm.InstaladorPronto);
        });
    }

    [Fact] // a pergunta so' promete o que acontece: fecha as janelas deste app, UAC, progresso visivel (/SILENT),
           // reabre ESTE app (/REABRIR=) — e pede para nao abrir o mOffice no meio (revisao final, I-1)
    public void PerguntaInstalar_SoPrometeOQueAcontece()
    {
        var q = DialogoConfiguracoesViewModel.PerguntaInstalar("1.0.1", "mTexto");
        Assert.StartsWith("Versão 1.0.1 disponível. Instalar agora?", q);
        Assert.Contains("As janelas do mTexto serão fechadas", q);
        Assert.Contains("O instalador mostra o progresso e o mTexto reabre sozinho em um ou dois minutos. Não abra o mOffice enquanto isso.", q);
    }

    [Fact] // abrir o dialogo NAO consulta nada (rede so' por clique)
    public void Construir_NaoConsultaAFonte()
    {
        JanelaSta.Executar(() =>
        {
            var fonte = new FonteDeReleasesFalsa(ReleasesDeTeste.Nova());
            var vm = Vm(fonte, _ => throw new InvalidOperationException("nao deveria perguntar"));
            Assert.Equal(0, fonte.Chamadas);
            Assert.Equal(EstadoAtualizacao.Ocioso, vm.EstadoAtualizacao);
            Assert.Null(vm.MensagemAtualizacao);
            return Task.CompletedTask;
        });
    }

    [Fact]
    public void JaNaUltima_MensagemComAVersao()
    {
        JanelaSta.Executar(async () =>
        {
            var fonte = new FonteDeReleasesFalsa(ReleasesDeTeste.Atual());
            var vm = Vm(fonte, _ => throw new InvalidOperationException("nao deveria perguntar"));
            await vm.VerificarAtualizacoesCommand.ExecuteAsync(null);
            Assert.Equal(1, fonte.Chamadas);
            Assert.Equal(EstadoAtualizacao.Atualizado, vm.EstadoAtualizacao);
            Assert.Equal($"Você está com a versão mais recente ({AplicativoOffice.Versao}).", vm.MensagemAtualizacao);
        });
    }

    [Fact] // disponivel + "Nao": nada baixa, nada fecha
    public void Disponivel_Nao_NadaAcontece()
    {
        JanelaSta.Executar(async () =>
        {
            var perguntas = new List<string>();
            var vm = Vm(new FonteDeReleasesFalsa(ReleasesDeTeste.Nova()), _ => false, perguntas);
            var fechou = false; vm.PedirFechar += () => fechou = true;
            await vm.VerificarAtualizacoesCommand.ExecuteAsync(null);
            Assert.StartsWith($"Versão {ReleasesDeTeste.VersaoNova} disponível. Instalar agora?", Assert.Single(perguntas));
            Assert.Equal(EstadoAtualizacao.Disponivel, vm.EstadoAtualizacao);
            Assert.Equal($"Versão {ReleasesDeTeste.VersaoNova} disponível.", vm.MensagemAtualizacao);
            Assert.Null(vm.InstaladorPronto);
            Assert.False(fechou);
        });
    }

    [Fact] // disponivel + "Sim", mas o download e' recusado (host fora da lista, sem rede): erro sem detalhe, nada fecha
    public void Disponivel_Sim_DownloadRecusado_ErroENaoFecha()
    {
        JanelaSta.Executar(async () =>
        {
            var vm = Vm(new FonteDeReleasesFalsa(ReleasesDeTeste.Nova()), _ => true);
            var fechou = false; vm.PedirFechar += () => fechou = true;
            await vm.VerificarAtualizacoesCommand.ExecuteAsync(null);
            Assert.Equal(EstadoAtualizacao.Erro, vm.EstadoAtualizacao);
            Assert.Equal("A origem do instalador não é confiável.", vm.MensagemAtualizacao);
            Assert.Null(vm.InstaladorPronto);
            Assert.False(fechou);
        });
    }

    [Fact] // erro de rede: a mensagem e' a pt-BR do servico, sem o texto da excecao
    public void ErroDeRede_MensagemSemDetalhe()
    {
        JanelaSta.Executar(async () =>
        {
            var vm = Vm(new FonteDeReleasesFalsa(null, new HttpRequestException("SocketException 10060 detalhe")), _ => true);
            await vm.VerificarAtualizacoesCommand.ExecuteAsync(null);
            Assert.Equal(EstadoAtualizacao.Erro, vm.EstadoAtualizacao);
            Assert.DoesNotContain("10060", vm.MensagemAtualizacao);
            Assert.Contains("internet", vm.MensagemAtualizacao);
        });
    }

    [Fact] // o dialogo real mostra o botao "Verificar atualizações" ligado ao comando
    public void Dialogo_TemOBotaoVerificarAtualizacoes()
    {
        JanelaSta.Executar(() =>
        {
            var vm = Vm(new FonteDeReleasesFalsa(null), _ => false);
            var d = new DialogoConfiguracoes(vm); JanelaWpf.Aplicar(d, escuro: false);
            try
            {
                d.Show(); d.UpdateLayout();
                var botao = Descendentes<Button>(d).Single(b => Equals(b.Content, "Verificar atualizações"));
                Assert.Same(vm.VerificarAtualizacoesCommand, botao.Command);
            }
            finally { d.Close(); }
            return Task.CompletedTask;
        });
    }

    private static IEnumerable<T> Descendentes<T>(DependencyObject raiz) where T : DependencyObject
    {
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(raiz); i++)
        {
            var filho = VisualTreeHelper.GetChild(raiz, i);
            if (filho is T t) yield return t;
            foreach (var d in Descendentes<T>(filho)) yield return d;
        }
    }
}

// =========================================================================================================
// Roteamento dos verbos e fiacao da subida (varredura de fonte: um Application nao sobe num teste)
// =========================================================================================================

public class AtualizacaoRoteamentoTests
{
    private static string Ler(params string[] partes) =>
        File.ReadAllText(Path.Combine(new[] { AplicativoOfficeTests.RaizRepo() }.Concat(partes).ToArray()));

    private static string Subida()
    {
        var f = Ler("src", "mOffice.Comum", "AplicativoOffice.cs");
        var i = f.IndexOf("protected override void OnStartup(", StringComparison.Ordinal);
        Assert.True(i >= 0);
        return f[i..f.IndexOf("protected override void OnExit(", i, StringComparison.Ordinal)];
    }

    [Theory] // os TRES executaveis sobem pelo OnStartup do AplicativoOffice: nenhum o substitui
    [InlineData("mTexto")]
    [InlineData("mPlanilha")]
    [InlineData("mSlides")]
    public void TresApps_HerdamASubidaComum(string app)
    {
        var cs = Ler("src", app, "App.xaml.cs");
        Assert.Contains(": mOffice.Comum.AplicativoOffice", cs);
        Assert.DoesNotContain("OnStartup", cs);
        Assert.DoesNotContain("Startup=", Ler("src", app, "App.xaml")); // nem um Startup= no XAML contornando
    }

    [Fact] // os verbos sao roteados ANTES da instancia unica e de qualquer outra coisa da subida
    public void Verbos_AntesDaInstanciaUnica()
    {
        var s = Subida();
        var iVerbo = s.IndexOf("ExecutorDeAtualizacao.VerboDe(e.Args)", StringComparison.Ordinal);
        var iInstancia = s.IndexOf("new InstanciaUnica(", StringComparison.Ordinal);
        var iPastas = s.IndexOf("Pastas.Garantir(", StringComparison.Ordinal);
        Assert.True(iVerbo >= 0 && iInstancia > iVerbo && iPastas > iVerbo, "os verbos do RMM tem de vir antes de tudo na subida");
        var regiao = s[iVerbo..iInstancia];
        // Task.Run NA regiao do roteamento (evita o deadlock do DispatcherSynchronizationContext).
        Assert.Contains("Task.Run(() => executor.ExecutarAsync(verbo, CancellationToken.None))", regiao);
        Assert.Contains("SaidaDeStatus.Escrever(linha);", regiao);
        Assert.Contains("Shutdown(ExecutorDeAtualizacao.MapearExitCode(status.Resultado));", regiao);
        Assert.Contains("return;", regiao[regiao.IndexOf("Shutdown(ExecutorDeAtualizacao", StringComparison.Ordinal)..]);
    }

    [Fact] // M-1 (revisao final): a subida do app e o /update varrem os downloads abandonados
    public void Varredura_NaSubidaENoUpdate()
    {
        var s = Subida();
        var iVarre = s.IndexOf("ServicoAtualizacao.VarrerDownloadsAntigos();", StringComparison.Ordinal);
        Assert.True(iVarre > s.IndexOf("ExecutorDeAtualizacao.VerboDe(e.Args)", StringComparison.Ordinal));
        var exec = Ler("src", "mOffice.Comum", "Atualizacao", "ExecutorDeAtualizacao.cs");
        var atualizar = exec[exec.IndexOf("public async Task<StatusAtualizacaoCli> AtualizarAsync(", StringComparison.Ordinal)..];
        Assert.True(atualizar.IndexOf("ServicoAtualizacao.VarrerDownloadsAntigos();", StringComparison.Ordinal)
            < atualizar.IndexOf("CriarServico()", StringComparison.Ordinal));
    }

    [Fact] // mutex de presenca: adquirido na subida (melhor esforco) e solto no OnExit
    public void MutexDePresenca_AdquiridoESolto()
    {
        var f = Ler("src", "mOffice.Comum", "AplicativoOffice.cs");
        Assert.Contains("AdquirirPresenca();", Subida());
        var adquirir = f[f.IndexOf("private void AdquirirPresenca()", StringComparison.Ordinal)..];
        Assert.Contains("try { _mutexDeAtualizacao = InstaladorSilencioso.AdquirirMutexDoApp(); }", adquirir[..adquirir.IndexOf("\n    }", StringComparison.Ordinal)]);
        var saida = f[f.IndexOf("protected override void OnExit(", StringComparison.Ordinal)..];
        Assert.Contains("_mutexDeAtualizacao?.Dispose();", saida[..saida.IndexOf('}')]);
    }

    [Fact] // o botao reabre o app; so' o RMM passa reabrir:false
    public void BotaoReabre_RmmNao()
    {
        var app = Ler("src", "mOffice.Comum", "AplicativoOffice.cs");
        // Task 3: com /REABRIR=<este app> (reabrir fica no padrao, true).
        Assert.Contains("InstaladorSilencioso.MontarStartInfo(a, app: _perfil.Nome))", app);
        Assert.DoesNotContain("reabrir: false", app);
        Assert.Contains("InstaladorSilencioso.MontarStartInfo(a, reabrir: false)", Ler("src", "mOffice.Comum", "Atualizacao", "ExecutorDeAtualizacao.cs"));
    }

    [Fact] // "instalar agora": fecha pelas janelas (fluxo normal) e o instalador sobe no encerramento
    public void InstalarAgora_FechaPeloFluxoNormal()
    {
        var app = Ler("src", "mOffice.Comum", "AplicativoOffice.cs");
        Assert.Contains("if (vm.InstaladorPronto is { } arquivo) Dispatcher.BeginInvoke(() => IniciarAtualizacao(arquivo));", app);
        Assert.Contains("_janelas.FecharTodas();", app);
        Assert.Contains("PerguntarAoFechar = DialogoFechar.PerguntarAsync,", app);
        // Plano 6, Task 5: as perguntas da trava com os dialogos da casca, e o "Cancelar" do arquivo em uso volta a' boas-vindas sem erro
        Assert.Contains("PerguntarEmUso = DialogoConflito.PerguntarEmUsoAsync,", app);
        Assert.Contains("PerguntarConflito = DialogoConflito.PerguntarAsync,", app);
        Assert.Contains("ex is JanelaEditor.ConversaoDePdfRecusadaException or JanelaEditor.AberturaEmUsoCanceladaException", app);
        // I-2: todo fechamento abandonado desarma, e voltar a trabalhar tambem
        Assert.Contains("j.FechamentoAbortado += () => DesarmarAtualizacao(", app);
        foreach (var metodo in new[] { "private async Task AbrirSeguroAsync(", "private async Task NovoSeguroAsync(",
                                       "private async Task RecuperarSeguroAsync(", "private void AoPedirAtivacao(" })
        {
            var corpo = app[app.IndexOf(metodo, StringComparison.Ordinal)..];
            corpo = corpo[..corpo.IndexOf("\n    }", StringComparison.Ordinal)];
            Assert.Contains("DesarmarAtualizacao(", corpo);
        }
        // I-3: o dialogo de producao recebe a deteccao dos outros apps, e a casca confere de novo antes de armar
        Assert.Contains("outrosAppsAbertos: () => OutrosAppsDaSuite.Abertos(_perfil)", app);
        var iniciar = app[app.IndexOf("private void IniciarAtualizacao(", StringComparison.Ordinal)..];
        Assert.True(iniciar.IndexOf("OutrosAppsDaSuite.Abertos(_perfil)", StringComparison.Ordinal)
            < iniciar.IndexOf("_atualizacao.Armar(arquivo);", StringComparison.Ordinal));
        // I-5: a presenca e' solta antes do lancamento e retomada na falha
        Assert.Contains("soltarPresenca: () => { _mutexDeAtualizacao?.Dispose(); _mutexDeAtualizacao = null; },", app);
        Assert.Contains("retomarPresenca: AdquirirPresenca", app);
        // M-3 (fix round 2): o encerramento confere de novo os outros apps
        Assert.Contains("outrosAppsAbertos: () => OutrosAppsDaSuite.Abertos(_perfil));", app);
        var encerrar = app[app.IndexOf("private void EncerrarAgora()", StringComparison.Ordinal)..];
        encerrar = encerrar[..encerrar.IndexOf("Shutdown();", StringComparison.Ordinal)];
        Assert.Contains("_atualizacao.IniciarSeArmada()", encerrar);
    }
}
