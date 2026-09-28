using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace mOffice.Comum.Atualizacao;

/// <summary>
/// Atualizacao sob demanda pelo GitHub Releases (Plano 5, Task 2). Porte de <c>UpdateService.cs</c> do mPDF
/// (o exemplar vence na mecanica; aqui so' mudam nomes e textos).
///
/// TODA a rede do produto vive NESTE ARQUIVO — e so' nele. O mOffice tem egresso zero no uso normal: a
/// unica conversa com a internet e' o clique em "Verificar atualizações" (Configurações) ou os verbos
/// <c>/checkupdate</c> e <c>/update</c> do RMM. Duas guardas estruturais vigiam isso:
/// <c>RedeTests</c> (mOffice.Nucleo.Tests: tipos de rede em <c>src/</c>, com exceção por NOME EXATO deste
/// arquivo) e <c>RedeConfinadaTests</c> (mOffice.Comum.Tests: o mesmo por outra varredura, mais "o
/// <see cref="ServicoAtualizacao"/> so' e' construido nos dois sitios deliberados"). Por isso a fonte REAL
/// (<see cref="FonteGitHub"/>) mora aqui e nao num arquivo proprio: a fronteira de rede e' por ARQUIVO.
/// </summary>
public sealed class ServicoAtualizacao : IDisposable
{
    /// <summary>Repositorio PUBLICO das releases do mOffice (a copia publica do codigo, Task 4 do Plano 5).</summary>
    public const string Repositorio = "maiconbiondo/moffice-desktop";

    /// <summary>A API de "ultima release" (o GitHub ignora pre-releases e rascunhos aqui).</summary>
    public const string UrlReleases = "https://api.github.com/repos/" + Repositorio + "/releases/latest";

    private readonly IFonteDeReleases _fonte;
    private readonly HttpClient _download;

    public ServicoAtualizacao(IFonteDeReleases fonte) : this(fonte, CriarClienteDeDownload()) { }

    /// <param name="limite">Limite do download INTEIRO (costura de teste; padrao <see cref="LimiteDoDownload"/>).</param>
    internal ServicoAtualizacao(IFonteDeReleases fonte, HttpClient download, TimeSpan? limite = null)
    {
        _fonte = fonte;
        _download = download;
        _limite = limite ?? LimiteDoDownload;
    }

    private readonly TimeSpan _limite;

    /// <summary>Limite do download inteiro. O instalador tem ~1,4 GB (motor + runtime): os 10 min do mPDF exigiriam
    /// ~19 Mbit/s. O cancelamento pelo usuario (fechar o dialogo) continua valendo pelo token.</summary>
    internal static readonly TimeSpan LimiteDoDownload = TimeSpan.FromMinutes(60);

    /// <summary>Maximo de redirecionamentos seguidos no download (o GitHub usa 1: github.com -> release-assets).</summary>
    internal const int MaxRedirecionamentos = 5;

    /// <summary>Redirecionamento AUTOMATICO desligado: cada <c>Location</c> passa por <see cref="HostPermitido"/> antes
    /// de ser seguido (ver <see cref="BaixarEVerificarAsync"/>). Com ele ligado, so' o primeiro host seria conferido.</summary>
    internal static SocketsHttpHandler CriarHandlerDeDownload() => new() { AllowAutoRedirect = false };

    private static HttpClient CriarClienteDeDownload() => new(CriarHandlerDeDownload()) { Timeout = LimiteDoDownload };

    /// <summary>A fonte de producao (a API do GitHub). Os testes nunca a usam: passam um falso.</summary>
    public static IFonteDeReleases CriarFontePadrao() => new FonteGitHub();

    public void Dispose()
    {
        _download.Dispose();
        (_fonte as IDisposable)?.Dispose(); // a FonteGitHub tem o HttpClient dela (achado do mPDF: vazava)
    }

    // ---- versao instalada -----------------------------------------------------------------------------

    /// <summary>Versao instalada, SEMPRE com 3 componentes. Sai de <see cref="AplicativoOffice.Versao"/>
    /// (a <c>&lt;Version&gt;</c> unica de <c>src/Directory.Build.props</c>).</summary>
    public static Version VersaoAtual() => Normalizar(Version.Parse(AplicativoOffice.Versao));

    public static string VersaoAtualTexto() => VersaoAtual().ToString();

    // ---- verificar ------------------------------------------------------------------------------------

    /// <summary>Consulta a fonte e devolve um resultado tipado. "Disponivel" so' com asset .exe E hash no
    /// corpo. Mensagens de erro em pt-BR, sem detalhe tecnico (o detalhe nao chega ao usuario).</summary>
    public async Task<ResultadoVerificacao> VerificarAsync(CancellationToken ct = default)
    {
        ReleaseMaisRecente? release;
        try
        {
            release = await _fonte.ObterUltimaAsync(ct);
        }
        // 403 (limite de requisicoes sem autenticacao) e 429 (Too Many Requests): e' limite de taxa, nao falta de internet.
        catch (HttpRequestException ex) when (ex.StatusCode is HttpStatusCode.Forbidden or HttpStatusCode.TooManyRequests)
        {
            return ResultadoVerificacao.Erro(TipoErroAtualizacao.LimiteTaxa, MensagemLimiteTaxa);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException
            or System.Net.Sockets.SocketException or JsonException or IOException)
        {
            return ResultadoVerificacao.Erro(TipoErroAtualizacao.SemRede, MensagemSemRede);
        }

        if (release is null || string.IsNullOrWhiteSpace(release.Tag))
            return ResultadoVerificacao.Erro(TipoErroAtualizacao.SemRede, MensagemSemRede);

        if (!EhMaisNova(release.Tag, VersaoAtualTexto()))
            return ResultadoVerificacao.Atualizado();

        if (string.IsNullOrWhiteSpace(release.UrlAsset))
            return ResultadoVerificacao.Erro(TipoErroAtualizacao.SemAsset, MensagemSemAsset);

        var sha = ExtrairSha256(release.Corpo);
        if (sha is null)
            return ResultadoVerificacao.Erro(TipoErroAtualizacao.SemHash, MensagemSemHash);

        return ResultadoVerificacao.Disponivel(new InfoAtualizacao(
            release.Tag, release.Corpo, release.UrlAsset, release.NomeAsset ?? "instalador.exe", release.TamanhoAsset, sha));
    }

    internal const string MensagemLimiteTaxa = "Muitas verificações recentes. Tente novamente em alguns minutos.";
    internal const string MensagemSemRede = "Não foi possível verificar se há atualização. Confira a conexão com a internet.";
    internal const string MensagemSemAsset = "A versão publicada não tem um instalador disponível.";
    /// <summary>O instalador mudou no disco depois de verificado (<see cref="ArquivoVerificado.Reconferir"/>).</summary>
    public const string MensagemAlteradoDepois =
        "O instalador baixado foi alterado depois da verificação e foi removido. Verifique as atualizações de novo.";

    internal const string MensagemSemHash = "A versão publicada não informa o código de verificação do instalador.";

    // ---- comparacao de versao (System.Version, nunca string: 1.10 > 1.9) --------------------------------

    /// <summary>Tag e versao atual aceitam o prefixo "v". Qualquer uma que nao parseie da' <c>false</c>
    /// (nunca oferece atualizacao a partir de dado invalido).</summary>
    internal static bool EhMaisNova(string tag, string atual)
    {
        var t = Parsear(tag);
        var a = Parsear(atual);
        if (t is null || a is null) return false;
        return Normalizar(t) > Normalizar(a);
    }

    /// <summary>A tag sem o "v" (o que o usuario le': "Versão 1.0.1 disponível").</summary>
    public static string VersaoDaTag(string tag) =>
        tag.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? tag[1..] : tag;

    private static Version? Parsear(string texto) => Version.TryParse(VersaoDaTag(texto.Trim()), out var v) ? v : null;

    /// <summary>Sempre Major.Minor.Build. Licao do mPDF: "1.0.0" do assembly e' 1.0.0.0 (Revision 0) e a tag
    /// "1.0.0" e' Revision -1 — comparar cru faria versoes IGUAIS parecerem diferentes.</summary>
    private static Version Normalizar(Version v) => new(v.Major, v.Minor, Math.Max(v.Build, 0));

    // ---- hash no corpo da release (contrato do release.ps1: a ULTIMA linha "SHA256: <64 hex>") ----------

    private static readonly Regex LinhaSha = new(@"^SHA256:\s*([0-9a-fA-F]{64})$", RegexOptions.Compiled);

    /// <summary>A ULTIMA linha do corpo que casa vence (notas livres que citem a string antes nao enganam).
    /// <c>null</c> sem nenhuma linha valida (release sem hash).</summary>
    internal static string? ExtrairSha256(string? corpo)
    {
        if (string.IsNullOrEmpty(corpo)) return null;
        string? achado = null;
        foreach (var linha in corpo.Split('\n'))
        {
            var m = LinhaSha.Match(linha.TrimEnd('\r'));
            if (m.Success) achado = m.Groups[1].Value.ToLowerInvariant();
        }
        return achado;
    }

    // ---- baixar + verificar -----------------------------------------------------------------------------

    /// <summary>Pasta-raiz dos downloads: <c>%TEMP%\mOffice</c> (cada chamada cria <c>update-&lt;guid&gt;</c>).</summary>
    internal static string PastaDeDownloads => Path.Combine(Path.GetTempPath(), "mOffice");

    /// <summary>Idade minima de uma pasta <c>update-*</c> para a varredura apaga-la: 1 dia (um download em curso, ou um
    /// instalador armado esperando as janelas fecharem, e' sempre mais novo que isso).</summary>
    internal static readonly TimeSpan IdadeMinimaDownloads = TimeSpan.FromDays(1);

    /// <summary>Varre os downloads antigos da pasta padrao (subida do app e <c>/update</c>). Nunca lanca.</summary>
    public static int VarrerDownloadsAntigos() => VarrerDownloadsAntigos(PastaDeDownloads, IdadeMinimaDownloads);

    /// <summary>
    /// Apaga as pastas <c>update-*</c> de <paramref name="raiz"/> mais velhas que <paramref name="idadeMinima"/> (cada
    /// instalador tem ~1,4 GB: um download abandonado nao pode ficar no %TEMP% para sempre). NUNCA segue link: pasta
    /// que e' ponto de reparse (junction/symlink) e' pulada, e uma pasta com qualquer coisa alem de arquivos comuns
    /// (subpasta, link) tambem — so' se apaga o que esta pasta mesma criou. <paramref name="emUso"/>: uma pasta que
    /// nunca e' tocada (o instalador armado). Devolve quantas apagou. Melhor esforco: nunca lanca.
    /// </summary>
    internal static int VarrerDownloadsAntigos(string raiz, TimeSpan idadeMinima, string? emUso = null)
    {
        var apagadas = 0;
        try
        {
            if (!Directory.Exists(raiz)) return 0;
            var limite = DateTime.UtcNow - idadeMinima;
            foreach (var pasta in Directory.EnumerateDirectories(raiz, "update-*"))
            {
                try
                {
                    var info = new DirectoryInfo(pasta);
                    if (info.Attributes.HasFlag(FileAttributes.ReparsePoint)) continue;
                    if (emUso is not null && string.Equals(Path.GetFullPath(pasta).TrimEnd('\\'),
                            Path.GetFullPath(emUso).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase)) continue;
                    if (info.LastWriteTimeUtc > limite || info.CreationTimeUtc > limite) continue;
                    var itens = info.GetFileSystemInfos();
                    if (itens.Any(i => i is DirectoryInfo || i.Attributes.HasFlag(FileAttributes.ReparsePoint))) continue;
                    foreach (var f in itens) { f.Attributes = FileAttributes.Normal; f.Delete(); }
                    info.Delete(recursive: false);
                    apagadas++;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* em uso ou sem permissao: fica */ }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return apagadas;
    }

    /// <summary>
    /// Baixa o instalador para <c>%TEMP%\mOffice\update-&lt;guid&gt;\&lt;nome&gt;</c> e verifica o SHA-256. Nome e host
    /// do asset vem da API (dado NAO confiavel): os dois sao validados ANTES de qualquer I/O. Cancelamento genuino
    /// propaga; qualquer outra falha vira um <see cref="ResultadoDownload"/> recusado, e o arquivo parcial e' apagado.
    /// </summary>
    public async Task<ResultadoDownload> BaixarEVerificarAsync(
        InfoAtualizacao info, IProgress<(long Baixados, long Total)>? progresso, CancellationToken ct)
    {
        var nome = SanearNomeDoAsset(info.NomeAsset);
        if (nome is null) return ResultadoDownload.Falhou("O instalador publicado tem um nome inválido.");
        if (!HostPermitido(info.UrlAsset))
            return ResultadoDownload.Falhou("A origem do instalador não é confiável.");

        var pasta = Path.Combine(PastaDeDownloads, $"update-{Guid.NewGuid():N}");
        Directory.CreateDirectory(pasta);
        var caminho = Path.Combine(pasta, nome);

        // O HttpClient.Timeout NAO cobre o corpo quando se le' com ResponseHeadersRead: um servidor que para de mandar
        // bytes no meio travaria o /update do RMM para sempre. O limite vale para o download INTEIRO (cabecalhos +
        // corpo), por um token ligado ao do chamador; estourar vira "nao foi possivel baixar", o cancelamento do
        // usuario continua propagando como cancelamento.
        using var limite = CancellationTokenSource.CreateLinkedTokenSource(ct);
        limite.CancelAfter(_limite);
        var tk = limite.Token;
        try
        {
            using var resposta = await ObterSeguindoRedirecionamentosAsync(new Uri(info.UrlAsset), tk);
            if (resposta is null)
            {
                TentarApagar(caminho);
                return ResultadoDownload.Falhou("A origem do instalador não é confiável.");
            }
            resposta.EnsureSuccessStatusCode();
            var total = resposta.Content.Headers.ContentLength ?? info.TamanhoBytes;

            await using (var origem = await resposta.Content.ReadAsStreamAsync(tk))
            await using (var destino = File.Create(caminho))
            {
                var buffer = new byte[81920];
                long lidos = 0;
                int n;
                while ((n = await origem.ReadAsync(buffer, tk)) > 0)
                {
                    await destino.WriteAsync(buffer.AsMemory(0, n), tk);
                    lidos += n;
                    progresso?.Report((lidos, total));
                }
            }
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            TentarApagar(caminho);
            throw;
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException or OperationCanceledException)
        {
            // OperationCanceledException aqui = o LIMITE estourou (o cancelamento do usuario ja' saiu no catch acima).
            TentarApagar(caminho);
            return ResultadoDownload.Falhou("Não foi possível baixar a atualização. Confira a conexão com a internet.");
        }

        return VerificarEFinalizar(caminho, info.Sha256Esperado);
    }

    /// <summary>
    /// GET seguindo os redirecionamentos A MAO: cada salto (inclusive o primeiro) tem de passar por
    /// <see cref="HostPermitido"/> ANTES de a requisicao sair, e no maximo <see cref="MaxRedirecionamentos"/> saltos.
    /// <c>null</c> = recusado (host fora da lista, <c>Location</c> ausente/invalido ou saltos demais) — a requisicao
    /// recusada nunca e' feita.
    /// </summary>
    private async Task<HttpResponseMessage?> ObterSeguindoRedirecionamentosAsync(Uri inicio, CancellationToken ct)
    {
        var atual = inicio;
        for (var salto = 0; salto <= MaxRedirecionamentos; salto++)
        {
            if (!HostPermitido(atual.AbsoluteUri)) return null;
            var resposta = await _download.GetAsync(atual, HttpCompletionOption.ResponseHeadersRead, ct);
            if (!EhRedirecionamento(resposta.StatusCode)) return resposta;
            var proximo = ProximoSalto(atual, resposta.Headers.Location);
            resposta.Dispose();
            if (proximo is null) return null;
            atual = proximo;
        }
        return null; // saltos demais
    }

    internal static bool EhRedirecionamento(HttpStatusCode s) =>
        s is HttpStatusCode.MovedPermanently or HttpStatusCode.Found or HttpStatusCode.SeeOther
            or HttpStatusCode.TemporaryRedirect or HttpStatusCode.PermanentRedirect;

    /// <summary>O destino de um redirecionamento (relativo resolvido contra o atual), so' se o host for permitido.</summary>
    internal static Uri? ProximoSalto(Uri atual, Uri? location)
    {
        if (location is null) return null;
        var destino = location.IsAbsoluteUri ? location : new Uri(atual, location);
        return HostPermitido(destino.AbsoluteUri) ? destino : null;
    }

    /// <summary>Compara o SHA-256 do arquivo JA' em disco (sem diferenciar maiusculas). Divergente: apaga e recusa.
    /// Batendo: devolve o UNICO <see cref="ArquivoVerificado"/> possivel.</summary>
    public static ResultadoDownload VerificarEFinalizar(string caminhoBaixado, string sha256Esperado) =>
        ArquivoVerificado.VerificarEFinalizar(caminhoBaixado, sha256Esperado);

    /// <summary>Apaga o arquivo e, se ficou vazia, a pasta <c>update-&lt;guid&gt;</c> dele (nada de lixo no %TEMP%).</summary>
    private static void TentarApagar(string caminho)
    {
        try
        {
            if (File.Exists(caminho)) File.Delete(caminho);
            var pasta = Path.GetDirectoryName(caminho);
            if (pasta is not null && Path.GetFileName(pasta).StartsWith("update-", StringComparison.Ordinal)
                && Directory.Exists(pasta) && !Directory.EnumerateFileSystemEntries(pasta).Any())
                Directory.Delete(pasta);
        }
        catch { /* melhor esforco */ }
    }

    /// <summary>So' o componente de ARQUIVO, e igual ao original: nome com separador, unidade ou UNC e' recusado
    /// inteiro (<c>Path.Combine</c> descartaria a pasta diante de um caminho absoluto).</summary>
    internal static string? SanearNomeDoAsset(string? nome)
    {
        if (string.IsNullOrWhiteSpace(nome)) return null;
        var arquivo = Path.GetFileName(nome);
        if (string.IsNullOrEmpty(arquivo)) return null;
        if (!string.Equals(arquivo, nome, StringComparison.Ordinal)) return null;
        if (Path.IsPathRooted(arquivo)) return null;
        return arquivo;
    }

    private static readonly string[] HostsPermitidos =
        ["github.com", "objects.githubusercontent.com", "release-assets.githubusercontent.com", "api.github.com"];

    /// <summary>HTTPS e host na lista (igual ou subdominio por SUFIXO com ponto — nunca <c>Contains</c>, que aceitaria
    /// "evilgithub.com" ou "github.com.evil.com").</summary>
    internal static bool HostPermitido(string? url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) return false;
        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) return false;
        foreach (var h in HostsPermitidos)
        {
            if (string.Equals(uri.Host, h, StringComparison.OrdinalIgnoreCase)) return true;
            if (uri.Host.EndsWith("." + h, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    /// <summary>Instalador BAIXADO cujo SHA-256 confere. O construtor e' privado e o unico que o chama e'
    /// <see cref="VerificarEFinalizar"/>, declarado DENTRO desta classe aninhada (em C# a classe externa nao ve'
    /// os privados da aninhada). Quem lanca o instalador so' aceita este tipo — nunca uma string crua.</summary>
    public sealed class ArquivoVerificado
    {
        public string Caminho { get; }

        /// <summary>O SHA-256 que o arquivo tinha na verificacao (o da release).</summary>
        public string Sha256 { get; }

        private ArquivoVerificado(string caminho, string sha256) { Caminho = caminho; Sha256 = sha256; }

        /// <summary>
        /// Confere o hash DE NOVO e TRAVA o arquivo: entre o download e o encerramento do app (o usuario salvando
        /// documentos) o arquivo fica no %TEMP%, onde outro processo poderia troca-lo. O hash e' calculado pelo MESMO
        /// handle que fica aberto (<see cref="FileShare.Read"/>: ninguem escreve, renomeia nem apaga enquanto ele vive),
        /// e quem lanca o segura ate' o <c>Process.Start</c> voltar — o arquivo que o Windows executa (inclusive
        /// durante a pergunta do UAC) e' o que foi conferido. Nao confere: apaga e devolve <c>null</c>.
        /// </summary>
        public IDisposable? ReconferirETravar()
        {
            FileStream? trava = null;
            try
            {
                trava = new FileStream(Caminho, FileMode.Open, FileAccess.Read, FileShare.Read);
                var real = Convert.ToHexString(SHA256.HashData(trava)).ToLowerInvariant();
                if (string.Equals(real, Sha256, StringComparison.OrdinalIgnoreCase)) return trava;
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* sumiu ou travado: nao confere */ }
            trava?.Dispose();
            TentarApagar(Caminho);
            return null;
        }

        /// <summary><see cref="ReconferirETravar"/> sem segurar a trava.</summary>
        public bool Reconferir()
        {
            using var t = ReconferirETravar();
            return t is not null;
        }

        private static string Hash(string caminho)
        {
            using var s = File.OpenRead(caminho);
            return Convert.ToHexString(SHA256.HashData(s)).ToLowerInvariant();
        }

        internal static ResultadoDownload VerificarEFinalizar(string caminho, string esperado)
        {
            var real = Hash(caminho);
            if (!string.Equals(real, esperado.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                TentarApagar(caminho);
                return ResultadoDownload.Falhou(
                    "O instalador baixado não confere com o código de verificação: pode estar corrompido ou adulterado. Ele foi removido.");
            }
            return ResultadoDownload.Ok(new ArquivoVerificado(caminho, real));
        }
    }
}

// ---- costura falsificavel: a consulta da ultima release ------------------------------------------------

/// <summary>A ultima release, reduzida ao que o servico usa. <c>NomeAsset</c>/<c>UrlAsset</c> nulos quando nao ha'
/// nenhum <c>.exe</c> entre os assets.</summary>
public sealed record ReleaseMaisRecente(string Tag, string? Corpo, string? NomeAsset, string? UrlAsset, long TamanhoAsset);

/// <summary>De onde vem a ultima release. Producao: <see cref="FonteGitHub"/>; testes: um falso (nenhum teste toca a
/// rede).</summary>
public interface IFonteDeReleases
{
    Task<ReleaseMaisRecente?> ObterUltimaAsync(CancellationToken ct);
}

/// <summary>A fonte REAL: GET em <see cref="ServicoAtualizacao.UrlReleases"/>, 15 s de limite, com <c>User-Agent</c>
/// (sem ele a API do GitHub recusa com 403). O asset e' o PRIMEIRO cujo nome termina em <c>.exe</c>.</summary>
internal sealed class FonteGitHub : IFonteDeReleases, IDisposable
{
    private readonly HttpClient _http;

    public FonteGitHub()
    {
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        _http.DefaultRequestHeaders.UserAgent.ParseAdd($"mOffice-Atualizacao/{ServicoAtualizacao.VersaoAtualTexto()}");
        _http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
    }

    public async Task<ReleaseMaisRecente?> ObterUltimaAsync(CancellationToken ct)
    {
        using var resposta = await _http.GetAsync(ServicoAtualizacao.UrlReleases, ct);
        resposta.EnsureSuccessStatusCode();
        await using var fluxo = await resposta.Content.ReadAsStreamAsync(ct);
        using var doc = await JsonDocument.ParseAsync(fluxo, cancellationToken: ct);
        return LerRelease(doc.RootElement);
    }

    /// <summary>O JSON da API reduzido a <see cref="ReleaseMaisRecente"/> (separado para teste, sem rede).</summary>
    internal static ReleaseMaisRecente LerRelease(JsonElement raiz)
    {
        var tag = raiz.TryGetProperty("tag_name", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString() ?? "" : "";
        var corpo = raiz.TryGetProperty("body", out var b) && b.ValueKind == JsonValueKind.String ? b.GetString() : null;
        string? nome = null, url = null;
        long tamanho = 0;
        if (raiz.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
        {
            foreach (var a in assets.EnumerateArray())
            {
                var n = a.TryGetProperty("name", out var ne) && ne.ValueKind == JsonValueKind.String ? ne.GetString() : null;
                if (n is null || !n.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)) continue;
                nome = n;
                url = a.TryGetProperty("browser_download_url", out var ue) && ue.ValueKind == JsonValueKind.String ? ue.GetString() : null;
                tamanho = a.TryGetProperty("size", out var se) && se.ValueKind == JsonValueKind.Number ? se.GetInt64() : 0;
                break;
            }
        }
        return new ReleaseMaisRecente(tag, corpo, nome, url, tamanho);
    }

    public void Dispose() => _http.Dispose();
}

// ---- resultados tipados ---------------------------------------------------------------------------------

public sealed record InfoAtualizacao(
    string TagVersao, string? Notas, string UrlAsset, string NomeAsset, long TamanhoBytes, string Sha256Esperado);

public enum StatusVerificacao { Disponivel, Atualizado, Erro }

public enum TipoErroAtualizacao { SemRede, LimiteTaxa, SemAsset, SemHash }

public sealed record ResultadoVerificacao(StatusVerificacao Status, InfoAtualizacao? Info, TipoErroAtualizacao? TipoErro, string? MensagemErro)
{
    public static ResultadoVerificacao Disponivel(InfoAtualizacao info) => new(StatusVerificacao.Disponivel, info, null, null);
    public static ResultadoVerificacao Atualizado() => new(StatusVerificacao.Atualizado, null, null, null);
    public static ResultadoVerificacao Erro(TipoErroAtualizacao tipo, string mensagem) => new(StatusVerificacao.Erro, null, tipo, mensagem);
}

public enum StatusDownload { Verificado, Recusado }

public sealed record ResultadoDownload(StatusDownload Status, ServicoAtualizacao.ArquivoVerificado? Arquivo, string? MensagemErro)
{
    public static ResultadoDownload Ok(ServicoAtualizacao.ArquivoVerificado arquivo) => new(StatusDownload.Verificado, arquivo, null);
    public static ResultadoDownload Falhou(string mensagem) => new(StatusDownload.Recusado, null, mensagem);
}
