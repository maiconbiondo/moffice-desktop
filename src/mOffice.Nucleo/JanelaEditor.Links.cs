using System.Diagnostics;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace mOffice.Nucleo;

/// <summary>
/// Plano 6, Task 6 (A1/CE-2/MvP-02, SO-04/PL-16) — links do documento e navegacao. NUNCA nasce pop-up do WebView2: ele seria criado
/// fora da <see cref="GuardaEgresso"/> (que e' por <c>CoreWebView2</c>) e carregaria a pagina (VERIFICADO ao vivo na auditoria).
/// O pedido vira "abrir no navegador padrao DEPOIS de o usuario confirmar" (como o Office), e so' para o que a
/// <see cref="PoliticaDeLink"/> permite.
/// </summary>
public partial class JanelaEditor
{
    /// <summary>Confirmacao antes de abrir um link fora do mOffice (<see cref="MensagemConfirmarLink"/>): <c>true</c> = abrir.
    /// A casca injeta o dialogo do design system; sem ela, um <c>MessageBox</c>.</summary>
    public Func<string, bool>? ConfirmarLink { get; init; }

    /// <summary>Costura de teste: no lugar do <c>Process.Start</c>. Recebe o que iria ao Windows (URL, ou o caminho local de um <c>file:</c>).</summary>
    internal Action<string>? AbrirNoSistemaParaTeste { get; set; }

    public static string MensagemConfirmarLink(string uri) => $"Abrir este endereço fora do mOffice?\n\n{uri}";
    public static string MensagemLinkRecusado(string uri) => $"O mOffice não abre este link por segurança: {uri}";

    // Uma confirmacao por vez: um documento que chama window.open em laco nao empilha dialogos (os pedidos seguintes sao descartados).
    private bool _linkEmConfirmacao;

    /// <summary>Paginas que carregam no documento de cima: a entrada do editor e a de teste do PoC. Qualquer outra navegacao de topo e' tratada como link.</summary>
    internal static bool EhPaginaDoHost(string? uri) =>
        uri is not null
        && (uri.StartsWith("https://moffice.local/web-apps/apps/api/documents/index.html", StringComparison.OrdinalIgnoreCase)
            || uri.StartsWith("https://moffice-poc.local/", StringComparison.OrdinalIgnoreCase)
            || uri.Equals("about:blank", StringComparison.OrdinalIgnoreCase));

    private static readonly string[] HostsInternos = ["moffice.local", "moffice-poc.local", "moffice-doc.local", "moffice-fontes.local"];

    /// <summary>Navegacao de QUADRO que o editor faz por dentro: paginas dos hosts virtuais do mOffice e esquemas locais (about/blob/data/ascdesktop).
    /// O resto (http/https de fora, mailto, ftp, esquemas de aplicativo) e' link: cancela e vira confirmacao.</summary>
    internal static bool EhNavegacaoInternaDeQuadro(string? uri)
    {
        if (string.IsNullOrEmpty(uri) || !Uri.TryCreate(uri, UriKind.Absolute, out var u)) return false;
        switch (u.Scheme.ToLowerInvariant())
        {
            case "about": case "blob": case "data": case "ascdesktop": return true;
            case "http": case "https": return HostsInternos.Contains(u.Host, StringComparer.OrdinalIgnoreCase);
            default: return false;
        }
    }

    private void InstalarLinks(CoreWebView2 wv)
    {
        wv.NewWindowRequested += (_, e) =>
        {
            e.Handled = true; // sem janela nova, nunca
            var uri = e.Uri;
            Dispatcher.BeginInvoke(() => { if (!_fechada) TratarLink(uri, "janela nova"); });
        };
        // Esquema de aplicativo externo (mailto:, ms-*:, zoommtg: ...): o WebView2 perguntaria ao Windows por conta propria. Nunca: vira link.
        wv.LaunchingExternalUriScheme += (_, e) =>
        {
            e.Cancel = true;
            var uri = e.Uri;
            Dispatcher.BeginInvoke(() => { if (!_fechada) TratarLink(uri, "esquema externo"); });
        };
        // Navegacao de iframe para fora (o editor e' um iframe): a GuardaEgresso ja' bloquearia a rede, mas o usuario veria um quadro 403 e o
        // pedido sumiria; aqui cancela e vira confirmacao como qualquer link.
        wv.FrameNavigationStarting += (_, e) =>
        {
            if (EhNavegacaoInternaDeQuadro(e.Uri)) return;
            e.Cancel = true;
            // file: num quadro: cancela sem perguntar (nenhum popup por conta de um iframe); so' log.
            if (e.Uri.StartsWith("file:", StringComparison.OrdinalIgnoreCase)) { _arquivos.Log("LINK navegacao de quadro para file: cancelada (so' log): " + Resumo(e.Uri)); return; }
            var uri = e.Uri;
            Dispatcher.BeginInvoke(() => { if (!_fechada) TratarLink(uri, "navegacao de quadro"); });
        };
        // Downloads da pagina. O editor salva ARQUIVOS pela ponte (Salvar como), mas alguns recursos do sdk baixam do lado do cliente por
        // <a download href="data:|blob:">: "Salvar como imagem" (asc_SaveDrawingAsPicture) e o auxiliar de blob (AscCommon, ex. exportar HTML).
        // data:/blob: -> o host pergunta ONDE gravar (nunca a pasta Downloads em silencio); qualquer outro esquema (http/https...) -> cancelado.
        // M4 da revisao final do Plano 6: o dialogo NAO abre dentro do evento (modal aninhado no tratador do WebView2): adiamento + BeginInvoke.
        // Um download por vez (o segundo, com o dialogo aberto, e' cancelado); so' extensoes da lista; nome sugerido saneado.
        wv.DownloadStarting += (_, e) =>
        {
            var uri = e.DownloadOperation.Uri ?? "";
            if (!uri.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && !uri.StartsWith("blob:", StringComparison.OrdinalIgnoreCase))
            {
                e.Cancel = true;
                _arquivos.Log("DOWNLOAD cancelado (a pagina nao baixa arquivos): " + Resumo(uri));
                return;
            }
            e.Handled = true; // sem o balao de downloads do navegador
            var bruto = Path.GetFileName(e.ResultFilePath ?? "");
            if (NomeDeDownloadSeguro(bruto) is not { } sugerido)
            {
                e.Cancel = true;
                _arquivos.Log("DOWNLOAD cancelado (extensao nao permitida): " + Resumo(bruto));
                return;
            }
            if (_downloadEmEscolha)
            {
                e.Cancel = true;
                _arquivos.Log("DOWNLOAD cancelado: ja ha um dialogo de download aberto (" + sugerido + ")");
                return;
            }
            _downloadEmEscolha = true;
            var adiamento = e.GetDeferral();
            Dispatcher.BeginInvoke(() =>
            {
                try
                {
                    string? destino = null;
                    if (!_fechada)
                    {
                        try { destino = (EscolherDestinoDeDownloadParaTeste ?? EscolherDestinoDeDownload)(sugerido); }
                        catch (Exception ex) { destino = null; _arquivos.Log("DOWNLOAD erro ao escolher o destino (cancelado): " + ex.Message); }
                    }
                    if (string.IsNullOrWhiteSpace(destino)) { e.Cancel = true; _arquivos.Log("DOWNLOAD cancelado pelo usuario: " + sugerido); return; }
                    if (!ExtensoesDeDownload.Contains(Path.GetExtension(destino).TrimStart('.')))
                    {
                        e.Cancel = true; _arquivos.Log("DOWNLOAD cancelado (destino com extensao nao permitida): " + Path.GetFileName(destino));
                        return;
                    }
                    e.ResultFilePath = destino;
                    _arquivos.Log("DOWNLOAD gravando em " + destino);
                }
                catch (Exception ex) { try { e.Cancel = true; } catch (Exception) { } _arquivos.Log("DOWNLOAD erro (cancelado): " + ex.Message); }
                finally
                {
                    _downloadEmEscolha = false;
                    try { adiamento.Complete(); } catch (Exception ex) { _arquivos.Log("DOWNLOAD adiamento: " + ex.Message); }
                }
            });
        };
        wv.NavigationStarting += (_, e) =>
        {
            if (EhPaginaDoHost(e.Uri)) return;
            e.Cancel = true; // a pagina do editor continua; o destino vira link
            var uri = e.Uri;
            Dispatcher.BeginInvoke(() => { if (!_fechada) TratarLink(uri, "navegacao de topo"); });
        };
    }

    private void TratarLink(string uri, string origem)
    {
        var acao = PoliticaDeLink.Classificar(uri);
        _arquivos.Log($"LINK {origem}: {acao} {uri}");
        switch (acao)
        {
            case AcaoDeLink.Descartar:
                return;
            case AcaoDeLink.Recusar:
                AvisarUsuario(MensagemLinkRecusado(uri));
                return;
        }
        if (_linkEmConfirmacao) { _arquivos.Log("LINK descartado: ja ha uma confirmacao aberta"); return; }
        _linkEmConfirmacao = true;
        try
        {
            bool sim;
            try { sim = (ConfirmarLink ?? ConfirmarLinkPadrao)(uri); }
            catch (Exception ex) { _arquivos.Log("LINK erro na confirmacao (nao abre): " + ex.Message); return; }
            if (!sim) { _arquivos.Log("LINK nao confirmado: " + uri); return; }
        }
        finally { _linkEmConfirmacao = false; }

        // Abre o que foi CONFIRMADO: URL absoluta normalizada, ou o caminho local de um file:.
        var u = new Uri(uri);
        var alvo = acao == AcaoDeLink.Arquivo ? u.LocalPath : u.AbsoluteUri;
        try { (AbrirNoSistemaParaTeste ?? AbrirNoSistema)(alvo); _arquivos.Log("LINK aberto no sistema: " + alvo); }
        catch (Exception ex) { _arquivos.Log("LINK nao abriu: " + ex.Message); AvisarUsuario($"Não foi possível abrir o endereço: {uri}"); }
    }

    // Um dialogo de download por vez (M4).
    private bool _downloadEmEscolha;

    /// <summary>O que a pagina pode baixar pelo host (M4): figuras do "Salvar como imagem", o HTML/TXT/CSV do auxiliar de blob e os formatos de
    /// documento que o mOffice abre. Nada executavel; sem "Todos os arquivos" no dialogo.</summary>
    internal static readonly IReadOnlySet<string> ExtensoesDeDownload = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "png", "jpg", "jpeg", "gif", "bmp", "svg", "html", "htm", "txt", "csv", "docx", "xlsx", "pptx", "odt", "ods", "odp", "pdf" };

    private static readonly HashSet<string> NomesReservados = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    };

    /// <summary>
    /// Nome sugerido pela PAGINA (<c>&lt;a download="..."&gt;</c>) saneado para o dialogo (M4), ou <c>null</c> se a extensao nao esta' em
    /// <see cref="ExtensoesDeDownload"/>. Tira controles de direcao do texto (RLO/LRO/LRE/RLE/PDF/LRI/RLI/FSI/PDI/LRM/RLM/ALM — "fotoexe.png"
    /// virando "foto<RLO>gnp.exe" na tela), outros caracteres de controle e invalidos em nome de arquivo, pontos e espacos no fim, e nao deixa
    /// nome reservado do Windows (CON, NUL, COM1...). Nome vazio vira "arquivo". `internal static`: testavel sem janela.
    /// </summary>
    internal static string? NomeDeDownloadSeguro(string? sugerido)
    {
        var invalidos = Path.GetInvalidFileNameChars();
        var sb = new System.Text.StringBuilder();
        foreach (var c in sugerido ?? "")
        {
            if (char.IsControl(c) || c is '‎' or '‏' or '؜' || c is >= '‪' and <= '‮' || c is >= '⁦' and <= '⁩') continue;
            sb.Append(Array.IndexOf(invalidos, c) >= 0 ? '_' : c);
        }
        var nome = sb.ToString().Trim().TrimEnd('.', ' ');
        var ext = Path.GetExtension(nome).TrimStart('.');
        if (!ExtensoesDeDownload.Contains(ext)) return null;
        var baseNome = Path.GetFileNameWithoutExtension(nome).TrimEnd('.', ' ');
        if (baseNome.Length == 0) baseNome = "arquivo";
        if (NomesReservados.Contains(baseNome.Split('.')[0])) baseNome = "_" + baseNome;
        return baseNome + "." + ext.ToLowerInvariant();
    }

    /// <summary>Costura de teste do destino do download (data:/blob:): recebe o nome sugerido, devolve o caminho ou null (cancelar).</summary>
    internal Func<string, string?>? EscolherDestinoDeDownloadParaTeste { get; set; }

    private static string Resumo(string uri) => uri.Length <= 120 ? uri : uri[..120] + $"... ({uri.Length} caracteres)";

    /// <summary>"Salvar imagem": o dialogo do Windows, com o filtro da extensao sugerida e a pasta do documento (ou Imagens) como inicio.</summary>
    private string? EscolherDestinoDeDownload(string sugerido)
    {
        var ext = Path.GetExtension(sugerido).TrimStart((char)46).ToLowerInvariant();
        var ehImagem = ext is "png" or "jpg" or "jpeg" or "gif" or "bmp" or "svg";
        var pasta = _sessao is { Novo: false } s && Path.GetDirectoryName(s.ArquivoOrigem) is { } d && Directory.Exists(d) ? d : Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title = ehImagem ? "Salvar imagem" : "Salvar arquivo",
            FileName = sugerido,
            InitialDirectory = pasta,
            OverwritePrompt = true,
            // So' a extensao do arquivo (M4: sem "Todos os arquivos" — o nome ja' passou por NomeDeDownloadSeguro).
            Filter = $"{ext.ToUpperInvariant()} (*.{ext})|*.{ext}",
            AddExtension = true,
            DefaultExt = ext,
        };
        return dlg.ShowDialog(this) == true ? dlg.FileName : null;
    }

    private bool ConfirmarLinkPadrao(string uri) =>
        MessageBox.Show(this, MensagemConfirmarLink(uri), NomeApp, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes; // Enter = Nao

    private static void AbrirNoSistema(string alvo) => Process.Start(new ProcessStartInfo(alvo) { UseShellExecute = true })?.Dispose();
}
