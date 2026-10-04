using Microsoft.Web.WebView2.Core;

namespace mOffice.Nucleo;

/// <summary>
/// Auditoria de egresso: recebe no <see cref="CoreWebView2.WebResourceRequested"/> (filtro "*") o que o
/// WebView2 entrega a esse evento. Uma URI cuja origem está em <see cref="OrigensPermitidas"/> passa sem
/// resposta (segue para o mapeamento de pasta / handlers); qualquer outra é negada com 403 (com
/// <c>Access-Control-Allow-Origin</c>, para que um <c>fetch()</c> da página consiga ler o status em vez de
/// receber um erro de CORS opaco) e fica registrada em <see cref="Bloqueadas"/> e no log.
///
/// Observado nesta revisão (teste de fumaça, `moffice-poc.local` mapeado via
/// <see cref="CoreWebView2.SetVirtualHostNameToFolderMapping"/>): NEM a navegação de nível superior para o
/// host virtual NEM o subrecurso `ponte-teste.js` carregado dele via &lt;script src&gt; passam pelo
/// <c>WebResourceRequested</c> — a página carrega e o script roda normalmente, mas nenhuma das duas URIs
/// aparece em <see cref="Permitidas"/> nem no log (confirmado: o log da execução só tem a linha JS do
/// grampo.js e a linha BLOQUEADA do fetch para example.com). Ou seja: este guard prova que egresso externo
/// é negado e logado, mas **não** prova, e não deve ser lido como provando, que todo tráfego de um host
/// virtual mapeado passa por auditoria — recursos servidos por
/// <c>SetVirtualHostNameToFolderMapping</c> aparentemente são resolvidos sem disparar este evento.
/// </summary>
public sealed class GuardaEgresso
{
    /// <summary><c>https://moffice-dic.local/</c> (Plano 6, Task 8): dicionarios de hifenizacao, pasta LOCAL do motor mapeada pela janela.</summary>
    public static readonly string[] OrigensPermitidas = ["https://moffice.local/", "https://moffice-doc.local/", "https://moffice-poc.local/", "https://moffice-dic.local/", CatalogoNoEditor.Origem, "file:///", "ascdesktop://"];

    /// <summary>
    /// Servidor local opcional para as origens PERMITIDAS que ninguem mais atende — o esquema nativo
    /// <c>ascdesktop://</c> (<see cref="ServidorAscDesktop"/>) e o host do catalogo de fontes (<see cref="CatalogoNoEditor"/>). Fica aqui, e nao num segundo handler do
    /// <c>WebResourceRequested</c>, para que <see cref="OrigensPermitidas"/> continue sendo a UNICA lista que
    /// decide o que passa: quem serve e' chamado depois da decisao, nunca em paralelo com ela.
    /// Devolver <c>null</c> deixa a requisicao seguir para o mapeamento de pasta / handlers padrao.
    /// </summary>
    public Func<CoreWebView2, string, CoreWebView2WebResourceResponse?>? ServidorLocal { get; set; }
    private readonly object _trava = new();
    private readonly List<string> _bloqueadas = [], _permitidas = [];
    private string? _log;
    public IReadOnlyList<string> Bloqueadas { get { lock (_trava) return _bloqueadas.ToArray(); } }
    public IReadOnlyList<string> Permitidas { get { lock (_trava) return _permitidas.ToArray(); } }

    public void Instalar(CoreWebView2 wv, string arquivoLog)
    {
        _log = arquivoLog;
        Directory.CreateDirectory(Path.GetDirectoryName(arquivoLog)!);
        wv.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
        // Esquema customizado nao e' coberto pelo filtro "*" — precisa ser pedido pelo nome.
        wv.AddWebResourceRequestedFilter("ascdesktop://*", CoreWebView2WebResourceContext.All);
        wv.WebResourceRequested += (_, e) =>
        {
            var uri = e.Request.Uri;
            if (OrigensPermitidas.Any(o => uri.StartsWith(o, StringComparison.OrdinalIgnoreCase)))
            {
                Registrar(_permitidas, "PERMITIDA", uri);
                // O servidor local le ARQUIVO DO DISCO, e isto roda na thread da interface: um arquivo
                // apagado, travado por outro processo ou sem permissao viraria excecao nao tratada dentro do
                // manipulador de evento e derrubaria a janela (achado da revisao final 2026-09-17). Segunda
                // rede, alem do try/catch do proprio servidor: aqui a falha vira 403 + log, nunca crash.
                CoreWebView2WebResourceResponse? resposta;
                try { resposta = ServidorLocal?.Invoke(wv, uri); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    lock (_trava) Escrever("ERRO DO SERVIDOR LOCAL " + uri + ": " + ex.Message);
                    e.Response = wv.Environment.CreateWebResourceResponse(null, 403, "Falha ao ler o recurso local", "Access-Control-Allow-Origin: *");
                    return;
                }
                // O host do catalogo NAO e' mapeado: sem resposta do servidor local a requisicao iria para a
                // rede (DNS de moffice-fontes.local). Egresso zero: sem servidor, 404 aqui mesmo.
                if (resposta is null && CatalogoNoEditor.EhDoHost(uri))
                    resposta = wv.Environment.CreateWebResourceResponse(null, 404, "Nao encontrado", "Access-Control-Allow-Origin: *");
                if (resposta is not null) e.Response = resposta;
                return; // sem servidor local: segue para o mapeamento de pasta / handlers
            }
            // Ordem importa: monta e atribui a resposta 403 ANTES de registrar/logar, para que uma falha ao
            // construir a resposta nunca deixe uma requisição registrada como bloqueada tendo, na verdade, saído.
            // Access-Control-Allow-Origin e necessario para que um fetch() cross-origin da pagina consiga LER
            // o status 403 (sem CORS, o browser transforma a resposta bloqueada num TypeError opaco e o teste
            // nao consegue provar nada sobre o status real).
            e.Response = wv.Environment.CreateWebResourceResponse(null, 403, "Bloqueado pela guarda de egresso do mOffice", "Access-Control-Allow-Origin: *");
            Registrar(_bloqueadas, "BLOQUEADA", uri);
        };
    }

    private void Registrar(List<string> lista, string rotulo, string uri)
    {
        lock (_trava)
        {
            lista.Add(uri);
            Escrever(rotulo + " " + uri);
        }
    }

    private void Escrever(string linha)
    {
        if (_log is null) return;
        File.AppendAllText(_log, $"{DateTime.Now:HH:mm:ss.fff} {linha}\n");
    }
}
