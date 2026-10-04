namespace mOffice.Nucleo;

/// <summary>
/// Plano 5, Task 1 (fix round 1): como o editor recebe o <c>AllFonts.js</c> SEM que ele precise existir dentro
/// do motor.
///
/// Antes, <c>CatalogoFontes.GarantirNoMotor</c> COPIAVA o catalogo para
/// <c>&lt;motor&gt;\editors\sdkjs\common\AllFonts.js</c>, porque o editor pede
/// <c>https://moffice.local/sdkjs/common/AllFonts.js</c> (modulo "allfonts" do requirejs) e o host virtual
/// <c>moffice.local</c> e' um mapeamento de pasta (<c>SetVirtualHostNameToFolderMapping</c>) — que NAO passa
/// pelo <c>WebResourceRequested</c> (medido: ver <see cref="GuardaEgresso"/>). Instalado em Program Files isso
/// quebra: um usuario sem elevacao nao escreve la', e abrir a pasta para escrita criaria um JS que o editor
/// executa e que qualquer usuario do PC poderia trocar.
///
/// Agora: um script de DocumentCreated (<see cref="ScriptReescrita"/>, antes de qualquer script da pagina)
/// troca o <c>src</c> de todo <c>&lt;script&gt;</c> que aponte para <c>https://moffice.local/.../sdkjs/common/AllFonts.js</c>
/// por <see cref="UrlAllFonts"/>, num host que NAO e' mapeado — e por isso passa pelo
/// <c>WebResourceRequested</c>, onde a janela responde com o arquivo do catalogo
/// (<see cref="mOffice.Conversor.ConfiguracaoMotor.PastaFontes"/>). O requirejs cria o no e atribui
/// <c>node.src = url</c> (web-apps/vendor/requirejs/require.js 9.4.0.129: <c>r.src=i</c>), entao o setter do
/// prototipo pega esse caminho. LIMITE DE ESCOPO: so' <c>.src</c> atribuido por JS; um <c>&lt;script src&gt;</c>
/// escrito no markup (inserido pelo parser) NAO passa pelo setter. As duas paginas do motor assim —
/// <c>web-apps/apps/api/documents/preload.html</c> e <c>cache-scripts.html</c>, dormentes no mOffice — sao
/// excluidas do pacote por <c>tools/installer/montar-pacote.ps1</c> (e o <c>PacoteTests</c> confere). O service worker do sdkjs so' mexe em URLs com segmento de versao (e ignora
/// AllFonts.js no desktop), entao nao entra no meio.
///
/// Seguranca: o host so' serve UMA URL, exata; o caminho do arquivo e' sempre
/// <c>&lt;PastaFontes&gt;\AllFonts.js</c> — nada da requisicao entra na montagem do caminho. Qualquer outra URL
/// do host recebe 404 (sem cair na rede).
/// </summary>
public static class CatalogoNoEditor
{
    public const string Host = "moffice-fontes.local";
    public const string Origem = "https://" + Host + "/";
    public const string UrlAllFonts = Origem + "AllFonts.js";
    public const string NomeArquivo = "AllFonts.js";

    /// <summary>A URI e' do host do catalogo (qualquer caminho)? Quem responde decide 200 ou 404.</summary>
    public static bool EhDoHost(string uri) => uri.StartsWith(Origem, StringComparison.OrdinalIgnoreCase);

    /// <summary>Arquivo a servir para <paramref name="uri"/>: <c>&lt;pastaCatalogo&gt;\AllFonts.js</c> SO' para a
    /// URL exata <see cref="UrlAllFonts"/> (sem query, sem fragmento, sem outro caminho); <c>null</c> para
    /// qualquer outra coisa. Nada de <paramref name="uri"/> vai para o caminho devolvido.</summary>
    public static string? Resolver(string uri, string pastaCatalogo)
        => string.Equals(uri, UrlAllFonts, StringComparison.Ordinal) ? Path.Combine(pastaCatalogo, NomeArquivo) : null;

    /// <summary>Script de DocumentCreated (roda em TODO documento: janela de cima e quadro do editor). Idempotente.</summary>
    public const string ScriptReescrita = """
        (function () {
          var P = HTMLScriptElement.prototype;
          if (P.__mofficeAllFonts) return;
          var d = Object.getOwnPropertyDescriptor(P, "src");
          if (!d || !d.set || !d.get) return;
          Object.defineProperty(P, "__mofficeAllFonts", { value: true });
          Object.defineProperty(P, "src", {
            configurable: true, enumerable: d.enumerable, get: d.get,
            set: function (v) {
              try {
                var u = new URL(String(v), document.baseURI);
                if (u.protocol === "https:" && u.hostname === "moffice.local" && /\/sdkjs\/common\/AllFonts\.js$/i.test(u.pathname))
                  v = "https://moffice-fontes.local/AllFonts.js";
              } catch (e) { }
              d.set.call(this, v);
            }
          });
        })();
        """;
}
