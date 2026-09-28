using System.Collections.Concurrent;
using System.Text;

namespace mOffice.Nucleo;

/// <summary>
/// Midia do documento aberto: a pasta <c>&lt;recovery&gt;/media</c> que o x2t preenche na conversao e onde
/// as imagens COLADAS/INSERIDAS pelo usuario tambem vao parar.
///
/// Por que existe (pergunta aberta do Marco 0, relatorio da Task 5A secao 5): o sdkjs guarda a pasta que
/// recebe no <c>DesktopOfflineAppDocumentEndLoad</c> em <c>AscCommon.Ss.tia</c> e, se ela nao comeca com
/// <c>file:</c>, prefixa <c>file://</c> (sdkjs/word/sdk-all.js:
/// <c>window.DesktopOfflineAppDocumentEndLoad=function(a,b,d){...AscCommon.Ss.tia=a; 0!=AscCommon.Ss.tia
/// .indexOf("file:")&amp;&amp;(...AscCommon.Ss.tia="file://"+AscCommon.Ss.tia)...}</c>). Cada imagem sai de
/// <c>tia + "/media/" + nome</c> (<c>fe.RS</c>/<c>fe.vdc</c> no mesmo arquivo). No app oficial (CEF) isso
/// funciona porque a pagina do editor TAMBEM e' <c>file://</c>; aqui a pagina e' <c>https://moffice.local</c>
/// e o Chromium recusa subrecurso <c>file://</c> — medido na Task 5A: <c>fetch</c> da <c>TypeError</c> e
/// <c>&lt;img&gt;</c> da <c>onerror</c>, sem chegar ao disco.
///
/// A pasta de recovery JA' e' servida por <c>https://moffice-doc.local</c> (mapeamento de host virtual, feito
/// no <see cref="JanelaEditor.AbrirAsync"/>). Como a condicao do <c>EndLoad</c> so' aceita string que comece
/// com <c>file:</c>, nao da para passar o host virtual como argumento: a base e' TROCADA logo depois do
/// <c>EndLoad</c> (uma propriedade, sem patch nenhum no motor) — ver
/// <see cref="JanelaEditor"/>.<c>ChamarEndLoadAsync</c>.
/// </summary>
public static class Midia
{
    /// <summary>Base das URLs do documento aberto, SEM barra final: o sdkjs monta <c>tia + "/media/" + nome</c>
    /// (<c>fe.RS</c>), entao uma barra final aqui viraria <c>//media/</c> na URL. O host e' o mesmo mapeado em
    /// <see cref="JanelaEditor.AbrirAsync"/> e autorizado em <see cref="GuardaEgresso.OrigensPermitidas"/>.</summary>
    public const string BaseUrlDocumento = "https://moffice-doc.local";

    /// <summary>Nome da pasta de midia dentro da pasta de recovery — o mesmo do app oficial
    /// (<c>m_sLocalFileFolderWithoutFile + L"/media/"</c>, client_renderer_wrapper.cpp).</summary>
    public const string PastaMedia = "media";

    /// <summary>Prefixo das URLs de midia deste documento (o que a instrumentacao de rede conta).</summary>
    public static string BaseUrlMedia => BaseUrlDocumento + "/" + PastaMedia + "/";

    public static bool EhMidiaDoDocumento(string? url) =>
        url is not null && url.StartsWith(BaseUrlMedia, StringComparison.OrdinalIgnoreCase);

    public static string PastaDeMidia(SessaoDocumento sessao) => Path.Combine(sessao.PastaRecovery, PastaMedia);

    // Serializa a reserva do nome: duas colagens seguidas (ou uma colagem enquanto o editor ainda grava outra
    // imagem) escolheriam o MESMO indice se lessem a pasta ao mesmo tempo, e a segunda gravacao sobrescreveria
    // a primeira — o documento ficaria com duas figuras iguais.
    //
    // Trava POR SESSAO, nao mais uma UNICA global (minor 4 da revisao 1): antes, colar uma imagem na janela A
    // esperava a janela B terminar de gravar a dela, mesmo as duas sendo documentos completamente
    // independentes, em pastas de recovery diferentes — nao ha corrida real entre elas, so' contencao
    // artificial. Chave = pasta de recovery (unica por SessaoDocumento, e' o mesmo valor que
    // <see cref="PastaDeMidia"/> usa). Nao removida quando a sessao fecha: seriam poucos SemaphoreSlim na vida
    // do processo do PoC (um por documento aberto), e coordenar a remocao com o fechamento da janela
    // (SessaoDocumento.ApagarRecovery) por esse volume nao compensa.
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> _travasPorSessao = new(StringComparer.OrdinalIgnoreCase);

    private static SemaphoreSlim TravaDaSessao(SessaoDocumento sessao) =>
        _travasPorSessao.GetOrAdd(sessao.PastaRecovery, _ => new SemaphoreSlim(1, 1));

    /// <summary>
    /// Grava <paramref name="bytes"/> em <c>&lt;recovery&gt;/media</c> e devolve o NOME do arquivo (nunca a
    /// URL: o sdkjs monta <c>tia + "/media/" + &lt;isto&gt;</c>).
    ///
    /// Numeracao igual a do app oficial (client_renderer_wrapper.cpp, <c>GetLocalImageUrlLocal</c>:
    /// <c>L"image" + std::to_wstring(m_nLocalImagesNextIndex++) + L".png"</c>, com
    /// <c>m_nLocalImagesNextIndex = (int)arMedia.size() + 1</c> em <c>LocalFileRecoverFolder</c>): o proximo
    /// indice e' a QUANTIDADE de arquivos ja' em media/ mais 1 — ou seja, a midia do proprio documento (a que
    /// o x2t escreveu) nunca e' sobrescrita.
    ///
    /// A extensao sai do CONTEUDO (<see cref="ExtensaoDaImagem"/>), nao do que o chamador disse: o chamador,
    /// aqui, e' a ponte — e do outro lado dela esta' conteudo web. Um "nome" vindo de la' nunca vira parte de
    /// caminho. Conteudo que nao e' imagem e' recusado com <see cref="InvalidDataException"/>, como o
    /// <c>"error"</c> que o app oficial devolve quando o <c>CImageFileFormatChecker</c> nao reconhece o arquivo.
    ///
    /// Todos os <c>await</c> vao com <c>ConfigureAwait(false)</c>: a ponte chama isto DENTRO de uma chamada
    /// sincrona de objeto de host (thread da interface, renderer bloqueado esperando o retorno) e bloqueia no
    /// resultado — com a continuacao marcada para voltar a thread da interface seria impasse garantido.
    /// </summary>
    public static async Task<string> GravarImagemAsync(SessaoDocumento sessao, byte[] bytes, string extensao)
    {
        ArgumentNullException.ThrowIfNull(sessao);
        ArgumentNullException.ThrowIfNull(bytes);
        var ext = ExtensaoDaImagem(bytes, extensao)
                  ?? throw new InvalidDataException("conteudo nao reconhecido como imagem (" + bytes.Length + " bytes)");
        var pasta = PastaDeMidia(sessao);
        Directory.CreateDirectory(pasta);
        var trava = TravaDaSessao(sessao);
        await trava.WaitAsync().ConfigureAwait(false);
        try
        {
            var indice = Directory.EnumerateFiles(pasta).Count() + 1;
            while (Directory.EnumerateFiles(pasta, "image" + indice + ".*").Any()) indice++;
            var nome = "image" + indice + "." + ext;
            await File.WriteAllBytesAsync(Path.Combine(pasta, nome), bytes).ConfigureAwait(false);
            return nome;
        }
        finally { trava.Release(); }
    }

    /// <summary>Extensao pelo CONTEUDO (assinatura dos primeiros bytes), no mesmo conjunto que o app oficial
    /// copia sem reconverter: png, jpg, gif, svg — mais bmp, que o proprio sdkjs exibe. <c>declarada</c> e'
    /// usada so' para o SVG, que e' texto e nao tem assinatura binaria propria. <c>null</c> = nao e' imagem.
    ///
    /// SVG so' e' aceito quando <paramref name="declarada"/> diz "svg" — ou seja, um arquivo ESCOLHIDO PELO
    /// USUARIO num dialogo do host (<see cref="PonteDesktop.PermitirArquivoEscolhido"/>), nunca uma `data:`
    /// URL vinda de colar/arrastar (essa chega aqui com <c>declarada = null</c>, entao nunca bate no `EhSvg`
    /// abaixo). Recusa DE PROPOSITO (minor 3 da revisao 1), nao "formato desconhecido": SVG e' XML e pode
    /// trazer um <c>&lt;script&gt;</c> valido dentro — o editor exibiria isso como midia do documento sem
    /// nenhuma sandbox extra por cima. <see cref="PonteDesktop.Resolver"/> intercepta `data:image/svg+xml`
    /// ANTES de chegar aqui e devolve um erro proprio (<c>"formato-nao-suportado"</c>, nao o <c>"error"</c>
    /// generico), exatamente para essa recusa ficar visivel em vez de se confundir com "conteudo ilegivel" —
    /// e, desde o fix 2 da revisao da Task 5, fareja tambem o CONTEUDO decodificado de qualquer <c>data:</c>
    /// (<see cref="PareceXml"/>), para um MIME mentido (<c>image/png</c> com texto SVG) cair na mesma recusa.</summary>
    public static string? ExtensaoDaImagem(byte[] bytes, string? declarada = null)
    {
        if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 'P' && bytes[2] == 'N' && bytes[3] == 'G') return "png";
        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF) return "jpg";
        if (bytes.Length >= 6 && bytes[0] == 'G' && bytes[1] == 'I' && bytes[2] == 'F' && bytes[3] == '8') return "gif";
        if (bytes.Length >= 2 && bytes[0] == 'B' && bytes[1] == 'M') return "bmp";
        if (EhSvg(bytes, declarada)) return "svg";
        return null;
    }

    private static bool EhSvg(byte[] bytes, string? declarada)
    {
        if (!string.Equals(declarada?.TrimStart('.'), "svg", StringComparison.OrdinalIgnoreCase)) return false;
        var inicio = Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, 512));
        return inicio.Contains("<svg", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>O conteudo PARECE texto XML/HTML/SVG: o primeiro byte util (pulando o BOM UTF-8 e espacos em
    /// branco) e' <c>&lt;</c>. E' o farejo que <see cref="PonteDesktop"/> aplica a QUALQUER <c>data:</c> depois
    /// de decodificar o base64 (fix 2 da revisao da Task 5): o MIME declarado na URL vem da pagina, entao
    /// <c>data:image/png;base64,&lt;svg…&gt;</c> nao pode escapar da recusa proposital de SVG so' por mentir o
    /// tipo. Nenhuma imagem binaria aceita (png/jpg/gif/bmp) comeca com <c>&lt;</c> — e um poliglota com
    /// magico PNG na frente e' tratado como png, gravado e servido como <c>image/png</c>, o que o navegador
    /// decodifica como bitmap sem executar nada.</summary>
    public static bool PareceXml(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        var i = 0;
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) i = 3;
        while (i < bytes.Length && (bytes[i] == ' ' || bytes[i] == '\t' || bytes[i] == '\r' || bytes[i] == '\n')) i++;
        return i < bytes.Length && bytes[i] == '<';
    }
}
