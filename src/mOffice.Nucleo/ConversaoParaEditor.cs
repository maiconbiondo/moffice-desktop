using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>Resultado de <see cref="ConversaoParaEditor.ConverterAsync"/>: <see cref="Pasta"/> = URL da pasta da conversao
/// (<c>https://moffice-doc.local/convertidos/&lt;id&gt;</c>, o 1o argumento de <c>onDocumentCompare</c>); <see cref="Bin"/> = o binario do editor;
/// <see cref="Imagens"/> = <c>media/&lt;nome&gt;</c> -&gt; caminho LOCAL do arquivo de midia que o x2t gravou (quem chama decide como a midia entra
/// no documento aberto — <see cref="Midia.ImportarMidiaConvertidaAsync"/>).</summary>
public sealed record ResultadoConversao(string Pasta, byte[] Bin, IReadOnlyDictionary<string, string> Imagens);

/// <summary>
/// Arquivo escolhido pelo usuario -&gt; binario do editor, para o que o sdkjs pede por <c>AscDesktopEditor.convertFile</c> (Inserir texto do arquivo,
/// vinculo externo, importar XML) e <c>CompareDocumentFile</c>/<c>MergeDocumentFile</c> (Plano 6, Task 8). Converte para
/// <c>&lt;recovery&gt;\convertidos\&lt;id&gt;</c> (servida por <c>https://moffice-doc.local/convertidos/&lt;id&gt;/</c>) com o MESMO x2t e o mesmo XML do
/// abrir (<see cref="X2tConversor.ParaBinAsync"/>: LCID pt-BR etc.), no canvas pedido. O arquivo do usuario e' so' LIDO (o x2t le' a origem; nada e'
/// escrito ao lado dele). A pasta sai em <see cref="Liberar"/> (o <c>close()</c> do objeto que o sdkjs recebe) e, de qualquer jeito, com a pasta de
/// recuperacao quando o documento fecha.
/// </summary>
public sealed class ConversaoParaEditor(X2tConversor conversor)
{
    public const string SubpastaConvertidos = "convertidos";

    public async Task<ResultadoConversao> ConverterAsync(string caminho, int formatoCanvas, string pastaRecovery, string id, CancellationToken ct = default)
    {
        var pasta = PastaDaConversao(pastaRecovery, id);
        var bin = await conversor.ParaBinAsync(caminho, pasta, ct, formatoCanvas: formatoCanvas).ConfigureAwait(false);
        var bytes = await File.ReadAllBytesAsync(bin, ct).ConfigureAwait(false);
        var media = Path.Combine(pasta, Midia.PastaMedia);
        var imagens = Directory.Exists(media)
            ? Directory.EnumerateFiles(media).ToDictionary(f => Midia.PastaMedia + "/" + Path.GetFileName(f), f => f, StringComparer.Ordinal)
            : new Dictionary<string, string>(StringComparer.Ordinal);
        return new ResultadoConversao(Midia.BaseUrlDocumento + "/" + SubpastaConvertidos + "/" + id, bytes, imagens);
    }

    /// <summary>Apaga a pasta da conversao <paramref name="id"/> (melhor esforco).</summary>
    public static void Liberar(string pastaRecovery, string id)
    {
        try { var p = PastaDaConversao(pastaRecovery, id); if (Directory.Exists(p)) Directory.Delete(p, true); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException) { }
    }

    /// <summary><c>&lt;recovery&gt;\convertidos\&lt;id&gt;</c>. O id vem da pagina no <c>close()</c>: so' letras e digitos (nunca vira outro caminho).</summary>
    internal static string PastaDaConversao(string pastaRecovery, string id)
    {
        if (string.IsNullOrEmpty(id) || id.Length > 64 || !id.All(char.IsAsciiLetterOrDigit))
            throw new ArgumentException("id de conversao invalido", nameof(id));
        return Path.Combine(pastaRecovery, SubpastaConvertidos, id);
    }
}
