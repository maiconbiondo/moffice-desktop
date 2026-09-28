using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>Parsing do pedido de salvamento que chega do EDITOR (Ctrl+S / botao Salvar / menu Salvar como):
/// ponte-desktop.js (LocalFileSave) empacota <c>{param, fileType, jsonOptions}</c> como JSON e manda pela
/// ponte (ver <see cref="JanelaEditor.SalvarPedidoDoEditorAsync"/>). Extraido de <see cref="JanelaEditor"/> na
/// revisao 2026-09-17 (fix review 1, minor 6) para ser testavel sem janela nenhuma.</summary>
public static class PedidoSalvar
{
    /// <summary>Le' o JSON que ponte-desktop.js monta no LocalFileSave: <c>{param, fileType, jsonOptions}</c>.
    /// JSON malformado (ou qualquer coisa inesperada) devolve o par neutro (salvar normal, sem fileType) em vez
    /// de propagar excecao — o chamador roda dentro de um fire-and-forget que nunca deve derrubar a ponte.</summary>
    public static (bool salvarComo, int fileType) Ler(string param)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(param);
            var p = doc.RootElement.TryGetProperty("param", out var vp) ? vp.GetString() ?? "" : "";
            var ft = doc.RootElement.TryGetProperty("fileType", out var vf) && vf.TryGetInt32(out var n) ? n : 0;
            return (p.Contains("saveas=true", StringComparison.OrdinalIgnoreCase), ft);
        }
        catch { return (false, 0); }
    }

    /// <summary>Mapeia o <c>fileType</c> (codigo AVS_OFFICESTUDIO_FILE_*, ver <see cref="FormatoDocumento"/>)
    /// para o formato correspondente, ou null se o codigo nao for um dos valores conhecidos (0 incluso —
    /// "nenhum formato escolhido pelo chamador").</summary>
    public static FormatoDocumento? FormatoPorCodigo(int fileType) =>
        Enum.IsDefined(typeof(FormatoDocumento), fileType) ? (FormatoDocumento)fileType : null;
}
