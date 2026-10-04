namespace mOffice.Nucleo;

public enum AcaoDeLink { Navegador, Arquivo, Descartar, Recusar }

/// <summary>O que fazer com um window.open/navegacao que o editor pede (A1/CE-2/MvP-02, auditoria 2026-10-01). Navegador =
/// http(s)/mailto, abre no navegador PADRAO depois de o usuario confirmar. Arquivo = file: de documento/pasta, abre pelo
/// Windows depois de confirmar (so' local; pasta so' se existir). Descartar = enderecos do proprio ONLYOFFICE (ajuda/suporte/feedback/loja) e about: — so' log.
/// Recusar = o resto (executaveis, scripts, esquemas desconhecidos) — log e aviso.</summary>
public static class PoliticaDeLink
{
    private static readonly IReadOnlySet<string> ArquivosPermitidos = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        ".pdf", ".docx", ".docm", ".doc", ".dotx", ".odt", ".rtf", ".txt", ".xlsx", ".xlsm", ".xls", ".ods", ".csv",
        ".pptx", ".pptm", ".ppt", ".ppsx", ".odp", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".msg", ".eml", ".zip",
    };

    public static AcaoDeLink Classificar(string? uri)
    {
        if (string.IsNullOrWhiteSpace(uri) || !Uri.TryCreate(uri, UriKind.Absolute, out var u)) return AcaoDeLink.Recusar;
        switch (u.Scheme.ToLowerInvariant())
        {
            case "about": return AcaoDeLink.Descartar;
            case "mailto": return AcaoDeLink.Navegador;
            case "http":
            case "https":
                var h = u.Host.ToLowerInvariant().TrimEnd((char)46); // "support.onlyoffice.com." (ponto final) e o mesmo host
                return h == "onlyoffice.com" || h.EndsWith(".onlyoffice.com", StringComparison.Ordinal) ? AcaoDeLink.Descartar : AcaoDeLink.Navegador;
            case "file":
                // So' arquivo LOCAL: file://servidor/... (UNC) e' recusado (abrir tocaria a rede e entregaria credenciais NTLM).
                if (!u.IsLoopback && u.Host.Length > 0) return AcaoDeLink.Recusar;
                // Sem extensao: so' se for uma PASTA que existe (abre no Explorer); nunca se manda ao shell um arquivo sem tipo.
                if (Directory.Exists(u.LocalPath)) return AcaoDeLink.Arquivo; // pasta existente (mesmo com ponto no nome)
                var ext = Path.GetExtension(u.LocalPath);
                if (ext.Length == 0) return AcaoDeLink.Recusar;
                return ArquivosPermitidos.Contains(ext) ? AcaoDeLink.Arquivo : AcaoDeLink.Recusar;
            default: return AcaoDeLink.Recusar;
        }
    }
}
