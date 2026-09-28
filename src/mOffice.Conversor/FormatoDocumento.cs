namespace mOffice.Conversor;

/// <summary>Codigos AVS_OFFICESTUDIO_FILE_* do x2t (core/Common/OfficeFileFormats.h, verificado 2026-09-16).</summary>
public enum FormatoDocumento { Docx = 65, Odt = 67, Xlsx = 257, Ods = 259, Pptx = 129, Odp = 131, Pdf = 513, PdfA = 521 }

public static class FormatoDocumentoExtensoes
{
    /// <summary>AVS_OFFICESTUDIO_FILE_CANVAS (8192): e o que o app oficial passa em m_nFormatTo ao gerar o Editor.bin; o x2t escolhe word/cell/slide.</summary>
    public const int CanvasGenerico = 8192;

    public static FormatoDocumento DaExtensao(string caminho) => Path.GetExtension(caminho).ToLowerInvariant() switch
    {
        ".docx" => FormatoDocumento.Docx, ".odt" => FormatoDocumento.Odt,
        ".xlsx" => FormatoDocumento.Xlsx, ".ods" => FormatoDocumento.Ods,
        ".pptx" => FormatoDocumento.Pptx, ".odp" => FormatoDocumento.Odp,
        ".pdf" => FormatoDocumento.Pdf,
        var e => throw new ErroConversao($"Extensão não suportada: '{e}'. Formatos aceitos: DOCX, ODT, XLSX, ODS, PPTX, ODP, PDF.", 0, "")
    };

    public static string Extensao(this FormatoDocumento f) => f == FormatoDocumento.PdfA ? "pdf" : f.ToString().ToLowerInvariant();

    /// <summary>
    /// Este formato e' uma EXPORTACAO (saida so' de leitura), nao um formato de documento que o editor
    /// reabre para continuar editando. Hoje, PDF e PDF/A.
    ///
    /// A distincao decide o que acontece com a SESSAO depois de gravar: salvar em formato de documento
    /// repoenta a sessao para o arquivo novo (o titulo muda, o proximo Ctrl+S grava la', o documento deixa
    /// de estar modificado); exportar NAO — quem exportou continua editando o .docx, e um "salvar" seguinte
    /// tem de voltar para o .docx, nao regravar o PDF. Um PDF tambem nao entra nos RECENTES do app: o
    /// mTexto nao abre PDF.
    /// </summary>
    public static bool EhExportacao(this FormatoDocumento f) => f is FormatoDocumento.Pdf or FormatoDocumento.PdfA;

    /// <summary>Valor de ?doctype= da pagina de entrada do desktop (word|cell|slide).</summary>
    public static string DocType(this FormatoDocumento f) => f switch
    {
        FormatoDocumento.Docx or FormatoDocumento.Odt => "word",
        FormatoDocumento.Xlsx or FormatoDocumento.Ods => "cell",
        FormatoDocumento.Pptx or FormatoDocumento.Odp => "slide",
        _ => "pdf"
    };
}
