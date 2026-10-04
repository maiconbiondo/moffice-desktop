namespace mOffice.Conversor;

/// <summary>
/// Codigos <c>c_oAscFileType</c> do sdkjs 9.4 (os mesmos AVS_OFFICESTUDIO_FILE_* do x2t, core/Common/OfficeFileFormats.h). Plano 6,
/// Task 2: alem dos seis de antes, o legado (.doc/.xls/.ppt: so' leitura), modelos, macros, apresentacao de slides, RTF e HTML.
/// A lista de EXTENSOES de cada formato mora em <see cref="TabelaFormatos"/>.
/// </summary>
public enum FormatoDocumento
{
    Docx = 65, Doc = 66, Odt = 67, Rtf = 68, Txt = 69, Html = 70, Docm = 75, Dotx = 76, Dotm = 77, Ott = 79,
    Pptx = 129, Ppt = 130, Odp = 131, Ppsx = 132, Pptm = 133, Ppsm = 134, Potx = 135, Potm = 136, Otp = 138,
    Xlsx = 257, Xls = 258, Ods = 259, Csv = 260, Xlsm = 261, Xltx = 262, Xltm = 263, Xlsb = 264, Ots = 266,
    Pdf = 513, PdfA = 521,
}

public static class FormatoDocumentoExtensoes
{
    /// <summary>AVS_OFFICESTUDIO_FILE_CANVAS (8192): e o que o app oficial passa em m_nFormatTo ao gerar o Editor.bin; o x2t escolhe word/cell/slide.</summary>
    public const int CanvasGenerico = 8192;

    public static FormatoDocumento DaExtensao(string caminho) =>
        TabelaFormatos.DaExtensao(caminho) is { } i ? i.Formato
        : throw new ErroConversao($"O mOffice não abre arquivos {Path.GetExtension(caminho).ToLowerInvariant()}.", 0, "");

    /// <summary>Extensao SEM ponto do formato (a da primeira linha dele na tabela; PDF/A usa a do PDF).</summary>
    public static string Extensao(this FormatoDocumento f) => TabelaFormatos.Principal(f).Extensao.TrimStart('.');

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

    /// <summary>Valor de ?doctype= da pagina de entrada do desktop (word|cell|slide); "pdf" para a exportacao.</summary>
    public static string DocType(this FormatoDocumento f) => TabelaFormatos.Principal(f).DocType;
}
