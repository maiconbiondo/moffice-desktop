namespace mOffice.Conversor;

/// <summary>Uma extensao que o mOffice conhece. <see cref="Abre"/>: o app a abre. <see cref="Grava"/>: aparece no
/// "Salvar como" e o x2t a grava (medido na auditoria de 2026-10-01). <see cref="Legado"/>: o x2t le' mas nao grava
/// (.doc/.xls/.ppt: 80/88) — abre em Modo de compatibilidade e salva como o moderno do tipo. <see cref="Modelo"/>:
/// pelo Explorer abre como documento NOVO. <see cref="Apresentar"/>: pelo Explorer abre ja' apresentando.
/// <see cref="Familia"/>: o que o <see cref="ValidadorDePacote"/> confere ("ooxml", "odf", "pdf", "rtf", "texto",
/// "binario"). <see cref="ProgId"/>: associacao no Explorer (null = nenhuma).</summary>
public sealed record InfoExtensao(string Extensao, FormatoDocumento Formato, string DocType, string Descricao,
    bool Abre, bool Grava, bool Legado, bool Modelo, bool Macro, bool Apresentar, bool PerdeFormatacao,
    string Familia, string? ProgId, bool PadraoNoRmm);

/// <summary>
/// A UNICA lista de formatos do produto (Plano 6, Task 2): dela saem o gate de abertura
/// (<see cref="FormatoDocumentoExtensoes.DaExtensao"/>), os filtros dos dialogos (PerfilApp), as linhas do "Salvar
/// como" (JanelaEditor.LinhasSalvar) e as associacoes do instalador (GeradorAssociacoes, nos testes). A ORDEM importa:
/// a primeira linha de cada <see cref="FormatoDocumento"/> e' a principal (extensao e descricao dele).
/// </summary>
public static class TabelaFormatos
{
    private static InfoExtensao L(string ext, FormatoDocumento f, string tipo, string desc, bool abre = true, bool grava = true,
        bool legado = false, bool modelo = false, bool macro = false, bool apresentar = false, bool perde = false,
        string familia = "ooxml", bool rmm = false, bool semProgId = false) =>
        new(ext, f, tipo, desc, abre, grava, legado, modelo, macro, apresentar, perde, familia,
            abre && !semProgId ? $"mOffice.{Prefixo(tipo)}.{ext.TrimStart('.')}" : null, rmm);

    private static string Prefixo(string tipo) => tipo switch { "cell" => "Planilha", "slide" => "Slides", _ => "Texto" };

    public static IReadOnlyList<InfoExtensao> Todas { get; } =
    [
        // --- texto
        L(".docx", FormatoDocumento.Docx, "word", "Documento do Word", rmm: true),
        L(".docm", FormatoDocumento.Docm, "word", "Documento Habilitado para Macro do Word", macro: true, rmm: true),
        L(".dotx", FormatoDocumento.Dotx, "word", "Modelo do Word", modelo: true, rmm: true),
        L(".dotm", FormatoDocumento.Dotm, "word", "Modelo Habilitado para Macro do Word", modelo: true, macro: true),
        L(".doc", FormatoDocumento.Doc, "word", "Documento do Word 97-2003", grava: false, legado: true, familia: "binario", rmm: true),
        L(".dot", FormatoDocumento.Doc, "word", "Modelo do Word 97-2003", grava: false, legado: true, modelo: true, familia: "binario"),
        L(".odt", FormatoDocumento.Odt, "word", "Documento de texto ODF", familia: "odf", rmm: true),
        L(".ott", FormatoDocumento.Ott, "word", "Modelo de texto ODF", modelo: true, familia: "odf"),
        L(".rtf", FormatoDocumento.Rtf, "word", "Formato Rich Text", familia: "rtf", rmm: true),
        // .txt abre e grava (Task 3), mas SEM ProgID: o Bloco de Notas continua o dono do .txt no Windows.
        L(".txt", FormatoDocumento.Txt, "word", "Texto sem formatação", perde: true, familia: "texto", semProgId: true),
        L(".html", FormatoDocumento.Html, "word", "Página da Web", abre: false, perde: true, familia: "texto"),
        L(".htm", FormatoDocumento.Html, "word", "Página da Web", abre: false, perde: true, familia: "texto"),   // alias: so' vale digitado (o Salvar como lista uma linha por formato)
        // --- planilha
        L(".xlsx", FormatoDocumento.Xlsx, "cell", "Pasta de trabalho do Excel", rmm: true),
        L(".xlsm", FormatoDocumento.Xlsm, "cell", "Pasta de trabalho habilitada para macro do Excel", macro: true, rmm: true),
        L(".xltx", FormatoDocumento.Xltx, "cell", "Modelo do Excel", modelo: true, rmm: true),
        L(".xltm", FormatoDocumento.Xltm, "cell", "Modelo habilitado para macro do Excel", modelo: true, macro: true),
        L(".xls", FormatoDocumento.Xls, "cell", "Pasta de trabalho do Excel 97-2003", grava: false, legado: true, familia: "binario", rmm: true),
        L(".xlt", FormatoDocumento.Xls, "cell", "Modelo do Excel 97-2003", grava: false, legado: true, modelo: true, familia: "binario"),
        L(".xlsb", FormatoDocumento.Xlsb, "cell", "Pasta de trabalho binária do Excel", grava: false, legado: true, familia: "binario"),
        L(".ods", FormatoDocumento.Ods, "cell", "Planilha ODF", familia: "odf", rmm: true),
        L(".ots", FormatoDocumento.Ots, "cell", "Modelo de planilha ODF", modelo: true, familia: "odf"),
        // .csv abre e grava (Task 3). Ganha o ProgID (aparece no "Abrir com"), mas nunca vira o padrao: o Excel e' o dono do .csv.
        L(".csv", FormatoDocumento.Csv, "cell", "CSV (separado por ponto e vírgula)", perde: true, familia: "texto"),
        // --- apresentacao
        L(".pptx", FormatoDocumento.Pptx, "slide", "Apresentação do PowerPoint", rmm: true),
        L(".pptm", FormatoDocumento.Pptm, "slide", "Apresentação habilitada para macro do PowerPoint", macro: true, rmm: true),
        L(".potx", FormatoDocumento.Potx, "slide", "Modelo do PowerPoint", modelo: true, rmm: true),
        L(".potm", FormatoDocumento.Potm, "slide", "Modelo habilitado para macro do PowerPoint", modelo: true, macro: true),
        L(".ppsx", FormatoDocumento.Ppsx, "slide", "Apresentação de slides do PowerPoint", apresentar: true, rmm: true),
        L(".ppsm", FormatoDocumento.Ppsm, "slide", "Apresentação de slides habilitada para macro do PowerPoint", macro: true, apresentar: true),
        L(".ppt", FormatoDocumento.Ppt, "slide", "Apresentação do PowerPoint 97-2003", grava: false, legado: true, familia: "binario", rmm: true),
        L(".pps", FormatoDocumento.Ppt, "slide", "Apresentação de slides do PowerPoint 97-2003", grava: false, legado: true, apresentar: true, familia: "binario"),
        L(".pot", FormatoDocumento.Ppt, "slide", "Modelo do PowerPoint 97-2003", grava: false, legado: true, modelo: true, familia: "binario"),
        L(".odp", FormatoDocumento.Odp, "slide", "Apresentação ODF", familia: "odf", rmm: true),
        L(".otp", FormatoDocumento.Otp, "slide", "Modelo de apresentação ODF", modelo: true, familia: "odf"),
        // --- exportacao (nenhum app abre pela tabela; o mTexto converte PDF em documento pelo Abrir/soltar: PerfilApp.AbrePdfComoDocumento)
        L(".pdf", FormatoDocumento.Pdf, "pdf", "PDF", abre: false, familia: "pdf"),
    ];

    public static InfoExtensao? DaExtensao(string caminhoOuExtensao)
    {
        var ext = caminhoOuExtensao.StartsWith('.') ? caminhoOuExtensao : Path.GetExtension(caminhoOuExtensao);
        return Todas.FirstOrDefault(i => string.Equals(i.Extensao, ext, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>A primeira linha do formato (PDF/A usa a do PDF).</summary>
    public static InfoExtensao Principal(FormatoDocumento f)
    {
        var alvo = f == FormatoDocumento.PdfA ? FormatoDocumento.Pdf : f;
        return Todas.First(i => i.Formato == alvo);
    }

    public static FormatoDocumento Moderno(string docType) => docType switch
    {
        "cell" => FormatoDocumento.Xlsx, "slide" => FormatoDocumento.Pptx, _ => FormatoDocumento.Docx,
    };

    public static IEnumerable<InfoExtensao> DoTipo(string docType) => Todas.Where(i => i.DocType == docType);

    /// <summary>Em que formato uma sessao aberta desta extensao grava: o proprio, ou o moderno do tipo quando e' legado.</summary>
    public static FormatoDocumento GravaComo(InfoExtensao i) => i.Grava ? i.Formato : Moderno(i.DocType);
}
