using mOffice.Conversor;

namespace mOffice.Comum;

/// <summary>Identidade de UM dos tres apps da suite. E' o unico lugar onde "mTexto abre docx/odt" vira dado:
/// a tela de boas-vindas, o arrastar-e-soltar, o OpenFileDialog e o titulo da janela leem tudo daqui, entao
/// a Task 6 monta cada app passando o perfil — sem ramificacao por app espalhada pela casca.</summary>
public sealed record PerfilApp(string Nome, string NomeExibicao, string Subtitulo, string ChaveLogo,
    FormatoDocumento FormatoNovo, IReadOnlyList<FormatoDocumento> Formatos, string FiltroAbrir)
{
    public static readonly PerfilApp Texto = new("mtexto", "mTexto", "Editor de textos · DOCX e ODT", "Logo.mTexto",
        FormatoDocumento.Docx, [FormatoDocumento.Docx, FormatoDocumento.Odt], "Documentos de texto|*.docx;*.odt|Todos os arquivos|*.*");
    public static readonly PerfilApp Planilha = new("mplanilha", "mPlanilha", "Editor de planilhas · XLSX e ODS", "Logo.mPlanilha",
        FormatoDocumento.Xlsx, [FormatoDocumento.Xlsx, FormatoDocumento.Ods], "Planilhas|*.xlsx;*.ods|Todos os arquivos|*.*");
    public static readonly PerfilApp Slides = new("mslides", "mSlides", "Editor de apresentações · PPTX e ODP", "Logo.mSlides",
        FormatoDocumento.Pptx, [FormatoDocumento.Pptx, FormatoDocumento.Odp], "Apresentações|*.pptx;*.odp|Todos os arquivos|*.*");

    public bool Aceita(string caminho) =>
        Formatos.Any(f => string.Equals(Path.GetExtension(caminho).TrimStart('.'), f.Extensao(), StringComparison.OrdinalIgnoreCase));

    /// <summary>Primeiro caminho aceito (arrastar-e-soltar e linha de comando); null se nenhum.</summary>
    public string? PrimeiroAceito(IEnumerable<string> caminhos) => caminhos.FirstOrDefault(Aceita);
}
