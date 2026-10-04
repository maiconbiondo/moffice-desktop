using mOffice.Conversor;

namespace mOffice.Comum;

/// <summary>Identidade de UM dos tres apps da suite. E' o unico lugar onde "mTexto abre docx/odt" vira dado:
/// a tela de boas-vindas, o arrastar-e-soltar, o OpenFileDialog e o titulo da janela leem tudo daqui, entao
/// a Task 6 monta cada app passando o perfil — sem ramificacao por app espalhada pela casca.
/// Plano 6, Task 2: o que o app abre (<see cref="Formatos"/>, <see cref="FiltroAbrir"/>, <see cref="Aceita"/>) sai da
/// <see cref="TabelaFormatos"/> pelo <see cref="DocType"/> do app, nao de listas escritas aqui.</summary>
public sealed record PerfilApp(string Nome, string NomeExibicao, string Descricao, string ChaveLogo, string DocType,
    FormatoDocumento FormatoNovo)
{
    public static readonly PerfilApp Texto = new("mtexto", "mTexto", "Editor de textos", "Logo.mTexto",
        "word", FormatoDocumento.Docx);
    public static readonly PerfilApp Planilha = new("mplanilha", "mPlanilha", "Editor de planilhas", "Logo.mPlanilha",
        "cell", FormatoDocumento.Xlsx);
    public static readonly PerfilApp Slides = new("mslides", "mSlides", "Editor de apresentações", "Logo.mSlides",
        "slide", FormatoDocumento.Pptx);

    /// <summary>O texto do botao "novo" da boas-vindas: "Documento em branco" / "Pasta de trabalho em branco" / "Apresentação em branco" (como o Office).</summary>
    public string RotuloNovo => DocType switch { "cell" => "Pasta de trabalho em branco", "slide" => "Apresentação em branco", _ => "Documento em branco" };

    /// <summary>"Editor de textos · DOCX, DOC, ODT e RTF": os formatos de documento que o app abre (sem modelo, macro nem apresentacao de
    /// slides), da <see cref="TabelaFormatos"/>.</summary>
    public string Subtitulo
    {
        get
        {
            var e = TabelaFormatos.DoTipo(DocType).Where(i => i.Abre && !i.Modelo && !i.Macro && !i.Apresentar)
                .Select(i => i.Extensao.TrimStart('.').ToUpperInvariant()).ToList();
            return $"{Descricao} · " + (e.Count > 1 ? string.Join(", ", e[..^1]) + " e " + e[^1] : string.Join("", e));
        }
    }

    /// <summary>Formatos que o app abre, sem repeticao (.doc e .dot sao o mesmo <see cref="FormatoDocumento"/>).</summary>
    public IReadOnlyList<FormatoDocumento> Formatos =>
        TabelaFormatos.DoTipo(DocType).Where(i => i.Abre).Select(i => i.Formato).Distinct().ToList();

    /// <summary>"Todos os documentos do mTexto|*.docx;*.docm;…|Documento do Word (*.docx)|*.docx|…|Todos os arquivos|*.*" —
    /// como o "Todos os Documentos do Word" do Office, a partir da tabela.</summary>
    public string FiltroAbrir
    {
        get
        {
            var abre = TabelaFormatos.DoTipo(DocType).Where(i => i.Abre).ToList();
            var todos = string.Join(";", abre.Select(i => "*" + i.Extensao));
            var linhas = abre.Select(i => $"{i.Descricao} (*{i.Extensao})|*{i.Extensao}").ToList();
            // FMT-07: so' o mTexto abre PDF (como documento editavel), e so' pelo Abrir/soltar: fica fora de "Todos os documentos".
            if (AbrePdfComoDocumento(".pdf")) linhas.Add($"{DescricaoDoPdf}|*.pdf");
            return $"Todos os documentos do {NomeExibicao}|{todos}|" + string.Join("|", linhas) + "|Todos os arquivos|*.*";
        }
    }

    /// <summary>A linha do PDF no filtro do Abrir do mTexto.</summary>
    public const string DescricaoDoPdf = "PDF (abrir como documento editável) (*.pdf)";

    /// <summary>FMT-07 (Plano 6, Task 3): este app abre <paramref name="caminho"/> (um .pdf) CONVERTENDO-o em documento editavel — so' o mTexto,
    /// so' pelo Abrir/soltar. <see cref="Aceita"/> continua falso para .pdf: o PDF nao tem ProgID no mTexto (o mPDF e' o dono), nao vem pela
    /// linha de comando nem entra nos recentes.</summary>
    public bool AbrePdfComoDocumento(string caminho) =>
        DocType == "word" && string.Equals(Path.GetExtension(caminho), ".pdf", StringComparison.OrdinalIgnoreCase);

    public bool Aceita(string caminho) => TabelaFormatos.DaExtensao(caminho) is { Abre: true } i && i.DocType == DocType;

    /// <summary>FMT-05: a faixa quando um arquivo que este app nao abre chega (Abrir com, soltar, pipe) — antes sumia em silencio.</summary>
    public string MensagemNaoAbre(string caminho) =>
        $"O {NomeExibicao} não abre arquivos {Path.GetExtension(caminho).ToUpperInvariant()}: abra-o pelo aplicativo do mOffice correspondente.";

    /// <summary>Primeiro caminho aceito (arrastar-e-soltar e linha de comando); null se nenhum.</summary>
    public string? PrimeiroAceito(IEnumerable<string> caminhos) => caminhos.FirstOrDefault(c => Aceita(c) || AbrePdfComoDocumento(c));
}
