using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>Uma linha do filtro do "Salvar como". <see cref="Extensao"/> com ponto.</summary>
internal sealed record LinhaSalvar(string Rotulo, string Extensao, FormatoDocumento Formato, bool Assinar = false, OpcoesTexto? Texto = null);

/// <summary>
/// Plano 6, Task 2 — as linhas do "Salvar como" (e a traducao do que o usuario escolheu em <see cref="DestinoSalvar"/>) saem da
/// <see cref="TabelaFormatos"/>. Antes moravam no <c>JanelaEditor.xaml.cs</c> como texto de filtro escrito a mao.
/// </summary>
public partial class JanelaEditor
{
    /// <summary>
    /// Linhas do "Salvar como" para o tipo do documento. As CINCO primeiras ficam onde estavam (OOXML, ODF, PDF, PDF/A, PDF
    /// assinado com mPDF — Plano 4): o painel "Salvar como" do editor pre-seleciona 3/4 pelo fileType e os testes do
    /// Plano 4 dependem disso. Depois vem o resto do que a tabela GRAVA para o tipo, na ordem da tabela, sem legado.
    /// Sem exportacao (caminho do FECHAMENTO), as tres linhas de PDF saem: ali o usuario pediu "Salvar" para NAO perder o
    /// documento, e PDF nao salva documento nenhum — e' exportacao, a sessao continua no original e a janela fecharia
    /// apagando a recuperacao (achado Critico da revisao final do Plano 2).
    ///
    /// Formato de exportacao com <paramref name="permitirExportacao"/> falso (uma sessao cujo formato e' PDF, que este editor
    /// nem abre) nao tem par editavel: cai nas linhas de TEXTO, o unico destino que ainda preserva o que o usuario
    /// escreveu. Nunca devolve lista vazia — um dialogo sem filtro nenhum seria pior do que um destino inesperado.
    /// </summary>
    internal static IReadOnlyList<LinhaSalvar> LinhasSalvar(FormatoDocumento documento, bool permitirExportacao)
    {
        var tipo = documento.DocType();
        // Sessao que ja' e' PDF (este editor nem abre PDF como sessao): com exportacao, so' PDF; sem exportacao, cai nas linhas de
        // TEXTO — nunca um filtro vazio (Pdf_SemExportacao_CaiNoFiltroDeTextoENuncaFicaVazio, Plano 2).
        if (tipo == "pdf") return permitirExportacao ? [new("PDF (*.pdf)", ".pdf", FormatoDocumento.Pdf)] : LinhasSalvar(FormatoDocumento.Docx, false);
        // Plano 6, Task 3: o CSV vira DUAS linhas (ANSI para o Excel pt-BR e UTF-8) e o TXT leva as opcoes de texto; as demais nao mudam.
        IEnumerable<LinhaSalvar> Da(InfoExtensao i) => i.Formato switch
        {
            FormatoDocumento.Csv => [new("CSV (separado por ponto e vírgula) (*.csv)", ".csv", FormatoDocumento.Csv, Texto: OpcoesTexto.CsvExcelPtBr),
                                     new("CSV UTF-8 (separado por ponto e vírgula) (*.csv)", ".csv", FormatoDocumento.Csv, Texto: OpcoesTexto.CsvUtf8)],
            FormatoDocumento.Txt => [new($"{i.Descricao} (*{i.Extensao})", i.Extensao, i.Formato, Texto: OpcoesTexto.TxtUtf8)],
            _ => [new($"{i.Descricao} (*{i.Extensao})", i.Extensao, i.Formato)],
        };
        // DistinctBy: uma linha por FORMATO (".htm" e' alias de ".html": so' vale digitado, ver DestinoDoDialogo).
        var doTipo = TabelaFormatos.DoTipo(tipo).Where(i => i.Grava && !i.Legado).DistinctBy(i => i.Formato).ToList();
        var moderno = doTipo.First(i => i.Familia == "ooxml");
        var odf = doTipo.First(i => i.Familia == "odf");
        var linhas = new List<LinhaSalvar>(Da(moderno).Concat(Da(odf)));
        if (permitirExportacao)
            linhas.AddRange([
                new("PDF (*.pdf)", ".pdf", FormatoDocumento.Pdf),
                new("PDF/A (*.pdf)", ".pdf", FormatoDocumento.PdfA),
                new("PDF assinado com mPDF (*.pdf)", ".pdf", FormatoDocumento.Pdf, Assinar: true)]);
        linhas.AddRange(doTipo.Where(i => i != moderno && i != odf).SelectMany(Da));
        return linhas;
    }

    /// <summary>Filtro do <c>SaveFileDialog</c> ("rotulo|*.ext|…") das <see cref="LinhasSalvar"/>. `internal` para o teste puro
    /// poder compara-lo sem abrir janela.</summary>
    internal static string FiltroSalvar(FormatoDocumento f, bool permitirExportacao = true) =>
        string.Join("|", LinhasSalvar(f, permitirExportacao).Select(l => $"{l.Rotulo}|*{l.Extensao}"));

    /// <summary>Posicao (base 1, a do <c>SaveFileDialog.FilterIndex</c>) do formato dentro de <see cref="LinhasSalvar"/> (com
    /// exportacao): 1 = OOXML, 2 = ODF, 3 = PDF, 4 = PDF/A, 5 = PDF com <paramref name="assinar"/>, depois as demais. Formato fora
    /// da lista (legado, TXT...) = 1. Numa sessao que ja' e' PDF (filtro de uma linha so', que este editor nem abre) o indice de
    /// PDF vira o da lista de texto e o dialogo cai na primeira linha.</summary>
    internal static int IndiceDoFiltro(FormatoDocumento f, bool assinar = false, bool permitirExportacao = true) =>
        IndiceNaLista(LinhasSalvar(f is FormatoDocumento.Pdf or FormatoDocumento.PdfA ? FormatoDocumento.Docx : f, permitirExportacao), f, assinar);

    /// <summary>Posicao (base 1) de <paramref name="f"/> NA LISTA que o dialogo mostra; fora dela = 1.</summary>
    internal static int IndiceNaLista(IReadOnlyList<LinhaSalvar> linhas, FormatoDocumento f, bool assinar = false, OpcoesTexto? atual = null)
    {
        var lista = linhas.ToList();
        // CSV/TXT: a linha com a MESMA codificacao da sessao (um CSV UTF-8 abre o dialogo em "CSV UTF-8", nao em ANSI: aceitar o sugerido
        // perderia os caracteres fora do ANSI em silencio); sem correspondencia, a primeira linha do formato.
        var i = atual is null ? -1 : lista.FindIndex(l => l.Formato == f && l.Assinar == assinar && l.Texto?.Codificacao == atual.Codificacao);
        if (i < 0) i = lista.FindIndex(l => l.Formato == f && l.Assinar == assinar);
        return i < 0 ? 1 : i + 1;
    }

    /// <summary>Filtro e linha pre-selecionada do dialogo, derivados de UMA so' chamada a <see cref="LinhasSalvar"/>: nunca divergem
    /// (o fechamento mostra a lista SEM as 3 linhas de PDF; contar com elas deslocava a selecao em 3).</summary>
    internal static (string Filtro, int Indice) FiltroEIndice(FormatoDocumento documento, FormatoDocumento sugerido, bool permitirExportacao, OpcoesTexto? atual = null)
    {
        var linhas = LinhasSalvar(documento, permitirExportacao);
        return (string.Join("|", linhas.Select(l => $"{l.Rotulo}|*{l.Extensao}")), IndiceNaLista(linhas, sugerido, atual: atual));
    }

    /// <summary>
    /// Traduz o que o usuario escolheu no dialogo (<paramref name="caminho"/> + <c>FilterIndex</c> base 1) no
    /// <see cref="DestinoSalvar"/>. Puro (`internal static`, testado sem janela). O formato vem da LINHA escolhida, porque as tres
    /// linhas de PDF tem a mesma extensao. Regras:
    /// <list type="bullet">
    /// <item>Linhas de PDF (so' com exportacao): uma extensao de DOCUMENTO do mesmo tipo e gravavel digitada vence (o usuario
    /// escreveu "a.odt" com a linha PDF escolhida: grava ODT; na linha "PDF assinado" o <c>Assinar</c> vai junto, para o host avisar
    /// que nada foi enviado para assinar — <see cref="MensagemAssinarSalvouDocumento"/>); senao o nome termina em <c>.pdf</c>
    /// (acrescentado se faltar — nunca grava PDF com outra extensao).</item>
    /// <item>Demais linhas: extensao digitada que seja GRAVAVEL e do mesmo tipo vence (o usuario digitou "a.rtf" com a linha
    /// .docx: grava RTF); senao o formato e' o da linha e a extensao e' corrigida (<see cref="ComExtensaoDe"/>: legado trocado,
    /// desconhecida acrescentada — A2).</item>
    /// </list>
    /// </summary>
    internal static DestinoSalvar DestinoDoDialogo(string caminho, int filterIndex, FormatoDocumento sugerido, bool permitirExportacao)
    {
        var linhas = LinhasSalvar(sugerido, permitirExportacao);
        var linha = filterIndex >= 1 && filterIndex <= linhas.Count ? linhas[filterIndex - 1] : linhas[0];
        var digitada = TabelaFormatos.DaExtensao(caminho);
        FormatoDocumento? docDigitado = digitada is { Grava: true, Legado: false } d && d.DocType == sugerido.DocType() ? d.Formato : null;
        if (linha.Extensao == ".pdf" && !sugerido.EhExportacao())
        {
            if (docDigitado is { } doc) return new DestinoSalvar(caminho, doc, Assinar: linha.Assinar, Texto: OpcoesTexto.PadraoDe(doc));
            var pdf = string.Equals(Path.GetExtension(caminho), ".pdf", StringComparison.OrdinalIgnoreCase) ? caminho : caminho + ".pdf";
            return linha.Formato == FormatoDocumento.PdfA ? DestinoPdfA(pdf) : new DestinoSalvar(pdf, FormatoDocumento.Pdf, Assinar: linha.Assinar);
        }
        // CSV/TXT digitado com a linha de OUTRO formato escolhida: sem as opcoes da linha, o padrao do formato (nunca sem opcoes).
        if (docDigitado is { } conhecido) return new DestinoSalvar(caminho, conhecido, Texto: conhecido == linha.Formato ? linha.Texto : OpcoesTexto.PadraoDe(conhecido));
        return new DestinoSalvar(ComExtensaoDe(caminho, linha.Formato), linha.Formato, Texto: linha.Texto);
    }

    /// <summary>
    /// Como a linha "PDF/A" do "Salvar como" gera o arquivo — o UNICO lugar dessa decisao (S6 do Plano 4).
    /// Hoje e' a decisao <b>A</b>: o proprio x2t exporta PDF/A (<see cref="FormatoDocumento.PdfA"/>), pelo mesmo
    /// caminho de exportacao do PDF (sessao nao repontada, <see cref="Exportado"/>, sem senha). O pedido real ao x2t e'
    /// <c>m_nFormatTo = 513</c> + <c>m_bIsPDFA = true</c> (<see cref="X2tConversor"/>): 521 a partir do Editor.bin sai com
    /// codigo 88 (MEDIDO). O arquivo DECLARA PDF/A-2A no XMP (<c>pdfaid:part 2</c>, <c>conformance A</c>); se ele e'
    /// conforme e' a Task 1/S6 que valida, no veraPDF, contra esse perfil declarado. Se reprovar (decisao <b>B</b>), a
    /// troca e' SO' aqui: devolver
    /// <c>FormatoDocumento.Pdf</c> com uma marca de "converter pelo mPDF (/pdfa)" no <see cref="DestinoSalvar"/>.
    /// </summary>
    private static DestinoSalvar DestinoPdfA(string pdf) => new(pdf, FormatoDocumento.PdfA);

    /// <summary>Nome (sem extensao) que o dialogo "Salvar como" sugere: o titulo generico de um documento novo; o
    /// nome do original com " (recuperado)" numa sessao recuperada desvinculada (I2 da revisao final do Plano 3 —
    /// aceitar o nome sugerido nunca sobrescreve o original alterado); senao o nome do original (num legado, o dialogo poe a
    /// extensao moderna da linha: "contrato.doc" -&gt; "contrato.docx"). `internal static` para o teste puro.</summary>
    internal static string NomeSugerido(SessaoDocumento s) =>
        s.Novo ? s.Titulo
        : s.Desvinculada ? Path.GetFileNameWithoutExtension(s.ArquivoOrigem) + " (recuperado)"
        : Path.GetFileNameWithoutExtension(s.ArquivoOrigem);

    /// <summary>Extensoes que NUNCA recebem conteudo (A2, Task 1): as da tabela que o mOffice abre mas o x2t nao grava
    /// (.doc/.xls/.ppt e variantes: 80/88, medido; .xlsb relido perde os textos).</summary>
    internal static IReadOnlySet<string> ExtensoesSoLeitura { get; } =
        new HashSet<string>(TabelaFormatos.Todas.Where(i => !i.Grava && i.Abre).Select(i => i.Extensao), StringComparer.OrdinalIgnoreCase);

    /// <summary>A2: extensao de formato so' de leitura e' TROCADA ("a.doc" -&gt; "a.docx"); qualquer outra fica no nome e a
    /// do formato e' acrescentada ("Contrato v1.2" -&gt; "Contrato v1.2.docx").</summary>
    internal static string ComExtensaoDe(string caminho, FormatoDocumento f) =>
        ExtensoesSoLeitura.Contains(Path.GetExtension(caminho))
            ? Path.ChangeExtension(caminho, f.Extensao())
            : caminho + "." + f.Extensao();

    /// <summary>Defesa em profundidade (A2): o destino tem uma extensao que a tabela GRAVA para o formato. Um destino
    /// "x.doc" nunca combina com conteudo moderno (a extensao legada nao esta' entre as que gravam).</summary>
    internal static bool ExtensaoCombina(string destino, FormatoDocumento f) =>
        TabelaFormatos.Todas.Any(i => i.Formato == (f == FormatoDocumento.PdfA ? FormatoDocumento.Pdf : f) && i.Grava
            && string.Equals(i.Extensao, Path.GetExtension(destino), StringComparison.OrdinalIgnoreCase));

    /// <summary>Faixa a CADA gravacao num formato que nao guarda a formatacao (TXT, CSV, HTML; <c>PerdeFormatacao</c> da tabela).</summary>
    public static string MensagemPerdeFormatacao(FormatoDocumento f) =>
        $"O arquivo foi salvo como {f.Extensao().ToUpperInvariant()}: formatação, imagens e tabelas não são guardadas nesse formato.";

    /// <summary>Faixa a cada gravacao em CSV: o x2t grava so' a planilha ativa (a primeira), como o Excel.</summary>
    public const string MensagemCsvSoPlanilhaAtiva = "O arquivo foi salvo como CSV: só a planilha ativa foi gravada.";

    public static string MensagemExtensaoNaoCombina(string destino, FormatoDocumento f) =>
        $"O arquivo \"{Path.GetFileName(destino)}\" não foi gravado: o conteúdo seria {f.Extensao().ToUpperInvariant()} com a extensão "
        + $"{Path.GetExtension(destino).ToLowerInvariant()}. Use Salvar como e escolha o tipo pela lista.";
}
