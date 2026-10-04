namespace mOffice.Conversor;

/// <summary>
/// Como ler/gravar um arquivo de texto delimitado (CSV) ou texto puro (TXT) pelo x2t (Plano 6, Task 3, FMT-03): a codificacao e o
/// delimitador sao os numeros que o x2t le' em <c>m_nCsvTxtEncoding</c> e <c>m_nCsvDelimiter</c> — os MESMOS que o dialogo de
/// importacao do editor manda em <c>SetAdvancedOptions</c>. <see cref="Lcid"/> vai em <c>m_nLcid</c> (o x2t usa para ler datas e
/// numeros do arquivo; pt-BR por padrao, como o salvar). <see cref="DelimitadorChar"/>: o "Outro" do dialogo (um caractere qualquer, ex.
/// <c>|</c>), que vai em <c>m_nCsvDelimiterChar</c> e vence o codigo (MEDIDO: com <c>|</c> o x2t separa as colunas pelo caractere).
/// </summary>
/// <param name="Codificacao">O ID da codificacao na lista do editor (sdkjs <c>AscCommon.c_oAscEncodings</c>, a mesma que o x2t le' em
/// <c>m_nCsvTxtEncoding</c>) — NAO o codepage do Windows: <see cref="Utf8"/> = 46 (nao 65001) e <see cref="Windows1252"/> = 44 (nao 1252).
/// MEDIDO no x2t 9.4.0.129: 1252 faz o x2t cair (0xC0000005); 3 le' como ISO-8859-4, 37 como ISO-8859-1.</param>
/// <param name="Delimitador">Codigo de <c>Asc.c_oAscCsvDelimiter</c> (conferido no web-apps 9.4: o dialogo do editor usa 1 tabulacao,
/// 2 ponto e virgula, 3 dois pontos, 4 virgula, 5 espaco): <see cref="SemDelimitador"/> (0) no TXT.</param>
public sealed record OpcoesTexto(int Codificacao, int Delimitador, int Lcid = X2tConversor.LcidPadrao, string? DelimitadorChar = null)
{
    public const int Utf8 = 46, Windows1252 = 44;

    /// <summary>O maior ID da lista de codificacoes do editor (51 = UTF-32BE). Um ID fora de 0..51 faz o x2t cair (MEDIDO com 1252).</summary>
    public const int MaiorIdDeCodificacao = 51;
    public const int Outro = -1, SemDelimitador = 0, Tabulacao = 1, PontoEVirgula = 2, DoisPontos = 3, Virgula = 4, Espaco = 5;

    /// <summary>CSV como o Excel pt-BR espera: ANSI (1252) e ponto e virgula. E' a primeira linha de CSV do "Salvar como".</summary>
    public static readonly OpcoesTexto CsvExcelPtBr = new(Windows1252, PontoEVirgula);

    /// <summary>CSV UTF-8 com ponto e virgula (a segunda linha de CSV do "Salvar como").</summary>
    public static readonly OpcoesTexto CsvUtf8 = new(Utf8, PontoEVirgula);

    /// <summary>TXT em UTF-8, sem delimitador.</summary>
    public static readonly OpcoesTexto TxtUtf8 = new(Utf8, SemDelimitador);

    /// <summary>O padrao de gravacao de um formato de texto (<c>null</c> para qualquer outro formato).</summary>
    public static OpcoesTexto? PadraoDe(FormatoDocumento f) => f switch
    {
        FormatoDocumento.Csv => CsvExcelPtBr,
        FormatoDocumento.Txt => TxtUtf8,
        _ => null,
    };
}
