using System.Text;

namespace mOffice.Conversor;

/// <summary>
/// A sugestao do host para o dialogo de importacao de CSV/TXT (Plano 6, Task 3): codificacao pelo conteudo e delimitador pelas
/// primeiras linhas. So' sugere — quem decide e' o usuario, no dialogo do editor. Nunca lanca: arquivo ilegivel cai no padrao pt-BR.
/// </summary>
public static class SugestaoDeTexto
{
    /// <summary>Quanto do comeco do arquivo e' examinado.</summary>
    public const int BytesExaminados = 64 * 1024;

    private const int LinhasExaminadas = 20;

    public static OpcoesTexto Detectar(string arquivo, bool csv)
    {
        byte[] b;
        try
        {
            using var f = File.OpenRead(arquivo);
            b = new byte[(int)Math.Min(BytesExaminados, f.Length)];
            f.ReadExactly(b);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return csv ? OpcoesTexto.CsvExcelPtBr : OpcoesTexto.TxtUtf8;
        }
        var utf8 = EhUtf8(b);
        var codificacao = utf8 ? OpcoesTexto.Utf8 : OpcoesTexto.Windows1252;
        if (!csv) return new OpcoesTexto(codificacao, OpcoesTexto.SemDelimitador);
        var texto = utf8 ? Encoding.UTF8.GetString(b) : Encoding.Latin1.GetString(b);
        return new OpcoesTexto(codificacao, DelimitadorMaisFrequente(texto));
    }

    /// <summary>BOM de UTF-8, ou bytes que decodificam como UTF-8 sem erro. Uma sequencia incompleta no FIM (o corte dos
    /// <see cref="BytesExaminados"/> no meio de um caractere) nao conta como erro: o decodificador a segura (<c>flush: false</c>).</summary>
    private static bool EhUtf8(byte[] b)
    {
        if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return true;
        try { new UTF8Encoding(false, throwOnInvalidBytes: true).GetDecoder().GetCharCount(b, 0, b.Length, flush: false); return true; }
        catch (DecoderFallbackException) { return false; }
    }

    /// <summary>';', ',' ou tabulacao: o que mais aparece FORA de aspas nas primeiras linhas. Nenhum ou empate com ';' = ';' (pt-BR).</summary>
    private static int DelimitadorMaisFrequente(string texto)
    {
        int pontoEVirgula = 0, virgula = 0, tab = 0, linhas = 0;
        var dentroDeAspas = false;
        foreach (var c in texto)
        {
            if (c == '"') { dentroDeAspas = !dentroDeAspas; continue; }
            if (dentroDeAspas) continue;
            switch (c)
            {
                case ';': pontoEVirgula++; break;
                case ',': virgula++; break;
                case '\t': tab++; break;
                case '\n': linhas++; break;
            }
            if (linhas >= LinhasExaminadas) break;
        }
        var maior = Math.Max(pontoEVirgula, Math.Max(virgula, tab));
        if (maior == 0 || pontoEVirgula == maior) return OpcoesTexto.PontoEVirgula;
        return virgula == maior ? OpcoesTexto.Virgula : OpcoesTexto.Tabulacao;
    }
}
