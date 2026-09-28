using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace mOffice.Conversor.Tests;

/// <summary>
/// Leitor MINIMO de PDF para os testes (Task 4 do Plano 3): quantas paginas o arquivo tem e o texto que o x2t
/// escreveu nelas. Nao e' um leitor de PDF geral — cobre o que o proprio x2t gera (objetos soltos, streams
/// FlateDecode, fontes Identity-H com <c>/ToUnicode</c>) e o PDF do Chromium (so' a contagem de paginas).
///
/// Por que nao basta procurar a string no arquivo: o x2t grava o texto como indices de glifo em hexadecimal
/// (<c>[&lt;0001000200030002&gt;-249.97 ...]TJ</c>, MEDIDO num PDF de 3 paginas gerado por ele) e o espaco
/// entre palavras vira ajuste de kerning, nao caractere. O texto so' existe depois de passar cada codigo pelo
/// mapa <c>/ToUnicode</c> da fonte ativa (<c>/F1 10 Tf</c>).
/// </summary>
public static class TextoDePdf
{
    private static readonly Regex Objeto = new(@"(\d+)\s+0\s+obj\b(.*?)\bendobj", RegexOptions.Singleline | RegexOptions.Compiled);

    /// <summary>Numero de paginas: objetos <c>/Type /Page</c> (nao <c>/Pages</c>).</summary>
    public static int Paginas(string pdf)
    {
        var texto = Encoding.Latin1.GetString(File.ReadAllBytes(pdf));
        return Regex.Matches(texto, @"/Type\s*/Page(?![a-zA-Z])").Count;
    }

    /// <summary>Texto de todas as paginas, na ordem dos objetos de conteudo (um paragrafo por operador
    /// <c>Td</c>/<c>T*</c>; kerning negativo grande vira espaco).</summary>
    public static string Extrair(string pdf)
    {
        var texto = Encoding.Latin1.GetString(File.ReadAllBytes(pdf));
        var objetos = new Dictionary<int, string>();
        foreach (Match m in Objeto.Matches(texto)) objetos[int.Parse(m.Groups[1].Value)] = m.Groups[2].Value;

        // Nome da fonte no recurso (/F1) -> mapa ToUnicode. Em conflito (o mesmo /F1 apontando para objetos
        // diferentes em paginas diferentes) fica o ultimo — suficiente para os documentos de uma fonte dos testes.
        var mapas = new Dictionary<string, Dictionary<int, string>>();
        foreach (var corpo in objetos.Values)
            foreach (Match f in Regex.Matches(corpo, @"/(F\w*)\s+(\d+)\s+0\s+R"))
            {
                if (!objetos.TryGetValue(int.Parse(f.Groups[2].Value), out var fonte)) continue;
                var tu = Regex.Match(fonte, @"/ToUnicode\s+(\d+)\s+0\s+R");
                if (!tu.Success || !objetos.TryGetValue(int.Parse(tu.Groups[1].Value), out var cmapObj)) continue;
                var cmap = Stream(cmapObj);
                if (cmap is not null) mapas[f.Groups[1].Value] = LerCMap(cmap);
            }

        var saida = new StringBuilder();
        foreach (var corpo in objetos.Values)
        {
            var conteudo = Stream(corpo);
            if (conteudo is null || !(conteudo.Contains("Tj") || conteudo.Contains("TJ"))) continue;
            Dictionary<int, string>? atual = null;
            // Tokens relevantes: /Nome tam Tf, <hex>, (literal), numero (kerning dentro de TJ), Td/TD/T*.
            foreach (Match t in Regex.Matches(conteudo, @"/(\w+)\s+[\d.]+\s+Tf|<([0-9A-Fa-f\s]*)>|\(((?:\\.|[^\\)])*)\)|(-?\d+(?:\.\d+)?)(?=\s*[<(\]])|\b(Td|TD|T\*)\b"))
            {
                if (t.Groups[1].Success) mapas.TryGetValue(t.Groups[1].Value, out atual);
                else if (t.Groups[2].Success) saida.Append(Decodificar(t.Groups[2].Value, atual));
                else if (t.Groups[3].Success) saida.Append(t.Groups[3].Value);
                else if (t.Groups[4].Success) { if (double.Parse(t.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture) < -150) saida.Append(' '); }
                else if (t.Groups[5].Success) saida.Append('\n');
            }
            saida.Append('\n');
        }
        return saida.ToString();
    }

    private static string Decodificar(string hex, Dictionary<int, string>? mapa)
    {
        hex = Regex.Replace(hex, @"\s", "");
        var sb = new StringBuilder();
        // Identity-H: codigos de 2 bytes (<0000> <FFFF> no codespacerange do x2t).
        for (var i = 0; i + 4 <= hex.Length; i += 4)
        {
            var codigo = Convert.ToInt32(hex.Substring(i, 4), 16);
            if (mapa is not null && mapa.TryGetValue(codigo, out var s)) sb.Append(s);
        }
        return sb.ToString();
    }

    private static Dictionary<int, string> LerCMap(string cmap)
    {
        var mapa = new Dictionary<int, string>();
        foreach (Match bloco in Regex.Matches(cmap, @"beginbfchar(.*?)endbfchar", RegexOptions.Singleline))
            foreach (Match par in Regex.Matches(bloco.Groups[1].Value, @"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>"))
                mapa[Convert.ToInt32(par.Groups[1].Value, 16)] = Utf16(par.Groups[2].Value);
        foreach (Match bloco in Regex.Matches(cmap, @"beginbfrange(.*?)endbfrange", RegexOptions.Singleline))
            foreach (Match faixa in Regex.Matches(bloco.Groups[1].Value, @"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>"))
            {
                int ini = Convert.ToInt32(faixa.Groups[1].Value, 16), fim = Convert.ToInt32(faixa.Groups[2].Value, 16);
                var destino = Convert.ToInt32(faixa.Groups[3].Value, 16);
                for (var c = ini; c <= fim && c - ini < 65536; c++) mapa[c] = char.ConvertFromUtf32(destino + c - ini);
            }
        return mapa;
    }

    private static string Utf16(string hex)
    {
        var bytes = Convert.FromHexString(hex.Length % 2 == 0 ? hex : "0" + hex);
        return Encoding.BigEndianUnicode.GetString(bytes);
    }

    /// <summary>Conteudo (descomprimido se FlateDecode) do stream de um objeto, ou <c>null</c>.</summary>
    private static string? Stream(string corpoObjeto)
    {
        var ini = corpoObjeto.IndexOf("stream", StringComparison.Ordinal);
        if (ini < 0) return null;
        ini += "stream".Length;
        if (ini < corpoObjeto.Length && corpoObjeto[ini] == '\r') ini++;
        if (ini < corpoObjeto.Length && corpoObjeto[ini] == '\n') ini++;
        var fim = corpoObjeto.IndexOf("endstream", ini, StringComparison.Ordinal);
        if (fim < 0) return null;
        var bruto = Encoding.Latin1.GetBytes(corpoObjeto[ini..fim]);
        if (!corpoObjeto[..ini].Contains("/FlateDecode")) return Encoding.Latin1.GetString(bruto);
        try
        {
            using var z = new ZLibStream(new MemoryStream(bruto), CompressionMode.Decompress);
            using var saida = new MemoryStream();
            z.CopyTo(saida);
            return Encoding.Latin1.GetString(saida.ToArray());
        }
        catch (InvalidDataException) { return null; }
    }
}
