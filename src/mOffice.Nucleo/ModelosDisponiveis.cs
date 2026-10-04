using System.Text;
using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>Um modelo da galeria "Novo": <see cref="Nome"/> para exibir, <see cref="Caminho"/> do arquivo (.dotx/.xltx/.potx…),
/// <see cref="DoUsuario"/> = veio da pasta de modelos do usuario (senao, dos modelos que o motor traz).</summary>
public sealed record ModeloDisponivel(string Nome, string Caminho, bool DoUsuario);

/// <summary>Lista os modelos da galeria (Plano 6, Task 9): primeiro os da pasta de modelos do usuario, depois os que o motor
/// traz em <c>converter\templates\PT</c>. Nunca lanca: uma pasta de rede fora do ar vira lista vazia.</summary>
public static class ModelosDisponiveis
{
    /// <param name="docType"><c>word</c>, <c>cell</c> ou <c>slide</c>.</param>
    /// <param name="pastaDoUsuario">Pasta de modelos pessoais (so' o nivel de cima); <c>null</c> = nenhuma.</param>
    /// <param name="pastaModelosDoMotor"><c>&lt;motor&gt;\converter\templates\PT</c>.</param>
    public static IReadOnlyList<ModeloDisponivel> Listar(string docType, string? pastaDoUsuario, string pastaModelosDoMotor)
    {
        var lista = new List<ModeloDisponivel>();
        lista.AddRange(DoUsuario(docType, pastaDoUsuario));
        lista.AddRange(DoMotor(docType, pastaModelosDoMotor));
        return lista;
    }

    private static IEnumerable<ModeloDisponivel> DoUsuario(string docType, string? pasta)
    {
        if (string.IsNullOrWhiteSpace(pasta)) return [];
        try
        {
            return Directory.EnumerateFiles(pasta)
                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal)
                            && !Path.GetFileNameWithoutExtension(f).Equals("Normal", StringComparison.OrdinalIgnoreCase)
                            && TabelaFormatos.DaExtensao(f) is { Abre: true, Modelo: true } i && i.DocType == docType)
                .Select(f => new ModeloDisponivel(Path.GetFileNameWithoutExtension(f), f, true))
                .OrderBy(m => m.Nome, StringComparer.OrdinalIgnoreCase).ThenBy(m => m.Caminho, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return []; }
    }

    private static IEnumerable<ModeloDisponivel> DoMotor(string docType, string pastaPt)
    {
        var sub = docType switch { "word" => "Documents", "cell" => "Spreadsheet", "slide" => "Presentation", _ => null };
        if (sub is null || string.IsNullOrWhiteSpace(pastaPt)) return [];
        try
        {
            var pasta = Path.Combine(pastaPt, sub);
            if (!Directory.Exists(pasta)) return [];
            return Directory.EnumerateFiles(pasta)
                .Where(f => TabelaFormatos.DaExtensao(f) is { Abre: true, Modelo: true } i && i.DocType == docType)
                .Select(f => new ModeloDisponivel(NomeDoMotor(f), f, false))
                .OrderBy(m => m.Nome, StringComparer.OrdinalIgnoreCase).ThenBy(m => m.Caminho, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return []; }
    }

    /// <summary>Nome legivel de um modelo do motor: o arquivo e' <c>[32]</c> + base32 (RFC 4648, com ou sem padding) do nome em UTF-8.
    /// Fora desse formato, devolve o nome do arquivo sem extensao.</summary>
    public static string NomeDoMotor(string arquivo)
    {
        var nome = Path.GetFileNameWithoutExtension(arquivo);
        if (!nome.StartsWith("[32]", StringComparison.Ordinal)) return nome;
        try { return Encoding.UTF8.GetString(Base32(nome[4..])); }
        catch (FormatException) { return nome; }
    }

    private static byte[] Base32(string texto)
    {
        const string alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
        var saida = new List<byte>();
        int buffer = 0, bits = 0;
        foreach (var c in texto.TrimEnd('=').ToUpperInvariant())
        {
            var v = alfabeto.IndexOf(c);
            if (v < 0) throw new FormatException("Caractere fora do alfabeto base32.");
            buffer = (buffer << 5) | v; bits += 5;
            if (bits >= 8) { bits -= 8; saida.Add((byte)((buffer >> bits) & 0xFF)); buffer &= (1 << bits) - 1; }
        }
        return saida.ToArray();
    }
}
