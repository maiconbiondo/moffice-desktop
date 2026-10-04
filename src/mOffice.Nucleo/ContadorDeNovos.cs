using System.Collections.Concurrent;

namespace mOffice.Nucleo;

/// <summary>Numeracao dos documentos novos como o Office pt-BR: Documento1, Documento2… por TIPO (word/cell/slide) e por processo, sem
/// reaproveitar o numero de um documento fechado (Plano 6, Task 7).</summary>
public static class ContadorDeNovos
{
    private static readonly ConcurrentDictionary<string, int> Contagem = new(StringComparer.Ordinal);

    /// <summary>Proximo numero (a partir de 1) para o tipo de documento <paramref name="docType"/>.</summary>
    public static int Proximo(string docType) => Contagem.AddOrUpdate(docType, 1, (_, n) => n + 1);
}
