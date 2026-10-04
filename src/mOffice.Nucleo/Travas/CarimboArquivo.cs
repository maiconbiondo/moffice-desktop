namespace mOffice.Nucleo.Travas;

/// <summary>Estado do arquivo do usuario no disco (data de escrita UTC + tamanho), tirado na abertura e renovado a cada
/// salvamento (Plano 6, Task 5 — SO-02). Se mudou na hora de salvar por cima, alguem gravou o arquivo por fora: o salvar
/// pergunta antes (<see cref="DecisaoConflito"/>).</summary>
public readonly record struct CarimboArquivo(DateTime EscritaUtc, long Tamanho)
{
    /// <summary><c>null</c> se o arquivo nao existe ou nao pode ser lido.</summary>
    public static CarimboArquivo? De(string caminho)
    {
        try
        {
            var i = new FileInfo(caminho);
            return i.Exists ? new CarimboArquivo(i.LastWriteTimeUtc, i.Length) : null;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return null; }
    }
}

/// <summary>Resposta a "o arquivo foi alterado por outra pessoa ou programa desde que voce o abriu".</summary>
public enum DecisaoConflito { SalvarCopia, Substituir, Cancelar }
