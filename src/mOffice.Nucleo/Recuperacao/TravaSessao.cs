namespace mOffice.Nucleo.Recuperacao;

/// <summary>
/// Marca "esta pasta de recuperacao esta' em uso por uma janela viva" (Task 2 do Plano 3):
/// <c>&lt;pasta&gt;\rec.lock</c> aberto com <see cref="FileShare.None"/> durante toda a sessao. Quem precisa
/// saber se a pasta e' de uma sessao que CAIU (e nao de uma janela aberta, talvez em outro processo) tenta abrir
/// o mesmo arquivo: se conseguir, ninguem o segura. O sistema operacional solta o arquivo sozinho quando o
/// processo morre — e' exatamente o que distingue "queda" de "aberto".
/// </summary>
public sealed class TravaSessao : IDisposable
{
    public const string NomeArquivo = "rec.lock";

    private FileStream? _fs;

    private TravaSessao(FileStream fs) => _fs = fs;

    private static FileStream Abrir(string pasta) =>
        new(Path.Combine(pasta, NomeArquivo), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);

    /// <summary>Trava a pasta, ou <c>null</c> se ela ja' esta' travada por outra sessao (ou nao da' para criar
    /// o arquivo).</summary>
    public static TravaSessao? Adquirir(string pasta)
    {
        try { return new TravaSessao(Abrir(pasta)); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return null; }
    }

    /// <summary><c>true</c> quando nenhuma sessao viva segura a pasta. NAO cria nada (fix round 1 da revisao da
    /// Task 2): a Task 3 sonda pastas que nao sao dela, e uma sonda que criasse <c>rec.lock</c> mudaria o
    /// frescor da pasta (a varredura olha a escrita mais recente dentro dela) e deixaria rastro. Sem a pasta ou
    /// sem o <c>rec.lock</c> nao ha' trava nenhuma: livre. Com o arquivo, abre-o (<see cref="FileMode.Open"/>)
    /// com <see cref="FileShare.None"/> e fecha na hora; se outro o segura, nao esta' livre. Sem permissao
    /// para abrir tambem conta como "nao livre" — o lado seguro para quem pensa em apagar a pasta.</summary>
    public static bool Livre(string pasta)
    {
        var arquivo = Path.Combine(pasta, NomeArquivo);
        if (!File.Exists(arquivo)) return true;
        try
        {
            using var fs = new FileStream(arquivo, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return true;
        }
        catch (Exception e) when (e is FileNotFoundException or DirectoryNotFoundException) { return true; } // sumiu entre o Exists e o Open
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return false; }
    }

    public void Dispose()
    {
        _fs?.Dispose();
        _fs = null;
    }
}
