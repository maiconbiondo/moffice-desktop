namespace mOffice.Nucleo.Recuperacao;

/// <summary>"Não salvos recentemente" (Plano 6, Task 10 — SO-11): o "Não salvar" do fechamento deixa de apagar a pasta de recuperação na hora
/// e a MOVE para <see cref="Pastas.Descartados"/>, onde ela fica <see cref="Retencao"/> (como o "Recuperar documentos não salvos" do Office).
/// A pasta continua com o formato da recuperação (<c>sessao.json</c>, <c>Editor.bin</c>, <c>changes/</c>): <see cref="Recuperaveis.Listar"/> e
/// <c>RecuperarAsync</c> funcionam sobre ela. Excluir é SEMPRE ato explícito do usuário (<see cref="Recuperaveis.Descartar"/>); a varredura só leva o que
/// passou da retenção e nunca toca pasta com a trava presa (janela que a recuperou e está aberta).</summary>
public static class Descartados
{
    /// <summary>Marca com a hora do descarte: a retenção conta daqui, não da última alteração (um documento editado há 5 dias e descartado agora fica 4 dias).</summary>
    public const string NomeMarca = "descartado.txt";

    /// <summary>A pasta esta' DENTRO da raiz dos descartados (comparacao de caminho completo, sem diferenciar maiusculas)?</summary>
    public static bool Contem(string raiz, string pasta)
    {
        try
        {
            var r = Path.GetFullPath(raiz).TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
            return Path.GetFullPath(pasta).StartsWith(r, StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException) { return false; }
    }

    /// <summary>Reabrir dos "Nao salvos" (backlog 56): a pasta volta para a RECUPERACAO — la' uma queda ou logoff a deixa listada como "Documentos recuperados",
    /// e nao sob a varredura de 4 dias. Move (rename) e SO' DEPOIS apaga a marca do descarte. Se o move falhar (alguem segura um arquivo), abre uma COPIA na
    /// recuperacao (sem trava nem marca) e deixa o descartado intocado: nunca se perde. Devolve a pasta a usar, ou null se nem a copia deu.</summary>
    public static string? Repor(string pasta, string raizRecuperacao)
    {
        try
        {
            Pastas.Garantir(raizRecuperacao);
            var nome = new DirectoryInfo(pasta).Name;
            var destino = Path.Combine(raizRecuperacao, nome);
            for (var n = 2; Directory.Exists(destino); n++) destino = Path.Combine(raizRecuperacao, nome + "-" + n);
            try
            {
                Directory.Move(pasta, destino);
                try { File.Delete(Path.Combine(destino, NomeMarca)); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
                return destino;
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                try { Copiar(pasta, destino); return destino; }
                catch (Exception e2) when (e2 is IOException or UnauthorizedAccessException)
                {
                    try { if (Directory.Exists(destino)) Directory.Delete(destino, true); } catch (Exception e3) when (e3 is IOException or UnauthorizedAccessException) { }
                    return null;
                }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return null; }
    }

    private static void Copiar(string origem, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var arq in Directory.EnumerateFiles(origem))
        {
            var nome = Path.GetFileName(arq);
            if (string.Equals(nome, TravaSessao.NomeArquivo, StringComparison.OrdinalIgnoreCase) || string.Equals(nome, NomeMarca, StringComparison.OrdinalIgnoreCase)) continue;
            using var de = new FileStream(arq, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var para = new FileStream(Path.Combine(destino, nome), FileMode.CreateNew, FileAccess.Write, FileShare.None);
            de.CopyTo(para);
        }
        foreach (var sub in Directory.EnumerateDirectories(origem)) Copiar(sub, Path.Combine(destino, Path.GetFileName(sub)));
    }

    public static TimeSpan Retencao { get; } = TimeSpan.FromDays(4);

    /// <summary>Move a pasta inteira (sem copiar nada: o rename na mesma unidade é atômico) para <paramref name="raizDescartados"/>. Devolve o novo caminho, ou
    /// null se não deu (pasta ausente, trava presa, E/S) — quem chama apaga como antes ou deixa para a varredura. Nunca lança.</summary>
    public static string? Guardar(string pastaRecovery, string raizDescartados)
    {
        try
        {
            if (!Directory.Exists(pastaRecovery)) return null;
            var info = new DirectoryInfo(pastaRecovery);
            if (info.LinkTarget is not null || info.Attributes.HasFlag(FileAttributes.ReparsePoint)) return null;
            Pastas.Garantir(raizDescartados);
            var destino = Path.Combine(raizDescartados, info.Name);
            for (var n = 2; Directory.Exists(destino); n++) destino = Path.Combine(raizDescartados, info.Name + "-" + n);
            // A marca vai ANTES do move: entre o move e a marca a pasta chegaria aos descartados com a data da ultima edicao e a varredura
            // da subida a levaria (documento editado ha 5 dias, descartado agora). Marca pronta = a retencao conta do descarte desde o 1o instante.
            try { File.WriteAllText(Path.Combine(pastaRecovery, NomeMarca), DateTimeOffset.Now.ToString("O")); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            Directory.Move(pastaRecovery, destino);
            return destino;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return null; }
    }

    /// <summary>Apaga as pastas de <paramref name="raiz"/> sem escrita há mais de <paramref name="idade"/> (maior data entre a pasta e tudo dentro dela,
    /// inclusive a marca do descarte). Pasta com a trava presa (recuperada e aberta numa janela) nunca é tocada. Raiz que é link = nada. Devolve quantas
    /// apagou; nunca lança.</summary>
    public static int Varrer(string raiz, TimeSpan idade)
    {
        var apagadas = 0;
        try
        {
            if (!Directory.Exists(raiz)) return 0;
            var infoRaiz = new DirectoryInfo(raiz);
            if (infoRaiz.LinkTarget is not null || infoRaiz.Attributes.HasFlag(FileAttributes.ReparsePoint)) return 0;
            var limite = DateTime.UtcNow - idade;
            foreach (var pasta in Directory.EnumerateDirectories(raiz))
            {
                try
                {
                    var p = new DirectoryInfo(pasta);
                    if (p.LinkTarget is not null || p.Attributes.HasFlag(FileAttributes.ReparsePoint)) continue;
                    if (UltimaEscritaUtc(pasta) >= limite) continue;
                    if (!TravaSessao.Livre(pasta)) continue;
                    Directory.Delete(pasta, true);
                    apagadas++;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return apagadas;
    }

    private static DateTime UltimaEscritaUtc(string pasta)
    {
        var maisRecente = Directory.GetLastWriteTimeUtc(pasta);
        try
        {
            foreach (var arq in Directory.EnumerateFiles(pasta, "*", SearchOption.AllDirectories))
            {
                try { var t = File.GetLastWriteTimeUtc(arq); if (t > maisRecente) maisRecente = t; }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return maisRecente;
    }
}
