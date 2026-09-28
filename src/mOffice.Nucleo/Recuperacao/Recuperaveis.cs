namespace mOffice.Nucleo.Recuperacao;

/// <summary>Um documento que uma sessao anterior deixou sem salvar (Task 3 do Plano 3): a pasta de recuperacao,
/// o titulo e a origem gravados em <c>sessao.json</c>, se era um documento NOVO (sem arquivo do usuario) e a
/// hora da ultima alteracao gravada (<c>changes/changes0.json</c>).</summary>
public sealed record DocumentoRecuperavel(string Pasta, string Titulo, string Origem, bool Novo, DateTime UltimaAlteracaoUtc)
{
    /// <summary>A mesma hora no fuso da maquina, para a tela (o record guarda UTC, como o sistema de arquivos).</summary>
    public DateTime UltimaAlteracaoLocal => DateTime.SpecifyKind(UltimaAlteracaoUtc, DateTimeKind.Utc).ToLocalTime();
}

/// <summary>
/// Quais pastas de <see cref="Pastas.Recuperacao"/> guardam trabalho nao salvo de uma sessao que CAIU (Task 3
/// do Plano 3), e o descarte delas. Recuperavel = <c>sessao.json</c> legivel, <c>Editor.bin</c>,
/// <c>changes/changes0.json</c> com pelo menos 1 change e a trava (<see cref="TravaSessao"/>) LIVRE — a pasta
/// de uma janela viva (neste processo ou em outro) nunca aparece, nunca e' descartada.
///
/// So' LE as pastas (spec §8, dado do usuario vem primeiro): a leitura dos changes e'
/// <see cref="ChangesGravados.Ler"/>, pura. O conserto de um <c>changes0.json</c> danificado
/// (<see cref="ChangesGravados.Reparar"/>) e' do dono da pasta — a janela que a recupera, com a trava na mao.
/// </summary>
public static class Recuperaveis
{
    /// <summary>Os documentos recuperaveis do app <paramref name="app"/> (mtexto/mplanilha/mslides — o
    /// <c>MetadadosSessao.App</c>, sem diferenciar maiusculas), do mais recente para o mais antigo. Nunca lanca:
    /// raiz ausente = lista vazia; raiz que e' link simbolico/juncao = lista vazia (mesma recusa de
    /// <see cref="SessaoDocumento.VarrerRecoveryAntigo"/>); uma pasta ilegivel fica de fora sem derrubar a lista.</summary>
    public static IReadOnlyList<DocumentoRecuperavel> Listar(string raiz, string app)
    {
        var lista = new List<DocumentoRecuperavel>();
        try
        {
            if (!Directory.Exists(raiz) || EhLink(raiz)) return lista;
            foreach (var pasta in Directory.EnumerateDirectories(raiz))
            {
                try
                {
                    if (EhLink(pasta)) continue;
                    // Trava PRIMEIRO (fix round 2 da revisao final do Plano 3): a pasta de uma sessao VIVA (desta ou de
                    // outra instancia) nunca e' lida — nem o sessao.json, nem o changes0.json. Ler um arquivo que a sessao
                    // viva esta' gravando atrapalhava a gravacao dela (a boas-vindas lista a cada recarga).
                    if (!TravaSessao.Livre(pasta)) continue;
                    var m = MetadadosSessao.Ler(pasta);
                    if (m is null || !string.Equals(m.App, app, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!EhRecuperavel(pasta)) continue;
                    lista.Add(new DocumentoRecuperavel(pasta, m.Titulo, m.Origem, m.Novo, File.GetLastWriteTimeUtc(ArquivoChanges(pasta))));
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return lista.OrderByDescending(d => d.UltimaAlteracaoUtc).ToList();
    }

    /// <summary>A pasta guarda trabalho nao salvo de uma sessao que caiu? So' sonda: nao cria nem altera nada
    /// (<see cref="TravaSessao.Livre"/> nao cria <c>rec.lock</c>; <see cref="ChangesGravados.Ler"/> so' le').</summary>
    // A trava vem ANTES do conteudo (fix round 2 da revisao final do Plano 3): de uma pasta viva nada e' lido.
    public static bool EhRecuperavel(string pasta) => TravaSessao.Livre(pasta) && TemConteudo(pasta);

    /// <summary>
    /// Apaga a pasta de um documento recuperavel — e SO' de um recuperavel: pasta de sessao viva, sem changes ou
    /// sem metadados fica intacta (nada acontece). Melhor esforco, nunca lanca: o que nao der para apagar agora
    /// continua na lista (quem chama recarrega e ve') ou cai na varredura.
    ///
    /// A pasta e' apagada COM A TRAVA NA MAO: um <c>Directory.Delete</c> recursivo direto sobre uma pasta que
    /// alguem acabou de abrir (uma janela recuperando-a entre a sonda e o apagar) apagaria o conteudo e falharia
    /// so' no <c>rec.lock</c> travado — a sessao viva ficaria sem <c>Editor.bin</c>. Com a trava, quem chegar
    /// depois nao consegue abrir a pasta; o <c>rec.lock</c> e a pasta vazia saem por ultimo.
    /// </summary>
    public static void Descartar(string pasta)
    {
        try
        {
            if (!Directory.Exists(pasta) || EhLink(pasta) || !EhRecuperavel(pasta)) return;
            using (var trava = TravaSessao.Adquirir(pasta))
            {
                // Alguem pegou a pasta entre a sonda e aqui, ou ela mudou: nao e' mais nossa para apagar.
                if (trava is null || !TemConteudo(pasta)) return;
                // sessao.json primeiro: se algo falhar no meio, a pasta ja' deixou de ser recuperavel (sem
                // metadados) em vez de reaparecer na lista pela metade; o resto cai na varredura.
                File.Delete(Path.Combine(pasta, MetadadosSessao.NomeArquivo));
                // Directory.Delete recursivo NAO segue juncao/link (apaga so' o link), entao uma subpasta assim
                // nao leva nada de fora da pasta de recuperacao.
                foreach (var sub in Directory.EnumerateDirectories(pasta)) Directory.Delete(sub, true);
                foreach (var arq in Directory.EnumerateFiles(pasta))
                    if (!string.Equals(Path.GetFileName(arq), TravaSessao.NomeArquivo, StringComparison.OrdinalIgnoreCase)) File.Delete(arq);
            }
            Directory.Delete(pasta, true);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }

    private static string ArquivoChanges(string pasta) => Path.Combine(pasta, "changes", "changes0.json");

    /// <summary>O que torna a pasta recuperavel, sem olhar a trava.</summary>
    internal static bool TemConteudo(string pasta) =>
        MetadadosSessao.Ler(pasta) is not null
        && File.Exists(Path.Combine(pasta, "Editor.bin"))
        && ChangesGravados.Ler(ArquivoChanges(pasta)).Quantidade > 0;

    private static bool EhLink(string pasta)
    {
        var info = new DirectoryInfo(pasta);
        return info.LinkTarget is not null || info.Attributes.HasFlag(FileAttributes.ReparsePoint);
    }
}
