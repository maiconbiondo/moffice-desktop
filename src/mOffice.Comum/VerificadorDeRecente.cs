namespace mOffice.Comum;

/// <summary>O que se sabe de um arquivo da lista de recentes. <see cref="SemResposta"/> NAO e' "nao existe": um caminho de rede fora do ar
/// nao permite concluir nada (o recente fica na lista, fixado ou nao).</summary>
public enum EstadoDoRecente { Existe, NaoExiste, SemResposta }

/// <summary>Confere se o recente ainda existe SEM nunca travar a interface: <c>File.Exists</c> num UNC fora do ar pode levar dezenas de segundos,
/// entao a sonda roda fora da thread da tela e so' vale o <c>limite</c> (Plano 6, Task 10 — SO-10).</summary>
public static class VerificadorDeRecente
{
    /// <summary>Sonda <paramref name="caminho"/> fora da thread de quem chama. Passou do <paramref name="limite"/> ou a sonda falhou
    /// (E/S, permissao, caminho invalido) = <see cref="EstadoDoRecente.SemResposta"/>: so' um "falso" explicito vira <see cref="EstadoDoRecente.NaoExiste"/>.
    /// A sonda que ficou para tras e' observada (nenhuma excecao tardia solta).</summary>
    public static async Task<EstadoDoRecente> VerificarAsync(string caminho, TimeSpan limite, Func<string, bool>? existe = null)
    {
        existe ??= File.Exists;
        var sonda = Task.Run(() => existe(caminho));
        try { return await sonda.WaitAsync(limite) ? EstadoDoRecente.Existe : EstadoDoRecente.NaoExiste; }
        catch (TimeoutException)
        {
            _ = sonda.ContinueWith(t => _ = t.Exception, TaskContinuationOptions.OnlyOnFaulted);
            return EstadoDoRecente.SemResposta;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return EstadoDoRecente.SemResposta;
        }
    }
}
