namespace mOffice.Comum;

public enum ModoAbertura { Abrir, NovoDoModelo, Apresentar, Imprimir }

/// <summary>O que a linha de comando (ou o pipe da segunda instancia) pede. Verbos do Explorer gerados pela tabela:
/// open de modelo = <c>/modelo "%1"</c>; open de apresentacao de slides = <c>/apresentar "%1"</c>; o resto = <c>"%1"</c>.
/// Imprimir (verbo <c>print</c> do Explorer, Plano 6 Task 10) = <c>/imprimir "%1"</c>: processo EFEMERO (converte, imprime e sai) — nunca vai pelo pipe.
/// Pelo pipe: o caminho puro (Abrir) ou <c>?modelo|&lt;caminho&gt;</c> / <c>?apresentar|&lt;caminho&gt;</c> — sempre com
/// caminho ABSOLUTO (a mesma regra que a InstanciaUnica ja' aplicava ao caminho puro).</summary>
public sealed record PedidoDeAbertura(string Caminho, ModoAbertura Modo)
{
    private static readonly IReadOnlyDictionary<string, ModoAbertura> Verbos = new Dictionary<string, ModoAbertura>(StringComparer.OrdinalIgnoreCase)
        { ["/modelo"] = ModoAbertura.NovoDoModelo, ["/apresentar"] = ModoAbertura.Apresentar, ["/imprimir"] = ModoAbertura.Imprimir };

    /// <summary>O pedido da linha de comando: o primeiro argumento que e' arquivo aceito pelo <paramref name="perfil"/>, com o modo do
    /// verbo que veio ANTES dele. <paramref name="rejeitado"/> = o primeiro arquivo que o app NAO abre (FMT-05), para a casca avisar.</summary>
    public static PedidoDeAbertura? DosArgumentos(string[] args, PerfilApp perfil, out string? rejeitado)
    {
        rejeitado = null;
        var modo = ModoAbertura.Abrir;
        foreach (var a in args)
        {
            if (string.IsNullOrWhiteSpace(a)) continue;
            if (Verbos.TryGetValue(a, out var m)) { modo = m; continue; }
            if (a.StartsWith('-') || a.StartsWith('/')) continue;
            string caminho;
            try { caminho = Path.GetFullPath(a); }
            catch (Exception e) when (e is ArgumentException or NotSupportedException or PathTooLongException or System.Security.SecurityException) { return null; }
            if (!perfil.Aceita(caminho)) { rejeitado ??= caminho; continue; }
            return new PedidoDeAbertura(caminho, modo);
        }
        return null;
    }

    /// <exception cref="InvalidOperationException">Para <see cref="ModoAbertura.Imprimir"/>: imprimir e' processo efemero, nao encaminha para a instancia primaria.</exception>
    public string ParaPipe() => Modo switch
    {
        ModoAbertura.Imprimir => throw new InvalidOperationException("O pedido de impressão não vai pelo pipe: é um processo efêmero."),
        ModoAbertura.NovoDoModelo => "?modelo|" + Caminho,
        ModoAbertura.Apresentar => "?apresentar|" + Caminho,
        _ => Caminho,
    };

    /// <summary>FMT-05: a segunda instancia recebeu um arquivo que o app NAO abre — manda <c>?rejeitado|&lt;caminho&gt;</c> para a primaria
    /// mostrar a mesma mensagem pt-BR (antes a segunda saia calada).</summary>
    public static string RejeicaoParaPipe(string caminho) => "?rejeitado|" + caminho;

    /// <summary>O caminho (absoluto) de uma linha <c>?rejeitado|..</c>, ou null.</summary>
    public static string? DoPipeRejeitado(string linha) =>
        linha.StartsWith("?rejeitado|", StringComparison.Ordinal) && Path.IsPathRooted(linha["?rejeitado|".Length..]) ? linha["?rejeitado|".Length..] : null;

    public static PedidoDeAbertura? DoPipe(string linha)
    {
        if (string.IsNullOrWhiteSpace(linha)) return null;
        var modo = ModoAbertura.Abrir;
        var caminho = linha;
        if (linha.StartsWith('?'))
        {
            var barra = linha.IndexOf('|');
            if (barra < 0) return null;
            modo = linha[..barra] switch { "?modelo" => ModoAbertura.NovoDoModelo, "?apresentar" => ModoAbertura.Apresentar, _ => (ModoAbertura)(-1) };
            if ((int)modo < 0) return null;
            caminho = linha[(barra + 1)..];
        }
        return Path.IsPathRooted(caminho) ? new PedidoDeAbertura(caminho, modo) : null;
    }
}
