namespace mOffice.Nucleo.Recuperacao;

/// <summary>
/// Substituicao ATOMICA de um arquivo da pasta de recuperacao (revisao final do Plano 3, M2/M7): grava tudo num
/// <c>&lt;destino&gt;.tmp</c> ao lado, manda para o DISCO (<c>Flush(true)</c>) e so' entao troca pelo nome certo
/// (<c>File.Move(..., true)</c>). Uma queda em qualquer ponto deixa o arquivo antigo inteiro (ou o <c>.tmp</c>
/// orfao), nunca o arquivo pela metade — e os arquivos que passam por aqui (<c>sessao.json</c>,
/// <c>changes0.json</c>) sao exatamente os que a recuperacao le' depois de uma queda.
///
/// Sem o <c>Flush(true)</c> a troca de nome pode chegar ao disco ANTES dos dados (o conteudo ainda no cache do
/// sistema): uma queda de energia logo depois deixaria o nome novo apontando para um arquivo vazio/rasgado —
/// o mesmo rasgo que a escrita atomica existe para evitar.
///
/// A troca de nome NAO consegue substituir um arquivo que outro processo esta' lendo naquele instante (MEDIDO: com
/// um leitor aberto mesmo com <c>FileShare.Delete</c>, o <c>File.Move(..., true)</c> falha) — e o antivirus costuma
/// ler justamente o arquivo que acabou de ser gravado. Por isso a troca tenta algumas vezes, com espera curta
/// (<see cref="Tentativas"/>, ~0,5 s no total: isto roda na thread da interface), antes de desistir; desistindo, o
/// arquivo antigo continua inteiro e o <c>.tmp</c> e' apagado.
/// </summary>
internal static class GravacaoAtomica
{
    internal const int Tentativas = 10;

    public static void Gravar(string destino, ReadOnlySpan<byte> conteudo)
    {
        var tmp = destino + ".tmp";
        using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            fs.Write(conteudo);
            fs.Flush(flushToDisk: true);
        }
        try { ComNovasTentativas(() => File.Move(tmp, destino, true)); }
        catch
        {
            try { File.Delete(tmp); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            throw;
        }
    }

    /// <summary>
    /// TRUNCA <paramref name="destino"/> para <paramref name="prefixo"/> (fix round 2 da revisao final do Plano 3,
    /// item 1): primeiro pela troca atomica (<see cref="Gravar"/>); se ela nao conseguir (leitor segurando o arquivo —
    /// a impressao le' o <c>changes0.json</c> pelo x2t enquanto o usuario desfaz), NO LUGAR: abre o arquivo existente
    /// compartilhando leitura/escrita/delecao, confere que <paramref name="prefixo"/> e' mesmo o comeco dele e so'
    /// corta o fim (<c>SetLength</c> + <c>Flush(true)</c>) — os bytes que ficam nao sao reescritos, entao uma queda no
    /// meio deixa o arquivo antigo ou o prefixo, nunca um meio-termo. Se o comeco NAO bate (so' com um arquivo
    /// estranho), reescreve o prefixo e corta. Lanca se nem assim der: quem chama falha FECHADO.
    /// </summary>
    public static void Truncar(string destino, ReadOnlySpan<byte> prefixo)
    {
        try { Gravar(destino, prefixo); return; }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        var bytes = prefixo.ToArray();
        ComNovasTentativas(() =>
        {
            using var fs = new FileStream(destino, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
            var atual = new byte[Math.Min(fs.Length, bytes.Length)];
            fs.ReadExactly(atual);
            if (!atual.AsSpan().SequenceEqual(bytes))
            {
                fs.Position = 0;
                fs.Write(bytes);
            }
            fs.SetLength(bytes.Length);
            fs.Flush(flushToDisk: true);
        });
    }

    /// <summary>Roda <paramref name="acao"/> de novo (ate' <see cref="Tentativas"/> vezes, espera crescente de 10 ms)
    /// enquanto ela falhar por E/S ou permissao — o arquivo preso por um instante (antivirus, backup, outra janela
    /// lendo). A ultima falha sobe para quem chamou.</summary>
    public static void ComNovasTentativas(Action acao)
    {
        for (var i = 1; ; i++)
        {
            try { acao(); return; }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException && i < Tentativas) { Thread.Sleep(10 * i); }
        }
    }

    /// <summary>
    /// Leitura que NUNCA atrapalha quem esta' gravando (fix round 2, item 2): <c>File.ReadAllBytes</c> abre com
    /// <c>FileShare.Read</c> e, enquanto esta' aberto, o <c>FileMode.Append</c> da sessao viva (e a troca atomica)
    /// FALHAM. Aqui a leitura compartilha leitura, escrita e delecao. Lanca como <c>File.ReadAllBytes</c>.
    /// </summary>
    public static byte[] LerCompartilhado(string arquivo)
    {
        AoLerParaTeste?.Invoke(arquivo);
        using var fs = new FileStream(arquivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var memoria = new MemoryStream();
        fs.CopyTo(memoria);
        return memoria.ToArray();
    }

    /// <summary>Costura de teste: chamada com o caminho a cada <see cref="LerCompartilhado"/> (o teste filtra pelos
    /// caminhos DELE — outros testes rodam em paralelo). `internal`.</summary>
    internal static Action<string>? AoLerParaTeste;
}
