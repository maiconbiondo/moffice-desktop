namespace mOffice.Nucleo;

/// <summary>Logs de diagnostico e egresso por janela, em %LOCALAPPDATA%\mOffice\logs (backlog 15).
/// Rotacao por IDADE e por TAMANHO: a varredura na subida do app apaga .log com mais de 30 dias
/// (<see cref="RetencaoPadrao"/>) e, depois disso, os mais ANTIGOS enquanto a pasta passar de 200 MB
/// (<see cref="RetencaoBytes"/>).</summary>
public static class Logs
{
    public static readonly TimeSpan RetencaoPadrao = TimeSpan.FromDays(30);

    /// <summary>Teto da pasta de logs. So' a idade nao basta (achado da revisao final do Plano 2): um unico
    /// dia de uso intenso — varias janelas abertas, cada uma com o seu .log — cabe folgadamente dentro dos
    /// 30 dias e mesmo assim pode encher o disco do usuario, e a spec §4 sempre falou em "rotacao por
    /// tamanho". 200 MB e' generoso para diagnostico (um log de sessao normal fica na casa das dezenas de
    /// KB) e pequeno demais para incomodar qualquer disco.</summary>
    public const long RetencaoBytes = 200L * 1024 * 1024;

    // Revisao 1 da Task 1 (achado do review): app+timestamp+pid e' unico por SEGUNDO, nao por JANELA — duas
    // janelas do MESMO app no MESMO processo no MESMO segundo (ex.: `mOffice.Sonda --abrir a.docx b.docx`,
    // que cria uma JanelaEditor por arquivo em sequencia rapida) cairiam no mesmo arquivo. Um contador por
    // PROCESSO, incrementado a cada chamada, desambigua sem depender do relogio.
    private static int _contador;

    public static string NovoArquivo(string app) => NovoArquivoEm(Pastas.Logs, app);

    /// <summary>Nome: <c>&lt;app&gt;-&lt;yyyyMMdd-HHmmss&gt;-&lt;pid&gt;-&lt;n&gt;.log</c>, com <c>n</c> a
    /// partir de 1 e unico por PROCESSO (nao reinicia por app nem por pasta) — e' o que garante que duas
    /// chamadas nunca devolvem o mesmo caminho, mesmo no mesmo segundo.</summary>
    public static string NovoArquivoEm(string pasta, string app)
    {
        Pastas.Garantir(pasta);
        var n = Interlocked.Increment(ref _contador);
        var nome = $"{app.ToLowerInvariant()}-{DateTime.Now:yyyyMMdd-HHmmss}-{Environment.ProcessId}-{n}.log";
        return Path.Combine(pasta, nome);
    }

    /// <summary>
    /// Apaga os .log mais velhos que <paramref name="idadeMinima"/> e, em seguida, os .log mais ANTIGOS que
    /// ainda restarem enquanto a soma dos tamanhos passar de <paramref name="limiteBytes"/>. Devolve quantos
    /// arquivos foram apagados nas duas passadas somadas. Nunca lanca: log e' diagnostico, nao pode derrubar
    /// a subida do app.
    ///
    /// A ordem importa: idade primeiro (regra barata e previsivel), tamanho depois, sobre o que sobrou.
    ///
    /// Por que o log da sessao corrente nao corre risco: a varredura roda UMA vez, na SUBIDA do app
    /// (<c>AplicativoOffice.OnStartup</c>, antes de existir qualquer janela), entao nenhum arquivo desta
    /// execucao existe ainda quando ela passa. Nao ha' nada aqui que proteja um <c>.log</c> em uso — e nao
    /// adiantaria haver: o <see cref="PonteArquivos"/> escreve com <c>File.AppendAllText</c>, que abre e
    /// fecha o arquivo a cada linha, entao entre duas linhas o arquivo NAO esta' travado e um
    /// <c>File.Delete</c> passaria liso. A garantia e' de MOMENTO (varrer antes de comecar a escrever), nao
    /// de trava de arquivo.
    /// </summary>
    public static int Varrer(string pasta, TimeSpan idadeMinima, long limiteBytes = RetencaoBytes)
    {
        if (!Directory.Exists(pasta)) return 0;
        var limite = DateTime.UtcNow - idadeMinima;
        var apagados = 0;
        // Sobrevive ao que foi apagado pela idade: a lista e' materializada e refiltrada, sem reenumerar a
        // pasta (que poderia pegar arquivos criados no meio da propria varredura).
        var restantes = new List<FileInfo>();
        foreach (var arquivo in Directory.EnumerateFiles(pasta, "*.log"))
        {
            try
            {
                var info = new FileInfo(arquivo);
                if (info.LastWriteTimeUtc < limite) { info.Delete(); apagados++; }
                else restantes.Add(info);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        var total = 0L;
        foreach (var info in restantes) { try { total += info.Length; } catch (IOException) { } }
        if (total <= limiteBytes) return apagados;

        // Do mais antigo para o mais novo, ate' caber. O tamanho e' o lido acima (FileInfo e' um instantaneo):
        // um .log que cresceu no meio disto so' faz a passada parar um pouco mais tarde, na proxima subida.
        foreach (var info in restantes.OrderBy(i => i.LastWriteTimeUtc))
        {
            if (total <= limiteBytes) break;
            try { var tamanho = info.Length; info.Delete(); total -= tamanho; apagados++; }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        return apagados;
    }
}
