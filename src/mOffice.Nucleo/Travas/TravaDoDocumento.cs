namespace mOffice.Nucleo.Travas;

/// <summary>
/// A trava de um documento aberto, como a do Office (Plano 6, Task 5 — SO-02/MvP-05): um handle sobre o original com
/// <c>FileAccess.Read</c> + <c>FileShare.Read</c> (os outros leem, ninguem grava — o Word do colega abre "somente leitura") e o arquivo
/// de dono <c>~$</c> com o nome do autor.
///
/// Salvar por cima (MEDIDO, <c>TravasTests.Trava_SalvamentoComATravaNaMao_SubstituiERetrava</c>): com o handle em
/// <c>FileShare.Read</c> o <c>File.Replace</c> do PROPRIO processo e' recusado (violacao de compartilhamento). Por isso, durante a troca,
/// <see cref="PrepararSubstituicao"/> passa a trava para um handle <c>Read</c> + <c>FileShare.Read | Delete</c> — terceiros continuam
/// sem gravar, e o ReplaceFile (que renomeia o original para o backup) passa. Esse handle acompanha o backup, entao depois da troca
/// <see cref="Retomar"/> abre a trava sobre o arquivo NOVO. O handle normal nao tem <c>FileShare.Delete</c> de proposito: como no Office,
/// o documento aberto nao pode ser apagado nem renomeado por fora.
/// </summary>
public sealed class TravaDoDocumento : IDisposable
{
    private string _documento;
    private readonly string _autor;
    private readonly Action<string> _log;
    private FileStream? _handle;
    private ArquivoDeDono? _dono;

    private TravaDoDocumento(string documento, string autor, Action<string> log, FileStream handle, ArquivoDeDono? dono)
    {
        _documento = documento; _autor = autor; _log = log; _handle = handle; _dono = dono;
    }

    /// <summary>O documento travado (muda no <see cref="Mover"/>).</summary>
    public string Documento => _documento;

    /// <summary>A trava esta' segurando o documento agora.</summary>
    public bool Segurando => _handle is not null;

    private static FileStream AbrirHandle(string documento, FileShare share) =>
        new(documento, FileMode.Open, FileAccess.Read, share);

    /// <summary>Trava <paramref name="documento"/> e cria o <c>~$</c>. Falha de E/S no documento = <c>null</c> com log (o documento abre do
    /// mesmo jeito, so' sem trava). Falha so' no <c>~$</c> (preso por outro, pasta sem escrita de arquivo novo) = trava sem <c>~$</c>, com log.</summary>
    public static TravaDoDocumento? Adquirir(string documento, string nomeDoAutor, Action<string> log)
    {
        FileStream handle;
        try { handle = AbrirHandle(documento, FileShare.Read); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            log($"TRAVA nao adquirida em {Path.GetFileName(documento)}: {e.Message} (a sessao segue sem trava)");
            return null;
        }
        var dono = ArquivoDeDono.Criar(documento, nomeDoAutor);
        if (dono is null) log($"TRAVA sem arquivo de dono (~$) para {Path.GetFileName(documento)}: ja' existe preso por outro processo, ou a pasta nao aceita (o documento continua travado)");
        log($"TRAVA adquirida em {Path.GetFileName(documento)}{(dono is null ? "" : " (+ " + Path.GetFileName(dono.Caminho) + ")")}");
        return new TravaDoDocumento(documento, nomeDoAutor, log, handle, dono);
    }

    /// <summary>Solta o documento (outros passam a poder gravar); o <c>~$</c> continua.</summary>
    public void Soltar()
    {
        _handle?.Dispose();
        _handle = null;
    }

    /// <summary>Durante a troca do salvamento: segura o documento com <c>FileShare.Read | Delete</c> (ver o comentario da classe) — o
    /// handle novo e' aberto ANTES de o normal fechar, entao nao ha' instante sem trava. <c>false</c> (com log) se nao deu: a trava fica
    /// solta e o salvamento segue (um <see cref="Retomar"/> depois reabre).</summary>
    public bool PrepararSubstituicao()
    {
        try
        {
            var troca = AbrirHandle(_documento, FileShare.Read | FileShare.Delete);
            _handle?.Dispose();
            _handle = troca;
            return true;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            _log($"TRAVA nao passou ao modo de troca ({e.Message}): solta durante o salvamento");
            Soltar();
            return false;
        }
    }

    /// <summary>Costura de teste: todo <see cref="Retomar"/> falha (como um arquivo que outro programa pegou logo depois da troca). `internal`.</summary>
    internal bool FalharRetomarParaTeste { get; set; }

    /// <summary>Volta a travar o documento (o handle e' aberto de novo pelo caminho: depois de um ReplaceFile ele aponta para o arquivo
    /// NOVO). Falha = <c>false</c> + log; a sessao segue sem trava.</summary>
    public bool Retomar()
    {
        try
        {
            if (FalharRetomarParaTeste) throw new IOException("falha simulada do Retomar (teste)");
            var novo = AbrirHandle(_documento, FileShare.Read);
            _handle?.Dispose();
            _handle = novo;
            // Sem ~$ (um Mover que falhou, ou o ~$ estava preso na aquisicao): tenta de novo agora que o documento esta' travado.
            _dono ??= ArquivoDeDono.Criar(_documento, _autor);
            return true;
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            // O handle antigo (se havia) aponta para o backup da troca ou para um arquivo que nao e' mais o documento: sai.
            Soltar();
            _log($"TRAVA nao retomada em {Path.GetFileName(_documento)}: {e.Message} (a sessao segue sem trava)");
            return false;
        }
    }

    /// <summary>"Salvar como": solta o documento antigo (e apaga o <c>~$</c> dele) e trava <paramref name="novo"/>. <c>false</c> = o novo
    /// nao pode ser travado (o antigo ja' foi solto do mesmo jeito).</summary>
    public bool Mover(string novo)
    {
        Soltar();
        _dono?.Dispose();
        _dono = null;
        _documento = novo;
        // Retomar cria o ~$ do novo; se falhar agora, o proximo Retomar que der certo cria (fix 1 da revisao, MINOR 5).
        if (!Retomar()) return false;
        _log($"TRAVA movida para {Path.GetFileName(novo)}");
        return true;
    }

    /// <summary>Solta o documento e apaga o <c>~$</c> (fechar a janela).</summary>
    public void Dispose()
    {
        Soltar();
        _dono?.Dispose();
        _dono = null;
    }
}
