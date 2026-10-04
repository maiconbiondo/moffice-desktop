using System.Text.RegularExpressions;

namespace mOffice.Nucleo;

/// <summary>A troca do arquivo falhou DEPOIS de o original ter sido movido para o backup e o retorno dele tambem falhou
/// (ramo 5b de <see cref="GravacaoNoDestino.Substituir"/>): o original so' existe em <c>*.moffice-bak</c>. A mensagem (pt-BR,
/// para o usuario) e' mostrada pelo chamador: no Ctrl+S/Salvar do editor pela faixa (<c>AvisoAoUsuario</c>); no caminho que
/// relanca (SalvarComoAsync -&gt; fechamento ou API do host) ela viaja na propria excecao (ErroAoFechar / casca).</summary>
internal sealed class ErroGravacaoParcial(string mensagem, Exception interna) : IOException(mensagem, interna);

/// <summary>Como o temporario virou o arquivo do usuario (<see cref="GravacaoNoDestino.Substituir"/>).</summary>
internal enum ResultadoGravacao
{
    /// <summary>Destino novo: movido, nada a preservar.</summary>
    Novo,
    /// <summary>ReplaceFile: ACL explicita, data de criacao, atributos e fluxos do destino preservados.</summary>
    MetadadosPreservados,
    /// <summary>O ReplaceFile falhou e o conteudo entrou por Move: gravado, mas sem preservar os metadados.</summary>
    SemPreservar,
}

/// <summary>Troca o arquivo do usuario pelo temporario ja' validado (SO-14, auditoria 2026-10-01). Com destino
/// existente usa ReplaceFile (<see cref="File.Replace(string, string, string?, bool)"/>), que preserva do DESTINO a ACL
/// explicita, a data de criacao, os atributos e os fluxos alternativos — o File.Move(…, true) de antes trocava o arquivo
/// inteiro.
///
/// INVARIANTE (revisao da Task 1): o usuario nunca fica sem o arquivo. O Replace roda COM um backup explicito ao lado do
/// destino, porque sem backup uma falha no meio dele (ERROR_UNABLE_TO_MOVE_REPLACEMENT 1176: destino ja' removido e o
/// substituto ainda com o nome temporario; 1177: destino so' sobrevive com outro nome) deixava o chamador apagar o
/// temporario — perda total. Falhou: se o destino sumiu, volta do backup; sem backup, o temporario (conteudo novo) vira o
/// destino; so' com o destino de pe' decide entre o Move de contingencia e relancar a excecao.</summary>
internal static class GravacaoNoDestino
{
    internal const string SufixoBackup = ".moffice-bak";

    /// <param name="aposTroca">Fix 1 da revisao da Task 5 do Plano 6: roda logo depois de o conteudo novo estar no destino e ANTES da
    /// limpeza dos backups — o salvar retrava o arquivo e tira o carimbo ali, para uma gravacao de terceiros no intervalo virar conflito
    /// e nunca a nova referencia. Melhor esforco: uma excecao dele so' vai ao log (a troca ja' deu certo).</param>
    public static ResultadoGravacao Substituir(string temporario, string destino, Action<string>? log = null,
        Action<string, string, string>? replace = null, Action? aposTroca = null)
    {
        replace ??= (origem, alvo, backup) => File.Replace(origem, alvo, backup, ignoreMetadataErrors: true);
        if (!File.Exists(destino))
        {
            // Destino ausente: backups ao lado podem ser o UNICO exemplar do original de uma gravacao que morreu no meio
            // (kill durante o ReplaceFile). Nunca se apaga; so' se registra para o suporte achar.
            foreach (var orfao in BackupsDe(destino))
                log?.Invoke($"PONTE ATENCAO: o destino nao existe e ha um backup de gravacao anterior ({orfao}); mantido, pode ser o unico exemplar do original");
            File.Move(temporario, destino);
            Avisar(aposTroca, log);
            return ResultadoGravacao.Novo;
        }
        var bak = BackupLivre(destino); // um backup antigo ao lado nao e' tocado aqui: so' depois de uma troca bem-sucedida
        try
        {
            replace(temporario, destino, bak);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
        {
            // 1) nunca ficar sem o destino
            if (!File.Exists(destino))
            {
                if (File.Exists(bak))
                {
                    try { File.Move(bak, destino); }                              // volta o arquivo velho
                    catch (Exception ex)
                    {
                        // Nem o retorno do backup deu certo: o original so' existe no .moffice-bak e a versao nova no temporario
                        // (mantido pelo chamador enquanto o destino nao existir). O log leva os dois caminhos completos; a mensagem ao
                        // usuario NAO cita o temporario (o proximo salvamento o apaga) — a versao nova continua aberta no editor.
                        log?.Invoke($"PONTE ERRO: nao consegui devolver o original. original={bak} nova={temporario}: {ex.Message}");
                        throw new ErroGravacaoParcial($"O arquivo original ficou em {bak}. A versão nova continua aberta no mOffice: salve de novo ou use Salvar como.", ex);
                    }
                }
                else if (File.Exists(temporario)) File.Move(temporario, destino); // sem backup: o conteudo novo entra no lugar
            }
            if (!File.Exists(destino)) throw;                                     // impossivel recuperar: o chamador preserva o temporario
            // 2) destino de pe': contingencia por Move (conteudo gravado sem preservar metadados) ou relanca
            if (File.Exists(temporario)) File.Move(temporario, destino, overwrite: true);
            Avisar(aposTroca, log);
            ApagarBackup(bak, log);
            LimparBackupsAntigos(destino, log);
            return ResultadoGravacao.SemPreservar;
        }
        Avisar(aposTroca, log);
        ApagarBackup(bak, log);
        LimparBackupsAntigos(destino, log);
        return ResultadoGravacao.MetadadosPreservados;
    }

    private static void Avisar(Action? aposTroca, Action<string>? log)
    {
        try { aposTroca?.Invoke(); }
        catch (Exception e) { log?.Invoke("PONTE aviso depois da troca falhou: " + e.Message); }
    }

    /// <summary>Backups de gravacao (<c>nome.moffice-bak</c>, <c>nome.moffice-bak2</c>...) ao lado do destino.</summary>
    private static string[] BackupsDe(string destino)
    {
        var pasta = Path.GetDirectoryName(Path.GetFullPath(destino));
        if (string.IsNullOrEmpty(pasta) || !Directory.Exists(pasta)) return [];
        // Exatamente "<nome>.moffice-bak" ou "<nome>.moffice-bak<N>": o curinga do GetFiles pegaria "<nome>.moffice-bakup.txt" etc.
        var exato = new Regex("^" + Regex.Escape(Path.GetFileName(destino) + SufixoBackup) + @"\d*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        return Directory.GetFiles(pasta, Path.GetFileName(destino) + SufixoBackup + "*").Where(f => exato.IsMatch(Path.GetFileName(f))).ToArray();
    }

    /// <summary>Depois de uma troca BEM-SUCEDIDA o destino atual supera qualquer backup antigo ao lado: apaga, registrando.
    /// (Uma gravacao que falhou nunca chega aqui, entao nunca descarta o backup anterior.)</summary>
    private static void LimparBackupsAntigos(string destino, Action<string>? log)
    {
        foreach (var velho in BackupsDe(destino))
        {
            log?.Invoke($"PONTE backup antigo de gravacao removido ({velho}): o destino foi gravado");
            ApagarBackup(velho, log);
        }
    }

    private static string BackupLivre(string destino)
    {
        var bak = destino + SufixoBackup;
        for (var i = 2; File.Exists(bak); i++) bak = destino + SufixoBackup + i;
        return bak;
    }

    private static void ApagarBackup(string bak, Action<string>? log)
    {
        try
        {
            if (!File.Exists(bak)) return;
            File.SetAttributes(bak, FileAttributes.Normal);
            File.Delete(bak);
        }
        catch (Exception ex) { log?.Invoke($"PONTE nao apaguei o backup da gravacao ({Path.GetFileName(bak)}): {ex.Message}"); }
    }
}
