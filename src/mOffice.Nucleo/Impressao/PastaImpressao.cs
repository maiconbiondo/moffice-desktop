namespace mOffice.Nucleo.Impressao;

/// <summary>
/// Retencao dos PDFs de impressao (Task 4 do Plano 3) em <see cref="Pastas.Impressao"/>. O PDF de impressao e' o
/// documento inteiro em claro, fora da pasta de recuperacao: cada janela apaga os que gerou ao fechar; o que ela
/// entregou a quem imprime (Task 5) e' apagado por ele depois do envio (<see cref="ApagarComEsperaAsync"/>); e a
/// subida do app varre o que sobrou de uma queda.
///
/// A pasta e' COMPARTILHADA pelos tres apps e por todas as janelas: por isso a janela apaga SO' os arquivos
/// dela (nunca a pasta inteira, que poderia levar junto o PDF que outra janela esta' mandando para a
/// impressora agora), e a varredura so' leva o que tem mais de <see cref="IdadeMinima"/>.
/// </summary>
public static class PastaImpressao
{
    /// <summary>1 dia: nenhuma impressao demora isso, e o lixo de uma queda nao fica semanas no disco.</summary>
    public static readonly TimeSpan IdadeMinima = TimeSpan.FromDays(1);

    /// <summary>Caminho novo e unico para um PDF de impressao (a pasta e' criada).</summary>
    public static string NovoArquivo(string pasta) =>
        Path.Combine(Pastas.Garantir(pasta), Guid.NewGuid().ToString("N") + ".pdf");

    /// <summary>Apaga os <c>*.pdf</c> de <paramref name="pasta"/> mais velhos que <paramref name="idadeMinima"/>.
    /// So' arquivos diretos da pasta, so' <c>.pdf</c>; nunca lanca. Devolve quantos apagou.</summary>
    public static int Varrer(string pasta, TimeSpan idadeMinima)
    {
        var apagados = 0;
        try
        {
            if (!Directory.Exists(pasta)) return 0;
            var limite = DateTime.UtcNow - idadeMinima;
            foreach (var arquivo in Directory.EnumerateFiles(pasta, "*.pdf", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(arquivo) >= limite) continue;
                    File.Delete(arquivo);
                    apagados++;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return apagados;
    }

    /// <summary>Apaga UM PDF de impressao depois do envio (Task 5). O processo do navegador do WebView2 pode segurar
    /// o arquivo por um instante depois de a WebView2 oculta fechar: tenta de novo algumas vezes antes de desistir
    /// (o que sobrar fica para a varredura da subida). Nunca lanca; devolve se o arquivo deixou de existir.</summary>
    public static async Task<bool> ApagarComEsperaAsync(string arquivo, int tentativas = 10, int intervaloMs = 300)
    {
        for (var i = 0; ; i++)
        {
            try { File.Delete(arquivo); return !File.Exists(arquivo); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                if (i + 1 >= tentativas) return false;
            }
            await Task.Delay(intervaloMs);
        }
    }

    /// <summary>Apaga UM PDF de impressao (melhor esforco: se a impressao ainda o segura, fica para a varredura).</summary>
    public static void Apagar(string arquivo)
    {
        try { File.Delete(arquivo); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }
}
