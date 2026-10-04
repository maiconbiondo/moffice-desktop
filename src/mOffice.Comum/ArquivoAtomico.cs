using System.Security.Cryptography;
using System.Text;

namespace mOffice.Comum;

/// <summary>Arquivos pequenos de configuracao compartilhados entre processos (os tres apps rodam juntos, e o logoff fecha todos de uma vez):
/// escrita ATOMICA (temporario na mesma pasta + troca) e exclusao mutua entre processos por um mutex nomeado do usuario. Sem isto, um leitor via
/// o arquivo truncado pela escrita do outro, lia "padroes" e gravava os padroes por cima das configuracoes do usuario.</summary>
internal static class ArquivoAtomico
{
    /// <summary>Mutex nomeado do usuario (<c>Local\</c>) para o arquivo: o nome vem do caminho completo, em minusculas.</summary>
    internal static string NomeDoMutex(string arquivo)
    {
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(arquivo).ToLowerInvariant())))[..16];
        return @"Local\mOffice-Config-" + hash;
    }

    /// <summary>Roda <paramref name="acao"/> com o mutex do arquivo na mao (espera ate' 10 s; sem conseguir, segue — melhor do que travar o app).</summary>
    internal static T ComExclusao<T>(string arquivo, Func<T> acao)
    {
        using var m = new Mutex(false, NomeDoMutex(arquivo));
        var tem = false;
        try { tem = m.WaitOne(TimeSpan.FromSeconds(10)); }
        catch (AbandonedMutexException) { tem = true; }   // o dono morreu no meio: a escrita atomica garante que o arquivo esta' inteiro
        try { return acao(); }
        finally { if (tem) m.ReleaseMutex(); }
    }

    /// <summary>Le o texto sem impedir que outro processo troque o arquivo (compartilha leitura, escrita e exclusao).</summary>
    internal static string Ler(string arquivo)
    {
        for (var tentativa = 1; ; tentativa++)
        {
            try
            {
                using var f = new FileStream(arquivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                using var r = new StreamReader(f, Encoding.UTF8);
                return r.ReadToEnd();
            }
            // A troca do File.Replace deixa o nome sem arquivo por um instante: espera e tenta de novo (arquivo que nunca existiu sai na 1a tentativa do chamador, que testa File.Exists).
            catch (IOException) when (tentativa < 20) { Thread.Sleep(5); }
        }
    }

    /// <summary>Grava num temporario ao lado e troca pelo destino (File.Replace/Move): o destino nunca fica pela metade. O temporario some em qualquer falha.</summary>
    internal static void Escrever(string arquivo, string texto)
    {
        var pasta = Path.GetDirectoryName(Path.GetFullPath(arquivo))!;
        Directory.CreateDirectory(pasta);
        var tmp = Path.Combine(pasta, Path.GetFileName(arquivo) + "." + Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            File.WriteAllText(tmp, texto, new UTF8Encoding(false));
            for (var tentativa = 1; ; tentativa++)
            {
                try
                {
                    if (File.Exists(arquivo)) File.Replace(tmp, arquivo, null);
                    else File.Move(tmp, arquivo);
                    return;
                }
                catch (IOException) when (tentativa < 8) { Thread.Sleep(25 * tentativa); }   // um leitor/antivirus segurou o arquivo por um instante
                catch (UnauthorizedAccessException) when (tentativa < 8) { Thread.Sleep(25 * tentativa); }
            }
        }
        finally { try { if (File.Exists(tmp)) File.Delete(tmp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
}
