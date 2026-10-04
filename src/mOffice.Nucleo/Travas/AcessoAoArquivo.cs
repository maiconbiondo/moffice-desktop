using System.IO.Compression;

namespace mOffice.Nucleo.Travas;

/// <summary>Como o arquivo pode ser aberto (Plano 6, Task 5 — SO-02/MvP-05).</summary>
public enum EstadoAcesso
{
    /// <summary>Ninguem grava nele e o usuario pode gravar: abre editavel e a sessao trava o original.</summary>
    Livre,
    /// <summary>Atributo "Somente leitura" do arquivo.</summary>
    SomenteLeituraAtributo,
    /// <summary>Sem permissao de escrita (ACL, pasta de rede so' de leitura).</summary>
    SemPermissaoDeEscrita,
    /// <summary>Outro programa/pessoa esta' com o arquivo aberto (violacao de compartilhamento ao abrir para escrita).</summary>
    EmUso,
}

public static class AcessoAoArquivo
{
    private const int ErroCompartilhamento = 32;   // ERROR_SHARING_VIOLATION
    private const int ErroTrava = 33;              // ERROR_LOCK_VIOLATION

    /// <summary>Sonda o arquivo como o Office faz ao abrir: atributo somente leitura; depois abre para leitura E escrita negando
    /// escrita aos outros (nada e' gravado). Arquivo inexistente/caminho invalido = <see cref="EstadoAcesso.Livre"/>: quem reclama e'
    /// a conversao, com a mensagem de sempre. So' o arquivo decide — o <c>~$</c> sozinho nunca manda (sobra de queda).</summary>
    public static EstadoAcesso Sondar(string caminho)
    {
        try
        {
            if (File.GetAttributes(caminho).HasFlag(FileAttributes.ReadOnly)) return EstadoAcesso.SomenteLeituraAtributo;
            using var _ = new FileStream(caminho, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
            return EstadoAcesso.Livre;
        }
        catch (UnauthorizedAccessException) { return EstadoAcesso.SemPermissaoDeEscrita; }
        catch (IOException e) when ((e.HResult & 0xFFFF) is ErroCompartilhamento or ErroTrava) { return EstadoAcesso.EmUso; }
        catch (Exception e) when (e is IOException or ArgumentException or NotSupportedException) { return EstadoAcesso.Livre; }
    }

    /// <summary>Pacote com assinatura digital XML: OOXML com parte em <c>_xmlsignatures/</c> ou ODF com
    /// <c>META-INF/documentsignatures.xml</c>. Lido com <c>FileShare.ReadWrite</c> (o arquivo pode estar em uso por outro).
    /// <c>false</c> em qualquer erro (nao e' zip, sem acesso): so' decide se a faixa aparece.</summary>
    public static bool TemAssinaturaDigital(string pacote)
    {
        try
        {
            using var f = new FileStream(pacote, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var z = new ZipArchive(f, ZipArchiveMode.Read);
            return z.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)
                || string.Equals(e.FullName, "META-INF/documentsignatures.xml", StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception e) when (e is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return false; }
    }
}
