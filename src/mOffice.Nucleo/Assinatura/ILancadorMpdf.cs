using System.Diagnostics;

namespace mOffice.Nucleo.Assinatura;

/// <summary>Abre o mPDF para assinar um PDF (Plano 4, Task 5). Interface para os testes usarem um falso: nenhum
/// teste abre o mPDF de verdade (o <c>/assinar</c> real e' passo humano no gate).</summary>
public interface ILancadorMpdf
{
    /// <summary>Lanca o mPDF com <c>/assinar "&lt;pdf&gt;"</c>. Lanca excecao se o processo nao puder ser iniciado.</summary>
    void Assinar(MpdfInstalado mpdf, string pdf);
}

/// <summary>O lancador de producao: <c>mPdf.App.exe /assinar &lt;pdf&gt;</c>, SEM shell e com
/// <see cref="ProcessStartInfo.ArgumentList"/> — o caminho do PDF vai como UM argumento, entao aspas, espacos ou
/// <c>&amp;</c> no nome do arquivo nunca viram outro argumento (nunca se concatena a linha de comando).</summary>
public sealed class LancadorMpdf : ILancadorMpdf
{
    /// <summary>O pedido de processo, separado para teste (o teste confere o pedido sem iniciar nada).</summary>
    public static ProcessStartInfo Pedido(MpdfInstalado mpdf, string pdf)
    {
        var psi = new ProcessStartInfo(mpdf.Exe)
        {
            UseShellExecute = false,
            WorkingDirectory = Path.GetDirectoryName(mpdf.Exe) ?? "",
        };
        psi.ArgumentList.Add("/assinar");
        psi.ArgumentList.Add(pdf);
        return psi;
    }

    public void Assinar(MpdfInstalado mpdf, string pdf)
    {
        using var p = Process.Start(Pedido(mpdf, pdf));
    }
}
