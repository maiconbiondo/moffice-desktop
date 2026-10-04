using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace mOffice.Comum;

/// <summary>Onde a janela estava, em pixels FISICOS (<c>RECT</c> do <c>WINDOWPLACEMENT</c>: posicao restaurada, valida mesmo com a janela
/// maximizada) e se estava maximizada.</summary>
public sealed record PosicaoSalva(int Esquerda, int Topo, int Direita, int Fundo, bool Maximizada);

/// <summary>Posicao das janelas de documento (Plano 6, Task 7, SO-07): maximizada ou onde o usuario deixou, nunca 1280x800 fixo.</summary>
public static class PosicaoDeJanela
{
    /// <summary>Onde a proxima janela de documento nasce (SO-07): com outra janela do app ativa, em cascata a partir dela (ou
    /// maximizada, se ela estiver); sem, a ultima posicao salva; sem nada salvo, maximizada (null) — nunca 1280x800 fixo.</summary>
    public static PosicaoSalva? Proxima(PosicaoSalva? salva, PosicaoSalva? ativa, int deslocamento)
    {
        if (ativa is { Maximizada: true }) return ativa;
        if (ativa is { } a) return a with { Esquerda = a.Esquerda + deslocamento, Topo = a.Topo + deslocamento, Direita = a.Direita + deslocamento, Fundo = a.Fundo + deslocamento };
        return salva;
    }

    /// <summary>Posicao atual da janela (restaurada + maximizada?). <c>null</c> se ela nao tem HWND ou a API falhar.</summary>
    public static PosicaoSalva? Ler(Window j)
    {
        var hwnd = new WindowInteropHelper(j).Handle;
        if (hwnd == IntPtr.Zero) return null;
        var p = new WINDOWPLACEMENT { length = Marshal.SizeOf<WINDOWPLACEMENT>() };
        if (!GetWindowPlacement(hwnd, ref p)) return null;
        var r = p.rcNormalPosition;
        // Minimizada: o estado a lembrar e' o de antes (WPF_RESTORETOMAXIMIZED); ShowCmd so' diz "minimizada".
        var maximizada = p.showCmd == SW_SHOWMAXIMIZED || (p.showCmd == SW_SHOWMINIMIZED && (p.flags & WPF_RESTORETOMAXIMIZED) != 0);
        return new PosicaoSalva(r.Left, r.Top, r.Right, r.Bottom, maximizada);
    }

    /// <summary>Aplica antes do <c>Show()</c>. <c>null</c> = maximizar. O <c>SetWindowPlacement</c> traz de volta sozinho uma posicao que
    /// ficou fora de qualquer monitor (monitor desligado desde a ultima vez).</summary>
    public static void Aplicar(Window j, PosicaoSalva? p)
    {
        if (p is null) { j.WindowState = WindowState.Maximized; return; }
        var hwnd = new WindowInteropHelper(j).EnsureHandle();
        var w = new WINDOWPLACEMENT
        {
            length = Marshal.SizeOf<WINDOWPLACEMENT>(),
            showCmd = p.Maximizada ? SW_SHOWMAXIMIZED : SW_SHOWNORMAL,
            rcNormalPosition = new RECT { Left = p.Esquerda, Top = p.Topo, Right = p.Direita, Bottom = p.Fundo },
        };
        if (!SetWindowPlacement(hwnd, ref w)) j.WindowState = p.Maximizada ? WindowState.Maximized : WindowState.Normal;
    }

    private const int SW_SHOWNORMAL = 1, SW_SHOWMINIMIZED = 2, SW_SHOWMAXIMIZED = 3;
    private const int WPF_RESTORETOMAXIMIZED = 0x2;

    [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWPLACEMENT
    {
        public int length, flags, showCmd;
        public POINT ptMinPosition, ptMaxPosition;
        public RECT rcNormalPosition;
    }

    [DllImport("user32.dll")] private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
    [DllImport("user32.dll")] private static extern bool SetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
}
