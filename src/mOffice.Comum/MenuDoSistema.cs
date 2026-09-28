using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace mOffice.Comum;

/// <summary>Acrescenta itens ao menu da janela do Windows (clique direito na barra de titulo, ou
/// Alt+Espaco). E' onde moram "Configuracoes do mOffice" e "Sobre" desde que a barra de comandos acima do
/// editor saiu (pedido do usuario em 2026-09-25): sao acoes raras, que nao merecem espaco fixo na tela,
/// e o menu Arquivo do editor nao e' nosso para estender.</summary>
public static class MenuDoSistema
{
    private const int WM_SYSCOMMAND = 0x0112;
    private const uint MF_STRING = 0x0000, MF_SEPARATOR = 0x0800;

    /// <summary>Primeiro id dos nossos itens. O Windows reserva os 4 bits baixos do wParam de WM_SYSCOMMAND
    /// e usa ids a partir de 0xF000 para os comandos do sistema: os nossos sao multiplos de 16 abaixo disso.</summary>
    internal const int IdBase = 0x1010;

    [DllImport("user32.dll")] private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool AppendMenu(IntPtr hMenu, uint uFlags, UIntPtr uIDNewItem, string? lpNewItem);

    /// <summary>Anexa um separador e os itens ao menu da janela. Pode ser chamado antes de a janela existir
    /// na tela: se o HWND ainda nao existe, espera o <c>SourceInitialized</c>.</summary>
    public static void Adicionar(Window janela, IReadOnlyList<(string Texto, Action Acao)> itens)
    {
        if (itens.Count == 0) return;
        void Instalar()
        {
            var hwnd = new WindowInteropHelper(janela).Handle;
            var menu = GetSystemMenu(hwnd, false);
            if (menu == IntPtr.Zero) return;
            AppendMenu(menu, MF_SEPARATOR, UIntPtr.Zero, null);
            var acoes = new Dictionary<int, Action>();
            for (var i = 0; i < itens.Count; i++)
            {
                var id = IdBase + i * 16;
                AppendMenu(menu, MF_STRING, (UIntPtr)(uint)id, itens[i].Texto);
                // Pela FILA do Dispatcher, nunca dentro do hook: as acoes abrem dialogos modais, e um
                // ShowDialog rodando dentro do tratamento de WM_SYSCOMMAND reentra no laco de mensagens
                // com o menu do sistema ainda no meio do seu proprio processamento.
                var acao = itens[i].Acao;
                acoes[id] = () => janela.Dispatcher.BeginInvoke(acao);
            }
            HwndSource.FromHwnd(hwnd)?.AddHook((IntPtr h, int msg, IntPtr w, IntPtr l, ref bool tratado) =>
            {
                if (Tratar(msg, w, acoes)) tratado = true;
                return IntPtr.Zero;
            });
        }
        if (new WindowInteropHelper(janela).Handle != IntPtr.Zero) Instalar();
        else janela.SourceInitialized += (_, _) => Instalar();
    }

    /// <summary>A decisao do hook, separada para teste: e' um WM_SYSCOMMAND com um id NOSSO? Entao roda a
    /// acao e marca como tratado (o Windows nao deve processar um comando que ele desconhece).</summary>
    internal static bool Tratar(int msg, IntPtr wParam, IReadOnlyDictionary<int, Action> acoes)
    {
        if (msg != WM_SYSCOMMAND) return false;
        var id = (int)((long)wParam & 0xFFF0);
        if (!acoes.TryGetValue(id, out var acao)) return false;
        acao();
        return true;
    }
}
