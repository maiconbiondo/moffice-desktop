using System.Windows;
using mOffice.Comum;
using mOffice.Nucleo.Tests;

namespace mOffice.Comum.Tests;

[Collection("janela real")]
public class MenuDoSistemaTests
{
    private const int WM_SYSCOMMAND = 0x0112;

    [Fact] // so' WM_SYSCOMMAND com um id NOSSO dispara a acao; os 4 bits baixos do wParam sao do Windows e sao ignorados
    public void Tratar_SoComandoNossoDisparaAcao()
    {
        var chamados = new List<int>();
        var acoes = new Dictionary<int, Action>
        {
            [MenuDoSistema.IdBase] = () => chamados.Add(1),
            [MenuDoSistema.IdBase + 16] = () => chamados.Add(2),
        };
        Assert.True(MenuDoSistema.Tratar(WM_SYSCOMMAND, (IntPtr)(MenuDoSistema.IdBase + 16 + 3), acoes)); // bits baixos ligados
        Assert.False(MenuDoSistema.Tratar(WM_SYSCOMMAND, (IntPtr)0xF060, acoes));      // SC_CLOSE: do Windows, nao nosso
        Assert.False(MenuDoSistema.Tratar(0x0111, (IntPtr)MenuDoSistema.IdBase, acoes)); // WM_COMMAND: outra mensagem
        Assert.Equal([2], chamados);
    }

    [Fact] // os ids nossos ficam abaixo da faixa reservada do sistema (0xF000) e sao multiplos de 16
    public void IdBase_ForaDaFaixaDoSistema() { Assert.True(MenuDoSistema.IdBase < 0xF000); Assert.Equal(0, MenuDoSistema.IdBase % 16); }

    [Fact] // textos do menu da janela com o nome do app
    public void Textos_ConfiguracoesESobreDoApp() =>
        Assert.Equal(["Configurações do mOffice…", "Sobre o mPlanilha…"], AplicativoOffice.TextosDoMenuDaJanela(PerfilApp.Planilha));

    [Fact] // numa janela real, os itens entram no menu da janela (separador + 2)
    public void Adicionar_AcrescentaItensAoMenuDaJanela()
    {
        JanelaSta.Executar(() =>
        {
            var j = new Window { Width = 200, Height = 100, ShowActivated = false, WindowStartupLocation = WindowStartupLocation.Manual, Left = -2000 };
            MenuDoSistema.Adicionar(j, [("Um", () => { }), ("Dois", () => { })]); // antes do HWND: espera o SourceInitialized
            try
            {
                j.Show();
                var hwnd = new System.Windows.Interop.WindowInteropHelper(j).Handle;
                var textos = Itens(GetSystemMenu(hwnd, false));
                Assert.Equal(["Um", "Dois"], textos.TakeLast(2));
            }
            finally { j.Close(); }
            return Task.CompletedTask;
        });
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern IntPtr GetSystemMenu(IntPtr h, bool r);
    [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern int GetMenuItemCount(IntPtr m);
    [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
    private static extern int GetMenuString(IntPtr m, uint pos, System.Text.StringBuilder s, int max, uint flags);

    private static List<string> Itens(IntPtr menu)
    {
        var r = new List<string>();
        for (var i = 0; i < GetMenuItemCount(menu); i++)
        {
            var sb = new System.Text.StringBuilder(128);
            GetMenuString(menu, (uint)i, sb, 128, 0x400 /* MF_BYPOSITION */);
            if (sb.Length > 0) r.Add(sb.ToString());
        }
        return r;
    }
}
