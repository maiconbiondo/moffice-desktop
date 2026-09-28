using System.Windows;
using System.Windows.Input;

namespace mOffice.Comum.Dialogos;

public partial class DialogoConfiguracoes : Window
{
    public DialogoConfiguracoes(DialogoConfiguracoesViewModel vm)
    {
        InitializeComponent();
        DataContext = vm;
        // "Instalar agora" com o instalador pronto: o dialogo fecha e o app segue o fechamento (AplicativoOffice).
        // So' existe sob ShowDialog (ver Mostrar), entao DialogResult e' valido aqui.
        vm.PedirFechar += () => Dispatcher.BeginInvoke(() => { if (IsVisible) DialogResult = true; });
        // Fechado no meio do download: cancela (senao ~1 GB seguiria baixando sem ninguem para instalar).
        Closed += (_, _) => vm.CancelarDownload();
    }

    /// <summary>Abre as Configurações MODAL sobre <paramref name="dono"/>. Este diálogo só pode ser
    /// mostrado assim: ele é sem moldura (<c>WindowStyle=None</c>), não tem X, e o único caminho de saída é
    /// o botão <c>IsCancel</c> — que fecha pelo <c>DialogResult</c> do próprio <see cref="Window"/> e
    /// portanto só funciona sob <c>ShowDialog</c>. Aberto com <c>Show()</c>, o usuário ficaria preso. A
    /// Task 6 usa este método, nunca o construtor direto.</summary>
    public static void Mostrar(Window dono, DialogoConfiguracoesViewModel vm) =>
        new DialogoConfiguracoes(vm) { Owner = dono }.ShowDialog();

    private void Cabecalho_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed) DragMove();
    }
}
