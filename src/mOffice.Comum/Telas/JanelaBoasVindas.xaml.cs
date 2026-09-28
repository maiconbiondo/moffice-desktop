using System.Windows;
using System.Windows.Media;

namespace mOffice.Comum.Telas;

public partial class JanelaBoasVindas : Window
{
    /// <param name="chaveLogo">Chave do logo do app (<see cref="PerfilApp.ChaveLogo"/>). Vira o icone da
    /// janela (barra de tarefas / Alt+Tab). A Task 6 pode sobrescrever <c>Icon</c> com o .ico do instalado.</param>
    public JanelaBoasVindas(TelaBoasVindasViewModel vm, string chaveLogo)
    {
        InitializeComponent();
        DataContext = vm;
        Title = vm.Titulo;
        // O logo vetorial e' um DrawingImage do tema mesclado acima; se a chave nao resolver (tema
        // trocado), a janela fica com o icone padrao em vez de lancar.
        if (TryFindResource(chaveLogo) is ImageSource icone) Icon = icone;
    }
}
