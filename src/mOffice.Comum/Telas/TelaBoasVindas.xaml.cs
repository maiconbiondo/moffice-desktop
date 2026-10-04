using System.Windows;
using System.Windows.Controls;

namespace mOffice.Comum.Telas;

/// <summary>Code-behind SO' do arrastar-e-soltar (o resto e' binding puro pro
/// <see cref="TelaBoasVindasViewModel"/>). O filtro de extensao vive no <see cref="PerfilApp"/>, nao aqui:
/// assim ele e' testavel sem WPF (BoasVindasTests.ArquivoSolto_FiltraPorPerfil).</summary>
public partial class TelaBoasVindas : UserControl
{
    public TelaBoasVindas()
    {
        InitializeComponent();
        DataContextChanged += (_, e) =>
        {
            if (e.OldValue is TelaBoasVindasViewModel antigo) antigo.RolarParaModelos -= RolarParaModelos;
            if (e.NewValue is TelaBoasVindasViewModel novo) novo.RolarParaModelos += RolarParaModelos;
        };
    }

    // O "Criar a partir de modelo" do editor leva a' galeria (Plano 6, Task 9).
    private void RolarParaModelos() => Dispatcher.BeginInvoke(() => PainelModelos.BringIntoView());

    private TelaBoasVindasViewModel? Vm => DataContext as TelaBoasVindasViewModel;

    private static string[] Caminhos(DragEventArgs e) =>
        e.Data.GetDataPresent(DataFormats.FileDrop) ? (string[])e.Data.GetData(DataFormats.FileDrop)! : [];

    private void Tela_DragOver(object sender, DragEventArgs e)
    {
        // Cursor de "copiar" so' quando ha' pelo menos um arquivo que ESTE app abre; qualquer outra coisa
        // fica com o cursor de proibido (o Drop abaixo tambem nao faria nada).
        var aceito = Vm is not null && Vm.Perfil.PrimeiroAceito(Caminhos(e)) is not null;
        e.Effects = aceito ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    // `async void` porque e' um handler de evento roteado — nao ha' quem espere a Task. Por isso NADA aqui
    // pode escapar: AbrirCaminhoAsync ja' converte falha de abertura em MensagemErro, e este try cobre o
    // resto (e.Data.GetData conversa com COM e pode lancar). Uma excecao solta daqui viraria excecao nao
    // tratada do Dispatcher, derrubando o app.
    private async void Tela_Drop(object sender, DragEventArgs e)
    {
        e.Handled = true;
        if (Vm is null) return;
        try
        {
            var caminho = Vm.Perfil.PrimeiroAceito(Caminhos(e));
            if (caminho is null) return;
            await Vm.AbrirCaminhoAsync(caminho);
        }
        catch (Exception ex) { Vm.MensagemErro = $"Não foi possível abrir o arquivo: {ex.Message}"; }
    }
}
