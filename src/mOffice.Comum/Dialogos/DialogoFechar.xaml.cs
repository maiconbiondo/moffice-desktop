using System.Windows;
using System.Windows.Input;
using mOffice.Nucleo;

namespace mOffice.Comum.Dialogos;

/// <summary>Salvar / Não salvar / Cancelar ao fechar um documento modificado. E' o que a Task 6 injeta em
/// <see cref="JanelaEditor.PerguntarAoFechar"/> no lugar do MessageBox padrao.</summary>
public partial class DialogoFechar : Window
{
    public DialogoFechar(string titulo, string nomeApp)
    {
        InitializeComponent();
        NomeAppTexto.Text = nomeApp;
        PerguntaTexto.Text = $"Deseja salvar as alterações em “{titulo}” antes de fechar?";
    }

    /// <summary>Cancelar e' o PADRAO: fechar a janela pelo X, pelo Esc ou por qualquer caminho que nao
    /// passe por um dos dois botoes de acao nao pode perder o documento do usuario.</summary>
    public DecisaoFechar Decisao { get; private set; } = DecisaoFechar.Cancelar;

    private void Salvar_Click(object sender, RoutedEventArgs e) { Decisao = DecisaoFechar.Salvar; Close(); }
    private void NaoSalvar_Click(object sender, RoutedEventArgs e) { Decisao = DecisaoFechar.NaoSalvar; Close(); }

    private void Cabecalho_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed) DragMove();
    }

    /// <summary>Abre modal sobre <paramref name="dono"/> e devolve a decisao. Sincrono por dentro
    /// (<c>ShowDialog</c> bloqueia), embrulhado em Task porque
    /// <see cref="JanelaEditor.PerguntarAoFechar"/> e' assincrono. Os dois botoes de acao fecham pelo
    /// <c>Close()</c> do handler (vale modal E nao-modal); Cancelar fecha pelo <c>IsCancel</c> do Window.</summary>
    public static Task<DecisaoFechar> PerguntarAsync(Window dono, string titulo, string nomeApp)
    {
        var d = new DialogoFechar(titulo, nomeApp) { Owner = dono };
        d.ShowDialog();
        return Task.FromResult(d.Decisao);
    }

    /// <summary>E' este overload que a Task 6 injeta em <see cref="JanelaEditor.PerguntarAoFechar"/>.</summary>
    public static Task<DecisaoFechar> PerguntarAsync(JanelaEditor dono) =>
        PerguntarAsync(dono, dono.SessaoAtual?.Titulo ?? "", dono.NomeApp);
}
