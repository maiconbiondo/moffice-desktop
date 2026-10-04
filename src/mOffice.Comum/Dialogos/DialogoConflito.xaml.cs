using System.Windows;
using System.Windows.Input;
using mOffice.Nucleo;
using mOffice.Nucleo.Travas;

namespace mOffice.Comum.Dialogos;

/// <summary>Plano 6, Task 5 — as perguntas da trava de arquivo no estilo da casca (em vez dos <c>MessageBox</c> padrao do nucleo):
/// <see cref="PerguntarAsync(JanelaEditor)"/> (conflito ao salvar: Salvar uma cópia / Substituir / Cancelar) e
/// <see cref="PerguntarEmUsoAsync(JanelaEditor, string, string?)"/> (arquivo em uso ao abrir: Somente leitura / Cancelar).</summary>
public partial class DialogoConflito : Window
{
    private readonly bool _emUso;

    public DialogoConflito(string titulo, string pergunta, string principal, bool comSubstituir, string nomeApp)
    {
        InitializeComponent();
        TituloTexto.Text = titulo;
        NomeAppTexto.Text = nomeApp;
        PerguntaTexto.Text = pergunta;
        BotaoPrincipal.Content = principal;
        BotaoSubstituir.Visibility = comSubstituir ? Visibility.Visible : Visibility.Collapsed;
        _emUso = !comSubstituir;
    }

    /// <summary>Cancelar e' o PADRAO: o X, o Esc ou qualquer caminho fora dos botoes de acao nao grava nem abre nada.</summary>
    public DecisaoConflito Decisao { get; private set; } = DecisaoConflito.Cancelar;

    /// <summary>Enter = Cancelar: tira o "padrao" do botao principal e poe o foco/padrao em Cancelar.</summary>
    internal void CancelarPadrao()
    {
        BotaoPrincipal.IsDefault = false;
        BotaoCancelar.IsDefault = true;
        Loaded += (_, _) => BotaoCancelar.Focus();
    }

    /// <summary>Espaco de largura zero depois de / ? & = . para o TextBlock poder quebrar um endereco sem espacos.</summary>
    internal static string ComPontosDeQuebra(string texto) =>
        string.Concat(texto.Select(c => "/?&=.".Contains(c) ? c + "\u200B" : c.ToString()));

    private void Principal_Click(object sender, RoutedEventArgs e) { Decisao = DecisaoConflito.SalvarCopia; Close(); }
    private void Substituir_Click(object sender, RoutedEventArgs e) { if (!_emUso) Decisao = DecisaoConflito.Substituir; Close(); }

    private void Cabecalho_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed) DragMove();
    }

    /// <summary>Conflito ao salvar por cima: e' o que a casca injeta em <see cref="JanelaEditor.PerguntarConflito"/>.</summary>
    public static Task<DecisaoConflito> PerguntarAsync(JanelaEditor dono)
    {
        var d = new DialogoConflito("Arquivo alterado", JanelaEditor.MensagemConflito + " O que deseja fazer com as suas alterações?",
            "Salvar uma cópia", comSubstituir: true, dono.NomeApp) { Owner = dono };
        d.ShowDialog();
        return Task.FromResult(d.Decisao);
    }

    /// <summary>Link do documento (Plano 6, Task 6): e' o que a casca injeta em <see cref="JanelaEditor.ConfirmarLink"/>. <c>true</c> = abrir fora do mOffice.
    /// <b>Cancelar e' o padrao</b>: Enter, Esc e o X nao abrem nada; so' o clique em "Abrir" abre. O endereco aparece numa area de altura
    /// limitada com rolagem, com pontos de quebra depois de / ? & = . (endereco longo nao estoura a janela).</summary>
    public static bool ConfirmarLink(JanelaEditor dono, string uri)
    {
        var d = new DialogoConflito("Abrir link", ComPontosDeQuebra(JanelaEditor.MensagemConfirmarLink(uri)), "Abrir", comSubstituir: false, dono.NomeApp) { Owner = dono };
        d.CancelarPadrao();
        d.ShowDialog();
        return d.Decisao == DecisaoConflito.SalvarCopia;
    }

    /// <summary>Vinculo externo da planilha (fix 1 da revisao da Task 8): e' o que a casca injeta em <see cref="JanelaEditor.ConfirmarVinculo"/>.
    /// <c>true</c> = permitir a leitura. <b>Cancelar e' o padrao</b> (Enter, Esc e o X negam); o caminho completo vem na mensagem.</summary>
    public static bool ConfirmarVinculo(JanelaEditor dono, string mensagem)
    {
        var d = new DialogoConflito("Vínculo externo", ComPontosDeQuebra(mensagem), "Permitir", comSubstituir: false, dono.NomeApp) { Owner = dono };
        d.CancelarPadrao();
        d.ShowDialog();
        return d.Decisao == DecisaoConflito.SalvarCopia;
    }

    /// <summary>Arquivo em uso ao abrir: e' o que a casca injeta em <see cref="JanelaEditor.PerguntarEmUso"/>. <c>true</c> = somente leitura.</summary>
    public static Task<bool> PerguntarEmUsoAsync(JanelaEditor dono, string arquivo, string? donoDoArquivo)
    {
        var d = new DialogoConflito("Arquivo em uso", JanelaEditor.MensagemEmUso(arquivo, donoDoArquivo),
            "Somente leitura", comSubstituir: false, dono.NomeApp) { Owner = dono };
        d.ShowDialog();
        return Task.FromResult(d.Decisao == DecisaoConflito.SalvarCopia);
    }
}
