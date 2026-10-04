using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace mOffice.Comum.Dialogos;

public partial class DialogoSobre : Window
{
    /// <summary>Endereco publico do codigo-fonte (AGPL-3.0). Texto puro na tela: o Sobre nao abre navegador
    /// (egresso zero fora da atualizacao).</summary>
    public const string CodigoFonte = "github.com/maiconbiondo/moffice-desktop";

    private readonly string? _pastaLicencas;
    private readonly Action<string> _abrirPasta;

    /// <param name="versao">Versao do app (Assembly/instalador) — a Task 6 passa.</param>
    /// <param name="versaoMotor">Versao do ONLYOFFICE embarcado (<c>mOffice.Nucleo.Motor.Versao</c>).</param>
    public DialogoSobre(PerfilApp perfil, string versao, string versaoMotor)
        : this(perfil, versao, versaoMotor, AppContext.BaseDirectory, null) { }

    /// <param name="pastaBase">Pasta do executavel; os testes passam uma pasta temporaria.</param>
    /// <param name="abrirPasta">Quem abre a pasta de licencas; null = Explorer.</param>
    internal DialogoSobre(PerfilApp perfil, string versao, string versaoMotor, string pastaBase, Action<string>? abrirPasta)
    {
        InitializeComponent();
        if (TryFindResource(perfil.ChaveLogo) is ImageSource logo) LogoImagem.Source = logo;
        TituloTexto.Text = $"{perfil.NomeExibicao} {versao}";
        SubtituloTexto.Text = perfil.Subtitulo;
        // Atribuicao EXIGIDA pela AGPL-3.0 e pela marca da Ascensio: versao do motor, onde estao as
        // modificacoes, a marca e o link do codigo-fonte. Nao editar sem ler docs/licencas.
        AtribuicaoTexto.Text =
            $"Baseado no ONLYOFFICE Desktop Editors {versaoMotor}, desenvolvido originalmente por Ascensio System SIA " +
            "(AGPL-3.0). O mOffice é uma versão modificada, não publicada nem endossada pela Ascensio System SIA; " +
            "as modificações estão listadas em MODIFICACOES-ONLYOFFICE.md. " +
            "ONLYOFFICE is a trademark of Ascensio System SIA.";
        FonteTexto.Text =
            $"O mOffice é software livre, sob a GNU AGPL-3.0. Código-fonte: {CodigoFonte}. " +
            "Os textos completos das licenças e o inventário dos componentes (INVENTARIO.md) estão na pasta de licenças.";
        ListaComponentes.ItemsSource = Atribuicoes.Itens(versao, versaoMotor);

        _pastaLicencas = LocalizarPastaLicencas(pastaBase);
        _abrirPasta = abrirPasta ?? AbrirNoExplorer;
        BotaoLicencas.IsEnabled = _pastaLicencas is not null;
        if (_pastaLicencas is null) BotaoLicencas.ToolTip = "A pasta de licenças não foi encontrada nesta instalação.";
    }

    /// <summary>Abre o Sobre MODAL sobre <paramref name="dono"/>. Este diálogo só pode ser mostrado assim:
    /// ele é sem moldura (<c>WindowStyle=None</c>), não tem X, e o único caminho de saída é o botão
    /// <c>IsCancel</c> — que fecha pelo <c>DialogResult</c> do próprio <see cref="Window"/> e portanto só
    /// funciona sob <c>ShowDialog</c>. Aberto com <c>Show()</c>, o usuário ficaria preso. A Task 6 usa
    /// este método, nunca o construtor direto.</summary>
    public static void Mostrar(Window dono, PerfilApp perfil, string versao, string versaoMotor) =>
        new DialogoSobre(perfil, versao, versaoMotor) { Owner = dono }.ShowDialog();

    /// <summary>A pasta de licencas: <c>&lt;pastaBase&gt;\licencas</c> (instalado: <c>{app}\licencas</c>) ou,
    /// em desenvolvimento, o primeiro <c>docs\licencas</c> subindo a partir de <paramref name="pastaBase"/>
    /// (o exe roda de <c>src\&lt;app&gt;\bin\...</c>). So' vale uma pasta com <c>INVENTARIO.md</c>: uma
    /// <c>docs\licencas</c> qualquer acima de Program Files nao e' a nossa. Null = nenhuma.</summary>
    internal static string? LocalizarPastaLicencas(string pastaBase)
    {
        static bool Valida(string p) => File.Exists(Path.Combine(p, "INVENTARIO.md"));
        var instalada = Path.Combine(pastaBase, "licencas");
        if (Valida(instalada)) return instalada;
        for (var d = new DirectoryInfo(pastaBase); d is not null; d = d.Parent)
        {
            var dev = Path.Combine(d.FullName, "docs", "licencas");
            if (Valida(dev)) return dev;
        }
        return null;
    }

    private static void AbrirNoExplorer(string pasta)
    {
        var psi = new ProcessStartInfo("explorer.exe") { UseShellExecute = false };
        psi.ArgumentList.Add(pasta);
        Process.Start(psi)?.Dispose();
    }

    private void AbrirLicencas_Click(object sender, RoutedEventArgs e)
    {
        if (_pastaLicencas is not null) _abrirPasta(_pastaLicencas);
    }

    private void Cabecalho_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left && e.ButtonState == MouseButtonState.Pressed) DragMove();
    }
}
