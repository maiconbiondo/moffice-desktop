using System.Windows;
using System.Windows.Controls;
using mOffice.Comum;
using mOffice.Comum.Dialogos;
using mOffice.Comum.Telas;
using mOffice.Nucleo;
using mOffice.Nucleo.Tests;
// <UseWindowsForms>true</UseWindowsForms> (preciso pro JanelaSta.SessaoRemota) traz System.Windows.Forms
// nos ImplicitUsings: sem estes aliases, Button/RadioButton/Size sao ambiguos entre WinForms e WPF.
using Button = System.Windows.Controls.Button;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using Image = System.Windows.Controls.Image;
using RadioButton = System.Windows.Controls.RadioButton;
using Size = System.Windows.Size;
using ToggleButton = System.Windows.Controls.Primitives.ToggleButton;

namespace mOffice.Comum.Tests;

/// <summary>Janelas REAIS numa thread STA. Nenhuma chama <c>ShowDialog</c> (bloquearia o frame do
/// <see cref="JanelaSta"/> para sempre): a arvore VISUAL — que e' o que <see cref="Descendentes{T}"/>
/// enxerga — so' existe depois de o template ser aplicado, entao seguimos o exemplar do mPDF
/// (tests/mPdf.App.Tests/ShellTests.cs): <c>Show()</c> + <c>UpdateLayout()</c>, e <c>Close()</c> no
/// finally. Nenhuma delas abre o editor (WebView2), por isso rodam tambem em sessao remota (RDP) — sem
/// gate em JanelaSta.SessaoRemota.</summary>
[Collection("janela real")]
public class DialogosTests
{
    private static Configuracoes Cfg() => new(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "config.json"));

    /// <summary>Materializa a arvore visual da janela sem ShowDialog e fecha sempre.</summary>
    private static void ComJanela(Window j, Action<Window> corpo)
    {
        try { j.Show(); j.UpdateLayout(); corpo(j); }
        finally { j.Close(); }
    }

    [Fact] // Sobre: versao, motor, atribuicao e marca aparecem no texto da janela
    public void Sobre_MostraVersaoMotorEAtribuicao()
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoSobre(PerfilApp.Texto, "1.0.0", Motor.Versao);
            JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                var texto = string.Join(" ", Descendentes<TextBlock>(j).Select(t => t.Text));
                Assert.Contains("mTexto 1.0.0", texto); Assert.Contains("9.4.0.129", texto);
                Assert.Contains("ONLYOFFICE is a trademark of Ascensio System SIA", texto);
                Assert.Contains("MODIFICACOES-ONLYOFFICE.md", texto);
                // O logo some em SILENCIO se a chave nao resolver (TryFindResource devolve null, nao lanca):
                // so' uma assercao guarda isso.
                Assert.NotNull(Descendentes<Image>(j).Single().Source);
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Sobre: cada componente do docs/licencas/INVENTARIO.md aparece no texto RENDERIZADO da janela (inclusive os
           // de baixo da lista rolavel), com o link do codigo-fonte
    public void Sobre_ListaTodoOInventario()
    {
        var inventario = LicencasTests.ComponentesDoInventario();
        Assert.True(inventario.Count >= 20, $"so {inventario.Count} componentes lidos do INVENTARIO.md");
        JanelaSta.Executar(() =>
        {
            var d = new DialogoSobre(PerfilApp.Planilha, "1.0.0", Motor.Versao);
            JanelaWpf.Aplicar(d, escuro: true);
            ComJanela(d, j =>
            {
                var textos = Descendentes<TextBlock>(j).Select(t => t.Text).ToList();
                foreach (var c in inventario)
                    Assert.True(textos.Contains(c), $"'{c}' (INVENTARIO.md) nao aparece no Sobre");
                var tudo = string.Join(" ", textos);
                Assert.Contains("github.com/maiconbiondo/moffice-desktop", tudo);
                Assert.Contains("153.0.4234.48", tudo);
                Assert.Contains("74.2", tudo);
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // "Abrir licencas": clique de verdade abre a pasta achada; sem pasta, o botao fica desligado
    public void Sobre_AbrirLicencas()
    {
        var tmp = Directory.CreateTempSubdirectory("moffice-sobre-").FullName;
        try
        {
            var licencas = Directory.CreateDirectory(Path.Combine(tmp, "app", "licencas")).FullName;
            File.WriteAllText(Path.Combine(licencas, "INVENTARIO.md"), "x");
            var semNada = Directory.CreateDirectory(Path.Combine(tmp, "vazio")).FullName;
            JanelaSta.Executar(() =>
            {
                var abertas = new List<string>();
                var d = new DialogoSobre(PerfilApp.Texto, "1.0.0", Motor.Versao, Path.Combine(tmp, "app"), abertas.Add);
                JanelaWpf.Aplicar(d, escuro: false);
                ComJanela(d, j =>
                {
                    var b = Descendentes<Button>(j).Single(x => x.Content is "Abrir licenças");
                    Assert.True(b.IsEnabled);
                    Clicar(b);
                    Assert.Equal([licencas], abertas);
                });
                var d2 = new DialogoSobre(PerfilApp.Texto, "1.0.0", Motor.Versao, semNada, abertas.Add);
                JanelaWpf.Aplicar(d2, escuro: false);
                ComJanela(d2, j => Assert.False(Descendentes<Button>(j).Single(x => x.Content is "Abrir licenças").IsEnabled));
                return Task.CompletedTask;
            });
        }
        finally { Directory.Delete(tmp, true); }
    }

    [Fact] // Configuracoes: mudar o tema no VM grava no config e troca os tokens
    public void Configuracoes_TemaGravaEAplica()
    {
        JanelaSta.Executar(() =>
        {
            var cfg = Cfg();
            var alvo = new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() };
            var tema = new TemaService(alvo);
            var vm = new DialogoConfiguracoesViewModel(cfg, tema);
            vm.ModoTema = ModoTema.Escuro;
            Assert.Equal(ModoTema.Escuro, cfg.Tema); Assert.True(tema.EscuroEfetivo);
            vm.Grampo = true; Assert.True(cfg.Grampo);
            var d = new DialogoConfiguracoes(vm); JanelaWpf.Aplicar(d, escuro: true);
            ComJanela(d, j => Assert.Equal(3, Descendentes<RadioButton>(j).Count()));
            return Task.CompletedTask;
        });
    }

    [Fact] // Fechar: os tres botoes existem e Cancelar e' o padrao
    public void Fechar_TresBotoesECancelarEPadrao()
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoFechar("petição.docx", "mTexto"); JanelaWpf.Aplicar(d, escuro: false);
            // Nasce Cancelar: fechar pelo Esc, pelo IsCancel ou por qualquer caminho que nao passe por um
            // botao de acao nunca pode perder o documento do usuario.
            Assert.Equal(DecisaoFechar.Cancelar, d.Decisao);
            ComJanela(d, j =>
            {
                var botoes = Descendentes<Button>(j).ToDictionary(b => (string)b.Content);
                Assert.True(botoes.ContainsKey("Salvar") && botoes.ContainsKey("Não salvar") && botoes.ContainsKey("Cancelar"));
            });
            return Task.CompletedTask;
        });
    }

    [Theory] // clique DE VERDADE em cada botao de acao: grava a decisao E fecha a janela
    [InlineData("Salvar", DecisaoFechar.Salvar)]
    [InlineData("Não salvar", DecisaoFechar.NaoSalvar)]
    public void Fechar_CliqueNoBotao_GravaDecisaoEFecha(string rotulo, DecisaoFechar esperada)
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoFechar("petição.docx", "mTexto"); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                // RaiseEvent(ClickEvent) percorre o MESMO evento roteado que o Click="…" do XAML assina —
                // e' o caminho de producao, nao um seam so' de teste (o EscolherParaTeste anterior saiu).
                Descendentes<Button>(j).Single(b => (string)b.Content == rotulo)
                    .RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                Assert.Equal(esperada, d.Decisao);
                // Sob Show() quem fecha e' o Close() do handler; sob ShowDialog() o mesmo Close() tambem
                // devolve o controle a PerguntarAsync.
                Assert.False(d.IsVisible);
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Boas-vindas e barra constroem com o tema e resolvem os estilos com chave mOffice.*
    public void BoasVindasEFaixa_ConstroemComTema()
    {
        JanelaSta.Executar(() =>
        {
            var rec = new Recentes(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "r.json"));
            // Um recente ANTES de construir a janela: sem item, o DataTemplate de RECENTES nunca e'
            // instanciado e nada dentro dele e' exercitado — os BasedOn, os bindings de comando por
            // RelativeSource e o {icons:PhosphorIcon x} (Ph.Glyph LANCA num nome fora do mapa de 58).
            rec.Registrar(@"C:\x\exemplo.docx");
            int configuracoes = 0, sobre = 0;
            var vm = new TelaBoasVindasViewModel(PerfilApp.Planilha, rec, () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null),
                configuracoes: () => configuracoes++, sobre: () => sobre++);
            var j = new JanelaBoasVindas(vm, PerfilApp.Planilha.ChaveLogo); JanelaWpf.Aplicar(j, escuro: true);
            ComJanela(j, _ =>
            {
                Assert.Equal("mPlanilha", j.Title);
                // Logo da janela e logo da tela: os dois vem de TryFindResource, que devolve null em vez de
                // lancar — sem estas duas assercoes, trocar a chave apagaria o logo em silencio.
                Assert.NotNull(j.Icon);
                Assert.NotNull(Descendentes<Image>(j).Single().Source);
                // A linha do recente montou de verdade: o ToggleButton de fixar e o botao de remover.
                Assert.NotEmpty(Descendentes<ToggleButton>(j));
                Assert.Contains(Descendentes<Button>(j), b => (b.ToolTip as string) == "Tirar da lista");
                // Rodape: Configuracoes e Sobre moram aqui desde que a barra de comandos saiu (2026-09-25).
                // Clique de VERDADE (evento roteado), nao so' o comando do VM.
                var links = Descendentes<Button>(j).Where(b => b.Content is "Configurações" or "Sobre").ToList();
                Assert.Equal(2, links.Count);
                Assert.All(links, b => Assert.Equal(Visibility.Visible, b.Visibility));
                foreach (var b in links) Clicar(b);
                Assert.Equal(1, configuracoes); Assert.Equal(1, sobre);
            });

            // A faixa de aviso que ficou no lugar da barra: altura ZERO sem aviso (nada de espaco morto acima do
            // editor), visivel com aviso, e o "x" dispensa.
            // Dentro de uma janela de verdade: fora dela o WPF adia os bindings e a medida sai antes do valor.
            var avisoVm = new FaixaAvisoViewModel();
            var faixa = new FaixaAviso { DataContext = avisoVm };
            var jf = new Window { Width = 800, Height = 200, Content = new StackPanel { Children = { faixa } } };
            JanelaWpf.Aplicar(jf, escuro: false);
            ComJanela(jf, _ =>
            {
                Assert.Equal(0, faixa.ActualHeight);
                avisoVm.Aviso = "Não foi possível salvar: teste";
                jf.UpdateLayout();
                Assert.True(faixa.ActualHeight > 0);
                Assert.Contains(Descendentes<TextBlock>(faixa), t => t.Text == "Não foi possível salvar: teste");
                Clicar(Descendentes<Button>(faixa).Single());
                Assert.Null(avisoVm.Aviso);
                jf.UpdateLayout();
                Assert.Equal(0, faixa.ActualHeight);
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Documentos recuperados (Task 3 do Plano 3): secao ACIMA dos recentes, "(não salvo)" so' no documento
           // novo, data local, e Recuperar/Descartar com clique de verdade chegando ao VM; sem itens, a secao some.
           // A pergunta de descarte e' um falso — a padrao e' um MessageBox, que travaria o teste.
    public void BoasVindas_DocumentosRecuperados()
    {
        JanelaSta.Executar(() =>
        {
            var rec = new Recentes(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "r.json"));
            var d1 = new mOffice.Nucleo.Recuperacao.DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false,
                new DateTime(2026, 9, 25, 13, 5, 0, DateTimeKind.Utc));
            var d2 = new mOffice.Nucleo.Recuperacao.DocumentoRecuperavel(@"C:\r\2", "Documento", @"C:\r\2\new.docx", true, DateTime.UtcNow);
            var lista = new List<mOffice.Nucleo.Recuperacao.DocumentoRecuperavel> { d1, d2 };
            mOffice.Nucleo.Recuperacao.DocumentoRecuperavel? recuperado = null, perguntado = null;
            var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null),
                recuperaveis: () => lista, recuperar: d => { recuperado = d; return Task.CompletedTask; },
                descartar: d => lista.Remove(d), confirmarDescarte: d => { perguntado = d; return true; });
            var j = new JanelaBoasVindas(vm, PerfilApp.Texto.ChaveLogo); JanelaWpf.Aplicar(j, escuro: false);
            ComJanela(j, _ =>
            {
                List<string> Visiveis() => Descendentes<TextBlock>(j).Where(t => t.IsVisible).Select(t => t.Text).ToList();
                var textos = Visiveis();
                Assert.Contains("O mOffice fechou sem salvar estes documentos. Recupere para continuar de onde parou.", textos);
                Assert.True(textos.IndexOf("DOCUMENTOS RECUPERADOS") is >= 0 and var i && i < textos.IndexOf("RECENTES"),
                    "a secao de recuperados tem de vir ACIMA dos recentes");
                Assert.Single(textos, t => t == " (não salvo)");
                Assert.Contains(d1.UltimaAlteracaoLocal.ToString("dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture), textos);
                Assert.Equal(2, Descendentes<Button>(j).Count(b => b.Content is "Recuperar"));

                Clicar(Descendentes<Button>(j).First(b => b.Content is "Recuperar"));
                Assert.Same(d1, recuperado);
                j.UpdateLayout(); // Recuperar relista do disco: as linhas sao recriadas
                Clicar(Descendentes<Button>(j).Last(b => b.Content is "Descartar"));
                Assert.Same(d2, perguntado);
                j.UpdateLayout();
                Assert.Single(Descendentes<Button>(j), b => b.Content is "Descartar");
                Clicar(Descendentes<Button>(j).Single(b => b.Content is "Descartar"));
                j.UpdateLayout();
                Assert.DoesNotContain("DOCUMENTOS RECUPERADOS", Visiveis());
                Assert.Contains("RECENTES", Visiveis());
            });
            return Task.CompletedTask;
        });
    }

    /// <summary>Clique como o do mouse: ButtonBase.OnClick (protegido) e' o que dispara o Click E executa o
    /// Command ligado. Levantar so' o ClickEvent roteado NAO executa o comando — um teste assim passaria com o
    /// binding do comando quebrado.</summary>
    internal static void Clicar(System.Windows.Controls.Primitives.ButtonBase b) =>
        typeof(System.Windows.Controls.Primitives.ButtonBase)
            .GetMethod("OnClick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .Invoke(b, null);

    internal static IEnumerable<T> Descendentes<T>(DependencyObject raiz) where T : DependencyObject
    {
        for (var i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(raiz); i++)
        {
            var filho = System.Windows.Media.VisualTreeHelper.GetChild(raiz, i);
            if (filho is T t) yield return t;
            foreach (var n in Descendentes<T>(filho)) yield return n;
        }
    }
}
