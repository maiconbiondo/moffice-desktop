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
using CheckBox = System.Windows.Controls.CheckBox;
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

    [Fact] // I2 da revisao final do Plano 6: config.json somente leitura — os setters do dialogo nao derrubam o app e o valor fica (memoria)
    public void Configuracoes_SomenteLeitura_SettersNaoLancam()
    {
        JanelaSta.Executar(() =>
        {
            var arq = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "config.json");
            var cfg = new Configuracoes(arq); cfg.Grampo = false;
            File.SetAttributes(arq, FileAttributes.ReadOnly);
            try
            {
                var vm = new DialogoConfiguracoesViewModel(cfg, new TemaService(new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() }));
                vm.NomeUsuario = "Dra. Ana"; vm.PastaPadraoSalvar = @"D:\Docs"; vm.FormatoOpenDocument = true; vm.Grampo = true; vm.ModoTema = ModoTema.Escuro;
                Assert.Equal("Dra. Ana", cfg.NomeUsuario); Assert.True(cfg.FormatoOdf); Assert.True(cfg.Grampo); Assert.Equal(ModoTema.Escuro, cfg.Tema);
            }
            finally { File.SetAttributes(arq, FileAttributes.Normal); }
            return Task.CompletedTask;
        });
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
            ComJanela(d, j => Assert.Equal(3, Descendentes<RadioButton>(j).Count(r => r.GroupName == "Tema")));
            return Task.CompletedTask;
        });
    }

    [Fact] // Plano 6, Task 4: as secoes Geral, Salvar e Modelos aparecem com os textos do plano
    public void Configuracoes_TemSecoesGeralSalvarModelos()
    {
        JanelaSta.Executar(() =>
        {
            var vm = new DialogoConfiguracoesViewModel(Cfg(), new TemaService(new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() }));
            var d = new DialogoConfiguracoes(vm); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                var texto = string.Join("\n", Descendentes<TextBlock>(j).Select(t => t.Text))
                    + "\n" + string.Join("\n", Descendentes<ContentControl>(j).Select(c => c.Content as string));
                foreach (var esperado in new[] { "Geral", "Salvar", "Modelos", "Atualizações", "Nome do autor", "Pasta padrão para salvar",
                    "Formato padrão", "Pasta de modelos", "Aparece nos comentários e nas alterações controladas.",
                    "Formatos do Microsoft Office (.docx, .xlsx, .pptx)", "OpenDocument (.odt, .ods, .odp)" })
                    Assert.Contains(esperado, texto);
                Assert.Equal(2, Descendentes<Button>(j).Count(b => b.Content is "Procurar…"));
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Plano 6, Task 4: a caixa do grampo nao aparece ate clicar em "Mostrar opcoes de diagnostico"
    public void Configuracoes_DiagnosticoEscondidoPorPadrao()
    {
        JanelaSta.Executar(() =>
        {
            var vm = new DialogoConfiguracoesViewModel(Cfg(), new TemaService(new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() }));
            var d = new DialogoConfiguracoes(vm); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                CheckBox Grampo() => Descendentes<CheckBox>(j).Single(c => c.Content is string s && s.StartsWith("Registrar chamadas da ponte"));
                Button Link() => Descendentes<Button>(j).Single(b => b.Content is "Mostrar opções de diagnóstico");
                Assert.False(Grampo().IsVisible); Assert.True(Link().IsVisible);
                Clicar(Link()); j.UpdateLayout();
                Assert.True(Grampo().IsVisible); Assert.False(Link().IsVisible);
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Plano 6, Task 4: o nome digitado, o formato e as pastas gravam no config na hora; trocar o tema do mOffice zera o do editor
    public void Configuracoes_NovosCamposGravam_TemaDoMOfficeZeraOTemaDoEditor()
    {
        JanelaSta.Executar(() =>
        {
            var cfg = Cfg();
            var vm = new DialogoConfiguracoesViewModel(cfg, new TemaService(new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() }));
            vm.NomeUsuario = "  Dra. Ana  "; Assert.Equal("Dra. Ana", cfg.NomeUsuario);
            vm.PastaPadraoSalvar = @"D:\Docs"; Assert.Equal(@"D:\Docs", cfg.PastaPadraoSalvar);
            vm.PastaModelos = @"D:\Mod"; Assert.Equal(@"D:\Mod", cfg.PastaModelos);
            vm.FormatoOpenDocument = true; Assert.True(cfg.FormatoOdf); Assert.False(vm.FormatoOoxml);
            vm.FormatoOoxml = true; Assert.False(cfg.FormatoOdf);
            cfg.TemaEditorId = "theme-night";
            vm.ModoTema = ModoTema.Escuro;
            Assert.Null(cfg.TemaEditorId);
            return Task.CompletedTask;
        });
    }

    private static DialogoConfiguracoesViewModel VmConfig(Configuracoes cfg) =>
        new(cfg, new TemaService(new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() }));

    [Fact] // Fix 1 (revisao da Task 4): o dialogo cabe na area de trabalho e o "Fechar" continua dentro dele em tela pequena
    public void Configuracoes_CabeNaAreaDeTrabalho_FecharVisivelQuandoRestrito()
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoConfiguracoes(VmConfig(Cfg())); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                Assert.True(j.ActualHeight <= SystemParameters.WorkArea.Height, $"{j.ActualHeight} > {SystemParameters.WorkArea.Height}");
            });
            var pequeno = new DialogoConfiguracoes(VmConfig(Cfg())) { MaxHeight = 576 }; JanelaWpf.Aplicar(pequeno, escuro: false);
            ComJanela(pequeno, j =>
            {
                Assert.True(j.ActualHeight <= 576.5, $"altura {j.ActualHeight}");
                var fechar = Descendentes<Button>(j).Single(b => b.Content is "Fechar");
                var caixa = fechar.TransformToAncestor(j).TransformBounds(new Rect(0, 0, fechar.ActualWidth, fechar.ActualHeight));
                Assert.True(caixa.Bottom <= j.ActualHeight + 0.5, $"Fechar em {caixa.Bottom}, janela {j.ActualHeight}");
                Assert.True(Descendentes<System.Windows.Controls.ScrollViewer>(j).Any(s => s.ScrollableHeight > 0), "sem rolagem do corpo");
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Fix 2: o que foi digitado vale mesmo se o dialogo fechar sem perder o foco (Esc/IsCancel); espaco no meio nao some
    public void Configuracoes_DigitadoGravaSemPerderFoco()
    {
        JanelaSta.Executar(() =>
        {
            var cfg = Cfg();
            var d = new DialogoConfiguracoes(VmConfig(cfg)); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                var caixas = Descendentes<System.Windows.Controls.TextBox>(j).ToList();
                Assert.Equal(3, caixas.Count);
                caixas[0].Text = "Dra. Ana ";           // nome do autor
                caixas[1].Text = @"D:\Docs";            // pasta de salvar
                caixas[2].Text = @"D:\Mod";             // pasta de modelos
                Assert.Equal("Dra. Ana", cfg.NomeUsuario);
                Assert.Equal(@"D:\Docs", cfg.PastaPadraoSalvar);
                Assert.Equal(@"D:\Mod", cfg.PastaModelos);
                Assert.Equal("Dra. Ana ", caixas[0].Text); // o texto digitado nao e' reescrito (aparado) no meio da digitacao
                Clicar(Descendentes<Button>(j).Single(b => b.Content is "Fechar"));
            });
            Assert.Equal("Dra. Ana", new Configuracoes(ArquivoDe(cfg)).NomeUsuario);
            return Task.CompletedTask;
        });
    }

    private static string ArquivoDe(Configuracoes cfg) =>
        (string)typeof(Configuracoes).GetField("_arquivo", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!.GetValue(cfg)!;

    [Fact] // minor: rodape avisa que vale para as proximas janelas; subtitulo atualizado
    public void Configuracoes_RodapeESubtitulo()
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoConfiguracoes(VmConfig(Cfg())); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                var t = string.Join("\n", Descendentes<TextBlock>(j).Select(x => x.Text));
                Assert.Contains("Nome, pastas e formato valem para as próximas janelas.", t);
                Assert.DoesNotContain("Aparência e diagnóstico", t);
            });
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
                var botoes = Descendentes<Button>(j).ToDictionary(b => ((string)b.Content).Replace("_", ""));
                Assert.True(botoes.ContainsKey("Salvar") && botoes.ContainsKey("Não salvar") && botoes.ContainsKey("Cancelar"));
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Plano 6, Task 7 (SO-17): Salvar e' o botao padrao (Enter) e os dois de acao tem tecla de acesso (Alt+S, Alt+N); Cancelar continua IsCancel
    public void DialogoFechar_SalvarEhPadrao_EComMnemonicos()
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoFechar("petição.docx", "mTexto"); JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                var botoes = Descendentes<Button>(j).ToList();
                Assert.Equal(["_Salvar", "_Não salvar", "Cancelar"], botoes.Select(b => (string)b.Content).ToList());
                Assert.True(botoes.Single(b => ((string)b.Content).Replace("_", "") == "Salvar").IsDefault);
                Assert.False(botoes.Single(b => ((string)b.Content).Replace("_", "") == "Não salvar").IsDefault);
                Assert.True(botoes.Single(b => (string)b.Content == "Cancelar").IsCancel);
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
                Descendentes<Button>(j).Single(b => ((string)b.Content).Replace("_", "") == rotulo)
                    .RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                Assert.Equal(esperada, d.Decisao);
                // Sob Show() quem fecha e' o Close() do handler; sob ShowDialog() o mesmo Close() tambem
                // devolve o controle a PerguntarAsync.
                Assert.False(d.IsVisible);
            });
            return Task.CompletedTask;
        });
    }

    [Theory] // Plano 6, Task 5 — conflito ao salvar: tres botoes, nasce Cancelar, clique grava a decisao e fecha
    [InlineData("Salvar uma cópia", mOffice.Nucleo.Travas.DecisaoConflito.SalvarCopia)]
    [InlineData("Substituir", mOffice.Nucleo.Travas.DecisaoConflito.Substituir)]
    public void Conflito_CliqueNoBotao_GravaDecisaoEFecha(string rotulo, mOffice.Nucleo.Travas.DecisaoConflito esperada)
    {
        JanelaSta.Executar(() =>
        {
            var d = new DialogoConflito("Arquivo alterado", JanelaEditor.MensagemConflito, "Salvar uma cópia", comSubstituir: true, "mTexto");
            JanelaWpf.Aplicar(d, escuro: false);
            Assert.Equal(mOffice.Nucleo.Travas.DecisaoConflito.Cancelar, d.Decisao);
            ComJanela(d, j =>
            {
                var botoes = Descendentes<Button>(j).Where(b => b.IsVisible).ToDictionary(b => (string)b.Content);
                Assert.Equal(["Salvar uma cópia", "Substituir", "Cancelar"], botoes.Keys.ToList());
                Assert.True(botoes["Salvar uma cópia"].IsDefault);
                Assert.True(botoes["Cancelar"].IsCancel);
                botoes[rotulo].RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                Assert.Equal(esperada, d.Decisao);
                Assert.False(d.IsVisible);
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Plano 6, Task 5 — arquivo em uso ao abrir: so' "Somente leitura" (padrao) e "Cancelar"; o texto e' o do nucleo
    public void EmUso_DoisBotoes_SomenteLeitura()
    {
        JanelaSta.Executar(() =>
        {
            var texto = JanelaEditor.MensagemEmUso("Peticao.docx", "Dr. Colega");
            Assert.Equal("\"Peticao.docx\" está em uso por Dr. Colega. Abrir somente leitura?", texto);
            var d = new DialogoConflito("Arquivo em uso", texto, "Somente leitura", comSubstituir: false, "mTexto");
            JanelaWpf.Aplicar(d, escuro: true);
            ComJanela(d, j =>
            {
                var visiveis = Descendentes<Button>(j).Where(b => b.IsVisible).Select(b => (string)b.Content).ToList();
                Assert.Equal(["Somente leitura", "Cancelar"], visiveis);
                Assert.Contains(Descendentes<TextBlock>(j), t => t.Text == texto);
                Descendentes<Button>(j).Single(b => (string)b.Content == "Somente leitura").RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                Assert.Equal(mOffice.Nucleo.Travas.DecisaoConflito.SalvarCopia, d.Decisao);   // "principal" = abrir somente leitura
            });
            return Task.CompletedTask;
        });
    }

    [Fact] // Plano 6, Task 6 (fix) — link: CANCELAR e' o padrao (Enter nao abre), rolagem com altura limitada, endereco longo quebra
    public void ConfirmarLink_CancelarEPadrao_ComRolagem()
    {
        JanelaSta.Executar(() =>
        {
            var longa = "https://exemplo.com.br/" + string.Concat(Enumerable.Repeat("pasta/arquivo?x=1&y=2/", 40));
            var d = new DialogoConflito("Abrir link", DialogoConflito.ComPontosDeQuebra(JanelaEditor.MensagemConfirmarLink(longa)), "Abrir", comSubstituir: false, "mTexto");
            d.CancelarPadrao();
            JanelaWpf.Aplicar(d, escuro: false);
            ComJanela(d, j =>
            {
                var botoes = Descendentes<Button>(j).Where(b => b.IsVisible).ToDictionary(b => (string)b.Content);
                Assert.Equal(["Abrir", "Cancelar"], botoes.Keys.ToList());
                Assert.False(botoes["Abrir"].IsDefault);
                Assert.True(botoes["Cancelar"].IsDefault);
                Assert.True(botoes["Cancelar"].IsCancel);
                var rolagem = Descendentes<ScrollViewer>(j).First(s => s.MaxHeight is > 0 and < 1000);
                Assert.True(rolagem.MaxHeight <= 240);
                Assert.True(j.ActualHeight < 600, "a janela cresceu sem limite: " + j.ActualHeight);
                Assert.Equal(mOffice.Nucleo.Travas.DecisaoConflito.Cancelar, d.Decisao);
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
