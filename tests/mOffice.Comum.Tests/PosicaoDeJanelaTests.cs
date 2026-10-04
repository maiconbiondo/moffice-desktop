using mOffice.Comum;
using mOffice.Nucleo.Tests;

namespace mOffice.Comum.Tests;

public class PosicaoDeJanelaTests
{
    private static readonly PosicaoSalva Salva = new(100, 80, 1100, 780, false);

    [Fact] // primeira janela sem nada salvo: maximizar (null)
    public void SemSalva_SemAtiva_Maximiza() => Assert.Null(PosicaoDeJanela.Proxima(null, null, 26));

    [Fact] // primeira janela com posicao salva: a salva (o SetWindowPlacement do Windows traz de volta se o monitor sumiu)
    public void ComSalva_UsaSalva() => Assert.Equal(Salva, PosicaoDeJanela.Proxima(Salva, null, 26));

    [Fact] // segunda janela com a ativa restaurada: cascata a partir dela
    public void ComAtivaRestaurada_Cascata() =>
        Assert.Equal(new PosicaoSalva(126, 106, 1126, 806, false), PosicaoDeJanela.Proxima(Salva, Salva, 26));

    [Fact] // ativa maximizada: a nova tambem maximiza
    public void ComAtivaMaximizada_Maximiza() =>
        Assert.True(PosicaoDeJanela.Proxima(null, Salva with { Maximizada = true }, 26)!.Maximizada);

    [Fact] // a posicao salva vai e volta pelo config.json (campo novo no fim; config antigo continua legivel)
    public void Configuracoes_Posicao_VaiEVolta()
    {
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        var cfg = new Configuracoes(Path.Combine(pasta, "config.json"));
        Assert.Null(cfg.Posicao("mtexto"));
        cfg.DefinirPosicao("mtexto", Salva);
        cfg.DefinirPosicao("mslides", Salva with { Maximizada = true });
        Assert.Equal(Salva, cfg.Posicao("mtexto"));
        Assert.True(cfg.Posicao("mslides")!.Maximizada);
        Assert.Null(cfg.Posicao("mplanilha"));
        // outras configuracoes seguem independentes
        cfg.FormatoOdf = true; Assert.Equal(Salva, cfg.Posicao("mtexto"));
    }

    [Fact] // Ler/Aplicar de verdade (WINDOWPLACEMENT) numa janela WPF: restaurada, maximizada com a restaurada guardada, e null = maximizar
    public void LerAplicar_IdaEVolta_NumaJanelaReal()
    {
        JanelaSta.Executar(() =>
        {
            var j = new System.Windows.Window { Width = 400, Height = 300, ShowInTaskbar = false };
            var p = new PosicaoSalva(120, 90, 820, 590, false);
            PosicaoDeJanela.Aplicar(j, p);
            j.Show();
            Assert.Equal(p, PosicaoDeJanela.Ler(j));
            j.Close();

            var m = new System.Windows.Window { ShowInTaskbar = false };
            PosicaoDeJanela.Aplicar(m, p with { Maximizada = true });
            m.Show();
            Assert.Equal(System.Windows.WindowState.Maximized, m.WindowState);
            Assert.Equal(p with { Maximizada = true }, PosicaoDeJanela.Ler(m));   // a restaurada fica guardada por baixo
            m.Close();

            var n = new System.Windows.Window { ShowInTaskbar = false };
            PosicaoDeJanela.Aplicar(n, null);                                     // nada salvo: maximiza
            n.Show();
            Assert.True(PosicaoDeJanela.Ler(n)!.Maximizada);
            n.Close();
            return Task.CompletedTask;
        });
    }
}
