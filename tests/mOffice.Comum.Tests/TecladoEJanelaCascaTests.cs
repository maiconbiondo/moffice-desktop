namespace mOffice.Comum.Tests;

/// <summary>Plano 6, Task 7: a fiacao da casca (AplicativoOffice) que so' existe com o app inteiro de pe' — varredura do fonte, como os
/// demais testes de fiacao desta suite. O comportamento de cada peca tem teste proprio (PosicaoDeJanela, GerenteDeJanelas, Atalhos).</summary>
public class TecladoEJanelaCascaTests
{
    private static string App() => File.ReadAllText(Path.Combine(AplicativoOfficeTests.RaizRepo(), "src", "mOffice.Comum", "AplicativoOffice.cs"));

    [Fact] // A14: abrir o app de novo com documento aberto mostra a tela inicial, pelo MESMO pedido de ativacao (?ativar) da segunda instancia
    public void AtivacaoComDocumentoAberto_MostraATelaInicial()
    {
        var app = App();
        var corpo = app[app.IndexOf("private void AoPedirAtivacao(", StringComparison.Ordinal)..];
        corpo = corpo[..corpo.IndexOf("\n    }", StringComparison.Ordinal)];
        Assert.Contains("if (_janelas.Abertas > 0) { Logar(\"ATIVACAO com documento aberto: mostra a tela inicial\"); MostrarBoasVindas(); return; }", corpo);
        Assert.DoesNotContain("(ignorada)", corpo);
        Assert.Contains("_instancia.AtivacaoRecebida += () => Dispatcher.BeginInvoke(AoPedirAtivacao);", app);
    }

    [Fact] // Ctrl+O abre o dialogo do app; Ctrl+F6 ativa a proxima janela da politica
    public void AtalhosDaCasca_AbrirEProximaJanela()
    {
        var app = App();
        Assert.Contains("if (a == AcaoDeAtalho.Abrir) _ = EscolherEAbrirAsync(j);", app);
        Assert.Contains("_janelas.ProximaDepois(adaptada)?.Ativar();", app);
    }

    [Fact] // SO-07: a posicao e' aplicada ANTES do Show() e salva no Closing que nao foi cancelado
    public void PosicaoDaJanela_AplicadaAntesDoShow_ESalvaNoFechar()
    {
        var app = App();
        var aplicar = app.IndexOf("PosicaoDeJanela.Aplicar(j, PosicaoDeJanela.Proxima(_config.Posicao(_perfil.Nome)", StringComparison.Ordinal);
        var show = app.IndexOf("        j.Show();\n", aplicar, StringComparison.Ordinal);
        if (show < 0) show = app.IndexOf("        j.Show();\r\n", aplicar, StringComparison.Ordinal);
        Assert.True(aplicar > 0 && show > aplicar, "a posicao precisa ser aplicada antes de j.Show()");
        Assert.Contains("if (e.Cancel || j.EmTelaCheia) return;", app);
        Assert.Contains("_config.DefinirPosicao(_perfil.Nome, p);", app);
    }
}
