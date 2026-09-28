using mOffice.Nucleo;

namespace mOffice.Nucleo.Tests;

public class LogsTests
{
    [Fact] // nome do arquivo carrega app, data e pid; a pasta e' criada
    public void NovoArquivo_CriaPastaENomeiaComAppEPid()
    {
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        var caminho = Logs.NovoArquivoEm(pasta, "mtexto");
        Assert.True(Directory.Exists(pasta));
        Assert.StartsWith("mtexto-", Path.GetFileName(caminho));
        Assert.Contains($"-{Environment.ProcessId}-", Path.GetFileName(caminho));
    }

    [Fact] // revisao 1 (achado do review): duas janelas do MESMO app no MESMO segundo nao podem cair no
           // mesmo arquivo — e' exatamente o que o contador por processo em NovoArquivoEm resolve.
    public void NovoArquivo_DuasChamadasSeguidas_DevolvemCaminhosDiferentes()
    {
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        var primeiro = Logs.NovoArquivoEm(pasta, "mtexto");
        var segundo = Logs.NovoArquivoEm(pasta, "mtexto");
        Assert.NotEqual(primeiro, segundo);
    }

    [Fact] // so' .log mais velhos que a idade minima sao apagados; outros arquivos ficam
    public void Varrer_ApagaSoLogsVelhos()
    {
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(pasta);
        var velho = Path.Combine(pasta, "mtexto-velho.log"); File.WriteAllText(velho, "x");
        File.SetLastWriteTimeUtc(velho, DateTime.UtcNow.AddDays(-40));
        var novo = Path.Combine(pasta, "mtexto-novo.log"); File.WriteAllText(novo, "x");
        var outro = Path.Combine(pasta, "nao-e-log.txt"); File.WriteAllText(outro, "x");
        File.SetLastWriteTimeUtc(outro, DateTime.UtcNow.AddDays(-40));
        Assert.Equal(1, Logs.Varrer(pasta, TimeSpan.FromDays(30)));
        Assert.False(File.Exists(velho)); Assert.True(File.Exists(novo)); Assert.True(File.Exists(outro));
    }

    [Fact] // pasta inexistente = 0, sem lancar
    public void Varrer_PastaInexistente_DevolveZero()
        => Assert.Equal(0, Logs.Varrer(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N")), TimeSpan.Zero));

    [Fact] // grampo desligado por padrao (o produto nao loga toda chamada da ponte)
    public void GrampoLigado_PadraoFalse() => Assert.False(JanelaEditor.GrampoLigado);

    // ---------------------------------------------------------------------------------------------------
    // Rotacao por TAMANHO (achado da revisao final do Plano 2: a spec §4 sempre disse "rotacao por tamanho",
    // e so' a idade estava implementada — 30 dias de logs de varias janelas cabem num disco cheio).
    // ---------------------------------------------------------------------------------------------------

    /// <summary>Cria <paramref name="quantos"/> .log de <paramref name="bytes"/> cada, um dia mais novo que
    /// o anterior (o primeiro e' o mais ANTIGO), todos dentro da janela de retencao por idade.</summary>
    private static (string pasta, string[] arquivos) SemearLogs(int quantos, int bytes)
    {
        var pasta = Path.Combine(DadosDeTeste.RaizDeTestes, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(pasta);
        var arquivos = new string[quantos];
        for (var i = 0; i < quantos; i++)
        {
            arquivos[i] = Path.Combine(pasta, $"mtexto-{i}.log");
            File.WriteAllBytes(arquivos[i], new byte[bytes]);
            File.SetLastWriteTimeUtc(arquivos[i], DateTime.UtcNow.AddHours(-(quantos - i)));
        }
        return (pasta, arquivos);
    }

    [Fact] // acima do teto: apaga do MAIS ANTIGO para o mais novo, so' ate' caber — e para
    public void Varrer_AcimaDoLimiteDeBytes_ApagaOsMaisAntigosAteCaber()
    {
        var (pasta, arquivos) = SemearLogs(5, 1000);  // 5 000 bytes no total
        try
        {
            // Teto de 2 500: precisa apagar os 3 mais antigos (sobram 2 000) — apagar 2 deixaria 3 000.
            Assert.Equal(3, Logs.Varrer(pasta, TimeSpan.FromDays(30), limiteBytes: 2500));
            Assert.False(File.Exists(arquivos[0])); Assert.False(File.Exists(arquivos[1])); Assert.False(File.Exists(arquivos[2]));
            Assert.True(File.Exists(arquivos[3])); Assert.True(File.Exists(arquivos[4]));  // os mais novos sobrevivem
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [Fact] // dentro do teto: nao apaga nada por tamanho
    public void Varrer_AbaixoDoLimiteDeBytes_NaoApagaNada()
    {
        var (pasta, arquivos) = SemearLogs(3, 1000);
        try
        {
            Assert.Equal(0, Logs.Varrer(pasta, TimeSpan.FromDays(30), limiteBytes: 10_000));
            Assert.All(arquivos, a => Assert.True(File.Exists(a)));
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [Fact] // as duas passadas contam junto: o velho sai pela IDADE e o resto ainda e' cortado pelo TAMANHO
    public void Varrer_IdadeETamanho_SomamNoTotalApagado()
    {
        var (pasta, arquivos) = SemearLogs(4, 1000);
        try
        {
            File.SetLastWriteTimeUtc(arquivos[0], DateTime.UtcNow.AddDays(-40));  // este sai pela idade
            // Sobram 3 000 bytes; teto de 1 500 corta mais 2 (sobra 1 000, o mais novo).
            Assert.Equal(3, Logs.Varrer(pasta, TimeSpan.FromDays(30), limiteBytes: 1500));
            Assert.True(File.Exists(arquivos[3]));
            Assert.False(File.Exists(arquivos[0])); Assert.False(File.Exists(arquivos[1])); Assert.False(File.Exists(arquivos[2]));
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    [Fact] // o teto de producao e' 200 MB e o padrao do parametro e' ele (a subida do app nao passa nada)
    public void RetencaoBytes_EhOTetoPadrao()
    {
        Assert.Equal(200L * 1024 * 1024, Logs.RetencaoBytes);
        var (pasta, arquivos) = SemearLogs(2, 1000);
        try
        {
            Assert.Equal(0, Logs.Varrer(pasta, TimeSpan.FromDays(30)));  // 2 KB < 200 MB: nada some
            Assert.All(arquivos, a => Assert.True(File.Exists(a)));
        }
        finally { try { Directory.Delete(pasta, true); } catch (IOException) { } }
    }

    // ---------------------------------------------------------------------------------------------------
    // Comandos silenciosos do NUCLEO (achado da revisao final do Plano 2: title:mousemove chega ~150x por
    // segundo e cada um virava uma linha no .log da janela).
    // ---------------------------------------------------------------------------------------------------

    [Theory] // os tres de mouse da barra de titulo nao vao para o log
    [InlineData("title:mousemove")]
    [InlineData("title:mousedown")]
    [InlineData("title:mouseup")]
    public void ComandosSilenciosos_ContemOsDeMouseDaBarraDeTitulo(string comando)
        => Assert.Contains(comando, JanelaEditor.ComandosSilenciosos);

    [Theory] // tudo o mais continua sendo registrado — a lista e' fechada de proposito
    [InlineData("create:new")]
    [InlineData("open:recent")]
    [InlineData("editor:event")]
    [InlineData("title:button")]
    [InlineData("uitheme:changed")]
    [InlineData("")]
    public void ComandosSilenciosos_NaoEngoleOResto(string comando)
        => Assert.DoesNotContain(comando, JanelaEditor.ComandosSilenciosos);
}
