namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 11 (PL-10): medida de memoria no log do host e compactacao do LOH depois de um EndLoad grande.</summary>
public class MedidaDeMemoriaTests
{
    [Fact] // linha ASCII do log, com GC, privados e WS em MiB
    public void Linha_Formato() => Assert.Matches(@"^MEMORIA depois-do-endload gc=\d+ priv=\d+ ws=\d+$", MedidaDeMemoria.Linha("depois-do-endload"));

    [Fact] // abaixo do limiar nao compacta (nada no log); acima, compacta e registra
    public void Compactar_SoAcimaDoLimiar()
    {
        var log = new List<string>();
        MedidaDeMemoria.CompactarSeGrande(1024, log.Add);
        Assert.Empty(log);
        MedidaDeMemoria.CompactarSeGrande(64L * 1024 * 1024, log.Add);
        Assert.Single(log);
    }

    [Fact] // o limiar e' de 20 MiB: um byte a menos nao compacta, o limiar exato compacta
    public void Compactar_LimiarExatoDe20MiB()
    {
        var log = new List<string>();
        MedidaDeMemoria.CompactarSeGrande(MedidaDeMemoria.LimiarBytes - 1, log.Add);
        Assert.Empty(log);
        MedidaDeMemoria.CompactarSeGrande(MedidaDeMemoria.LimiarBytes, log.Add);
        Assert.Single(log);
        Assert.Equal(20L * 1024 * 1024, MedidaDeMemoria.LimiarBytes);
    }

    [Fact] // a linha de compactacao leva a pausa medida e o antes/depois (diagnostico), em ASCII
    public void Compactar_RegistraPausaEAntesDepois()
    {
        var log = new List<string>();
        MedidaDeMemoria.CompactarSeGrande(64L * 1024 * 1024, log.Add);
        Assert.Matches(@"^MEMORIA compactou-loh pausa=\d+ms gc=\d+->\d+ priv=\d+->\d+$", log[0]);
    }

    [Fact] // um log que lanca nao derruba a compactacao (melhor esforco: nunca propaga para a abertura do documento)
    public void Compactar_LogQueLancaNaoPropaga() =>
        MedidaDeMemoria.CompactarSeGrande(64L * 1024 * 1024, _ => throw new InvalidOperationException("log quebrado"));
}
