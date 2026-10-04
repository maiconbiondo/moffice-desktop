using System.Diagnostics;
using System.Globalization;
using System.Runtime;

namespace mOffice.Nucleo;

/// <summary>
/// Plano 6, Task 11 (PL-10): medida de memoria do HOST e compactacao do LOH depois de um documento grande. A auditoria mediu
/// o mPlanilha com 50 mil linhas em 1,17 GB privados no host: o <c>Editor.bin</c> inteiro viaja em base64 dentro do script do
/// <c>EndLoad</c> (bytes -> base64 -> JSON -> script), tres copias grandes no LOH que o GC so' devolve ao sistema operacional
/// quando compacta. Tudo aqui e' diagnostico local (so' o log da janela) e melhor esforco: nunca propaga excecao.
/// </summary>
public static class MedidaDeMemoria
{
    /// <summary>A partir de quantos bytes manipulados o host compacta o LOH (20 MiB).</summary>
    public const long LimiarBytes = 20L * 1024 * 1024;

    private const long Mib = 1024 * 1024;

    /// <summary>Linha ASCII do log: <c>MEMORIA &lt;momento&gt; gc=&lt;MiB&gt; priv=&lt;MiB&gt; ws=&lt;MiB&gt;</c> — heap gerenciado, memoria
    /// privada e working set do processo, em MiB.</summary>
    public static string Linha(string momento)
    {
        var (gc, priv, ws) = Ler();
        return string.Create(CultureInfo.InvariantCulture, $"MEMORIA {momento} gc={gc} priv={priv} ws={ws}");
    }

    private static (long Gc, long Priv, long Ws) Ler()
    {
        using var p = Process.GetCurrentProcess();
        return (GC.GetTotalMemory(false) / Mib, p.PrivateMemorySize64 / Mib, p.WorkingSet64 / Mib);
    }

    /// <summary>Se <paramref name="bytesManipulados"/> chegou ao <see cref="LimiarBytes"/>, compacta o LOH uma vez (GC bloqueante de
    /// geracao 2, compactando) e registra a pausa e o antes/depois. Abaixo do limiar nao faz nada nem escreve no log. Bloqueia o
    /// chamador pelo tempo da coleta: quem chama da thread de interface decide se roda isto fora dela (ver
    /// <c>JanelaEditor.CompactarAposEndLoad</c>).</summary>
    public static void CompactarSeGrande(long bytesManipulados, Action<string> log)
    {
        if (bytesManipulados < LimiarBytes) return;
        try
        {
            var (gc0, priv0, _) = Ler();
            var relogio = Stopwatch.StartNew();
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            relogio.Stop();
            var (gc1, priv1, _) = Ler();
            log(string.Create(CultureInfo.InvariantCulture,
                $"MEMORIA compactou-loh pausa={relogio.ElapsedMilliseconds}ms gc={gc0}->{gc1} priv={priv0}->{priv1}"));
        }
        catch { /* diagnostico: uma falha aqui (ou no log) nunca pode atrapalhar a abertura do documento */ }
    }
}
