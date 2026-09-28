using System.Runtime.ExceptionServices;
using System.Windows.Threading;

namespace mOffice.Nucleo.Tests;

public static class JanelaSta
{
    public static bool SessaoRemota => System.Windows.Forms.SystemInformation.TerminalServerSession; // precisa <UseWindowsForms>true</UseWindowsForms> no csproj de testes

    /// <summary>Roda `corpo` numa thread STA com Dispatcher ativo; propaga excecao; falha se passar do timeout.
    /// Se a thread nao terminar a tempo (hang do WPF/WebView2), pede o shutdown do Dispatcher dela pelo
    /// Dispatcher capturado (thread-safe entre threads) antes de reportar falha, para nao deixar a thread
    /// STA (e a janela/WebView2 dela) vazando pelo resto da execucao dos testes.</summary>
    public static void Executar(Func<Task> corpo, TimeSpan? timeout = null)
    {
        Exception? erro = null;
        var dispatcherPronto = new TaskCompletionSource<Dispatcher>();
        var t = new Thread(() =>
        {
            dispatcherPronto.SetResult(Dispatcher.CurrentDispatcher);
            var frame = new DispatcherFrame();
            _ = Dispatcher.CurrentDispatcher.InvokeAsync(async () =>
            {
                try { await corpo(); } catch (Exception ex) { erro = ex; } finally { frame.Continue = false; }
            });
            Dispatcher.PushFrame(frame);
            Dispatcher.CurrentDispatcher.InvokeShutdown();
        });
        t.SetApartmentState(ApartmentState.STA); t.IsBackground = true; t.Start();
        var dispatcher = dispatcherPronto.Task.GetAwaiter().GetResult();
        try
        {
            Assert.True(t.Join(timeout ?? TimeSpan.FromSeconds(90)), "thread STA não terminou no tempo (hang do WPF/WebView2?)");
        }
        finally
        {
            if (t.IsAlive)
            {
                dispatcher.BeginInvokeShutdown(DispatcherPriority.Send);
                t.Join(TimeSpan.FromSeconds(5));
            }
        }
        if (erro is not null) ExceptionDispatchInfo.Capture(erro).Throw();
    }
}
