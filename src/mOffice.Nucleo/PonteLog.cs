using System.Runtime.InteropServices;

namespace mOffice.Nucleo;

/// <summary>
/// Unico objeto de host exposto ao QUADRO do editor alem de <see cref="PonteDesktop"/> (achado 1 da revisao
/// 2026-09-17): o editor roda documento de terceiros (parsing de DOCX/ODT/etc arbitrario) e nao pode ter
/// acesso a <see cref="PonteArquivos"/> (LerBase64/GravarBase64/dialogos = I/O de arquivo arbitrario). O
/// quadro so precisa poder logar; <c>arquivos</c> fica so na janela (usado pela pagina de teste do PoC,
/// que roda no NIVEL DA JANELA, nao no quadro do editor).
/// </summary>
[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]
public sealed class PonteLog
{
    private readonly string _log;
    private readonly object _trava = new();

    public PonteLog(string arquivoLog) { _log = arquivoLog; }

    public void Log(string linha)
    {
        lock (_trava) File.AppendAllText(_log, $"{DateTime.Now:HH:mm:ss.fff} JS {linha}\n");
    }
}
