using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace mOffice.Nucleo;

/// <summary>Exposto ao JS como chrome.webview.hostObjects.arquivos (async) e .sync.arquivos. Base64 porque o proxy COM nao passa byte[] com seguranca.</summary>
[ClassInterface(ClassInterfaceType.AutoDual)]
[ComVisible(true)]
public sealed class PonteArquivos
{
    private readonly string _log;
    private readonly object _trava = new();
    public PonteArquivos(string arquivoLog) { _log = arquivoLog; }

    public string LerBase64(string caminho) => Convert.ToBase64String(File.ReadAllBytes(caminho));
    public void GravarBase64(string caminho, string base64) => File.WriteAllBytes(caminho, Convert.FromBase64String(base64));
    public string DialogoAbrir(string filtro) { var d = new OpenFileDialog { Filter = filtro }; return d.ShowDialog() == true ? d.FileName : ""; }
    public string DialogoSalvar(string filtro, string nomeSugerido) { var d = new SaveFileDialog { Filter = filtro, FileName = nomeSugerido }; return d.ShowDialog() == true ? d.FileName : ""; }
    public void Log(string linha) { lock (_trava) File.AppendAllText(_log, $"{DateTime.Now:HH:mm:ss.fff} JS {linha}\n"); }
}
