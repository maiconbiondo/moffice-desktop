namespace mOffice.Nucleo;

/// <summary>
/// Plano 6, Task 7 (SO-05/A4, SO-06, SO-07) — atalhos do Office que o editor web nao tem. A ponte (<c>ponte-desktop.js</c>) captura a
/// tecla em fase de CAPTURA, antes do editor, e manda o nome pelo aviso "atalho". Canal: a PAGINA, nao o WPF — com o WebView2 focado o
/// <c>PreviewKeyDown</c> da janela nao ve' a tecla, e as aceleradoras do navegador estao desligadas (Task 6).
/// </summary>
public partial class JanelaEditor
{
    /// <summary>Atalho do Office tratado (todos tambem chegam aqui, inclusive os que o nucleo ja' tratou sozinho): a casca trata
    /// <see cref="AcaoDeAtalho.Abrir"/> (dialogo do app) e <see cref="AcaoDeAtalho.ProximaJanela"/> (politica de janelas).</summary>
    public event Action<AcaoDeAtalho>? AtalhoRecebido;

    /// <summary>Aviso "atalho" da ponte (so' da sessao ATUAL: a pagina anterior vive ate' o NavigationCompleted da nova). Roda dentro da chamada
    /// sincrona do JS: tudo que abre dialogo ou mexe no editor sai pelo Dispatcher.</summary>
    private void ReceberAtalho(string nome)
    {
        if (AtalhosDoHost.DoAviso(nome) is not { } acao) { _arquivos.Log("ATALHO desconhecido ignorado: " + nome); return; }
        Dispatcher.BeginInvoke(() =>
        {
            if (_fechada) return;
            try
            {
                switch (acao)
                {
                    case AcaoDeAtalho.Fechar: Close(); break;
                    // F12 / Shift+F12: o MESMO caminho do menu Arquivo e do Ctrl+S do editor (asc_Save -> LocalFileSave ->
                    // SalvarPedidoDoEditorAsync): a regra de "Salvar vira Salvar como" (novo, legado, somente leitura, recuperado),
                    // o conflito, o EndSave e a faixa de erro valem iguais — nada paralelo.
                    case AcaoDeAtalho.SalvarComo: _ = PedirSalvarAoEditorAsync(salvarComo: true); break;
                    case AcaoDeAtalho.Salvar: _ = PedirSalvarAoEditorAsync(salvarComo: false); break;
                }
            }
            catch (Exception ex) { _arquivos.Log("ATALHO " + acao + " falhou: " + ex.Message); }
            try { AtalhoRecebido?.Invoke(acao); }
            catch (Exception ex) { _arquivos.Log("ATALHO erro no assinante (" + acao + "): " + ex.Message); }
        });
    }

    private async Task PedirSalvarAoEditorAsync(bool salvarComo)
    {
        try
        {
            var r = (await ExecutarJsAsync(PedirSalvarAoEditor.Replace("SALVARCOMO", salvarComo ? "true" : "false"))).Trim('"');
            if (r != "ok") { UltimoErro = "O editor não aceitou o pedido de salvamento: " + r; _arquivos.Log("ATALHO salvar: " + r); AvisarUsuario(UltimoErro); }
        }
        catch (Exception ex) { UltimoErro = ex.Message; _arquivos.Log("ATALHO salvar falhou: " + ex.Message); AvisarUsuario("Não foi possível salvar: " + ex.Message); }
    }

    /// <summary>Costura de teste: uma tecla com modificadores (CDP <c>Input.dispatchKeyEvent</c>, como o teclado de verdade) no editor.
    /// <paramref name="modificadores"/>: Alt=1, Ctrl=2, Meta=4, Shift=8. Modelo: <c>EnviarAtalhoAsync</c> (colar/copiar/recortar).</summary>
    internal async Task EnviarTeclaParaTesteAsync(string key, string code, int vk, int modificadores)
    {
        var wv = await ProntoAsync();
        Web.Focus();
        string Evento(string tipo) => System.Text.Json.JsonSerializer.Serialize(new
        {
            type = tipo, modifiers = modificadores, windowsVirtualKeyCode = vk, nativeVirtualKeyCode = vk, key, code,
        });
        await wv.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", Evento("rawKeyDown"));
        await wv.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", Evento("keyUp"));
    }
}
