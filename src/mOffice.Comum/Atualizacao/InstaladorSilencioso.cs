using System.Diagnostics;

namespace mOffice.Comum.Atualizacao;

/// <summary>
/// Instalacao SILENCIOSA do instalador ja' verificado + o mutex <c>Global\</c> de presenca que o instalador
/// (<c>AppMutex=</c> do <c>.iss</c>, Task 3) usa para saber que ha' um app da suite rodando. Porte de
/// <c>SilentUpdateInstaller.cs</c> do mPDF. Pecas puras: nenhum teste inicia processo.
/// </summary>
public static class InstaladorSilencioso
{
    /// <summary>Flags do Inno Setup para o <c>/update</c> do RMM: sem janela nenhuma (<c>/VERYSILENT</c>), sem caixas de
    /// mensagem (<c>/SUPPRESSMSGBOXES</c>) e nunca reinicia o Windows (<c>/NORESTART</c>). Sempre com
    /// <see cref="ParametroNaoReabrir"/>: a atualizacao em massa e' invisivel.</summary>
    public const string ArgumentosSilenciosos = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART";

    /// <summary>
    /// Flags do botao "Verificar atualizações" (revisao final do Plano 5, I-1): <c>/SILENT</c> MOSTRA a janela de
    /// progresso (sem perguntas: <c>/SUPPRESSMSGBOXES</c>). Com <c>/VERYSILENT</c> o usuario nao via nada por 1–2 min,
    /// reabria o app pelo Menu Iniciar no meio da instalacao, o app travava arquivos que o [InstallDelete] ja' tinha
    /// apagado, e a pergunta "arquivo em uso" suprimida respondia Abortar: instalacao quebrada. O <c>WizardSilent</c>
    /// do <c>.iss</c> vale para <c>/SILENT</c> e <c>/VERYSILENT</c>, entao o <c>MutexDoApp</c> continua vazio aqui.
    /// </summary>
    public const string ArgumentosComProgresso = "/SILENT /SUPPRESSMSGBOXES /NORESTART";

    /// <summary>Parametro proprio (o Inno ignora o que nao conhece; o <c>[Code]</c> do <c>.iss</c> le') para NAO
    /// reabrir o app ao fim. So' o <c>/update</c> do RMM usa: a atualizacao em massa tem de ser invisivel. O botao
    /// das Configurações NAO passa (o usuario clicou e espera o app voltar).</summary>
    public const string ParametroNaoReabrir = "/NORELAUNCH";

    /// <summary>Parametro proprio (Task 3): <c>/REABRIR=&lt;app&gt;</c> diz ao <c>.iss</c> QUAL dos tres apps reabrir no fim
    /// (<c>DeveReabrir</c> do <c>[Code]</c>, comparacao sem diferenciar maiusculas) — o instalador e um so para a suite.
    /// Sem ele, nada e reaberto (instalacao a mao).</summary>
    public const string ParametroReabrir = "/REABRIR=";

    /// <summary>O <see cref="ProcessStartInfo"/> do instalador VERIFICADO. <c>UseShellExecute=true</c> faz o
    /// manifesto de administrador do instalador disparar o UAC (sem isso o <c>Process.Start</c> falha).</summary>
    /// <paramref name="app"/>: o app a reabrir no fim (<see cref="ParametroReabrir"/>); so vale com <paramref name="reabrir"/>.
    public static ProcessStartInfo MontarStartInfo(ServicoAtualizacao.ArquivoVerificado arquivo, bool reabrir = true, string? app = null) =>
        MontarStartInfo(arquivo.Caminho, reabrir, app);

    /// <summary>Forma crua, so' para os testes das flags: a producao passa sempre um
    /// <see cref="ServicoAtualizacao.ArquivoVerificado"/> (a sobrecarga acima).</summary>
    internal static ProcessStartInfo MontarStartInfo(string caminhoVerificado, bool reabrir, string? app = null) =>
        new(caminhoVerificado)
        {
            Arguments = !reabrir ? ArgumentosSilenciosos + " " + ParametroNaoReabrir
                : string.IsNullOrEmpty(app) ? ArgumentosComProgresso
                : ArgumentosComProgresso + " " + ParametroReabrir + app,
            UseShellExecute = true,
        };

    /// <summary>Abre (ou cria) o mutex <see cref="NomesDeInstancia.MutexDeAtualizacao"/> sem ser dono: o objeto
    /// NOMEADO existe enquanto qualquer handle estiver aberto, e e' essa existencia que o instalador ve'. O chamador
    /// segura a referencia pela vida do processo e a descarta no encerramento. Pode lancar (sem permissao para
    /// <c>Global\</c>): quem chama trata como melhor esforco.</summary>
    public static Mutex AdquirirMutexDoApp() => new(initiallyOwned: false, NomesDeInstancia.MutexDeAtualizacao);
}

/// <summary>
/// "Instalar agora" do botão das Configurações: o instalador verificado fica ARMADO enquanto as janelas fecham pelo
/// fluxo normal (cada documento modificado pergunta se salva). Quando o app vai de fato encerrar, o instalador
/// armado e' iniciado (com reabertura). Qualquer fechamento ABANDONADO (Cancelar, "Salvar como" cancelado, falha ao
/// salvar) ou o usuario voltando a trabalhar (abrir, novo, recuperar, ativar o app) DESARMA — quem chama faz isso.
/// Separado do <see cref="AplicativoOffice"/> para ser testavel.
///
/// No lancamento: outro app da suite aberto desarma (o instalador o fecharia sem perguntar se salva); o hash e'
/// conferido DE NOVO pelo handle que TRAVA o arquivo ate' o <c>Process.Start</c> voltar
/// (<see cref="ServicoAtualizacao.ArquivoVerificado.ReconferirETravar"/>); e a
/// presenca deste processo (o mutex do <c>AppMutex=</c>) e' solta ANTES do <c>Process.Start</c> — o instalador nunca
/// ve' este processo como "app aberto". Se o lancamento falhar, a presenca e' retomada (o app continua aberto).
/// </summary>
public sealed class AtualizacaoAoEncerrar(
    Action<ServicoAtualizacao.ArquivoVerificado> lancar,
    Action<Exception>? logar = null,
    Action? soltarPresenca = null,
    Action? retomarPresenca = null,
    Func<IReadOnlyList<string>>? outrosAppsAbertos = null)
{
    /// <summary>ERROR_CANCELLED: o usuario recusou o UAC.</summary>
    internal const int UacRecusado = 1223;

    private ServicoAtualizacao.ArquivoVerificado? _armado;

    public bool Armada => _armado is not null;

    /// <summary>O caminho do instalador armado (para as mensagens), ou <c>null</c>.</summary>
    public string? Caminho => _armado?.Caminho;

    public void Armar(ServicoAtualizacao.ArquivoVerificado arquivo) => _armado = arquivo;

    /// <summary>Desarma; devolve se estava armada (quem chama so' avisa o usuario uma vez).</summary>
    public bool Desarmar()
    {
        var estava = _armado is not null;
        _armado = null;
        return estava;
    }

    /// <summary>Encerramento: inicia o instalador armado (uma vez so'). Devolve <c>null</c> se nao havia nada ou se
    /// iniciou; a mensagem pt-BR se nao iniciou (o app entao NAO deve encerrar — ver quem chama).</summary>
    public string? IniciarSeArmada()
    {
        if (_armado is not { } a) return null;
        _armado = null;
        // Outro app da suite aberto DEPOIS de armar (as janelas deste levaram tempo para fechar): desarma.
        if (outrosAppsAbertos?.Invoke() is { Count: > 0 } outros) return OutrosAppsDaSuite.MensagemFecharOutros(outros);
        using var trava = a.ReconferirETravar(); // aberta ate' o fim deste metodo, isto e', ate' o lancar voltar
        if (trava is null) return ServicoAtualizacao.MensagemAlteradoDepois;
        soltarPresenca?.Invoke();
        try
        {
            lancar(a);
            return null;
        }
        catch (Exception ex)
        {
            logar?.Invoke(ex);
            retomarPresenca?.Invoke();
            return ex is System.ComponentModel.Win32Exception { NativeErrorCode: UacRecusado }
                ? MensagemCancelada
                : MensagemFalhaAoIniciar(a.Caminho);
        }
    }

    public const string MensagemCancelada = "Atualização cancelada.";

    public static string MensagemFalhaAoIniciar(string caminho) =>
        $"Não foi possível iniciar a atualização. O instalador baixado está em: {caminho}";

    public static string MensagemAdiada(string caminho) =>
        $"Atualização adiada porque um documento não foi fechado. O instalador baixado está em: {caminho}";
}

/// <summary>
/// Os OUTROS apps da suite abertos nesta sessao (Plano 5, Task 2, fix round 1). O "Instalar agora" so' fecha as janelas
/// DESTE processo; o instalador fecharia os outros por fora (<c>CloseApplications</c>) sem perguntar se salvam. Por
/// isso o botao recusa enquanto houver outro app aberto. Detecta pelo mutex de instancia unica de cada app
/// (<see cref="NomesDeInstancia.Mutex"/>, por usuario e sessao): outro USUARIO/sessao nao e' visto — e' o caso do
/// RMM, que roda desacompanhado com os usuarios deslogados, onde vale o CloseApplications + recuperacao.
/// </summary>
public static class OutrosAppsDaSuite
{
    public static readonly IReadOnlyList<PerfilApp> Todos = [PerfilApp.Texto, PerfilApp.Planilha, PerfilApp.Slides];

    /// <summary>Os nomes de exibicao dos outros apps abertos (ex.: ["mPlanilha"]), fora <paramref name="atual"/>.</summary>
    public static IReadOnlyList<string> Abertos(PerfilApp atual) => Abertos(atual, p => NomesDeInstancia.Mutex(p.Nome));

    /// <summary>Costura de teste: o nome do mutex de cada perfil (os testes usam nomes proprios, nunca os reais).</summary>
    internal static IReadOnlyList<string> Abertos(PerfilApp atual, Func<PerfilApp, string> nomeDoMutex) =>
        Todos.Where(p => p.Nome != atual.Nome && Existe(nomeDoMutex(p))).Select(p => p.NomeExibicao).ToList();

    private static bool Existe(string nome)
    {
        try
        {
            if (!Mutex.TryOpenExisting(nome, out var m)) return false;
            m.Dispose();
            return true;
        }
        catch (UnauthorizedAccessException) { return true; } // existe, so' nao podemos abri-lo
    }

    /// <summary>"Feche o mPlanilha antes de atualizar (...)" / "Feche o mPlanilha e o mSlides antes de atualizar (...)".</summary>
    public static string MensagemFecharOutros(IReadOnlyList<string> apps) =>
        apps.Count == 1
            ? $"Feche o {apps[0]} antes de atualizar (documentos abertos nele não seriam salvos)."
            : $"Feche o {string.Join(", o ", apps.Take(apps.Count - 1))} e o {apps[^1]} antes de atualizar (documentos abertos neles não seriam salvos).";
}
