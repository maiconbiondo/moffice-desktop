using mOffice.Nucleo.Impressao;

namespace mOffice.Comum;

/// <summary>
/// As impressoes em andamento do app e o encerramento que espera por elas (fix round 1 da Task 5 do Plano 3).
///
/// O problema: a impressao roda numa WebView2 oculta na thread de interface do app. Fechar a ULTIMA janela de
/// documento no meio do envio fazia o gerente de janelas encerrar o app na hora (<c>Shutdown</c>) — o Dispatcher
/// morria, a WebView2 oculta junto, e o trabalho se perdia sem aviso. Agora o encerramento pedido pela politica de
/// janelas (<see cref="PedirEncerramento"/>) ESPERA as impressoes em curso — o app fica vivo, sem janela nenhuma —
/// ate' <see cref="LimiteDeEspera"/>; so' entao encerra. Durante a espera nada abre sozinho; se o USUARIO abrir o
/// app de novo (arquivo pela instancia unica, ou o app sem arquivo), a espera e' cancelada (<see cref="Cancelar"/>)
/// e o app segue normalmente.
///
/// Fix round 4 da revisao final do Plano 3 (defeito da era do Plano 2): o encerramento tambem espera os SALVAMENTOS em
/// curso (<see cref="AcompanharAsync"/>). "Ctrl+S e fechar a ultima janela" encerrava o app com o salvamento no meio —
/// as continuacoes dele iam para um Dispatcher morto, o x2t terminava deixando <c>&lt;doc&gt;.moffice-tmp</c> ao lado
/// e o arquivo do usuario ficava VELHO, em silencio. As contagens (e as mensagens de log) sao as mesmas das impressoes:
/// "trabalho em curso" e' impressao ou salvamento.
///
/// Tambem e' o dono do PDF depois do <c>ImprimirPedido</c>: <see cref="ImprimirAsync"/> apaga o PDF no fim, com
/// sucesso, cancelamento ou erro. Sem WPF: testavel com envios falsos. Chamado na thread de interface; as
/// contagens ficam sob trava para os testes (que nao tem Dispatcher) poderem chamar de qualquer thread.
/// </summary>
internal sealed class ImpressoesEmCurso
{
    /// <summary>Quanto o encerramento espera as impressoes em curso: generoso — um documento grande num driver lento
    /// leva minutos — mas finito: um envio travado nao pode deixar um processo invisivel vivo para sempre.</summary>
    public static readonly TimeSpan LimiteDeEspera = TimeSpan.FromMinutes(3);

    private readonly Action _encerrar;
    private readonly Action<string> _log;
    private readonly TimeSpan _limite;
    private readonly Func<TimeSpan, Task> _esperar;
    private readonly object _trava = new();
    private int _emCurso;
    private int _geracao;
    private bool _drenando;

    /// <param name="encerrar">Encerra o app de verdade (Shutdown).</param>
    /// <param name="esperar">Costura de teste para o relogio do limite (padrao <see cref="Task.Delay(TimeSpan)"/>).</param>
    public ImpressoesEmCurso(Action encerrar, Action<string> log, TimeSpan? limite = null, Func<TimeSpan, Task>? esperar = null)
    {
        _encerrar = encerrar;
        _log = log;
        _limite = limite ?? LimiteDeEspera;
        _esperar = esperar ?? (t => Task.Delay(t));
    }

    public int EmCurso { get { lock (_trava) return _emCurso; } }

    /// <summary>O ultimo documento fechou e o app esta' esperando impressoes em curso para encerrar.</summary>
    public bool Drenando { get { lock (_trava) return _drenando; } }

    /// <summary>
    /// Um envio completo: conta como "em curso" do inicio ao fim, mostra o resultado em <paramref name="avisar"/>
    /// ("Enviado para …" ou o erro, sempre em pt-BR; cancelamento do dialogo do Windows nao mostra nada) e APAGA o
    /// PDF no fim (com espera, o navegador pode segura-lo um instante; o que sobrar fica para a varredura de 1 dia).
    /// Se o app estava esperando para encerrar e esta era a ultima, encerra DEPOIS de apagar. Nunca lanca.
    /// </summary>
    public async Task ImprimirAsync(string pdf, Func<Task<ImpressaoEnviada?>> imprimir, Action<string> avisar)
    {
        lock (_trava) _emCurso++;
        try
        {
            var enviada = await imprimir();
            if (enviada is null) _log("IMPRESSAO cancelada no dialogo do sistema");
            else
            {
                _log($"IMPRESSAO enviada para {enviada.Impressora} (offline={enviada.Offline})");
                Avisar(avisar, AplicativoOffice.MensagemEnviadoParaImpressora(enviada.Impressora, enviada.Offline));
            }
        }
        catch (Exception ex)
        {
            _log("IMPRESSAO falhou: " + ex);
            // So' mensagem pt-BR vai para a faixa: as do nucleo (InvalidOperationException) ja' sao; o resto fica no log.
            Avisar(avisar, AplicativoOffice.MensagemErroAoEnviar(ex is InvalidOperationException ? ex.Message : MensagemInesperada));
        }
        finally
        {
            try
            {
                if (!await PastaImpressao.ApagarComEsperaAsync(pdf))
                    _log("IMPRESSAO PDF de impressao ainda em uso (fica para a varredura): " + Path.GetFileName(pdf));
            }
            catch (Exception ex) { _log("IMPRESSAO falha ao apagar o PDF: " + ex.Message); }
            Terminou("impressao");
        }
    }

    /// <summary>
    /// Um trabalho em curso que nao e' impressao (fix round 4: um SALVAMENTO de documento — <c>JanelaEditor.SalvamentoIniciado</c>):
    /// conta como "em curso" desde a chamada (SINCRONO: antes de qualquer await, para o encerramento pedido logo depois
    /// ja' o ver) ate' <paramref name="trabalho"/> terminar, com sucesso ou nao. Se o app esperava para encerrar e este
    /// era o ultimo, encerra. Nunca lanca.
    /// </summary>
    public async Task AcompanharAsync(Task trabalho, string descricao)
    {
        lock (_trava) _emCurso++;
        try { await trabalho; }
        catch (Exception ex) { _log(descricao.ToUpperInvariant() + " em curso terminou com erro: " + ex.Message); }
        finally { Terminou(descricao); }
    }

    private void Terminou(string descricao)
    {
        bool encerrarAgora;
        lock (_trava)
        {
            _emCurso--;
            encerrarAgora = _emCurso == 0 && _drenando;
            if (encerrarAgora) { _drenando = false; _geracao++; }
        }
        if (encerrarAgora)
        {
            _log($"ENCERRAMENTO o ultimo trabalho em curso ({descricao}) terminou: encerrando");
            Encerrar();
        }
    }

    /// <summary>Texto de uma falha sem mensagem pt-BR propria (o detalhe fica no log).</summary>
    internal const string MensagemInesperada = "ocorreu um erro inesperado (os detalhes ficaram no registro do mOffice).";

    /// <summary>
    /// A politica de janelas quer encerrar (ultima janela fechada). Sem impressao em curso: encerra na hora. Com:
    /// espera todas terminarem (<see cref="ImprimirAsync"/> encerra depois da ultima) ou <see cref="LimiteDeEspera"/>
    /// — o que vier primeiro.
    /// </summary>
    public void PedirEncerramento()
    {
        int geracao;
        lock (_trava)
        {
            if (_emCurso == 0) geracao = -1;
            else if (_drenando) return;
            else { _drenando = true; geracao = ++_geracao; }
        }
        if (geracao < 0) { Encerrar(); return; }
        _log($"ENCERRAMENTO esperando {EmCurso} impressao(oes) em curso (ate' {_limite.TotalMinutes:0.#} min)");
        _ = EncerrarNoLimiteAsync(geracao);
    }

    private async Task EncerrarNoLimiteAsync(int geracao)
    {
        await _esperar(_limite);
        lock (_trava)
        {
            if (!_drenando || geracao != _geracao) return; // terminou antes, ou o usuario reabriu o app
            _drenando = false;
            _geracao++;
        }
        _log($"ENCERRAMENTO limite de {_limite.TotalMinutes:0.#} min com {EmCurso} impressao(oes) ainda em curso: encerrando assim mesmo");
        Encerrar();
    }

    /// <summary>O usuario voltou a usar o app durante a espera (abriu um arquivo, um documento novo, ou o app pelo
    /// menu Iniciar): a espera acaba e o app segue. Devolve se havia espera.</summary>
    public bool Cancelar()
    {
        lock (_trava)
        {
            if (!_drenando) return false;
            _drenando = false;
            _geracao++;
        }
        _log("ENCERRAMENTO cancelado: o app foi usado de novo durante a espera das impressoes");
        return true;
    }

    private void Encerrar()
    {
        try { _encerrar(); }
        catch (Exception ex) { _log("ENCERRAMENTO falhou: " + ex); }
    }

    private void Avisar(Action<string> avisar, string mensagem)
    {
        try { avisar(mensagem); }
        catch (Exception ex) { _log("IMPRESSAO falha ao mostrar o aviso: " + ex.Message); }
    }
}
