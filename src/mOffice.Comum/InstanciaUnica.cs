using System.IO.Pipes;
using System.Text;

namespace mOffice.Comum;

/// Instancia unica por app (mutex nomeado + pipe nomeado). Copia traduzida de SingleInstanceService e
/// SingleInstanceLaunchGate do mPDF (src\mPdf.App\Services\ do repositorio do mPDF); onde o
/// esboco da task e o arquivo do mPDF divergirem em detalhe de implementacao, o arquivo do mPDF venceu.
///
/// Protocolo do pipe: 1 caminho ABSOLUTO por conexao (um caminho RELATIVO chegando pelo pipe e' tratado
/// como malformado e ignorado, nunca repassado adiante -- ver Path.IsPathRooted no loop principal),
/// UTF-8, terminado em '\n' (CRLF tolerado -- o '\r' final e' cortado). 1 caminho por CONEXAO (nao um
/// pipe persistente com varias linhas) -- mais simples, e o caso de uso real (NomesDeInstancia + o
/// App.xaml.cs de cada app) so precisa mandar 1 argv[0] por processo secundario que nasce e morre.
///
/// Endurecimento do protocolo (mesma disciplina do mPDF):
///  - Teto de tamanho de linha (<see cref="MaxLineLengthBytes"/>, 32KB -- nenhum caminho de arquivo
///    legitimo chega perto disso): uma linha que ultrapassa o teto SEM nunca achar '\n' e' descartada
///    (`LerLinhaComTetoAsync` devolve null), nunca cresce sem limite na memoria do processo primario.
///  - Linha vazia/so espaco OU nao-absoluta e' IGNORADA -- nunca dispara CaminhoRecebido com lixo nem
///    com um caminho que nao daria pra resolver de forma previsivel (relativo a QUAL diretorio de
///    trabalho? o do processo secundario ja morto).
///  - Crash do cliente NO MEIO do envio (desconexao abrupta, sem '\n') derruba a stream com
///    IOException -- capturada no loop principal, que volta a escutar a PROXIMA conexao. O processo
///    primario nunca trava esperando um '\n' que nao vai chegar.
///  - Falha SUSTENTADA de aceitar conexao (ex.: nome de pipe "squatted" por outro processo do MESMO
///    usuario, impedindo o bind) tambem cai no `catch (IOException)`, mas SEM o backoff limitado
///    (<see cref="BackoffMs"/>, 250ms) o loop giraria em CPU cheia pro resto da vida do processo
///    primario -- uma falha transitoria de 1 cliente que desconecta no meio e uma falha PERMANENTE de
///    bind passam pelo MESMO catch, entao o backoff se aplica aos dois.
///  - ACL do pipe: o DEFAULT do NamedPipeServerStream (sem opcao nenhuma) NAO e' "so o usuario atual" --
///    e' o descritor de seguranca padrao do Win32 (controle total pro LocalSystem/Administradores/dono,
///    mas LEITURA tambem pro grupo Everyone e pra conta anonima). `PipeOptions.CurrentUserOnly`
///    (servidor E cliente, abaixo) troca isso por um descritor que so permite o MESMO usuario (e no
///    Windows, tambem a MESMA elevacao) -- 2a camada de verdade por cima do nome unico por
///    usuario+sessao de NomesDeInstancia (que ja evita COLISAO de nome entre usuarios; CurrentUserOnly
///    evita que um usuario B, mesmo SABENDO o nome do pipe do usuario A por algum outro meio, consiga
///    conectar nele).
///
/// FORA do modelo de ameaca (residual aceito, mesma pratica do mPDF): um processo MALICIOSO rodando
/// como o MESMO usuario (mesma elevacao) ainda pode "ocupar" o nome do mutex/pipe primeiro (squatting)
/// -- nenhuma ACL de pipe/mutex nomeado distingue processos do MESMO usuario entre si. Aceito: um
/// atacante que ja roda codigo como o usuario tem vetores muito mais baratos e diretos do que orquestrar
/// uma corrida de squatting contra este mutex/pipe especificamente.
public sealed class InstanciaUnica : IDisposable
{
    private const int MaxLineLengthBytes = 32 * 1024;
    private const int ConnectTimeoutMs = 2000;
    private const int BackoffMs = 250;

    private readonly string _mutexName;
    private readonly string _pipeName;
    private CancellationTokenSource? _cts;
    private Task? _listenTask;
    private bool _disposed;

    // Divergencia deliberada do arquivo do mPDF (relatada no report da Task 4): la', TryAcquire guarda
    // o Mutex num campo e Dispose chama ReleaseMutex direto -- funciona porque TryAcquire/Dispose sao
    // sempre chamados pela MESMA thread do SO (a UI/main thread do WPF, de forma sincrona). Um mutex
    // nomeado do Win32 e' de propriedade da THREAD do SO que o adquiriu (nao do processo): so' essa
    // MESMA thread pode chamar ReleaseMutex -- de outra thread, a chamada lanca ApplicationException
    // ("... unsynchronized block of code"). Os testes desta classe (InstanciaUnicaTests, ports do
    // brief) usam `async Task` com `await` sobre um TaskCompletionSource resolvido POR OUTRA thread
    // (o loop de escuta do pipe, via Task.Run) -- sem SynchronizationContext capturado, a continuacao
    // do teste apos o `await` pode retomar num thread do POOL diferente do que chamou TentarAssumir,
    // e o Dispose() (chamado pelo `using` ao fim do teste) roda nesse thread diferente. Pra nao
    // depender de "quem chama TentarAssumir/Dispose usa sempre a mesma thread do SO" (verdade no WPF
    // sincrono do mPDF, mas nao garantida aqui), a posse do mutex fica confinada a UMA thread dedicada
    // e criada sob demanda, do Acquire ao Release: sempre a MESMA thread do SO, nunca a do chamador.
    private Thread? _threadDoMutex;
    private readonly ManualResetEventSlim _pedidoDeLiberacao = new(initialState: false);
    private bool _souPrimaria;

    public event Action<string>? CaminhoRecebido;

    /// <summary>Uma segunda instancia foi aberta SEM arquivo (menu Iniciar) — fix round 1 da Task 5 do Plano 3: o
    /// app pode estar vivo sem janela nenhuma (esperando uma impressao terminar para encerrar), e antes esse
    /// lancamento morria calado, sem mostrar nada.</summary>
    public event Action? AtivacaoRecebida;

    /// <summary>Linha do protocolo que pede "mostre-se" em vez de abrir um caminho. NAO e' um caminho absoluto, entao
    /// uma primaria de versao anterior (que so' aceita caminho absoluto) a ignora, como sempre ignorou lixo.</summary>
    public const string PedidoDeAtivacao = "?ativar";

    public InstanciaUnica(string mutex, string pipe)
    {
        _mutexName = mutex;
        _pipeName = pipe;
    }

    public bool TentarAssumir(string? caminhoParaEncaminhar)
    {
        var pronto = new ManualResetEventSlim(initialState: false);
        var criadoAgora = false;
        Exception? falhaNaThread = null;

        _threadDoMutex = new Thread(() =>
        {
            try
            {
                using var mutex = new System.Threading.Mutex(initiallyOwned: true, name: _mutexName, createdNew: out criadoAgora);
                pronto.Set();
                if (criadoAgora)
                {
                    // So' esta thread, que criou o mutex, pode libera-lo -- espera o Dispose() pedir.
                    _pedidoDeLiberacao.Wait();
                    mutex.ReleaseMutex();
                }
            }
            catch (Exception ex) { falhaNaThread = ex; pronto.Set(); }
        })
        { IsBackground = true, Name = "InstanciaUnica-Mutex" };
        _threadDoMutex.Start();
        pronto.Wait();
        pronto.Dispose();
        if (falhaNaThread is not null) throw falhaNaThread;

        _souPrimaria = criadoAgora;
        if (criadoAgora)
        {
            _cts = new CancellationTokenSource();
            _listenTask = Task.Run(() => LoopDeEscutaAsync(_cts.Token));
            return true;
        }

        // Nao conseguimos o mutex: ja existe uma instancia primaria (a thread dedicada acima so'
        // abriu um handle pro objeto NOMEADO ja existente -- Win32 nao concede posse nesse caso, ver
        // doc do CreateMutex -- e ja terminou sozinha, sem entrar no `_pedidoDeLiberacao.Wait()`
        // porque `criadoAgora` e' false). Nao ha' nada pra liberar/possuir aqui.
        // Sem arquivo: pede so' que a primaria se mostre (ver AtivacaoRecebida).
        try { EnviarParaPrimaria(string.IsNullOrWhiteSpace(caminhoParaEncaminhar) ? PedidoDeAtivacao : caminhoParaEncaminhar); }
        catch (Exception ex) when (ex is IOException or TimeoutException or ObjectDisposedException)
        {
            // A instancia primaria pode ter fechado bem entre a checagem do mutex acima e a
            // conexao do pipe aqui (janela de corrida pequena, mas real) -- best-effort: a
            // secundaria ainda sai (retorna false), so nao consegue avisar ninguem.
        }
        return false;
    }

    private void EnviarParaPrimaria(string caminho)
    {
        using var cliente = new NamedPipeClientStream(".", _pipeName, PipeDirection.Out, PipeOptions.CurrentUserOnly);
        cliente.Connect(ConnectTimeoutMs); // lanca TimeoutException se a primaria nao aceitar a tempo
        var bytes = Encoding.UTF8.GetBytes(caminho + "\n");
        cliente.Write(bytes, 0, bytes.Length);
        cliente.Flush();
    }

    private async Task LoopDeEscutaAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            Stream? servidor = null;
            try
            {
                servidor = await AceitarConexaoAsync(token).ConfigureAwait(false);

                var linha = await LerLinhaComTetoAsync(servidor, token).ConfigureAwait(false);
                // So repassa se for um caminho ABSOLUTO de verdade -- sem esta checagem, uma linha
                // relativa (protocolo malformado OU um bug em algum futuro remetente) chegaria a
                // CaminhoRecebido e o consumidor tentaria abrir relativo ao diretorio de trabalho ATUAL
                // do processo primario -- quase certamente nao e' o que o usuario quis.
                if (linha == PedidoDeAtivacao) AtivacaoRecebida?.Invoke();
                else if (!string.IsNullOrWhiteSpace(linha) && Path.IsPathRooted(linha))
                {
                    CaminhoRecebido?.Invoke(linha);
                }
            }
            catch (OperationCanceledException) { break; } // Dispose cancelou -- encerra o loop
            catch (ObjectDisposedException) { break; }     // idem, corrida de teardown
            catch (IOException)
            {
                // 2 causas possiveis, MESMO catch (ver doc XML da classe): (a) cliente desconectou/
                // crashou no meio de UMA conexao -- transitorio; ou (b) falha SUSTENTADA de aceitar
                // conexao nenhuma (ex.: pipe "squatted"). Sem o backoff abaixo, o caso (b) giraria em
                // CPU cheia pro resto da vida do processo -- o `catch(OperationCanceledException)`
                // aninhado cobre o Dispose cancelando ENQUANTO o backoff esta em andamento.
                try { await Task.Delay(BackoffMs, token).ConfigureAwait(false); }
                catch (OperationCanceledException) { break; }
            }
            finally { servidor?.Dispose(); }
        }
    }

    /// Cria uma `NamedPipeServerStream` nova (endurecida com `CurrentUserOnly`, ver doc XML da classe)
    /// e espera 1 conexao. Lanca IOException/OperationCanceledException tanto num bind que falha quanto
    /// numa espera cancelada -- o chamador (`LoopDeEscutaAsync`) trata os 2 casos.
    private async Task<Stream> AceitarConexaoAsync(CancellationToken token)
    {
        var servidor = new NamedPipeServerStream(
            _pipeName, PipeDirection.In, maxNumberOfServerInstances: 1,
            PipeTransmissionMode.Byte, PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        try
        {
            await servidor.WaitForConnectionAsync(token).ConfigureAwait(false);
            return servidor;
        }
        catch
        {
            servidor.Dispose();
            throw;
        }
    }

    /// Le uma linha terminada em '\n' de `stream`, com teto de <see cref="MaxLineLengthBytes"/> bytes.
    /// Devolve null (linha "invalida", ignorada pelo chamador) em 3 casos: EOF sem nunca achar '\n'
    /// (cliente fechou/crashou no meio), teto excedido sem achar '\n' (protocolo malformado), ou
    /// IOException durante a leitura (mesma categoria de desconexao abrupta).
    private static async Task<string?> LerLinhaComTetoAsync(Stream stream, CancellationToken token)
    {
        using var buffer = new MemoryStream();
        var chunk = new byte[4096];
        while (true)
        {
            int lidos;
            try { lidos = await stream.ReadAsync(chunk.AsMemory(0, chunk.Length), token).ConfigureAwait(false); }
            catch (IOException) { return null; }
            if (lidos == 0) return null; // EOF sem '\n' -- linha incompleta

            int indiceDaQuebra = Array.IndexOf(chunk, (byte)'\n', 0, lidos);
            if (indiceDaQuebra >= 0)
            {
                buffer.Write(chunk, 0, indiceDaQuebra);
                break;
            }

            buffer.Write(chunk, 0, lidos);
            if (buffer.Length > MaxLineLengthBytes) return null; // excedeu o teto -- descarta
        }

        var bytes = buffer.ToArray();
        if (bytes.Length > 0 && bytes[^1] == (byte)'\r') bytes = bytes[..^1]; // tolera CRLF
        return Encoding.UTF8.GetString(bytes);
    }

    /// Decide se o lancamento deve CONTINUAR (mostrar a janela normalmente) ou SAIR (uma instancia
    /// primaria ja existe e recebeu o caminho). Instancia unica e' BEST-EFFORT: uma falha em
    /// `tentarAssumir` (ex.: UnauthorizedAccessException/WaitHandleCannotBeOpenedException -- um mutex
    /// NOMEADO ja existindo como um tipo de handle diferente, ou permissao negada por alguma politica do
    /// ambiente) NUNCA pode impedir o app de abrir. Captura qualquer excecao, registra via `logar`
    /// (producao: log do app) e devolve true -- falha ABERTA: o lancamento segue como se esta fosse uma
    /// instancia primaria normal. Custo aceito do fail-open: esta sessao especificamente perde a
    /// protecao de instancia unica (pode acabar com 2 janelas se a causa raiz persistir) -- estritamente
    /// melhor que recusar abrir o app por causa de um mecanismo auxiliar (instancia unica nao e' o
    /// motivo de existir de nenhum app do mOffice).
    public static bool DeveContinuar(Func<bool> tentarAssumir, Action<Exception> logar)
    {
        try
        {
            return tentarAssumir();
        }
        catch (Exception e)
        {
            logar(e);
            return true; // fail-open -- ver doc XML do metodo
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _cts?.Cancel();
        try { _listenTask?.Wait(TimeSpan.FromSeconds(2)); }
        catch (AggregateException) { /* loop ja encerrando/encerrado -- nao pode travar o Dispose */ }
        _cts?.Dispose();

        if (_souPrimaria)
        {
            // Sinaliza a thread dedicada (ver campo `_threadDoMutex`) pra ela mesma chamar
            // ReleaseMutex -- so' ela, a que criou o mutex, pode faze-lo (posse e' por THREAD do SO).
            _pedidoDeLiberacao.Set();
            _threadDoMutex?.Join(TimeSpan.FromSeconds(2));
        }
        _pedidoDeLiberacao.Dispose();
    }
}
