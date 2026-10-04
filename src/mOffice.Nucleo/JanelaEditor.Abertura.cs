using System.IO.Compression;
using System.Windows;
using mOffice.Conversor;
using mOffice.Nucleo.Travas;

namespace mOffice.Nucleo;

/// <summary>
/// Plano 6, Task 2 — abertura de formatos legados (Modo de compatibilidade), modelos (documento novo a partir do modelo),
/// macros (VBA preservado, nao executado) e apresentacao de slides (abre e inicia a apresentacao).
/// </summary>
public partial class JanelaEditor
{
    /// <summary>A sessao atual veio de um formato legado (.doc/.xls/.ppt…): o x2t nao o grava, entao o primeiro salvar
    /// vira "Salvar como" no formato moderno e o original nunca e' tocado (decisao do dono, Plano 6).</summary>
    public bool ModoCompatibilidade => _sessao?.Compatibilidade == true;

    /// <summary>O pacote aberto tem <c>vbaProject.bin</c>: as macros VBA ficam no arquivo e nao rodam.</summary>
    public bool MacrosPreservadas { get; private set; }

    public const string MensagemMacros =
        "Este documento contém macros do Office (VBA). Elas são mantidas no arquivo, mas não são executadas no mOffice.";

    /// <summary>Faixa do Modo de compatibilidade: "Este documento está no formato antigo do Word (.doc). Ao salvar, o mOffice
    /// grava uma cópia em .docx; o arquivo original não é alterado." (Word/Excel/PowerPoint pelo tipo do arquivo).</summary>
    public static string MensagemCompatibilidade(InfoExtensao legado) =>
        $"Este documento está no formato antigo do {legado.DocType switch { "cell" => "Excel", "slide" => "PowerPoint", _ => "Word" }} ({legado.Extensao}). "
        + $"Ao salvar, o mOffice grava uma cópia em {TabelaFormatos.Principal(TabelaFormatos.Moderno(legado.DocType)).Extensao}; o arquivo original não é alterado.";

    /// <summary>O usuario recusou a conversao do PDF em documento (<see cref="AbrirPdfComoDocumentoAsync"/>): nao e' erro. Tipo PROPRIO para a
    /// casca nao engolir qualquer outro <see cref="OperationCanceledException"/> como se fosse uma recusa.</summary>
    public sealed class ConversaoDePdfRecusadaException() : OperationCanceledException("O usuário não confirmou a conversão do PDF.");

    /// <summary>A pergunta antes de converter um PDF em documento editavel (FMT-07).</summary>
    public const string MensagemPdfComoDocumento =
        "O mTexto vai converter este PDF em um documento editável. O resultado pode não ficar idêntico ao original, principalmente "
        + "se o PDF tiver muitos gráficos ou tiver sido digitalizado. O PDF original não é alterado. Continuar?";

    /// <summary>Costura de teste da confirmacao (o <c>MessageBox</c> do host e' modal). Recebe o PDF; <c>true</c> = continuar.</summary>
    internal Func<string, bool>? ConfirmarAbrirPdfParaTeste { get; set; }

    /// <summary>
    /// FMT-07 (como o "reflow" do Word): depois da confirmacao, o x2t converte o PDF em DOCX numa pasta propria e o resultado abre como
    /// documento NOVO (o titulo e' o nome do PDF; salvar pergunta o destino; o PDF original nunca e' alterado). A pasta da conversao sai
    /// no fim: <see cref="AbrirSessaoAsync"/> copia o arquivo para a pasta de recuperacao da sessao nova.
    /// </summary>
    /// <exception cref="OperationCanceledException">O usuario recusou a confirmacao. A janela NAO se fecha aqui: quem a criou (o gerente
    /// de janelas) a fecha e mostra a boas-vindas — fechar daqui, sendo ela a unica janela, encerraria o app (o <c>AoFechar</c> do gerente).</exception>
    public async Task AbrirPdfComoDocumentoAsync(string pdf)
    {
        Func<string, bool> confirmar = ConfirmarAbrirPdfParaTeste ?? new Func<string, bool>(_ => MessageBox.Show(this, MensagemPdfComoDocumento, NomeApp, MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK);
        if (!confirmar(pdf)) throw new ConversaoDePdfRecusadaException();
        var pasta = Pastas.Garantir(Path.Combine(Pastas.Raiz, "conversao", Guid.NewGuid().ToString("N")));
        try
        {
            var temp = Path.Combine(pasta, Path.GetFileNameWithoutExtension(pdf) + ".docx");
            await _conv.ConverterAsync(pdf, FormatoDocumento.Docx, temp);
            await AbrirSessaoAsync(temp, FormatoDocumento.Docx, novo: true, tituloDoNovo: Path.GetFileNameWithoutExtension(pdf));
        }
        finally { try { Directory.Delete(pasta, true); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { } }
    }

    /// <summary>Avisos que a casca mostra na faixa quando o documento fica pronto, em ordem de importancia. Recalculados a
    /// cada abertura. As Tasks 3 e 5 acrescentam os delas aqui.</summary>
    public IReadOnlyList<string> AvisosDeAbertura => _avisosDeAbertura;
    private List<string> _avisosDeAbertura = [];

    /// <summary>Sessao que deve iniciar a apresentacao quando o editor ficar pronto (<see cref="AbrirEApresentarAsync"/>). Presa
    /// a SESSAO, nao a um booleano da janela: uma abertura que falha, ou um outro documento aberto antes do "pronto", nunca herda
    /// o pedido de apresentar.</summary>
    private SessaoDocumento? _apresentarSessao;

    /// <summary>Novo documento a partir de um modelo (.dotx/.xltx/.potx/.ott/…; duplo clique no Explorer): mesmo caminho
    /// do <see cref="NovoAsync"/> — a COPIA do modelo vai para a pasta de recuperacao e a sessao e' <c>Novo</c>, no formato
    /// moderno do tipo (ou no de macro, se o modelo tiver macro: .dotm -&gt; .docm).</summary>
    public async Task NovoDoModeloAsync(string modelo)
    {
        var info = TabelaFormatos.DaExtensao(modelo) is { Abre: true, Modelo: true } m ? m
            : throw new ErroConversao("Este arquivo não é um modelo que o mOffice abre.", 0, "");
        var formato = info.Macro ? info.DocType switch { "cell" => FormatoDocumento.Xlsm, "slide" => FormatoDocumento.Pptm, _ => FormatoDocumento.Docm }
                                 : TabelaFormatos.Moderno(info.DocType);
        await AbrirSessaoAsync(modelo, formato, novo: true);
    }

    /// <summary>Apresentacao de slides (.ppsx/.pps; duplo clique no Explorer): abre e, quando o editor fica pronto, inicia
    /// a apresentacao do primeiro slide pela ponte (<c>window.__moffice.iniciarApresentacao(false)</c>).</summary>
    public async Task AbrirEApresentarAsync(string caminho)
    {
        await AbrirSessaoAsync(caminho, FormatoDocumentoExtensoes.DaExtensao(caminho), novo: false, apresentar: true);
    }

    /// <summary>O editor avisou "pronto" para <paramref name="origem"/>: se ela foi aberta para apresentar, inicia a apresentacao.
    /// Chamado de dentro da chamada sincrona do JS — o script roda fora dela (<see cref="Dispatcher"/>).</summary>
    private void IniciarApresentacaoSePedida(SessaoDocumento origem)
    {
        if (!ReferenceEquals(origem, _apresentarSessao)) return;
        _apresentarSessao = null;
        _ = Dispatcher.InvokeAsync(async () =>
        {
            if (_fechada || !ReferenceEquals(origem, _sessao)) return;
            try
            {
                var r = (await ExecutarJsAsync("window.frames[0].__moffice.iniciarApresentacao(false)")).Trim('"');
                _arquivos.Log("APRESENTAR iniciarApresentacao(false): " + r);
            }
            catch (Exception ex) { _arquivos.Log("APRESENTAR falhou ao iniciar a apresentacao: " + ex.Message); }
        });
    }

    /// <summary>Avisos e estado de abertura de uma sessao nova (ou recuperada), na ordem de importancia (a faixa mostra o primeiro):
    /// somente leitura (Task 5), compatibilidade do legado, assinatura digital (Task 5) e macros preservadas. <paramref name="paraLer"/>
    /// e' o arquivo que da' para abrir como zip (a COPIA, quando o original esta' em uso por outro).</summary>
    private void CalcularAvisosDeAbertura(SessaoDocumento sessao, string arquivoDeEntrada, bool novo, string? paraLer = null)
    {
        paraLer ??= arquivoDeEntrada;
        _semReservaAvisadoPara = null;
        var info = TabelaFormatos.DaExtensao(arquivoDeEntrada);
        var avisos = new List<string>();
        if (sessao.SomenteLeitura) avisos.Add(MensagemSomenteLeitura(sessao.MotivoSomenteLeitura, _donoDoArquivoEmUso));
        if (sessao.Compatibilidade && info is not null) avisos.Add(MensagemCompatibilidade(info));
        AssinaturaDigital = !novo && AcessoAoArquivo.TemAssinaturaDigital(paraLer);
        if (AssinaturaDigital) avisos.Add(AvisoAssinaturaDigital);
        MacrosPreservadas = !novo && info is { Macro: true } && TemVba(paraLer);
        if (MacrosPreservadas) avisos.Add(MensagemMacros);
        _avisosDeAbertura = avisos;
    }

    // --- Trava, somente leitura e conflito (Plano 6, Task 5 — SO-02/MvP-05, SO-14 parte 1, MvP-12 minimo) -----------------

    /// <summary>A sessao atual nao grava no arquivo de origem (em uso por outro, atributo somente leitura, sem permissao).</summary>
    public bool SomenteLeitura => _sessao?.SomenteLeitura == true;

    /// <summary>O pacote aberto tem assinatura digital XML (<c>_xmlsignatures/</c> no OOXML, <c>META-INF/documentsignatures.xml</c> no ODF).</summary>
    public bool AssinaturaDigital { get; private set; }

    public const string AvisoAssinaturaDigital =
        "Este documento tem assinatura digital. Ao salvar alterações, a assinatura deixa de valer.";

    /// <summary>Faixa quando o original nao pode ser travado (rede que nao segura o handle, outro programa com ele aberto so' para leitura
    /// com compartilhamento estranho...) na abertura ou depois de um salvamento. O carimbo continua: o conflito ainda e' detectado.</summary>
    public const string AvisoSemReserva =
        "Não foi possível reservar este arquivo; outras pessoas podem abri-lo ao mesmo tempo. O mOffice vai avisar antes de gravar por cima de alterações de outra pessoa.";

    /// <summary>Costura de teste: a aquisicao da trava do original falha (como numa rede que nao segura o handle). `internal`.</summary>
    internal bool FalharTravaParaTeste { get; set; }

    private TravaDoDocumento? AdquirirTrava(string documento) =>
        FalharTravaParaTeste ? null : TravaDoDocumento.Adquirir(documento, Autor.Nome, _arquivos.Log);

    /// <summary>Arquivo para o qual a faixa "sem reserva" ja' foi mostrada e que continua sem trava (minor da Task 5 na revisao final do Plano 6):
    /// num salvamento, o Retomar do aposTroca e o de <see cref="TravarArquivoGravado"/> falhavam um atras do outro e a faixa aparecia DUAS vezes.
    /// Volta a <c>null</c> quando o arquivo e' travado de novo.</summary>
    private string? _semReservaAvisadoPara;

    /// <summary>A trava nao veio (ou nao voltou depois de um salvamento): log + faixa (uma vez enquanto o mesmo arquivo seguir sem trava).</summary>
    private void AvisarSemReserva(string arquivo, bool naAbertura)
    {
        _arquivos.Log($"TRAVA sem reserva de {Path.GetFileName(arquivo)}: avisando na faixa (o carimbo continua conferindo conflitos)");
        if (_semReservaAvisadoPara is { } ja && MesmoArquivo(ja, arquivo)) { _arquivos.Log("TRAVA faixa sem reserva ja' mostrada para este arquivo"); return; }
        _semReservaAvisadoPara = arquivo;
        if (naAbertura) _avisosDeAbertura.Insert(0, AvisoSemReserva);
        else AvisarUsuario(AvisoSemReserva);
    }

    public const string MensagemConflito =
        "O arquivo foi alterado por outra pessoa ou programa desde que você o abriu.";

    /// <summary>Recusa de gravar por cima do original numa sessao <see cref="SomenteLeitura"/> (pelo "Salvar como" escolhendo o proprio original).</summary>
    public const string MensagemSomenteLeituraRecusado =
        "Este documento foi aberto somente leitura: o arquivo original não pode ser alterado. Use Salvar como para gravar uma cópia com outro nome.";

    /// <summary>"&lt;arquivo&gt; está em uso por &lt;nome&gt;. Abrir somente leitura?" (sem nome: "por outra pessoa ou programa").</summary>
    public static string MensagemEmUso(string arquivo, string? dono) =>
        $"\"{arquivo}\" está em uso por {(string.IsNullOrWhiteSpace(dono) ? "outra pessoa ou programa" : dono.Trim())}. Abrir somente leitura?";

    /// <summary>Faixa da sessao somente leitura: "Somente leitura: &lt;motivo&gt;. Para guardar as alterações, use Salvar como."</summary>
    public static string MensagemSomenteLeitura(EstadoAcesso estado, string? dono = null) =>
        $"Somente leitura: {estado switch
        {
            EstadoAcesso.EmUso => $"o arquivo está em uso por {(string.IsNullOrWhiteSpace(dono) ? "outra pessoa ou programa" : dono.Trim())}",
            EstadoAcesso.SomenteLeituraAtributo => "o arquivo está marcado como somente leitura",
            EstadoAcesso.SemPermissaoDeEscrita => "você não tem permissão para alterar este arquivo",
            _ => "o arquivo não pode ser alterado",
        }}. Para guardar as alterações, use Salvar como.";

    /// <summary>O usuario escolheu "Cancelar" na pergunta do arquivo em uso: nao e' erro. Tipo PROPRIO (como
    /// <see cref="ConversaoDePdfRecusadaException"/>): a casca volta a' boas-vindas sem mensagem. A janela nao se fecha aqui — quem a
    /// criou (o gerente de janelas) a fecha.</summary>
    public sealed class AberturaEmUsoCanceladaException() : OperationCanceledException("O usuário cancelou a abertura do arquivo em uso.");

    /// <summary>Pergunta "&lt;arquivo&gt; está em uso por &lt;nome&gt;. Abrir somente leitura?" — argumentos: a janela, o NOME do arquivo e o
    /// dono lido do <c>~$</c> (ou <c>null</c>). <c>true</c> = abrir somente leitura; <c>false</c> = cancelar a abertura. O padrao e' um
    /// <c>MessageBox</c> Sim/Nao; a casca troca pelo dialogo dela e os testes por uma resposta fixa. `set` (e nao `init`, como no plano):
    /// como <see cref="PerguntarAoFechar"/>, a fabrica de janelas dos testes injeta uma resposta fixa — um modal de verdade travaria a
    /// thread STA de um teste que esbarrasse num arquivo em uso.</summary>
    public Func<JanelaEditor, string, string?, Task<bool>> PerguntarEmUso { get; set; } = PerguntaEmUsoPadrao;

    /// <summary>Pergunta do conflito ao salvar por cima (<see cref="MensagemConflito"/>): Salvar uma cópia / Substituir / Cancelar.</summary>
    public Func<JanelaEditor, Task<DecisaoConflito>> PerguntarConflito { get; set; } = PerguntaConflitoPadrao;

    private static Task<bool> PerguntaEmUsoPadrao(JanelaEditor j, string arquivo, string? dono) =>
        Task.FromResult(MessageBox.Show(j, MensagemEmUso(arquivo, dono), j.NomeApp, MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes);

    private static Task<DecisaoConflito> PerguntaConflitoPadrao(JanelaEditor j) =>
        Task.FromResult(MessageBox.Show(j, MensagemConflito + "\n\nSim: salvar uma cópia com outro nome.\nNão: substituir o arquivo pelas suas alterações.",
            j.NomeApp, MessageBoxButton.YesNoCancel, MessageBoxImage.Warning) switch
        {
            MessageBoxResult.Yes => DecisaoConflito.SalvarCopia,
            MessageBoxResult.No => DecisaoConflito.Substituir,
            _ => DecisaoConflito.Cancelar,
        });

    /// <summary>Dono lido do <c>~$</c> na abertura em uso (so' para o texto da faixa).</summary>
    private string? _donoDoArquivoEmUso;

    /// <summary>
    /// Antes do x2t, para um arquivo do usuario (nao para documento novo): sonda o acesso (<see cref="AcessoAoArquivo.Sondar"/>). Em uso
    /// por outro: pergunta (<see cref="PerguntarEmUso"/>); "Cancelar" lanca <see cref="AberturaEmUsoCanceladaException"/> sem ter criado
    /// nada. Devolve o estado e o dono (lido do <c>~$</c>).
    /// </summary>
    private async Task<(EstadoAcesso Estado, string? Dono)> SondarAcessoAsync(string original)
    {
        var estado = AcessoAoArquivo.Sondar(original);
        string? dono = null;
        if (estado == EstadoAcesso.EmUso)
        {
            dono = ArquivoDeDono.LerNome(ArquivoDeDono.CaminhoPara(original));
            _arquivos.Log($"TRAVA {Path.GetFileName(original)} em uso por outro processo (dono no ~$: {(dono is null ? "nenhum" : "sim")}): perguntando");
            if (!await PerguntarEmUso(this, Path.GetFileName(original), dono))
            {
                _arquivos.Log("TRAVA abertura cancelada pelo usuario (arquivo em uso)");
                throw new AberturaEmUsoCanceladaException();
            }
        }
        if (estado != EstadoAcesso.Livre) _arquivos.Log($"TRAVA {Path.GetFileName(original)} abre somente leitura ({estado})");
        return (estado, dono);
    }

    /// <summary>Copia o original em uso para a pasta da sessao (o x2t le' a copia), lendo com <c>FileShare.ReadWrite | Delete</c> —
    /// o "Word do colega" segura o arquivo para escrita e o x2t nao conseguiria abri-lo.</summary>
    private static string CopiarParaLeitura(string original, string pastaDaSessao)
    {
        var copia = Path.Combine(pastaDaSessao, "origem-somente-leitura" + Path.GetExtension(original));
        using var de = new FileStream(original, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var para = new FileStream(copia, FileMode.Create, FileAccess.Write, FileShare.None);
        de.CopyTo(para);
        return copia;
    }

    /// <summary>Depois que o x2t leu o original (abertura <see cref="EstadoAcesso.Livre"/>): trava o arquivo e cria o <c>~$</c> com o autor.
    /// Nunca numa janela ja' fechada nem numa sessao que nao e' mais a atual (a trava nunca mais seria solta).</summary>
    private void TravarOriginalSePedido(SessaoDocumento sessao)
    {
        if (!sessao.TravarAoAbrir) return;
        sessao.TravarAoAbrir = false;
        if (_fechada || !ReferenceEquals(sessao, _sessao)) return;
        sessao.TravaDoOriginal = AdquirirTrava(sessao.ArquivoOrigem);
        if (sessao.TravaDoOriginal is null) AvisarSemReserva(sessao.ArquivoOrigem, naAbertura: true);
    }

    /// <summary>Depois de um salvamento de DOCUMENTO bem-sucedido: a trava passa a ser do arquivo gravado (no "Salvar como", sai do antigo;
    /// numa sessao sem trava — documento novo, recuperada, somente leitura — passa a existir) e o carimbo e' renovado.</summary>
    /// <param name="carimboDaTroca">O carimbo tirado logo depois da troca, dentro do <see cref="GravacaoNoDestino.Substituir"/> (MINOR 1 da
    /// revisao): e' ele a nova referencia — nunca um carimbo relido depois, que poderia ja' incluir uma gravacao de terceiros.</param>
    private void TravarArquivoGravado(SessaoDocumento s, CarimboArquivo? carimboDaTroca)
    {
        if (_fechada || !ReferenceEquals(s, _sessao)) return;
        var travado = true;
        if (s.TravaDoOriginal is { } t)
        {
            if (!MesmoArquivo(t.Documento, s.ArquivoOrigem)) travado = t.Mover(s.ArquivoOrigem);
            else if (!t.Segurando) travado = t.Retomar();
        }
        else travado = (s.TravaDoOriginal = AdquirirTrava(s.ArquivoOrigem)) is not null;
        if (!travado) AvisarSemReserva(s.ArquivoOrigem, naAbertura: false);
        else _semReservaAvisadoPara = null;
        s.Carimbo = carimboDaTroca ?? CarimboArquivo.De(s.ArquivoOrigem);
    }

    /// <summary>Conflito (SO-14 parte 1): o original mudou (ou sumiu) desde a abertura/ultimo salvamento. Pergunta ANTES de qualquer E/S.
    /// <c>null</c> = pode gravar (sem conflito, ou "Substituir"); senao a excecao do cancelamento — em "Salvar uma cópia", o "Salvar como"
    /// e' aberto logo depois que este salvamento terminar.</summary>
    private async Task<Exception?> ConferirConflitoAsync(SessaoDocumento s, string destino, FormatoDocumento fmt)
    {
        if (fmt.EhExportacao() || s.Carimbo is not { } carimbo || !MesmoArquivo(destino, s.ArquivoOrigem)) return null;
        if (CarimboArquivo.De(destino) == carimbo) return null;
        _arquivos.Log("CONFLITO o original mudou desde a abertura/ultimo salvamento: perguntando antes de gravar");
        DecisaoConflito decisao;
        try { decisao = await PerguntarConflito(this); }
        catch (Exception ex) { _arquivos.Log("CONFLITO a pergunta falhou (" + ex.Message + "): tratado como Cancelar"); decisao = DecisaoConflito.Cancelar; }
        _arquivos.Log("CONFLITO decisao: " + decisao);
        if (decisao == DecisaoConflito.Substituir) return null;
        if (decisao == DecisaoConflito.SalvarCopia)
            _ = Dispatcher.BeginInvoke(async () =>
            {
                if (_fechada) return;
                try { await SalvarComoComDialogoAsync(); }
                catch (Exception ex) { _arquivos.Log("CONFLITO salvar uma copia falhou: " + ex.Message); }
            });
        return new SalvamentoCanceladoException(decisao == DecisaoConflito.SalvarCopia
            ? MensagemConflito + " As alterações não foram gravadas nele: salve uma cópia com outro nome."
            : MensagemConflito + " O salvamento foi cancelado.");
    }

    /// <summary>O usuario cancelou o salvamento na pergunta do conflito. Mesmo caminho de uma gravacao que falhou (EndSave 2, documento
    /// continua modificado; excecao para quem chamou pela API do host).</summary>
    public sealed class SalvamentoCanceladoException(string mensagem) : OperationCanceledException(mensagem);

    /// <summary>O pacote tem um <c>vbaProject.bin</c> (<c>word/</c>, <c>xl/</c> ou <c>ppt/</c>). <c>false</c> em qualquer erro de E/S
    /// ou zip invalido — so' decide se a faixa de macros aparece.</summary>
    internal static bool TemVba(string pacote)
    {
        try
        {
            using var z = ZipFile.OpenRead(pacote);
            return z.Entries.Any(e => e.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception e) when (e is IOException or InvalidDataException or UnauthorizedAccessException) { return false; }
    }
}
