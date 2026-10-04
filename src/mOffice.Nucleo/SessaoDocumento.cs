using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>Um documento aberto: o arquivo de origem e a pasta de recovery que o x2t preencheu
/// (Editor.bin + media/ + changes/), no mesmo desenho do app oficial.</summary>
public sealed class SessaoDocumento
{
    private string _arquivoOrigem = null!;
    private FormatoDocumento _formato;
    private bool _novo;
    private bool _compat;

    // `init` (nao um `set` publico) para a construcao normal em JanelaEditor.AbrirAsync: o unico outro lugar
    // que muda estes dois e' Repontar, que grava direto no campo (permitido dentro da propria classe) — assim
    // fica explicito que "salvar como" e' o UNICO caminho de mudanca depois da abertura, achado 1 da revisao
    // 2026-09-17 (fix review 1): sem isto a sessao continuava nomeando o arquivo antigo depois de um "salvar
    // como" bem-sucedido.
    public required string ArquivoOrigem { get => _arquivoOrigem; init => _arquivoOrigem = value; }
    public required FormatoDocumento Formato { get => _formato; init => _formato = value; }
    public required string PastaRecovery { get; init; }
    public string Bin => Path.Combine(PastaRecovery, "Editor.bin");
    public string ArquivoChanges => Path.Combine(PastaRecovery, "changes", "changes0.json");
    public bool Modificado { get; set; }
    /// <summary>Changes gravados NESTA sessao (o indice que o editor enxerga). Numa sessao recuperada comeca
    /// em 0 mesmo com <see cref="ChangesAbertos"/> changes ja' no arquivo — o editor numera os proprios changes
    /// a partir do que ele mesmo gravou.</summary>
    public int IndiceChanges { get; set; }

    // --- Recuperacao apos queda (Task 2 do Plano 3) --------------------------------------------------------

    /// <summary>Quantos changes ja' estavam em <see cref="ArquivoChanges"/> quando a sessao abriu (recuperados
    /// da sessao que caiu e reaplicados antes do EndLoad). O truncamento de <c>LocalFileSaveChanges</c> e'
    /// relativo a ESTA sessao, entao este prefixo nunca e' removido (<see cref="PonteDesktop.SalvarChanges"/>).</summary>
    public int ChangesAbertos { get; init; }

    /// <summary>O que <c>AscDesktopEditor.CheckUserId()</c> responde: 1 na primeira abertura da pasta, +1 a cada
    /// recuperacao (ver <see cref="Recuperacao.MetadadosSessao"/>).</summary>
    public int IdUsuario { get; init; } = 1;

    /// <summary>A sessao reabriu uma pasta de recuperacao (sem <c>x2t</c>; changes reaplicados antes do EndLoad).</summary>
    public bool Recuperada { get; init; }

    /// <summary>
    /// Senha do documento (Task 7 do Plano 3): a digitada no dialogo do editor ao abrir um arquivo protegido e,
    /// depois, a que o editor manda em cada <c>LocalFileSave</c> (Arquivo -&gt; Proteger troca ou remove). Vai para o
    /// <c>m_sSavePassword</c> do x2t no salvar — sem ela o arquivo gravado perderia a protecao. <c>null</c> = sem senha.
    ///
    /// SO' EM MEMORIA: nunca em disco (nem no <c>sessao.json</c>, que so' guarda <see cref="Recuperacao.MetadadosSessao.Protegido"/>),
    /// nunca no log, nunca de volta para a pagina (<c>PonteDesktop.SessaoJson</c> nao a inclui). `internal`: a
    /// sessao e' exposta a' casca por <c>JanelaEditor.SessaoAtual</c>, e a senha nao e' dado da casca.
    /// </summary>
    internal string? Senha { get; set; }

    /// <summary>
    /// Sessao RECUPERADA de um documento protegido por senha, ainda sem senha (revisao final do Plano 3, I1): o
    /// caminho do arquivo CIFRADO de origem, que esta sessao nao pode regravar — o <c>Editor.bin</c> recuperado ja'
    /// esta' decifrado e a senha se perdeu com a sessao que caiu, entao um "Salvar" gravaria por cima do original
    /// cifrado uma copia SEM protecao, em silencio. Zerado quando uma senha e' definida de novo (Arquivo -&gt;
    /// Proteger). So' em memoria; <c>null</c> em toda sessao que nao seja esse caso.
    /// </summary>
    internal string? OriginalProtegidoSemSenha { get; set; }

    /// <summary>O aviso "salvo sem senha" desta sessao recuperada ja' apareceu (fix round 2: uma vez por sessao).</summary>
    internal bool AvisouSalvoSemSenha { get; set; }

    /// <summary>
    /// Sessao RECUPERADA cujo arquivo de origem mudou (ou sumiu) depois da queda (revisao final do Plano 3, I2 —
    /// ver <see cref="Recuperacao.MetadadosSessao.OrigemMudou"/>): o usuario provavelmente reabriu e salvou o
    /// original, e gravar a recuperacao por cima perderia a versao mais nova. Enquanto verdadeiro, o primeiro
    /// salvamento vira "Salvar como" (nome sugerido "&lt;nome&gt; (recuperado)") e gravar no arquivo de origem e'
    /// recusado. Zerado pelo primeiro salvamento de documento que der certo.
    /// </summary>
    public bool Desvinculada { get; internal set; }

    /// <summary>Metadados gravados em <c>sessao.json</c> ao abrir, para regravar no "salvar como" (<see cref="Repontar"/>).
    /// Nulo em sessoes criadas sem janela (testes puros).</summary>
    internal Recuperacao.MetadadosSessao? Metadados { get; set; }

    /// <summary><c>rec.lock</c> desta sessao (<see cref="Recuperacao.TravaSessao"/>). Liberada por
    /// <see cref="ApagarRecovery"/> ou <see cref="LiberarTrava"/>.</summary>
    internal Recuperacao.TravaSessao? Trava { get; set; }

    /// <summary>Solta a trava sem apagar nada (fechar deixando a pasta para uma recuperacao — so' testes). Solta tambem a trava do
    /// ORIGINAL (Plano 6, Task 5): todo fechamento da janela passa por aqui ou por <see cref="ApagarRecovery"/>.</summary>
    internal void LiberarTrava()
    {
        Trava?.Dispose();
        Trava = null;
        SoltarTravaDoOriginal();
    }

    // --- Trava do arquivo do usuario (Plano 6, Task 5 — SO-02/MvP-05) ----------------------------------------------------

    /// <summary>Handle de leitura sobre o original + arquivo de dono <c>~$</c> (<see cref="Travas.TravaDoDocumento"/>). Adquirida quando
    /// o x2t termina de ler um original <see cref="Travas.EstadoAcesso.Livre"/>; muda para o arquivo novo no "Salvar como"; solta e
    /// apagada (o <c>~$</c>) no fechamento. Independente da pasta de recuperacao: um fechamento que MANTEM a pasta (queda simulada,
    /// encerramento do Windows) tambem solta o original. Processo morto: o SO solta o handle e o <c>~$</c> fica (como no Word) — a
    /// proxima abertura, sem handle de ninguem, sonda <c>Livre</c> e o substitui.</summary>
    internal Travas.TravaDoDocumento? TravaDoOriginal { get; set; }

    /// <summary>Estado do original na abertura/ultimo salvamento desta sessao (<see cref="Travas.CarimboArquivo"/>). Mudou na hora de
    /// gravar por cima = conflito. <c>null</c>: sem original (novo) ou sessao recuperada (o I2 do Plano 3 trata o original mudado).</summary>
    internal Travas.CarimboArquivo? Carimbo { get; set; }

    /// <summary>A sessao abriu um original que nao pode ser gravado: em uso por outro, atributo somente leitura ou sem permissao. O
    /// editor fica com a restricao ReadOnly (cabecalho "Somente leitura"), o "Salvar" vira "Salvar como" (<see cref="ExigeSalvarComo"/>)
    /// e gravar no original e' recusado. Zerado no "Salvar como" que der certo (<see cref="Repontar"/>): o arquivo novo e' do usuario.</summary>
    public bool SomenteLeitura { get; internal set; }

    /// <summary>Por que a sessao e' <see cref="SomenteLeitura"/> (<see cref="Travas.EstadoAcesso.Livre"/> quando nao e').</summary>
    public Travas.EstadoAcesso MotivoSomenteLeitura { get; internal set; }

    /// <summary>Abertura <see cref="Travas.EstadoAcesso.Livre"/> ainda sem trava: o original e' travado quando o x2t terminar de le-lo
    /// (inclusive depois da senha/opcoes de CSV, que reconvertem).</summary>
    internal bool TravarAoAbrir { get; set; }

    /// <summary>Copia do original em uso por outro, na pasta da sessao: e' ela que o x2t le' (tambem nas reconversoes de senha e CSV).
    /// <c>null</c> fora desse caso.</summary>
    internal string? CopiaDeLeitura { get; init; }

    /// <summary>O arquivo que o x2t le' para abrir esta sessao: a copia do original em uso, ou o proprio original.</summary>
    internal string EntradaDeLeitura => CopiaDeLeitura ?? ArquivoOrigem;

    /// <summary>Solta e descarta a trava do original (apaga o <c>~$</c>). Idempotente.</summary>
    internal void SoltarTravaDoOriginal()
    {
        TravaDoOriginal?.Dispose();
        TravaDoOriginal = null;
    }

    /// <summary>Codificacao e delimitador do CSV/TXT desta sessao (Plano 6, Task 3): os que o usuario escolheu no dialogo de abertura
    /// ou na linha do "Salvar como". O "Salvar" seguinte (mesmo arquivo) grava com eles — um CSV ANSI aberto salva ANSI, como o Excel.
    /// <c>null</c> fora de CSV/TXT (e numa sessao recuperada, que cai no padrao do formato).</summary>
    public OpcoesTexto? OpcoesTexto
    {
        get => _opcoesTexto;
        set
        {
            _opcoesTexto = value;
            if (value is not null) OpcoesTextoDesconhecidas = false;
            // O sessao.json acompanha: uma sessao RECUPERADA de CSV/TXT tem de regravar o arquivo com a MESMA codificacao/delimitador.
            if (Metadados is { } m)
            {
                Metadados = ComOpcoesDeTexto(m);
                try { Metadados.Gravar(PastaRecovery); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
    }
    private OpcoesTexto? _opcoesTexto;

    private Recuperacao.MetadadosSessao ComOpcoesDeTexto(Recuperacao.MetadadosSessao m) =>
        m with { TextoCodificacao = _opcoesTexto?.Codificacao, TextoDelimitador = _opcoesTexto?.Delimitador, TextoDelimitadorChar = _opcoesTexto?.DelimitadorChar };

    /// <summary>Sessao CSV/TXT recuperada de um <c>sessao.json</c> SEM as opcoes de texto (ou ilegiveis): nao se sabe com que codificacao o
    /// original foi lido. O "Salvar" vira "Salvar como" (<see cref="ExigeSalvarComo"/>) e gravar por cima do original sem escolher as
    /// opcoes numa linha do dialogo e' recusado — nunca se adivinha uma codificacao para o arquivo que o usuario ja' tem. Zerado ao
    /// definir <see cref="OpcoesTexto"/> ou no "Salvar como" que der certo.</summary>
    internal bool OpcoesTextoDesconhecidas { get; set; }

    /// <summary>Titulo do documento NOVO que nasce de um arquivo convertido (PDF aberto como documento, Plano 6, Task 3): o nome do
    /// PDF em vez do rotulo generico. Tem precedencia sobre o rotulo; so' vale enquanto o documento for <see cref="Novo"/>.</summary>
    public string? TituloDoNovo { get; init; }

    /// <summary>Documento NOVO (Task 2 do Marco 1 / Plano 2): a sessao nasceu de uma copia do modelo vazio do
    /// motor (<see cref="ModelosVazios"/>), nao de um arquivo do usuario. Consequencias: o titulo e' generico
    /// ("Documento1"/"Pasta1"/"Apresentação1" pelo ContadorDeNovos, nao "new.docx") e "Salvar" tem de virar "Salvar como" — nao ha
    /// arquivo de origem para gravar por cima. Campo de apoio (e nao `{ get; init; }` puro) porque
    /// <see cref="Repontar"/> precisa zera-lo: depois do primeiro "salvar como" a sessao passa a ter arquivo
    /// de verdade e deixa de ser nova.</summary>
    public bool Novo { get => _novo; init => _novo = value; }

    /// <summary>Modo de compatibilidade (Plano 6, Task 2): a sessao veio de um formato LEGADO (.doc/.xls/.ppt…) que o x2t le' mas
    /// nao grava. <see cref="Formato"/> ja' e' o moderno do tipo (.docx/.xlsx/.pptx) e <see cref="ArquivoOrigem"/> continua o
    /// arquivo antigo, que nunca e' regravado: o primeiro "Salvar" vira "Salvar como" (<see cref="ExigeSalvarComo"/>). Campo de
    /// apoio porque <see cref="Repontar"/> o zera — depois do primeiro "Salvar como" a sessao aponta para o arquivo novo.</summary>
    public bool Compatibilidade { get => _compat; init => _compat = value; }

    /// <summary>O "Salvar" desta sessao tem de perguntar o destino: nao ha' arquivo gravavel de origem (documento novo, original
    /// de uma sessao recuperada que mudou, formato legado, original somente leitura ou em uso — Plano 6, Task 5).</summary>
    public bool ExigeSalvarComo => Novo || Desvinculada || Compatibilidade || OpcoesTextoDesconhecidas || SomenteLeitura;

    /// <summary>O que aparece na barra de titulo da janela e no <c>SetDocumentName</c> do editor: o nome do
    /// arquivo, ou o rotulo generico por tipo enquanto o documento for <see cref="Novo"/>.</summary>
    public string Titulo => Novo
        ? TituloDoNovo ?? (NumeroDoNovo > 0 ? RotuloDoNovo + NumeroDoNovo : RotuloDoNovo)
        : Path.GetFileName(ArquivoOrigem);

    /// <summary>Numero do documento novo no processo (Documento1, Pasta2…; <see cref="ContadorDeNovos"/>). 0 = sem numero: sessao criada
    /// fora da janela (testes, sessao recuperada) fica so' com o rotulo.</summary>
    public int NumeroDoNovo { get; init; }

    private string RotuloDoNovo => Formato.DocType() switch { "cell" => "Pasta", "slide" => "Apresentação", _ => "Documento" };

    /// <summary>Depois de um "Salvar como" bem-sucedido a sessao passa a apontar para o arquivo NOVO — mesma
    /// decisao dos editores reais (o Word/LibreOffice tambem "adotam" o destino do salvar como). A partir daqui
    /// um <see cref="JanelaEditor.SalvarAsync"/> grava no arquivo novo, nao mais no antigo.</summary>
    public void Repontar(string novoArquivo, FormatoDocumento novoFormato, OpcoesTexto? texto = null)
    {
        _arquivoOrigem = novoArquivo;
        _formato = novoFormato;
        // Plano 6, Task 3: so' CSV/TXT guardam codificacao e delimitador; qualquer outro formato zera (um "Salvar" seguinte nao herda).
        _opcoesTexto = novoFormato is FormatoDocumento.Csv or FormatoDocumento.Txt ? texto : null;
        OpcoesTextoDesconhecidas = false;
        // Task 2 do Marco 1 / Plano 2: o "salvar como" e' EXATAMENTE o momento em que um documento novo deixa
        // de ser novo — passa a ter arquivo de verdade, titulo com o nome dele e "Salvar" grava por cima.
        _novo = false;
        // Plano 6, Task 2: depois do "Salvar como" de um legado a sessao aponta para o arquivo MODERNO novo.
        _compat = false;
        // Plano 6, Task 5: o arquivo novo foi gravado pelo usuario — deixa de ser somente leitura (como no Office).
        SomenteLeitura = false;
        MotivoSomenteLeitura = Travas.EstadoAcesso.Livre;
        // Task 2 do Plano 3: uma queda DEPOIS do "salvar como" tem de reabrir apontando para o arquivo novo
        // (e com o titulo dele). O IdUsuario continua o mesmo — e' a mesma abertura da pasta. Melhor esforco:
        // o salvamento ja' deu certo, e falhar aqui so' pioraria a recuperacao, nao o documento.
        if (Metadados is { } m)
        {
            Metadados = ComOpcoesDeTexto(m) with { Origem = novoArquivo, Formato = (int)novoFormato, Titulo = Titulo, Novo = false, Compatibilidade = false };
            try { Metadados.Gravar(PastaRecovery); } catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        }
    }

    // ---------------------------------------------------------------------------------------------------
    // Retencao (achado da revisao final 2026-09-17). Cada abertura cria %TEMP%\mOffice\recovery\<guid> com o
    // Editor.bin (o documento INTEIRO, em claro) + media/ + changes/, e nada apagava: 353 pastas orfas
    // medidas nesta maquina. O PoC nao tem recuperacao apos queda — nenhum caminho do codigo procura uma
    // pasta de recovery de uma sessao anterior —, entao guardar isso nao serve para nada e so' acumula
    // conteudo de documento no disco. Politica: apagar ao fechar a janela + varredura do que ficou para tras.
    // (A pasta temporaria do x2t e' outra coisa e ja' e' limpa por X2tConversor a cada chamada.)
    // ---------------------------------------------------------------------------------------------------

    // Task 1 do Marco 1 / Plano 2 (backlog 15): morava em %TEMP%\mOffice\recovery — um limpador de disco
    // podia apagar por baixo uma sessao em uso, e nada do produto deveria gravar em pasta temporaria do
    // sistema. Agora vive em <see cref="Pastas.Recuperacao"/> (%LOCALAPPDATA%\mOffice\recuperacao), a mesma
    // raiz de dados por usuario do resto do produto.
    public static string RaizRecovery => Pastas.Recuperacao;

    /// <summary>Apaga a pasta de recovery desta sessao. Melhor esforco: roda no fechamento da janela, onde uma
    /// excecao nao tem para onde ir (e um arquivo ainda travado pelo x2t nao e' motivo para quebrar o fim do
    /// programa). O que escapar aqui cai na varredura da proxima execucao.</summary>
    public void ApagarRecovery()
    {
        // sessao.json PRIMEIRO, ainda com a trava na mao (revisao final do Plano 3, M6 — a mesma ordem de
        // Recuperaveis.Descartar): sem ele a pasta deixa de ser recuperavel na hora. Um apagar que pare no meio
        // (processo morto, arquivo segurado por antivirus/backup) nunca deixa uma pasta que ainda PARECE
        // recuperavel — o que faria o documento que o usuario mandou descartar ("Nao salvar") reaparecer como
        // "documento recuperado" na proxima subida. O resto cai na varredura.
        try { File.Delete(Path.Combine(PastaRecovery, Recuperacao.MetadadosSessao.NomeArquivo)); } catch { }
        // A trava (rec.lock, FileShare.None) fica DENTRO da pasta: sem solta-la antes, o Directory.Delete
        // recursivo falharia nela depois de ja' ter apagado o resto (Task 2 do Plano 3).
        LiberarTrava();
        try
        {
            AntesDoApagarRecursivoParaTeste?.Invoke();
            if (Directory.Exists(PastaRecovery)) Directory.Delete(PastaRecovery, true);
        }
        catch { }
    }

    /// <summary>Costura de teste (M6): roda depois de o <c>sessao.json</c> sair e antes do apagar recursivo — o
    /// teste lanca aqui para simular um apagar interrompido. `internal`; por INSTANCIA (nunca estatica: os testes de
    /// janela real rodam em paralelo e chamam <see cref="ApagarRecovery"/> das sessoes deles).</summary>
    internal Action? AntesDoApagarRecursivoParaTeste { get; set; }

    /// <summary>Apaga as pastas de <paramref name="raiz"/> cujo frescor (<see cref="UltimaEscritaUtc"/>) e' mais
    /// antigo que <paramref name="idadeMinima"/> — o que sobrou de execucoes anteriores (queda, kill do
    /// processo, ou as 353 que ja' estavam la'). Devolve quantas apagou. Nunca lanca: e' chamada na subida do
    /// app. NUNCA apaga uma pasta recuperavel (<see cref="Recuperacao.Recuperaveis.EhRecuperavel"/>) nem uma com
    /// a trava presa (Task 3 do Plano 3); pastas sem changes (abriu e nao editou) continuam saindo pela idade.
    ///
    /// Recusa a RAIZ inteira, sem apagar nada, se ela propria for um link simbolico/juncao (achado 3 da
    /// revisao final, fix review 2): seguir o link levaria a varredura a apagar conteudo do outro lado dele,
    /// fora de <c>%TEMP%\mOffice\recovery</c> — a raiz e' fixa (<see cref="RaizRecovery"/>), entao nao ha
    /// motivo legitimo para ela ser um reparse point.</summary>
    public static int VarrerRecoveryAntigo(string raiz, TimeSpan idadeMinima)
    {
        var apagadas = 0;
        try
        {
            if (!Directory.Exists(raiz)) return 0;
            var infoRaiz = new DirectoryInfo(raiz);
            if (infoRaiz.LinkTarget is not null || infoRaiz.Attributes.HasFlag(FileAttributes.ReparsePoint)) return 0;
            var limite = DateTime.UtcNow - idadeMinima;
            foreach (var pasta in Directory.EnumerateDirectories(raiz))
            {
                try
                {
                    if (UltimaEscritaUtc(pasta) >= limite) continue;
                    // Task 3 do Plano 3 (spec §8, dado do usuario vem primeiro): um documento RECUPERAVEL nunca e'
                    // varrido, qualquer que seja a idade — so' some por "Recuperar" (e fechamento limpo) ou
                    // "Descartar". E uma pasta com a trava presa (sessao viva, talvez ociosa ha' mais de 30 dias
                    // em outra instancia) nunca e' tocada: o Delete recursivo apagaria o conteudo e so' falharia
                    // no rec.lock, deixando a janela viva sem Editor.bin. As duas checagens ficam logo antes do
                    // Delete (depois da idade, que e' a mais barata) para encurtar a janela entre olhar e apagar.
                    if (!Recuperacao.TravaSessao.Livre(pasta)) continue;
                    if (Recuperacao.Recuperaveis.EhRecuperavel(pasta)) continue;
                    Directory.Delete(pasta, true);
                    apagadas++;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return apagadas;
    }

    /// <summary>Frescor real de uma pasta de recovery: o MAIOR entre a data da propria pasta e a do arquivo mais
    /// recente dentro dela, recursivamente (achado 2 da revisao final, fix review 2). A pasta em si so' muda de
    /// <c>LastWriteTimeUtc</c> quando um arquivo e' CRIADO ou APAGADO dentro dela — gravar em cima de um
    /// arquivo ja' existente (o editor apendando <c>changes/changes0.json</c>, o x2t reescrevendo
    /// <c>Editor.bin</c>) NAO toca a data da pasta. MEDIDO na revisao final: apendar em
    /// <c>changes/changes0.json</c> e regravar <c>Editor.bin</c> de uma pasta ja' criada deixou
    /// <c>Directory.GetLastWriteTimeUtc</c> da pasta INALTERADO. Sem isto, um documento aberto ha' mais de
    /// <c>idadeMinima</c> numa instancia diferente do app — com escrita recente DENTRO da pasta — teria a
    /// pasta de recovery apagada por baixo enquanto ainda em uso. Erros de E/S por arquivo (removido, travado
    /// por outro processo, sem permissao) sao ignorados individualmente: o que importa e' o maior valor que
    /// deu para ler, nao que TODOS os arquivos respondam.</summary>
    private static DateTime UltimaEscritaUtc(string pasta)
    {
        var maisRecente = Directory.GetLastWriteTimeUtc(pasta);
        IEnumerable<string> arquivos;
        try { arquivos = Directory.EnumerateFiles(pasta, "*", SearchOption.AllDirectories).ToArray(); }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { return maisRecente; }
        foreach (var arquivo in arquivos)
        {
            try
            {
                var data = File.GetLastWriteTimeUtc(arquivo);
                if (data > maisRecente) maisRecente = data;
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        }
        return maisRecente;
    }
}
