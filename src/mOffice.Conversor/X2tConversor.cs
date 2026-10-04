using System.Diagnostics;
using System.Text;
using System.Xml.Linq;

namespace mOffice.Conversor;

/// <summary>
/// Chama x2t.exe como processo com um XML de parametros (TaskQueueDataConvert), igual ao app oficial
/// (desktop-sdk/.../fileconverter.h, verificado 2026-09-16). Nunca expoe Process/XML para fora.
/// </summary>
public sealed class X2tConversor(ConfiguracaoMotor cfg)
{
    public static readonly TimeSpan Timeout = TimeSpan.FromMinutes(3);

    /// <summary>Codigo de saida do x2t para arquivo protegido por senha/DRM aberto sem senha (Task 7 do Plano 3,
    /// MEDIDO no x2t 9.4.0.129 — ver <c>Senha_CicloCompleto</c>). E' o mesmo numero que o editor recebe em
    /// <c>asc_initAdvancedOptions</c>.</summary>
    public const int CodigoProtegido = 90;

    /// <summary>Codigo de saida do x2t para senha errada (Task 7 do Plano 3, MEDIDO).</summary>
    public const int CodigoSenhaErrada = 91;

    /// <summary>Codigo de saida do x2t para CSV/TXT aberto SEM codificacao/delimitador (Plano 6, Task 3, MEDIDO no x2t 9.4.0.129):
    /// o app precisa perguntar. E' o mesmo numero que o editor recebe em <c>asc_initAdvancedOptions</c> (o dialogo de importacao).</summary>
    public const int CodigoPrecisaDeOpcoes = 89;

    /// <summary>
    /// Formatos em que o <c>m_sSavePassword</c> do salvar e' aplicado (Task 7 do Plano 3). E' a lista do app
    /// oficial (<c>IsFormatSupportCrypto</c>, desktop-sdk v9.4.0.129 <c>fileconverter.h</c>) cruzada com os formatos
    /// que o mOffice salva, e CONFERIDA no x2t desta versao (<c>Senha_SalvarMantemProtecao_PorFormato</c>: cada
    /// documento sai cifrado — reabrir sem senha da' <see cref="CodigoProtegido"/>).
    ///
    /// PDF fica FORA de proposito (decisao do controlador, fix round 1 da Task 7), embora o x2t o cifre e o
    /// <c>IsFormatSupportCrypto</c> do oficial o liste: os sistemas de peticionamento (PJe, e-SAJ) recusam PDF com
    /// senha, e o Word exporta PDF sem senha. PDF/A tambem (a norma proibe criptografia). O PDF de IMPRESSAO nunca
    /// passa por aqui — <c>GerarPdfDeImpressaoAsync</c> chama <see cref="DeBinAsync"/> sem senha.
    /// </summary>
    ///
    /// Plano 6, Task 2 (MEDIDO no x2t 9.4, <c>Senha_SalvarMantemProtecao_PorFormato</c>): os formatos novos de macro, modelo e
    /// apresentacao de slides (OOXML e ODF) tambem saem cifrados — reabrir sem senha da' 90, com a senha abre. RTF e HTML NAO
    /// guardam senha (o x2t ignora o <c>m_sSavePassword</c>: abriu sem senha), entao ficam fora e o aviso "salvo sem senha" os cobre.
    public static bool FormatoGuardaSenha(FormatoDocumento f) => f is
        FormatoDocumento.Docx or FormatoDocumento.Docm or FormatoDocumento.Dotx or FormatoDocumento.Dotm or
        FormatoDocumento.Xlsx or FormatoDocumento.Xlsm or FormatoDocumento.Xltx or FormatoDocumento.Xltm or
        FormatoDocumento.Pptx or FormatoDocumento.Pptm or FormatoDocumento.Potx or FormatoDocumento.Potm or
        FormatoDocumento.Ppsx or FormatoDocumento.Ppsm or
        FormatoDocumento.Odt or FormatoDocumento.Ott or FormatoDocumento.Ods or FormatoDocumento.Ots or
        FormatoDocumento.Odp or FormatoDocumento.Otp;

    /// <summary>Arquivo do usuario -&gt; <c>Editor.bin</c> na pasta de recuperacao. <paramref name="senha"/> (Task 7 do
    /// Plano 3, parametro novo NO FIM) vira o <c>m_sPassword</c> do x2t, o mesmo elemento que o app oficial escreve
    /// quando o usuario digita a senha de um arquivo protegido (desktop-sdk <c>fileconverter.h</c>). Sem senha, um
    /// arquivo protegido sai com <see cref="ErroConversao.CodigoSaida"/> <see cref="CodigoProtegido"/>; com a senha
    /// errada, <see cref="CodigoSenhaErrada"/>. A senha nunca vai para log nenhum: ela so' existe no XML temporario,
    /// apagado no <c>finally</c> de <see cref="ExecutarAsync"/>.</summary>
    /// <remarks><paramref name="texto"/> (Plano 6, Task 3, parametro novo NO FIM): a codificacao e o delimitador que o usuario escolheu no
    /// dialogo do editor para CSV/TXT; sem ele, o x2t sai com <see cref="CodigoPrecisaDeOpcoes"/>.</remarks>
    /// <remarks><paramref name="formatoCanvas"/> (Plano 6, Task 8, parametro novo NO FIM): o <c>m_nFormatTo</c> — 8192 (generico: o x2t escolhe pelo
    /// tipo do arquivo, o de sempre), 8193 (word), 8194 (cell) ou 8195 (slide). E' o formato que o sdkjs pede em <c>convertFile</c> (Inserir texto do
    /// arquivo = <c>Asc.Xu.ttb</c> = 8193; vinculo externo e importar XML = 8194). Qualquer outro valor vira 8192.</remarks>
    public async Task<string> ParaBinAsync(string arquivoOrigem, string pastaRecovery, CancellationToken ct = default, string? senha = null, OpcoesTexto? texto = null,
        int formatoCanvas = FormatoDocumentoExtensoes.CanvasGenerico)
    {
        if (!File.Exists(arquivoOrigem)) throw new ErroConversao($"Arquivo não encontrado: {arquivoOrigem}", 0, "");
        Directory.CreateDirectory(Path.Combine(pastaRecovery, "media"));
        Directory.CreateDirectory(Path.Combine(pastaRecovery, "changes"));
        var bin = Path.Combine(pastaRecovery, "Editor.bin");
        var tempDir = CriarPastaTemp();
        try
        {
            // SEM m_sAllFontsPath, de proposito (revisao final 2026-09-17). As outras duas conversoes o levam;
            // esta nao, porque o XML do app oficial para ESTE sentido — CASCFileConverterToEditor::ThreadProc,
            // em desktop-sdk/ChromiumBasedEditors/lib/src/fileconverter.h — nao o lista (so' m_sFileFrom,
            // m_sFileTo, m_nFormatTo=8192, m_sThemeDir, m_bDontSaveAdditional, m_sFontDir, m_sTempDir); e' o
            // CASCFileConverterFromEditor (salvar) que traz m_sAllFontsPath. Bate com o experimento da Task 2:
            // o AllFonts.js e' obrigatorio nas conversoes que passam pelo doctrenderer (bin -> arquivo, ->PDF),
            // e a geracao do Editor.bin nao passa por ele. Acrescentar aqui seria divergir da tabela de fatos
            // sem medida que justifique.
            // CSV UTF-8 terminado em quebra de linha: o x2t le' a ULTIMA quebra como uma linha vazia e a regrava como "," a cada
            // abrir -> salvar (MEDIDO; acumulava). O x2t le' uma COPIA na pasta temporaria da conversao (apagada no finally) sem UM terminador
            // final; o arquivo do usuario nunca e' tocado e a copia nao e' alvo de nada.
            var entrada = texto is not null && texto.Codificacao == OpcoesTexto.Utf8 && TabelaFormatos.DaExtensao(arquivoOrigem) is { Formato: FormatoDocumento.Csv }
                ? CopiaSemTerminadorFinal(arquivoOrigem, tempDir) ?? arquivoOrigem : arquivoOrigem;
            var xml = XmlParaBin(entrada, bin, tempDir, senha, texto, formatoCanvas);
            await ExecutarAsync(xml, arquivoOrigem, tempDir, ct);
            if (!File.Exists(bin) || new FileInfo(bin).Length == 0)
                throw new ErroConversao($"Não foi possível converter '{Path.GetFileName(arquivoOrigem)}' para o formato do editor (Editor.bin vazio).", 0, "");
            return bin;
        }
        finally { try { Directory.Delete(tempDir, true); } catch { } }
    }

    /// <summary>Editor.bin (+ changes) da pasta de recuperacao -&gt; <paramref name="arquivoDestino"/> no formato
    /// <paramref name="alvo"/>. <paramref name="jsonParams"/> (Task 4 do Plano 3, parametro novo NO FIM para os
    /// chamadores de antes nao mudarem) vira o <c>m_sJsonParams</c> do x2t — o JSON que o editor manda no
    /// <c>Print</c> sem as opcoes da impressora (<c>documentLayout.isPrint</c> etc.), para o PDF de impressao sair
    /// como o editor imprime.</summary>
    /// <remarks><paramref name="senhaSalvar"/> (Task 7 do Plano 3, parametro novo NO FIM) vira o
    /// <c>m_sSavePassword</c>: o arquivo gravado sai protegido com essa senha — sem ele, um documento aberto com
    /// senha seria salvo SEM protecao. O PDF de impressao (Task 4) nunca passa senha.</remarks>
    /// <paramref name="texto"/> (Plano 6, Task 3, parametro novo NO FIM): codificacao e delimitador do CSV/TXT gravado — substituem os
    /// fixos de antes (46/4), que davam UTF-8 com BOM e virgula (o Excel pt-BR abre numa coluna so').
    public async Task DeBinAsync(string pastaRecovery, FormatoDocumento alvo, string arquivoDestino, CancellationToken ct = default, string? jsonParams = null, string? senhaSalvar = null, OpcoesTexto? texto = null)
    {
        var bin = Path.Combine(pastaRecovery, "Editor.bin");
        if (!File.Exists(bin)) throw new ErroConversao($"Editor.bin não encontrado em {pastaRecovery}", 0, "");
        var temChanges = File.Exists(Path.Combine(pastaRecovery, "changes", "changes0.json"));
        var tempDir = CriarPastaTemp();
        try
        {
            await ExecutarAsync(XmlDeBin(bin, alvo, arquivoDestino, temChanges, tempDir, jsonParams, senhaSalvar, texto), bin, tempDir, ct);
            if (!File.Exists(arquivoDestino) || new FileInfo(arquivoDestino).Length == 0)
                throw new ErroConversao($"Não foi possível salvar como {alvo.Extensao().ToUpperInvariant()} (arquivo de saída vazio).", 0, "");
        }
        finally { try { Directory.Delete(tempDir, true); } catch { } }
    }

    /// <summary>LCID do pt-BR (PL-01, auditoria 2026-10-01). Sem ele o x2t recalcula/grava o cache das formulas de texto
    /// (TEXTO, concatenacao de numeros, datas por extenso) em en-US — e o mPlanilha mostra o cache ao reabrir
    /// ("31/12/Thursday"). MEDIDO: com 1046 sai "31/12/2026", "x1,43", "1.234,50". O x2t.exe aceita a chave.</summary>
    public const int LcidPadrao = 1046;

    /// <summary>XML de parametros do <see cref="DeBinAsync"/> (extraido na Task 4 do Plano 3 para ser testavel
    /// sem rodar o x2t). A ordem dos elementos e' a de antes da extracao; <c>m_sJsonParams</c> so' entra com
    /// <paramref name="jsonParams"/> nao vazio, logo depois de <c>m_nDoctParams</c> — sem ele o XML do salvar
    /// fica identico ao de sempre. <paramref name="senhaSalvar"/> (Task 7) entra como <c>m_sSavePassword</c> logo
    /// depois de <c>m_sJsonParams</c>, so' quando nao nulo nem vazio (vazio = "sem senha", o que o editor manda
    /// depois de remover a protecao). `internal`: nao e' API do conversor.</summary>
    internal XElement XmlDeBin(string bin, FormatoDocumento alvo, string destino, bool temChanges, string tempDir, string? jsonParams, string? senhaSalvar = null, OpcoesTexto? texto = null) =>
        new("TaskQueueDataConvert",
            new XElement("m_sFileFrom", bin),
            new XElement("m_sFileTo", destino),
            new XElement("m_nFormatTo", (int)FormatoParaOX2t(alvo)),
            MarcaPdfA(alvo),
            new XElement("m_bFromChanges", temChanges ? "true" : "false"),
            new XElement("m_bDontSaveAdditional", "true"),
            new XElement("m_sAllFontsPath", Path.Combine(cfg.PastaFontes, "AllFonts.js")),
            new XElement("m_nCsvTxtEncoding", texto?.Codificacao ?? 46),
            new XElement("m_nCsvDelimiter", texto?.Delimitador ?? 4),
            string.IsNullOrEmpty(texto?.DelimitadorChar) ? null : new XElement("m_nCsvDelimiterChar", texto!.DelimitadorChar),
            new XElement("m_nDoctParams", 1),
            new XElement("m_nLcid", texto?.Lcid ?? LcidPadrao),
            string.IsNullOrEmpty(jsonParams) ? null : new XElement("m_sJsonParams", jsonParams),
            string.IsNullOrEmpty(senhaSalvar) ? null : new XElement("m_sSavePassword", senhaSalvar),
            ElementosComuns(tempDir));

    /// <summary>Copia de <paramref name="arquivo"/> em <paramref name="pasta"/> (mesmo nome, mesma extensao) SEM o ultimo terminador de registro
    /// (<c>\r\n</c> ou <c>\n</c>; so' UM). <c>null</c> se o arquivo nao termina em quebra de linha (nada a tirar: o x2t le' o original).
    /// MEDIDO: UTF-16 LE/BE e ANSI (e o TXT) nao acumulam a linha vazia; so' o UTF-8 (<see cref="OpcoesTexto.Utf8"/>). <c>internal</c>: testado
    /// sem rodar o x2t.</summary>
    internal static string? CopiaSemTerminadorFinal(string arquivo, string pasta)
    {
        using var origem = File.OpenRead(arquivo);
        var tamanho = origem.Length;
        var cauda = new byte[(int)Math.Min(2, tamanho)];
        origem.Seek(tamanho - cauda.Length, SeekOrigin.Begin);
        origem.ReadExactly(cauda);
        var tirar = cauda.Length >= 2 && cauda[^2] == (byte)'\r' && cauda[^1] == (byte)'\n' ? 2
                  : cauda.Length >= 1 && cauda[^1] == (byte)'\n' ? 1 : 0;
        if (tirar == 0) return null;
        var destino = Path.Combine(pasta, Path.GetFileName(arquivo));
        origem.Seek(0, SeekOrigin.Begin);
        using (var saida = File.Create(destino))
        {
            var restante = tamanho - tirar;
            var buffer = new byte[81920];
            while (restante > 0)
            {
                var n = origem.Read(buffer, 0, (int)Math.Min(buffer.Length, restante));
                if (n <= 0) break;
                saida.Write(buffer, 0, n);
                restante -= n;
            }
        }
        return destino;
    }

    /// <summary>XML de parametros do <see cref="ParaBinAsync"/> (extraido na Task 3 do Plano 6, como o <see cref="XmlDeBin"/>, para ser
    /// testavel sem rodar o x2t). Com <paramref name="texto"/> entram <c>m_nCsvTxtEncoding</c>, <c>m_nCsvDelimiter</c> e <c>m_nLcid</c>;
    /// sem ele, nenhum dos tres (o XML de todos os outros formatos fica identico ao de sempre).</summary>
    internal XElement XmlParaBin(string arquivoOrigem, string bin, string tempDir, string? senha, OpcoesTexto? texto,
        int formatoCanvas = FormatoDocumentoExtensoes.CanvasGenerico) =>
        new("TaskQueueDataConvert",
            new XElement("m_sFileFrom", arquivoOrigem),
            new XElement("m_sFileTo", bin),
            new XElement("m_nFormatTo", FormatoCanvasValido(formatoCanvas)),
            new XElement("m_bDontSaveAdditional", "true"),
            senha is null ? null : new XElement("m_sPassword", senha),
            // MEDIDO no x2t 9.4.0.129: SEM m_nFormatFrom (a hipotese do plano de que seria preciso caiu): o x2t deduz CSV/TXT pela extensao;
            // COM ele o mesmo pedido sai com 88. m_nCsvTxtEncoding e' o ID da lista de codificacoes do editor (44 = windows-1252,
            // 46 = UTF-8), nao o codepage do Windows: 1252 crasha o x2t (0xC0000005).
            texto is null ? null : new XElement("m_nCsvTxtEncoding", texto.Codificacao),
            texto is null ? null : new XElement("m_nCsvDelimiter", texto.Delimitador),
            string.IsNullOrEmpty(texto?.DelimitadorChar) ? null : new XElement("m_nCsvDelimiterChar", texto!.DelimitadorChar),
            texto is null ? null : new XElement("m_nLcid", texto.Lcid),
            ElementosComuns(tempDir));

    /// <summary>8192–8195 (AVS_OFFICESTUDIO_FILE_CANVAS[_WORD|_SPREADSHEET|_PRESENTATION]); qualquer outro -&gt; 8192 (Plano 6, Task 8).</summary>
    internal static int FormatoCanvasValido(int formato) =>
        formato is >= FormatoDocumentoExtensoes.CanvasGenerico and <= FormatoDocumentoExtensoes.CanvasGenerico + 3 ? formato : FormatoDocumentoExtensoes.CanvasGenerico;

    /// <summary>PDF/A (Plano 4, Task 2) — MEDIDO no x2t 9.4: <c>m_nFormatTo = 521</c> sai com codigo 88 (nao converte),
    /// tanto a partir do <c>Editor.bin</c> (<see cref="DeBinAsync"/>) quanto de arquivo a arquivo (<see cref="ConverterAsync"/>). O x2t le' o PDF/A como <c>m_nFormatTo = 513</c> (PDF) + <c>m_bIsPDFA = true</c>
    /// (a convencao do DocumentServer; a chave existe no x2t.exe e o PdfFile.dll grava o XMP
    /// <c>pdfaid:part 2</c> / <c>conformance A</c>). Os demais formatos passam como estao.</summary>
    internal static FormatoDocumento FormatoParaOX2t(FormatoDocumento alvo) =>
        alvo == FormatoDocumento.PdfA ? FormatoDocumento.Pdf : alvo;

    /// <summary>O <c>m_bIsPDFA</c> que acompanha <see cref="FormatoParaOX2t"/>: so' no PDF/A (<c>null</c> = sem elemento).</summary>
    private static XElement? MarcaPdfA(FormatoDocumento alvo) =>
        alvo == FormatoDocumento.PdfA ? new XElement("m_bIsPDFA", "true") : null;

    public async Task ConverterAsync(string arquivoOrigem, FormatoDocumento alvo, string arquivoDestino, CancellationToken ct = default)
    {
        if (!File.Exists(arquivoOrigem)) throw new ErroConversao($"Arquivo não encontrado: {arquivoOrigem}", 0, "");
        var tempDir = CriarPastaTemp();
        try
        {
            var xml = new XElement("TaskQueueDataConvert",
                new XElement("m_sFileFrom", arquivoOrigem),
                new XElement("m_sFileTo", arquivoDestino),
                new XElement("m_nFormatTo", (int)FormatoParaOX2t(alvo)),
                MarcaPdfA(alvo),
                new XElement("m_bDontSaveAdditional", "true"),
                new XElement("m_sAllFontsPath", Path.Combine(cfg.PastaFontes, "AllFonts.js")),
                new XElement("m_nCsvTxtEncoding", 46),
                new XElement("m_nCsvDelimiter", 4),
                new XElement("m_nLcid", LcidPadrao),
                ElementosComuns(tempDir));
            await ExecutarAsync(xml, arquivoOrigem, tempDir, ct);
            if (!File.Exists(arquivoDestino) || new FileInfo(arquivoDestino).Length == 0)
                throw new ErroConversao($"Não foi possível converter '{Path.GetFileName(arquivoOrigem)}' para {alvo.Extensao().ToUpperInvariant()}.", 0, "");
        }
        finally { try { Directory.Delete(tempDir, true); } catch { } }
    }

    /// <summary>Cria (se preciso) a pasta-pai cfg.PastaTemp e uma subpasta unica por chamada (m_sTempDir do x2t).
    /// O chamador e' responsavel por apagar a subpasta retornada; pastaRecovery (Editor.bin/media/changes) nunca e' tocada aqui.</summary>
    private string CriarPastaTemp()
    {
        Directory.CreateDirectory(cfg.PastaTemp);
        var temp = Path.Combine(cfg.PastaTemp, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        return temp;
    }

    private XElement[] ElementosComuns(string tempDir)
    {
        var lista = new List<XElement> { new("m_sFontDir", cfg.PastaFontes), new("m_sTempDir", tempDir) };
        if (Directory.Exists(cfg.PastaTemas)) lista.Add(new XElement("m_sThemeDir", cfg.PastaTemas));
        return lista.ToArray();
    }

    /// <summary>Nome do XML de parametros DENTRO da pasta temporaria da chamada (ver <see cref="ExecutarAsync"/>).</summary>
    internal const string NomeXmlParametros = "parametros.xml";

    /// <param name="tempDir">A pasta temporaria DESTA chamada (<see cref="CriarPastaTemp"/>), apagada pelo chamador no
    /// <c>finally</c>. O XML de parametros vai para dentro dela (M3 da revisao final do Plano 3): ele pode levar a
    /// senha do documento em claro (<c>m_sPassword</c>/<c>m_sSavePassword</c>), e antes ficava na RAIZ do
    /// <c>%TEMP%</c> — onde um processo morto no meio (queda, kill) o deixava para tras, lido por qualquer programa
    /// do usuario. Dentro da pasta da chamada ele some junto com ela. A sobra de uma queda fica em
    /// <c>%TEMP%\mOffice\x2t\&lt;guid&gt;</c> (<see cref="ConfiguracaoMotor.PastaTemp"/>), e a subida do app a varre
    /// (<see cref="VarrerPastasTemporarias"/>).</param>
    private async Task ExecutarAsync(XElement parametros, string contexto, string tempDir, CancellationToken ct)
    {
        if (!File.Exists(cfg.X2t)) throw new ErroConversao($"Conversor x2t.exe não encontrado em '{cfg.X2t}'. Rode tools/onlyoffice/obter-motor.ps1.", 0, "");
        var arquivoXml = Path.Combine(tempDir, NomeXmlParametros);
        // A gravacao fica DENTRO do try (Task 7 do Plano 3): o XML pode levar a senha do documento em claro
        // (m_sPassword/m_sSavePassword), e uma gravacao interrompida no meio (cancelamento, disco cheio) deixaria o
        // arquivo parcial para tras se o finally so' cobrisse a execucao do x2t.
        try
        {
            await File.WriteAllTextAsync(arquivoXml, new XDocument(new XDeclaration("1.0", "utf-8", null), parametros).ToString(), new UTF8Encoding(false), ct);
            var psi = new ProcessStartInfo(cfg.X2t, $"\"{arquivoXml}\"")
            {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardError = true, RedirectStandardOutput = true,
                WorkingDirectory = Path.GetDirectoryName(cfg.X2t)!,
            };
            using var p = Process.Start(psi) ?? throw new ErroConversao("Não foi possível iniciar o x2t.exe.", 0, "");
            var stderr = p.StandardError.ReadToEndAsync(ct);
            var stdout = p.StandardOutput.ReadToEndAsync(ct);
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct); cts.CancelAfter(Timeout);
            try { await p.WaitForExitAsync(cts.Token); }
            catch (OperationCanceledException)
            {
                try { p.Kill(true); } catch { }
                var stderrTexto = await ObterOuVazio(stderr);
                await ObterOuVazio(stdout); // observa a task mesmo sem usar o conteudo, evita excecao nao observada
                throw new ErroConversao($"O conversor demorou mais de {Timeout.TotalSeconds:0}s com '{Path.GetFileName(contexto)}' e foi interrompido.", -1, stderrTexto);
            }
            var stderrFinal = await ObterOuVazio(stderr);
            var stdoutFinal = await ObterOuVazio(stdout);
            if (p.ExitCode != 0)
                throw new ErroConversao($"Não foi possível converter '{Path.GetFileName(contexto)}' (x2t código {p.ExitCode}).", p.ExitCode, stderrFinal + stdoutFinal);
        }
        finally { try { File.Delete(arquivoXml); } catch { } }
    }

    /// <summary>Prefixo dos XMLs de parametros que versoes anteriores gravavam na raiz do <c>%TEMP%</c>.</summary>
    public const string PrefixoXmlAntigo = "moffice-x2t-";

    /// <summary>
    /// Apaga de <paramref name="pasta"/> (a raiz do <c>%TEMP%</c>, na subida do app) os <c>moffice-x2t-*.xml</c> com
    /// mais de <paramref name="idadeMinima"/> (M3 da revisao final do Plano 3): XML de parametros de uma conversao
    /// cujo processo morreu no meio, gravado ali pelas versoes anteriores — pode levar a senha de um documento em
    /// claro. A idade protege uma conversao EM ANDAMENTO de uma instancia antiga ainda aberta. So' o nivel de cima
    /// (nada de subpastas), so' o padrao exato. Nunca lanca. Devolve quantos apagou.
    /// </summary>
    public static int VarrerParametrosOrfaos(string pasta, TimeSpan idadeMinima)
    {
        var apagados = 0;
        try
        {
            if (!Directory.Exists(pasta)) return 0;
            var limite = DateTime.UtcNow - idadeMinima;
            foreach (var arquivo in Directory.EnumerateFiles(pasta, PrefixoXmlAntigo + "*.xml", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    // O padrao do EnumerateFiles tambem casa nomes 8.3 e extensoes mais longas (".xmlx"): confere de novo.
                    var nome = Path.GetFileName(arquivo);
                    if (!nome.StartsWith(PrefixoXmlAntigo, StringComparison.OrdinalIgnoreCase)
                        || !nome.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                    if (File.GetLastWriteTimeUtc(arquivo) >= limite) continue;
                    File.Delete(arquivo);
                    apagados++;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return apagados;
    }

    /// <summary>
    /// Apaga as subpastas de <paramref name="pastaTemp"/> (<see cref="ConfiguracaoMotor.PastaTemp"/>:
    /// <c>%TEMP%\mOffice\x2t\&lt;guid&gt;</c>, uma por conversao) cujo conteudo mais recente e' mais velho que
    /// <paramref name="idadeMinima"/> — o que sobrou de um processo morto no meio de uma conversao (fix round 2 da
    /// revisao final do Plano 3): pode ter o XML de parametros com a senha em claro e pedacos do documento. A idade
    /// protege a conversao EM ANDAMENTO de outra instancia (o limite do x2t e' 3 min). Links/juncoes (a raiz ou uma
    /// subpasta) nunca sao seguidos. Nunca lanca. Devolve quantas apagou.
    /// </summary>
    public static int VarrerPastasTemporarias(string pastaTemp, TimeSpan idadeMinima)
    {
        var apagadas = 0;
        try
        {
            if (!Directory.Exists(pastaTemp) || EhLink(pastaTemp)) return 0;
            var limite = DateTime.UtcNow - idadeMinima;
            foreach (var pasta in Directory.EnumerateDirectories(pastaTemp))
            {
                try
                {
                    if (EhLink(pasta)) continue;
                    var maisRecente = Directory.GetLastWriteTimeUtc(pasta);
                    foreach (var f in Directory.EnumerateFileSystemEntries(pasta, "*", SearchOption.AllDirectories))
                    {
                        var t = File.GetLastWriteTimeUtc(f);
                        if (t > maisRecente) maisRecente = t;
                    }
                    if (maisRecente >= limite) continue;
                    Directory.Delete(pasta, true);
                    apagadas++;
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return apagadas;
    }

    private static bool EhLink(string pasta)
    {
        var info = new DirectoryInfo(pasta);
        return info.LinkTarget is not null || info.Attributes.HasFlag(FileAttributes.ReparsePoint);
    }

    private static async Task<string> ObterOuVazio(Task<string> tarefa)
    {
        try { return await tarefa; } catch { return ""; }
    }
}
