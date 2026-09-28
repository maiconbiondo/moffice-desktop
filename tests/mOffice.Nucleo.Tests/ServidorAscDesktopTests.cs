using mOffice.Nucleo;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Contencao do esquema nativo ascdesktop://fonts/. Sem janela: e' so resolucao de caminho — e e' exatamente
/// onde um erro vira leitura de arquivo arbitrario, entao cada eixo de fuga tem um caso proprio (travessia,
/// raiz irma com o mesmo prefixo textual, caminho relativo, arquivo inexistente, prefixo errado).
/// </summary>
public class ServidorAscDesktopTests : IDisposable
{
    private readonly string _raiz = Directory.CreateTempSubdirectory("moffice-fontes-").FullName;
    private readonly string _irma;
    private readonly ServidorAscDesktop _servidor;
    private readonly Xunit.Abstractions.ITestOutputHelper _saida;

    public ServidorAscDesktopTests(Xunit.Abstractions.ITestOutputHelper saida)
    {
        _saida = saida;
        // Irma com o MESMO prefixo textual da raiz permitida: "C:\...\raiz" vs "C:\...\raiz-fora".
        _irma = _raiz + "-fora";
        Directory.CreateDirectory(_irma);
        File.WriteAllText(Path.Combine(_raiz, "arial.ttf"), "fonte");
        File.WriteAllText(Path.Combine(_irma, "segredo.ttf"), "nao deveria sair");
        _servidor = new ServidorAscDesktop([_raiz]);
    }

    public void Dispose()
    {
        try { Directory.Delete(_raiz, true); } catch { }
        try { Directory.Delete(_irma, true); } catch { }
        GC.SuppressFinalize(this);
    }

    private string Uri(string caminho) => "ascdesktop://fonts/" + caminho.Replace('\\', '/');

    [Fact] // controle positivo: sem ele os negativos abaixo passariam mesmo com o servidor sempre recusando
    public void ArquivoDentroDaRaiz_EServido()
        => Assert.Equal(Path.Combine(_raiz, "arial.ttf"), _servidor.ResolverArquivo(Uri(Path.Combine(_raiz, "arial.ttf"))));

    [Fact]
    public void ArquivoDentroDaRaiz_ComEscapePercentual_EServido()
        => Assert.Equal(Path.Combine(_raiz, "arial.ttf"),
            _servidor.ResolverArquivo("ascdesktop://fonts/" + _raiz.Replace('\\', '/').Replace(":", "%3A") + "/arial.ttf"));

    [Fact]
    public void Travessia_ComPontoPonto_ERecusada()
        => Assert.Null(_servidor.ResolverArquivo(Uri(Path.Combine(_raiz, "..", Path.GetFileName(_irma), "segredo.ttf"))));

    [Fact] // achado 6 da revisao 2026-09-17: a MESMA travessia, mas com "..\" percent-encoded (%2e%2e%5c) —
           // Uri.UnescapeDataString roda ANTES do Replace('/', separador); se rodasse depois, ou se a ordem
           // fosse trocada, um "%2e%2e%5c" cru poderia sobreviver ao filtro e so virar ".." tarde demais.
    public void Travessia_ComPontoPontoPercentEncoded_ERecusada()
    {
        var alvo = Path.Combine(_raiz, "..", Path.GetFileName(_irma), "segredo.ttf");
        var bruto = alvo.Replace("..", "%2e%2e").Replace("\\", "%5c");
        Assert.Null(_servidor.ResolverArquivo("ascdesktop://fonts/" + bruto));
    }

    [Fact] // achado 6: caminho UNC (\\servidor\share\...) tem de ser recusado como qualquer outro fora da raiz —
           // e' fully-qualified (Path.IsPathFullyQualified aceita), entao so a comparacao com a raiz protege.
    public void CaminhoUnc_ERecusado()
        => Assert.Null(_servidor.ResolverArquivo(Uri(@"\\servidor\share\x.ttf")));

    [Fact] // achado 6: prefixo de dispositivo \\?\ (caminho estendido do Win32) apontando pra fora da raiz —
           // prova que o prefixo verbatim nao muda o resultado da comparacao com a raiz permitida.
    public void CaminhoComPrefixoDeDispositivo_ERecusado()
        => Assert.Null(_servidor.ResolverArquivo(Uri(@"\\?\" + Path.Combine(_irma, "segredo.ttf"))));

    [Fact] // raiz irma com o mesmo prefixo textual: StartsWith sem separador deixaria passar
    public void RaizIrmaComMesmoPrefixo_ERecusada()
        => Assert.Null(_servidor.ResolverArquivo(Uri(Path.Combine(_irma, "segredo.ttf"))));

    [Fact]
    public void CaminhoRelativo_ERecusado()
        => Assert.Null(_servidor.ResolverArquivo(Uri("arial.ttf")));

    [Fact]
    public void ArquivoInexistenteDentroDaRaiz_ERecusado()
        => Assert.Null(_servidor.ResolverArquivo(Uri(Path.Combine(_raiz, "nao-existe.ttf"))));

    [Fact]
    public void OutroPrefixoDoEsquema_ERecusado()
        => Assert.Null(_servidor.ResolverArquivo("ascdesktop://plugin_content/" + _raiz.Replace('\\', '/') + "/arial.ttf"));

    [Fact] // a pasta de recovery muda a cada documento: entra por delegate, e nao na lista fixa
    public void PastaDeRecovery_EntraComoRaizDinamica()
    {
        var recovery = Directory.CreateTempSubdirectory("moffice-rec-").FullName;
        try
        {
            var arquivo = Path.Combine(recovery, "embutida.ttf");
            File.WriteAllText(arquivo, "fonte");
            Assert.Null(_servidor.ResolverArquivo(Uri(arquivo)));
            _servidor.RaizExtra = () => recovery;
            Assert.Equal(arquivo, _servidor.ResolverArquivo(Uri(arquivo)));
        }
        finally { try { Directory.Delete(recovery, true); } catch { } }
    }

    // -----------------------------------------------------------------------------------------------------
    // Segunda lista da contencao (Task 7): o AllFonts.js que o ONLYOFFICE oficial gera nesta maquina nomeia
    // fontes FORA de qualquer raiz fixa — %LOCALAPPDATA%\Microsoft\FontCache\...\CloudFonts (Microsoft 365) e
    // C:\Program Files\ONLYOFFICE\DesktopEditors\fonts. Sem elas o editor `slide` nem abria uma PPTX com tema
    // Aptos. A whitelist e' por ARQUIVO: estar na mesma pasta de um arquivo do catalogo nao basta.
    // -----------------------------------------------------------------------------------------------------

    [Fact] // controle positivo da whitelist: fora de toda raiz, mas nomeado pelo catalogo -> servido
    public void ArquivoForaDasRaizes_MasNoCatalogo_EServido()
    {
        var fonte = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(fonte, "fonte");
        var servidor = new ServidorAscDesktop([_raiz], [fonte]);
        Assert.Equal(fonte, servidor.ResolverArquivo(Uri(fonte)));
    }

    [Fact] // a whitelist NAO promove a pasta: o vizinho do arquivo catalogado continua recusado
    public void VizinhoDeArquivoDoCatalogo_ContinuaRecusado()
    {
        var fonte = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(fonte, "fonte");
        var servidor = new ServidorAscDesktop([_raiz], [fonte]);
        Assert.Null(servidor.ResolverArquivo(Uri(Path.Combine(_irma, "segredo.ttf"))));
    }

    [Fact] // catalogo nomeando arquivo que nao existe mais no disco: File.Exists continua valendo
    public void ArquivoDoCatalogoInexistente_ERecusado()
        => Assert.Null(new ServidorAscDesktop([_raiz], [Path.Combine(_irma, "sumiu.ttf")]).ResolverArquivo(Uri(Path.Combine(_irma, "sumiu.ttf"))));

    [Fact] // travessia continua recusada mesmo com catalogo ligado (o catalogo so ACRESCENTA arquivos exatos)
    public void Travessia_ComCatalogo_ContinuaRecusada()
    {
        var fonte = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(fonte, "fonte");
        var servidor = new ServidorAscDesktop([_raiz], [fonte]);
        Assert.Null(servidor.ResolverArquivo(Uri(Path.Combine(_raiz, "..", Path.GetFileName(_irma), "segredo.ttf"))));
    }

    [Fact] // formato real do AllFonts.js: window["__fonts_files"] = [ "...", ... ]; entrada relativa e' descartada
    public void LerCatalogo_LeOsCaminhosAbsolutosDoAllFontsJs()
    {
        var js = Path.Combine(_raiz, "AllFonts.js");
        File.WriteAllText(js, """
            window["__all_fonts_js_version__"] = 2;

            window["__fonts_files"] = [
            "C:/Program Files/ONLYOFFICE/DesktopEditors/fonts/ASC.ttf",
            "C:/WINDOWS/Fonts/UbuntuMono[wght].ttf",
            "C:/Users/alguem/AppData/Local/Microsoft/FontCache/4/CloudFonts/Aptos/1.ttf",
            "relativa/ignorada.ttf"
            ];
            window["__fonts_infos"] = ["nao deve entrar"];
            """);
        var lista = ServidorAscDesktop.LerCatalogo(js);
        // O ']' do meio e' parte de um NOME DE FONTE (existe de verdade: UbuntuMono[wght].ttf). Procurar o
        // primeiro ']' cru cortava o catalogo desta maquina em 47 de 630 arquivos, em silencio.
        Assert.Equal(3, lista.Count);
        Assert.Contains(lista, c => c.Contains("UbuntuMono[wght]", StringComparison.Ordinal));
        Assert.Contains(lista, c => c.Contains("CloudFonts", StringComparison.Ordinal));
        Assert.Contains(lista, c => c.EndsWith("ASC.ttf", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(lista, c => c.Contains("ignorada", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(lista, c => c.Contains("nao deve entrar", StringComparison.OrdinalIgnoreCase));
    }

    [Fact] // sem catalogo no disco o servidor degrada para "so raizes" em vez de quebrar
    public void LerCatalogo_ArquivoAusente_DevolveVazio()
        => Assert.Empty(ServidorAscDesktop.LerCatalogo(Path.Combine(_raiz, "nao-existe-AllFonts.js")));

    // -----------------------------------------------------------------------------------------------------
    // Fix review 1, achado 1: o parser tem de falhar FECHADO. O AllFonts.js e' escrito por OUTRO programa
    // (o ONLYOFFICE oficial) numa pasta gravavel pelo usuario; ler um arquivo truncado — ou lido no meio da
    // regravacao — nao pode virar "autoriza tudo que parece caminho neste arquivo".
    // -----------------------------------------------------------------------------------------------------

    /// <summary>Escreve um AllFonts.js com o cabecalho real e as entradas dadas; <paramref name="terminar"/>
    /// false omite o `];` (arquivo truncado).</summary>
    private string EscreverCatalogo(string nome, bool terminar, params string[] entradas)
    {
        var js = Path.Combine(_raiz, nome);
        var corpo = string.Join(",\n", entradas.Select(e => "\"" + e.Replace('\\', '/') + "\""));
        File.WriteAllText(js, "window[\"__all_fonts_js_version__\"] = 2;\n\nwindow[\"__fonts_files\"] = [\n"
            + corpo + (terminar ? "\n];\n" : "\n"));
        return js;
    }

    [Fact] // array sem o `];`: zero fontes autorizadas, E um caminho que ESTA listado continua recusado.
           // Controle de mutacao: tirando o `if (!fim.Success) return [];` de LerCatalogo, a varredura vai ate
           // o fim do arquivo, a lista deixa de ser vazia e as duas asserções caem.
    public void LerCatalogo_ArrayNaoTerminado_NaoAutorizaNadaEAvisa()
    {
        var fonte = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(fonte, "fonte");
        var js = EscreverCatalogo("AllFonts-truncado.js", terminar: false, fonte);

        var avisos = new List<string>();
        var lista = ServidorAscDesktop.LerCatalogo(js, avisos.Add);

        Assert.Empty(lista);
        Assert.Contains(avisos, a => a.Contains("malformado", StringComparison.OrdinalIgnoreCase));
        // O que importa no fim: o arquivo LISTADO num catalogo malformado nao e' servido.
        Assert.Null(new ServidorAscDesktop([_raiz], lista).ResolverArquivo(Uri(fonte)));
    }

    [Fact] // controle positivo do caso acima: o MESMO catalogo, agora terminado, autoriza a mesma fonte —
           // prova que o "recusado" de cima vem do terminador ausente e nao do arquivo/caminho em si.
    public void LerCatalogo_MesmoArrayTerminado_AutorizaAFonte()
    {
        var fonte = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(fonte, "fonte");
        var js = EscreverCatalogo("AllFonts-ok.js", terminar: true, fonte);

        var avisos = new List<string>();
        var lista = ServidorAscDesktop.LerCatalogo(js, avisos.Add);

        Assert.Equal(fonte, Path.GetFullPath(Assert.Single(lista)));
        Assert.Empty(avisos);
        Assert.Equal(fonte, new ServidorAscDesktop([_raiz], lista).ResolverArquivo(Uri(fonte)));
    }

    [Fact] // entrada com extensao que nao e' de fonte: nunca entra na whitelist, nem num catalogo bem formado.
           // (Medido no AllFonts.js real desta maquina: 630 entradas, so .ttf/.otf/.ttc — o filtro nao perde
           // nenhuma fonte de verdade.)
    public void LerCatalogo_EntradaSemExtensaoDeFonte_ERecusada()
    {
        var dll = Path.Combine(_irma, "carga.dll");
        var ttf = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(dll, "binario"); File.WriteAllText(ttf, "fonte");
        var js = EscreverCatalogo("AllFonts-com-dll.js", terminar: true, dll, ttf);

        var lista = ServidorAscDesktop.LerCatalogo(js);
        Assert.Equal(ttf, Path.GetFullPath(Assert.Single(lista)));

        var servidor = new ServidorAscDesktop([_raiz], lista);
        Assert.Null(servidor.ResolverArquivo(Uri(dll)));      // negativo
        Assert.Equal(ttf, servidor.ResolverArquivo(Uri(ttf))); // positivo, no mesmo catalogo
    }

    [Theory] // as tres extensoes que o catalogo real usa, em maiuscula: a comparacao e' sem caixa
    [InlineData("MAIUSCULA.TTF")] [InlineData("Postscript.OTF")] [InlineData("Colecao.TtC")]
    public void LerCatalogo_ExtensaoDeFonteEmQualquerCaixa_EAceita(string nome)
    {
        var fonte = Path.Combine(_irma, nome);
        File.WriteAllText(fonte, "fonte");
        var lista = ServidorAscDesktop.LerCatalogo(EscreverCatalogo("AllFonts-caixa.js", terminar: true, fonte));
        Assert.Equal(fonte, Path.GetFullPath(Assert.Single(lista)));
    }

    [Fact] // nome com VIRGULA: e' permitido em nome de arquivo no Windows e o parser tem de manter a entrada
           // inteira (a delimitacao e' pelas aspas). Aspas escapadas (\") nao sao testadas porque `"` e'
           // caractere PROIBIDO em nome de arquivo no Windows — nenhum caminho real pode conter uma.
    public void LerCatalogo_NomeComVirgula_EntraInteiro()
    {
        var fonte = Path.Combine(_irma, "Arial, Narrow.ttf");
        File.WriteAllText(fonte, "fonte");
        var lista = ServidorAscDesktop.LerCatalogo(EscreverCatalogo("AllFonts-virgula.js", terminar: true, fonte));
        Assert.Equal(fonte, Path.GetFullPath(Assert.Single(lista)));
    }

    // -----------------------------------------------------------------------------------------------------
    // Revisao final 2026-09-17: o filtro de extensao existia SO' na leitura do catalogo, e o ramo das RAIZES
    // nao tinha filtro nenhum — uma das raizes reais e' a pasta do motor, que guarda x2t.exe e as DLLs do
    // core. E o caminho UNC so' era barrado pela comparacao com a raiz: num CATALOGO adulterado ele passava
    // direto para o File.Exists, que num host inexistente fica segundos no timeout de SMB, na thread da UI.
    // -----------------------------------------------------------------------------------------------------

    [Fact] // DENTRO da raiz permitida, mas executavel: recusado pela extensao. Controle de mutacao: tirando a
           // linha de ExtensoesServiveis de ResolverArquivo, este teste cai e o positivo do .ttf continua verde.
    public void ExecutavelDentroDaRaiz_ERecusado()
    {
        var exe = Path.Combine(_raiz, "x2t.exe");
        File.WriteAllText(exe, "MZ");
        Assert.Null(_servidor.ResolverArquivo(Uri(exe)));
        // positivo pareado, na MESMA pasta: prova que o "recusado" veio da extensao, nao da raiz.
        Assert.Equal(Path.Combine(_raiz, "arial.ttf"), _servidor.ResolverArquivo(Uri(Path.Combine(_raiz, "arial.ttf"))));
    }

    [Theory] // os outros tipos que moram na pasta do motor / de recovery e nunca podem sair pela pagina
    [InlineData("core.dll")] [InlineData("x2t.config")] [InlineData("Editor.bin.json")] [InlineData("fonts.log")]
    public void ArquivoNaoServivelDentroDaRaiz_ERecusado(string nome)
    {
        var arquivo = Path.Combine(_raiz, nome);
        File.WriteAllText(arquivo, "conteudo");
        Assert.Null(_servidor.ResolverArquivo(Uri(arquivo)));
    }

    [Fact] // o thumbnail de fontes E' pedido de verdade pelo editor (311 vezes nos logs de egresso desta
           // maquina, sempre fonts_thumbnail*.png, na pasta do cache de fontes — que e' uma das raizes).
           // Sem este caso, apertar a lista so' para .ttf/.otf/.ttc passaria na suite e quebraria o editor.
    public void ThumbnailPngDentroDaRaiz_EServido()
    {
        var png = Path.Combine(_raiz, "fonts_thumbnail@2x.png");
        File.WriteAllText(png, "PNG");
        Assert.Equal(png, _servidor.ResolverArquivo(Uri(png)));
    }

    [Fact] // achado 1 da revisao final (fix review 2): Editor.bin mora justamente na raiz DINAMICA (a pasta de
           // recovery, via RaizExtra) — e' o documento inteiro do usuario, em claro. A medicao da revisao
           // anterior mostrou que o .bin NUNCA foi pedido por ascdesktop:// nesta maquina; admiti-lo em
           // ExtensoesServiveis "por precaucao" reabria, pelo esquema de fontes, a leitura arbitraria que o
           // filtro de extensao existe para fechar. Quem le/grava o Editor.bin e' PonteDesktop/X2tConversor,
           // nunca este servidor.
    public void EditorBinDentroDaPastaDeRecovery_ERecusado()
    {
        var recovery = Directory.CreateTempSubdirectory("moffice-rec-").FullName;
        try
        {
            var bin = Path.Combine(recovery, "Editor.bin");
            File.WriteAllText(bin, "documento inteiro em claro");
            _servidor.RaizExtra = () => recovery;
            Assert.Null(_servidor.ResolverArquivo(Uri(bin)));
        }
        finally { try { Directory.Delete(recovery, true); } catch { } }
    }

    [Theory] // as duas grafias Win32 de UNC, no CATALOGO: descartadas na leitura, antes de virar whitelist.
    [InlineData(@"\\servidor\share\remota.ttf")]
    [InlineData(@"\\?\UNC\servidor\share\remota.ttf")]
    public void LerCatalogo_EntradaUnc_EDescartada(string unc)
    {
        var local = Path.Combine(_irma, "nuvem.ttf");
        File.WriteAllText(local, "fonte");
        var lista = ServidorAscDesktop.LerCatalogo(EscreverCatalogo("AllFonts-unc.js", terminar: true, unc, local));
        // so' a local sobrou — o positivo pareado prova que o catalogo foi lido e que o descarte foi da UNC.
        Assert.Equal(local, Path.GetFullPath(Assert.Single(lista)));
    }

    [Theory] // pedido UNC com a entrada JA' na whitelist (o pior caso: o ramo do catalogo nao compara raiz).
             // O que se afirma nao e' so' "recusado": e' recusado SEM TOCAR NA REDE. 500 ms e' folgado — um
             // File.Exists num host SMB inexistente leva segundos (5,4 s medidos na revisao final).
    [InlineData(@"\\servidor-que-nao-existe-moffice\share\remota.ttf")]
    [InlineData(@"\\?\UNC\servidor-que-nao-existe-moffice\share\remota.ttf")]
    public void PedidoUnc_ERecusadoSemTocarNaRede(string unc)
    {
        // O catalogo entra pelo CONSTRUTOR (e nao por LerCatalogo), de proposito: assim o teste prova a
        // defesa de ResolverArquivo mesmo que a da leitura do catalogo seja removida.
        var servidor = new ServidorAscDesktop([_raiz], [unc]);
        var relogio = System.Diagnostics.Stopwatch.StartNew();
        var r = servidor.ResolverArquivo(Uri(unc));
        relogio.Stop();
        _saida.WriteLine($"UNC recusada em {relogio.ElapsedMilliseconds} ms: {unc}");
        Assert.Null(r);
        Assert.True(relogio.ElapsedMilliseconds < 500, $"resolucao levou {relogio.ElapsedMilliseconds} ms — sinal de que houve I/O de rede");
    }
}
