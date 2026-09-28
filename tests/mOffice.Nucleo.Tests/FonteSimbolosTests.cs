using System.Buffers.Binary;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 1 do Marco 1 (decisao 1 do gate, 2026-09-19): a fonte de simbolos internos que o motor traz embutida em
/// base64 nos 4 <c>sdk-all.js</c> (familia "ASCW3" da Ascensio, sem licenca localizavel) e' trocada pela NOSSA,
/// <c>tools/fontes/mOffice-Simbolos.ttf</c> (OFL 1.1, desenhada em codigo). Tres camadas de prova:
///
/// <list type="bullet">
/// <item><b>Unidade</b> (sem motor, sem janela): a fonte commitada tem exatamente os cmaps/glifos que o motor
///   pressupoe, nao menciona marca de terceiros, e' pequena e e' REPRODUZIVEL byte a byte pelo script Python.</item>
/// <item><b>Guarda do pacote</b> (precisa do motor): nenhum dos 4 bundles contem mais a fonte original (procurada
///   pelo CONTEUDO: sha256 do trecho decodificado) e todos contem a nossa 1x. Mais a prova de que o detector
///   DISPARA no bundle original (<c>.orig</c>), senao "zero ocorrencias" nao provaria nada.</item>
/// <item><b>Janela real</b> (sessao console): documento sintetico com tabulacoes, quebra de linha e caixas de
///   selecao; ¶ ligado; captura. CONTROLE NEGATIVO: a mesma janela sobre um motor temporario cujo unico
///   bundle diferente e' o <c>word/sdk-all.js</c> com uma fonte de MESMOS cmaps/metricas e contornos VAZIOS
///   — os pixels que somem sao, por construcao, os que vinham dos nossos contornos.</item>
/// </list>
/// </summary>
public class FonteSimbolosTests
{
    /// <summary>sha256 do blob original da Ascensio (o que NAO pode mais estar no pacote). E' o mesmo de
    /// <c>fonts/ASC.ttf</c> — ver docs/licencas/AVISO-fontes.md, secao 1.</summary>
    public const string ShaOriginalAscensio = "35ff902a3b9c9254ce3c57d610dad9e57008b8a3351691c463bee775644c1e18";

    public const string FamiliaEsperada = "mOffice Simbolos";

    /// <summary>Ordem fixa dos glifos no script (indices que os cmaps devem apontar).</summary>
    private static readonly string[] Ordem =
        [".notdef", ".null", "uni000D", "uniF020", "uniF021", "uniF022", "uniF038", "uniF039", "uni25C9", "circle", "uni2610", "uni2611"];

    private static readonly Dictionary<int, string> CmapSimbolicaEsperada = new()
    {
        [0x25C9] = "uni25C9", [0x25CB] = "circle", [0x2610] = "uni2610", [0x2611] = "uni2611",
        [0xF020] = "uniF020", [0xF021] = "uniF021", [0xF022] = "uniF022", [0xF038] = "uniF038", [0xF039] = "uniF039",
    };

    private static readonly Dictionary<int, string> CmapMacEsperada = new()
    {
        [0x20] = "uniF020", [0x21] = "uniF021", [0x22] = "uniF022", [0x38] = "uniF038", [0x39] = "uniF039",
    };

    internal static string RaizRepositorio()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "mOffice.slnx"))) return dir.FullName;
            dir = dir.Parent;
        }
        throw new FileNotFoundException("mOffice.slnx nao encontrado subindo a partir de " + AppContext.BaseDirectory);
    }

    internal static string CaminhoFonte => Path.Combine(RaizRepositorio(), "tools", "fontes", "mOffice-Simbolos.ttf");
    internal static string CaminhoFonteVazia => Path.Combine(RaizRepositorio(), "tests", "mOffice.Nucleo.Tests", "Fixtures", "mOffice-Simbolos-vazia.ttf");
    internal static string CaminhoScript => Path.Combine(RaizRepositorio(), "tools", "fontes", "construir-simbolos.py");

    internal static string Sha256Hex(byte[] dados) => Convert.ToHexString(SHA256.HashData(dados)).ToLowerInvariant();

    // =====================================================================================================
    // Unidade
    // =====================================================================================================

    [Fact] // a fonte commitada tem EXATAMENTE os cmaps e glifos que o motor pressupoe, e e' pequena
    public void Fonte_TemOsCmapsEGlifosEsperados()
    {
        var bytes = File.ReadAllBytes(CaminhoFonte);
        Assert.True(bytes.Length < 10 * 1024, $"fonte com {bytes.Length} bytes — esperado < 10 KB");
        var ttf = TtfMinimo.Ler(bytes);

        Assert.Equal(Ordem.Length, ttf.NumGlyphs);
        Assert.Equal(Ordem, ttf.NomesDosGlifos);
        Assert.Equal(2048, ttf.UnitsPerEm);

        AssertCmap(CmapSimbolicaEsperada, ttf.Cmap(3, 0), "3/0 simbolica");
        AssertCmap(CmapMacEsperada, ttf.Cmap(1, 0), "1/0 (formato 6)");

        // Todo glifo alcancavel desenha alguma coisa — menos o espaco (uniF020), que e' so' avanco. Sem isto,
        // um glifo que sumisse do script mas continuasse no cmap passaria calado.
        foreach (var nome in CmapSimbolicaEsperada.Values.Distinct())
        {
            var contornos = ttf.ContornosDoGlifo(Array.IndexOf(Ordem, nome));
            if (nome == "uniF020") Assert.Equal(0, contornos);
            else Assert.True(contornos > 0, $"glifo {nome} sem contorno");
        }
        // Larguras de avanco iguais as que o motor pressupoe (layout das marcas nao muda).
        Assert.Equal(1825, ttf.AvancoDoGlifo(Array.IndexOf(Ordem, "uniF021")));
        Assert.Equal(1135, ttf.AvancoDoGlifo(Array.IndexOf(Ordem, "uniF038")));
        Assert.Equal(1764, ttf.AvancoDoGlifo(Array.IndexOf(Ordem, "uni2611")));
        Assert.Equal(512, ttf.AvancoDoGlifo(Array.IndexOf(Ordem, "uniF020")));
    }

    private static void AssertCmap(Dictionary<int, string> esperado, Dictionary<int, int>? real, string qual)
    {
        Assert.True(real is not null, $"subtabela cmap {qual} ausente");
        var esperadoPorIndice = esperado.ToDictionary(kv => kv.Key, kv => Array.IndexOf(Ordem, kv.Value));
        var realSemNotdef = real!.Where(kv => kv.Value != 0).ToDictionary(kv => kv.Key, kv => kv.Value);
        Assert.True(esperadoPorIndice.OrderBy(k => k.Key).SequenceEqual(realSemNotdef.OrderBy(k => k.Key)),
            $"cmap {qual} diferente do esperado. Esperado: {Descrever(esperadoPorIndice)}; real: {Descrever(realSemNotdef)}");
    }

    private static string Descrever(Dictionary<int, int> m) => string.Join(", ", m.OrderBy(k => k.Key).Select(k => $"U+{k.Key:X4}->{k.Value}"));

    [Fact] // tabela name nossa: familia, copyright, OFL; nenhuma marca de terceiros em lugar nenhum do arquivo
    public void Fonte_NaoMencionaMarcaDeTerceiros()
    {
        var bytes = File.ReadAllBytes(CaminhoFonte);
        var ttf = TtfMinimo.Ler(bytes);
        var nomes = ttf.Nomes;
        Assert.Contains(nomes, n => n.Id == 1 && n.Texto == FamiliaEsperada);
        Assert.Contains(nomes, n => n.Id == 0 && n.Texto == "Copyright (c) 2026 mOffice");
        Assert.Contains(nomes, n => n.Id == 13 && n.Texto.Contains("SIL Open Font License, Version 1.1"));
        Assert.Contains(nomes, n => n.Id == 14 && n.Texto.Contains("openfontlicense.org"));
        Assert.DoesNotContain(nomes, n => n.Id == 7); // nenhuma marca registrada declarada
        foreach (var n in nomes)
        {
            Assert.DoesNotContain("ASCW3", n.Texto, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ascensio", n.Texto, StringComparison.OrdinalIgnoreCase);
        }
        // E no arquivo inteiro (ASCII e UTF-16BE), nao so' na tabela name.
        foreach (var marca in new[] { "ASCW3", "Ascensio" })
        {
            AssertNaoContemBytes(Encoding.ASCII.GetBytes(marca), bytes, marca + " (ASCII)");
            AssertNaoContemBytes(Encoding.BigEndianUnicode.GetBytes(marca), bytes, marca + " (UTF-16BE)");
        }
    }

    private static void AssertNaoContemBytes(byte[] agulha, byte[] palheiro, string oQue) =>
        Assert.True(palheiro.AsSpan().IndexOf(agulha) < 0, $"'{oQue}' encontrado nos bytes da fonte");

    [Fact] // o controle negativo (fixture) e' a MESMA fonte sem contornos: mesmos cmaps, mesmas larguras, zero tinta
    public void FonteVazia_MesmosCmapsSemContornos()
    {
        var ttf = TtfMinimo.Ler(File.ReadAllBytes(CaminhoFonteVazia));
        Assert.Equal(Ordem, ttf.NomesDosGlifos);
        AssertCmap(CmapSimbolicaEsperada, ttf.Cmap(3, 0), "3/0 simbolica (vazia)");
        AssertCmap(CmapMacEsperada, ttf.Cmap(1, 0), "1/0 (vazia)");
        for (var g = 0; g < ttf.NumGlyphs; g++) Assert.Equal(0, ttf.ContornosDoGlifo(g));
        Assert.Equal(1825, ttf.AvancoDoGlifo(Array.IndexOf(Ordem, "uniF021")));
    }

    [SkippableFact] // mesmo script -> mesmos bytes: a fonte commitada e' exatamente o que o Python produz
    // A igualdade byte a byte esta' presa ao fontTools instalado (4.60.1): uma versao nova pode mudar a
    // serializacao (ordem/alinhamento de tabelas) sem mudar o significado — se isto falhar so' por isso,
    // regenerar a fonte commitada e re-medir, nao afrouxar o teste.
    public void Fonte_ReproduzivelPeloScript()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-fonte-").FullName;
        var regenerada = Path.Combine(pasta, "regenerada.ttf");
        var vazia = Path.Combine(pasta, "vazia.ttf");
        RodarScriptOuPular($"\"{CaminhoScript}\" --saida \"{regenerada}\"");
        RodarScriptOuPular($"\"{CaminhoScript}\" --vazia --saida \"{vazia}\"");
        Assert.Equal(File.ReadAllBytes(CaminhoFonte), File.ReadAllBytes(regenerada));
        Assert.Equal(File.ReadAllBytes(CaminhoFonteVazia), File.ReadAllBytes(vazia));
    }

    private static void RodarScriptOuPular(string argumentos)
    {
        ProcessStartInfo psi = new("python", argumentos) { RedirectStandardError = true, RedirectStandardOutput = true, UseShellExecute = false };
        Process? p;
        try { p = Process.Start(psi); }
        catch (System.ComponentModel.Win32Exception) { Skip.If(true, "python nao encontrado no PATH"); return; }
        Assert.NotNull(p);
        var erro = p.StandardError.ReadToEnd();
        p.StandardOutput.ReadToEnd();
        p.WaitForExit();
        Skip.If(erro.Contains("No module named") && erro.Contains("fontTools"), "fontTools nao instalado (pip install fonttools)");
        Assert.True(p.ExitCode == 0, "construir-simbolos.py falhou: " + erro);
    }

    // =====================================================================================================
    // Guarda do pacote (o artefato que o motor de verdade carrega)
    // =====================================================================================================

    private static readonly string[] Editores = ["word", "cell", "slide", "visio"];

    private static string? PastaMotorOuNula()
    {
        try { return ConfiguracaoMotor.Localizar().PastaMotor; }
        catch (ErroConversao) { return null; }
    }

    /// <summary>Os literais base64 de um bundle que decodificam para um TrueType, com o sha256 de cada um.
    /// Mesmo criterio do <c>aplicar-patches.ps1</c>, implementado de novo em C# (duas implementacoes
    /// independentes conferindo o mesmo artefato).</summary>
    internal static List<(int Posicao, string Sha)> FontesEmbutidas(string textoLatin1)
    {
        var lista = new List<(int, string)>();
        foreach (Match m in Regex.Matches(textoLatin1, "\"(AAEAAA[A-Za-z0-9+/]{1000,}={0,2})\""))
        {
            byte[] bytes;
            try { bytes = Convert.FromBase64String(m.Groups[1].Value); } catch (FormatException) { continue; }
            lista.Add((m.Index, Sha256Hex(bytes)));
        }
        return lista;
    }

    private static string LerLatin1(string caminho) => Encoding.Latin1.GetString(File.ReadAllBytes(caminho));

    [SkippableFact] // guarda: nenhum dos 4 bundles traz a fonte da Ascensio; todos trazem a nossa exatamente 1x
    public void Bundles_NaoContemAFonteDaAscensioEContemANossa()
    {
        var motor = PastaMotorOuNula();
        Skip.If(motor is null, "third_party/onlyoffice ausente — rode tools/onlyoffice/obter-motor.ps1");
        var shaNossa = Sha256Hex(File.ReadAllBytes(CaminhoFonte));
        foreach (var e in Editores)
        {
            var bundle = Path.Combine(motor!, "editors", "sdkjs", e, "sdk-all.js");
            Assert.True(File.Exists(bundle), "bundle ausente: " + bundle);
            var fontes = FontesEmbutidas(LerLatin1(bundle));
            Assert.True(fontes.All(f => f.Sha != ShaOriginalAscensio),
                $"{e}/sdk-all.js ainda contem a fonte original da Ascensio — rode tools/onlyoffice/aplicar-patches.ps1");
            Assert.True(fontes.Count(f => f.Sha == shaNossa) == 1,
                $"{e}/sdk-all.js deveria conter a nossa fonte exatamente 1x (tem {fontes.Count(f => f.Sha == shaNossa)}; fontes embutidas: {fontes.Count})");
        }
    }

    [SkippableFact] // prova de disparo: o detector ACHA a fonte original onde ela existe (bundle .orig guardado pelo patch)
    public void Detector_AchaAFonteOriginalNoBundleOriginal()
    {
        var motor = PastaMotorOuNula();
        Skip.If(motor is null, "third_party/onlyoffice ausente");
        var orig = Path.Combine(motor!, "editors", "sdkjs", "word", "sdk-all.js.orig");
        Skip.If(!File.Exists(orig), "word/sdk-all.js.orig ausente (motor nunca patchado nesta maquina)");
        var fontes = FontesEmbutidas(LerLatin1(orig));
        Assert.Equal(1, fontes.Count(f => f.Sha == ShaOriginalAscensio));
    }
}

/// <summary>Testes de janela real da fonte de simbolos — em serie com os outros (colecao <see cref="JanelaReal"/>).</summary>
[Collection(JanelaReal.Nome)]
public class FonteSimbolosJanelaRealTests
{
    private readonly ITestOutputHelper _saida;
    public FonteSimbolosJanelaRealTests(ITestOutputHelper saida) { _saida = saida; }

    /// <summary>Liga/desliga as marcas de formatacao pela API publica do editor de texto (o mesmo
    /// <c>put_ShowParaMarks</c> que o botao ¶ da barra chama, exportado em <c>sdk-all-min.js</c>).</summary>
    private static string MostrarMarcas(bool ligar) => $$"""
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api) return "sem api";
            if (typeof api.put_ShowParaMarks !== "function") return "sem put_ShowParaMarks";
            api.put_ShowParaMarks({{(ligar ? "true" : "false")}});
            return "ok " + String(api.get_ShowParaMarks());
          } catch (e) { return "erro: " + e; }
        })()
        """;

    /// <summary>Zoom do editor (<c>api.zoom(percent)</c>, o mesmo que o seletor da barra de status chama). As
    /// marcas de formatacao sao desenhadas a 10 pt fixos pelo motor; a 300 % cada uma vira centenas de pixels.</summary>
    private static string Ampliar(int percentual) => $$"""
        (function () {
          try {
            var w = window.frames[0];
            var api = (w.Asc && w.Asc.editor) ? w.Asc.editor : w.editor;
            if (!api || typeof api.zoom !== "function") return "sem api.zoom";
            api.zoom({{percentual}});
            return "ok";
          } catch (e) { return "erro: " + e; }
        })()
        """;

    private const int ZoomDaCaptura = 300;

    [SkippableFact] // janela real: as marcas (¶ ligado) sao desenhadas com a NOSSA fonte — controle negativo com a fonte VAZIA
    public void MarcasDeFormatacao_VemDaNossaFonte()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-simbolos-").FullName;
        // Caixas com "MS Gothic", como o Word grava. (Medido nesta janela: pedir a familia interna "ASCW3" como
        // fonte do run NAO chega a fonte embutida — o texto do documento passa pelo catalogo de fontes, onde
        // ASCW3 nao existe, e os simbolos saem de outra fonte. A fonte embutida so' serve as marcas de
        // formatacao e o substituto de ultimo recurso da serializacao, como o AVISO-fontes.md, secao 1.2, diz.)
        var docx = FabricaDocumentos.DocxMarcasDeFormatacao(pasta);
        var motorReal = mOffice.Conversor.Tests.MotorDeTeste.Localizar();

        // O web-apps GUARDA o estado do ¶ no localStorage do perfil (de-show-hiddenchars): uma rodada anterior
        // deixa a janela seguinte ja' com as marcas ligadas. Por isso cada janela DESLIGA antes de capturar
        // "sem marcas" e LIGA de novo — o estado inicial do perfil nao entra na comparacao.
        byte[] semMarcas = [], comMarcas = [];
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
            Assert.Equal("ok", (await j.ExecutarJsAsync(Ampliar(ZoomDaCaptura))).Trim('"'));
            Assert.Equal("ok false", (await j.ExecutarJsAsync(MostrarMarcas(false))).Trim('"'));
            semMarcas = await CapturaEstavel(j);
            var r = (await j.ExecutarJsAsync(MostrarMarcas(true))).Trim('"');
            _saida.WriteLine("put_ShowParaMarks (nossa fonte): " + r);
            Assert.Equal("ok true", r);
            comMarcas = await CapturaEstavel(j);
            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));

        // CONTROLE NEGATIVO: motor temporario = juncoes para tudo, menos word/sdk-all.js com a fonte de contornos
        // vazios. Perfil WebView2 proprio ("sonda-controle-fonte"): o bundle tem a mesma URL nos dois motores, e
        // um cache do navegador serviria o bundle da nossa fonte por cima do vazio — o controle ficaria igual ao
        // positivo e o teste falharia "sem motivo".
        byte[] controle = [];
        using (var motorControle = MotorDeControle.Criar(motorReal, File.ReadAllBytes(FonteSimbolosTests.CaminhoFonte), File.ReadAllBytes(FonteSimbolosTests.CaminhoFonteVazia)))
        {
            _saida.WriteLine("motor de controle: " + motorControle.Pasta);
            var config = new ConfiguracaoMotor(motorControle.Pasta);
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.PrepararParaTeste(new JanelaEditor("sonda-controle-fonte", config));
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal("ok", (await j.ExecutarJsAsync(Ampliar(ZoomDaCaptura))).Trim('"'));
                Assert.Equal("ok false", (await j.ExecutarJsAsync(MostrarMarcas(false))).Trim('"'));
                await CapturaEstavel(j);
                var r = (await j.ExecutarJsAsync(MostrarMarcas(true))).Trim('"');
                _saida.WriteLine("put_ShowParaMarks (fonte vazia): " + r);
                Assert.Equal("ok true", r);
                controle = await CapturaEstavel(j);
                Assert.Empty(j.Guarda.Bloqueadas);
                j.Close();
            }, TimeSpan.FromMinutes(3));
        }

        // JanelaEditor nao expoe mais PastaRelatorios (Task 1 do Marco 1 / Plano 2: log por janela virou
        // Logs.NovoArquivo, e poc/relatorios ficou so' com a Sonda) — este e' artefato de bancada do PROPRIO
        // teste (capturas para comparacao visual), entao ele calcula o caminho por conta propria.
        var relatorios = Path.Combine(FonteSimbolosTests.RaizRepositorio(), "poc", "relatorios");
        Directory.CreateDirectory(relatorios);
        File.WriteAllBytes(Path.Combine(relatorios, "ascw3-livre-formulario.png"), semMarcas);
        File.WriteAllBytes(Path.Combine(relatorios, "ascw3-livre-marcas.png"), comMarcas);
        File.WriteAllBytes(Path.Combine(relatorios, "ascw3-livre-marcas-controle-vazio.png"), controle);

        var a = Pixels.Decodificar(comMarcas);
        var b = Pixels.Decodificar(controle);
        Assert.True(a.Largura == b.Largura && a.Altura == b.Altura, $"capturas de tamanhos diferentes: {a.Largura}x{a.Altura} vs {b.Largura}x{b.Altura}");

        // So' a PAGINA entra na comparacao (abaixo dos 20 % de cima, onde ficam a barra de ferramentas e a regua):
        // o botao ¶ da barra reflete o estado guardado no perfil, nao a chamada da API, e varia entre as janelas.
        var linhaInicial = a.Altura / 5;

        // (i) a regiao das marcas tem pixels desenhados; (ii) com a fonte vazia a MESMA regiao fica sem tinta.
        // O que sobra de diferenca com a fonte vazia e' ruido (cursor piscando): dezenas de pixels, contra as
        // centenas de cada marca com o zoom a 300 % (texto a 12 pt; as marcas sao desenhadas a 10 pt fixos, e'
        // o zoom que as amplia).
        var diff = Pixels.Diferentes(a, b, linhaInicial: linhaInicial);
        var tintaNossa = diff.Count(p => Pixels.Escuro(a, p));
        var tintaVazia = diff.Count(p => Pixels.Escuro(b, p));
        _saida.WriteLine($"pixels diferentes (pagina) entre nossa fonte e fonte vazia: {diff.Count} de {a.Largura * (a.Altura - linhaInicial)}; escuros na nossa: {tintaNossa}; escuros na vazia: {tintaVazia}");
        File.WriteAllBytes(Path.Combine(relatorios, "ascw3-livre-marcas-diff.png"), Pixels.MascaraPng(a, diff));

        // ¶ ligado x desligado na nossa fonte tambem difere (sanidade: as marcas foram mesmo ligadas).
        var ligouMarcas = Pixels.Diferentes(a, Pixels.Decodificar(semMarcas), linhaInicial: linhaInicial).Count;
        _saida.WriteLine($"pixels diferentes (pagina) entre ¶ ligado e desligado (nossa fonte): {ligouMarcas}");
        Assert.True(ligouMarcas >= 200, $"ligar o ¶ mudou so' {ligouMarcas} pixels da pagina — as marcas nao apareceram");

        Assert.True(tintaNossa >= 200, $"as marcas da nossa fonte deveriam ter centenas de pixels de tinta que a fonte vazia nao tem (tem {tintaNossa})");
        Assert.True(tintaNossa > 5 * tintaVazia,
            $"os pixels que diferem deveriam ser tinta da NOSSA fonte, nao da vazia (nossa {tintaNossa}, vazia {tintaVazia})");
    }

    /// <summary>Captura quando duas seguidas, 300 ms apart, saem iguais (o desenho assentou).</summary>
    private static async Task<byte[]> CapturaEstavel(JanelaEditor j)
    {
        var anterior = await j.CapturarPngAsync();
        for (var i = 0; i < 30; i++)
        {
            await Task.Delay(300);
            var atual = await j.CapturarPngAsync();
            if (atual.AsSpan().SequenceEqual(anterior)) return atual;
            anterior = atual;
        }
        return anterior;
    }
}

/// <summary>Motor temporario para o controle negativo, sem tocar no motor compartilhado:
/// <list type="bullet">
/// <item><c>converter</c>, <c>fonts</c>, <c>dictionaries</c>, <c>cache</c>: JUNCOES (mklink /J, sem privilegio) —
///   quem le essas pastas e' codigo nosso (.NET/x2t), que segue juncoes normalmente.</item>
/// <item><c>editors</c>: COPIA de verdade, so' do que o editor de texto pede — <c>web-apps</c>, <c>webext</c>,
///   <c>sdkjs/common</c> e <c>sdkjs/word</c> (~170 MB) — com o <c>word/sdk-all.js</c> reescrito com a fonte de
///   simbolos VAZIA. Tem de ser copia: e' a pasta que <c>SetVirtualHostNameToFolderMapping</c> serve, e o
///   Chromium resolve juncoes e recusa o que cai fora da raiz mapeada (medido: com juncoes a navegacao caia em
///   <c>chrome-error://chromewebdata/</c>).</item>
/// </list>
/// Nada e' escrito no motor: o AllFonts.js vem do catalogo pelo host (<c>CatalogoNoEditor</c>).</summary>
internal sealed class MotorDeControle : IDisposable
{
    public string Pasta { get; }
    private readonly List<string> _juncoes = [];

    private MotorDeControle(string pasta) { Pasta = pasta; }

    public static MotorDeControle Criar(ConfiguracaoMotor real, byte[] fonteNossa, byte[] fonteVazia)
    {
        var m = new MotorDeControle(Directory.CreateTempSubdirectory("moffice-motor-controle-").FullName);
        try
        {
            m.Espelhar(real.PastaMotor, m.Pasta, exceto: "editors");
            var editorsReal = Path.Combine(real.PastaMotor, "editors");
            var editors = Path.Combine(m.Pasta, "editors");
            foreach (var pasta in new[] { "web-apps", "webext", Path.Combine("sdkjs", "common") })
                Copiar(Path.Combine(editorsReal, pasta), Path.Combine(editors, pasta));
            var word = Path.Combine(editors, "sdkjs", "word");
            Directory.CreateDirectory(word);
            foreach (var f in Directory.GetFiles(Path.Combine(editorsReal, "sdkjs", "word")))
            {
                var nome = Path.GetFileName(f);
                if (nome.EndsWith(".orig", StringComparison.OrdinalIgnoreCase)) continue;
                if (nome != "sdk-all.js") { File.Copy(f, Path.Combine(word, nome)); continue; }
                var texto = Encoding.Latin1.GetString(File.ReadAllBytes(f));
                var b64Nossa = Convert.ToBase64String(fonteNossa);
                var b64Vazia = Convert.ToBase64String(fonteVazia);
                var i = texto.IndexOf(b64Nossa, StringComparison.Ordinal);
                if (i < 0 || texto.IndexOf(b64Nossa, i + 1, StringComparison.Ordinal) >= 0)
                    throw new InvalidOperationException("word/sdk-all.js do motor real nao contem a nossa fonte exatamente 1x — rode tools/onlyoffice/aplicar-patches.ps1");
                var trocado = texto[..i] + b64Vazia + texto[(i + b64Nossa.Length)..];
                File.WriteAllBytes(Path.Combine(word, nome), Encoding.Latin1.GetBytes(trocado));
            }
            return m;
        }
        catch { m.Dispose(); throw; }
    }

    private static void Copiar(string origem, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var f in Directory.GetFiles(origem)) File.Copy(f, Path.Combine(destino, Path.GetFileName(f)));
        foreach (var d in Directory.GetDirectories(origem)) Copiar(d, Path.Combine(destino, Path.GetFileName(d)));
    }

    /// <summary>Arquivos: copia. Pastas: juncao — menos <paramref name="exceto"/>, que o chamador monta.</summary>
    private void Espelhar(string origem, string destino, string exceto)
    {
        Directory.CreateDirectory(destino);
        foreach (var f in Directory.GetFiles(origem)) File.Copy(f, Path.Combine(destino, Path.GetFileName(f)));
        foreach (var d in Directory.GetDirectories(origem))
        {
            var nome = Path.GetFileName(d);
            if (string.Equals(nome, exceto, StringComparison.OrdinalIgnoreCase)) continue;
            var link = Path.Combine(destino, nome);
            var psi = new ProcessStartInfo("cmd.exe", $"/c mklink /J \"{link}\" \"{d}\"") { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
            using var p = Process.Start(psi)!;
            var saida = p.StandardOutput.ReadToEnd() + p.StandardError.ReadToEnd();
            p.WaitForExit();
            if (p.ExitCode != 0 || !Directory.Exists(link)) throw new InvalidOperationException($"mklink /J falhou para {link}: {saida}");
            _juncoes.Add(link);
        }
    }

    public void Dispose()
    {
        // Juncoes primeiro, uma a uma e SEM recursao: apagar recursivamente por cima de uma juncao poderia
        // descer no motor real. Depois que nenhuma sobrou, o resto e' so' arquivo nosso.
        foreach (var j in _juncoes.AsEnumerable().Reverse())
        {
            try { if (Directory.Exists(j)) Directory.Delete(j, recursive: false); } catch (IOException) { }
        }
        if (_juncoes.Any(Directory.Exists)) return; // nao arrisca: deixa a pasta temporaria para tras
        try { Directory.Delete(Pasta, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
    }
}

/// <summary>Leitor minimo de TrueType para os testes: tabelas <c>maxp</c>, <c>head</c>, <c>hhea</c>/<c>hmtx</c>,
/// <c>loca</c>/<c>glyf</c> (contagem de contornos), <c>cmap</c> (formatos 4 e 6) e <c>name</c>, mais os nomes
/// de glifo do <c>post</c> formato 2. So' o que o teste precisa; nao e' um parser geral.</summary>
internal sealed class TtfMinimo
{
    private readonly byte[] _d;
    private readonly Dictionary<string, (int Offset, int Length)> _tabelas = new();
    public int NumGlyphs { get; private set; }
    public int UnitsPerEm { get; private set; }
    public List<(int Id, int Plataforma, string Texto)> Nomes { get; } = [];
    public List<string> NomesDosGlifos { get; } = [];

    private TtfMinimo(byte[] d) { _d = d; }

    public static TtfMinimo Ler(byte[] d)
    {
        var t = new TtfMinimo(d);
        Assert.Equal(0x00010000u, U32(d, 0)); // TrueType
        int n = U16(d, 4);
        for (var i = 0; i < n; i++)
        {
            var rec = 12 + i * 16;
            var tag = Encoding.ASCII.GetString(d, rec, 4);
            t._tabelas[tag] = ((int)U32(d, rec + 8), (int)U32(d, rec + 12));
        }
        t.NumGlyphs = U16(d, t.Tabela("maxp") + 4);
        t.UnitsPerEm = U16(d, t.Tabela("head") + 18);
        t.LerNomes();
        t.LerNomesDosGlifos();
        return t;
    }

    private int Tabela(string tag) => _tabelas.TryGetValue(tag, out var v) ? v.Offset : throw new InvalidDataException("tabela ausente: " + tag);

    private void LerNomes()
    {
        var b = Tabela("name");
        int count = U16(_d, b + 2), strings = b + U16(_d, b + 4);
        for (var i = 0; i < count; i++)
        {
            var r = b + 6 + i * 12;
            int plat = U16(_d, r), id = U16(_d, r + 6), len = U16(_d, r + 8), off = U16(_d, r + 10);
            var texto = plat == 3 || plat == 0 ? Encoding.BigEndianUnicode.GetString(_d, strings + off, len) : Encoding.Latin1.GetString(_d, strings + off, len);
            Nomes.Add((id, plat, texto));
        }
    }

    private void LerNomesDosGlifos()
    {
        var b = Tabela("post");
        if (U32(_d, b) != 0x00020000) return; // so' formato 2.0 traz nomes
        int n = U16(_d, b + 32);
        var indices = new int[n];
        for (var i = 0; i < n; i++) indices[i] = U16(_d, b + 34 + i * 2);
        var pascal = new List<string>();
        var p = b + 34 + n * 2;
        var fim = b + _tabelas["post"].Length;
        while (p < fim) { int len = _d[p]; pascal.Add(Encoding.Latin1.GetString(_d, p + 1, len)); p += 1 + len; }
        foreach (var idx in indices) NomesDosGlifos.Add(idx >= 258 ? pascal[idx - 258] : NomesMac[idx]);
    }

    // So' os nomes Macintosh padrao que esta fonte usa (indices 0..2); qualquer outro seria nome customizado.
    private static readonly string[] NomesMac = [".notdef", ".null", "nonmarkingreturn"];

    public Dictionary<int, int>? Cmap(int plataforma, int codificacao)
    {
        var b = Tabela("cmap");
        int n = U16(_d, b + 2);
        for (var i = 0; i < n; i++)
        {
            var r = b + 4 + i * 8;
            if (U16(_d, r) != plataforma || U16(_d, r + 2) != codificacao) continue;
            var sub = b + (int)U32(_d, r + 4);
            return U16(_d, sub) switch
            {
                4 => LerFormato4(sub),
                6 => LerFormato6(sub),
                var f => throw new InvalidDataException("formato de cmap nao suportado pelo teste: " + f),
            };
        }
        return null;
    }

    private Dictionary<int, int> LerFormato4(int sub)
    {
        var m = new Dictionary<int, int>();
        int segX2 = U16(_d, sub + 6), seg = segX2 / 2;
        int endBase = sub + 14, startBase = endBase + segX2 + 2, deltaBase = startBase + segX2, rangeBase = deltaBase + segX2;
        for (var s = 0; s < seg; s++)
        {
            int end = U16(_d, endBase + s * 2), start = U16(_d, startBase + s * 2);
            int delta = U16(_d, deltaBase + s * 2), range = U16(_d, rangeBase + s * 2);
            for (var c = start; c <= end && c != 0xFFFF; c++)
            {
                int g;
                if (range == 0) g = (c + delta) & 0xFFFF;
                else
                {
                    var addr = rangeBase + s * 2 + range + (c - start) * 2;
                    g = U16(_d, addr);
                    if (g != 0) g = (g + delta) & 0xFFFF;
                }
                if (g != 0) m[c] = g;
            }
        }
        return m;
    }

    private Dictionary<int, int> LerFormato6(int sub)
    {
        var m = new Dictionary<int, int>();
        int first = U16(_d, sub + 6), count = U16(_d, sub + 8);
        for (var i = 0; i < count; i++) { var g = U16(_d, sub + 10 + i * 2); if (g != 0) m[first + i] = g; }
        return m;
    }

    public int ContornosDoGlifo(int g)
    {
        var loca = Tabela("loca");
        var longo = (short)U16(_d, Tabela("head") + 50) == 1;
        int ini = longo ? (int)U32(_d, loca + g * 4) : U16(_d, loca + g * 2) * 2;
        int fim = longo ? (int)U32(_d, loca + (g + 1) * 4) : U16(_d, loca + (g + 1) * 2) * 2;
        if (fim <= ini) return 0;
        return (short)U16(_d, Tabela("glyf") + ini);
    }

    public int AvancoDoGlifo(int g)
    {
        int numH = U16(_d, Tabela("hhea") + 34);
        return U16(_d, Tabela("hmtx") + Math.Min(g, numH - 1) * 4);
    }

    private static ushort U16(byte[] d, int o) => BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o, 2));
    private static uint U32(byte[] d, int o) => BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o, 4));
}
