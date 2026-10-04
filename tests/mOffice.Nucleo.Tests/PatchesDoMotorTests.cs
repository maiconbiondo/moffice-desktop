using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace mOffice.Nucleo.Tests;

/// <summary>Guarda dos patches de texto do motor (Plano 6, Task 6). Implementacao INDEPENDENTE do aplicar-patches.ps1 (como
/// FonteSimbolosTests faz com o patch da fonte): duas implementacoes do mesmo criterio.</summary>
public class PatchesDoMotorTests
{
    private static string Raiz => FronteiraTests.RaizRepo();
    private static string PastaPatches => Path.Combine(Raiz, "tools", "onlyoffice", "patches");
    private static string Script => Path.Combine(Raiz, "tools", "onlyoffice", "aplicar-patches.ps1");
    private const string Versao = "9.4.0.129";

    public static IReadOnlyList<(string Arquivo, string De, string Para, int Ocorrencias)> Trocas()
    {
        var lista = new List<(string, string, string, int)>();
        foreach (var json in Directory.EnumerateFiles(PastaPatches, "*.json").Order(StringComparer.Ordinal))
        {
            using var d = JsonDocument.Parse(File.ReadAllText(json));
            foreach (var a in d.RootElement.GetProperty("arquivos").EnumerateArray())
                foreach (var t in a.GetProperty("trocas").EnumerateArray())
                    lista.Add((a.GetProperty("arquivo").GetString()!, t.GetProperty("de").GetString()!, t.GetProperty("para").GetString()!, t.GetProperty("ocorrencias").GetInt32()));
        }
        return lista;
    }

    private static int Contar(string texto, string trecho) { int n = 0, i = 0; while ((i = texto.IndexOf(trecho, i, StringComparison.Ordinal)) >= 0) { n++; i += trecho.Length; } return n; }
    private static string Sha(string arquivo) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(arquivo)));
    private static string Caminho(string raiz, string rel) => Path.Combine(raiz, rel.Replace('/', Path.DirectorySeparatorChar));
    private static string LerEstrito(string arquivo)
    {
        var b = File.ReadAllBytes(arquivo);
        var off = b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF ? 3 : 0;
        return new UTF8Encoding(false, true).GetString(b, off, b.Length - off);
    }

    [Fact] // todo patch declara a versao do motor que esta em VERSAO.txt e tem ao menos uma troca
    public void Patches_BemFormados()
    {
        var arquivos = Directory.EnumerateFiles(PastaPatches, "*.json").ToList();
        Assert.NotEmpty(arquivos);
        foreach (var f in arquivos)
        {
            using var d = JsonDocument.Parse(File.ReadAllText(f));
            Assert.Equal(Versao, d.RootElement.GetProperty("versaoMotor").GetString());
        }
        Assert.NotEmpty(Trocas());
        // Cada arquivo do motor pertence a UM patch so': o aplicador planeja cada patch a partir do texto do disco.
        var porPatch = arquivos.SelectMany(f => { using var d = JsonDocument.Parse(File.ReadAllText(f));
            return d.RootElement.GetProperty("arquivos").EnumerateArray().Select(a => a.GetProperty("arquivo").GetString()!).ToList(); }).ToList();
        Assert.Equal(porPatch.Count, porPatch.Select(p => p.ToLowerInvariant()).Distinct().Count());
        // sdk-all.js e' do patch da fonte (mesmo .orig): patch de texto ali sobrescreveria o original guardado.
        Assert.DoesNotContain(porPatch, p => System.Text.RegularExpressions.Regex.IsMatch(p, @"sdkjs/(word|cell|slide|visio)/sdk-all\.js$"));
        // Cada troca precisa ser reconhecivel depois de aplicada: "de" nao pode ser subtexto de "para" (senao nunca fica "ja aplicado").
        foreach (var (arq, de, para, n) in Trocas())
        {
            Assert.False(string.IsNullOrEmpty(de), arq);
            Assert.True(n >= 1, arq);
            Assert.True(Contar(para, de) == 0, $"{arq}: o 'de' aparece dentro do 'para' - o patch nao seria idempotente");
        }
    }

    [SkippableFact] // o motor LOCAL (third_party) esta' com todos os patches aplicados - senao os testes de janela mentem
    public void MotorLocal_ComTodosOsPatches()
    {
        var motor = Path.Combine(Raiz, "third_party", "onlyoffice");
        Skip.If(!File.Exists(Path.Combine(motor, "VERSAO.txt")), "motor nao obtido (rode tools/onlyoffice/obter-motor.ps1)");
        Assert.Equal(Versao, File.ReadAllText(Path.Combine(motor, "VERSAO.txt")).TrimStart((char)0xFEFF).Split((char)10)[0].Trim()); // os patches foram medidos nesta versao
        foreach (var (arquivo, de, para, n) in Trocas())
        {
            var texto = LerEstrito(Caminho(motor, arquivo));
            Assert.True(Contar(texto, de) == 0 && Contar(texto, para) == n,
                $"{arquivo}: patch nao aplicado (rode tools/onlyoffice/aplicar-patches.ps1). de={Contar(texto, de)} para={Contar(texto, para)}");
        }
    }

    // ---- o aplicador de verdade (powershell.exe) ----

    private static (int Codigo, string Saida) Rodar(string pastaMotor, string pastaPatches, params string[] extras)
    {
        var psi = new ProcessStartInfo("powershell.exe") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (var a in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", Script, "-PastaMotor", pastaMotor, "-PastaPatches", pastaPatches, "-SoPatchesDeTexto" }.Concat(extras)) psi.ArgumentList.Add(a);
        using var p = Process.Start(psi)!;
        var so = p.StandardOutput.ReadToEndAsync(); var se = p.StandardError.ReadToEndAsync();
        Assert.True(p.WaitForExit(120_000), "aplicar-patches.ps1 travou");
        return (p.ExitCode, so.Result + se.Result);
    }

    private static string NovaPasta() => Directory.CreateTempSubdirectory("moffice-patches-").FullName;
    private static void VersaoTxt(string motor) { Directory.CreateDirectory(motor); File.WriteAllText(Path.Combine(motor, "VERSAO.txt"), Versao + "\r\n"); }
    private static void Escrever(string raiz, string rel, byte[] bytes) { var c = Caminho(raiz, rel); Directory.CreateDirectory(Path.GetDirectoryName(c)!); File.WriteAllBytes(c, bytes); }
    private static byte[] Utf8(string s) => new UTF8Encoding(false).GetBytes(s);
    private static byte[] ComBom(string s) => [0xEF, 0xBB, 0xBF, .. Utf8(s)];
    private static string Json(string arquivo, string de, string para, int n, string versao = Versao) =>
        JsonSerializer.Serialize(new { descricao = "teste", versaoMotor = versao, arquivos = new[] { new { arquivo, trocas = new[] { new { de, para, ocorrencias = n } } } } });

    [SkippableFact] // os patches REAIS num motor-miniatura: aplica, e' idempotente, reverte byte a byte
    public void Aplicador_PatchesReais_AplicaIdempotenteEReverte()
    {
        var motorReal = Path.Combine(Raiz, "third_party", "onlyoffice");
        Skip.If(!File.Exists(Path.Combine(motorReal, "VERSAO.txt")), "motor nao obtido");
        var tmp = NovaPasta();
        try
        {
            var motor = Path.Combine(tmp, "motor"); VersaoTxt(motor);
            var originais = new Dictionary<string, string>(); // arquivo -> sha do ORIGINAL (o motor local desfeito pela troca inversa)
            foreach (var g in Trocas().GroupBy(t => t.Arquivo))
            {
                var origem = Caminho(motorReal, g.Key);
                var bytes = File.ReadAllBytes(origem);
                var bom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
                var texto = LerEstrito(origem);
                foreach (var t in g) texto = texto.Replace(t.Para, t.De); // volta ao original (o motor local pode ja' estar patchado)
                Escrever(motor, g.Key, bom ? ComBom(texto) : Utf8(texto));
                originais[g.Key] = Sha(Caminho(motor, g.Key));
            }
            var (c1, s1) = Rodar(motor, PastaPatches);
            Assert.True(c1 == 0, s1);
            foreach (var (arq, de, para, n) in Trocas())
            {
                var t = LerEstrito(Caminho(motor, arq));
                Assert.True(Contar(t, de) == 0 && Contar(t, para) == n, arq);
                Assert.True(File.Exists(Caminho(motor, arq) + ".orig"), arq + ".orig");
            }
            var aposAplicar = originais.Keys.ToDictionary(k => k, k => Sha(Caminho(motor, k)));
            var (c2, s2) = Rodar(motor, PastaPatches);
            Assert.True(c2 == 0, s2);
            Assert.Contains("ja aplicado", s2);
            Assert.DoesNotContain("aplicado     ", s2);
            foreach (var k in aposAplicar.Keys) Assert.Equal(aposAplicar[k], Sha(Caminho(motor, k)));
            var (c3, s3) = Rodar(motor, PastaPatches, "-Reverter");
            Assert.True(c3 == 0, s3);
            foreach (var k in originais.Keys)
            {
                Assert.Equal(originais[k], Sha(Caminho(motor, k)));
                Assert.False(File.Exists(Caminho(motor, k) + ".orig"));
            }
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }

    [Fact] // atomicidade: um trecho ausente em QUALQUER arquivo para tudo - o arquivo bom do patch anterior fica intocado, sem .orig
    public void Aplicador_TrechoAusente_NaoGravaNada()
    {
        var tmp = NovaPasta();
        try
        {
            var motor = Path.Combine(tmp, "motor"); var patches = Path.Combine(tmp, "patches"); VersaoTxt(motor); Directory.CreateDirectory(patches);
            Escrever(motor, "a/bom.txt", ComBom("antes AAA depois"));
            Escrever(motor, "a/ruim.txt", Utf8("nao tem o trecho"));
            File.WriteAllText(Path.Combine(patches, "01-bom.json"), Json("a/bom.txt", "AAA", "ZZZ", 1));
            File.WriteAllText(Path.Combine(patches, "02-ruim.json"), Json("a/ruim.txt", "XXX", "YYY", 1));
            var bom = Sha(Path.Combine(motor, "a", "bom.txt")); var ruim = Sha(Path.Combine(motor, "a", "ruim.txt"));
            var (c, s) = Rodar(motor, patches);
            Assert.NotEqual(0, c);
            Assert.Contains("02-ruim.json", s);
            Assert.Equal(bom, Sha(Path.Combine(motor, "a", "bom.txt"))); Assert.Equal(ruim, Sha(Path.Combine(motor, "a", "ruim.txt")));
            Assert.False(File.Exists(Path.Combine(motor, "a", "bom.txt.orig")));
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }

    [Fact] // BOM e acentos sobrevivem; o resultado e' o esperado; .orig guarda o original; reverter devolve o sha
    public void Aplicador_PreservaBomEAcentos_ERevertePeloOrig()
    {
        var tmp = NovaPasta();
        try
        {
            var motor = Path.Combine(tmp, "motor"); var patches = Path.Combine(tmp, "patches"); VersaoTxt(motor); Directory.CreateDirectory(patches);
            Escrever(motor, "a/pt.json", ComBom("{\"x\": \"Não\", \"y\": \"Não\"}\r\nfim"));
            File.WriteAllText(Path.Combine(patches, "01-acento.json"), Json("a/pt.json", "\"x\": \"Não\"", "\"x\": \"Nunca ação\"", 1), new UTF8Encoding(false));
            var alvo = Path.Combine(motor, "a", "pt.json"); var sha0 = Sha(alvo);
            var (c, s) = Rodar(motor, patches); Assert.True(c == 0, s);
            var b = File.ReadAllBytes(alvo);
            Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, b[..3]);
            Assert.Equal("{\"x\": \"Nunca ação\", \"y\": \"Não\"}\r\nfim", LerEstrito(alvo));
            Assert.Equal(sha0, Sha(alvo + ".orig"));
            var (c2, s2) = Rodar(motor, patches, "-Reverter"); Assert.True(c2 == 0, s2);
            Assert.Equal(sha0, Sha(alvo)); Assert.False(File.Exists(alvo + ".orig"));
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }

    [Theory] // contagem errada, versao errada, ambiguo, estado misto, "de" vazio, "de" dentro de "para": todos param sem gravar
    [InlineData("AAA", "ZZZ", 2, Versao, "antes AAA depois")]
    [InlineData("AAA", "ZZZ", 1, "9.9.9.9", "antes AAA depois")]
    [InlineData("AAA", "ZZZ", 1, Versao, "AAA e AAA")]
    [InlineData("AAA", "ZZZ", 1, Versao, "AAA e ZZZ")]
    [InlineData("", "ZZZ", 1, Versao, "AAA")]
    [InlineData("AAA", "xAAAx", 1, Versao, "AAA")]
    public void Aplicador_RecusaPatchInvalido_SemGravar(string de, string para, int n, string versao, string conteudo)
    {
        var tmp = NovaPasta();
        try
        {
            var motor = Path.Combine(tmp, "motor"); var patches = Path.Combine(tmp, "patches"); VersaoTxt(motor); Directory.CreateDirectory(patches);
            Escrever(motor, "a/f.txt", Utf8(conteudo));
            File.WriteAllText(Path.Combine(patches, "01-x.json"), Json("a/f.txt", de, para, n, versao));
            var antes = Sha(Path.Combine(motor, "a", "f.txt"));
            var (c, _) = Rodar(motor, patches);
            Assert.NotEqual(0, c);
            Assert.Equal(antes, Sha(Path.Combine(motor, "a", "f.txt")));
            Assert.False(File.Exists(Path.Combine(motor, "a", "f.txt.orig")));
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }

    [Fact] // dois patches no mesmo arquivo, ou patch de texto num sdk-all.js (e' do patch da fonte): recusado
    public void Aplicador_RecusaArquivoDeDoisPatches_ESdkAll()
    {
        var tmp = NovaPasta();
        try
        {
            var motor = Path.Combine(tmp, "motor"); var patches = Path.Combine(tmp, "patches"); VersaoTxt(motor); Directory.CreateDirectory(patches);
            Escrever(motor, "a/f.txt", Utf8("AAA BBB"));
            File.WriteAllText(Path.Combine(patches, "01-x.json"), Json("a/f.txt", "AAA", "ZZZ", 1));
            File.WriteAllText(Path.Combine(patches, "02-y.json"), Json("a/f.txt", "BBB", "WWW", 1));
            var antes = Sha(Path.Combine(motor, "a", "f.txt"));
            var (c, s) = Rodar(motor, patches);
            Assert.NotEqual(0, c); Assert.Contains("02-y.json", s);
            Assert.Equal(antes, Sha(Path.Combine(motor, "a", "f.txt")));

            File.Delete(Path.Combine(patches, "02-y.json")); File.Delete(Path.Combine(patches, "01-x.json"));
            Escrever(motor, "editors/sdkjs/word/sdk-all.js", Utf8("AAA"));
            File.WriteAllText(Path.Combine(patches, "03-sdk.json"), Json("editors/sdkjs/word/sdk-all.js", "AAA", "ZZZ", 1));
            var (c2, s2) = Rodar(motor, patches);
            Assert.NotEqual(0, c2); Assert.Contains("sdk-all", s2);
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }

    [Fact] // reverter com o arquivo patchado e SEM .orig nao inventa nada: recusa
    public void Reverter_SemOrigEPatchado_Recusa()
    {
        var tmp = NovaPasta();
        try
        {
            var motor = Path.Combine(tmp, "motor"); var patches = Path.Combine(tmp, "patches"); VersaoTxt(motor); Directory.CreateDirectory(patches);
            Escrever(motor, "a/f.txt", Utf8("ZZZ"));
            File.WriteAllText(Path.Combine(patches, "01-x.json"), Json("a/f.txt", "AAA", "ZZZ", 1));
            var (c, _) = Rodar(motor, patches, "-Reverter");
            Assert.NotEqual(0, c);
            Assert.Equal("ZZZ", File.ReadAllText(Path.Combine(motor, "a", "f.txt")));
        }
        finally { try { Directory.Delete(tmp, true); } catch { } }
    }
}
