using System.Text.RegularExpressions;

namespace mOffice.Nucleo.Tests;

/// <summary>Guarda contra o erro de 2026-09-25: o sdkjs chamava <c>AscDesktopEditor.CheckNeedWheel()</c> a
/// cada giro da roda do mouse, a ponte nao tinha a funcao, o TypeError travava o documento com "Ocorreu um
/// erro. Use a opcao 'Baixar como'". Esta guarda varre os bundles do motor embarcado e exige que TODA
/// chamada <c>AscDesktopEditor.X(</c> / <c>window.desktop.X(</c> tenha X implementado em ponte-desktop.js —
/// menos as que o proprio motor testa antes de chamar (<see cref="SoComTesteDeExistencia"/>). Uma versao
/// nova do motor que passe a chamar algo novo falha AQUI, nao na mao do usuario.</summary>
public class PonteCompletaTests
{
    /// <summary>Chamadas que o motor so' faz depois de conferir que a funcao existe (<c>AscDesktopEditor.X &amp;&amp; …X()</c>,
    /// <c>AscDesktopEditor.X ? …X() : padrao</c> ou <c>if (… AscDesktopEditor.X) { …X() }</c>): ausentes, o motor segue o
    /// caminho padrao — e IMPLEMENTAR mudaria o comportamento (ex. isSupportMacroses devolvendo false desligaria as macros).
    /// Conferido nos bundles 9.4.0.129 em 2026-09-25; os 7 ultimos, nos sdk-all-min.js, na revisao final do Plano 4
    /// (cada chamada provada por <see cref="ListaSoComTeste_SoTemNomesQueOMotorTestaAntes"/>).</summary>
    private static readonly HashSet<string> SoComTesteDeExistencia =
        ["IsFilePrinting", "SetLocalRestrictions", "getEngineVersion", "isSupportPlugins",
         "CallMediaPlayerCommand", "GetImageOriginalSize", "getDictionariesPath", "isSupportMacroses", "onFileLockedClose",
         "SetPdfCloudPrintFileInfo", "IsSignaturesSupport"];

    private static readonly Regex Chamada = new(@"(?:AscDesktopEditor|window\.desktop)\.([A-Za-z_]+)\s*\(", RegexOptions.Compiled);
    private static readonly Regex Implementada = new(@"^\s+([A-Za-z_]+)\s*:", RegexOptions.Compiled | RegexOptions.Multiline);

    private static IEnumerable<string> Bundles(string motor)
    {
        foreach (var e in new[] { "word", "cell", "slide", "visio" })
        {
            yield return Path.Combine(motor, "editors", "sdkjs", e, "sdk-all.js");
            // Revisao final do Plano 4: o sdk-all-min.js e' o bundle que o editor CARREGA primeiro (o sdk-all.js so' vem
            // depois, e so' no desktop) e tinha ~30 chamadas fora da varredura (OpenWorkbook, openExternalReference, ...).
            yield return Path.Combine(motor, "editors", "sdkjs", e, "sdk-all-min.js");
        }
        foreach (var f in Directory.EnumerateFiles(Path.Combine(motor, "editors", "sdkjs", "common"), "*.js"))
            yield return f;
        foreach (var e in new[] { "documenteditor", "spreadsheeteditor", "presentationeditor" })
            yield return Path.Combine(motor, "editors", "web-apps", "apps", e, "main", "app.js");
        yield return Path.Combine(motor, "editors", "web-apps", "apps", "api", "documents", "api.js");
    }

    [SkippableFact] // toda funcao da ponte que o motor chama sem testar existe em ponte-desktop.js
    public void TodaChamadaDoMotor_TemImplementacaoNaPonte()
    {
        var raiz = FronteiraTests.RaizRepo();
        var motor = Path.Combine(raiz, "third_party", "onlyoffice");
        Skip.If(!Directory.Exists(Path.Combine(motor, "editors", "sdkjs")), "motor ONLYOFFICE nao obtido (rode tools/onlyoffice/obter-motor.ps1)");

        var chamadas = new SortedSet<string>(StringComparer.Ordinal);
        var lidos = 0;
        foreach (var arquivo in Bundles(motor).Where(File.Exists))
        {
            lidos++;
            foreach (Match m in Chamada.Matches(File.ReadAllText(arquivo))) chamadas.Add(m.Groups[1].Value);
        }
        Assert.True(lidos >= 12, $"varredura vazia: so' {lidos} bundles lidos"); // piso: a guarda nao pode passar por nao ter olhado nada
        Assert.Contains("CheckNeedWheel", chamadas); // controle positivo: o caso real de 2026-09-25 e' enxergado
        Assert.Contains("OpenWorkbook", chamadas);   // controle positivo 2: so' existe no cell/sdk-all-min.js (revisao final do Plano 4)

        var ponte = File.ReadAllText(Path.Combine(raiz, "src", "mOffice.Nucleo", "Ponte", "ponte-desktop.js"));
        var implementadas = Implementada.Matches(ponte).Select(m => m.Groups[1].Value).ToHashSet(StringComparer.Ordinal);

        var faltando = chamadas.Where(c => !implementadas.Contains(c) && !SoComTesteDeExistencia.Contains(c)).ToList();
        Assert.True(faltando.Count == 0,
            "O motor chama funcoes da ponte que ponte-desktop.js nao implementa (cada uma trava o editor com " +
            "\"Ocorreu um erro\" quando alcancada): " + string.Join(", ", faltando));
    }

    /// <summary>Toda chamada <c>AscDesktopEditor.X(</c> tem, nos 400 caracteres anteriores DO MESMO ARQUIVO, um teste de
    /// existencia de X — <c>AscDesktopEditor.X&amp;&amp;</c>, <c>AscDesktopEditor.X?</c> ou <c>AscDesktopEditor.X){</c> — SEM
    /// <c>;</c> entre o teste e a chamada (um <c>;</c> encerra a expressao: a chamada seguinte ja' nao esta' protegida).
    /// Arquivo a arquivo (re-revisao do Plano 4): na concatenacao, o fim de um bundle podia "proteger" o comeco do
    /// seguinte. Devolve o trecho da primeira chamada SEM teste, ou null.</summary>
    internal static string? ChamadaSemTeste(IEnumerable<string> textos, string nome)
    {
        var teste = new Regex(@"AscDesktopEditor\." + nome + @"\s*(?:&&|\?|\)\s*\{)");
        var chamada = new Regex(@"AscDesktopEditor\." + nome + @"\s*\(");
        foreach (var texto in textos)
            foreach (Match m in chamada.Matches(texto))
            {
                var ini = Math.Max(0, m.Index - 400);
                var janela = texto[ini..m.Index];
                var testes = teste.Matches(janela);
                var protegida = testes.Count > 0 && !janela[(testes[^1].Index + testes[^1].Length)..].Contains(';');
                if (!protegida) return texto[Math.Max(0, m.Index - 120)..Math.Min(texto.Length, m.Index + 60)];
            }
        return null;
    }

    [Fact] // a regra do "so' com teste": controles sinteticos
    public void ChamadaSemTeste_Controles()
    {
        Assert.Null(ChamadaSemTeste(["a.AscDesktopEditor.Foo&&a.AscDesktopEditor.Foo(1)"], "Foo"));
        Assert.Null(ChamadaSemTeste(["x=a.AscDesktopEditor.Foo?a.AscDesktopEditor.Foo():!1"], "Foo"));
        Assert.Null(ChamadaSemTeste(["if(a.AscDesktopEditor&&a.AscDesktopEditor.Foo){var r=a.AscDesktopEditor.Foo()}"], "Foo"));
        Assert.Null(ChamadaSemTeste(["a.AscDesktopEditor.Foo&&(g(function(){a.AscDesktopEditor.Foo(!0)}),g(function(){a.AscDesktopEditor.Foo(!1)}))"], "Foo"));
        // Protegida seguida de uma NAO protegida depois de ";": a segunda e' apontada.
        Assert.NotNull(ChamadaSemTeste(["a.AscDesktopEditor.Foo&&a.AscDesktopEditor.Foo();a.AscDesktopEditor.Foo()"], "Foo"));
        Assert.NotNull(ChamadaSemTeste(["if(a.AscDesktopEditor.Foo){a.AscDesktopEditor.Foo()};b();a.AscDesktopEditor.Foo()"], "Foo"));
        // Sem teste nenhum.
        Assert.NotNull(ChamadaSemTeste(["a.AscDesktopEditor.Foo()"], "Foo"));
        // Arquivo a arquivo: o teste no FIM de um arquivo nao protege a chamada no COMECO do seguinte.
        Assert.NotNull(ChamadaSemTeste(["...a.AscDesktopEditor.Foo&&", "a.AscDesktopEditor.Foo()"], "Foo"));
    }

    [SkippableFact] // a lista de "so' com teste" nao pode esconder uma chamada que o motor faz SEM testar
    public void ListaSoComTeste_SoTemNomesQueOMotorTestaAntes()
    {
        var motor = Path.Combine(FronteiraTests.RaizRepo(), "third_party", "onlyoffice");
        Skip.If(!Directory.Exists(Path.Combine(motor, "editors", "sdkjs")), "motor ONLYOFFICE nao obtido (rode tools/onlyoffice/obter-motor.ps1)");
        var textos = Bundles(motor).Where(File.Exists).Select(File.ReadAllText).ToList();
        // Controles negativos: chamadas SEM teste sao reconhecidas como tal (a regra nao aceita tudo).
        Assert.NotNull(ChamadaSemTeste(textos, "OpenWorkbook"));
        Assert.NotNull(ChamadaSemTeste(textos, "emulateCloudPrinting"));   // parte das chamadas testa, parte nao
        foreach (var nome in SoComTesteDeExistencia)
        {
            Assert.True(textos.Any(t => Regex.IsMatch(t, @"AscDesktopEditor\." + nome + @"\s*\(")), $"{nome} nao e' mais chamado — tire da lista");
            var sem = ChamadaSemTeste(textos, nome);
            Assert.True(sem is null, $"{nome}: chamada SEM teste de existencia — implemente na ponte. Trecho: {sem}");
        }
    }
}
