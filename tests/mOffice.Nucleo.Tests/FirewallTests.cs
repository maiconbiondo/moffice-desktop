using System.Diagnostics;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 3 (S2), fix da revisao 1, achado 3: a regra de firewall (tools/egresso/firewall.ps1) aponta para um
/// caminho VERSIONADO — third_party/webview2/&lt;versao&gt;/.../&lt;exe&gt;. Se o runtime for atualizado (nova
/// versao, pasta nova) e ninguem rodar o script de novo, a(s) regra(s) continuam apontando para a pasta
/// ANTIGA e o egresso zero (S2) para de valer sem que nada avise disso. Este teste NAO precisa de
/// administrador (<c>netsh advfirewall firewall show rule</c> so' LE, nao precisa de elevacao) e roda em
/// qualquer maquina: sem regra nenhuma (dev que nunca rodou tools/egresso/firewall.ps1), ele pula.
///
/// Reprocedimento ao trocar de versao do runtime: rodar `tools/egresso/firewall.ps1` de novo (idempotente:
/// apaga as regras antigas pelo prefixo do nome e recria apontando para o caminho novo) depois de
/// `tools/webview2/obter-runtime.ps1`. Ate isso acontecer, este teste fica VERMELHO de proposito — e' o
/// guarda contra uma regra pendurada num caminho que pode nem existir mais.
/// </summary>
public class FirewallTests
{
    private const string Prefixo = "mOffice WebView2 - bloqueio de saida";

    [SkippableFact]
    public void RegraApontaParaORuntimeAtual()
    {
        var regras = LerRegrasMOffice();
        Skip.If(regras.Count == 0, "nenhuma regra de firewall do mOffice presente — rode tools/egresso/firewall.ps1");

        var raiz = AmbienteWebView2.PastaRuntimeFixo;
        // Regras de OUTRA arvore (ex.: o repositorio de desenvolvimento, nesta mesma maquina) e nenhum runtime NESTA:
        // nao ha' o que conferir aqui. Regra que aponta para dentro desta arvore sem runtime continua sendo orfa (falha).
        var raizArvore = FronteiraTests.RaizRepo();
        Skip.If(string.IsNullOrEmpty(raiz) && regras.All(r => r.Programa is not null &&
                    !r.Programa.StartsWith(raizArvore + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)),
                "runtime fixo ausente desta arvore e as regras de firewall pertencem a outra copia do repositorio");
        Assert.False(string.IsNullOrEmpty(raiz),
            "ha' regra(s) de firewall do mOffice mas nenhum runtime fixo foi encontrado (PastaRuntimeFixo nulo) — " +
            "rode tools/webview2/obter-runtime.ps1, ou as regras estao orfas (tools/egresso/firewall.ps1 -Remover)");

        foreach (var (nome, programa) in regras)
        {
            Assert.False(string.IsNullOrEmpty(programa), $"regra '{nome}' sem caminho de Program (formato inesperado do netsh)");
            Assert.True(programa!.StartsWith(raiz!, StringComparison.OrdinalIgnoreCase),
                $"regra '{nome}' aponta para '{programa}', fora do runtime fixo ATUAL ('{raiz}') — " +
                "rode tools/egresso/firewall.ps1 de novo depois de trocar a versao do runtime");
        }
    }

    /// <summary>Le as regras (nome + caminho do Program) cujo nome comeca com <see cref="Prefixo"/>, parseando
    /// a saida do netsh — que so' precisa de leitura, sem elevacao. Aceita os dois idiomas (a maquina desta
    /// task e' pt-BR; a bancada de CI pode ser en-US) porque o rotulo do netsh muda com o idioma do Windows,
    /// nao com o parametro da chamada.</summary>
    private static List<(string Nome, string? Programa)> LerRegrasMOffice()
    {
        var psi = new ProcessStartInfo("netsh", "advfirewall firewall show rule name=all verbose")
        {
            RedirectStandardOutput = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        using var p = Process.Start(psi) ?? throw new InvalidOperationException("nao foi possivel iniciar o netsh");
        var saida = p.StandardOutput.ReadToEnd();
        p.WaitForExit();

        var regras = new List<(string, string?)>();
        string? nomeAtual = null;
        string? programaAtual = null;
        void FechaRegraAtual()
        {
            if (nomeAtual is not null && nomeAtual.StartsWith(Prefixo, StringComparison.Ordinal))
                regras.Add((nomeAtual, programaAtual));
        }
        foreach (var linhaCrua in saida.Split('\n'))
        {
            var linha = linhaCrua.TrimEnd('\r');
            if (linha.StartsWith("Rule Name:", StringComparison.Ordinal) || linha.StartsWith("Nome da Regra:", StringComparison.Ordinal))
            {
                FechaRegraAtual();
                nomeAtual = ExtrairValor(linha);
                programaAtual = null;
            }
            else if (linha.StartsWith("Program:", StringComparison.Ordinal) || linha.StartsWith("Programa:", StringComparison.Ordinal))
            {
                programaAtual = ExtrairValor(linha);
            }
        }
        FechaRegraAtual();
        return regras;
    }

    private static string ExtrairValor(string linha)
    {
        var i = linha.IndexOf(':');
        return (i < 0 ? linha : linha[(i + 1)..]).Trim();
    }
}
