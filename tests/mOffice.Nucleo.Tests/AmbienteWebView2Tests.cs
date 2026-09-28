using Microsoft.Web.WebView2.Core;
using mOffice.Nucleo;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 3 (S2 — egresso zero). Cobre o contrato do ambiente compartilhado (um por thread de interface, o que
/// no produto e' um por PROCESSO): os argumentos de linha de comando do navegador e os nomes de feature
/// conferidos no binario, a pasta de dados por app e — quando o runtime fixo esta' baixado — o fato de o
/// ambiente rodar sobre ELE e ser reutilizado por todas as janelas da thread.
/// </summary>
public class AmbienteWebView2Tests
{
    [Fact] // os argumentos do navegador contem as flags de rede de fundo (spec 7.2)
    public void Argumentos_ContemFlagsDeRede()
    {
        var a = AmbienteWebView2.ArgumentosNavegador;
        Assert.Contains("--disable-background-networking", a);
        Assert.Contains("--disable-component-update", a);
        Assert.Contains("--disable-features=", a);
    }

    [Fact] // cada flag da linha de comando tem um motivo escrito: uma flag sem justificativa e' culto de carga
    public void Flags_TodasTemMotivo()
    {
        Assert.NotEmpty(AmbienteWebView2.Flags);
        foreach (var (flag, motivo) in AmbienteWebView2.Flags)
        {
            Assert.StartsWith("--", flag);
            Assert.False(string.IsNullOrWhiteSpace(motivo), $"flag sem motivo: {flag}");
        }
        // Nenhuma flag repetida (duas --disable-features= na mesma linha: a segunda vence em silencio).
        Assert.Equal(AmbienteWebView2.Flags.Count, AmbienteWebView2.Flags.Select(f => f.Flag.Split('=')[0]).Distinct().Count());
    }

    [SkippableFact] // o que DE VERDADE derruba uma flag: um nome de --disable-features que nao existe no build
    public void Features_ExistemNoBinarioDoRuntimePinado()
    {
        var pasta = AmbienteWebView2.PastaRuntimeFixo;
        Skip.If(pasta is null, "runtime fixo ausente — rode tools/webview2/obter-runtime.ps1");
        var dll = Path.Combine(pasta!, "msedge.dll");
        Skip.If(!File.Exists(dll), "msedge.dll nao encontrado no runtime fixo");

        var nomes = AmbienteWebView2.Flags
            .Where(f => f.Flag.StartsWith("--disable-features="))
            .SelectMany(f => f.Flag["--disable-features=".Length..].Split(',', StringSplitOptions.RemoveEmptyEntries))
            .ToArray();
        Assert.NotEmpty(nomes);

        var achados = ProcurarNoBinario(dll, nomes);
        // Controle negativo: um nome inventado NAO pode ser "encontrado" — senao a varredura acha qualquer coisa.
        Assert.DoesNotContain("msFeatureQueNaoExiste2026", ProcurarNoBinario(dll, ["msFeatureQueNaoExiste2026"]));
        foreach (var n in nomes)
            Assert.True(achados.Contains(n), $"--disable-features={n} nao existe no runtime {AmbienteWebView2.VersaoRuntimeFixo}: flag ignorada em silencio");
    }

    /// <summary>Varre o binario (centenas de MB) em blocos, procurando cada nome em ASCII. Os blocos se
    /// sobrepoem pelo tamanho do maior nome, senao um nome partido na fronteira escaparia.</summary>
    private static HashSet<string> ProcurarNoBinario(string caminho, string[] nomes)
    {
        var achados = new HashSet<string>(StringComparer.Ordinal);
        var sobra = nomes.Max(n => n.Length);
        var bloco = new byte[8 * 1024 * 1024];
        using var fs = File.OpenRead(caminho);
        var inicio = 0;
        int lidos;
        while ((lidos = fs.Read(bloco, inicio, bloco.Length - inicio)) > 0 || inicio > 0)
        {
            var total = inicio + lidos;
            var texto = System.Text.Encoding.ASCII.GetString(bloco, 0, total);
            foreach (var n in nomes)
                if (texto.Contains(n, StringComparison.Ordinal)) achados.Add(n);
            if (lidos == 0) break;
            inicio = Math.Min(sobra, total);
            Array.Copy(bloco, total - inicio, bloco, 0, inicio);
        }
        return achados;
    }

    [Fact] // a origem do DOCUMENTO (moffice-doc.local: Editor.bin + midia do x2t, inclusive SVG derivado do documento)
    // NAO pode falar com a ponte nem com ascdesktop://: e' conteudo derivado de arquivo de terceiros, e so' e'
    // buscado pela pagina do editor. Estar na lista daria "ponte"/"log" a qualquer quadro dessa origem.
    public void OrigensPonte_NaoIncluiAOrigemDoDocumento()
    {
        Assert.DoesNotContain(Midia.BaseUrlDocumento, AmbienteWebView2.OrigensPonte);
        Assert.DoesNotContain("https://moffice-doc.local", AmbienteWebView2.OrigensPonte);
        Assert.Contains("https://moffice.local", AmbienteWebView2.OrigensPonte); // controle positivo: a lista nao ficou vazia
    }

    [Fact] // pasta de dados por app dentro de Pastas.Raiz (%LOCALAPPDATA%\mOffice sem a variavel), nunca em %TEMP%
    public void PastaDados_PorApp()
    {
        var p = AmbienteWebView2.PastaDadosUsuario("mTexto");
        Assert.StartsWith(Path.Combine(Pastas.Raiz, "mTexto"), p);
        Assert.DoesNotContain(Path.GetTempPath(), p);
        // A promessa do PRODUTO ("nada em %TEMP%", spec paragrafo 4) e' sobre a raiz PADRAO: a afirmacao
        // acima depende do arranjo de teste manter MOFFICE_DADOS_DIR fora de %TEMP% (ver
        // DadosDeTeste.PastaDados); esta nao depende de nada (revisao 2).
        Assert.DoesNotContain(Path.GetTempPath(), Path.Combine(Pastas.RaizPadrao(), "mTexto", "WebView2"));
    }

    [Fact] // um app so' pode nomear uma PASTA: "..\\..\\Windows" nao pode virar pasta de dados
    public void PastaDados_RecusaNomeComCaminho()
    {
        Assert.Throws<ArgumentException>(() => AmbienteWebView2.PastaDadosUsuario(@"..\..\Windows"));
        Assert.Throws<ArgumentException>(() => AmbienteWebView2.PastaDadosUsuario(""));
    }

    [SkippableFact] // runtime fixo presente -> UsandoRuntimeFixo, versao do motor = a do exe embarcado, e UM ambiente por processo
    public void ObterAsync_UsaRuntimeFixo_QuandoPresente()
    {
        Skip.If(AmbienteWebView2.VersaoRuntimeFixo is null, "runtime fixo ausente — rode tools/webview2/obter-runtime.ps1");
        Skip.If(JanelaSta.SessaoRemota, "janela real: so' em sessao local");
        JanelaSta.Executar(async () =>
        {
            var env = await AmbienteWebView2.ObterAsync("testes");
            Assert.True(AmbienteWebView2.UsandoRuntimeFixo);
            Assert.Null(AmbienteWebView2.Aviso);
            Assert.Equal(AmbienteWebView2.VersaoRuntimeFixo, env.BrowserVersionString.Split(' ')[0]);
            var env2 = await AmbienteWebView2.ObterAsync("testes");
            Assert.Same(env, env2); // um ambiente para todas as janelas da mesma thread de interface
        });
    }

    [SkippableFact]
    // Regressao medida na Task 3: um cache estatico simples devolvia a uma SEGUNDA thread de interface o
    // mesmo objeto COM, criado na primeira — e usa-lo dava InvalidComObjectException ("COM object that has
    // been separated from its underlying RCW") assim que a thread original morria. O ambiente e' por thread
    // de interface (no produto, uma so'), e cada uma tem de receber um ambiente VIVO.
    public void ObterAsync_AmbientePorThreadDeInterface_ContinuaUsavel()
    {
        Skip.If(AmbienteWebView2.VersaoRuntimeFixo is null, "runtime fixo ausente — rode tools/webview2/obter-runtime.ps1");
        Skip.If(JanelaSta.SessaoRemota, "janela real: so' em sessao local");
        CoreWebView2Environment? primeiro = null;
        JanelaSta.Executar(async () => { primeiro = await AmbienteWebView2.ObterAsync("testes"); });
        // A thread STA acima ja' terminou aqui. A segunda tem de ganhar um ambiente proprio e utilizavel.
        JanelaSta.Executar(async () =>
        {
            var segundo = await AmbienteWebView2.ObterAsync("testes");
            Assert.NotSame(primeiro, segundo);
            Assert.Equal(AmbienteWebView2.VersaoRuntimeFixo, segundo.BrowserVersionString.Split(' ')[0]);
        });
    }
}
