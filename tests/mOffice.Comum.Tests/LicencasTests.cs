using System.Text;
using System.Text.RegularExpressions;
using mOffice.Comum.Dialogos;

namespace mOffice.Comum.Tests;

/// <summary>Guardas dos arquivos de licenca que o instalador leva em <c>{app}\licencas</c> (Plano 5, Task 4):
/// o inventario (<c>docs/licencas/INVENTARIO.md</c>) contra a pasta, a lista do Sobre contra o inventario e
/// o texto de termos do instalador (<c>AVISO-mOffice.txt</c>). O Sobre renderizado (janela real) esta em
/// <c>DialogosTests</c>.</summary>
public class LicencasTests
{
    internal static string Raiz => AplicativoOfficeTests.RaizRepo();
    internal static string PastaLicencas => Path.Combine(Raiz, "docs", "licencas");

    /// <summary>Primeira coluna de cada linha da tabela "## Componentes" do INVENTARIO.md, sem formatacao.</summary>
    internal static List<string> ComponentesDoInventario()
    {
        var linhas = File.ReadAllLines(Path.Combine(PastaLicencas, "INVENTARIO.md"));
        var ini = Array.FindIndex(linhas, l => l.Trim() == "## Componentes");
        Assert.True(ini >= 0, "INVENTARIO.md sem a secao '## Componentes'");
        var nomes = new List<string>();
        foreach (var l in linhas.Skip(ini + 1))
        {
            if (l.StartsWith("## ")) break;
            if (!l.StartsWith('|')) continue;
            var celula = l.Split('|')[1].Trim().Replace("**", "").Replace("`", "");
            if (celula is "Componente" || Regex.IsMatch(celula, "^-+$")) continue;
            nomes.Add(celula);
        }
        return nomes;
    }

    [Fact] // a tabela e' lida de verdade (sem isto, um parse quebrado devolveria 0 linhas e tudo "passaria")
    public void Inventario_TabelaLidaENaoVazia()
    {
        var nomes = ComponentesDoInventario();
        Assert.True(nomes.Count >= 20, $"so {nomes.Count} componentes lidos do INVENTARIO.md");
        Assert.Contains("ICU", nomes);
        Assert.Contains("ONLYOFFICE Desktop Editors", nomes);
        Assert.Equal(nomes.Count, nomes.Distinct().Count());
    }

    [Fact] // as duas direcoes: cada linha do inventario tem item no Sobre e cada item do Sobre esta no inventario
    public void Atribuicoes_IgualAoInventario()
    {
        var inventario = ComponentesDoInventario();
        var sobre = Atribuicoes.Itens("1.0.0", "9.4.0.129").Select(i => i.Componente).ToList();
        Assert.Empty(inventario.Except(sobre));
        Assert.Empty(sobre.Except(inventario));
    }

    [Fact] // todo arquivo que o inventario cita existe; todo arquivo da pasta (menos o rascunho pessoal) e' citado
    public void Inventario_ArquivosExistemETodosSaoCitados()
    {
        var texto = File.ReadAllText(Path.Combine(PastaLicencas, "INVENTARIO.md"));
        var citados = Regex.Matches(texto, @"`((?:LICENCA|AVISO|MODIFICACOES)-[^`/\\]+\.(?:txt|md))`")
            .Select(m => m.Groups[1].Value).Distinct().ToList();
        Assert.True(citados.Count >= 15, $"so {citados.Count} arquivos citados");
        foreach (var c in citados)
            Assert.True(File.Exists(Path.Combine(PastaLicencas, c)), $"INVENTARIO.md cita {c}, que nao existe em docs/licencas");
        var naPasta = Directory.GetFiles(PastaLicencas).Select(Path.GetFileName)
            .Where(n => !n!.StartsWith("email-", StringComparison.OrdinalIgnoreCase) && n != "INVENTARIO.md");
        foreach (var n in naPasta)
            Assert.True(texto.Contains(n!, StringComparison.Ordinal), $"{n} vai no instalador mas o INVENTARIO.md nao o cita");
    }

    [Fact] // a versao do mOffice no inventario aponta para a fonte unica (src/Directory.Build.props), sem numero fixo
    public void Inventario_VersaoDoMOffice_SemNumeroFixo()
    {
        var linha = File.ReadAllLines(Path.Combine(PastaLicencas, "INVENTARIO.md"))
            .Single(l => l.StartsWith("| mOffice (mTexto, mPlanilha, mSlides) |", StringComparison.Ordinal));
        var versao = linha.Split('|')[2];
        Assert.Contains("src/Directory.Build.props", versao);
        Assert.DoesNotMatch(@"\d+\.\d+\.\d+", versao);
    }

    [Fact] // a licenca do mOffice que viaja no pacote e' a do repositorio, byte a byte
    public void LicencaMOffice_IgualAoLicenseDaRaiz() =>
        Assert.Equal(File.ReadAllBytes(Path.Combine(Raiz, "LICENSE")),
                     File.ReadAllBytes(Path.Combine(PastaLicencas, "LICENCA-mOffice-AGPL-3.0.txt")));

    [Fact] // o texto de termos do instalador: UTF-8 com BOM (o Inno le .txt sem BOM como ANSI e estraga os acentos)
           // e os elementos obrigatorios
    public void AvisoMOffice_BomEConteudo()
    {
        var caminho = Path.Combine(PastaLicencas, "AVISO-mOffice.txt");
        var bytes = File.ReadAllBytes(caminho);
        Assert.True(bytes.Length > 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "AVISO-mOffice.txt sem BOM UTF-8");
        var t = new UTF8Encoding(false, true).GetString(bytes, 3, bytes.Length - 3); // lanca se nao for UTF-8 valido
        Assert.Contains("AGPL-3.0", t);
        Assert.Contains(DialogoSobre.CodigoFonte, t);
        Assert.Contains("FONTE-CORRESPONDENTE.md", t);
        Assert.Contains("ONLYOFFICE is a trademark of Ascensio System SIA.", t);
        Assert.Contains("2(b)(ii)", t);
        Assert.Contains(@"licencas\LICENCA-WebView2-Runtime.txt", t);
        Assert.Contains("SEM NENHUMA", t);
        Assert.Contains("GARANTIA", t);
        // a versao do runtime citada e' a mesma do arquivo de licenca do runtime pinado
        var versaoRuntime = Regex.Match(File.ReadAllText(Path.Combine(PastaLicencas, "LICENCA-WebView2-Runtime.txt")),
            @"Versao \.+: (\S+)").Groups[1].Value;
        Assert.Matches(@"^\d+\.\d+\.\d+\.\d+$", versaoRuntime);
        Assert.Contains($"versão {versaoRuntime}", t);
        Assert.Contains(Atribuicoes.Itens("x", "y"), i => i.Componente == "Microsoft Edge WebView2 Runtime" && i.Versao == versaoRuntime);
    }

    [Fact] // pasta de licencas: instalado ({app}\licencas) vence; em desenvolvimento, docs\licencas subindo; so' com INVENTARIO.md
    public void Sobre_LocalizaPastaDeLicencas()
    {
        var tmp = Directory.CreateTempSubdirectory("moffice-lic-").FullName;
        try
        {
            // desenvolvimento: <raiz>\docs\licencas, exe em <raiz>\src\mTexto\bin\Debug
            var bin = Directory.CreateDirectory(Path.Combine(tmp, "src", "mTexto", "bin", "Debug")).FullName;
            Assert.Null(DialogoSobre.LocalizarPastaLicencas(bin));
            var dev = Directory.CreateDirectory(Path.Combine(tmp, "docs", "licencas")).FullName;
            Assert.Null(DialogoSobre.LocalizarPastaLicencas(bin)); // sem INVENTARIO.md nao conta
            File.WriteAllText(Path.Combine(dev, "INVENTARIO.md"), "x");
            Assert.Equal(dev, DialogoSobre.LocalizarPastaLicencas(bin));
            // instalado: <app>\licencas vence o de desenvolvimento
            var inst = Directory.CreateDirectory(Path.Combine(bin, "licencas")).FullName;
            Assert.Equal(dev, DialogoSobre.LocalizarPastaLicencas(bin)); // ainda sem INVENTARIO.md
            File.WriteAllText(Path.Combine(inst, "INVENTARIO.md"), "x");
            Assert.Equal(inst, DialogoSobre.LocalizarPastaLicencas(bin));
        }
        finally { Directory.Delete(tmp, true); }
    }
}
