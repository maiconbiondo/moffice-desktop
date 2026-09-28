using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using mOffice.Comum.Atualizacao;

namespace mOffice.Comum.Tests;

/// <summary>Guardas do <c>tools/export-public.ps1</c> (Plano 5, Task 5), a copia curada que vai para o repositorio
/// publico. O script roda de verdade: numa MAQUETE de repositorio (git init em %TEMP%, com o script copiado para
/// o <c>tools/</c> dela) para os casos plantados, e no repositorio real para provar que a arvore de hoje exporta
/// limpa.
///
/// Este arquivo VAI para o publico, entao nao pode conter nenhum token proibido: o token literal plantado e' lido
/// do proprio script (que nunca e' exportado), o usuario vem de <see cref="Environment.UserName"/> e o slug solto
/// de <see cref="ServicoAtualizacao.Repositorio"/>. Na copia publica o script nao existe e tudo aqui e' pulado.</summary>
public class ExportPublicoTests
{
    private static string Raiz => AplicativoOfficeTests.RaizRepo();
    private static string Script => Path.Combine(Raiz, "tools", "export-public.ps1");

    private static void ExigeScript() =>
        Skip.IfNot(File.Exists(Script), "tools/export-public.ps1 ausente (copia publica: o script nunca e' exportado)");

    /// <summary>Os literais de <c>$TokensLiterais = @( ... )</c> do script.</summary>
    internal static List<string> TokensLiterais(string textoScript)
    {
        var bloco = Regex.Match(textoScript, @"\$TokensLiterais = @\((?<c>.*?)\r?\n\)", RegexOptions.Singleline);
        Assert.True(bloco.Success, "bloco $TokensLiterais = @( ... ) nao encontrado no script");
        return Regex.Matches(bloco.Groups["c"].Value, "'([^']+)'").Select(m => m.Groups[1].Value).ToList();
    }

    /// <summary>Nomes genericos do Windows que NAO contam como segmento sensivel do caminho do repositorio - a MESMA
    /// lista de <c>$SegmentosGenericos</c> do script (<see cref="Script_SegmentosGenericosIguaisAosDoTeste"/> confere).
    /// Usada tambem por <see cref="InstaladorTests.SemCaminhoAbsolutoDeMaquina"/>.</summary>
    internal static readonly string[] SegmentosGenericos =
        ["Users", "AppData", "Local", "LocalLow", "Roaming", "Temp", "Tmp", "home", "source", "repos", "src", "git", "dev",
         "Documents", "Desktop", "Downloads", "OneDrive", "Projects", "Public"];

    private static (int Codigo, string Saida) Rodar(string script, string destino, bool semAmostras = false, string? perfil = null)
    {
        var psi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, @"WindowsPowerShell\v1.0\powershell.exe"))
        { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
        foreach (var a in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-Destino", destino })
            psi.ArgumentList.Add(a);
        if (semAmostras) psi.ArgumentList.Add("-SemAmostras");
        if (perfil is not null) psi.Environment["USERPROFILE"] = perfil; // perfil FALSO, dentro da maquete (nunca o real)
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync();
        var saida = p.StandardOutput.ReadToEnd();
        if (!p.WaitForExit(300_000)) { p.Kill(true); Assert.Fail("export-public.ps1 nao terminou em 5 min"); }
        return (p.ExitCode, saida + err.Result);
    }

    private static string Git(string pasta, params string[] args)
    {
        var psi = new ProcessStartInfo("git")
        {
            UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
        };
        foreach (var a in new[] { "-C", pasta, "-c", "user.name=teste", "-c", "user.email=teste@example.invalid", "-c", "core.autocrlf=false",
                                  "-c", "core.quotepath=off" }.Concat(args))
            psi.ArgumentList.Add(a);
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync();
        var saida = p.StandardOutput.ReadToEnd();
        Assert.True(p.WaitForExit(60_000), "git nao terminou");
        Assert.True(p.ExitCode == 0, $"git {string.Join(' ', args)} falhou: {err.Result}");
        return saida;
    }

    /// <summary>Repositorio de mentira com a mesma forma do real: arquivos que entram (src, tests, tools, docs/licencas,
    /// raiz), arquivos que NUNCA entram (rastreados de proposito, para a exclusao ser posta a prova) e amostras
    /// "reais" fora do git em poc/amostras/. O script (real ou mutado) e' copiado para tools/.</summary>
    private sealed class Maquete : IDisposable
    {
        public const string AmostraAscii = "Relatorio Sigiloso Cliente";
        public const string AmostraAcento = "Balanço Anual Confidencial";
        public readonly string Base, Repo, Destino, NomeBase;
        public string ScriptNaMaquete => Path.Combine(Repo, "tools", "export-public.ps1");
        private readonly bool _comAmostras;

        public Maquete(string? textoScript = null, bool comAmostras = true)
        {
            _comAmostras = comAmostras;
            NomeBase = "moexp" + Guid.NewGuid().ToString("N")[..8];
            Base = Path.Combine(Path.GetTempPath(), NomeBase);
            Repo = Path.Combine(Base, "repo");
            Destino = Path.Combine(Base, "saida");
            try { Montar(textoScript); }
            catch { Dispose(); throw; }
        }

        private void Montar(string? textoScript)
        {
            Directory.CreateDirectory(Repo);
            Escrever("mOffice.slnx", "<Solution />\n");
            Escrever("Directory.Build.props", "<Project />\n");
            Escrever(".gitignore", "poc/amostras/\n");
            Escrever("LICENSE", "licenca de teste\n");
            Escrever("src/Directory.Build.props", "<Project><PropertyGroup><Version>9.8.7</Version></PropertyGroup></Project>\n");
            for (var i = 0; i < 110; i++) Escrever($"src/App/Arquivo{i:000}.cs", $"// arquivo {i}\n");
            Escrever("tests/T/Teste.cs", "// teste\n");
            Escrever("src/App/Configuração.cs", "// nome com acento: o git nao pode devolver o caminho escapado\n");
            Escrever("tools/util.ps1", "# util\n");
            // ZIP limpo: a varredura das entradas nao pode dar falso positivo (controle positivo do ramo ZIP)
            var docx = Path.Combine(Repo, "tests", "T", "Fixtures", "limpo.docx");
            Directory.CreateDirectory(Path.GetDirectoryName(docx)!);
            File.WriteAllBytes(docx, DocxComAutor("Autor Ficticio"));
            Escrever("docs/licencas/LICENCA-Terceiro.txt", "This License shall not be construed as a waiver.\n");
            Escrever("docs/licencas/email-rascunho.md", "rascunho pessoal\n");
            Escrever("docs/publico/README.md", "# Projeto publico\n");
            Escrever("docs/superpowers/plano.md", "historico interno\n");
            Escrever("poc/relatorios/r.md", "relatorio\n");
            Escrever("third_party/motor/x.txt", "binario\n");
            Escrever(".superpowers/s.md", "scratch\n");
            if (_comAmostras)
            {
                Escrever("poc/amostras/" + AmostraAscii + ".xlsx", "x");
                Escrever("poc/amostras/" + AmostraAcento + ".docx", "x");
            }
            Escrever("tools/export-public.ps1", textoScript ?? File.ReadAllText(Script));
            Git(Repo, "init", "-q", "-b", "main");
            Git(Repo, "add", "-A");
            Git(Repo, "commit", "-q", "-m", "maquete");
            var rastreados = Git(Repo, "ls-files");
            foreach (var r in new[] { "tools/export-public.ps1", "docs/licencas/email-rascunho.md", "docs/superpowers/plano.md",
                                      "poc/relatorios/r.md", "third_party/motor/x.txt", ".superpowers/s.md" })
                Assert.Contains(r, rastreados); // a exclusao so' prova algo se o arquivo proibido estiver RASTREADO
            Assert.DoesNotContain("poc/amostras", rastreados);
        }

        public void Escrever(string rel, string conteudo)
        {
            var caminho = Path.Combine(Repo, rel.Replace('/', '\\'));
            Directory.CreateDirectory(Path.GetDirectoryName(caminho)!);
            File.WriteAllText(caminho, conteudo, new UTF8Encoding(false));
        }

        /// <summary>Planta num arquivo JA rastreado (o script copia da arvore de trabalho).</summary>
        public void Plantar(string conteudo) => Escrever("src/App/Arquivo042.cs", "// arquivo 42\n" + conteudo + "\n");

        /// <summary>Arquivo NOVO, rastreado, com bytes quaisquer (UTF-16, docx...).</summary>
        public void PlantarBytes(string rel, byte[] bytes)
        {
            var caminho = Path.Combine(Repo, rel.Replace('/', '\\'));
            Directory.CreateDirectory(Path.GetDirectoryName(caminho)!);
            File.WriteAllBytes(caminho, bytes);
            Git(Repo, "add", "-A");
        }

        /// <summary>Arquivo NOVO, rastreado, cujo NOME carrega o token (conteudo inocente).</summary>
        public void PlantarNome(string nome)
        {
            Escrever("src/App/" + nome + ".txt", "inocente\n");
            Git(Repo, "add", "-A");
        }

        public (int Codigo, string Saida) Rodar(string? destino = null, bool semAmostras = false, string? perfil = null) =>
            ExportPublicoTests.Rodar(ScriptNaMaquete, destino ?? Destino, semAmostras, perfil);

        public void Dispose()
        {
            for (var i = 0; i < 5; i++)
            {
                try
                {
                    if (Directory.Exists(Base))
                    {
                        foreach (var f in Directory.EnumerateFiles(Base, "*", SearchOption.AllDirectories)) File.SetAttributes(f, FileAttributes.Normal);
                        Directory.Delete(Base, true);
                    }
                    return;
                }
                catch (Exception) when (i < 4) { Thread.Sleep(300); }
            }
        }
    }

    [SkippableFact] // a lista literal existe e tem o nome do escritorio e o pessoal (os dois lados do "e-mail pessoal exato")
    public void Script_ListaLiteralNaoVazia()
    {
        ExigeScript();
        var tokens = TokensLiterais(File.ReadAllText(Script));
        Assert.True(tokens.Count >= 5, $"so {tokens.Count} tokens literais");
        Assert.Contains(tokens, t => t.Contains('@'));
    }

    [SkippableFact] // lista vazia = guarda incapaz de verificar = erro (sai 1, nada no destino). Mutacao conferida.
    public void ListaVazia_EErro()
    {
        ExigeScript();
        var original = File.ReadAllText(Script);
        var mutado = Regex.Replace(original, @"\$TokensLiterais = @\(.*?\r?\n\)", "$$TokensLiterais = @()", RegexOptions.Singleline);
        Assert.NotEqual(original, mutado);
        Assert.Contains("$TokensLiterais = @()", mutado);
        using var m = new Maquete(mutado);
        var (codigo, saida) = m.Rodar();
        Assert.Equal(1, codigo);
        Assert.Contains("Lista de tokens proibidos vazia", saida);
        Assert.False(Directory.Exists(m.Destino));
    }

    [SkippableFact] // controle POSITIVO: a maquete limpa exporta; o que nunca entra ficou de fora; autor neutro
    public void MaqueteLimpa_ExportaSoAAllowlist()
    {
        ExigeScript();
        using var m = new Maquete();
        var (codigo, saida) = m.Rodar();
        Assert.True(codigo == 0, saida);
        Assert.True(Directory.Exists(Path.Combine(m.Destino, ".git")));
        var exportados = Git(m.Destino, "ls-files").Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Contains("README.md", exportados);
        Assert.Contains("LICENSE", exportados);
        Assert.Contains("mOffice.slnx", exportados);
        Assert.Contains("Directory.Build.props", exportados);
        Assert.Contains("src/Directory.Build.props", exportados);
        Assert.Contains("tests/T/Teste.cs", exportados);
        Assert.Contains("src/App/Configuração.cs", exportados);
        Assert.Contains("tests/T/Fixtures/limpo.docx", exportados);
        Assert.Contains("tools/util.ps1", exportados);
        Assert.Contains("docs/licencas/LICENCA-Terceiro.txt", exportados); // tem "construed": nao dispara o usuario
        Assert.Equal("# Projeto publico\n", File.ReadAllText(Path.Combine(m.Destino, "README.md")));
        foreach (var proibido in new[] { "tools/export-public.ps1", "docs/licencas/email-", "docs/superpowers/", "docs/publico/",
                                         "poc/", "third_party/", ".superpowers/" })
            Assert.DoesNotContain(exportados, f => f.StartsWith(proibido, StringComparison.OrdinalIgnoreCase));
        // 3 nomeados da raiz + src/Directory.Build.props + 110 de src/App + Configuração.cs + tests (2) + tools/util + 1 licenca + LICENSE + README
        Assert.Equal(110 + 11, exportados.Length);
        Assert.Contains("entradas de ZIP", saida);

        // uma exportacao anterior no mesmo destino e' substituida (e so ela - ver Destino_ComConteudoAlheio_Recusa)
        var (codigo2, saida2) = m.Rodar();
        Assert.True(codigo2 == 0, saida2);
        Assert.Equal("mOffice <moffice@users.noreply.github.com>|mOffice v9.8.7",
                     Git(m.Destino, "log", "--format=%an <%ae>|%s").Trim());
    }

    public static TheoryData<string> Plantios => new()
    {
        "literal", "literal-maiusculas", "usuario-em-caminho", "caminho-da-maquete", "segmento-da-maquete",
        "amostra", "amostra-sem-acento", "amostra-utf8", "amostra-maiusculas-acento", "gmail", "slug-solto",
        "outro-repo-do-dono", "nome-de-arquivo", "utf16le", "utf16be", "docx-core-xml", "zip-aninhado", "zip-cifrado",
    };

    /// <summary>ZIP em memoria (compactado: o texto das entradas NAO aparece cru nos bytes).</summary>
    private static byte[] Zip(params (string Nome, byte[] Conteudo)[] entradas)
    {
        using var ms = new MemoryStream();
        using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var (nome, conteudo) in entradas)
            {
                using var s = z.CreateEntry(nome, System.IO.Compression.CompressionLevel.Optimal).Open();
                s.Write(conteudo);
            }
        return ms.ToArray();
    }

    /// <summary>Um .docx minimo com o token no autor (docProps/core.xml), como o Word grava.</summary>
    private static byte[] DocxComAutor(string autor) => Zip(
        ("[Content_Types].xml", Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")),
        ("word/document.xml", Encoding.UTF8.GetBytes("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>" + new string(' ', 400))),
        ("docProps/core.xml", Encoding.UTF8.GetBytes("<cp:coreProperties xmlns:cp=\"x\" xmlns:dc=\"y\"><dc:creator>" + autor + "</dc:creator>" +
                                                     new string(' ', 400) + "</cp:coreProperties>")));

    /// <summary>Liga o bit 0 ("cifrada") do general-purpose flag em todo cabecalho local e do diretorio central.</summary>
    private static byte[] MarcarCifrado(byte[] zip)
    {
        var b = (byte[])zip.Clone();
        var marcados = 0;
        for (var i = 0; i + 4 <= b.Length; i++)
        {
            var sig = BitConverter.ToUInt32(b, i);
            if (sig == 0x04034b50) { b[i + 6] |= 1; marcados++; }
            else if (sig == 0x02014b50) { b[i + 8] |= 1; marcados++; }
        }
        Assert.Equal(2, marcados); // um local + um central: a mutacao pegou
        return b;
    }

    /// <summary>A mutacao so' prova algo se o token NAO estiver cru nos bytes (senao a varredura latin1 ja' pegaria).</summary>
    private static void AssertTokenNaoCru(byte[] bytes, string token) =>
        Assert.DoesNotContain(token, Encoding.Latin1.GetString(bytes), StringComparison.OrdinalIgnoreCase);

    [SkippableTheory] // controle NEGATIVO: cada tipo de token plantado faz o script sair 1 e APAGAR o destino
    [MemberData(nameof(Plantios))]
    public void TokenPlantado_RecusaEApagaODestino(string tipo)
    {
        ExigeScript();
        using var m = new Maquete();
        var literal = TokensLiterais(File.ReadAllText(Script))[0];
        var alvo = "src/App/Arquivo042.cs";
        string? esperado = null;
        switch (tipo)
        {
            case "literal": m.Plantar("// " + literal); break;
            case "literal-maiusculas": m.Plantar("// " + literal.ToUpperInvariant()); break;
            case "usuario-em-caminho": m.Plantar(@"// C:\Users\" + Environment.UserName + @"\AppData"); break;
            case "caminho-da-maquete": m.Plantar("// " + m.Repo); break;
            case "segmento-da-maquete": m.Plantar(@"// D:\outro\" + m.NomeBase + @"\x"); break;
            case "amostra": m.Plantar("// " + Maquete.AmostraAscii); break;
            case "amostra-sem-acento": m.Plantar("// Balanco Anual Confidencial"); break;
            case "amostra-utf8": m.Plantar("// " + Maquete.AmostraAcento); break;
            case "gmail": m.Plantar("// fulano.teste" + "@" + "gmail.com"); break;
            case "slug-solto": m.Plantar("var dono = \"" + ServicoAtualizacao.Repositorio.Split('/')[0] + "\";"); break;
            case "amostra-maiusculas-acento": m.Plantar("// " + Maquete.AmostraAcento.ToUpperInvariant()); break;
            case "outro-repo-do-dono": m.Plantar("// https://github.com/" + ServicoAtualizacao.Repositorio.Split('/')[0] + "/outro-repositorio"); break;
            case "nome-de-arquivo": m.PlantarNome(Maquete.AmostraAscii); alvo = "src/App/" + Maquete.AmostraAscii + ".txt"; break;
            case "utf16le":
            {
                var b = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("autor: " + literal + "\r\n")).ToArray();
                AssertTokenNaoCru(b, literal);
                alvo = "src/App/utf16le.txt"; m.PlantarBytes(alvo, b); break;
            }
            case "utf16be":
            {
                var b = Encoding.BigEndianUnicode.GetBytes("x autor: " + literal);
                AssertTokenNaoCru(b, literal);
                alvo = "src/App/utf16be.bin"; m.PlantarBytes(alvo, b); break;
            }
            case "docx-core-xml":
            {
                var b = DocxComAutor(literal);
                AssertTokenNaoCru(b, literal);
                alvo = "tests/T/Fixtures/modelo.docx!/docProps/core.xml"; m.PlantarBytes("tests/T/Fixtures/modelo.docx", b); break;
            }
            case "zip-cifrado":
            {
                // entrada INOCENTE, mas marcada como cifrada (bit 0 do general-purpose flag): nao da' para varrer -> recusa
                var b = MarcarCifrado(Zip(("segredo.txt", Encoding.UTF8.GetBytes("inocente " + new string('x', 200)))));
                alvo = "tests/T/Fixtures/cifrado.zip"; m.PlantarBytes(alvo, b); esperado = "CIFRADA(s): segredo.txt"; break;
            }
            case "zip-aninhado":
            {
                var b = Zip(("interno/pacote.zip", DocxComAutor(literal)));
                AssertTokenNaoCru(b, literal);
                alvo = "tests/T/Fixtures/lote.zip!/interno/pacote.zip!/docProps/core.xml"; m.PlantarBytes("tests/T/Fixtures/lote.zip", b); break;
            }
            default: throw new ArgumentOutOfRangeException(nameof(tipo));
        }
        var (codigo, saida) = m.Rodar();
        Assert.True(codigo == 1, $"{tipo}: esperava exit 1, veio {codigo}\n{saida}");
        Assert.Contains("GUARDA DE CONTEUDO RECUSOU", saida);
        if (esperado is not null) Assert.Contains(esperado, saida);
        Assert.Contains(alvo, saida);
        Assert.False(Directory.Exists(m.Destino), $"{tipo}: o destino nao foi apagado");
    }

    [SkippableFact] // o usuario casa por palavra inteira: "construed" (textos GPL/AGPL) e o nome colado a letras nao disparam;
                    // o slug publico em contexto de URL/repositorio tambem nao
    public void UsuarioColadoALetrasESlugEmUrl_NaoDisparam()
    {
        ExigeScript();
        using var m = new Maquete();
        var u = Environment.UserName;
        m.Plantar($"// shall not be construed; {u}d {u}s x{u}\n" +
                  $"// https://github.com/{ServicoAtualizacao.Repositorio}/releases\n" +
                  $"// https://api.github.com/repos/{ServicoAtualizacao.Repositorio}/releases/latest\n" +
                  $"const string R = \"{ServicoAtualizacao.Repositorio}\";");
        var (codigo, saida) = m.Rodar();
        Assert.True(codigo == 0, saida);
        Assert.True(Directory.Exists(Path.Combine(m.Destino, ".git")));
    }

    [SkippableTheory] // a guarda ESTRUTURAL (lista final conferida contra as pastas proibidas) pega uma exclusao removida,
                      // mesmo antes da varredura de conteudo. Mutacao conferida.
    [InlineData("$Exclusoes += $EsteScript", "tools/export-public.ps1")]
    [InlineData("$Exclusoes += 'docs/licencas/email-'", "docs/licencas/email-rascunho.md")]
    public void ExclusaoRemovida_AGuardaEstruturalRecusa(string linhaDaExclusao, string arquivoQueVazaria)
    {
        ExigeScript();
        var original = File.ReadAllText(Script);
        Assert.Single(Regex.Matches(original, Regex.Escape(linhaDaExclusao)));
        var mutado = original.Replace(linhaDaExclusao, "# exclusao removida pelo teste");
        Assert.NotEqual(original, mutado);
        using var m = new Maquete(mutado);
        var (codigo, saida) = m.Rodar();
        Assert.Equal(1, codigo);
        Assert.Contains($"GUARDA: a allowlist contem '{arquivoQueVazaria}'", saida);
        Assert.False(Directory.Exists(m.Destino));
    }

    [SkippableFact] // a lista de segmentos genericos do script e' a mesma que os testes usam (InstaladorTests)
    public void Script_SegmentosGenericosIguaisAosDoTeste()
    {
        ExigeScript();
        var bloco = Regex.Match(File.ReadAllText(Script), @"\$SegmentosGenericos = @\((?<c>.*?)\)", RegexOptions.Singleline);
        Assert.True(bloco.Success, "$SegmentosGenericos nao encontrado no script");
        var doScript = Regex.Matches(bloco.Groups["c"].Value, "'([^']+)'").Select(m => m.Groups[1].Value).ToArray();
        Assert.Equal(SegmentosGenericos, doScript);
    }

    [SkippableFact] // sem poc/amostras/ a guarda nao pode conferir os nomes: erro; -SemAmostras aceita explicitamente
    public void SemPastaDeAmostras_EErro_SalvoComSemAmostras()
    {
        ExigeScript();
        using var m = new Maquete(comAmostras: false);
        var (codigo, saida) = m.Rodar();
        Assert.Equal(1, codigo);
        Assert.Contains("poc/amostras/ ausente ou sem arquivos", saida);
        Assert.False(Directory.Exists(m.Destino));
        var (codigo2, saida2) = m.Rodar(semAmostras: true);
        Assert.True(codigo2 == 0, saida2);
        Assert.Contains("-SemAmostras", saida2);
    }

    public static TheoryData<string> DestinosProibidos => new()
    {
        "repositorio", "dentro-do-repositorio", "acima-do-repositorio", "raiz-do-drive", "perfil", "acima-do-perfil", "conteudo-alheio",
    };

    [SkippableTheory] // o destino e' conferido ANTES de apagar qualquer coisa: cada caso recusa (exit 1) e o sentinela sobrevive.
                      // O perfil e' FALSO (USERPROFILE do processo filho aponta para dentro da maquete): uma regressao
                      // nunca poderia apagar o perfil real.
    [MemberData(nameof(DestinosProibidos))]
    public void Destino_Perigoso_Recusa(string caso)
    {
        ExigeScript();
        using var m = new Maquete();
        var perfil = Path.Combine(m.Base, "casa", "perfil");
        Directory.CreateDirectory(perfil);
        File.WriteAllText(Path.Combine(perfil, "sentinela.txt"), "x");
        var alheio = Path.Combine(m.Base, "alheio");
        Directory.CreateDirectory(alheio);
        File.WriteAllText(Path.Combine(alheio, "sentinela.txt"), "x");
        var (destino, mensagem) = caso switch
        {
            "repositorio" => (m.Repo, "e' o repositorio"),
            "dentro-do-repositorio" => (Path.Combine(m.Repo, "src"), "e' o repositorio"),
            "acima-do-repositorio" => (m.Base, "e' o repositorio"),
            "raiz-do-drive" => (Path.GetPathRoot(m.Base)!, "raiz de um drive"),
            "perfil" => (perfil, "perfil do usuario"),
            "acima-do-perfil" => (Path.Combine(m.Base, "casa"), "perfil do usuario"),
            "conteudo-alheio" => (alheio, "nao e' uma exportacao anterior"),
            _ => throw new ArgumentOutOfRangeException(nameof(caso)),
        };
        var (codigo, saida) = m.Rodar(destino, perfil: perfil);
        Assert.True(codigo == 1, $"{caso}: esperava exit 1, veio {codigo}\n{saida}");
        Assert.Contains("DESTINO RECUSADO", saida);
        Assert.Contains(mensagem, saida);
        Assert.True(File.Exists(Path.Combine(perfil, "sentinela.txt")), $"{caso}: o perfil falso foi mexido");
        Assert.True(File.Exists(Path.Combine(alheio, "sentinela.txt")), $"{caso}: a pasta alheia foi mexida");
        Assert.True(File.Exists(Path.Combine(m.Repo, "mOffice.slnx")), $"{caso}: o repositorio foi mexido");
    }

    // ---- Outro nome para o mesmo lugar (Critico da re-revisao) -------------------------------------------------
    // SEMPRE contra a maquete em %TEMP%, nunca o repositorio real. Sentinelas: as amostras "reais" (poc/amostras, fora
    // do git - o que se perderia de verdade) e o mOffice.slnx da maquete.

    /// <summary><c>subst</c> de uma letra livre para a pasta temporaria; desfeito no Dispose.</summary>
    private sealed class Subst : IDisposable
    {
        public readonly string? Letra;
        public Subst(string alvo)
        {
            var usadas = DriveInfo.GetDrives().Select(d => char.ToUpperInvariant(d.Name[0])).ToHashSet();
            var livre = "ZYXWVUTSRQPONMLKJIHG".FirstOrDefault(c => !usadas.Contains(c));
            if (livre == default) return;
            if (SubstExe($"{livre}: \"{alvo.TrimEnd('\\')}\"") == 0 && Directory.Exists($"{livre}:\\")) Letra = $"{livre}:";
        }
        private static int SubstExe(string args)
        {
            var psi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "subst.exe"), args)
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
            using var p = Process.Start(psi)!;
            p.WaitForExit(15_000);
            return p.ExitCode;
        }
        public void Dispose() { if (Letra is not null) SubstExe($"{Letra} /D"); }
    }

    private static string Unc(string caminho) => $@"\\localhost\{caminho[0]}$\{caminho[3..]}";

    private static void SentinelasIntactas(Maquete m, string caso)
    {
        Assert.True(File.Exists(Path.Combine(m.Repo, "poc", "amostras", Maquete.AmostraAscii + ".xlsx")), $"{caso}: amostra (sentinela) apagada");
        Assert.True(File.Exists(Path.Combine(m.Repo, "mOffice.slnx")), $"{caso}: a maquete do repositorio foi mexida");
        Assert.False(Directory.Exists(Path.Combine(m.Repo, "novo")), $"{caso}: exportou para dentro do repositorio");
    }

    public static TheoryData<string, string> Aliases => new()
    {
        { "unc", "repo" }, { "unc", "dentro-novo" }, { "unc", "acima" },
        { "subst", "repo" }, { "subst", "dentro-novo" }, { "subst", "acima" },
    };

    [SkippableTheory] // \\localhost\C$\... e letra de subst passam pelo texto; a IDENTIDADE (canario) recusa antes de apagar
    [MemberData(nameof(Aliases))]
    public void Destino_OutroNomeParaORepositorio_Recusa(string via, string onde)
    {
        ExigeScript();
        using var m = new Maquete();
        using var s = via == "subst" ? new Subst(Path.GetTempPath()) : null;
        string baseAlias;
        if (via == "unc")
        {
            baseAlias = Unc(m.Base);
            Skip.IfNot(Directory.Exists(Path.Combine(baseAlias, "repo")), "compartilhamento administrativo \\\\localhost\\<drive>$ inacessivel nesta maquina");
        }
        else
        {
            Skip.If(s!.Letra is null, "subst indisponivel (sem letra livre?)");
            baseAlias = Path.Combine(s.Letra + "\\", m.NomeBase);
            Assert.True(Directory.Exists(Path.Combine(baseAlias, "repo")), "o subst nao mostra a maquete");
        }
        var destino = onde switch
        {
            "repo" => Path.Combine(baseAlias, "repo"),
            "dentro-novo" => Path.Combine(baseAlias, "repo", "novo"),
            "acima" => baseAlias,
            _ => throw new ArgumentOutOfRangeException(nameof(onde)),
        };
        Assert.False(destino.StartsWith(m.Repo, StringComparison.OrdinalIgnoreCase)); // o texto NAO denuncia o alias
        var (codigo, saida) = m.Rodar(destino, perfil: Path.Combine(m.Base, "perfil-falso"));
        Assert.True(codigo == 1, $"{via}/{onde}: esperava exit 1, veio {codigo}\n{saida}");
        Assert.Contains("MESMO LUGAR que o repositorio", saida);
        SentinelasIntactas(m, $"{via}/{onde}");
        Assert.Empty(Directory.GetFiles(m.Repo, ".moffice-canario-*")); // o canario sempre e' apagado
    }

    [SkippableFact] // o perfil (FALSO, dentro da maquete) por outro nome tambem e' reconhecido pelo canario dele
    public void Destino_OutroNomeParaOPerfil_Recusa()
    {
        ExigeScript();
        using var m = new Maquete();
        var perfil = Path.Combine(m.Base, "casa", "perfil");
        Directory.CreateDirectory(perfil);
        File.WriteAllText(Path.Combine(perfil, "sentinela.txt"), "x");
        using var s = new Subst(Path.GetTempPath());
        Skip.If(s.Letra is null, "subst indisponivel (sem letra livre?)");
        foreach (var destino in new[] { Path.Combine(s.Letra + "\\", m.NomeBase, "casa", "perfil"), Path.Combine(s.Letra + "\\", m.NomeBase, "casa") })
        {
            var (codigo, saida) = m.Rodar(destino, perfil: perfil);
            Assert.True(codigo == 1, saida);
            Assert.True(saida.Contains("MESMO LUGAR que o perfil"), saida);
            Assert.True(File.Exists(Path.Combine(perfil, "sentinela.txt")));
            Assert.Empty(Directory.GetFiles(perfil, ".moffice-canario-*"));
        }
    }

    [SkippableFact] // MUTACAO: sem a checagem de identidade, (1) o alias do repositorio ainda e' barrado ANTES de apagar pela
                    // camada "exportacao anterior" (a maquete nao tem o marcador) - sentinelas intactas; (2) o alias para uma
                    // pasta NOVA dentro do repositorio passa (exit 0): prova que era a identidade que barrava.
    public void SemChecagemDeIdentidade_OutraCamadaAindaSalvaORepositorio()
    {
        ExigeScript();
        var original = File.ReadAllText(Script);
        const string linha = "if (-not $recusa -and (Conferir-Identidade $DestinoCheio $RepoRoot $nomeCanario)) {";
        Assert.Single(Regex.Matches(original, Regex.Escape(linha)));
        var mutado = original.Replace(linha, "if ($false) {");
        using var m = new Maquete(mutado);
        using var s = new Subst(Path.GetTempPath());
        Skip.If(s.Letra is null, "subst indisponivel (sem letra livre?)");
        var baseAlias = Path.Combine(s.Letra + "\\", m.NomeBase);
        var perfil = Path.Combine(m.Base, "perfil-falso");

        var (codigo, saida) = m.Rodar(Path.Combine(baseAlias, "repo"), perfil: perfil);
        Assert.True(codigo == 1, saida);
        Assert.True(saida.Contains("nao e' uma exportacao anterior deste script"), saida);
        Assert.DoesNotContain("MESMO LUGAR", saida);
        SentinelasIntactas(m, "mutacao/repo");

        var (codigo2, saida2) = m.Rodar(Path.Combine(baseAlias, "repo", "novo"), perfil: perfil);
        Assert.True(codigo2 == 0, "a mutacao nao tirou a checagem: " + saida2);
        Assert.True(Directory.Exists(Path.Combine(m.Repo, "novo")));
    }

    public static TheoryData<string> ExportacoesFalsas => new()
    {
        "sem-marcador", "com-pasta-interna", "sem-git", "dois-commits", "outro-autor",
        "arquivo-alterado", "arquivo-novo", "arquivo-apagado",
    };

    [SkippableTheory] // so' uma exportacao DESTE script e' substituida; o resto recusa sem apagar nada (sentinela sobrevive)
    [MemberData(nameof(ExportacoesFalsas))]
    public void Destino_NaoEExportacaoAnterior_Recusa(string caso)
    {
        ExigeScript();
        using var m = new Maquete();
        var perfil = Path.Combine(m.Base, "perfil-falso");
        var (c0, s0) = m.Rodar(perfil: perfil);
        Assert.True(c0 == 0, s0);
        var marcador = Path.Combine(m.Destino, ".moffice-export-publica");
        Assert.True(File.Exists(marcador));
        Assert.DoesNotContain(".moffice-export-publica", Git(m.Destino, "ls-files")); // nunca vai para o publico
        // sentinela IGNORADA pelo git do destino (info/exclude): sozinha ela nao conta como "alteracao nao commitada" -
        // assim cada caso abaixo recusa pelo SEU motivo, e a sentinela prova que nada foi apagado
        File.WriteAllText(Path.Combine(m.Destino, "sentinela.txt"), "x");
        File.AppendAllText(Path.Combine(m.Destino, ".git", "info", "exclude"), "/sentinela.txt\n");
        Assert.Equal("", Git(m.Destino, "status", "--porcelain", "--untracked-files=all").Trim()); // controle: limpo antes da mutacao
        string esperado;
        switch (caso)
        {
            case "sem-marcador": File.Delete(marcador); esperado = "sem o marcador"; break;
            case "com-pasta-interna": Directory.CreateDirectory(Path.Combine(m.Destino, "poc")); esperado = "contem poc"; break;
            case "dois-commits":
                Git(m.Destino, "commit", "-q", "--allow-empty", "-m", "outro", "--author", "Fulano <fulano@example.invalid>");
                esperado = "historico git"; break;
            case "sem-git":
                foreach (var f in Directory.EnumerateFiles(Path.Combine(m.Destino, ".git"), "*", SearchOption.AllDirectories)) File.SetAttributes(f, FileAttributes.Normal);
                Directory.Delete(Path.Combine(m.Destino, ".git"), true);
                esperado = "sem .git"; break;
            case "arquivo-alterado": File.AppendAllText(Path.Combine(m.Destino, "README.md"), "mexido\n"); esperado = "nao commitadas"; break;
            case "arquivo-novo": File.WriteAllText(Path.Combine(m.Destino, "src", "Novo.cs"), "// trabalho de alguem\n"); esperado = "nao commitadas"; break;
            case "arquivo-apagado": File.Delete(Path.Combine(m.Destino, "tools", "util.ps1")); esperado = "nao commitadas"; break;
            case "outro-autor":
                Git(m.Destino, "commit", "-q", "--amend", "--no-edit", "--reset-author"); // autor = o do Git() (teste@example.invalid)
                esperado = "historico git"; break;
            default: throw new ArgumentOutOfRangeException(nameof(caso));
        }
        var (codigo, saida) = m.Rodar(perfil: perfil);
        Assert.True(codigo == 1, $"{caso}: esperava exit 1, veio {codigo}\n{saida}");
        Assert.Contains("nao e' uma exportacao anterior deste script", saida);
        Assert.Contains(esperado, saida);
        Assert.True(File.Exists(Path.Combine(m.Destino, "sentinela.txt")), $"{caso}: o destino foi apagado");
    }

    [SkippableFact] // a arvore REAL de hoje exporta limpa (tokens, amostras e caminhos desta maquina), sem o que nunca vai
    public void RepositorioReal_ExportaLimpo()
    {
        ExigeScript();
        var destino = Path.Combine(Path.GetTempPath(), "moexp-real-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            var (codigo, saida) = Rodar(Script, destino);
            Assert.True(codigo == 0, saida);
            var exportados = Git(destino, "ls-files").Split('\n', StringSplitOptions.RemoveEmptyEntries);
            Assert.True(exportados.Length >= 200, $"so {exportados.Length} arquivos exportados");
            Assert.Contains("README.md", exportados);
            Assert.Contains("docs/licencas/FONTE-CORRESPONDENTE.md", exportados);
            Assert.Contains("tests/mOffice.Comum.Tests/ExportPublicoTests.cs", exportados);
            foreach (var proibido in new[] { "tools/export-public.ps1", "docs/licencas/email-", "docs/superpowers/", "poc/",
                                             "third_party/", ".superpowers/" })
                Assert.DoesNotContain(exportados, f => f.StartsWith(proibido, StringComparison.OrdinalIgnoreCase));
        }
        finally
        {
            if (Directory.Exists(destino))
            {
                foreach (var f in Directory.EnumerateFiles(destino, "*", SearchOption.AllDirectories)) File.SetAttributes(f, FileAttributes.Normal);
                Directory.Delete(destino, true);
            }
        }
    }
}
