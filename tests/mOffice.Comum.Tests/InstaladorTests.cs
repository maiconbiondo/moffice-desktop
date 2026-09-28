using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Win32;
using mOffice.Comum.Atualizacao;
using mOffice.Nucleo;

namespace mOffice.Comum.Tests;

/// <summary>
/// Plano 5, Task 3 — varredura estrutural do <c>tools/installer/moffice.iss</c> (como os testes do <c>mpdf.iss</c> no
/// mPDF): o que o instalador faz e verificado lendo o script, porque instalar de verdade so acontece no PC limpo
/// (Task 7). Mais: o verbo <c>/gerarcatalogo</c> que o <c>[Run]</c> chama, o <c>/REABRIR=</c> que o app passa, a
/// compilacao real do <c>.iss</c> pelo ISCC (sem gerar instalador) e a politica de associacao do RMM rodando
/// contra chaves de TESTE em HKCU.
/// </summary>
public class InstaladorTests
{
    private static string Raiz => AplicativoOfficeTests.RaizRepo();
    private static string CaminhoIss => Path.Combine(Raiz, "tools", "installer", "moffice.iss");
    private static string Iss => File.ReadAllText(CaminhoIss);

    /// <summary>GUID do produto: fixo para sempre (uma versao nova com outro AppId instalaria AO LADO da antiga).</summary>
    private const string AppIdFixo = "BEA3C718-DFC3-42D6-BF84-41C61CCAC772";

    internal static readonly (string ProgId, string Ext, string Exe)[] ProgIds =
    [
        ("mOffice.Texto.docx", ".docx", "mTexto.exe"),
        ("mOffice.Texto.odt", ".odt", "mTexto.exe"),
        ("mOffice.Planilha.xlsx", ".xlsx", "mPlanilha.exe"),
        ("mOffice.Planilha.ods", ".ods", "mPlanilha.exe"),
        ("mOffice.Slides.pptx", ".pptx", "mSlides.exe"),
        ("mOffice.Slides.odp", ".odp", "mSlides.exe"),
    ];

    /// <summary>Linhas de uma secao (<c>[Setup]</c>, <c>[Run]</c>...), sem comentarios nem linhas vazias.</summary>
    internal static List<string> Secao(string iss, string nome)
    {
        var linhas = new List<string>();
        string? atual = null;
        foreach (var bruta in iss.Split('\n'))
        {
            var l = bruta.Trim();
            var m = Regex.Match(l, @"^\[(\w+)\]$");
            if (m.Success) { atual = m.Groups[1].Value; continue; }
            if (atual != nome || l.Length == 0 || l.StartsWith(';') || l.StartsWith("//")) continue;
            linhas.Add(l);
        }
        return linhas;
    }

    private static string? Diretiva(string nome)
    {
        var achadas = Secao(Iss, "Setup").Where(l => l.StartsWith(nome + "=", StringComparison.Ordinal)).ToList();
        Assert.True(achadas.Count <= 1, $"{nome}= aparece {achadas.Count} vezes no [Setup]");
        return achadas.Count == 0 ? null : achadas[0][(nome.Length + 1)..].Trim();
    }

    private static string Define(string nome)
    {
        var m = Regex.Match(Iss, "^#define " + nome + " \"([^\"]*)\"", RegexOptions.Multiline);
        Assert.True(m.Success, $"#define {nome} ausente");
        return m.Groups[1].Value;
    }

    private static string Codigo(string funcao)
    {
        var code = Iss[Iss.IndexOf("\n[Code]", StringComparison.Ordinal)..];
        var i = code.IndexOf("function " + funcao + "(", StringComparison.Ordinal);
        if (i < 0) i = code.IndexOf("procedure " + funcao + "(", StringComparison.Ordinal);
        Assert.True(i >= 0, $"funcao {funcao} ausente no [Code]");
        var fim = code.IndexOf("\nend;", i, StringComparison.Ordinal);
        return code[i..fim];
    }

    // ---- [Setup] ------------------------------------------------------------------------------------------

    [Fact] // AppMutex: o texto do .iss == a constante do app; vem do [Code] (so no interativo e no desinstalador)
    public void AppMutex_CasaComAConstante_EDependeDoModo()
    {
        Assert.Equal(NomesDeInstancia.MutexDeAtualizacao, Define("MutexDeAtualizacao"));
        Assert.Equal("{code:MutexDoApp}", Diretiva("AppMutex"));
        var f = Codigo("MutexDoApp");
        Assert.Contains("Result := '{#MutexDeAtualizacao}';", f);
        // So o instalador SILENCIOSO fica sem checagem (ali a pergunta suprimida aborta).
        Assert.Matches(@"if WizardSilent then\s+Result := '';", f);
        // O desinstalador usa o AppMutex GRAVADO na instalacao (vazio se foi silenciosa): checa por conta propria.
        var u = Codigo("InitializeUninstall");
        Assert.Contains("while CheckForMutexes('{#MutexDeAtualizacao}') do", u);
        Assert.Contains("Result := False;", u);
        Assert.Equal("yes", Diretiva("CloseApplications"));
        Assert.Equal("no", Diretiva("RestartApplications"));
    }

    [Fact] // AppId presente e FIXO (o teste pina o GUID: regenerar quebra aqui, de proposito)
    public void AppId_FixoENome()
    {
        Assert.Equal("{{" + AppIdFixo + "}", Diretiva("AppId"));
        Assert.Contains("NUNCA regenerar", Iss);
        Assert.Equal("{#MyAppName}", Diretiva("AppName"));
        Assert.Equal("mOffice", Define("MyAppName"));
    }

    [Fact] // diretivas pedidas pelo plano
    public void Setup_Diretivas()
    {
        Assert.Equal("admin", Diretiva("PrivilegesRequired"));
        Assert.Equal("x64compatible", Diretiva("ArchitecturesAllowed"));
        Assert.Equal("yes", Diretiva("DisableWelcomePage"));
        Assert.Equal("yes", Diretiva("DisableReadyPage"));
        Assert.Equal("lzma2/ultra64", Diretiva("Compression"));
        Assert.Equal("yes", Diretiva("SolidCompression"));
        Assert.Equal("no", Diretiva("DiskSpanning"));
        Assert.Equal("yes", Diretiva("ChangesAssociations"));
        Assert.Equal("mOffice-Setup-{#MyAppVersion}", Diretiva("OutputBaseFilename"));
        Assert.Null(Diretiva("Password"));
    }

    [Fact] // termos da Task 4 antes de instalar; sem o arquivo o .iss nao compila e o build para ANTES da suite
    public void InfoBeforeFile_EAvisoDaTask4()
    {
        Assert.Equal("{#ArquivoAviso}", Diretiva("InfoBeforeFile"));
        Assert.Equal(@"..\..\docs\licencas\AVISO-mOffice.txt", Define("ArquivoAviso"));
        Assert.Contains("#if !FileExists(AddBackslash(SourcePath) + ArquivoAviso)", Iss);
        var build = File.ReadAllText(Path.Combine(Raiz, "tools", "installer", "build-installer.ps1"));
        var iAviso = build.IndexOf(@"'docs\licencas\AVISO-mOffice.txt'", StringComparison.Ordinal);
        var iCheca = build.IndexOf("if (-not (Test-Path $Aviso))", StringComparison.Ordinal);
        var iTeste = build.IndexOf("& dotnet test $SlnFile", StringComparison.Ordinal);
        Assert.True(iAviso >= 0 && iCheca > iAviso && iTeste > iCheca, "build-installer.ps1 confere o AVISO antes da suite");
    }

    // ---- [Tasks] / [Icons] --------------------------------------------------------------------------------

    [Fact] // associar e firewall marcadas; atalho na area de trabalho desmarcado; Menu Iniciar sempre
    public void Tarefas_EAtalhos()
    {
        var tarefas = Secao(Iss, "Tasks");
        string Tarefa(string n) => Assert.Single(tarefas, l => l.StartsWith($"Name: \"{n}\";", StringComparison.Ordinal));
        Assert.DoesNotContain("unchecked", Tarefa("associar"));
        Assert.DoesNotContain("unchecked", Tarefa("firewall"));
        Assert.Contains("Flags: unchecked", Tarefa("atalhodesktop"));
        Assert.Equal(3, tarefas.Count);

        var icones = Secao(Iss, "Icons");
        foreach (var exe in new[] { "mTexto", "mPlanilha", "mSlides" })
        {
            Assert.Contains(icones, l => l == $"Name: \"{{autoprograms}}\\{exe}\"; Filename: \"{{app}}\\{exe}.exe\"");
            Assert.Contains(icones, l => l == $"Name: \"{{autodesktop}}\\{exe}\"; Filename: \"{{app}}\\{exe}.exe\"; Tasks: atalhodesktop");
        }
    }

    // ---- [Registry] ---------------------------------------------------------------------------------------

    [Fact] // os 6 ProgIDs, com icone e abrir; OpenWithProgids; Capabilities; RegisteredApplications; SupportedTypes
    public void SeisProgIds_EAplicativoRegistrado()
    {
        var reg = Secao(Iss, "Registry");
        Assert.All(reg, l => Assert.EndsWith("Tasks: associar", l)); // tudo preso a task
        foreach (var (progId, ext, exe) in ProgIds)
        {
            var app = exe[..^4];
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\Classes\\{progId}\"; ValueType: string; ValueName: \"\";") && l.Contains("Flags: uninsdeletekey"));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\Classes\\{progId}\\DefaultIcon\"") && l.Contains($"ValueData: \"\"\"{{app}}\\{exe}\"\",0\""));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\Classes\\{progId}\\shell\\open\\command\"") && l.Contains($"ValueData: \"\"\"{{app}}\\{exe}\"\" \"\"%1\"\"\""));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\Classes\\{ext}\\OpenWithProgids\"; ValueType: string; ValueName: \"{progId}\"; ValueData: \"\"; Flags: uninsdeletevalue;"));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\mOffice\\{app}\\Capabilities\\FileAssociations\"; ValueType: string; ValueName: \"{ext}\"; ValueData: \"{progId}\";"));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\Classes\\Applications\\{exe}\\SupportedTypes\"; ValueType: string; ValueName: \"{ext}\";"));
        }
        foreach (var app in new[] { "mTexto", "mPlanilha", "mSlides" })
        {
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\mOffice\\{app}\"; Flags: uninsdeletekey;"));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\Classes\\Applications\\{app}.exe\"; ValueType: string; ValueName: \"FriendlyAppName\"") && l.Contains("Flags: uninsdeletekey"));
            Assert.Contains(reg, l => l.Contains($"Subkey: \"Software\\RegisteredApplications\"; ValueType: string; ValueName: \"mOffice.{app}\"; ValueData: \"Software\\mOffice\\{app}\\Capabilities\"; Flags: uninsdeletevalue;"));
        }
        // Chaves COMPARTILHADAS nunca levam uninsdeletekey (desinstalar apagaria o registro de outros programas).
        Assert.DoesNotContain(reg, l => (l.Contains("\\OpenWithProgids\"") || l.Contains("\"Software\\RegisteredApplications\"")) && l.Contains("uninsdeletekey"));
        // Exatamente os 6 ProgIDs, nenhum a mais.
        var progIdsNoIss = reg.Select(l => Regex.Match(l, @"Subkey: ""Software\\Classes\\(mOffice\.[^""\\]+)""")).Where(m => m.Success).Select(m => m.Groups[1].Value).Distinct().Order();
        Assert.Equal(ProgIds.Select(p => p.ProgId).Order(), progIdsNoIss);
    }

    [Fact] // mesma lista de ProgIDs no registro de desenvolvimento, no XML e no script do RMM
    public void ProgIds_BatemComDevERmm()
    {
        var dev = File.ReadAllText(Path.Combine(Raiz, "tools", "explorer", "registrar-dev.ps1"));
        var xml = XDocument.Load(Path.Combine(Raiz, "tools", "rmm", "moffice-padrao-office.xml"));
        var aplicar = File.ReadAllText(Path.Combine(Raiz, "tools", "rmm", "aplicar-padrao-office.ps1"));
        var noXml = xml.Root!.Elements("Association").Select(a => ((string)a.Attribute("Identifier")!, (string)a.Attribute("ProgId")!, (string)a.Attribute("ApplicationName")!)).ToList();
        Assert.Equal(6, noXml.Count);
        foreach (var (progId, ext, exe) in ProgIds)
        {
            Assert.Matches($@"ProgId = '{Regex.Escape(progId)}';\s+Ext = '{Regex.Escape(ext)}';\s+Projeto = '{exe[..^4]}'", dev);
            Assert.Contains((ext, progId, exe[..^4]), noXml);
            Assert.Matches($@"Ext = '{Regex.Escape(ext)}';\s+ProgId = '{Regex.Escape(progId)}';\s+App = '{exe[..^4]}'", aplicar);
        }
    }

    // ---- [Files] / [InstallDelete] / [Dirs] ---------------------------------------------------------------

    [Fact] // o pacote inteiro, sem .pdb e sem sdkjs-plugins (defesa em profundidade sobre o montar-pacote.ps1)
    public void Arquivos_SemPdbESemSdkjsPlugins()
    {
        var files = Secao(Iss, "Files");
        var pacote = Assert.Single(files);
        Assert.StartsWith("Source: \"{#PacoteDir}\\*\";", pacote);
        var excludes = Regex.Match(pacote, "Excludes: \"([^\"]*)\"").Groups[1].Value.Split(',');
        Assert.Contains("*.pdb", excludes);
        Assert.Contains("sdkjs-plugins", excludes);
        Assert.Contains("AllFonts.js", excludes);
        Assert.Contains("ASC.ttf", excludes);
        Assert.Contains("DestDir: \"{app}\"", pacote);
        Assert.Contains("recursesubdirs", pacote);
    }

    [Fact] // atualizacao por cima troca o motor e o runtime inteiros; o catalogo por maquina e recriado. Na
           // desinstalacao nada de [UninstallDelete]/[UninstallRun]: o catalogo sai pelo [Code], depois da checagem de link
    public void InstallDelete_MotorRuntimeECatalogo()
    {
        var del = Secao(Iss, "InstallDelete");
        Assert.Contains("Type: filesandordirs; Name: \"{app}\\motor\"", del);
        Assert.Contains("Type: filesandordirs; Name: \"{app}\\WebView2Runtime\"", del);
        Assert.Contains("Type: filesandordirs; Name: \"{commonappdata}\\mOffice\\fontes\"", del);
        Assert.Contains("Name: \"{commonappdata}\\mOffice\\fontes\"; Flags: uninsalwaysuninstall", Secao(Iss, "Dirs"));
        Assert.Empty(Secao(Iss, "UninstallDelete"));
        Assert.Empty(Secao(Iss, "UninstallRun"));
        Assert.DoesNotMatch(@"(?m)^\[Uninstall(Delete|Run)\]", Iss);
        var u = Codigo("CurUninstallStepChanged");
        var iNossa = u.IndexOf("if PastaEhDaInstalacao() and RemoverLinks() then", StringComparison.Ordinal);
        var iDel = u.IndexOf("DelTree(PastaDoMOffice() + '\\fontes', True, True, True);", StringComparison.Ordinal);
        Assert.True(iNossa >= 0 && iDel > iNossa, "o catalogo so e apagado na pasta DESTA instalacao, sem link");
        // a marca de identidade so sai quando a pasta saiu (senao a proxima instalacao renomearia a pasta com o XML do RMM)
        var post = u[u.IndexOf("else if CurUninstallStep = usPostUninstall then", StringComparison.Ordinal)..];
        Assert.Contains("if not DirExists(PastaDoMOffice()) then", post);
        Assert.Contains("RegDeleteKeyIncludingSubkeys(ChaveRaizDaMarca(), ChaveInstalacao);", post);
    }

    // ---- Pasta de dados {commonappdata}\mOffice: renomeia a de outra origem, ACL na pasta VAZIA, sem /T ----------

    [Fact] // ssInstall (antes do [InstallDelete]/[Dirs]) prepara a pasta ou aborta
    public void PastaDeDados_PreparadaNoSsInstallOuAborta()
    {
        var passo = Codigo("CurStepChanged");
        var ssInstall = passo[passo.IndexOf("if CurStep = ssInstall then", StringComparison.Ordinal)..passo.IndexOf("else if CurStep = ssPostInstall then", StringComparison.Ordinal)];
        Assert.Contains("if not PrepararPastaDeDados() then", ssInstall);
        Assert.Contains("SuppressibleMsgBox(", ssInstall);
        Assert.Matches(@"(?m)^\s+Abort;", ssInstall); // Abort em CurStepChanged(ssInstall) encerra o Setup (ajuda do Inno)
        Assert.Equal("yes", Diretiva("RedirectionGuard"));
    }

    [Fact] // pasta de outra origem e RENOMEADA antes de qualquer ACL e nunca percorrida; a nova recebe a ACL vazia
    public void PastaDeDados_RenomeiaAntesDaAclENaoPercorre()
    {
        var f = Codigo("PrepararPastaDeDados");
        int Em(string trecho)
        {
            var i = f.IndexOf(trecho, StringComparison.Ordinal);
            Assert.True(i >= 0, "ausente: " + trecho);
            return i;
        }
        var iLink = Em("if EhLink(Raiz) then");
        var iNossa = Em("if DirExists(Raiz) and PastaEhDaInstalacao() then");
        var iRenomeia = Em("if not RenameFile(Raiz, Antiga) then");
        var iCria = Em("if not CreateDir(Nova) then");
        var iAclNova = Em("if not AplicarAcl(Nova) then");
        var iMove = Em("if not RenameFile(Nova, Raiz) then");
        var iMarca = Em("RegWriteStringValue(ChaveRaizDaMarca(), ChaveInstalacao, ValorIdentidade, Id)");
        Assert.True(iLink < iNossa && iNossa < iRenomeia && iRenomeia < iCria && iCria < iAclNova && iAclNova < iMove && iMove < iMarca,
            "link -> pasta da instalacao? -> renomeia a outra -> cria vazia em {tmp} -> ACL -> move -> marca");
        // Antes do rename, a UNICA ACL e a do ramo "pasta desta instalacao" (que sai com Exit): nenhuma ACL numa pasta
        // de outra origem.
        var antes = f[..iRenomeia];
        Assert.Single(Regex.Matches(antes, @"AplicarAcl\("));
        Assert.Matches(@"Result := AplicarAcl\(Raiz\);\s+Exit;", antes[iNossa..]);
        Assert.DoesNotContain("Exec(", antes);
        Assert.Contains("Nova := ExpandConstant('{tmp}\\mOffice-dados');", f);
        // A antiga so aparece no nome, no rename e no log: nada a percorre, copia ou apaga com direitos de admin.
        foreach (var l in f.Split('\n').Where(l => l.Contains("Antiga")))
            Assert.Matches(@"Antiga, Nova, Id: String;|Antiga := NomeDaAntiga\(Raiz\);|RenameFile\(Raiz, Antiga\)|Log\(|\+ Antiga \+", l);
    }

    [Fact] // nome da antiga: data + sufixo ALEATORIO, e uma segunda tentativa com outro nome antes de desistir
    public void PastaDeDados_NomeDaAntigaAleatorioComNovaTentativa()
    {
        Assert.Contains("Result := Raiz + '.antigo-' + GetDateTimeString('yyyymmddhhnnss', #0, #0) + '-' + IntToStr(Random(", Codigo("NomeDaAntiga"));
        var f = Codigo("PrepararPastaDeDados");
        Assert.Equal(2, Regex.Matches(f, @"Antiga := NomeDaAntiga\(Raiz\);").Count);
        Assert.Equal(2, Regex.Matches(f, @"if not RenameFile\(Raiz, Antiga\) then").Count);
        var i1 = f.IndexOf("Antiga := NomeDaAntiga(Raiz);", StringComparison.Ordinal);
        var i2 = f.IndexOf("Antiga := NomeDaAntiga(Raiz);", i1 + 1, StringComparison.Ordinal);
        Assert.True(f.IndexOf("if not RenameFile(Raiz, Antiga) then", StringComparison.Ordinal) < i2, "a segunda tentativa so depois de a primeira falhar");
    }

    [Fact] // move de {tmp} falhou (outro volume): FALHA FECHADO - nunca cria a pasta direto em %ProgramData%
    public void PastaDeDados_MoveFalhouFalhaFechado()
    {
        var f = Codigo("PrepararPastaDeDados");
        var iMove = f.IndexOf("if not RenameFile(Nova, Raiz) then", StringComparison.Ordinal);
        Assert.True(iMove >= 0);
        var ramo = f[iMove..f.IndexOf("end;", iMove, StringComparison.Ordinal)];
        Assert.Contains("Log('mOffice: FALHA ao mover", ramo);
        Assert.Contains("Exit;", ramo);
        Assert.DoesNotContain("CreateDir(Raiz)", f); // em lugar nenhum: a pasta so chega por move, ja protegida
        Assert.DoesNotContain("PastaVazia", Iss);
        // e o Exit vira aborto com mensagem pt-BR no ssInstall
        var passo = Codigo("CurStepChanged");
        Assert.Contains("if not PrepararPastaDeDados() then", passo);
        Assert.Contains("Não foi possível preparar com segurança a pasta", passo);
        // Identidade = volume + indice NTFS da PROPRIA entrada (sem seguir link), so para pasta
        var id = Codigo("IdentidadeDaPasta");
        Assert.Contains("CreateFileW(Caminho, $80, 7, 0, 3, $02000000 or $00200000, 0)", id); // BACKUP_SEMANTICS + OPEN_REPARSE_POINT
        Assert.Contains("(Info.Atributos and FILE_ATTRIBUTE_REPARSE_POINT) = 0", id);
        Assert.Contains("IntToStr(Info.Volume) + '-' + IntToStr(Info.IndiceHi) + '-' + IntToStr(Info.IndiceLo)", id);
        Assert.Contains("(Marca = Id)", Codigo("PastaEhDaInstalacao"));
        Assert.Contains("Result := HKLM;", Codigo("ChaveRaizDaMarca")); // a marca so pode ser escrita por admin
    }

    [Fact] // ACL final: dono Administradores, heranca cortada, Usuarios so RX, /L - e NENHUM /T em lugar nenhum
    public void PastaDeDados_AclSemT()
    {
        var acl = Codigo("AplicarAcl");
        Assert.Contains("'\" /setowner *S-1-5-32-544 /L'", acl);
        Assert.Contains("'\" /inheritance:r /grant:r *S-1-5-18:(OI)(CI)F *S-1-5-32-544:(OI)(CI)F *S-1-5-32-545:(OI)(CI)RX /L'", acl);
        Assert.Contains("Result := (Dono = 0) and (Acl = 0);", acl);
        var codigo = Iss.Split('\n').Select(l => l.Trim()).Where(l => !l.StartsWith(';') && !l.StartsWith("//")).ToList();
        var icacls = codigo.Where(l => l.Contains("icacls", StringComparison.OrdinalIgnoreCase)).ToList();
        Assert.Equal(2, icacls.Count); // so as duas do AplicarAcl; nenhuma no [Run]
        Assert.All(icacls, l => Assert.DoesNotMatch(@"(?i)/T\b", l));
        Assert.All(icacls, l => Assert.Contains(" /L'", l));
        Assert.DoesNotContain(icacls, l => Regex.IsMatch(l, @"S-1-5-32-545:\([^)]*\)\([^)]*\)(F|M|W)"));
        Assert.DoesNotContain(Secao(Iss, "Run"), l => l.Contains("icacls", StringComparison.OrdinalIgnoreCase));
    }

    [Fact] // links: a varredura (sem entrar em link) continua como defesa em profundidade antes do catalogo e na desinstalacao
    public void Links_VarreduraComoDefesa()
    {
        var eh = Codigo("EhLink");
        Assert.Contains("if FindFirst(Caminho, R) then", eh); // sem curinga: a PROPRIA entrada, sem seguir o link
        Assert.Contains("(R.Attributes and FILE_ATTRIBUTE_REPARSE_POINT) <> 0", eh);
        var apaga = Codigo("ApagarLink");
        Assert.Contains("if not RemoveDir(Caminho) then", apaga);
        Assert.Contains("DeleteFile(Caminho);", apaga);
        Assert.Contains("Result := not EhLink(Caminho);", apaga);
        var arvore = Codigo("LinksRestantes");
        Assert.Contains("if not ApagarLink(Filho) then", arvore);
        Assert.Matches(@"else if \(R\.Attributes and FILE_ATTRIBUTE_DIRECTORY\) <> 0 then\s+Result := Result \+ LinksRestantes\(Filho\);", arvore);
        Assert.Contains("Result := RemoverLinks();", Codigo("SemLinks"));
    }

    [Fact] // backlog 20: o catalogo e gerado como admin DEPOIS de a pasta estar protegida, com a varredura logo antes
    public void CatalogoPorMaquina_GeradoNaPastaProtegida()
    {
        var run = Secao(Iss, "Run");
        var gera = Assert.Single(run, l => l.StartsWith("Filename: \"{app}\\mTexto.exe\"; Parameters: \"" + VerboCatalogo.Verbo + "\"", StringComparison.Ordinal));
        Assert.Contains("waituntilterminated", gera);
        Assert.EndsWith("Check: SemLinks", gera);
        Assert.DoesNotContain("Tasks:", gera); // sempre, sem tarefa
        // O XML da politica do RMM mora na mesma pasta: o aplicar recria o arquivo (nunca herda dono de outro).
        var aplicar = File.ReadAllText(Path.Combine(Raiz, "tools", "rmm", "aplicar-padrao-office.ps1"));
        Assert.True(aplicar.IndexOf("Remove-Item -LiteralPath $xmlPath", StringComparison.Ordinal) < aplicar.IndexOf("[System.Xml.XmlWriter]::Create($xmlPath", StringComparison.Ordinal));
    }

    // ---- Firewall (backlog 22) ------------------------------------------------------------------------------

    /// <summary>O prefixo das regras no tools/egresso/firewall.ps1 (a fonte que o FirewallTests e o -Remover usam).</summary>
    private static string PrefixoDoScript()
    {
        var script = File.ReadAllText(Path.Combine(Raiz, "tools", "egresso", "firewall.ps1"));
        var m = Regex.Match(script, @"\[string\]\$Nome = ""([^""]+)""");
        Assert.True(m.Success);
        Assert.Contains("$nomeRegra = \"$Nome ($($exe.Name))\"", script); // formato "<prefixo> (<exe>)"
        return m.Groups[1].Value;
    }

    [Fact] // mesmos nomes do script de desenvolvimento: "<prefixo> (<exe>)", o prefixo IGUAL ao do firewall.ps1
    public void Firewall_NomesIguaisAosDoScript()
    {
        Assert.Contains($"PrefixoRegra = '{PrefixoDoScript()}';", Iss);
        Assert.Contains("Result := PrefixoRegra + ' (' + ExtractFileName(Exe) + ')';", Codigo("NomeDaRegra"));
        // a lista de .exe e a do script: todos, recursivo, sem entrar em link
        var coleta = Codigo("ColetarExes");
        Assert.Contains("if (R.Attributes and FILE_ATTRIBUTE_REPARSE_POINT) = 0 then", coleta);
        Assert.Contains("ColetarExes(Pasta + '\\' + R.Name, Lista);", coleta);
        Assert.Contains("CompareText(ExtractFileExt(R.Name), '.exe') = 0", coleta);
        // leitura das regras existentes com os MESMOS rotulos do script (ingles e portugues)
        var existentes = Codigo("RegrasExistentes");
        Assert.Contains("'Rule Name'", existentes);
        Assert.Contains("'Nome da Regra'", existentes);
        Assert.Contains("if Pos(PrefixoRegra, L) = 1 then", existentes);
        Assert.Contains("'^(Rule Name|Nome da Regra):", File.ReadAllText(Path.Combine(Raiz, "tools", "egresso", "firewall.ps1")));
    }

    [Fact] // criada pelo netsh (sem PowerShell), com o codigo de CADA regra conferido; falha vai ao log e ao aviso
    public void Firewall_NetshComCodigoConferidoEAviso()
    {
        Assert.Equal("yes", Diretiva("SetupLogging"));
        Assert.Equal("yes", Diretiva("UninstallLogging"));
        var semComentario = Iss.Split('\n').Select(l => l.Trim()).Where(l => !l.StartsWith(';') && !l.StartsWith("//"));
        Assert.DoesNotContain(semComentario, l => l.Contains("powershell", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("firewall.ps1\"", Iss); // o script nao vai mais para {app}\tools

        Assert.Contains("Exec(ExpandConstant('{sys}\\netsh.exe'), Parametros, '', SW_HIDE, ewWaitUntilTerminated, Result)", Codigo("Netsh"));
        var cria = Codigo("CriarRegrasFirewall");
        Assert.Contains("'advfirewall firewall add rule name=\"' + NomeDaRegra(Exes[i]) + '\" dir=out action=block program=\"' +", cria);
        Assert.Contains("'\" enable=yes profile=any'", cria);
        Assert.Contains("if Codigo = 0 then", cria);
        Assert.Contains("Log('mOffice: FALHA ao criar a regra de firewall ' + NomeDaRegra(Exes[i])", cria);
        Assert.Contains("Result := (Falhas = 0) and (GetArrayLength(Exes) > 0);", cria); // zero .exe tambem e falha

        var passo = Codigo("CurStepChanged");
        var pos = passo[passo.IndexOf("else if CurStep = ssPostInstall then", StringComparison.Ordinal)..];
        var iRemove = pos.IndexOf("RemoverRegrasFirewall();", StringComparison.Ordinal);
        var iCria = pos.IndexOf("if WizardIsTaskSelected('firewall') then", StringComparison.Ordinal);
        Assert.True(iRemove >= 0 && iCria > iRemove, "sempre remove as antigas; cria so com a tarefa");
        Assert.Contains("if not CriarRegrasFirewall() then\n        SuppressibleMsgBox(AvisoFirewall, mbError, MB_OK, IDOK);", pos.Replace("\r", ""));
        Assert.Contains("Não foi possível criar no Firewall do Windows", Iss); // aviso em pt-BR
    }

    [Fact] // desinstalar remove as regras pelo netsh (os mesmos nomes), antes de os arquivos sairem
    public void Firewall_RemocaoNaDesinstalacao()
    {
        var remove = Codigo("RemoverRegrasFirewall");
        Assert.Contains("Nomes := RegrasExistentes();", remove);
        Assert.Contains("Acrescentar(Nomes, NomeDaRegra(Exes[i]));", remove); // idioma sem rotulo conhecido: pelos .exe
        Assert.Contains("Netsh('advfirewall firewall delete rule name=\"' + Nomes[i] + '\"');", remove);
        Assert.Contains("Result := GetArrayLength(RegrasExistentes());", remove); // confere de verdade depois
        var u = Codigo("CurUninstallStepChanged");
        Assert.Contains("if CurUninstallStep = usUninstall then", u);
        Assert.Contains("if RemoverRegrasFirewall() > 0 then", u);
    }

    [Fact] // desinstalar desfaz a politica do RMM so se ela aponta para o XML DO mOffice, tirando so as seis
    public void Rmm_DesinstalacaoRemoveSoAsNossas()
    {
        var p = Codigo("RemoverPoliticaDoRmm");
        Assert.Contains("Chave := 'SOFTWARE\\Policies\\Microsoft\\Windows\\System';", p);
        Assert.Contains("Xml := PastaDoMOffice() + '\\moffice-padrao-office.xml';", p);
        Assert.Contains("if CompareText(Valor, Xml) <> 0 then", p);
        Assert.Contains("RegDeleteValue(HKLM, Chave, 'DefaultAssociationsConfiguration');", p);
        Assert.Contains("SaveStringsToUTF8FileWithoutBOM(Xml, Resto, False);", p);
        // XML que existe mas e link ou nao pode ser lido: registra e SAI antes de apagar o valor da politica.
        var iIlegivel = p.IndexOf("if EhLink(Xml) or (not LoadStringsFromFile(Xml, Linhas)) then", StringComparison.Ordinal);
        Assert.True(iIlegivel >= 0);
        var ramo = p[iIlegivel..p.IndexOf("end;", iIlegivel, StringComparison.Ordinal)];
        Assert.Contains("NAO alterada", ramo);
        Assert.Contains("Exit;", ramo);
        Assert.True(iIlegivel < p.IndexOf("RegDeleteValue(", StringComparison.Ordinal));
        Assert.Contains("if FileExists(Xml) or EhLink(Xml) then", p);
        // as seis extensoes, iguais as do XML de referencia do RMM
        var nossa = Codigo("AssociacaoDoMOffice");
        foreach (var (_, ext, _) in ProgIds) Assert.Contains($"Pos('Identifier=\"{ext}\"', L) > 0", nossa);
        Assert.Equal(6, Regex.Matches(nossa, "Identifier=").Count);
        // mesmo nome de arquivo que o aplicar grava
        Assert.Contains("$xmlPath = Join-Path $Pasta 'moffice-padrao-office.xml'", File.ReadAllText(Path.Combine(Raiz, "tools", "rmm", "aplicar-padrao-office.ps1")));
        var u = Codigo("CurUninstallStepChanged");
        var iLinks = u.IndexOf("if PastaEhDaInstalacao() and RemoverLinks() then", StringComparison.Ordinal);
        Assert.True(iLinks >= 0 && u.IndexOf("RemoverPoliticaDoRmm();", StringComparison.Ordinal) > iLinks);
    }

    [Fact] // reabre SO o app que pediu (/REABRIR=), nunca com /NORELAUNCH; des-elevado; sem skipifsilent
    public void DeveReabrir_RelancaSoOAppQuePediu()
    {
        var run = Secao(Iss, "Run");
        foreach (var app in new[] { "mTexto", "mPlanilha", "mSlides" })
        {
            var l = Assert.Single(run, x => x.Contains($"Check: DeveReabrir('{app}')"));
            Assert.StartsWith($"Filename: \"{{app}}\\{app}.exe\";", l);
            Assert.Contains("runasoriginaluser", l);
            Assert.Contains("nowait", l);
            Assert.Contains("postinstall", l);
            Assert.DoesNotContain("skipifsilent", l); // senao o /VERYSILENT pularia o relaunch
        }
        var f = Codigo("DeveReabrir");
        Assert.Contains($"if TemParametro('{InstaladorSilencioso.ParametroNaoReabrir}') then", f);
        Assert.Contains($"ValorDoParametro('{InstaladorSilencioso.ParametroReabrir}')", f);
        Assert.Contains("Result := False;", f);
    }

    [Fact] // nenhum caminho absoluto de maquina, nem nada do caminho de build (tokens derivados em runtime, como a guarda)
    public void SemCaminhoAbsolutoDeMaquina()
    {
        var arquivos = new[]
        {
            CaminhoIss,
            Path.Combine(Raiz, "tools", "installer", "build-installer.ps1"),
            Path.Combine(Raiz, "tools", "rmm", "aplicar-padrao-office.ps1"),
            Path.Combine(Raiz, "tools", "rmm", "remover-padrao-office.ps1"),
            Path.Combine(Raiz, "tools", "rmm", "moffice-padrao-office.xml"),
        };
        var segmentos = Raiz.Split(['\\', '/'], StringSplitOptions.RemoveEmptyEntries);
        // Mesma regra da guarda do tools/export-public.ps1: nomes genericos do Windows (Users, AppData, Temp...) nao
        // identificam ninguem e apareceriam em qualquer texto (ex.: um clone em %TEMP%). ExportPublicoTests confere
        // que a lista do script e' esta.
        var sensiveis = (segmentos.Length > 2 ? segmentos[1..^1] : [])
            .Where(s => s.Length >= 3 && !ExportPublicoTests.SegmentosGenericos.Contains(s, StringComparer.OrdinalIgnoreCase)).ToArray();
        foreach (var a in arquivos)
        {
            var texto = File.ReadAllText(a);
            Assert.DoesNotContain(Raiz, texto, StringComparison.OrdinalIgnoreCase);
            foreach (var s in sensiveis)
                Assert.DoesNotMatch(new Regex("(?<![A-Za-z0-9])" + Regex.Escape(s) + "(?![A-Za-z0-9])", RegexOptions.IgnoreCase), texto);
        }
        // No .iss, fora de comentario, nenhum "X:\" (tudo e relativo ao script ou constante do Inno).
        var codigo = Iss.Split('\n').Select(l => l.Trim()).Where(l => !l.StartsWith(';') && !l.StartsWith("//"));
        Assert.DoesNotContain(codigo, l => Regex.IsMatch(l, @"(?<![A-Za-z])[A-Za-z]:\\"));
        Assert.DoesNotContain(codigo, l => l.Contains(@"\Users\", StringComparison.OrdinalIgnoreCase));
    }

    [Fact] // guarda de vazamento do build: tokens do repo E do usuario que gera o build, lidos do ambiente (nunca fixos)
    public void Build_GuardaIncluiUsuarioEPerfil()
    {
        var b = File.ReadAllText(Path.Combine(Raiz, "tools", "installer", "build-installer.ps1"));
        Assert.Contains(@"$caminhosLiteral += @($env:USERPROFILE, ($env:USERPROFILE -replace '\\', '/'))", b);
        Assert.Contains("$segmentosSensiveis = @($segmentosSensiveis) + $env:USERNAME", b);
        var iTokens = b.IndexOf("$env:USERNAME", StringComparison.Ordinal);
        Assert.True(iTokens > 0 && iTokens < b.IndexOf("foreach ($bin in $arquivosBin)", StringComparison.Ordinal), "tokens antes da varredura");
        Assert.DoesNotContain(Environment.UserName, b.Replace("$env:USERNAME", ""), StringComparison.OrdinalIgnoreCase);
    }

    // ---- Compilacao real pelo ISCC (sem gerar instalador) -------------------------------------------------

    private static string? Iscc()
    {
        var candidatos = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Inno Setup 6", "ISCC.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Inno Setup 6", "ISCC.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Inno Setup 6", "ISCC.exe"),
        };
        return candidatos.FirstOrDefault(File.Exists);
    }

    private static (int Codigo, string Saida) Compilar(string iss, string pacote)
    {
        var psi = new ProcessStartInfo(Iscc()!) { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
        psi.ArgumentList.Add("/O-"); // so compila: nao grava instalador
        psi.ArgumentList.Add("/DMyAppVersion=9.9.9");
        psi.ArgumentList.Add("/DPacoteDir=" + pacote);
        psi.ArgumentList.Add(iss);
        using var p = Process.Start(psi)!;
        var err = p.StandardError.ReadToEndAsync();
        var saida = p.StandardOutput.ReadToEnd();
        Assert.True(p.WaitForExit(120_000), "ISCC nao terminou em 2 min");
        return (p.ExitCode, saida + err.Result);
    }

    [SkippableFact] // o .iss REAL: sem o AVISO da Task 4 recusa com a mensagem; com ele, compila
    public void Iscc_IssReal_SemAvisoRecusaComMensagem()
    {
        Skip.If(Iscc() is null, "Inno Setup 6 (ISCC.exe) nao instalado");
        var pacote = Directory.CreateTempSubdirectory("moffice-iss-").FullName;
        try
        {
            File.WriteAllText(Path.Combine(pacote, "mTexto.exe"), "x");
            var (codigo, saida) = Compilar(CaminhoIss, pacote);
            if (File.Exists(Path.Combine(Raiz, "docs", "licencas", "AVISO-mOffice.txt")))
                Assert.True(codigo == 0, saida);
            else
            {
                Assert.NotEqual(0, codigo);
                Assert.Contains(@"docs\licencas\AVISO-mOffice.txt nao existe", saida);
            }
        }
        finally { Directory.Delete(pacote, true); }
    }

    [SkippableFact] // o .iss inteiro ([Setup], [Registry], [Run], [Code]) compila: copia com o AVISO trocado por um
                    // arquivo de TESTE temporario (nunca no repositorio) e os caminhos relativos resolvidos
    public void Iscc_IssCompila()
    {
        Skip.If(Iscc() is null, "Inno Setup 6 (ISCC.exe) nao instalado");
        var tmp = Directory.CreateTempSubdirectory("moffice-iss-").FullName;
        try
        {
            var pacote = Path.Combine(tmp, "pacote");
            Directory.CreateDirectory(pacote);
            File.WriteAllText(Path.Combine(pacote, "mTexto.exe"), "x");
            var aviso = Path.Combine(tmp, "aviso-teste.txt");
            File.WriteAllText(aviso, "texto de teste");
            var linhaAviso = "#define ArquivoAviso \"..\\..\\docs\\licencas\\AVISO-mOffice.txt\"";
            Assert.Contains(linhaAviso, Iss);
            var copia = Iss
                .Replace(linhaAviso, $"#define ArquivoAviso \"{aviso}\"")
                .Replace("#if !FileExists(AddBackslash(SourcePath) + ArquivoAviso)", "#if !FileExists(ArquivoAviso)")
                .Replace("..\\..\\", Raiz + "\\")
                .Replace("..\\egresso\\", Path.Combine(Raiz, "tools", "egresso") + "\\");
            var iss = Path.Combine(tmp, "moffice.iss");
            File.WriteAllText(iss, copia);
            var (codigo, saida) = Compilar(iss, pacote);
            Assert.True(codigo == 0, saida);
        }
        finally { Directory.Delete(tmp, true); }
    }

    // ---- Verbo /gerarcatalogo -----------------------------------------------------------------------------

    [Theory]
    [InlineData(new[] { "/gerarcatalogo" }, true)]
    [InlineData(new[] { "/GerarCatalogo" }, true)]
    [InlineData(new[] { " /gerarcatalogo " }, true)]
    [InlineData(new[] { "arquivo.docx", "/gerarcatalogo" }, false)] // so o primeiro argumento conta
    [InlineData(new[] { "/gerarcatalogo2" }, false)]
    [InlineData(new[] { "--gerar-catalogo" }, false)]
    [InlineData(new string[0], false)]
    public void Verbo_Reconhecido(string[] args, bool esperado) => Assert.Equal(esperado, VerboCatalogo.EhVerbo(args));

    [Fact] // o verbo nao colide com os da atualizacao nem vira arquivo a abrir
    public void Verbo_NaoColide()
    {
        Assert.Null(ExecutorDeAtualizacao.VerboDe([VerboCatalogo.Verbo]));
        Assert.Null(AplicativoOffice.PrimeiroArquivoAceito([VerboCatalogo.Verbo], PerfilApp.Texto));
    }

    [Fact] // gerador ausente: linha de erro ASCII (a mensagem pt-BR tem acento) e codigo 20, sem excecao
    public async Task Verbo_GeradorAusente_ErroAscii()
    {
        var tmp = Directory.CreateTempSubdirectory("moffice-verbo-").FullName;
        try
        {
            var (linha, codigo) = await VerboCatalogo.ExecutarAsync(() => new CatalogoFontes(Path.Combine(tmp, "cat"), Path.Combine(tmp, "nao-existe.exe"), [tmp]));
            Assert.Equal(VerboCatalogo.CodigoErro, codigo);
            Assert.StartsWith("mOffice /gerarcatalogo: erro=O gerador de catalogo de fontes nao foi encontrado", linha);
            Assert.All(linha, c => Assert.InRange(c, ' ', '~'));
        }
        finally { Directory.Delete(tmp, true); }
    }

    [Fact] // motor nao encontrado (Localizar lanca) tambem vira linha de erro, nunca excecao
    public async Task Verbo_Excecao_ViraErro()
    {
        var (linha, codigo) = await VerboCatalogo.ExecutarAsync(() => throw new InvalidOperationException("Motor não encontrado"));
        Assert.Equal(20, codigo);
        Assert.Equal("mOffice /gerarcatalogo: erro=Motor nao encontrado", linha);
    }

    [SkippableFact] // com o allfontsgen de verdade: gera (regenerou=sim) e, na segunda vez, ja esta atual (regenerou=nao)
    public async Task Verbo_GeraECatalogoFicaAtual()
    {
        var motor = mOffice.Conversor.Tests.MotorDeTeste.Localizar().PastaMotor;
        var gerador = Path.Combine(motor, "converter", "allfontsgen.exe");
        Skip.If(!File.Exists(gerador), "allfontsgen.exe ausente - rode tools/onlyoffice/obter-allfontsgen.ps1");
        var tmp = Directory.CreateTempSubdirectory("moffice-verbo-").FullName;
        try
        {
            CatalogoFontes Criar() => new(Path.Combine(tmp, "cat"), gerador, [Path.Combine(motor, "fonts")]);
            var (linha1, codigo1) = await VerboCatalogo.ExecutarAsync(Criar);
            Assert.Equal(0, codigo1);
            Assert.Matches(@"^mOffice /gerarcatalogo: pasta=.+ fontes=[1-9]\d* regenerou=sim$", linha1);
            Assert.True(File.Exists(Path.Combine(tmp, "cat", "AllFonts.js")));
            Assert.True(File.Exists(Path.Combine(tmp, "cat", "impressao-digital.txt"))); // o que o allfontsgen direto NAO grava
            var (linha2, codigo2) = await VerboCatalogo.ExecutarAsync(Criar);
            Assert.Equal(0, codigo2);
            Assert.EndsWith("regenerou=nao", linha2);
        }
        finally { Directory.Delete(tmp, true); }
    }

    [Fact] // roteado na subida comum ANTES da instancia unica e das pastas; producao usa CatalogoFontes.Padrao
    public void Verbo_RoteadoAntesDeTudo()
    {
        var f = File.ReadAllText(Path.Combine(Raiz, "src", "mOffice.Comum", "AplicativoOffice.cs"));
        var s = f[f.IndexOf("protected override void OnStartup(", StringComparison.Ordinal)..];
        var iVerbo = s.IndexOf("VerboCatalogo.EhVerbo(e.Args)", StringComparison.Ordinal);
        var iInstancia = s.IndexOf("new InstanciaUnica(", StringComparison.Ordinal);
        var iPastas = s.IndexOf("Pastas.Garantir(", StringComparison.Ordinal);
        Assert.True(iVerbo >= 0 && iInstancia > iVerbo && iPastas > iVerbo, "o verbo tem de vir antes das pastas e da instancia unica");
        var regiao = s[iVerbo..iPastas];
        Assert.Contains("Task.Run(() => VerboCatalogo.ExecutarAsync(() => CatalogoFontes.Padrao(ConfiguracaoMotor.Localizar())))", regiao);
        Assert.Contains("SaidaDeStatus.Escrever(linha);", regiao);
        Assert.Contains("Shutdown(codigo);", regiao);
        Assert.Contains("return;", regiao[regiao.IndexOf("Shutdown(codigo);", StringComparison.Ordinal)..]);
    }

    // ---- /REABRIR= ----------------------------------------------------------------------------------------

    [Fact] // o botao passa o proprio app; o RMM (/NORELAUNCH) nunca passa /REABRIR
    public void StartInfo_Reabrir()
    {
        Assert.Equal("/REABRIR=", InstaladorSilencioso.ParametroReabrir);
        Assert.Equal("/SILENT /SUPPRESSMSGBOXES /NORESTART /REABRIR=mPlanilha",
            InstaladorSilencioso.MontarStartInfo(@"C:\Temp\s.exe", reabrir: true, app: "mPlanilha").Arguments);
        Assert.Equal("/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /NORELAUNCH",
            InstaladorSilencioso.MontarStartInfo(@"C:\Temp\s.exe", reabrir: false, app: "mPlanilha").Arguments);
        var app = File.ReadAllText(Path.Combine(Raiz, "src", "mOffice.Comum", "AplicativoOffice.cs"));
        Assert.Contains("InstaladorSilencioso.MontarStartInfo(a, app: _perfil.Nome)", app);
        // O /REABRIR= leva PerfilApp.Nome ("mtexto"); o DeveReabrir('mTexto') do .iss compara com CompareText (sem caixa).
        Assert.Equal("mTexto", PerfilApp.Texto.Nome, ignoreCase: true);
        Assert.Equal("mPlanilha", PerfilApp.Planilha.Nome, ignoreCase: true);
        Assert.Equal("mSlides", PerfilApp.Slides.Nome, ignoreCase: true);
        Assert.Contains("Result := CompareText(ValorDoParametro('/REABRIR='), App) = 0;", Codigo("DeveReabrir"));
    }

    // ---- Politica de associacao do RMM (contra chaves de TESTE em HKCU e pasta temporaria) -----------------

    private sealed class AmbienteRmm : IDisposable
    {
        public readonly string Chave = @"Software\mOffice.Testes\Rmm-" + Guid.NewGuid().ToString("N");
        public readonly string Pasta = Directory.CreateTempSubdirectory("moffice-rmm-").FullName;
        public string Politica => Chave + @"\Politica";
        public string Classes => Chave + @"\Classes";
        public string Xml => Path.Combine(Pasta, "moffice-padrao-office.xml");

        public void RegistrarProgIds()
        {
            foreach (var p in ProgIds) Registry.CurrentUser.CreateSubKey(Classes + "\\" + p.ProgId).Dispose();
        }

        public string? ValorPolitica()
        {
            using var k = Registry.CurrentUser.OpenSubKey(Politica);
            return k?.GetValue("DefaultAssociationsConfiguration") as string;
        }

        public void ApontarPolitica(string arquivo)
        {
            using var k = Registry.CurrentUser.CreateSubKey(Politica);
            k.SetValue("DefaultAssociationsConfiguration", arquivo);
        }

        public (int Codigo, string Saida) Rodar(string script, bool comClasses = true, string? pasta = null)
        {
            var psi = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, @"WindowsPowerShell\v1.0\powershell.exe"))
            { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
            foreach (var a in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", Path.Combine(Raiz, "tools", "rmm", script),
                         "-ChavePolitica", @"HKCU:\" + Politica, "-Pasta", pasta ?? Pasta })
                psi.ArgumentList.Add(a);
            if (comClasses) { psi.ArgumentList.Add("-ChaveClasses"); psi.ArgumentList.Add(@"HKCU:\" + Classes); }
            using var p = Process.Start(psi)!;
            var err = p.StandardError.ReadToEndAsync();
            var saida = p.StandardOutput.ReadToEnd();
            Assert.True(p.WaitForExit(60_000));
            return (p.ExitCode, saida + err.Result);
        }

        public List<(string Ext, string ProgId)> Associacoes(string arquivo) =>
            XDocument.Load(arquivo).Root!.Elements("Association").Select(a => ((string)a.Attribute("Identifier")!, (string)a.Attribute("ProgId")!)).ToList();

        public void Dispose()
        {
            Registry.CurrentUser.DeleteSubKeyTree(Chave, false);
            Directory.Delete(Pasta, true);
        }
    }

    [Fact] // sem os ProgIDs instalados: sai 1 e nao mexe em nada
    public void Rmm_SemProgIds_NaoAplica()
    {
        using var amb = new AmbienteRmm();
        var (codigo, saida) = amb.Rodar("aplicar-padrao-office.ps1");
        Assert.Equal(1, codigo);
        Assert.Contains("ProgIDs ausentes", saida);
        Assert.Null(amb.ValorPolitica());
        Assert.False(File.Exists(amb.Xml));
    }

    [Fact] // aplica as 6; JUNTA a associacao de outro app (.pdf do mPDF) em vez de apaga-la; idempotente;
           // o remover tira so as nossas e mantem a politica enquanto sobrar associacao de outro app
    public void Rmm_JuntaComOutraPolitica_ERemoveSoAsNossas()
    {
        using var amb = new AmbienteRmm();
        amb.RegistrarProgIds();
        var mpdf = Path.Combine(amb.Pasta, "mpdf-default-pdf.xml");
        File.WriteAllText(mpdf, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<DefaultAssociations>\n  <Association Identifier=\".pdf\" ProgId=\"mPDF.Document\" ApplicationName=\"mPDF\" />\n  <Association Identifier=\".docx\" ProgId=\"Outro.Docx\" ApplicationName=\"Outro\" />\n</DefaultAssociations>\n");
        amb.ApontarPolitica(mpdf);

        var (c1, s1) = amb.Rodar("aplicar-padrao-office.ps1");
        Assert.True(c1 == 0, s1);
        Assert.Equal(amb.Xml, amb.ValorPolitica());
        var depois = amb.Associacoes(amb.Xml);
        Assert.Contains((".pdf", "mPDF.Document"), depois);
        Assert.DoesNotContain((".docx", "Outro.Docx"), depois); // a extensao e nossa: vale a do mOffice
        foreach (var p in ProgIds) Assert.Contains((p.Ext, p.ProgId), depois);
        Assert.Equal(7, depois.Count);
        var bytes = File.ReadAllBytes(amb.Xml);
        Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "XML da politica sem BOM");

        var (c2, s2) = amb.Rodar("aplicar-padrao-office.ps1"); // idempotente: ja aponta para o nosso XML
        Assert.True(c2 == 0, s2);
        Assert.Equal(depois, amb.Associacoes(amb.Xml));

        var (c3, s3) = amb.Rodar("remover-padrao-office.ps1", comClasses: false);
        Assert.True(c3 == 0, s3);
        Assert.Equal(amb.Xml, amb.ValorPolitica()); // sobrou o .pdf: a politica fica
        Assert.Equal(new List<(string, string)> { (".pdf", "mPDF.Document") }, amb.Associacoes(amb.Xml));
    }

    [Fact] // so as nossas: o remover apaga o VALOR (a chave fica) e o XML; rodar de novo e inofensivo
    public void Rmm_RemoverSemOutras_ApagaValorEXml()
    {
        using var amb = new AmbienteRmm();
        amb.RegistrarProgIds();
        using (var k = Registry.CurrentUser.CreateSubKey(amb.Politica)) k.SetValue("OutraPolitica", 1);
        Assert.Equal(0, amb.Rodar("aplicar-padrao-office.ps1").Codigo);
        Assert.Equal(6, amb.Associacoes(amb.Xml).Count);

        var (c, s) = amb.Rodar("remover-padrao-office.ps1", comClasses: false);
        Assert.True(c == 0, s);
        Assert.Null(amb.ValorPolitica());
        Assert.False(File.Exists(amb.Xml));
        using (var k = Registry.CurrentUser.OpenSubKey(amb.Politica)) Assert.Equal(1, k?.GetValue("OutraPolitica")); // a chave e o resto ficam

        Assert.Equal(0, amb.Rodar("remover-padrao-office.ps1", comClasses: false).Codigo);
    }

    [Fact] // pasta do XML trocada por uma juncao (usuario comum pode criar em %ProgramData%): aplicar e remover
           // recusam e NADA e gravado no destino da juncao
    public void Rmm_PastaQueELink_Recusa()
    {
        using var amb = new AmbienteRmm();
        amb.RegistrarProgIds();
        var alvo = Path.Combine(amb.Pasta, "alvo");
        Directory.CreateDirectory(alvo);
        var link = Path.Combine(amb.Pasta, "link");
        // Juncao (mklink /J nao exige administrador, ao contrario do link simbolico): o que um usuario comum criaria.
        using (var mk = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"), $"/c mklink /J \"{link}\" \"{alvo}\"") { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true })!)
        {
            mk.StandardOutput.ReadToEnd();
            mk.WaitForExit();
        }
        Assert.True(new DirectoryInfo(link).Attributes.HasFlag(FileAttributes.ReparsePoint), "juncao de teste nao foi criada");
        try
        {
            foreach (var script in new[] { "aplicar-padrao-office.ps1", "remover-padrao-office.ps1" })
            {
                var (codigo, saida) = amb.Rodar(script, comClasses: script.StartsWith("aplicar"), pasta: link);
                Assert.True(codigo == 1, script + ": " + saida);
                Assert.Contains("e um link", saida);
                Assert.Empty(Directory.GetFiles(alvo));
            }
            Assert.Null(amb.ValorPolitica());
        }
        finally { Directory.Delete(link); } // apaga a JUNCAO (nao recursivo); o Dispose apaga o resto
    }

    [Fact] // politica de OUTRO arquivo (o mPDF sobrescreveu depois): o remover nao a toca
    public void Rmm_RemoverNaoTocaPoliticaAlheia()
    {
        using var amb = new AmbienteRmm();
        var alheio = Path.Combine(amb.Pasta, "outro.xml");
        File.WriteAllText(alheio, "<DefaultAssociations />");
        amb.ApontarPolitica(alheio);
        var (c, s) = amb.Rodar("remover-padrao-office.ps1", comClasses: false);
        Assert.True(c == 0, s);
        Assert.Equal(alheio, amb.ValorPolitica());
        Assert.True(File.Exists(alheio));
    }
}
