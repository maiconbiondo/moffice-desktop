using System.Text;
using System.Text.Json;
using mOffice.Conversor;

namespace mOffice.Comum.Tests;

/// <summary>Gera tools/installer/associacoes.iss, tools/installer/associacoes.json e tools/rmm/moffice-padrao-office.xml a
/// partir da TabelaFormatos (Plano 6, Task 2). O teste AssociacoesGeradas_EmDia falha se os arquivos commitados
/// divergirem; com MOFFICE_REGENERAR=1 ele os reescreve (e falha mesmo assim, para o commit ser consciente).
/// Fim de linha SEMPRE "\n" (o git normaliza por maquina; o teste compara sem olhar CRLF).</summary>
public static class GeradorAssociacoes
{
    public static string Exe(string docType) => docType switch { "cell" => "mPlanilha.exe", "slide" => "mSlides.exe", _ => "mTexto.exe" };
    private static string App(string docType) => Exe(docType)[..^4];

    /// <summary>O papel do arquivo para o icone (Plano 6, Task 10): apresentacao de slides, modelo, macro, legado; senao documento. A ordem decide quando ha mais de um (.dotm = modelo).</summary>
    public static string Papel(InfoExtensao i) => i.Apresentar ? "apresentacao" : i.Modelo ? "modelo" : i.Macro ? "macro" : i.Legado ? "legado" : "documento";

    /// <summary>Nome do .ico do tipo (sem pasta): <c>mtexto-documento.ico</c>, <c>mslides-apresentacao.ico</c>... (tools/icones/gerar-icones-documento.py gera os 13).</summary>
    public static string ArquivoIcone(InfoExtensao i) => $"{App(i.DocType).ToLowerInvariant()}-{Papel(i)}.ico";

    /// <summary>Icone do tipo: o .ico por papel instalado em {app}\Icones\Documentos (o Explorer deixa de mostrar o icone do .exe em todo arquivo).</summary>
    public static Func<InfoExtensao, string> Icone { get; set; } = i => $"\"\"{{app}}\\Icones\\Documentos\\{ArquivoIcone(i)}\"\",0";

    /// <summary>As extensoes que ganham "Novo" no Explorer (clique direito > Novo): o modelo vazio proprio (Modelos pt-BR, new.*; Task 9).</summary>
    public static readonly string[] ExtensoesComNovo = [".docx", ".xlsx", ".pptx"];

    public static string Iss()
    {
        var sb = new StringBuilder();
        void L(string linha) => sb.Append(linha).Append('\n');
        L("; GERADO por tests/mOffice.Comum.Tests/GeradorAssociacoes.cs a partir de TabelaFormatos - nao editar a mao.");
        L("; Regenerar: $env:MOFFICE_REGENERAR=1; dotnet test tests/mOffice.Comum.Tests --filter AssociacoesGeradas_EmDia");
        L("; Incluido pelo [Registry] do moffice.iss (#include \"associacoes.iss\"): comentarios e regras ficam la.");
        var comProgId = TabelaFormatos.Todas.Where(i => i.ProgId is not null).ToList();
        L("; --- ProgIDs: um por extensao que o mOffice abre (icone do tipo; abrir = \"<exe>\" \"%1\"; modelo = /modelo, apresentacao = /apresentar; imprimir = /imprimir) ---");
        foreach (var i in comProgId)
        {
            var exe = Exe(i.DocType);
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.ProgId}\"; ValueType: string; ValueName: \"\"; ValueData: \"{i.Descricao} ({App(i.DocType)})\"; Flags: uninsdeletekey; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.ProgId}\\DefaultIcon\"; ValueType: string; ValueName: \"\"; ValueData: \"{Icone(i)}\"; Tasks: associar");
            var verbo = i.Modelo ? "/modelo " : i.Apresentar ? "/apresentar " : "";
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.ProgId}\\shell\\open\\command\"; ValueType: string; ValueName: \"\"; ValueData: \"\"\"{{app}}\\{exe}\"\" {verbo}\"\"%1\"\"\"; Tasks: associar");
            // Verbo Imprimir do Explorer (Plano 6, Task 10): processo efemero (/imprimir converte, manda para a impressora padrao e sai).
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.ProgId}\\shell\\print\\command\"; ValueType: string; ValueName: \"\"; ValueData: \"\"\"{{app}}\\{exe}\"\" /imprimir \"\"%1\"\"\"; Tasks: associar");
            if (i.Modelo || i.Apresentar)
            {
                // Editar = abrir o PROPRIO arquivo (o modelo para alterar o modelo; a apresentacao de slides para editar os slides).
                L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.ProgId}\\shell\\edit\"; ValueType: string; ValueName: \"\"; ValueData: \"Editar\"; Tasks: associar");
                L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.ProgId}\\shell\\edit\\command\"; ValueType: string; ValueName: \"\"; ValueData: \"\"\"{{app}}\\{exe}\"\" \"\"%1\"\"\"; Tasks: associar");
            }
        }
        L("; --- OpenWithProgids: o app aparece no \"Abrir com\" da extensao (valor VAZIO; o que conta e o nome) ---");
        foreach (var i in comProgId)
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.Extensao}\\OpenWithProgids\"; ValueType: string; ValueName: \"{i.ProgId}\"; ValueData: \"\"; Flags: uninsdeletevalue; Tasks: associar");
        L("; --- ShellNew: clique direito > Novo (.docx/.xlsx/.pptx). Chave da EXTENSAO, compartilhada com o Office: so' o nosso valor sai (uninsdeletevalue) e a chave so' se vazia; NUNCA uninsdeletekey. ---");
        foreach (var i in comProgId.Where(i => ExtensoesComNovo.Contains(i.Extensao)))
        {
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.Extensao}\\ShellNew\"; ValueType: string; ValueName: \"FileName\"; ValueData: \"{{app}}\\Modelos\\pt-BR\\new{i.Extensao}\"; Flags: uninsdeletevalue; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\Classes\\{i.Extensao}\\ShellNew\"; Flags: uninsdeletekeyifempty; Tasks: associar");
        }
        L("; --- Aplicativo REGISTRADO: e disso que o Windows 11 monta \"Escolher outro aplicativo\" e Aplicativos padrao ---");
        var primeiro = true;
        foreach (var tipo in new[] { "word", "cell", "slide" })
        {
            var exe = Exe(tipo); var app = App(tipo);
            var desc = tipo switch { "cell" => "Planilhas do mOffice", "slide" => "Apresentações do mOffice", _ => "Editor de textos do mOffice" };
            L($"; {app}");
            L($"Root: HKLM; Subkey: \"Software\\Classes\\Applications\\{exe}\"; ValueType: string; ValueName: \"FriendlyAppName\"; ValueData: \"{app}\"; Flags: uninsdeletekey; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\Classes\\Applications\\{exe}\\DefaultIcon\"; ValueType: string; ValueName: \"\"; ValueData: \"\"\"{{app}}\\{exe}\"\",0\"; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\Classes\\Applications\\{exe}\\shell\\open\\command\"; ValueType: string; ValueName: \"\"; ValueData: \"\"\"{{app}}\\{exe}\"\" \"\"%1\"\"\"; Tasks: associar");
            foreach (var i in comProgId.Where(i => i.DocType == tipo))
                L($"Root: HKLM; Subkey: \"Software\\Classes\\Applications\\{exe}\\SupportedTypes\"; ValueType: string; ValueName: \"{i.Extensao}\"; ValueData: \"\"; Tasks: associar");
            if (primeiro)
            {
                // A chave pai e' compartilhada pelos tres apps: some so' quando o ultimo sai (uninsdeletekeyifempty).
                L("Root: HKLM; Subkey: \"Software\\mOffice\"; Flags: uninsdeletekeyifempty; Tasks: associar");
                primeiro = false;
            }
            L($"Root: HKLM; Subkey: \"Software\\mOffice\\{app}\"; Flags: uninsdeletekey; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\mOffice\\{app}\\Capabilities\"; ValueType: string; ValueName: \"ApplicationName\"; ValueData: \"{app}\"; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\mOffice\\{app}\\Capabilities\"; ValueType: string; ValueName: \"ApplicationDescription\"; ValueData: \"{desc}\"; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\mOffice\\{app}\\Capabilities\"; ValueType: string; ValueName: \"ApplicationIcon\"; ValueData: \"\"\"{{app}}\\{exe}\"\",0\"; Tasks: associar");
            foreach (var i in comProgId.Where(i => i.DocType == tipo))
                L($"Root: HKLM; Subkey: \"Software\\mOffice\\{app}\\Capabilities\\FileAssociations\"; ValueType: string; ValueName: \"{i.Extensao}\"; ValueData: \"{i.ProgId}\"; Tasks: associar");
            L($"Root: HKLM; Subkey: \"Software\\RegisteredApplications\"; ValueType: string; ValueName: \"mOffice.{app}\"; ValueData: \"Software\\mOffice\\{app}\\Capabilities\"; Flags: uninsdeletevalue; Tasks: associar");
        }
        return sb.ToString();
    }

    /// <summary>Lista para os scripts de desenvolvimento e do RMM (PowerShell le' com ConvertFrom-Json).</summary>
    public static string Json() => JsonSerializer.Serialize(
        TabelaFormatos.Todas.Where(i => i.ProgId is not null)
            .Select(i => new { Ext = i.Extensao, ProgId = i.ProgId, App = App(i.DocType), Exe = Exe(i.DocType), Nome = $"{i.Descricao} ({App(i.DocType)})",
                               Rmm = i.PadraoNoRmm, Modelo = i.Modelo, Apresentar = i.Apresentar }),
        new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping })
        .Replace("\r\n", "\n") + "\n";

    /// <summary>DefaultAssociationsConfiguration do RMM: so' as linhas PadraoNoRmm (as 17 que o Office faz padrao: documentos,
    /// planilhas e apresentacoes comuns, com e sem macro, e os formatos antigos; modelos e ODF de modelo ficam de fora).</summary>
    public static string XmlRmm()
    {
        var sb = new StringBuilder();
        void L(string linha) => sb.Append(linha).Append('\n');
        L("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        L("<!-- GERADO por tests/mOffice.Comum.Tests/GeradorAssociacoes.cs a partir de TabelaFormatos - nao editar a mao.");
        L("     Associacoes padrao do Windows (DefaultAssociationsConfiguration) - REFERENCIA. Forca documentos, planilhas e");
        L("     apresentacoes (.docx/.docm/.doc/.odt/.rtf, .xlsx/.xlsm/.xls/.ods, .pptx/.pptm/.ppsx/.ppt/.odp, e os modelos .dotx/.xltx/.potx)");
        L("     para mTexto, mPlanilha e mSlides a cada logon, para TODOS os usuarios da maquina. Os ProgIDs sao registrados pela task");
        L("     \"associar\" do instalador (moffice.iss): a maquina precisa ter o mOffice instalado com essa opcao marcada.");
        L("     Aplicado pelo aplicar-padrao-office.ps1 (mass-command do RMM), que le' a lista de tools/installer/associacoes.json (gerada da");
        L("     mesma tabela) e a JUNTA com as associacoes que a politica ja tinha (ex.: a do mPDF para .pdf), em vez de sobrescreve-las. -->");
        L("<DefaultAssociations>");
        foreach (var i in TabelaFormatos.Todas.Where(i => i.ProgId is not null && i.PadraoNoRmm))
            L($"  <Association Identifier=\"{i.Extensao}\" ProgId=\"{i.ProgId}\" ApplicationName=\"{App(i.DocType)}\" />");
        L("</DefaultAssociations>");
        return sb.ToString();
    }
}
