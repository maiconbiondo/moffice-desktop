namespace mOffice.Nucleo;

/// <summary>
/// Onde o mOffice grava dados por usuario (spec do Marco 1, §4): %LOCALAPPDATA%\mOffice. A variavel de
/// ambiente MOFFICE_DADOS_DIR redireciona tudo (isolamento de teste; producao nunca a define) — mesmo
/// padrao do MPDF_CONFIG_DIR do mPDF. Nada do produto grava em %TEMP% (backlog 15).
/// </summary>
public static class Pastas
{
    public const string VariavelDeAmbiente = "MOFFICE_DADOS_DIR";

    public static string RaizPadrao() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "mOffice");

    /// <summary>Raiz efetiva: a variavel, quando nao esta' em branco; senao o padrao.</summary>
    public static string Resolver(string? variavel) =>
        string.IsNullOrWhiteSpace(variavel) ? RaizPadrao() : variavel.Trim();

    public static string Raiz => Resolver(Environment.GetEnvironmentVariable(VariavelDeAmbiente));

    public static string Recuperacao => RecuperacaoEm(Raiz);
    public static string Logs => LogsEm(Raiz);
    public static string Config => ConfigEm(Raiz);
    public static string Recentes(string app) => RecentesEm(Raiz, app);
    /// <summary>Palavras do "Adicionar ao dicionario" (Task 1 do Plano 3), uma por linha; compartilhado pelos
    /// tres apps.</summary>
    public static string DicionarioUsuario => DicionarioUsuarioEm(Raiz);
    /// <summary>PDFs de impressao (Task 4 do Plano 3): o x2t gera aqui o PDF que vai para a impressora. Cada
    /// janela apaga os dela ao fechar; a subida do app varre os que sobraram (<see cref="mOffice.Nucleo.Impressao.PastaImpressao"/>).</summary>
    public static string Impressao => ImpressaoEm(Raiz);

    /// <summary>"Não salvos recentemente" (Plano 6, Task 10): o que o "Não salvar" do fechamento move da recuperação e fica 4 dias.</summary>
    public static string Descartados => DescartadosEm(Raiz);

    public static string RecuperacaoEm(string raiz) => Path.Combine(raiz, "recuperacao");
    public static string DescartadosEm(string raiz) => Path.Combine(raiz, "descartados");
    public static string LogsEm(string raiz) => Path.Combine(raiz, "logs");
    public static string ConfigEm(string raiz) => Path.Combine(raiz, "config.json");
    public static string RecentesEm(string raiz, string app) => Path.Combine(raiz, $"recentes-{app.ToLowerInvariant()}.json");
    public static string DicionarioUsuarioEm(string raiz) => Path.Combine(raiz, "dicionario-usuario.txt");
    public static string ImpressaoEm(string raiz) => Path.Combine(raiz, "impressao");

    /// <summary>Cria a pasta (idempotente) e devolve o mesmo caminho, para encadear.</summary>
    public static string Garantir(string pasta) { Directory.CreateDirectory(pasta); return pasta; }
}
