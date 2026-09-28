using System.Runtime.CompilerServices;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Higiene de teste (achado menor da revisao final do Plano 2). Os testes de janela real usam o produto de
/// verdade, e o produto grava em <c>Pastas.Raiz</c> — ou seja, ate' aqui a suite escrevia dentro do
/// <c>%LOCALAPPDATA%\mOffice</c> REAL do desenvolvedor: pastas de recuperacao, arquivos <c>.log</c> por
/// janela, perfis do WebView2 e <c>recentes-*.json</c> misturados com os do app que ele usa para trabalhar.
/// Alem de sujo, e' arriscado nos dois sentidos: a varredura de retencao de um teste pode apagar coisa de
/// uma sessao real, e um app aberto pode segurar arquivo que o teste espera poder apagar.
///
/// Conserto: apontar <c>MOFFICE_DADOS_DIR</c> (a variavel que o proprio produto le', ver <c>Pastas</c>) para
/// <c>%LOCALAPPDATA%\mOffice-testes\dados</c> (fora de <c>%TEMP%</c> — ver <see cref="PastaDados"/>) ANTES
/// de qualquer teste rodar. O <c>[ModuleInitializer]</c> e' o unico
/// gancho que roda garantidamente antes do primeiro tipo desta assembly ser usado, independente da ordem dos
/// testes — um <c>IAsyncLifetime</c> ou um fixture nao dariam essa garantia.
///
/// Uma variavel ja' definida por FORA e' respeitada: e' assim que uma execucao em CI, ou o desenvolvedor
/// investigando um caso, aponta a suite para outro lugar sem editar codigo.
/// </summary>
public static class DadosDeTeste
{
    /// <summary>Pastas de TRABALHO dos testes (documentos sinteticos, motores temporarios): as
    /// <c>%TEMP%\mOffice-testes\&lt;guid&gt;</c> que a suite ja' criava e nem sempre limpava. E' o que a
    /// varredura de <see cref="Preparar"/> poda.</summary>
    public static string RaizDeTestes => Path.Combine(Path.GetTempPath(), "mOffice-testes");

    /// <summary>
    /// Onde <c>Pastas.Raiz</c> aponta durante a suite (recuperacao, logs, config, recentes, perfis do
    /// WebView2).
    ///
    /// FORA de <c>%TEMP%</c>, e nao dentro de <see cref="RaizDeTestes"/> como seria mais simetrico: a spec §4
    /// diz "nada em %TEMP%" e DUAS guardas do produto verificam isso lendo o caminho EFETIVO
    /// (<c>PastasTests.SessaoDocumento_RaizRecovery_EstaSobPastas</c> e
    /// <c>AmbienteWebView2Tests.PastaDados_PorApp</c>). Apontar a variavel para dentro de <c>%TEMP%</c> as
    /// fazia falhar — corretamente: elas nao tem como distinguir "o produto escolheu %TEMP%" de "o arranjo do
    /// teste escolheu %TEMP%". Preferiu-se mover a pasta a afrouxar as guardas para olharem so' a raiz padrao,
    /// que nenhum codigo usa durante a suite. <c>mOffice-testes</c> ao lado (e nunca DENTRO) de
    /// <c>mOffice</c>: a pasta real do desenvolvedor continua intocada.
    /// </summary>
    public static string PastaDados => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "mOffice-testes", "dados");

    /// <summary>Quanto uma pasta de trabalho de teste pode ficar em <see cref="RaizDeTestes"/> antes de a
    /// proxima execucao apaga-la. 1 dia: mais do que qualquer execucao da suite dura, e curto o bastante
    /// para o lixo nao se acumular por semanas.</summary>
    public static readonly TimeSpan IdadeMaximaDoLixo = TimeSpan.FromDays(1);

    private static int _preparado;

    /// <summary>Idempotente: as duas assemblies de teste (Nucleo e Comum) chamam, e so' a primeira faz algo.</summary>
    public static void Preparar()
    {
        if (Interlocked.Exchange(ref _preparado, 1) == 1) return;
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(Pastas.VariavelDeAmbiente)))
            Environment.SetEnvironmentVariable(Pastas.VariavelDeAmbiente, PastaDados);
        VarrerLixoAntigo();
    }

    /// <summary>
    /// Apaga as pastas de trabalho de execucoes anteriores (as <c>&lt;guid&gt;</c> que os testes criam e que
    /// nem sempre limpam) com mais de <see cref="IdadeMaximaDoLixo"/>.
    ///
    /// Contencao de proposito: so' olha DENTRO de <see cref="RaizDeTestes"/> — nada fora dela e' candidato,
    /// em nenhuma circunstancia. <see cref="PastaDados"/> nao corre risco porque nem mora aqui (ver o
    /// comentario la'); a guarda por igualdade de caminho fica assim mesmo, barata, para o dia em que alguem
    /// mover a pasta de dados para ca' sem se lembrar disto. Nunca lanca: uma falha de limpeza nao pode
    /// derrubar a suite inteira antes do primeiro teste.
    /// </summary>
    private static void VarrerLixoAntigo()
    {
        try
        {
            if (!Directory.Exists(RaizDeTestes)) return;
            var limite = DateTime.UtcNow - IdadeMaximaDoLixo;
            var dados = Path.GetFullPath(PastaDados);
            foreach (var pasta in Directory.EnumerateDirectories(RaizDeTestes))
            {
                try
                {
                    var completo = Path.GetFullPath(pasta);
                    if (dados.StartsWith(completo, StringComparison.OrdinalIgnoreCase)) continue;
                    if (Directory.GetLastWriteTimeUtc(pasta) >= limite) continue;
                    Directory.Delete(pasta, recursive: true);
                }
                catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
            }
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }
}

internal static class InicializacaoModulo
{
    [ModuleInitializer]
    internal static void Preparar() => DadosDeTeste.Preparar();
}
