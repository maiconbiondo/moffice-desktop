using mOffice.Comum.Atualizacao;
using mOffice.Nucleo;

namespace mOffice.Comum;

/// <summary>
/// Verbo <c>/gerarcatalogo</c> (Plano 5, Task 3): gera o catalogo de fontes SEM abrir janela, pelo MESMO caminho
/// que a janela do editor usa (<see cref="CatalogoFontes.Padrao"/> + <see cref="CatalogoFontes.GarantirAtualizadoAsync"/>),
/// e sai. Quem chama e o instalador (<c>[Run]</c> do <c>moffice.iss</c>), como administrador: ali a pasta por
/// maquina (<c>%ProgramData%\mOffice\fontes</c>, so leitura para usuarios) e gravavel, e o catalogo nasce la.
///
/// Por que um verbo e nao o <c>allfontsgen.exe</c> direto no <c>[Run]</c>: o arquivo
/// <c>impressao-digital.txt</c> que o <see cref="CatalogoFontes"/> grava por ultimo e o que diz ao app que o
/// catalogo esta atual. Sem ele (gerador chamado por fora), a primeira janela regeneraria tudo de novo — e as
/// pastas de fontes e argumentos teriam de ser repetidos no <c>.iss</c>, uma segunda definicao para divergir.
///
/// Saida: UMA linha ASCII no stdout (mesmo <see cref="SaidaDeStatus"/> dos verbos de atualizacao) e codigo 0
/// (catalogo pronto, gerado agora ou ja atual) ou 20 (erro). Nunca lanca.
/// </summary>
public static class VerboCatalogo
{
    public const string Verbo = "/gerarcatalogo";

    public const int CodigoOk = 0;
    public const int CodigoErro = 20;

    /// <summary>So o PRIMEIRO argumento conta (mesma regra do <see cref="ExecutorDeAtualizacao.VerboDe"/>).</summary>
    public static bool EhVerbo(string[] args) =>
        args.Length > 0 && string.Equals(args[0].Trim(), Verbo, StringComparison.OrdinalIgnoreCase);

    /// <summary>Gera (se preciso) e devolve a linha de status e o codigo de saida. <paramref name="criar"/> e a
    /// costura de teste; a producao passa <c>() =&gt; CatalogoFontes.Padrao(ConfiguracaoMotor.Localizar())</c>.
    /// Qualquer excecao (motor nao encontrado, por exemplo) vira linha de erro.</summary>
    public static async Task<(string Linha, int Codigo)> ExecutarAsync(Func<CatalogoFontes> criar)
    {
        try
        {
            var catalogo = criar();
            var r = await catalogo.GarantirAtualizadoAsync();
            if (r.Erro is not null) return (Formatar("erro=" + r.Erro), CodigoErro);
            return (Formatar($"pasta={catalogo.PastaCatalogo} fontes={r.Fontes} regenerou={(r.Regenerou ? "sim" : "nao")}"), CodigoOk);
        }
        catch (Exception e)
        {
            return (Formatar("erro=" + e.Message), CodigoErro);
        }
    }

    private static string Formatar(string resto) => ExecutorDeAtualizacao.ParaAscii($"mOffice {Verbo}: {resto}");
}
