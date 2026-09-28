namespace mOffice.Nucleo.Egresso;

/// <summary>
/// Camada 2 do egresso zero (Marco 1 / Plano 1 / Task 3, spec 7.2): politicas de REGISTRO do WebView2 que
/// desligam a comunicacao de fundo que nenhuma linha de comando alcanca — o servico de experimentacao e
/// configuracao do Edge (o que baixa as "flags" remotas do Chromium) e' ligado pelo processo do navegador
/// antes de qualquer argumento nosso chegar a ele.
///
/// Fonte: documentacao de politicas do WebView2 (learn.microsoft.com, "Microsoft Edge WebView2 - Policies"):
/// <c>ExperimentationAndConfigurationServiceControl</c> = 0 significa <c>RestrictedMode</c> — o servico nao e'
/// consultado. A chave fica em HKLM (por maquina): e' politica, nao preferencia de usuario.
///
/// Quem aplica em producao e' o INSTALADOR, como administrador, com <see cref="ComandosReg"/>; a
/// desinstalacao roda <see cref="ComandosRegRemover"/>. tools/egresso/politicas.ps1 faz o mesmo na bancada de
/// medida, e <c>PoliticasTests.ScriptDaBancada_NaoDivergiuDaLista</c> impede que as duas listas divirjam.
/// </summary>
public static class Politicas
{
    public const string Chave = @"HKLM\SOFTWARE\Policies\Microsoft\Edge\WebView2";

    /// <summary>Politicas aplicadas pelo mOffice. Hoje uma so': a medida da Task 3 nao achou outra politica
    /// documentada que mudasse o numero de conexoes (ver poc/relatorios/egresso-S2-*.md). Acrescentar aqui
    /// exige medida, nao palpite.</summary>
    public static IReadOnlyList<(string Chave, string Nome, object Valor)> Recomendadas { get; } =
    [
        (Chave, "ExperimentationAndConfigurationServiceControl", 0),
    ];

    /// <summary>`reg add` idempotentes (/f): rodar duas vezes da' o mesmo resultado, e' o que o instalador precisa.</summary>
    public static string ComandosReg() => string.Join(Environment.NewLine,
        Recomendadas.Select(p => $"reg add \"{p.Chave}\" /v {p.Nome} /t REG_DWORD /d {p.Valor} /f"));

    /// <summary>O desfazer da <see cref="ComandosReg"/>: um `reg delete` por politica. Apaga o VALOR, nunca a
    /// chave inteira — outros produtos (o proprio Edge corporativo) podem ter politicas suas ali.</summary>
    public static string ComandosRegRemover() => string.Join(Environment.NewLine,
        Recomendadas.Select(p => $"reg delete \"{p.Chave}\" /v {p.Nome} /f"));
}
