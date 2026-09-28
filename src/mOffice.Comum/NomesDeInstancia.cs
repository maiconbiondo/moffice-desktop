using System.Diagnostics;

namespace mOffice.Comum;

/// Nomes de mutex/pipe de instancia unica, por APP (mtexto, mplanilha, ...). Copia traduzida de
/// SingleInstanceNames do mPDF (src\mPdf.App\Services\ do repositorio do mPDF).
///
/// Seguranca multiusuario (mesma justificativa do mPDF): o nome embute Environment.UserName E o
/// SessionId do processo atual. UserName sozinho ja bastaria pra separar duas CONTAS diferentes na
/// mesma maquina (o namespace "Local\" do Mutex ja e' por sessao do Windows, entao duas contas seriam
/// sessoes diferentes de qualquer forma) -- mas pipes nomeados NAO sao isolados por sessao do jeito que
/// objetos "Local\" sao (o namespace \\.\pipe\ e' compartilhado pela maquina inteira, entre sessoes).
/// Sem o SessionId, duas sessoes RDP do MESMO usuario poderiam acabar com pipes de mesmo nome -- a
/// instancia secundaria da sessao A enviaria o caminho pro pipe da sessao B por engano (janela errada
/// recebendo o arquivo). Incluir os dois garante que mutex e pipe ficam sempre "emparelhados" com a
/// MESMA execucao: cada usuario, em cada sessao, ve sua propria instancia unica por app.
///
/// Testes (InstanciaUnicaTests) NUNCA usam estes nomes: cada teste gera seu proprio par via
/// Guid.NewGuid(), pra nunca colidir entre execucoes concorrentes/CI nem com uma instancia real de um
/// app rodando na maquina do dev.
public static class NomesDeInstancia
{
    // GUID fixo da SUITE mOffice -- gerado uma vez, NUNCA trocar (trocar quebraria a deteccao de
    // instancia unica entre uma versao ja rodando e outra recem-instalada por cima; toda versao NOVA
    // continua vendo esta mesma constante, entao a troca continua funcionando entre atualizacoes).
    private const string AppId = "3f6e1c2a-9b4d-4e0b-8a7c-2d5f6e7a8b9c";

    /// <summary>
    /// Mutex <c>Global\</c> de PRESENCA para o instalador (<c>AppMutex=</c> do <c>.iss</c>, Task 3 do Plano 5): cada
    /// um dos tres apps o adquire na subida (melhor esforco, <see cref="Atualizacao.InstaladorSilencioso.AdquirirMutexDoApp"/>)
    /// e o solta ao sair. UM nome para a SUITE (o instalador troca os tres de uma vez) e INDEPENDENTE dos nomes de
    /// instancia unica por app acima — nao participa de decisao nenhuma de instancia unica. Gerado uma vez; NUNCA
    /// trocar: o <c>.iss</c> tem de usar exatamente este texto (teste estrutural da Task 3), e uma versao nova
    /// instalada por cima de uma antiga tem de reconhecer o mutex da antiga.
    /// </summary>
    public const string MutexDeAtualizacao = @"Global\mOffice-Update-Mutex-73C377DD";

    public static string Mutex(string app) => Nome(app, "mutex");
    public static string Pipe(string app) => Nome(app, "pipe");

    private static string Nome(string app, string tipo)
    {
        // Backslash nao e' valido em nome de mutex/pipe (fora o prefixo "Local\"/"Global\") -- um
        // UserName no formato "DOMINIO\usuario" (raro com Environment.UserName, mas nao impossivel em
        // alguns cenarios de dominio) quebraria a construcao do nome; saneia por garantia barata.
        var usuario = Environment.UserName.Replace('\\', '_');
        var sessao = Process.GetCurrentProcess().SessionId;
        return $"mOffice-{app}-{AppId}-{tipo}-{usuario}-{sessao}";
    }
}
