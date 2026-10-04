namespace mOffice.Nucleo;

/// <summary>Atalhos do Office que o editor web nao tem e que o HOST trata (Plano 6, Task 7). A ponte (ponte-desktop.js) captura a
/// tecla antes do editor e manda o NOME pelo aviso "atalho".</summary>
public enum AcaoDeAtalho { Abrir, Fechar, SalvarComo, Salvar, ProximaJanela }

public static class AtalhosDoHost
{
    /// <summary>Nome que a ponte manda -> acao. Nome desconhecido (ou com outra caixa) = <c>null</c>: a pagina nao inventa acao do host.</summary>
    public static AcaoDeAtalho? DoAviso(string nome) => nome switch
    {
        "abrir" => AcaoDeAtalho.Abrir,
        "fechar" => AcaoDeAtalho.Fechar,
        "salvarComo" => AcaoDeAtalho.SalvarComo,
        "salvar" => AcaoDeAtalho.Salvar,
        "proximaJanela" => AcaoDeAtalho.ProximaJanela,
        _ => null,
    };
}
