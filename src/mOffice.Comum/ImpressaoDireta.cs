using mOffice.Conversor;
using mOffice.Nucleo;
using mOffice.Nucleo.Impressao;

namespace mOffice.Comum;

/// <summary>Verbo "Imprimir" do Explorer (<c>/imprimir "%1"</c>, Plano 6 Task 10 — SO-09): um arquivo que NAO esta' aberto vai para a impressora padrao sem abrir janela
/// de documento. Reaproveita o caminho de impressao do Plano 3 (x2t gera o PDF, a <see cref="Impressora"/> o entrega) — nenhum motor de impressao novo.
/// Codigos de saida como os verbos do RMM: 0 = enviado, 20 = erro.</summary>
public static class ImpressaoDireta
{
    public const int Ok = 0;
    public const int Erro = 20;

    /// <summary>Converte <paramref name="caminho"/> em PDF numa pasta de impressao, chama <paramref name="enviar"/>(pdf, app) e APAGA o PDF no fim (deu certo ou nao).
    /// <paramref name="enviar"/> padrao: <see cref="Impressora.ImprimirAsync"/> com <see cref="OpcoesImpressao.Padrao"/> (impressora padrao, colorida) — precisa do
    /// Dispatcher rodando (a WebView2 oculta), entao NAO se espera esta tarefa com <c>GetAwaiter().GetResult()</c> na thread de interface.
    /// <paramref name="converter"/>: costura de teste (o x2t de verdade e' o padrao). Nunca lanca: falha = <see cref="Erro"/>.</summary>
    public static async Task<int> ImprimirAsync(string caminho, string app, Func<string, string, Task>? enviar = null,
        Func<string, string, Task>? converter = null, string? pastaImpressao = null, Action<string>? log = null, Action<string>? avisar = null)
    {
        string? pdf = null;
        try
        {
            if (!File.Exists(caminho))
            {
                log?.Invoke("IMPRIMIR arquivo nao encontrado: " + caminho);
                Avisar(avisar, caminho, "o arquivo não foi encontrado.");
                return Erro;
            }
            pdf = PastaImpressao.NovoArquivo(pastaImpressao ?? Pastas.Impressao);
            converter ??= (origem, destino) => new X2tConversor(ConfiguracaoMotor.Localizar()).ConverterAsync(origem, FormatoDocumento.Pdf, destino);
            enviar ??= async (arquivoPdf, nomeApp) => await new Impressora(nomeApp).ImprimirAsync(arquivoPdf, OpcoesImpressao.Padrao);
            await converter(caminho, pdf);
            await enviar(pdf, app);
            return Ok;
        }
        catch (Exception e)
        {
            log?.Invoke("IMPRIMIR falhou: " + e.Message);
            Avisar(avisar, caminho, MotivoDaFalha(e));
            return Erro;
        }
        finally { if (pdf is not null) await PastaImpressao.ApagarComEsperaAsync(pdf); }
    }

    /// <summary>O Office pergunta/avisa quando nao consegue imprimir: o processo efemero tambem — uma caixa pt-BR antes do 20 (nunca falha calado).</summary>
    private static void Avisar(Action<string>? avisar, string caminho, string motivo)
    {
        try { avisar?.Invoke($"Não foi possível imprimir \"{Path.GetFileName(caminho)}\": {motivo}"); }
        catch (Exception) { /* o aviso nunca atrapalha o encerramento */ }
    }

    private static string MotivoDaFalha(Exception e) => e switch
    {
        ErroConversao => "o arquivo não pôde ser aberto para impressão. Ele pode estar protegido por senha, em uso por outro programa ou danificado.",
        _ => e.Message.TrimEnd() is { Length: > 0 } m ? (m.EndsWith('.') ? m : m + ".") : "erro inesperado.",
    };
}
