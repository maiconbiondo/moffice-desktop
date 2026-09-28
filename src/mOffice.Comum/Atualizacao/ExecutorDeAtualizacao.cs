using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace mOffice.Comum.Atualizacao;

public enum ResultadoAtualizacaoCli { JaNaUltima, AtualizacaoDisponivel, AtualizacaoIniciada, Erro }

public sealed record StatusAtualizacaoCli(
    ResultadoAtualizacaoCli Resultado, string? VersaoInstalada, string? VersaoDisponivel, string? MensagemErro);

/// <summary>
/// Os verbos do RMM, em QUALQUER dos tres executaveis: <c>/checkupdate</c> (so' verifica) e <c>/update</c> (baixa,
/// verifica o SHA-256 e roda o instalador silencioso com <c>/NORELAUNCH</c>). Sem interface e sem confirmacao.
/// Porte de <c>SilentUpdateRunner.cs</c> do mPDF. Codigo de saida: 0 = ja' na ultima, 10 = ha' versao nova (ou a
/// instalacao foi iniciada), 20 = erro. Constroi o <see cref="ServicoAtualizacao"/> — um dos DOIS sitios
/// deliberados (o outro e' o botao das Configurações; ver RedeConfinadaTests).
///
/// DESACOMPANHADO de proposito (fix round 1): ao contrario do botao, o <c>/update</c> NAO confere se ha' apps da suite
/// abertos nem pergunta nada. O RMM roda quando os usuarios estao deslogados (ou fora do expediente); se houver um app
/// aberto, o instalador o fecha (<c>CloseApplications</c> do <c>.iss</c>) e o trabalho nao salvo fica na pasta de
/// recuperacao, que a proxima abertura oferece ("Documentos recuperados"). Este processo efemero nao segura o mutex de
/// presenca (o roteamento do verbo vem antes de adquiri-lo), entao o instalador nunca o ve' como app aberto.
/// </summary>
public sealed class ExecutorDeAtualizacao
{
    public const string VerboVerificar = "/checkupdate";
    public const string VerboAtualizar = "/update";

    private readonly Func<IFonteDeReleases> _criarFonte;
    private readonly Action<ServicoAtualizacao.ArquivoVerificado> _lancarInstalador;

    public ExecutorDeAtualizacao(Func<IFonteDeReleases>? criarFonte = null,
        Action<ServicoAtualizacao.ArquivoVerificado>? lancarInstalador = null)
    {
        _criarFonte = criarFonte ?? ServicoAtualizacao.CriarFontePadrao;
        // reabrir:false — o /update do RMM nunca reabre o app na estacao (atualizacao em massa invisivel).
        _lancarInstalador = lancarInstalador ?? (a => Process.Start(InstaladorSilencioso.MontarStartInfo(a, reabrir: false)));
    }

    /// <summary>O primeiro argumento e' um dos verbos? Devolve o verbo normalizado (minusculas) ou <c>null</c>.</summary>
    public static string? VerboDe(string[] args)
    {
        if (args.Length == 0) return null;
        var v = args[0].Trim().ToLowerInvariant();
        return v is VerboVerificar or VerboAtualizar ? v : null;
    }

    private ServicoAtualizacao CriarServico() => new ServicoAtualizacao(_criarFonte());

    /// <summary>Executa o verbo (<see cref="VerboVerificar"/> ou <see cref="VerboAtualizar"/>).</summary>
    public Task<StatusAtualizacaoCli> ExecutarAsync(string verbo, CancellationToken ct) =>
        verbo == VerboAtualizar ? AtualizarAsync(ct) : VerificarAsync(ct);

    public async Task<StatusAtualizacaoCli> VerificarAsync(CancellationToken ct)
    {
        var instalada = ServicoAtualizacao.VersaoAtualTexto();
        try
        {
            using var servico = CriarServico();
            var r = await servico.VerificarAsync(ct);
            return r.Status switch
            {
                StatusVerificacao.Atualizado => new(ResultadoAtualizacaoCli.JaNaUltima, instalada, instalada, null),
                StatusVerificacao.Disponivel => new(ResultadoAtualizacaoCli.AtualizacaoDisponivel, instalada, ServicoAtualizacao.VersaoDaTag(r.Info!.TagVersao), null),
                _ => new(ResultadoAtualizacaoCli.Erro, instalada, null, r.MensagemErro),
            };
        }
        catch (Exception ex) { return new(ResultadoAtualizacaoCli.Erro, instalada, null, ex.Message); }
    }

    public async Task<StatusAtualizacaoCli> AtualizarAsync(CancellationToken ct)
    {
        var instalada = ServicoAtualizacao.VersaoAtualTexto();
        // Downloads abandonados (> 1 dia) de execucoes anteriores: o /update roda sem a subida do app, que tambem varre.
        ServicoAtualizacao.VarrerDownloadsAntigos();
        try
        {
            using var servico = CriarServico();
            var r = await servico.VerificarAsync(ct);
            if (r.Status == StatusVerificacao.Atualizado)
                return new(ResultadoAtualizacaoCli.JaNaUltima, instalada, instalada, null);
            if (r.Status != StatusVerificacao.Disponivel)
                return new(ResultadoAtualizacaoCli.Erro, instalada, null, r.MensagemErro);

            var disponivel = ServicoAtualizacao.VersaoDaTag(r.Info!.TagVersao);
            var d = await servico.BaixarEVerificarAsync(r.Info, null, ct);
            if (d.Status != StatusDownload.Verificado || d.Arquivo is null)
                return new(ResultadoAtualizacaoCli.Erro, instalada, disponivel, d.MensagemErro);
            return InstalarVerificado(d.Arquivo, instalada, disponivel);
        }
        catch (Exception ex) { return new(ResultadoAtualizacaoCli.Erro, instalada, null, ex.Message); }
    }

    /// <summary>So' um <see cref="ServicoAtualizacao.ArquivoVerificado"/> chega ao lancador.</summary>
    internal StatusAtualizacaoCli InstalarVerificado(ServicoAtualizacao.ArquivoVerificado arquivo, string? instalada, string? disponivel)
    {
        // Confere o hash de novo e TRAVA o arquivo ate' o Process.Start voltar (ver ReconferirETravar).
        using var trava = arquivo.ReconferirETravar();
        if (trava is null)
            return new(ResultadoAtualizacaoCli.Erro, instalada, disponivel, ServicoAtualizacao.MensagemAlteradoDepois);
        try
        {
            _lancarInstalador(arquivo);
            return new(ResultadoAtualizacaoCli.AtualizacaoIniciada, instalada, disponivel, null);
        }
        catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == AtualizacaoAoEncerrar.UacRecusado)
        {
            return new(ResultadoAtualizacaoCli.Erro, instalada, disponivel, AtualizacaoAoEncerrar.MensagemCancelada);
        }
        catch (Exception ex) { return new(ResultadoAtualizacaoCli.Erro, instalada, disponivel, ex.Message); }
    }

    public static int MapearExitCode(ResultadoAtualizacaoCli r) => r switch
    {
        ResultadoAtualizacaoCli.JaNaUltima => 0,
        ResultadoAtualizacaoCli.AtualizacaoDisponivel => 10,
        ResultadoAtualizacaoCli.AtualizacaoIniciada => 10,
        _ => 20,
    };

    /// <summary>A linha de status, SO' ASCII (o agente do RMM le' a saida sem saber a codificacao). Ex.:
    /// <c>mOffice /checkupdate: instalada=1.0.0 disponivel=1.0.1 atualizar=sim</c>.</summary>
    public static string FormatarStatus(string comando, StatusAtualizacaoCli s)
    {
        var inst = s.VersaoInstalada ?? "-";
        var linha = s.Resultado switch
        {
            ResultadoAtualizacaoCli.Erro => $"mOffice {comando}: instalada={inst} erro={s.MensagemErro ?? "desconhecido"}",
            ResultadoAtualizacaoCli.AtualizacaoIniciada => $"mOffice {comando}: instalada={inst} disponivel={s.VersaoDisponivel ?? "-"} acao=instalando",
            ResultadoAtualizacaoCli.AtualizacaoDisponivel => $"mOffice {comando}: instalada={inst} disponivel={s.VersaoDisponivel ?? "-"} atualizar=sim",
            _ => $"mOffice {comando}: instalada={inst} disponivel={inst} atualizar=nao",
        };
        return ParaAscii(linha);
    }

    /// <summary>Tira os acentos (NFD sem as marcas) e troca o que sobrar fora do ASCII imprimivel por '?': as
    /// mensagens de erro sao pt-BR ("Não foi possível…") e a linha tem de continuar ASCII.</summary>
    internal static string ParaAscii(string texto)
    {
        var sb = new StringBuilder(texto.Length);
        foreach (var c in texto.Normalize(NormalizationForm.FormD))
        {
            if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
            sb.Append(c is >= ' ' and <= '~' ? c : c is '\r' or '\n' or '\t' ? ' ' : '?');
        }
        return sb.ToString();
    }
}

/// <summary>
/// Escreve a linha de status para quem chamou <c>mTexto.exe /update</c> (os apps sao WinExe, sem console proprio).
/// Porte de <c>ConsoleStatus.cs</c> do mPDF: (1) stdout HERDADO — o agente do RMM redireciona a saida e o processo a
/// herda; (2) <c>AttachConsole</c> do pai como reserva. Sem nenhum dos dois (aberto pelo Explorer), nada: o codigo
/// de saida e' o sinal principal. Melhor esforco: nunca derruba o processo.
/// </summary>
public static class SaidaDeStatus
{
    private const uint AttachParentProcess = 0xFFFFFFFF;
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AttachConsole(uint dwProcessId);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool FreeConsole();

    public static void Escrever(string linha)
    {
        if (NoStdoutHerdado(linha)) return;
        ViaAttachConsole(linha);
    }

    private static bool NoStdoutHerdado(string linha)
    {
        try
        {
            using var stdout = Console.OpenStandardOutput();
            if (stdout == Stream.Null) return false;
            var bytes = Encoding.ASCII.GetBytes(linha + "\r\n");
            stdout.Write(bytes, 0, bytes.Length);
            stdout.Flush();
            return true;
        }
        catch { return false; }
    }

    private static void ViaAttachConsole(string linha)
    {
        try
        {
            if (!AttachConsole(AttachParentProcess)) return;
            try { Console.Out.WriteLine(linha); Console.Out.Flush(); }
            finally { FreeConsole(); }
        }
        catch { /* melhor esforco */ }
    }
}
