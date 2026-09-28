using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 7b — descobre as amostras REAIS do escritorio em <c>poc/amostras/</c> (fora do git, CONFIDENCIAL) e
/// produz o rotulo neutro de cada uma (real-docx/real-xlsx/real-pptx/real-odt/real-ods) mais o ODP derivado do
/// PPTX real pelo LibreOffice (nenhum ODP foi entregue pelo escritorio). O mapeamento rotulo -&gt; nome real fica
/// SO em <c>poc/amostras/trabalho/MAPA.txt</c>, tambem fora do git — nada aqui, nem em teste, relatorio ou log
/// commitado, deve conter o nome real de um arquivo do escritorio.
/// </summary>
public static class AmostrasReaisCatalogo
{
    public sealed record Caso(string Rotulo, string Extensao, string CaminhoOrigem, bool Derivada);

    private static readonly object Trava = new();
    private static List<Caso>? _casos;

    public static string PastaAmostras => Path.GetFullPath(Path.Combine(
        mOffice.Conversor.Tests.MotorDeTeste.Localizar().PastaMotor, "..", "..", "poc", "amostras"));

    public static string PastaTrabalho => Path.Combine(PastaAmostras, "trabalho");

    /// <summary>Idempotente: a primeira chamada descobre e cacheia; chamadas seguintes (a suite chama isto tanto
    /// no <see cref="TheoryData{T}"/> quanto dentro de cada caso) devolvem a mesma lista, sem gerar o ODP
    /// derivado de novo.</summary>
    public static IReadOnlyList<Caso> Obter()
    {
        lock (Trava)
        {
            _casos ??= Descobrir();
            return _casos;
        }
    }

    private static List<Caso> Descobrir()
    {
        var lista = new List<Caso>();
        if (!Directory.Exists(PastaAmostras)) return lista;

        var extensoesReais = new[] { ".docx", ".xlsx", ".pptx", ".odt", ".ods" };
        var arquivos = Directory.GetFiles(PastaAmostras, "*", SearchOption.TopDirectoryOnly)
            .Where(f => extensoesReais.Contains(Path.GetExtension(f).ToLowerInvariant()))
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (arquivos.Count == 0) return lista;

        var mapa = new List<string> { "# Rotulo -> nome real do arquivo (NUNCA commitar este arquivo nem seu conteudo)" };
        foreach (var arq in arquivos)
        {
            var ext = Path.GetExtension(arq).ToLowerInvariant();
            var rotulo = "real-" + ext.TrimStart('.');
            lista.Add(new Caso(rotulo, ext, arq, Derivada: false));
            mapa.Add($"{rotulo} = {Path.GetFileName(arq)}");
        }

        // Nenhum ODP foi entregue: deriva-se do PPTX real pelo LibreOffice headless (--convert-to odp),
        // rotulado como derivado — nunca comparavel a um ODP de escritorio de verdade.
        var pptx = arquivos.FirstOrDefault(f => Path.GetExtension(f).Equals(".pptx", StringComparison.OrdinalIgnoreCase));
        if (pptx is not null)
        {
            Directory.CreateDirectory(PastaTrabalho);
            var odpDerivado = Path.Combine(PastaTrabalho, "real-odp-derivado.odp");
            if (!File.Exists(odpDerivado)) GerarOdpDerivado(pptx, PastaTrabalho, odpDerivado);
            if (File.Exists(odpDerivado))
            {
                lista.Add(new Caso("real-odp", ".odp", odpDerivado, Derivada: true));
                mapa.Add($"real-odp = derivado de {Path.GetFileName(pptx)} pelo LibreOffice --convert-to odp (NAO e' ODP de escritorio real)");
            }
        }

        try
        {
            Directory.CreateDirectory(PastaTrabalho);
            File.WriteAllLines(Path.Combine(PastaTrabalho, "MAPA.txt"), mapa);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { /* MAPA.txt e' conveniencia local; nunca derruba a descoberta */ }

        return lista;
    }

    /// <summary>`soffice --headless --convert-to odp --outdir <trabalho> <pptx>`. ProcessStartInfo.ArgumentList
    /// (nao uma string concatenada) evita a armadilha ja documentada em fidelidade.ps1/reabrir.ps1: caminho com
    /// espaco/acento chegando partido no processo filho.</summary>
    private static void GerarOdpDerivado(string pptxOrigem, string pastaTrabalho, string destino)
    {
        const string soffice = @"C:\Program Files\LibreOffice\program\soffice.exe";
        if (!File.Exists(soffice)) return;
        var psi = new System.Diagnostics.ProcessStartInfo(soffice)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        foreach (var a in new[] { "--headless", "--norestore", "--convert-to", "odp", "--outdir", pastaTrabalho, pptxOrigem })
            psi.ArgumentList.Add(a);
        using var p = System.Diagnostics.Process.Start(psi);
        if (p is null) return;
        if (!p.WaitForExit(120_000)) { try { p.Kill(entireProcessTree: true); } catch { } return; }

        // O LibreOffice nomeia o arquivo gerado pelo NOME REAL do PPTX de origem: renomeia para o nome neutro e
        // apaga o intermediario, para o nome real do arquivo do escritorio nao sobreviver em outro lugar da
        // pasta de trabalho (mesmo fora do git, e' bom nao espalhar).
        var gerado = Path.Combine(pastaTrabalho, Path.GetFileNameWithoutExtension(pptxOrigem) + ".odp");
        if (!File.Exists(gerado)) return;
        if (File.Exists(destino)) File.Delete(destino);
        File.Move(gerado, destino);
    }
}
