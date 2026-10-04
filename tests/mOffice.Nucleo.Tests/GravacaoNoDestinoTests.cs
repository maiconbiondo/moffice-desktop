namespace mOffice.Nucleo.Tests;

public class GravacaoNoDestinoTests : IDisposable
{
    private readonly List<string> _pastas = [];

    private string Pasta()
    {
        var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(p);
        _pastas.Add(p);
        return p;
    }

    public void Dispose() { foreach (var p in _pastas) try { Directory.Delete(p, true); } catch { } }

    private static string Backups(string pasta) => string.Join(",", Directory.GetFiles(pasta, "*" + GravacaoNoDestino.SufixoBackup + "*"));

    [Fact] // SO-14: substituir um arquivo existente preserva a data de criacao dele (ReplaceFile), e o conteudo e' o novo
    public void Substituir_DestinoExistente_PreservaDataDeCriacao()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "peticao.docx"); File.WriteAllText(destino, "velho");
        var criado = new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        File.SetCreationTimeUtc(destino, criado);
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        Assert.Equal(ResultadoGravacao.MetadadosPreservados, GravacaoNoDestino.Substituir(tmp, destino));
        Assert.Equal("novo", File.ReadAllText(destino));
        Assert.Equal(criado, File.GetCreationTimeUtc(destino));
        Assert.False(File.Exists(tmp));
        Assert.Equal("", Backups(p)); // backup apagado
    }

    [Fact] // destino inexistente: move simples (nada a preservar) — e' o estado "Novo", nao uma falha do Replace
    public void Substituir_SemDestino_Move()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "novo.docx");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "x");
        Assert.Equal(ResultadoGravacao.Novo, GravacaoNoDestino.Substituir(tmp, destino));
        Assert.Equal("x", File.ReadAllText(destino));
    }

    [Fact] // 1176: o Replace removeu o destino e parou com o substituto ainda no nome temporario -> o conteudo novo entra no lugar
    public void Replace_FalhaComDestinoRemovido_SemBackup_TemporarioViraDestino()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        var r = GravacaoNoDestino.Substituir(tmp, destino, replace: (_, alvo, _) =>
        {
            File.Delete(alvo);
            throw new IOException("1176 simulado");
        });
        Assert.Equal(ResultadoGravacao.SemPreservar, r);
        Assert.Equal("novo", File.ReadAllText(destino));
        Assert.False(File.Exists(tmp));
    }

    [Fact] // 1177: destino so' sobrevive no backup -> volta do backup e o Move de contingencia grava o conteudo novo
    public void Replace_FalhaComDestinoSoNoBackup_Restaura_EGrava()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        var r = GravacaoNoDestino.Substituir(tmp, destino, replace: (_, alvo, bak) =>
        {
            File.Move(alvo, bak);
            throw new IOException("1177 simulado");
        });
        Assert.Equal(ResultadoGravacao.SemPreservar, r);
        Assert.Equal("novo", File.ReadAllText(destino));
        Assert.Equal("", Backups(p));
    }

    [Fact] // Replace recusado sem tocar em nada (rede antiga): contingencia por Move, conteudo novo gravado
    public void Replace_Recusado_Contingencia_Move()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        var r = GravacaoNoDestino.Substituir(tmp, destino, replace: (_, _, _) => throw new IOException("recusado"));
        Assert.Equal(ResultadoGravacao.SemPreservar, r);
        Assert.Equal("novo", File.ReadAllText(destino));
        Assert.False(File.Exists(tmp));
    }

    [Fact] // destino travado de verdade (leitor sem FileShare): o Replace real falha, o Move tambem -> relanca e o arquivo velho segue intacto.
           // O temporario SOBRA aqui porque Substituir nao o apaga; quem o apaga e' o finally do chamador, e so' com o destino existindo (regra do finally: ver Fix 1).
    public void DestinoTravado_Relanca_ArquivoVelhoIntacto_TemporarioPreservado()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        using (new FileStream(destino, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            Assert.ThrowsAny<Exception>(() => GravacaoNoDestino.Substituir(tmp, destino));
        }
        Assert.Equal("velho", File.ReadAllText(destino));
        Assert.Equal("novo", File.ReadAllText(tmp)); // Substituir nao mexe no temporario quando relanca
        Assert.Equal("", Backups(p));
    }

    [Fact] // impossivel recuperar o destino (sem backup e sem temporario): relanca, nada e' inventado
    public void SemDestinoSemBackupSemTemporario_Relanca()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        Assert.Throws<IOException>(() => GravacaoNoDestino.Substituir(tmp, destino, replace: (t, alvo, _) =>
        {
            File.Delete(alvo); File.Delete(t);
            throw new IOException("perdido");
        }));
    }

    [Fact] // destino de pe' + backup antigo ao lado (lixo de gravacao anterior): removido, com registro; a gravacao segue
    public void DestinoExiste_BackupAntigo_ERemovidoELogado()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var antigo = destino + GravacaoNoDestino.SufixoBackup; File.WriteAllText(antigo, "lixo");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        var log = new List<string>();
        Assert.Equal(ResultadoGravacao.MetadadosPreservados, GravacaoNoDestino.Substituir(tmp, destino, log.Add));
        Assert.Equal("novo", File.ReadAllText(destino));
        Assert.Equal("", Backups(p));
        Assert.Contains(log, l => l.Contains("backup antigo"));
    }

    [Fact] // uma gravacao que FALHA nunca descarta o backup anterior: o lixo so' sai depois de uma troca bem-sucedida
    public void FalhaNaoDescartaBackupAnterior()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var antigo = destino + GravacaoNoDestino.SufixoBackup; File.WriteAllText(antigo, "anterior");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        using (new FileStream(destino, FileMode.Open, FileAccess.Read, FileShare.None))
            Assert.ThrowsAny<Exception>(() => GravacaoNoDestino.Substituir(tmp, destino));
        Assert.Equal("anterior", File.ReadAllText(antigo));
    }

    [Fact] // so' "<nome>.moffice-bak" e "<nome>.moffice-bak<N>" contam como backup: parecidos (bakup, .txt) ficam
    public void LimpezaDeBackups_CasaExatamenteONome()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var apagar = new[] { destino + ".moffice-bak", destino + ".moffice-bak7" };
        var manter = new[] { destino + ".moffice-bakup", destino + ".moffice-bak.txt", Path.Combine(p, "b.docx.moffice-bak") };
        foreach (var f in apagar.Concat(manter)) File.WriteAllText(f, "x");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        GravacaoNoDestino.Substituir(tmp, destino);
        Assert.All(apagar, f => Assert.False(File.Exists(f), f));
        Assert.All(manter, f => Assert.True(File.Exists(f), f));
    }

    [Fact] // destino AUSENTE + backup ao lado (kill no meio do ReplaceFile): o backup e' o unico original -> nunca apagado, so' logado
    public void DestinoAusente_BackupOrfao_MantidoELogado()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx");
        var orfao = destino + GravacaoNoDestino.SufixoBackup; File.WriteAllText(orfao, "original");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        var log = new List<string>();
        Assert.Equal(ResultadoGravacao.Novo, GravacaoNoDestino.Substituir(tmp, destino, log.Add));
        Assert.Equal("original", File.ReadAllText(orfao));
        Assert.Equal("novo", File.ReadAllText(destino));
        Assert.Contains(log, l => l.Contains(orfao) && l.Contains("ATENCAO"));
    }

    [Fact] // 5b: 1177 (original so' no backup) e o retorno do backup TAMBEM falha: erro pt-BR nomeando os dois caminhos, nada apagado
    public void Replace_1177_RetornoDoBackupFalha_ErroNomeiaOsDoisArquivos()
    {
        var p = Pasta();
        var destino = Path.Combine(p, "a.docx"); File.WriteAllText(destino, "velho");
        var tmp = destino + ".moffice-tmp"; File.WriteAllText(tmp, "novo");
        var bak = destino + GravacaoNoDestino.SufixoBackup;
        var log = new List<string>();
        var e = Assert.Throws<ErroGravacaoParcial>(() => GravacaoNoDestino.Substituir(tmp, destino, log.Add, replace: (_, alvo, b) =>
        {
            File.Move(alvo, b);
            Directory.CreateDirectory(alvo); // um diretorio no lugar do destino: File.Exists(destino) = false e o Move(bak, destino) falha
            throw new IOException("1177 simulado");
        }));
        Assert.Equal($"O arquivo original ficou em {bak}. A versão nova continua aberta no mOffice: salve de novo ou use Salvar como.", e.Message);
        Assert.DoesNotContain(".moffice-tmp", e.Message); // o temporario e' apagado no proximo salvamento: nao se aponta o usuario para ele
        Assert.Equal("velho", File.ReadAllText(bak));
        Assert.Equal("novo", File.ReadAllText(tmp));
        Assert.Contains(log, l => l.Contains(bak) && l.Contains(tmp));
    }
}
