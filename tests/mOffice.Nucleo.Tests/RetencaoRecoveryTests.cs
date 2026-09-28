using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Retencao da pasta de recovery (achado da revisao final 2026-09-17). Cada documento aberto cria
/// <c>%TEMP%\mOffice\recovery\&lt;guid&gt;</c> com o <c>Editor.bin</c> — o documento INTEIRO, em claro — e nada
/// apagava: 353 pastas orfas medidas nesta maquina. A politica tem duas metades: apagar ao fechar a janela
/// (provado em <see cref="AbrirDocumentoTests"/>, que e' teste de janela real) e varrer na subida do app o que
/// ficou para tras de uma execucao anterior (queda, kill do processo) — que e' o que este arquivo prova, sem
/// janela nenhuma.
/// </summary>
public class RetencaoRecoveryTests
{
    [Fact] // a antiga sai, a recente FICA. O par e' o teste: uma varredura que apagasse tudo passaria no
           // primeiro assert sozinho, e ela apagaria a pasta do documento aberto no momento.
    public void VarrerRecoveryAntigo_ApagaSoAsMaisVelhasQueOLimite()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-varredura-").FullName;
        try
        {
            var antiga = Directory.CreateDirectory(Path.Combine(raiz, "antiga"));
            var recente = Directory.CreateDirectory(Path.Combine(raiz, "recente"));
            // Conteudo de verdade nas duas: a varredura apaga pasta com arquivo dentro (Directory.Delete
            // recursivo), e sem isto o teste nao distinguiria "apagou" de "a pasta estava vazia".
            var binAntigo = Path.Combine(antiga.FullName, "Editor.bin");
            File.WriteAllText(binAntigo, "bin antigo");
            File.WriteAllText(Path.Combine(recente.FullName, "Editor.bin"), "bin recente");
            // O frescor conta pasta E arquivos dentro dela (achado 2 da revisao final, fix review 2): as duas
            // datas tem de voltar 3 dias, senao o Editor.bin recem-gravado (mtime = agora) sozinho manteria a
            // pasta "recente" mesmo com a pasta datada para tras.
            var dataAntiga = DateTime.UtcNow.AddDays(-3);
            File.SetLastWriteTimeUtc(binAntigo, dataAntiga);
            Directory.SetLastWriteTimeUtc(antiga.FullName, dataAntiga);

            var apagadas = SessaoDocumento.VarrerRecoveryAntigo(raiz, TimeSpan.FromDays(1));

            Assert.Equal(1, apagadas);
            Assert.False(Directory.Exists(antiga.FullName), "a pasta de 3 dias deveria ter sido apagada");
            Assert.True(Directory.Exists(recente.FullName), "a pasta recente NAO pode ser apagada — pode ser a do documento aberto agora");
        }
        finally { try { Directory.Delete(raiz, true); } catch { } }
    }

    [Fact] // primeira execucao da maquina: a raiz nem existe. Nao pode lancar na subida do app.
    public void VarrerRecoveryAntigo_RaizInexistente_NaoLancaEDevolveZero()
        => Assert.Equal(0, SessaoDocumento.VarrerRecoveryAntigo(
            Path.Combine(Path.GetTempPath(), "moffice-raiz-que-nao-existe-" + Guid.NewGuid().ToString("N")), TimeSpan.FromDays(1)));

    // -----------------------------------------------------------------------------------------------------
    // Achado 2 da revisao final (fix review 2): LastWriteTime da PASTA nao muda quando um arquivo DENTRO dela
    // e' regravado (so' quando um arquivo e' criado/apagado) — medido apendando changes/changes0.json e
    // regravando Editor.bin de uma pasta ja' criada. A pasta de um documento aberto ha' mais de idadeMinima
    // numa OUTRA instancia do app, mas com escrita recente dentro, seria apagada por baixo.
    // -----------------------------------------------------------------------------------------------------

    [Fact] // a pasta e' VELHA (mtime de 3 dias atras), mas o changes0.json de dentro acabou de ser gravado —
           // a varredura tem de olhar para o arquivo, nao so' para a pasta, e NAO apagar.
    public void VarrerRecoveryAntigo_PastaVelhaComArquivoRecenteDentro_NaoApaga()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-varredura-").FullName;
        try
        {
            var pasta = Directory.CreateDirectory(Path.Combine(raiz, "em-uso-em-outra-instancia"));
            var changes = Directory.CreateDirectory(Path.Combine(pasta.FullName, "changes"));
            File.WriteAllText(Path.Combine(pasta.FullName, "Editor.bin"), "bin");
            // A pasta (e o Editor.bin criado dentro dela) ficam com mtime de 3 dias atras...
            Directory.SetLastWriteTimeUtc(pasta.FullName, DateTime.UtcNow.AddDays(-3));
            File.SetLastWriteTimeUtc(Path.Combine(pasta.FullName, "Editor.bin"), DateTime.UtcNow.AddDays(-3));
            // ...mas o changes0.json acabou de ser apendado pelo editor, na sessao ainda aberta em outro
            // processo — sem TOCAR na data da pasta (e' exatamente isto que o achado mediu).
            File.WriteAllText(Path.Combine(changes.FullName, "changes0.json"), "\"c1\",");

            var apagadas = SessaoDocumento.VarrerRecoveryAntigo(raiz, TimeSpan.FromDays(1));

            Assert.Equal(0, apagadas);
            Assert.True(Directory.Exists(pasta.FullName), "pasta com escrita recente DENTRO nao pode ser apagada, mesmo com a pasta em si datada");
        }
        finally { try { Directory.Delete(raiz, true); } catch { } }
    }

    [Fact] // controle positivo do caso acima: a MESMA forma de pasta, mas com TODOS os arquivos velhos (pasta
           // E conteudo) — prova que o "nao apaga" de cima veio do arquivo recente, e nao de alguma outra
           // condicao acidental (ex.: SearchOption.AllDirectories nao descendo em changes/).
    public void VarrerRecoveryAntigo_PastaEArquivosTodosVelhos_Apaga()
    {
        var raiz = Directory.CreateTempSubdirectory("moffice-varredura-").FullName;
        try
        {
            var pasta = Directory.CreateDirectory(Path.Combine(raiz, "abandonada"));
            var changes = Directory.CreateDirectory(Path.Combine(pasta.FullName, "changes"));
            var bin = Path.Combine(pasta.FullName, "Editor.bin");
            var changesJson = Path.Combine(changes.FullName, "changes0.json");
            File.WriteAllText(bin, "bin");
            File.WriteAllText(changesJson, "\"c1\",");
            var velho = DateTime.UtcNow.AddDays(-3);
            File.SetLastWriteTimeUtc(bin, velho);
            File.SetLastWriteTimeUtc(changesJson, velho);
            Directory.SetLastWriteTimeUtc(changes.FullName, velho);
            Directory.SetLastWriteTimeUtc(pasta.FullName, velho);

            var apagadas = SessaoDocumento.VarrerRecoveryAntigo(raiz, TimeSpan.FromDays(1));

            Assert.Equal(1, apagadas);
            Assert.False(Directory.Exists(pasta.FullName), "pasta e conteudo todos velhos deveriam ter sido apagados");
        }
        finally { try { Directory.Delete(raiz, true); } catch { } }
    }

    // -----------------------------------------------------------------------------------------------------
    // Achado 3 da revisao final (fix review 2): a RAIZ da varredura sendo um link simbolico/juncao para outro
    // lugar do disco. Seguir o link levaria a varredura a apagar conteudo do outro lado dele. Testar com uma
    // juncao de verdade e' incomodo (mklink /j nao tem API direta no .NET); um symlink de diretorio, criado
    // com Directory.CreateSymbolicLink, prova a mesma defesa (LinkTarget/ReparsePoint), sem depender de p/invoke.
    // -----------------------------------------------------------------------------------------------------

    [SkippableFact] // criar link simbolico de diretorio no Windows exige SeCreateSymbolicLinkPrivilege (admin)
                     // ou Modo de Desenvolvedor ligado; sem isto o teste e' pulado, nao falha.
    public void VarrerRecoveryAntigo_RaizELinkSimbolico_RecusaSemApagarNada()
    {
        var alvo = Directory.CreateTempSubdirectory("moffice-varredura-alvo-").FullName;
        var pastaDoLink = Directory.CreateTempSubdirectory("moffice-varredura-link-").FullName;
        var link = Path.Combine(pastaDoLink, "link-para-o-alvo");
        try
        {
            var antiga = Directory.CreateDirectory(Path.Combine(alvo, "antiga"));
            File.WriteAllText(Path.Combine(antiga.FullName, "Editor.bin"), "bin antigo");
            Directory.SetLastWriteTimeUtc(antiga.FullName, DateTime.UtcNow.AddDays(-3));
            File.SetLastWriteTimeUtc(Path.Combine(antiga.FullName, "Editor.bin"), DateTime.UtcNow.AddDays(-3));

            string? semPrivilegio = null;
            try { Directory.CreateSymbolicLink(link, alvo); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                semPrivilegio = "sem privilegio para criar link simbolico de diretorio nesta conta "
                    + "(precisa admin ou Modo de Desenvolvedor): " + e.Message;
            }
            Skip.If(semPrivilegio is not null, semPrivilegio!);

            var apagadas = SessaoDocumento.VarrerRecoveryAntigo(link, TimeSpan.FromDays(1));

            Assert.Equal(0, apagadas);
            Assert.True(Directory.Exists(antiga.FullName), "a raiz sendo link simbolico nao pode apagar nada do outro lado dele");
        }
        finally
        {
            try { Directory.Delete(link); } catch { }
            try { Directory.Delete(pastaDoLink, true); } catch { }
            try { Directory.Delete(alvo, true); } catch { }
        }
    }

    [Fact] // ApagarRecovery e' melhor esforco: chamada com a pasta ja' ausente nao pode lancar (roda no
           // fechamento da janela, onde uma excecao nao tem para onde ir).
    public void ApagarRecovery_PastaAusente_NaoLanca()
    {
        var s = new SessaoDocumento
        {
            ArquivoOrigem = @"C:\nada\x.docx",
            Formato = FormatoDocumento.Docx,
            PastaRecovery = Path.Combine(Path.GetTempPath(), "moffice-rec-ausente-" + Guid.NewGuid().ToString("N")),
        };
        s.ApagarRecovery();
        Assert.False(Directory.Exists(s.PastaRecovery));
    }

    [Fact] // controle positivo do ApagarRecovery: com conteudo dentro, some tudo.
    public void ApagarRecovery_ComConteudo_ApagaAPastaInteira()
    {
        var pasta = Directory.CreateTempSubdirectory("moffice-rec-").FullName;
        Directory.CreateDirectory(Path.Combine(pasta, "changes"));
        File.WriteAllText(Path.Combine(pasta, "Editor.bin"), "documento em claro");
        File.WriteAllText(Path.Combine(pasta, "changes", "changes0.json"), "\"c1\",");
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\nada\x.docx", Formato = FormatoDocumento.Docx, PastaRecovery = pasta };

        s.ApagarRecovery();

        Assert.False(Directory.Exists(pasta));
    }
}
