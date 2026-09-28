using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

/// <summary>M3 da revisao final do Plano 3: o XML de parametros do x2t (que pode levar a senha do documento em claro)
/// nasce DENTRO da pasta temporaria da chamada — nunca na raiz do %TEMP% — e a subida do app varre os
/// <c>moffice-x2t-*.xml</c> que versoes anteriores deixaram la'.</summary>
public class X2tParametrosTests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("mOffice-testes-x2tparam-").FullName;

    public void Dispose() { try { Directory.Delete(_pasta, true); } catch { } }

    [Fact] // conversao real: nenhum moffice-x2t-*.xml aparece na raiz do %TEMP%; o XML aparece (e some) dentro da pasta
           // temporaria da chamada (controle positivo: o observador enxerga a criacao do XML)
    public async Task Conversao_XmlDeParametrosNasceNaPastaDaChamada()
    {
        var pastaTempPai = Path.Combine(_pasta, "tmp-x2t");
        Directory.CreateDirectory(pastaTempPai);
        var conv = new X2tConversor(MotorDeTeste.Localizar() with { PastaTemp = pastaTempPai });
        var naRaiz = new List<string>();
        var naPastaDaChamada = new List<string>();
        using var raiz = new FileSystemWatcher(Path.GetTempPath(), X2tConversor.PrefixoXmlAntigo + "*.xml") { IncludeSubdirectories = false };
        using var chamada = new FileSystemWatcher(pastaTempPai, "*.xml") { IncludeSubdirectories = true };
        raiz.Created += (_, e) => { lock (naRaiz) naRaiz.Add(e.FullPath); };
        chamada.Created += (_, e) => { lock (naPastaDaChamada) naPastaDaChamada.Add(e.FullPath); };
        raiz.EnableRaisingEvents = true;
        chamada.EnableRaisingEvents = true;

        var origem = FabricaDocumentos.Docx(_pasta, X2tConversorTests.Frase);
        await conv.ParaBinAsync(origem, Path.Combine(_pasta, "rec"));
        await conv.DeBinAsync(Path.Combine(_pasta, "rec"), FormatoDocumento.Docx, Path.Combine(_pasta, "volta.docx"));
        await Task.Delay(500);   // os eventos do observador chegam numa thread propria

        lock (naRaiz) Assert.Empty(naRaiz);
        lock (naPastaDaChamada)
        {
            Assert.Equal(2, naPastaDaChamada.Count(p => Path.GetFileName(p) == "parametros.xml"));
            // Cada XML nasceu numa subpasta da chamada (a <guid> de CriarPastaTemp), nunca na pasta-pai.
            Assert.All(naPastaDaChamada, p => Assert.NotEqual(Path.GetFullPath(pastaTempPai).TrimEnd('\\'), Path.GetDirectoryName(Path.GetFullPath(p))));
        }
        // ...e nao sobrou nada: a pasta da chamada (com o XML dentro) foi apagada no finally.
        Assert.Empty(Directory.GetFileSystemEntries(pastaTempPai));
    }

    [Fact] // varredura da subida: so' os moffice-x2t-*.xml VELHOS do nivel de cima saem; o recente (conversao em
           // andamento de outra instancia), outro nome, outra extensao e subpasta ficam
    public void VarrerParametrosOrfaos_SoOsVelhosDoPadrao()
    {
        var velho = Criar("moffice-x2t-" + Guid.NewGuid().ToString("N") + ".xml", horas: -2);
        var recente = Criar("moffice-x2t-" + Guid.NewGuid().ToString("N") + ".xml", horas: 0);
        var outroNome = Criar("outro-" + Guid.NewGuid().ToString("N") + ".xml", horas: -2);
        var outraExtensao = Criar("moffice-x2t-" + Guid.NewGuid().ToString("N") + ".xmlx", horas: -2);
        var sub = Directory.CreateDirectory(Path.Combine(_pasta, "sub")).FullName;
        var naSubpasta = Path.Combine(sub, "moffice-x2t-sub.xml");
        File.WriteAllText(naSubpasta, "<x/>"); File.SetLastWriteTimeUtc(naSubpasta, DateTime.UtcNow.AddHours(-2));

        Assert.Equal(1, X2tConversor.VarrerParametrosOrfaos(_pasta, TimeSpan.FromHours(1)));
        Assert.False(File.Exists(velho));
        Assert.True(File.Exists(recente));
        Assert.True(File.Exists(outroNome));
        Assert.True(File.Exists(outraExtensao));
        Assert.True(File.Exists(naSubpasta));
        // Pasta inexistente: nao lanca.
        Assert.Equal(0, X2tConversor.VarrerParametrosOrfaos(Path.Combine(_pasta, "nao-existe"), TimeSpan.FromHours(1)));
    }

    [Fact] // fix round 2: a subida tambem varre as pastas por chamada que um processo morto deixou em PastaTemp (la' dentro
           // fica o XML de parametros de hoje) — so' as velhas; a de uma conversao em andamento (recente) fica
    public void VarrerPastasTemporarias_SoAsVelhas()
    {
        var pai = Directory.CreateDirectory(Path.Combine(_pasta, "x2t")).FullName;
        var velha = Directory.CreateDirectory(Path.Combine(pai, Guid.NewGuid().ToString("N"))).FullName;
        File.WriteAllText(Path.Combine(velha, "parametros.xml"), "<m_sPassword>segredo</m_sPassword>");
        File.SetLastWriteTimeUtc(Path.Combine(velha, "parametros.xml"), DateTime.UtcNow.AddHours(-3));
        Directory.SetLastWriteTimeUtc(velha, DateTime.UtcNow.AddHours(-3));
        var recente = Directory.CreateDirectory(Path.Combine(pai, Guid.NewGuid().ToString("N"))).FullName;
        File.WriteAllText(Path.Combine(recente, "parametros.xml"), "<x/>");
        // Pasta velha com um arquivo RECENTE dentro (conversao longa ainda escrevendo): fica.
        var velhaComArquivoNovo = Directory.CreateDirectory(Path.Combine(pai, Guid.NewGuid().ToString("N"))).FullName;
        File.WriteAllText(Path.Combine(velhaComArquivoNovo, "saida.bin"), "x");
        Directory.SetLastWriteTimeUtc(velhaComArquivoNovo, DateTime.UtcNow.AddHours(-3));

        Assert.Equal(1, X2tConversor.VarrerPastasTemporarias(pai, TimeSpan.FromHours(1)));
        Assert.False(Directory.Exists(velha));
        Assert.True(Directory.Exists(recente));
        Assert.True(Directory.Exists(velhaComArquivoNovo));
        Assert.Equal(0, X2tConversor.VarrerPastasTemporarias(Path.Combine(_pasta, "nao-existe"), TimeSpan.FromHours(1)));
        Assert.Equal(Path.Combine(Path.GetTempPath(), "mOffice", "x2t"), ConfiguracaoMotor.PastaTempPadrao);
        Assert.Equal(ConfiguracaoMotor.PastaTempPadrao, MotorDeTeste.Localizar().PastaTemp);
    }

    private string Criar(string nome, int horas)
    {
        var p = Path.Combine(_pasta, nome);
        File.WriteAllText(p, "<TaskQueueDataConvert><m_sPassword>segredo</m_sPassword></TaskQueueDataConvert>");
        File.SetLastWriteTimeUtc(p, DateTime.UtcNow.AddHours(horas));
        return p;
    }
}
