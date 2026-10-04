namespace mOffice.Nucleo.Tests;

public class ModelosDisponiveisTests
{
    [Fact] // nome dos modelos do motor: "[32]" + base32 do nome
    public void NomeDoMotor_DecodificaBase32() =>
        Assert.Equal("Ordem de compra", ModelosDisponiveis.NomeDoMotor(@"C:\m\Documents\[32]J5ZGIZLNEBSGKIDDN5WXA4TB.dotx"));

    [Fact] // com padding, com acento (UTF-8), sem o prefixo e com base32 invalido (nao lanca)
    public void NomeDoMotor_PaddingAcentoEFormatosInvalidos()
    {
        Assert.Equal("Vale-Presente", ModelosDisponiveis.NomeDoMotor(@"[32]KZQWYZJNKBZGK43FNZ2GK===.dotx"));
        Assert.Equal("Procedimento Operacional Padrão", ModelosDisponiveis.NomeDoMotor(@"[32]KBZG6Y3FMRUW2ZLOORXSAT3QMVZGCY3JN5XGC3BAKBQWI4WDUNXQ====.dotx"));
        Assert.Equal("arquivo comum", ModelosDisponiveis.NomeDoMotor(@"C:\x\arquivo comum.dotx"));
        Assert.Equal("[32]!!!", ModelosDisponiveis.NomeDoMotor(@"[32]!!!.dotx"));
    }

    [Fact] // do usuario: so' modelos do tipo, sem Normal.dotm nem ~$; do usuario primeiro, depois os do motor
    public void Listar_UsuarioPrimeiroSemNormal()
    {
        var u = Directory.CreateTempSubdirectory("moffice-mod-u-").FullName; var m = Directory.CreateTempSubdirectory("moffice-mod-m-").FullName;
        try
        {
            foreach (var f in new[] { "Peticao inicial.dotx", "Normal.dotm", "~$ticao.dotx", "Custos.xltx", "nota.txt", "Contrato.dotm" }) File.WriteAllText(Path.Combine(u, f), "x");
            Directory.CreateDirectory(Path.Combine(u, "sub")); File.WriteAllText(Path.Combine(u, "sub", "Escondido.dotx"), "x"); // so' o nivel de cima
            Directory.CreateDirectory(Path.Combine(m, "Documents"));
            File.WriteAllText(Path.Combine(m, "Documents", "[32]J5ZGIZLNEBSGKIDDN5WXA4TB.dotx"), "x");
            var l = ModelosDisponiveis.Listar("word", u, m);
            Assert.Equal([("Contrato", true), ("Peticao inicial", true), ("Ordem de compra", false)], l.Select(x => (x.Nome, x.DoUsuario)));
            Assert.Equal([("Custos", true)], ModelosDisponiveis.Listar("cell", u, m).Select(x => (x.Nome, x.DoUsuario)));
        }
        finally { Directory.Delete(u, true); Directory.Delete(m, true); }
    }

    [Fact] // pasta do usuario inexistente (rede fora do ar): lista so' os do motor, sem lancar
    public void Listar_PastaDoUsuarioInexistente() =>
        Assert.Empty(ModelosDisponiveis.Listar("cell", @"Z:\nao\existe", Path.Combine(Path.GetTempPath(), "nao-existe-tambem")));

    [Fact] // sem pasta do usuario (null) e tipo desconhecido
    public void Listar_SemPastaOuTipoDesconhecido()
    {
        Assert.Empty(ModelosDisponiveis.Listar("word", null, Path.Combine(Path.GetTempPath(), "nao-existe-mesmo")));
        Assert.Empty(ModelosDisponiveis.Listar("visio", null, Path.GetTempPath()));
    }

    [SkippableFact] // o motor real: 10 + 6 + 10 modelos (Forms fica de fora: sao PDFs de formulario)
    public void Listar_MotorReal()
    {
        var pt = Path.Combine(FronteiraTests.RaizRepo(), "third_party", "onlyoffice", "converter", "templates", "PT");
        Skip.If(!Directory.Exists(pt), "motor nao obtido");
        Assert.Equal(10, ModelosDisponiveis.Listar("word", null, pt).Count);
        Assert.Equal(6, ModelosDisponiveis.Listar("cell", null, pt).Count);
        Assert.Equal(10, ModelosDisponiveis.Listar("slide", null, pt).Count);
        Assert.Contains("Ordem de compra", ModelosDisponiveis.Listar("word", null, pt).Select(x => x.Nome));
    }
}
