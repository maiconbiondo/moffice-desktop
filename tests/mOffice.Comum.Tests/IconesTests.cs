using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace mOffice.Comum.Tests;

public class IconesTests
{
    [Theory] // cada .ico tem 4 quadros PNG (256/48/32/16), como o mpdf.ico
    [InlineData("mTexto", "mtexto.ico")] [InlineData("mPlanilha", "mplanilha.ico")] [InlineData("mSlides", "mslides.ico")]
    public void Ico_TemQuatroQuadrosPng(string app, string ico)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AplicativoOfficeTests.RaizRepo(), "src", app, "Assets", ico));
        Assert.Equal(0, BitConverter.ToUInt16(bytes, 0)); Assert.Equal(1, BitConverter.ToUInt16(bytes, 2)); // reservado, tipo=icone
        var quadros = BitConverter.ToUInt16(bytes, 4);
        Assert.Equal(4, quadros);
        var tamanhos = new List<int>();
        for (var i = 0; i < quadros; i++)
        {
            var e = 6 + i * 16;
            var largura = bytes[e] == 0 ? 256 : bytes[e];
            var offset = BitConverter.ToInt32(bytes, e + 12);
            Assert.Equal(0x89, bytes[offset]); Assert.Equal((byte)'P', bytes[offset + 1]); // PNG cru
            tamanhos.Add(largura);
        }
        Assert.Equal([256, 48, 32, 16], tamanhos);
    }

    /// <summary>
    /// Regeneracao reprodutivel: rodar tools/icones/gerar-icones.ps1 devolve o MESMO desenho que esta'
    /// commitado. E' o que prova que o .ico e' artefato derivado do tema (mOfficeTema.xaml) e nao um binario
    /// solto que ninguem sabe mais refazer — e o que pega um logo editado no tema sem os icones regerados.
    ///
    /// A comparacao e' de PIXEL, nao de byte (achado da revisao 1): os bytes do PNG dependem da versao do
    /// Pillow (tabela de Huffman, nivel de compressao do `optimize=True`), que nao esta' fixada em lugar
    /// nenhum — um `pip install -U pillow` quebraria o teste sem nada ter mudado no desenho. O criterio e'
    /// o que o USUARIO ve': mesmo numero de quadros, mesmas dimensoes e ao menos 99 % dos pixels iguais
    /// (a folga de 1 % cobre diferenca de arredondamento no antialiasing entre versoes da biblioteca).
    /// </summary>
    [SkippableFact]
    public void Ico_RegeneracaoReprodutivel()
    {
        var raiz = AplicativoOfficeTests.RaizRepo();
        Skip.IfNot(TemPillow(), "Python com Pillow >= 10 nao disponivel (pip install 'pillow>=10')");
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        try
        {
            var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("powershell",
                $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(raiz, "tools", "icones", "gerar-icones.ps1")}\" -Saida \"{pasta}\"") { UseShellExecute = false })!;
            // Com teto: um script que trava (um `python` que abre um prompt, por exemplo) penduraria a
            // suite inteira em vez de falhar.
            if (!p.WaitForExit(120000))
            {
                // Mata ANTES de falhar: um powershell/python pendurado sobreviveria a' suite inteira e
                // ainda seguraria a pasta temporaria que o finally tenta apagar.
                try { p.Kill(entireProcessTree: true); } catch (Exception) { }
                Assert.Fail("gerar-icones.ps1 nao terminou em 120s");
            }
            Assert.Equal(0, p.ExitCode);

            foreach (var (app, ico) in new[] { ("mTexto", "mtexto.ico"), ("mPlanilha", "mplanilha.ico"), ("mSlides", "mslides.ico") })
            {
                var commitado = Quadros(File.ReadAllBytes(Path.Combine(raiz, "src", app, "Assets", ico)));
                var gerado = Quadros(File.ReadAllBytes(Path.Combine(pasta, ico)));
                Assert.Equal(commitado.Count, gerado.Count);
                for (var i = 0; i < commitado.Count; i++)
                {
                    var (lc, ac, pc) = commitado[i];
                    var (lg, ag, pg) = gerado[i];
                    Assert.Equal((lc, ac), (lg, ag));
                    var total = lc * ac;
                    var iguais = 0;
                    for (var k = 0; k < total; k++)
                        if (pc[k * 4] == pg[k * 4] && pc[k * 4 + 1] == pg[k * 4 + 1]
                            && pc[k * 4 + 2] == pg[k * 4 + 2] && pc[k * 4 + 3] == pg[k * 4 + 3]) iguais++;
                    Assert.True(iguais * 100.0 / total >= 99.0,
                        $"{ico} quadro {lc}px: só {iguais * 100.0 / total:0.00}% dos pixels batem com o commitado");
                }
            }
        }
        finally { try { Directory.Delete(pasta, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }

    /// <summary>Decodifica os quadros PNG de dentro de um .ico para (largura, altura, pixels BGRA). Usa o
    /// decodificador do WPF: e' DECODIFICACAO (WIC, so' CPU), nao render — nao depende de sessao grafica,
    /// ao contrario do RenderTargetBitmap (ver tools/icones/render-logos.py).</summary>
    private static List<(int Largura, int Altura, byte[] Pixels)> Quadros(byte[] ico)
    {
        var lista = new List<(int, int, byte[])>();
        var n = BitConverter.ToUInt16(ico, 4);
        for (var i = 0; i < n; i++)
        {
            var e = 6 + i * 16;
            var tamanho = BitConverter.ToInt32(ico, e + 8);
            var offset = BitConverter.ToInt32(ico, e + 12);
            using var ms = new MemoryStream(ico, offset, tamanho);
            var dec = new PngBitmapDecoder(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
            var conv = new FormatConvertedBitmap(dec.Frames[0], PixelFormats.Bgra32, null, 0);
            var px = new byte[conv.PixelWidth * conv.PixelHeight * 4];
            conv.CopyPixels(px, conv.PixelWidth * 4, 0);
            lista.Add((conv.PixelWidth, conv.PixelHeight, px));
        }
        return lista;
    }

    private static bool TemPillow()
    {
        try
        {
            var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("python",
                "-c \"import PIL,sys; sys.exit(0 if int(PIL.__version__.split('.')[0])>=10 else 1)\"")
            { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true })!;
            return p.WaitForExit(30000) && p.ExitCode == 0;
        }
        catch (Exception e) when (e is System.ComponentModel.Win32Exception or InvalidOperationException) { return false; }
    }
}
