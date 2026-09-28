using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace mOffice.Nucleo.Tests;

/// <summary>Leitura de pixels de uma captura PNG da janela (<see cref="JanelaEditor.CapturarPngAsync"/>).
/// E' o unico jeito de provar que uma FIGURA aparece na tela: a API do editor diz que o desenho existe no
/// modelo, mas nao diz que os bytes da imagem chegaram (na Task 5A eles nao chegavam — `file://` recusado como
/// subrecurso de uma pagina `https://` — e o documento abria com o quadro da figura VAZIO).</summary>
public static class Pixels
{
    /// <summary>Quantos pixels "vermelho forte" (R &gt; 200, G &lt; 60, B &lt; 60) a captura tem. O criterio e'
    /// folgado de proposito: a figura e' reescalada pelo zoom do editor e reamostrada na captura, entao as
    /// bordas ficam misturadas com o branco da pagina — o miolo continua vermelho puro.</summary>
    public static int ContarVermelhos(byte[] png)
    {
        using var ms = new MemoryStream(png);
        var quadro = BitmapFrame.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
        var convertido = new FormatConvertedBitmap(quadro, PixelFormats.Bgra32, null, 0);
        var stride = convertido.PixelWidth * 4;
        var buffer = new byte[stride * convertido.PixelHeight];
        convertido.CopyPixels(buffer, stride, 0);
        var n = 0;
        for (var i = 0; i + 3 < buffer.Length; i += 4)
        {
            byte b = buffer[i], g = buffer[i + 1], r = buffer[i + 2];
            if (r > 200 && g < 60 && b < 60) n++;
        }
        return n;
    }

    public static bool ContemBlocoVermelho(byte[] png, int minimoPixels) => ContarVermelhos(png) >= minimoPixels;

    // -----------------------------------------------------------------------------------------------------
    // Comparacao de duas capturas (Task 1 do Marco 1, fonte de simbolos): o teste de janela real prova que os
    // pixels das marcas de formatacao vem da NOSSA fonte comparando a captura normal com a de um motor de
    // controle cuja fonte tem contornos vazios — o que difere e', por construcao, a tinta dos nossos glifos.
    // -----------------------------------------------------------------------------------------------------

    /// <summary>Captura decodificada em BGRA (4 bytes por pixel).</summary>
    public sealed record Imagem(int Largura, int Altura, byte[] Bgra);

    public static Imagem Decodificar(byte[] png)
    {
        using var ms = new MemoryStream(png);
        var quadro = BitmapFrame.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
        var convertido = new FormatConvertedBitmap(quadro, PixelFormats.Bgra32, null, 0);
        var stride = convertido.PixelWidth * 4;
        var buffer = new byte[stride * convertido.PixelHeight];
        convertido.CopyPixels(buffer, stride, 0);
        return new Imagem(convertido.PixelWidth, convertido.PixelHeight, buffer);
    }

    /// <summary>Indices (linha * largura + coluna) dos pixels em que algum canal difere mais que
    /// <paramref name="limiar"/> entre as duas imagens (mesmo tamanho), a partir da linha
    /// <paramref name="linhaInicial"/> — para deixar a barra de ferramentas (cujo estado de botao varia)
    /// fora da comparacao e olhar so' a pagina.</summary>
    public static List<int> Diferentes(Imagem a, Imagem b, int limiar = 40, int linhaInicial = 0)
    {
        if (a.Largura != b.Largura || a.Altura != b.Altura) throw new ArgumentException("imagens de tamanhos diferentes");
        var lista = new List<int>();
        for (var i = linhaInicial * a.Largura; i < a.Largura * a.Altura; i++)
        {
            var o = i * 4;
            if (Math.Abs(a.Bgra[o] - b.Bgra[o]) > limiar || Math.Abs(a.Bgra[o + 1] - b.Bgra[o + 1]) > limiar || Math.Abs(a.Bgra[o + 2] - b.Bgra[o + 2]) > limiar)
                lista.Add(i);
        }
        return lista;
    }

    /// <summary>Pixel "com tinta": luminancia abaixo de 128 (texto e marcas sao pretos sobre a pagina branca).</summary>
    public static bool Escuro(Imagem img, int indice)
    {
        var o = indice * 4;
        var luma = (img.Bgra[o + 2] * 299 + img.Bgra[o + 1] * 587 + img.Bgra[o] * 114) / 1000;
        return luma < 128;
    }

    /// <summary>PNG do mesmo tamanho de <paramref name="a"/>: branco, com os pixels de <paramref name="indices"/>
    /// em vermelho — a "mascara" do que mudou, para o relatorio.</summary>
    public static byte[] MascaraPng(Imagem a, IEnumerable<int> indices)
    {
        var stride = a.Largura * 4;
        var buffer = new byte[stride * a.Altura];
        Array.Fill(buffer, (byte)255);
        foreach (var i in indices) { buffer[i * 4] = 0; buffer[i * 4 + 1] = 0; buffer[i * 4 + 2] = 255; buffer[i * 4 + 3] = 255; }
        var bmp = BitmapSource.Create(a.Largura, a.Altura, 96, 96, PixelFormats.Bgra32, null, buffer, stride);
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(bmp));
        using var saida = new MemoryStream();
        enc.Save(saida);
        return saida.ToArray();
    }
}
