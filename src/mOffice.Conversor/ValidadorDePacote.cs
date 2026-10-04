using System.IO.Compression;
using System.Text;

namespace mOffice.Conversor;

/// <summary>
/// Confere o arquivo que o x2t acabou de gravar ANTES de ele substituir o do usuario (FMT-10, auditoria 2026-10-01):
/// o x2t 9.4 devolve codigo 0 com pacote truncado quando o m_sTempDir passa de ~220 caracteres (MEDIDO: 10 de 12
/// partes; zip vazio de 22 bytes), e o "codigo 0 e tamanho &gt; 0" de antes deixava isso passar por cima do original.
/// So' estrutura minima por familia (<see cref="InfoExtensao.Familia"/>, da <see cref="TabelaFormatos"/>) — nada de
/// validar conteudo: OOXML e ODF pela parte principal do zip, PDF pelo cabecalho e pelo fim, RTF e HTML pela marca do
/// comeco. Formato que o x2t nao grava ("binario": .doc/.xls/.ppt/.xlsb) e' RECUSADO: so' chegaria aqui por bug de quem chamou.
/// </summary>
public static class ValidadorDePacote
{
    private static readonly byte[] AssinaturaOle = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    /// <summary>Quantos bytes do comeco do arquivo as familias de texto (RTF/HTML) examinam.</summary>
    private const int BytesDeMarca = 4096;

    public static void Validar(string arquivo, FormatoDocumento formato, bool cifrado = false)
    {
        var info = new FileInfo(arquivo);
        if (!info.Exists || info.Length == 0) throw Falha(arquivo, "arquivo vazio");
        var cabeca = LerInicio(arquivo, 8);
        // Documento com senha sai como conteiner OLE (d0cf11e0, medido no Plano 3) — nao e' zip.
        if (cifrado && cabeca.AsSpan().SequenceEqual(AssinaturaOle)) return;
        switch (Familia(formato))
        {
            case "ooxml": ExigirZip(arquivo, "[Content_Types].xml", ParteOoxml(formato)); break;
            case "odf": ExigirZip(arquivo, "mimetype", "content.xml"); break;
            case "pdf":
                if (cabeca.Length < 5 || Encoding.ASCII.GetString(cabeca, 0, 5) != "%PDF-") throw Falha(arquivo, "sem o cabeçalho %PDF-");
                if (!FimContem(arquivo, "%%EOF")) throw Falha(arquivo, "sem o fim %%EOF");
                break;
            case "rtf":
                if (!InicioContem(arquivo, "{\\rtf", aposEspacos: true)) throw Falha(arquivo, "não começa com {\\rtf");
                break;
            case "texto":
                // HTML: nao vazio e com marca de HTML. TXT/CSV (Plano 6, Task 3): texto de verdade — nao um zip/OLE (o x2t gravando o
                // formato errado) nem binario/UTF-16 por engano (NUL nos primeiros bytes; so' UTF-8 e 1252 sao gravados). Nada alem disso:
                // o conteudo e' livre e a codificacao e' a que o usuario escolheu.
                if (formato == FormatoDocumento.Html && !TemMarcaDeHtml(arquivo)) throw Falha(arquivo, "sem a marca <html>");
                if (formato is FormatoDocumento.Txt or FormatoDocumento.Csv) ExigirTextoPuro(arquivo);
                break;
            default:
                throw Falha(arquivo, $"o formato {formato.Extensao().ToUpperInvariant()} não pode ser gravado pelo mOffice");
        }
    }

    /// <summary>Familia do formato para a validacao, da <see cref="TabelaFormatos"/> (primeira linha do formato).</summary>
    internal static string Familia(FormatoDocumento f) => TabelaFormatos.Principal(f).Familia;

    private static string ParteOoxml(FormatoDocumento f) => f.DocType() switch
    {
        "cell" => "xl/workbook.xml",
        "slide" => "ppt/presentation.xml",
        _ => "word/document.xml",
    };

    private static void ExigirZip(string arquivo, params string[] partes)
    {
        try
        {
            using var z = ZipFile.OpenRead(arquivo);
            foreach (var parte in partes)
            {
                var e = z.GetEntry(parte);
                if (e is null || e.Length == 0) throw Falha(arquivo, "falta a parte " + parte);
            }
        }
        catch (InvalidDataException) { throw Falha(arquivo, "pacote ZIP ilegível"); }
    }

    private static void ExigirTextoPuro(string arquivo)
    {
        var b = LerInicio(arquivo, BytesDeMarca);
        if (b.Length >= 4 && b[0] == 0x50 && b[1] == 0x4B && b[2] is 0x03 or 0x05 && b[3] is 0x04 or 0x06) throw Falha(arquivo, "é um pacote ZIP, não texto");
        if (b.AsSpan().IndexOf((byte)0) >= 0) throw Falha(arquivo, "contém bytes binários");
    }

    private static byte[] LerInicio(string arquivo, int n)
    {
        using var f = File.OpenRead(arquivo);
        var b = new byte[Math.Min(n, (int)Math.Min(int.MaxValue, f.Length))];
        f.ReadExactly(b);
        return b;
    }

    /// <summary>O comeco do arquivo (BOM e espacos em branco iniciais descartados) e' <paramref name="marca"/>.</summary>
    private static bool InicioContem(string arquivo, string marca, bool aposEspacos)
    {
        var s = TextoInicial(arquivo);
        if (aposEspacos) s = s.TrimStart();
        return s.StartsWith(marca, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TemMarcaDeHtml(string arquivo)
    {
        var s = TextoInicial(arquivo);
        return s.Contains("<html", StringComparison.OrdinalIgnoreCase) || s.Contains("<!doctype html", StringComparison.OrdinalIgnoreCase);
    }

    private static string TextoInicial(string arquivo)
    {
        var b = LerInicio(arquivo, BytesDeMarca);
        var inicio = b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF ? 3 : 0;
        return Encoding.UTF8.GetString(b, inicio, b.Length - inicio).TrimStart('﻿');
    }

    private static bool FimContem(string arquivo, string marca)
    {
        using var f = File.OpenRead(arquivo);
        var n = (int)Math.Min(1024, f.Length);
        f.Seek(-n, SeekOrigin.End);
        var b = new byte[n];
        f.ReadExactly(b);
        return Encoding.ASCII.GetString(b).Contains(marca, StringComparison.Ordinal);
    }

    private static ErroConversao Falha(string arquivo, string motivo) =>
        new($"O arquivo gerado para \"{Path.GetFileName(arquivo)}\" está incompleto ({motivo}); o original não foi alterado.", 0, "");
}
