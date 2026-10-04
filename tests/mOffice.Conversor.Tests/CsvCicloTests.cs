using System.Text;
using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

/// <summary>
/// Fix 2 da Task 3 (Plano 6): abrir -> salvar -> abrir -> salvar de um CSV UTF-8 que termina em quebra de linha ACUMULAVA uma linha so' de
/// delimitador a cada ciclo (o x2t le' a ultima quebra como linha vazia). O conserto e' na ABERTURA: o x2t le' uma copia sem UM terminador final
/// (<see cref="X2tConversor.CopiaSemTerminadorFinal"/>). x2t real, sem janela.
/// </summary>
public class CsvCicloTests : IDisposable
{
    private readonly string _pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", "csv-ciclo-" + Guid.NewGuid().ToString("N"));
    public CsvCicloTests() => Directory.CreateDirectory(_pasta);
    public void Dispose() { try { Directory.Delete(_pasta, true); } catch { } }

    private static readonly UTF8Encoding Utf8 = new(false, true);

    /// <summary>N ciclos abrir -> salvar (cada um abre a SAIDA do anterior); devolve os bytes de cada saida.</summary>
    private async Task<List<byte[]>> Ciclos(byte[] inicial, OpcoesTexto op, int n = 3)
    {
        var c = new X2tConversor(MotorDeTeste.Localizar());
        var atual = Path.Combine(_pasta, "in-" + Guid.NewGuid().ToString("N") + ".csv");
        File.WriteAllBytes(atual, inicial);
        var saidas = new List<byte[]>();
        for (var i = 1; i <= n; i++)
        {
            var rec = Path.Combine(_pasta, "rec-" + Guid.NewGuid().ToString("N"));
            await c.ParaBinAsync(atual, rec, texto: op);
            var saida = Path.Combine(_pasta, "out-" + Guid.NewGuid().ToString("N") + ".csv");
            await c.DeBinAsync(rec, FormatoDocumento.Csv, saida, texto: op);
            saidas.Add(File.ReadAllBytes(saida));
            atual = saida;
        }
        return saidas;
    }

    private static string Texto(byte[] b) => Utf8.GetString(b).TrimStart('﻿').Replace("\r\n", "\n");

    [Theory] // UTF-8 terminado em LF ou CRLF: a saida do ciclo 1 e' byte a byte a do ciclo 3 (antes crescia 2 bytes por ciclo) e nao tem linha extra
    [InlineData("\n")] [InlineData("\r\n")]
    public async Task Utf8_TerminadoEmQuebra_NaoAcumulaLinhaVazia(string fim)
    {
        var s = await Ciclos(Utf8.GetBytes("Nome,Cidade" + fim + "Ação,São Paulo" + fim), new OpcoesTexto(OpcoesTexto.Utf8, OpcoesTexto.Virgula));
        Assert.Equal(s[0], s[1]); Assert.Equal(s[0], s[2]);
        Assert.Equal("Nome,Cidade\nAção,São Paulo\n", Texto(s[0]));
    }

    [Fact] // sem quebra final: nada a tirar; a saida (com o LF que o x2t acrescenta) tambem e' estavel
    public async Task Utf8_SemQuebraFinal_Estavel()
    {
        var s = await Ciclos(Utf8.GetBytes("Nome,Cidade\nAção,São Paulo"), new OpcoesTexto(OpcoesTexto.Utf8, OpcoesTexto.Virgula));
        Assert.Equal(s[0], s[2]);
        Assert.Equal("Nome,Cidade\nAção,São Paulo\n", Texto(s[0]));
    }

    [Fact] // controle negativo: ANSI ja' era estavel e continua (o arquivo nao passa pela copia)
    public async Task Ansi_ContinuaEstavel()
    {
        var s = await Ciclos(Encoding.Latin1.GetBytes("Nome;Cidade\nAção;São Paulo\n"), OpcoesTexto.CsvExcelPtBr);
        Assert.Equal(s[0], s[2]);
        Assert.Equal("Nome;Cidade\nAção;São Paulo\n", Encoding.Latin1.GetString(s[0]).Replace("\r\n", "\n"));
    }

    [Theory] // UTF-16 LE/BE (ids 48/49) MEDIDOS: nao acumulam; continuam intactos (nao entram na copia)
    [InlineData(48)] [InlineData(49)]
    public async Task Utf16_ContinuaEstavel(int id)
    {
        var enc = id == 48 ? Encoding.Unicode : Encoding.BigEndianUnicode;
        byte[] bom = id == 48 ? [0xFF, 0xFE] : [0xFE, 0xFF];
        var s = await Ciclos([.. bom, .. enc.GetBytes("Nome,Cidade\nAção,São Paulo\n")], new OpcoesTexto(id, OpcoesTexto.Virgula));
        Assert.Equal(s[0], s[2]);
        Assert.Equal("Nome,Cidade\nAção,São Paulo\n", enc.GetString(s[0]).TrimStart('﻿').Replace("\r\n", "\n"));
    }

    [Fact] // uma linha vazia INTENCIONAL no meio do conteudo e uma linha final so' de delimitador (celulas vazias) sobrevivem aos ciclos
    public async Task LinhasVaziasIntencionais_Preservadas()
    {
        var op = new OpcoesTexto(OpcoesTexto.Utf8, OpcoesTexto.Virgula);
        var meio = await Ciclos(Utf8.GetBytes("a,b\n,\nc,d\n"), op);
        Assert.Equal(meio[0], meio[2]);
        Assert.Equal("a,b\n,\nc,d\n", Texto(meio[0]));
        var fim = await Ciclos(Utf8.GetBytes("a,b\n,\n"), op);      // a ultima linha ',' e' dado do usuario
        Assert.Equal(fim[0], fim[2]);
        Assert.Equal("a,b\n,\n", Texto(fim[0]));
    }

    [Fact] // TXT nao acumula (MEDIDO): ciclos de abrir -> salvar do mesmo texto UTF-8 saem identicos
    public async Task Txt_NaoAcumula()
    {
        var c = new X2tConversor(MotorDeTeste.Localizar());
        var atual = Path.Combine(_pasta, "t.txt");
        File.WriteAllBytes(atual, Utf8.GetBytes("linha1\nlinha2\n"));
        var saidas = new List<byte[]>();
        for (var i = 1; i <= 3; i++)
        {
            var rec = Path.Combine(_pasta, "rect-" + i);
            await c.ParaBinAsync(atual, rec, texto: OpcoesTexto.TxtUtf8);
            var saida = Path.Combine(_pasta, $"t-out-{i}.txt");
            await c.DeBinAsync(rec, FormatoDocumento.Txt, saida, texto: OpcoesTexto.TxtUtf8);
            saidas.Add(File.ReadAllBytes(saida)); atual = saida;
        }
        Assert.Equal(saidas[0], saidas[2]);
    }

    [Theory] // a copia: tira exatamente UM terminador (LF ou CRLF), nao toca no original e nao existe quando nao ha' o que tirar
    [InlineData("a,b\n", 1)] [InlineData("a,b\r\n", 2)] [InlineData("a,b\n\n", 1)] [InlineData("a,b\r\n\r\n", 2)] [InlineData("\n", 1)]
    public void Copia_TiraUmTerminador(string conteudo, int tirados)
    {
        var f = Path.Combine(_pasta, "o.csv"); var original = Utf8.GetBytes(conteudo); File.WriteAllBytes(f, original);
        var pasta = Path.Combine(_pasta, "copia"); Directory.CreateDirectory(pasta);
        var copia = X2tConversor.CopiaSemTerminadorFinal(f, pasta)!;
        Assert.Equal(pasta, Path.GetDirectoryName(copia));
        Assert.Equal(".csv", Path.GetExtension(copia));
        Assert.Equal(original[..^tirados], File.ReadAllBytes(copia));
        Assert.Equal(original, File.ReadAllBytes(f));
    }

    [Theory]
    [InlineData("a,b")] [InlineData("a,b\r")] [InlineData("")] [InlineData("a,\"b\n\"")]
    public void Copia_SemTerminadorFinal_NaoCopia(string conteudo)
    {
        var f = Path.Combine(_pasta, "o.csv"); File.WriteAllBytes(f, Utf8.GetBytes(conteudo));
        Assert.Null(X2tConversor.CopiaSemTerminadorFinal(f, _pasta));
    }

    [Fact] // so' o CSV em UTF-8 passa pela copia: o XML de abrir leva a copia, e o arquivo de ORIGEM dos outros casos segue o do usuario
    public async Task OArquivoDoUsuarioNaoMuda()
    {
        var f = Path.Combine(_pasta, "u.csv"); var original = Utf8.GetBytes("a,b\nc,d\n"); File.WriteAllBytes(f, original);
        var antes = File.GetLastWriteTimeUtc(f);
        var c = new X2tConversor(MotorDeTeste.Localizar());
        await c.ParaBinAsync(f, Path.Combine(_pasta, "rec-u"), texto: new OpcoesTexto(OpcoesTexto.Utf8, OpcoesTexto.Virgula));
        Assert.Equal(original, File.ReadAllBytes(f));
        Assert.Equal(antes, File.GetLastWriteTimeUtc(f));
        Assert.Equal([f], Directory.GetFiles(_pasta, "*.csv"));   // a copia ficou so' na pasta temporaria do x2t, ja' apagada
    }
}
