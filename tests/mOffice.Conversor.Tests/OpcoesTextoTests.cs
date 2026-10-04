using System.Text;
using mOffice.Conversor;

namespace mOffice.Conversor.Tests;

public class OpcoesTextoTests
{
    private static string Arquivo(byte[] bytes, string ext = ".csv")
    {
        var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N") + ext);
        Directory.CreateDirectory(Path.GetDirectoryName(p)!); File.WriteAllBytes(p, bytes); return p;
    }

    [Fact] // BOM UTF-8 -> UTF-8 (46)
    public void ComBom_Utf8() =>
        Assert.Equal(OpcoesTexto.Utf8, SugestaoDeTexto.Detectar(Arquivo([0xEF, 0xBB, 0xBF, .. Encoding.UTF8.GetBytes("Nome;Valor\nJoão;1,5\n")]), csv: true).Codificacao);

    [Fact] // UTF-8 valido sem BOM -> UTF-8
    public void SemBomUtf8Valido_Utf8() =>
        Assert.Equal(OpcoesTexto.Utf8, SugestaoDeTexto.Detectar(Arquivo(Encoding.UTF8.GetBytes("Nome;Cidade\nAção;São Paulo\n")), csv: true).Codificacao);

    [Fact] // bytes 1252 (acento invalido em UTF-8) -> 1252, o padrao do Excel pt-BR
    public void Ansi_1252() =>
        Assert.Equal(OpcoesTexto.Windows1252, SugestaoDeTexto.Detectar(Arquivo(Encoding.Latin1.GetBytes("Nome;Cidade\nAção;São Paulo\n")), csv: true).Codificacao);

    [Fact] // o corte dos 64 KiB no meio de um caractere UTF-8 de varios bytes nao faz um arquivo valido virar 1252
    public void CorteNoMeioDeCaractere_Utf8()
    {
        var linha = "Ação;São Paulo;ç\n";
        var sb = new StringBuilder();
        while (Encoding.UTF8.GetByteCount(sb.ToString()) < 70_000) sb.Append(linha);
        var bytes = Encoding.UTF8.GetBytes(sb.ToString());
        Assert.Equal(OpcoesTexto.Utf8, SugestaoDeTexto.Detectar(Arquivo(bytes), csv: true).Codificacao);
    }

    [Theory] // separador mais frequente nas primeiras linhas (fora de aspas); empate ou nenhum -> ';' (pt-BR)
    [InlineData("a;b;c\n1;2;3\n", OpcoesTexto.PontoEVirgula)]
    [InlineData("a,b,c\n1,2,3\n", OpcoesTexto.Virgula)]
    [InlineData("a\tb\n1\t2\n", OpcoesTexto.Tabulacao)]
    [InlineData("\"x,y\";b\n\"1,5\";2\n", OpcoesTexto.PontoEVirgula)]
    [InlineData("so uma coluna\n", OpcoesTexto.PontoEVirgula)]
    [InlineData("a;b,c\n1;2,3\n", OpcoesTexto.PontoEVirgula)]   // empate: ';'
    public void Delimitador(string conteudo, int esperado) =>
        Assert.Equal(esperado, SugestaoDeTexto.Detectar(Arquivo(Encoding.UTF8.GetBytes(conteudo)), csv: true).Delimitador);

    [Fact] // TXT nao tem delimitador
    public void Txt_SemDelimitador() =>
        Assert.Equal(OpcoesTexto.SemDelimitador, SugestaoDeTexto.Detectar(Arquivo(Encoding.UTF8.GetBytes("linha;com;ponto"), ".txt"), csv: false).Delimitador);

    [Fact] // arquivo inexistente/ilegivel: o padrao pt-BR em vez de excecao (a sugestao nunca derruba a abertura)
    public void ArquivoIlegivel_PadraoPtBr() =>
        Assert.Equal(OpcoesTexto.CsvExcelPtBr, SugestaoDeTexto.Detectar(Path.Combine(Path.GetTempPath(), "nao-existe-" + Guid.NewGuid().ToString("N") + ".csv"), csv: true) with { Lcid = OpcoesTexto.CsvExcelPtBr.Lcid });

    [Fact] // os codigos sao os do editor (conferidos no bundle 9.4): delimitador 1 tab, 2 ';', 3 ':', 4 ',', 5 espaco; codificacao = ID da lista
           // c_oAscEncodings do sdkjs (44 = windows-1252, 46 = UTF-8), NAO o codepage do Windows (1252 crasha o x2t, medido)
    public void CodigosLiterais()
    {
        Assert.Equal((46, 44, 0, 1, 2, 3, 4, 5), (OpcoesTexto.Utf8, OpcoesTexto.Windows1252, OpcoesTexto.SemDelimitador,
            OpcoesTexto.Tabulacao, OpcoesTexto.PontoEVirgula, OpcoesTexto.DoisPontos, OpcoesTexto.Virgula, OpcoesTexto.Espaco));
        Assert.Equal(new OpcoesTexto(44, 2, 1046), OpcoesTexto.CsvExcelPtBr);
        Assert.Equal(new OpcoesTexto(46, 2, 1046), OpcoesTexto.CsvUtf8);
        Assert.Equal(new OpcoesTexto(46, 0, 1046), OpcoesTexto.TxtUtf8);
    }
}
