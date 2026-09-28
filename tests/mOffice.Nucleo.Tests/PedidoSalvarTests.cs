using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>Parsing do pedido de salvamento que chega do editor (<see cref="PedidoSalvar"/>), sem precisar de
/// janela nenhuma — minor 6 da revisao 2026-09-17 (fix review 1): antes so era exercitado indiretamente pelos
/// testes de janela real, que so cobrem o "salvar normal" (fileType 0, param vazio).</summary>
public class PedidoSalvarTests
{
    [Fact]
    public void Ler_SaveAsTrue_ReconheceSalvarComo()
    {
        var (salvarComo, fileType) = PedidoSalvar.Ler("""{"param":"saveas=true;","fileType":0,"jsonOptions":""}""");
        Assert.True(salvarComo);
        Assert.Equal(0, fileType);
    }

    [Fact]
    public void Ler_FileType65_DevolveDocx()
    {
        var (_, fileType) = PedidoSalvar.Ler("""{"param":"","fileType":65,"jsonOptions":""}""");
        Assert.Equal(65, fileType);
        Assert.Equal(FormatoDocumento.Docx, PedidoSalvar.FormatoPorCodigo(fileType));
    }

    [Fact]
    public void Ler_FileType67_DevolveOdt()
    {
        var (_, fileType) = PedidoSalvar.Ler("""{"param":"","fileType":67,"jsonOptions":""}""");
        Assert.Equal(67, fileType);
        Assert.Equal(FormatoDocumento.Odt, PedidoSalvar.FormatoPorCodigo(fileType));
    }

    [Fact]
    public void Ler_ParamVazioFileTypeZero_ReconheceSalvarNormal()
    {
        var (salvarComo, fileType) = PedidoSalvar.Ler("""{"param":"","fileType":0,"jsonOptions":""}""");
        Assert.False(salvarComo);
        Assert.Equal(0, fileType);
    }

    [Fact]
    public void Ler_JsonMalformado_DevolveSalvarNormalSemFileType()
    {
        // Fire-and-forget (SalvarPedidoDoEditorAsync): um JSON inesperado nao pode propagar excecao daqui,
        // senao derrubaria a ponte inteira por uma mensagem que o proprio C# nao controla.
        var (salvarComo, fileType) = PedidoSalvar.Ler("isto nao e' json");
        Assert.False(salvarComo);
        Assert.Equal(0, fileType);
    }

    [Fact]
    public void FormatoPorCodigo_CodigoDesconhecido_DevolveNulo()
    {
        Assert.Null(PedidoSalvar.FormatoPorCodigo(9999));
    }

    [Fact]
    public void FormatoPorCodigo_Zero_DevolveNulo()
    {
        // 0 = "nenhum formato escolhido pelo chamador" (asc_Save sem asc_DownloadAs) — nao e' um FormatoDocumento valido.
        Assert.Null(PedidoSalvar.FormatoPorCodigo(0));
    }
}
