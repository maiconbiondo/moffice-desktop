namespace mOffice.Conversor;

public sealed class ErroConversao(string mensagem, int codigoSaida, string stderr) : Exception(mensagem)
{
    public int CodigoSaida { get; } = codigoSaida;
    public string Stderr { get; } = stderr;
}
