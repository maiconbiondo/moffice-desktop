using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>"Novo documento" = copia do modelo vazio que o proprio motor traz em converter/empty/&lt;idioma&gt;/
/// (spec §6). Nao ha modelo ODF: os formatos ODF usam o irmao OOXML e o primeiro "Salvar" converte.</summary>
public static class ModelosVazios
{
    public static string NomeDoModelo(FormatoDocumento formato) => formato switch
    {
        FormatoDocumento.Docx or FormatoDocumento.Odt => "new.docx",
        FormatoDocumento.Xlsx or FormatoDocumento.Ods => "new.xlsx",
        FormatoDocumento.Pptx or FormatoDocumento.Odp => "new.pptx",
        _ => throw new ArgumentOutOfRangeException(nameof(formato), formato, "Sem modelo vazio para este formato."),
    };

    public static string Caminho(ConfiguracaoMotor motor, FormatoDocumento formato)
    {
        var nome = NomeDoModelo(formato);
        var raiz = Path.Combine(motor.PastaMotor, "converter", "empty");
        foreach (var idioma in new[] { "pt-BR", "default" })
        {
            var candidato = Path.Combine(raiz, idioma, nome);
            if (File.Exists(candidato)) return candidato;
        }
        throw new FileNotFoundException($"Modelo vazio {nome} não encontrado em {raiz} (pt-BR nem default).", nome);
    }
}
