using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>"Novo documento" = copia do modelo vazio que o proprio motor traz em converter/empty/&lt;idioma&gt;/
/// (spec §6), mas antes procura o modelo proprio do mOffice em <c>Modelos\pt-BR</c> ao lado do exe
/// (derivado do do motor por tools/modelos/gerar-modelos.py). Nao ha modelo ODF: os formatos ODF usam o irmao OOXML e o primeiro "Salvar" converte.</summary>
public static class ModelosVazios
{
    public static string NomeDoModelo(FormatoDocumento formato) => formato switch
    {
        FormatoDocumento.Docx or FormatoDocumento.Odt => "new.docx",
        FormatoDocumento.Xlsx or FormatoDocumento.Ods => "new.xlsx",
        FormatoDocumento.Pptx or FormatoDocumento.Odp => "new.pptx",
        _ => throw new ArgumentOutOfRangeException(nameof(formato), formato, "Sem modelo vazio para este formato."),
    };

    /// <param name="pastaProprios">Onde ficam os modelos proprios; padrao <c>&lt;exe&gt;\Modelos\pt-BR</c> (parametro para teste).</param>
    public static string Caminho(ConfiguracaoMotor motor, FormatoDocumento formato, string? pastaProprios = null)
    {
        var nome = NomeDoModelo(formato);
        // Os modelos PROPRIOS do mOffice (Calibri 11, 1,08, 8 pt depois; layouts do PowerPoint pt-BR) vem antes dos do motor.
        var proprio = Path.Combine(pastaProprios ?? Path.Combine(AppContext.BaseDirectory, "Modelos", "pt-BR"), nome);
        if (File.Exists(proprio)) return proprio;
        var raiz = Path.Combine(motor.PastaMotor, "converter", "empty");
        foreach (var idioma in new[] { "pt-BR", "default" })
        {
            var candidato = Path.Combine(raiz, idioma, nome);
            if (File.Exists(candidato)) return candidato;
        }
        throw new FileNotFoundException($"Modelo vazio {nome} não encontrado em {raiz} (pt-BR nem default).", nome);
    }
}
