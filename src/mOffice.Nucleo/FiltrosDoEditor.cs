namespace mOffice.Nucleo;

/// <summary>O que o caminho escolhido num dialogo do host autoriza a pagina a fazer com ele (Plano 6, Task 8; MvP-15/PL-08). Consentimento de
/// UM uso por caminho, como o <c>g_pLocalResolver</c> do app oficial: <see cref="Imagem"/> = copiar para media/
/// (<see cref="PonteDesktop.PermitirArquivoEscolhido"/>); <see cref="Conversao"/> = converter com o x2t para o binario do editor
/// (<see cref="PonteDesktop.PermitirConversao"/>); <see cref="Leitura"/> = entregar os bytes (<see cref="PonteDesktop.PermitirLeitura"/>);
/// <see cref="Nenhum"/> = so' o caminho volta ao editor (hiperlink para arquivo), nada e' lido.</summary>
public enum Consentimento { Imagem, Conversao, Leitura, Nenhum }

/// <summary>Titulo e filtro do dialogo do Windows e o consentimento que a escolha da' (ver <see cref="FiltrosDoEditor"/>).</summary>
public sealed record PedidoDeArquivo(string Titulo, string Filtro, Consentimento Consentimento);

/// <summary>
/// <c>AscDesktopEditor.OpenFilenameDialog(filtro, multi, cb)</c> -&gt; o dialogo que o host abre. Filtros MEDIDOS nos bundles 9.4.0.129: "images"
/// (inserir imagem, todos os editores), "word" (Inserir texto do arquivo — multi —, Comparar/Combinar do arquivo: sdkjs/word), "cell" (gráfico/vinculo),
/// "(*.xml)" (Dados &gt; XML: cell), "csv/txt" (Dados &gt; De texto/CSV: cell), "any" (hiperlink para arquivo: cell <c>asc_getFilePath</c>),
/// "audio"/"video" (so' com <c>IsSupportMedia</c>, que o mOffice responde false). Antes desta task todo filtro que nao fosse "word"/"csv/txt"
/// caia no dialogo "Inserir imagem" e dava consentimento de copia para media/ a um arquivo que nao era imagem (VERIFICADO na auditoria).
/// </summary>
public static class FiltrosDoEditor
{
    /// <summary>Pedido para o filtro do motor, ou <c>null</c> (audio/video e qualquer filtro desconhecido: nenhum dialogo).</summary>
    public static PedidoDeArquivo? Para(string filtroDoMotor) => filtroDoMotor switch
    {
        "images" => new("Inserir imagem", "Imagens|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.svg|Todos os arquivos|*.*", Consentimento.Imagem),
        "word" => new("Selecionar documento", "Documentos|*.docx;*.docm;*.dotx;*.doc;*.odt;*.rtf;*.txt|Todos os arquivos|*.*", Consentimento.Conversao),
        "cell" => new("Selecionar planilha", "Planilhas|*.xlsx;*.xlsm;*.xls;*.ods;*.csv|Todos os arquivos|*.*", Consentimento.Conversao),
        "(*.xml)" => new("Importar XML", "XML|*.xml|Todos os arquivos|*.*", Consentimento.Conversao),
        "csv/txt" => new("Importar texto ou CSV", "Texto e CSV|*.csv;*.txt|Todos os arquivos|*.*", Consentimento.Leitura),
        "any" => new("Selecionar arquivo", "Todos os arquivos|*.*", Consentimento.Nenhum),
        _ => null,
    };
}
