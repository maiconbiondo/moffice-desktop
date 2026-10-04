using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using mOffice.Conversor;

namespace mOffice.Nucleo;

/// <summary>
/// O que o editor responde em <c>AscDesktopEditor.SetAdvancedOptions(xml)</c> (Task 7 do Plano 3). MEDIDO nos bundles
/// 9.4.0.129: para arquivo protegido os tres editores mandam exatamente
/// <c>"&lt;m_sPassword&gt;" + escape(senha) + "&lt;/m_sPassword&gt;"</c> (word <c>Asc.asc_docs_api.prototype.asc_setAdvancedOptions</c>
/// em <c>sdkjs/word/sdk-all.js</c>, cell <c>spreadsheet_api.prototype.asc_setAdvancedOptions</c>, slide idem), onde o
/// escape e' o de XML (<c>&amp; &lt; &gt; &apos; &quot;</c>, <c>AscCommon.lyd</c> no word). Para CSV/TXT o mesmo canal
/// leva <c>&lt;m_nCsvTxtEncoding&gt;</c>/<c>&lt;m_nCsvDelimiter&gt;</c>/<c>&lt;m_nCsvDelimiterChar&gt;</c> — um
/// fragmento com varios elementos soltos, sem raiz. Por isso a leitura embrulha o texto numa raiz antes de parsear.
/// </summary>
public static class OpcoesAvancadas
{
    /// <summary>Texto de <c>&lt;m_sPassword&gt;</c>, ja' decodificado (<c>&amp;amp;</c> -&gt; <c>&amp;</c>). Ausente,
    /// XML quebrado ou vazio de entrada = <c>null</c>. Uma senha vazia (<c>&lt;m_sPassword&gt;&lt;/m_sPassword&gt;</c>)
    /// volta como <c>""</c>: quem chama decide o que ela significa.</summary>
    public static string? Senha(string? xml)
    {
        if (string.IsNullOrEmpty(xml)) return null;
        try
        {
            // Leitor EXPLICITO (M4 da revisao final do Plano 3): o texto vem da pagina. O comentario antigo dizia
            // "DTD proibido (padrao do XmlReader)", mas o XElement.Parse monta o PROPRIO XmlReaderSettings com
            // DtdProcessing.Parse — so' a raiz embrulhada impedia um DOCTYPE. Aqui o DTD e' proibido de verdade e
            // sem resolver de entidade externa, e a raiz continua embrulhada (o fragmento de CSV nao tem raiz).
            using var leitor = XmlReader.Create(new StringReader("<opcoes>" + xml + "</opcoes>"), Configuracao);
            // PreserveWhitespace (e IgnoreWhitespace = false no leitor): uma senha feita SO' de espacos e' uma senha
            // valida; com LoadOptions.None esse texto era descartado como espaco "insignificante" e a senha virava "".
            var raiz = XElement.Load(leitor, LoadOptions.PreserveWhitespace);
            return raiz.Element("m_sPassword")?.Value;
        }
        catch (XmlException) { return null; }
    }

    /// <summary>A resposta do dialogo de importacao de CSV/TXT do editor (Plano 6, Task 3): <c>m_nCsvTxtEncoding</c> (o ID da lista de
    /// codificacoes do editor, 0 a <see cref="OpcoesTexto.MaiorIdDeCodificacao"/>), <c>m_nCsvDelimiter</c> (ausente no TXT = sem
    /// delimitador) e <c>m_nCsvDelimiterChar</c> (o "Outro"). Sem codificacao valida, XML quebrado ou vazio = <c>null</c> (a abertura e'
    /// cancelada: um ID fora da lista faria o x2t cair — MEDIDO com 1252). Mesmo leitor de <see cref="Senha"/> (DTD proibido).</summary>
    public static OpcoesTexto? Texto(string? xml)
    {
        if (string.IsNullOrEmpty(xml)) return null;
        try
        {
            using var leitor = XmlReader.Create(new StringReader("<opcoes>" + xml + "</opcoes>"), Configuracao);
            var raiz = XElement.Load(leitor, LoadOptions.PreserveWhitespace);
            if (!int.TryParse(raiz.Element("m_nCsvTxtEncoding")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var codificacao)
                || codificacao is < 0 or > OpcoesTexto.MaiorIdDeCodificacao) return null;
            var delimitador = int.TryParse(raiz.Element("m_nCsvDelimiter")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var d)
                ? d : OpcoesTexto.SemDelimitador;
            // So' os codigos MEDIDOS no x2t: 1..5 (tab, ';', ':', ',', espaco), 0 (TXT / "Outro" so' com caractere) e -1 ("Outro" do dialogo).
            if (delimitador is < OpcoesTexto.Outro or > OpcoesTexto.Espaco) return null;
            var outro = raiz.Element("m_nCsvDelimiterChar")?.Value;
            return new OpcoesTexto(codificacao, delimitador, DelimitadorChar: string.IsNullOrEmpty(outro) ? null : outro);
        }
        catch (XmlException) { return null; }
    }

    /// <summary>Configuracao do leitor de <see cref="Senha"/>: DTD proibido (um DOCTYPE lanca
    /// <see cref="XmlException"/>), nenhum resolver externo, espacos preservados. `internal` para o teste da regra.</summary>
    internal static readonly XmlReaderSettings Configuracao = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreWhitespace = false,
        MaxCharactersFromEntities = 1024,
    };
}
