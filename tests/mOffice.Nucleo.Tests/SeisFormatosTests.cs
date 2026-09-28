using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 7 — os TRES editores (word/cell/slide) e os SEIS formatos (DOCX/ODT/XLSX/ODS/PPTX/ODP) pelo caminho
/// completo do produto: abrir no editor real -> editar pela API publica DELE -> salvar pelo
/// asc_Save/LocalFileSave -> reler o arquivo com OpenXML/ZIP.
///
/// A Task 6 provou o ciclo so para `word` + DOCX. Aqui a mesma mecanica vale para os outros dois editores e
/// para as tres saidas ODF. As APIs de edicao de cell e slide foram levantadas no sdkjs desta versao e estao
/// documentadas em <see cref="EditorReal"/> (as hipoteses do brief nao eram as deste build).
/// </summary>
[Collection(JanelaReal.Nome)]
public class SeisFormatosTests
{
    /// <summary>Frase inserida pelo editor. Nao pode existir no documento de origem.</summary>
    public const string Inserido = "mOffice PoC 2026-09";

    /// <summary>Texto que ja existe no documento de origem — tem de SOBREVIVER a edicao + salvamento.</summary>
    public const string Original = "Original mOffice";

    /// <summary>Celula usada na planilha: A1 guarda o texto de origem, entao a edicao vai para B1 (o brief ja
    /// previa essa adaptacao). Registrada no relatorio.</summary>
    private const string CelulaDaEdicao = "B1";

    private readonly ITestOutputHelper _saida;
    public SeisFormatosTests(ITestOutputHelper saida) { _saida = saida; }

    public static TheoryData<FormatoDocumento> Formatos =>
    [
        FormatoDocumento.Docx, FormatoDocumento.Odt,
        FormatoDocumento.Xlsx, FormatoDocumento.Ods,
        FormatoDocumento.Pptx, FormatoDocumento.Odp,
    ];

    [SkippableTheory] // janela real: abrir -> editar -> salvar -> reler, para cada um dos 6 formatos
    [MemberData(nameof(Formatos))]
    public void AbreEditaSalva_Formato(FormatoDocumento fmt)
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory($"moffice-6f-{fmt}-").FullName;
        var arquivo = Fabricar(pasta, fmt);
        var doctype = fmt.DocType();

        // Linha de base lida do ARQUIVO (nao do editor): em cell/slide um asc_EditSelectAll antes de editar
        // mexeria na selecao e a insercao cairia no lugar errado (em word isso ja custou um bug na Task 6).
        var antes = Ler(arquivo, fmt);
        Assert.Contains(Original, antes);
        Assert.DoesNotContain(Inserido, antes);

        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.NovaJanela();
            j.Show();
            await j.AbrirAsync(arquivo);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));

            // O sinal de determinismo tem de COMECAR falso: se `asc_isDocumentCanSave()` ja viesse `true` de
            // um documento recem-aberto, a espera abaixo ("virou true") nao provaria nada — passaria sem a
            // edicao ter chegado ao modelo. Asserir, e nao so registrar (fix review 1, minor 4).
            var podeSalvarAntes = (await j.ExecutarJsAsync(EditorReal.PodeSalvarNoEditor)).Trim('"');
            _saida.WriteLine($"[{fmt}/{doctype}] asc_isDocumentCanSave antes de editar: {podeSalvarAntes}");
            Assert.Equal("false", podeSalvarAntes);

            var inserido = (await j.ExecutarJsAsync(EditorReal.InserirPorDocType(doctype, Inserido, CelulaDaEdicao))).Trim('"');
            _saida.WriteLine($"[{fmt}/{doctype}] insercao: {inserido}");
            Assert.StartsWith("ok", inserido);

            // Espera DETERMINISTICA: em cell/slide a colagem passa por callback (pluginMethod_PasteText ->
            // Ktb/teb com continuacao), entao a mudanca nem sempre esta no modelo quando o script retorna.
            // asc_isDocumentCanSave() e' o sinal publico de "ha alteracao nao salva" nos tres editores.
            var podeSalvar = await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 100);
            Assert.Equal("true", podeSalvar);

            await j.SalvarAsync();
            Assert.Null(j.UltimoErro);

            // A edicao passou pela ponte (changes0.json): sem ela o x2t so reescreveria o Editor.bin original.
            Assert.True(File.Exists(j.SessaoAtual!.ArquivoChanges), "nenhum change gravado — a edição não chegou na ponte");

            var depois = Ler(arquivo, fmt);
            _saida.WriteLine($"[{fmt}/{doctype}] texto salvo ({depois.Length} chars): {Resumo(depois)}");
            Assert.Contains(Inserido, depois);
            Assert.Contains(Original, depois);

            // XLSX: o texto integral nao distingue "inseriu em B1" de "sobrescreveu A1 e o Original sobrou em
            // outro lugar". Aqui a asserção e' por CELULA — A1 EXATAMENTE o texto de origem, B1 com o inserido
            // (fix review 1, minor 4). So vale para XLSX: em ODS o arquivo salvo e' ODF, lido pelo content.xml.
            if (fmt == FormatoDocumento.Xlsx)
            {
                Assert.Equal(Original, FabricaDocumentos.TextoXlsxCelula(arquivo, "A1"));
                Assert.Contains(Inserido, FabricaDocumentos.TextoXlsxCelula(arquivo, CelulaDaEdicao));
            }

            Assert.Empty(j.Guarda.Bloqueadas);
            j.Close();
        }, TimeSpan.FromMinutes(3));
    }

    // -----------------------------------------------------------------------------------------------------

    /// <summary>OOXML vem da fabrica; ODF vem do OOXML convertido pelo x2t (a Task 2 ja cobre essa conversao).</summary>
    private static string Fabricar(string pasta, FormatoDocumento fmt)
    {
        var ooxml = fmt switch
        {
            FormatoDocumento.Docx or FormatoDocumento.Odt => FabricaDocumentos.Docx(pasta, Original),
            FormatoDocumento.Xlsx or FormatoDocumento.Ods => FabricaDocumentos.Xlsx(pasta, Original),
            _ => FabricaDocumentos.Pptx(pasta, Original),
        };
        if (fmt is FormatoDocumento.Docx or FormatoDocumento.Xlsx or FormatoDocumento.Pptx) return ooxml;
        var odf = Path.Combine(pasta, "origem." + fmt.Extensao());
        new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar()).ConverterAsync(ooxml, fmt, odf).GetAwaiter().GetResult();
        return odf;
    }

    /// <summary>Texto integral do arquivo, por formato. ODT/ODS/ODP sao ZIP: o content.xml inteiro serve
    /// (o que interessa e' "a frase esta la'"), e as contagens objetivas ficam com tools/fidelidade.ps1.</summary>
    private static string Ler(string caminho, FormatoDocumento fmt) => fmt switch
    {
        FormatoDocumento.Docx => FabricaDocumentos.TextoDocx(caminho),
        FormatoDocumento.Xlsx => FabricaDocumentos.TextoXlsx(caminho),
        FormatoDocumento.Pptx => FabricaDocumentos.TextoPptx(caminho),
        _ => FabricaDocumentos.TextoOdf(caminho),
    };

    private static string Resumo(string texto)
    {
        var limpo = System.Text.RegularExpressions.Regex.Replace(texto, @"\s+", " ").Trim();
        return limpo.Length <= 300 ? limpo : limpo[..300] + "…";
    }
}
