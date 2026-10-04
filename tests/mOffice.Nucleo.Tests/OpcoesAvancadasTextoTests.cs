using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

public class OpcoesAvancadasTextoTests
{
    [Fact] // a resposta do dialogo de CSV do editor vira OpcoesTexto
    public void Texto_LeCodificacaoEDelimitador() =>
        Assert.Equal(new OpcoesTexto(44, 2), OpcoesAvancadas.Texto("<m_nCsvTxtEncoding>44</m_nCsvTxtEncoding><m_nCsvDelimiter>2</m_nCsvDelimiter>"));

    [Fact] // resposta de senha nao e' opcao de texto
    public void Texto_RespostaDeSenha_Nulo() => Assert.Null(OpcoesAvancadas.Texto("<m_sPassword>x</m_sPassword>"));

    [Fact] // TXT: so' a codificacao (o dialogo de texto nao tem delimitador) -> delimitador 0
    public void Texto_SoCodificacao_SemDelimitador() =>
        Assert.Equal(new OpcoesTexto(46, OpcoesTexto.SemDelimitador), OpcoesAvancadas.Texto("<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>"));

    [Fact] // "Outro" no dialogo: o caractere vai junto (m_nCsvDelimiterChar) e o codigo, se vier, e' preservado
    public void Texto_DelimitadorOutro()
    {
        Assert.Equal(new OpcoesTexto(44, OpcoesTexto.SemDelimitador, DelimitadorChar: "|"),
            OpcoesAvancadas.Texto("<m_nCsvTxtEncoding>44</m_nCsvTxtEncoding><m_nCsvDelimiterChar>|</m_nCsvDelimiterChar>"));
        Assert.Equal(new OpcoesTexto(46, -1, DelimitadorChar: "#"),
            OpcoesAvancadas.Texto("<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding><m_nCsvDelimiter>-1</m_nCsvDelimiter><m_nCsvDelimiterChar>#</m_nCsvDelimiterChar>"));
    }

    [Theory] // vazio, quebrado, DTD e valores nao numericos: nulo (a abertura e' cancelada, nunca abre com lixo)
    [InlineData(null)] [InlineData("")] [InlineData("<m_nCsvTxtEncoding>")] [InlineData("<m_nCsvTxtEncoding>abc</m_nCsvTxtEncoding>")]
    [InlineData("<!DOCTYPE x [<!ENTITY e \"1\">]><m_nCsvTxtEncoding>&e;</m_nCsvTxtEncoding>")]
    public void Texto_Invalido_Nulo(string? xml) => Assert.Null(OpcoesAvancadas.Texto(xml));
}

public class SessaoDocumentoTextoTests
{
    private static SessaoDocumento Sessao(FormatoDocumento f, bool novo = false, string? titulo = null) =>
        new() { ArquivoOrigem = @"C:\x\a." + f.Extensao(), Formato = f, PastaRecovery = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")), Novo = novo, TituloDoNovo = titulo };

    [Fact] // PDF aberto como documento: o titulo do documento NOVO e' o nome do PDF; sem ele, o rotulo generico
    public void TituloDoNovo_TemPrecedenciaSoEmDocumentoNovo()
    {
        Assert.Equal("contrato", Sessao(FormatoDocumento.Docx, novo: true, titulo: "contrato").Titulo);
        Assert.Equal("Documento", Sessao(FormatoDocumento.Docx, novo: true).Titulo);
        Assert.Equal("a.docx", Sessao(FormatoDocumento.Docx, novo: false, titulo: "contrato").Titulo);   // depois do primeiro "Salvar como" o titulo e' o do arquivo
    }

    [Fact] // so' CSV/TXT guardam as opcoes ao repontar; qualquer outro formato zera (o Ctrl+S seguinte nao herda)
    public void Repontar_GuardaOpcoesSoDeTexto()
    {
        var s = Sessao(FormatoDocumento.Xlsx);
        s.Repontar(@"C:\x\a.csv", FormatoDocumento.Csv, OpcoesTexto.CsvUtf8);
        Assert.Equal(OpcoesTexto.CsvUtf8, s.OpcoesTexto);
        s.Repontar(@"C:\x\a.xlsx", FormatoDocumento.Xlsx, OpcoesTexto.CsvUtf8);
        Assert.Null(s.OpcoesTexto);
    }
}

public class OpcoesTextoPersistenciaTests
{
    private static string PastaNova() => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"))).FullName;

    [Theory] // so' os codigos de delimitador MEDIDOS no x2t: -1 (Outro) a 5; qualquer outro e' resposta invalida
    [InlineData(-1, true)] [InlineData(0, true)] [InlineData(1, true)] [InlineData(5, true)] [InlineData(6, false)] [InlineData(-2, false)] [InlineData(99, false)]
    public void Texto_FaixaDoDelimitador(int codigo, bool valido) =>
        Assert.Equal(valido, OpcoesAvancadas.Texto($"<m_nCsvTxtEncoding>44</m_nCsvTxtEncoding><m_nCsvDelimiter>{codigo}</m_nCsvDelimiter>") is not null);

    [Fact] // as opcoes vao e voltam pelo sessao.json; um sessao.json ANTIGO (sem os campos) le' com null
    public void Metadados_GuardamAsOpcoesDeTexto()
    {
        var pasta = PastaNova();
        new mOffice.Nucleo.Recuperacao.MetadadosSessao(@"C:\x\a.csv", (int)FormatoDocumento.Csv, "a.csv", false, "mplanilha", 1, TextoCodificacao: 46, TextoDelimitador: 4, TextoDelimitadorChar: "|").Gravar(pasta);
        var m = mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(pasta)!;
        Assert.Equal((46, 4, "|"), (m.TextoCodificacao, m.TextoDelimitador, m.TextoDelimitadorChar));
        File.WriteAllText(Path.Combine(pasta, "sessao.json"), """{"Origem":"C:/x/a.csv","Formato":260,"Titulo":"a.csv","Novo":false,"App":"mplanilha","IdUsuario":1}""");
        var antigo = mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(pasta)!;
        Assert.Null(antigo.TextoCodificacao); Assert.Null(antigo.TextoDelimitador);
    }

    [Fact] // definir as opcoes na sessao regrava o sessao.json; Repontar tambem (e zera fora de CSV/TXT)
    public void Sessao_PersisteAsOpcoes()
    {
        var pasta = PastaNova();
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\a.csv", Formato = FormatoDocumento.Csv, PastaRecovery = pasta };
        s.Metadados = new mOffice.Nucleo.Recuperacao.MetadadosSessao(@"C:\x\a.csv", (int)FormatoDocumento.Csv, "a.csv", false, "mplanilha", 1);
        s.OpcoesTexto = OpcoesTexto.CsvUtf8;
        Assert.Equal((46, 2), (mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(pasta)!.TextoCodificacao, mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(pasta)!.TextoDelimitador));
        s.Repontar(@"C:\x\b.csv", FormatoDocumento.Csv, OpcoesTexto.CsvExcelPtBr);
        Assert.Equal(44, mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(pasta)!.TextoCodificacao);
        s.Repontar(@"C:\x\b.xlsx", FormatoDocumento.Xlsx);
        Assert.Null(mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(pasta)!.TextoCodificacao);
    }

    [Fact] // sessao com opcoes desconhecidas exige Salvar como; definir as opcoes (ou Repontar) tira a exigencia
    public void OpcoesDesconhecidas_ExigemSalvarComo()
    {
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\a.csv", Formato = FormatoDocumento.Csv, PastaRecovery = PastaNova(), OpcoesTextoDesconhecidas = true };
        Assert.True(s.ExigeSalvarComo);
        s.OpcoesTexto = OpcoesTexto.CsvUtf8;
        Assert.False(s.ExigeSalvarComo);
        s.OpcoesTextoDesconhecidas = true; s.Repontar(@"C:\x\b.csv", FormatoDocumento.Csv, OpcoesTexto.CsvUtf8);
        Assert.False(s.ExigeSalvarComo);
    }

    [Fact] // a pre-selecao do Salvar como acompanha a codificacao da sessao: CSV UTF-8 abre em "CSV UTF-8", nunca em ANSI (perderia caracteres)
    public void Indice_AcompanhaACodificacaoDaSessao()
    {
        var linhas = JanelaEditor.LinhasSalvar(FormatoDocumento.Csv, true).ToList();
        var utf8 = linhas.FindIndex(l => l.Texto == OpcoesTexto.CsvUtf8) + 1;
        var ansi = linhas.FindIndex(l => l.Texto == OpcoesTexto.CsvExcelPtBr) + 1;
        Assert.Equal(utf8, JanelaEditor.FiltroEIndice(FormatoDocumento.Csv, FormatoDocumento.Csv, true, new OpcoesTexto(46, 4)).Indice);
        Assert.Equal(ansi, JanelaEditor.FiltroEIndice(FormatoDocumento.Csv, FormatoDocumento.Csv, true, new OpcoesTexto(44, 2)).Indice);
        Assert.Equal(ansi, JanelaEditor.FiltroEIndice(FormatoDocumento.Csv, FormatoDocumento.Csv, true, new OpcoesTexto(37, 2)).Indice);   // sem correspondencia: a primeira
        Assert.Equal(ansi, JanelaEditor.FiltroEIndice(FormatoDocumento.Csv, FormatoDocumento.Csv, true).Indice);
    }

    [Fact] // o base64 da previa nunca passa por Replace: "CODIGO" dentro dele (e' base64 valido) chega intacto; senha e CSV levam o codigo certo
    public void ScriptDeOpcoes_Base64NaoSofreReplace()
    {
        var js = JanelaEditor.MontarScriptDeOpcoes(89, new OpcoesTexto(44, 2), "CODIGOCHAMADACODIGO");
        Assert.Contains("atob(\"CODIGOCHAMADACODIGO\")", js);
        Assert.Contains("asc_initAdvancedOptions(89, undefined,", js);
        Assert.Contains("sse-settings-csv-encoding", js);
        // JS valido: chaves simples (uma chave dobrada por engano num string interpolado quebrava o dialogo inteiro)
        Assert.Contains("try { w.asc_initAdvancedOptions(89, undefined, ", js);
        Assert.Contains("} finally { w.__moffice.dadosAbertos = null; }", js);
        Assert.DoesNotContain("}}", js); Assert.DoesNotContain("{{", js);
        var senha = JanelaEditor.MontarScriptDeOpcoes(90, null, null);
        Assert.Contains("w.asc_initAdvancedOptions(90);", senha);
        Assert.DoesNotContain("CODIGO", senha); Assert.DoesNotContain("CHAMADA", senha);
    }
}
