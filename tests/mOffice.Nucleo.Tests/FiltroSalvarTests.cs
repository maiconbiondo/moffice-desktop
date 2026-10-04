using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// O construtor do filtro do diálogo "Salvar como" (<c>JanelaEditor.FiltroSalvar</c>), sem janela nenhuma.
/// É a prova do achado CRÍTICO da revisão final do Plano 2: no caminho de FECHAMENTO o diálogo não pode
/// oferecer PDF — escolher PDF ali executa o ramo de EXPORTAÇÃO (grava o PDF, dispara `Exportado`, não
/// repõe a sessão), o fechamento segue mesmo assim e o `Closed` apaga a pasta de recuperação: o documento
/// editável do usuário some e sobra só o PDF.
/// </summary>
public class FiltroSalvarTests
{
    [Theory] // salvar-como normal: OOXML, ODF e as tres linhas de PDF (PDF, PDF/A, assinado), nessa ordem, por tipo de documento
    [InlineData(FormatoDocumento.Docx, ".docx", ".odt")]
    [InlineData(FormatoDocumento.Odt, ".docx", ".odt")]
    [InlineData(FormatoDocumento.Xlsx, ".xlsx", ".ods")]
    [InlineData(FormatoDocumento.Ods, ".xlsx", ".ods")]
    [InlineData(FormatoDocumento.Pptx, ".pptx", ".odp")]
    [InlineData(FormatoDocumento.Odp, ".pptx", ".odp")]
    public void ComExportacao_TemOsTresFormatos(FormatoDocumento f, string ooxml, string odf)
    {
        var filtro = JanelaEditor.FiltroSalvar(f, permitirExportacao: true);
        Assert.Contains("*" + ooxml, filtro);
        Assert.Contains("*" + odf, filtro);
        Assert.Contains("*.pdf", filtro);
        // A ordem é o contrato de IndiceDoFiltro (base 1: OOXML, ODF, PDF).
        Assert.True(filtro.IndexOf("*" + ooxml, StringComparison.Ordinal) < filtro.IndexOf("*" + odf, StringComparison.Ordinal));
        Assert.True(filtro.IndexOf("*" + odf, StringComparison.Ordinal) < filtro.IndexOf("*.pdf", StringComparison.Ordinal));
        // Plano 4, Task 2: as linhas 3, 4 e 5 sao PDF, PDF/A e "PDF assinado com mPDF" (o formato sai do INDICE).
        // Plano 6: as linhas 3, 4 e 5 continuam PDF, PDF/A e "PDF assinado com mPDF"; depois delas vem o resto que a tabela grava.
        Assert.Equal(["PDF (*.pdf)", "PDF/A (*.pdf)", "PDF assinado com mPDF (*.pdf)"],
            JanelaEditor.LinhasSalvar(f, permitirExportacao: true).Skip(2).Take(3).Select(l => l.Rotulo));
    }

    [Theory] // caminho de FECHAMENTO: só formatos de documento — nenhuma menção a PDF, em nenhum lugar do filtro
    [InlineData(FormatoDocumento.Docx, ".docx", ".odt")]
    [InlineData(FormatoDocumento.Odt, ".docx", ".odt")]
    [InlineData(FormatoDocumento.Xlsx, ".xlsx", ".ods")]
    [InlineData(FormatoDocumento.Ods, ".xlsx", ".ods")]
    [InlineData(FormatoDocumento.Pptx, ".pptx", ".odp")]
    [InlineData(FormatoDocumento.Odp, ".pptx", ".odp")]
    public void SemExportacao_NaoOferecePdf(FormatoDocumento f, string ooxml, string odf)
    {
        var filtro = JanelaEditor.FiltroSalvar(f, permitirExportacao: false);
        Assert.Contains("*" + ooxml, filtro);
        Assert.Contains("*" + odf, filtro);
        Assert.DoesNotContain("pdf", filtro, StringComparison.OrdinalIgnoreCase);
    }

    [Fact] // PDF sugerido: com exportação o filtro é só PDF (é o "baixar como PDF" do próprio editor)
    public void Pdf_ComExportacao_SoPdf()
    {
        var filtro = JanelaEditor.FiltroSalvar(FormatoDocumento.Pdf, permitirExportacao: true);
        Assert.Equal("PDF (*.pdf)|*.pdf", filtro);
    }

    [Fact] // PDF sugerido SEM exportação (sessão que este editor nem abre): cai no filtro de texto, nunca vazio
    public void Pdf_SemExportacao_CaiNoFiltroDeTextoENuncaFicaVazio()
    {
        var filtro = JanelaEditor.FiltroSalvar(FormatoDocumento.Pdf, permitirExportacao: false);
        Assert.DoesNotContain("pdf", filtro, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("*.docx", filtro);
        Assert.NotEqual("", filtro);
    }

    [Fact] // o padrão do parâmetro é o comportamento antigo (salvar-como da barra e do editor continuam com PDF)
    public void Padrao_EhPermitirExportacao()
        => Assert.Equal(JanelaEditor.FiltroSalvar(FormatoDocumento.Docx, permitirExportacao: true),
                        JanelaEditor.FiltroSalvar(FormatoDocumento.Docx));

    // ---- Plano 4, Task 2: PDF/A e "PDF assinado com mPDF" no "Salvar como" ----

    [Fact] // as cinco primeiras linhas nao mudam de lugar (o painel do editor pre-seleciona 3/4/5); depois, as novas
    public void Linhas_Texto_ComExportacao() =>
        Assert.Equal(["Documento do Word (*.docx)", "Documento de texto ODF (*.odt)", "PDF (*.pdf)", "PDF/A (*.pdf)", "PDF assinado com mPDF (*.pdf)",
                      "Documento Habilitado para Macro do Word (*.docm)", "Modelo do Word (*.dotx)", "Modelo Habilitado para Macro do Word (*.dotm)",
                      "Modelo de texto ODF (*.ott)", "Formato Rich Text (*.rtf)", "Texto sem formatação (*.txt)", "Página da Web (*.html)"],
            JanelaEditor.LinhasSalvar(FormatoDocumento.Docx, permitirExportacao: true).Select(l => l.Rotulo));

    [Fact] // fechamento (sem exportacao): sem PDF; planilha nunca oferece XLSB nem XLS
    public void Linhas_Planilha_SemExportacao() =>
        Assert.Equal(["Pasta de trabalho do Excel (*.xlsx)", "Planilha ODF (*.ods)", "Pasta de trabalho habilitada para macro do Excel (*.xlsm)",
                      "Modelo do Excel (*.xltx)", "Modelo habilitado para macro do Excel (*.xltm)", "Modelo de planilha ODF (*.ots)",
                      "CSV (separado por ponto e vírgula) (*.csv)", "CSV UTF-8 (separado por ponto e vírgula) (*.csv)"],
            JanelaEditor.LinhasSalvar(FormatoDocumento.Xlsx, permitirExportacao: false).Select(l => l.Rotulo));

    [Fact] // apresentacao: ppsx/ppsm/potx entram (A9)
    public void Linhas_Slides_TemApresentacaoDeSlidesEModelo() =>
        Assert.Contains("Apresentação de slides do PowerPoint (*.ppsx)", JanelaEditor.LinhasSalvar(FormatoDocumento.Pptx, true).Select(l => l.Rotulo));

    [Theory] // legado e XLSB (so' leitura) NUNCA aparecem no Salvar como, em nenhum tipo (FMT-08)
    [InlineData(FormatoDocumento.Docx)] [InlineData(FormatoDocumento.Xlsx)] [InlineData(FormatoDocumento.Pptx)]
    public void Linhas_NuncaOferecemLegadoNemXlsb(FormatoDocumento f)
    {
        foreach (var exportacao in new[] { true, false })
            Assert.DoesNotContain(JanelaEditor.LinhasSalvar(f, exportacao), l => l.Extensao is ".doc" or ".dot" or ".xls" or ".xlt" or ".xlsb" or ".ppt" or ".pps" or ".pot");
    }

    [Theory] // o formato vem da LINHA; extensao de documento do mesmo tipo e gravavel digitada vence; legado vira o da linha
    [InlineData(@"C:\x\a", 7, FormatoDocumento.Docx, @"C:\x\a.dotx", FormatoDocumento.Dotx)]
    [InlineData(@"C:\x\a.rtf", 1, FormatoDocumento.Docx, @"C:\x\a.rtf", FormatoDocumento.Rtf)]
    [InlineData(@"C:\x\a.doc", 7, FormatoDocumento.Docx, @"C:\x\a.dotx", FormatoDocumento.Dotx)]
    [InlineData(@"C:\x\s.ppsx", 1, FormatoDocumento.Pptx, @"C:\x\s.ppsx", FormatoDocumento.Ppsx)]
    public void Destino_LinhasNovas(string digitado, int indice, FormatoDocumento sessao, string caminho, FormatoDocumento esperado)
    {
        var d = JanelaEditor.DestinoDoDialogo(digitado, indice, sessao, permitirExportacao: true);
        Assert.Equal((caminho, esperado), (d.Caminho, d.Formato));
    }

    [Theory] // pre-selecao: sessao de modelo/macro abre o dialogo na linha dela
    [InlineData(FormatoDocumento.Dotx, 7)] [InlineData(FormatoDocumento.Docm, 6)] [InlineData(FormatoDocumento.Rtf, 10)]
    public void Indice_LinhasNovas(FormatoDocumento f, int esperado) => Assert.Equal(esperado, JanelaEditor.IndiceDoFiltro(f));

    [Theory] // FECHAMENTO x Salvar como: a linha pre-selecionada e' contada na lista que o dialogo MOSTRA (sem as 3 de PDF no fechamento)
    [InlineData(FormatoDocumento.Docx, true, 1)] [InlineData(FormatoDocumento.Docx, false, 1)]
    [InlineData(FormatoDocumento.Docm, true, 6)] [InlineData(FormatoDocumento.Docm, false, 3)]
    [InlineData(FormatoDocumento.Xlsm, true, 6)] [InlineData(FormatoDocumento.Xlsm, false, 3)]
    [InlineData(FormatoDocumento.Pptm, true, 6)] [InlineData(FormatoDocumento.Pptm, false, 3)]
    [InlineData(FormatoDocumento.Ppsx, true, 9)] [InlineData(FormatoDocumento.Ppsx, false, 6)]
    [InlineData(FormatoDocumento.Odt, true, 2)] [InlineData(FormatoDocumento.Odt, false, 2)]
    public void FiltroEIndice_LiteraisPorFormatoEFlag(FormatoDocumento f, bool exportacao, int esperado) =>
        Assert.Equal(esperado, JanelaEditor.FiltroEIndice(f, f, exportacao).Indice);

    [Fact] // para TODO formato gravavel de cada tipo e nos dois fluxos, a linha selecionada do filtro tem o formato da sessao (nunca outro tipo)
    public void FiltroEIndice_LinhaSelecionadaEhOFormatoDaSessao()
    {
        foreach (var i in TabelaFormatos.Todas.Where(i => i.Grava && !i.Legado && i.DocType != "pdf" && i.Abre))
            foreach (var exportacao in new[] { true, false })
            {
                var (filtro, indice) = JanelaEditor.FiltroEIndice(i.Formato, i.Formato, exportacao);
                var linhas = JanelaEditor.LinhasSalvar(i.Formato, exportacao);
                Assert.Equal(filtro, JanelaEditor.FiltroSalvar(i.Formato, exportacao));          // um so' filtro
                Assert.Equal(i.Formato, linhas[indice - 1].Formato);                              // a linha pre-selecionada
                Assert.Equal(i.Extensao, linhas[indice - 1].Extensao);
                var extensoesDoFiltro = filtro.Split('|').Where((_, k) => k % 2 == 1).ToList();
                Assert.Equal("*" + i.Extensao, extensoesDoFiltro[indice - 1]);
            }
    }

    [Fact] // .htm digitado e' a linha de Pagina da Web (alias), sem criar uma segunda linha no dialogo
    public void Htm_EhAliasDoHtml()
    {
        var d = JanelaEditor.DestinoDoDialogo(@"C:\x\a.htm", 1, FormatoDocumento.Docx, true);
        Assert.Equal((@"C:\x\a.htm", FormatoDocumento.Html), (d.Caminho, d.Formato));
        Assert.True(JanelaEditor.ExtensaoCombina(@"C:\x\a.HTM", FormatoDocumento.Html));
        Assert.Single(JanelaEditor.LinhasSalvar(FormatoDocumento.Docx, true), l => l.Formato == FormatoDocumento.Html);
    }

    [Fact] // a linha de HTML grava HTML ("Pagina da Web"); sessao legada pre-seleciona o moderno do tipo
    public void Html_GravaComoPaginaDaWeb()
    {
        var d = JanelaEditor.DestinoDoDialogo(@"C:\x\a", 12, FormatoDocumento.Docx, true);
        Assert.Equal((@"C:\x\a.html", FormatoDocumento.Html), (d.Caminho, d.Formato));
    }

    [Theory] // o formato vem do INDICE do filtro (as tres linhas de PDF tem a mesma extensao)
    [InlineData(@"C:\x\a.pdf", 3, FormatoDocumento.Pdf, false)]
    [InlineData(@"C:\x\a.pdf", 4, FormatoDocumento.PdfA, false)]
    [InlineData(@"C:\x\a.pdf", 5, FormatoDocumento.Pdf, true)]
    [InlineData(@"C:\x\a.docx", 1, FormatoDocumento.Docx, false)]
    [InlineData(@"C:\x\a.odt", 2, FormatoDocumento.Odt, false)]
    public void Destino_PeloIndice(string caminho, int indice, FormatoDocumento esperado, bool assinar)
    {
        var d = JanelaEditor.DestinoDoDialogo(caminho, indice, FormatoDocumento.Docx, permitirExportacao: true);
        Assert.Equal((esperado, assinar), (d.Formato, d.Assinar));
    }

    [Fact] // linha de PDF escolhida e nome sem .pdf: acrescenta .pdf (nunca grava PDF com extensao de documento)
    public void Destino_LinhaPdfSemExtensao_AcrescentaPdf() =>
        Assert.Equal(@"C:\x\peticao.pdf", JanelaEditor.DestinoDoDialogo(@"C:\x\peticao", 4, FormatoDocumento.Docx, true).Caminho);

    [Fact] // extensao de documento digitada com a linha PDF escolhida: vence a extensao (comportamento de hoje)
    public void Destino_ExtensaoDeDocumentoVence() =>
        Assert.Equal(FormatoDocumento.Odt, JanelaEditor.DestinoDoDialogo(@"C:\x\a.odt", 3, FormatoDocumento.Docx, true).Formato);

    [Theory] // linha "PDF assinado" (5) com extensao de documento digitada: grava o documento e leva o Assinar para o host
             // AVISAR que nada foi para assinar; nas linhas 3/4 nao ha' pedido de assinatura (revisao final do Plano 4)
    [InlineData(@"C:\x\a.odt", 5, FormatoDocumento.Odt, true)]
    [InlineData(@"C:\x\a.docx", 5, FormatoDocumento.Docx, true)]
    [InlineData(@"C:\x\a.odt", 3, FormatoDocumento.Odt, false)]
    [InlineData(@"C:\x\a.odt", 4, FormatoDocumento.Odt, false)]
    public void Destino_ExtensaoDeDocumento_NaLinhaAssinada_LevaOPedido(string caminho, int indice, FormatoDocumento fmt, bool assinar)
    {
        var d = JanelaEditor.DestinoDoDialogo(caminho, indice, FormatoDocumento.Docx, true);
        Assert.Equal((caminho, fmt, assinar), (d.Caminho, d.Formato, d.Assinar));
    }

    [Fact]
    public void MensagemAssinarSalvouDocumento_TemAExtensao() =>
        Assert.Equal("O arquivo foi salvo como documento (ODT), não como PDF: nada foi enviado para assinar.",
            JanelaEditor.MensagemAssinarSalvouDocumento(FormatoDocumento.Odt));

    [Theory] // pre-selecao pelo fileType do painel do editor
    [InlineData(FormatoDocumento.Pdf, false, 3)]
    [InlineData(FormatoDocumento.PdfA, false, 4)]
    [InlineData(FormatoDocumento.Pdf, true, 5)]
    [InlineData(FormatoDocumento.Ods, false, 2)]
    public void Indice_DoFormato(FormatoDocumento f, bool assinar, int esperado) =>
        Assert.Equal(esperado, JanelaEditor.IndiceDoFiltro(f, assinar));

    [Theory] // o fileType que o painel "Salvar como" do editor manda (c_oAscFileType.PDF/PDFA, medidos no bundle 9.4) cai na linha certa
    [InlineData(513, 3)]
    [InlineData(521, 4)]
    public void FileTypeDoPainel_PreSelecionaALinha(int fileType, int esperado) =>
        Assert.Equal(esperado, JanelaEditor.IndiceDoFiltro(PedidoSalvar.FormatoPorCodigo(fileType)!.Value));

    [Theory] // ida e volta: o indice que pre-seleciona um formato e' o mesmo que o destino le' de volta como esse formato
    [InlineData(FormatoDocumento.Docx, false)]
    [InlineData(FormatoDocumento.Odt, false)]
    [InlineData(FormatoDocumento.Pdf, false)]
    [InlineData(FormatoDocumento.PdfA, false)]
    [InlineData(FormatoDocumento.Pdf, true)]
    public void Indice_IdaEVolta(FormatoDocumento f, bool assinar)
    {
        var d = JanelaEditor.DestinoDoDialogo(@"C:\x\a." + f.Extensao(), JanelaEditor.IndiceDoFiltro(f, assinar), FormatoDocumento.Docx, true);
        Assert.Equal((f, assinar), (d.Formato, d.Assinar));
    }

    [Theory] // sem exportacao (fechamento) o indice nao cria PDF; nome sem extensao ganha a do formato
    [InlineData(@"C:\x\a.docx", 3, @"C:\x\a.docx", FormatoDocumento.Docx)]
    [InlineData(@"C:\x\a.odt", 5, @"C:\x\a.odt", FormatoDocumento.Odt)]
    [InlineData(@"C:\x\a", 1, @"C:\x\a.docx", FormatoDocumento.Docx)] // Plano 6: no fechamento so' ha linhas de documento (o indice 4 virou .docm)
    public void Destino_SemExportacao_SoPelaExtensao(string digitado, int indice, string caminho, FormatoDocumento esperado)
    {
        var d = JanelaEditor.DestinoDoDialogo(digitado, indice, FormatoDocumento.Docx, permitirExportacao: false);
        Assert.Equal((caminho, esperado, false), (d.Caminho, d.Formato, d.Assinar));
    }

    [Fact] // .PDF em maiusculas com a linha PDF/A: nao acrescenta outro .pdf
    public void Destino_PdfMaiusculo_NaoDuplicaExtensao() =>
        Assert.Equal(@"C:\x\A.PDF", JanelaEditor.DestinoDoDialogo(@"C:\x\A.PDF", 4, FormatoDocumento.Xlsx, true).Caminho);

    [Theory] // A2: extensao que o mOffice so' LE e' trocada pela da linha; extensao desconhecida ganha a da linha no fim
    [InlineData(@"C:\x\a.doc", 1, FormatoDocumento.Docx, @"C:\x\a.docx", FormatoDocumento.Docx)]
    [InlineData(@"C:\x\a.doc", 2, FormatoDocumento.Docx, @"C:\x\a.odt", FormatoDocumento.Odt)]
    [InlineData(@"C:\x\a.XLS", 1, FormatoDocumento.Xlsx, @"C:\x\a.xlsx", FormatoDocumento.Xlsx)]
    [InlineData(@"C:\x\Contrato v1.2", 1, FormatoDocumento.Docx, @"C:\x\Contrato v1.2.docx", FormatoDocumento.Docx)]
    [InlineData(@"C:\x\a.ppt", 2, FormatoDocumento.Pptx, @"C:\x\a.odp", FormatoDocumento.Odp)]
    public void Destino_ExtensaoDesconhecida_GanhaAExtensaoDaLinha(string digitado, int indice, FormatoDocumento sessao, string caminho, FormatoDocumento formato)
    {
        var d = JanelaEditor.DestinoDoDialogo(digitado, indice, sessao, permitirExportacao: true);
        Assert.Equal((caminho, formato), (d.Caminho, d.Formato));
    }

    [Theory] // extensao conhecida de OUTRO app (a.xlsx digitado no mTexto) nao vira formato de planilha: e' tratada como desconhecida
    [InlineData(@"C:\x\a.xlsx", 1, FormatoDocumento.Docx, @"C:\x\a.xlsx.docx", FormatoDocumento.Docx)]
    public void Destino_ExtensaoDeOutroApp_NaoMudaOTipo(string digitado, int indice, FormatoDocumento sessao, string caminho, FormatoDocumento formato)
    {
        var d = JanelaEditor.DestinoDoDialogo(digitado, indice, sessao, permitirExportacao: true);
        Assert.Equal((caminho, formato), (d.Caminho, d.Formato));
    }

    [Theory] // guarda de defesa: a extensao do destino tem de ser a do formato gravado
    [InlineData(@"C:\x\a.docx", FormatoDocumento.Docx, true)]
    [InlineData(@"C:\x\a.DOCX", FormatoDocumento.Docx, true)]
    [InlineData(@"C:\x\a.pdf", FormatoDocumento.PdfA, true)]
    [InlineData(@"C:\x\a.doc", FormatoDocumento.Docx, false)]
    [InlineData(@"C:\x\a.odt", FormatoDocumento.Docx, false)]
    public void ExtensaoCombina_SoComAExtensaoDoFormato(string destino, FormatoDocumento f, bool esperado) =>
        Assert.Equal(esperado, JanelaEditor.ExtensaoCombina(destino, f));

    [Fact]
    public void MensagemExtensaoNaoCombina_CitaArquivoEFormato() =>
        Assert.Equal("O arquivo \"a.doc\" não foi gravado: o conteúdo seria DOCX com a extensão .doc. Use Salvar como e escolha o tipo pela lista.",
            JanelaEditor.MensagemExtensaoNaoCombina(@"C:\x\a.doc", FormatoDocumento.Docx));

    // ---- Plano 6, Task 3: TXT e CSV no Salvar como ----

    [Fact] // planilha: duas linhas de CSV, ambas com ponto e virgula - ANSI (Excel pt-BR) e UTF-8
    public void Linhas_Planilha_TemCsvAnsiEUtf8()
    {
        var csv = JanelaEditor.LinhasSalvar(FormatoDocumento.Xlsx, true).Where(l => l.Formato == FormatoDocumento.Csv).ToList();
        Assert.Equal([OpcoesTexto.CsvExcelPtBr, OpcoesTexto.CsvUtf8], csv.Select(l => l.Texto));
        Assert.Equal(["CSV (separado por ponto e vírgula) (*.csv)", "CSV UTF-8 (separado por ponto e vírgula) (*.csv)"], csv.Select(l => l.Rotulo));
    }

    [Fact] // texto: "Texto sem formatação (*.txt)" em UTF-8
    public void Linhas_Texto_TemTxt() =>
        Assert.Contains(JanelaEditor.LinhasSalvar(FormatoDocumento.Docx, true), l => l.Formato == FormatoDocumento.Txt && l.Texto == OpcoesTexto.TxtUtf8);

    [Fact] // so' as linhas de texto levam opcoes: as demais (docx, pdf...) nao
    public void OutrasLinhas_SemOpcoesDeTexto() =>
        Assert.All(JanelaEditor.LinhasSalvar(FormatoDocumento.Docx, true).Where(l => l.Formato != FormatoDocumento.Txt), l => Assert.Null(l.Texto));

    [Fact] // a linha escolhida leva as opcoes para o destino
    public void Destino_LinhaCsv_LevaAsOpcoes()
    {
        var linhas = JanelaEditor.LinhasSalvar(FormatoDocumento.Xlsx, true);
        var i = linhas.ToList().FindIndex(l => l.Texto == OpcoesTexto.CsvUtf8) + 1;
        var d = JanelaEditor.DestinoDoDialogo(@"C:\x\a", i, FormatoDocumento.Xlsx, true);
        Assert.Equal((@"C:\x\a.csv", FormatoDocumento.Csv, OpcoesTexto.CsvUtf8), (d.Caminho, d.Formato, d.Texto));
    }

    [Fact] // CSV digitado com outra linha escolhida (a.csv na linha .xlsx): grava CSV com o padrao do Excel pt-BR, nunca sem opcoes
    public void Destino_CsvDigitado_ComLinhaDeOutroFormato_UsaOPadrao()
    {
        var d = JanelaEditor.DestinoDoDialogo(@"C:\x\a.csv", 1, FormatoDocumento.Xlsx, true);
        Assert.Equal((FormatoDocumento.Csv, OpcoesTexto.CsvExcelPtBr), (d.Formato, d.Texto));
        var t = JanelaEditor.DestinoDoDialogo(@"C:\x\a.txt", 1, FormatoDocumento.Docx, true);
        Assert.Equal((FormatoDocumento.Txt, OpcoesTexto.TxtUtf8), (t.Formato, t.Texto));
    }

    [Fact] // formato que nao e' de texto nunca leva opcoes
    public void Destino_Docx_SemOpcoes() =>
        Assert.Null(JanelaEditor.DestinoDoDialogo(@"C:\x\a.docx", 1, FormatoDocumento.Docx, true).Texto);

    [Fact] // a pre-selecao de uma sessao CSV/TXT cai na primeira linha do proprio formato
    public void Indice_SessaoCsvETxt()
    {
        Assert.Equal(JanelaEditor.LinhasSalvar(FormatoDocumento.Csv, true).ToList().FindIndex(l => l.Formato == FormatoDocumento.Csv) + 1, JanelaEditor.IndiceDoFiltro(FormatoDocumento.Csv));
        Assert.Equal(JanelaEditor.LinhasSalvar(FormatoDocumento.Txt, true).ToList().FindIndex(l => l.Formato == FormatoDocumento.Txt) + 1, JanelaEditor.IndiceDoFiltro(FormatoDocumento.Txt));
    }

    [Theory] // a mensagem da faixa a cada gravacao em formato que perde formatacao
    [InlineData(FormatoDocumento.Txt, "O arquivo foi salvo como TXT: formatação, imagens e tabelas não são guardadas nesse formato.")]
    [InlineData(FormatoDocumento.Csv, "O arquivo foi salvo como CSV: formatação, imagens e tabelas não são guardadas nesse formato.")]
    [InlineData(FormatoDocumento.Html, "O arquivo foi salvo como HTML: formatação, imagens e tabelas não são guardadas nesse formato.")]
    public void MensagemPerdeFormatacao_Literal(FormatoDocumento f, string esperado) => Assert.Equal(esperado, JanelaEditor.MensagemPerdeFormatacao(f));

    [Fact]
    public void MensagemCsvSoPlanilhaAtiva_Literal() =>
        Assert.Equal("O arquivo foi salvo como CSV: só a planilha ativa foi gravada.", JanelaEditor.MensagemCsvSoPlanilhaAtiva);
}
