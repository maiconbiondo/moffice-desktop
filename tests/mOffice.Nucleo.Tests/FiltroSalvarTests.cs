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
        Assert.EndsWith("|PDF (*.pdf)|*.pdf|PDF/A (*.pdf)|*.pdf|PDF assinado com mPDF (*.pdf)|*.pdf", filtro);
        Assert.Equal(10, filtro.Split('|').Length);
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

    [Fact] // com exportacao: OOXML, ODF, PDF, PDF/A, PDF assinado — nesta ordem
    public void Filtro_ComExportacao_TemPdfAEAssinado() =>
        Assert.Equal("Documento do Word (*.docx)|*.docx|Documento de texto ODF (*.odt)|*.odt|PDF (*.pdf)|*.pdf|PDF/A (*.pdf)|*.pdf|PDF assinado com mPDF (*.pdf)|*.pdf",
            JanelaEditor.FiltroSalvar(FormatoDocumento.Docx));

    [Fact] // no fechamento (sem exportacao) nada de PDF — o texto de hoje nao muda
    public void Filtro_SemExportacao_Inalterado() =>
        Assert.Equal("Pasta de trabalho do Excel (*.xlsx)|*.xlsx|Planilha ODF (*.ods)|*.ods",
            JanelaEditor.FiltroSalvar(FormatoDocumento.Xlsx, permitirExportacao: false));

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

    [Theory] // sem exportacao (fechamento) o indice nao cria PDF: tudo pela extensao, como antes
    [InlineData(@"C:\x\a.docx", 3, FormatoDocumento.Docx)]
    [InlineData(@"C:\x\a.odt", 5, FormatoDocumento.Odt)]
    [InlineData(@"C:\x\a", 4, FormatoDocumento.Docx)]
    public void Destino_SemExportacao_SoPelaExtensao(string caminho, int indice, FormatoDocumento esperado)
    {
        var d = JanelaEditor.DestinoDoDialogo(caminho, indice, FormatoDocumento.Docx, permitirExportacao: false);
        Assert.Equal((caminho, esperado, false), (d.Caminho, d.Formato, d.Assinar));
    }

    [Fact] // .PDF em maiusculas com a linha PDF/A: nao acrescenta outro .pdf
    public void Destino_PdfMaiusculo_NaoDuplicaExtensao() =>
        Assert.Equal(@"C:\x\A.PDF", JanelaEditor.DestinoDoDialogo(@"C:\x\A.PDF", 4, FormatoDocumento.Xlsx, true).Caminho);
}
