using mOffice.Comum;

namespace mOffice.Comum.Tests;

/// <summary>Plano 6, Task 3 (FMT-07): o mTexto abre PDF como documento editavel — so' pelo Abrir/soltar; o .pdf continua sem
/// ProgID (o mPDF e' o dono) e fora da linha de comando e dos recentes.</summary>
public class PdfNoMTextoTests
{
    [Fact] // o filtro do Abrir do mTexto ganha a linha do PDF logo antes de "Todos os arquivos"
    public void FiltroAbrir_Texto_TemLinhaDoPdf() =>
        Assert.EndsWith("|PDF (abrir como documento editável) (*.pdf)|*.pdf|Todos os arquivos|*.*", PerfilApp.Texto.FiltroAbrir);

    [Fact] // a linha "Todos os documentos" nao inclui PDF (um duplo clique em .pdf nunca e' do mTexto) e o resto do filtro nao muda de lugar
    public void FiltroAbrir_Texto_TodosOsDocumentosSemPdf()
    {
        var todos = PerfilApp.Texto.FiltroAbrir.Split('|')[1];
        Assert.DoesNotContain("pdf", todos);
        Assert.Contains("*.txt", todos);
    }

    [Theory] // planilha e slides nao abrem PDF
    [InlineData("mplanilha")] [InlineData("mslides")]
    public void FiltroAbrir_OutrosApps_SemPdf(string app) =>
        Assert.DoesNotContain("pdf", (app == "mplanilha" ? PerfilApp.Planilha : PerfilApp.Slides).FiltroAbrir, StringComparison.OrdinalIgnoreCase);

    [Fact] // TXT no mTexto e CSV no mPlanilha entram no filtro (da tabela)
    public void FiltroAbrir_TxtECsv()
    {
        Assert.Contains("Texto sem formatação (*.txt)|*.txt", PerfilApp.Texto.FiltroAbrir);
        Assert.Contains("CSV (separado por ponto e vírgula) (*.csv)|*.csv", PerfilApp.Planilha.FiltroAbrir);
    }

    [Fact] // so' o mTexto abre PDF como documento
    public void AbrePdfComoDocumento_SoOMTexto()
    {
        Assert.True(PerfilApp.Texto.AbrePdfComoDocumento(@"C:\x\a.pdf"));
        Assert.True(PerfilApp.Texto.AbrePdfComoDocumento(@"C:\x\A.PDF"));
        Assert.False(PerfilApp.Texto.AbrePdfComoDocumento(@"C:\x\a.docx"));
        Assert.False(PerfilApp.Planilha.AbrePdfComoDocumento(@"C:\x\a.pdf"));
        Assert.False(PerfilApp.Slides.AbrePdfComoDocumento(@"C:\x\a.pdf"));
    }

    [Fact] // o gate geral (linha de comando, pipe, recentes) continua recusando .pdf; so' o soltar/Abrir o aceita
    public void Aceita_Pdf_ContinuaFalso_PrimeiroAceitoPeloSoltarAceita()
    {
        Assert.False(PerfilApp.Texto.Aceita(@"C:\x\a.pdf"));
        Assert.Equal(@"C:\x\b.pdf", PerfilApp.Texto.PrimeiroAceito([@"C:\x\a.xlsx", @"C:\x\b.pdf"]));
        Assert.Null(PerfilApp.Planilha.PrimeiroAceito([@"C:\x\b.pdf"]));
    }

    [Fact] // Abrir com / linha de comando do Explorer: .pdf nao e' aceito pelo mTexto (o mPDF e' o dono)
    public void LinhaDeComando_PdfRejeitado()
    {
        Assert.Null(PedidoDeAbertura.DosArgumentos([@"C:\x\a.pdf"], PerfilApp.Texto, out var rejeitado));
        Assert.Equal(@"C:\x\a.pdf", rejeitado);
    }
}
