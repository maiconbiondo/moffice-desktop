using mOffice.Comum;
using mOffice.Conversor;

namespace mOffice.Comum.Tests;

public class PedidoDeAberturaTests
{
    [Fact] // "%1" simples = abrir
    public void Caminho_Abrir()
    {
        var p = PedidoDeAbertura.DosArgumentos([@"C:\x\a.doc"], PerfilApp.Texto, out var rejeitado);
        Assert.Equal(new PedidoDeAbertura(@"C:\x\a.doc", ModoAbertura.Abrir), p);
        Assert.Null(rejeitado);
    }

    [Fact] // /modelo "%1" (verbo open do .dotx no Explorer) = novo a partir do modelo
    public void Modelo_NovoDoModelo() =>
        Assert.Equal(ModoAbertura.NovoDoModelo, PedidoDeAbertura.DosArgumentos(["/modelo", @"C:\x\m.dotx"], PerfilApp.Texto, out _)!.Modo);

    [Fact] // /apresentar "%1" (verbo open do .ppsx) = abrir e apresentar
    public void Apresentar() =>
        Assert.Equal(ModoAbertura.Apresentar, PedidoDeAbertura.DosArgumentos(["/apresentar", @"C:\x\s.ppsx"], PerfilApp.Slides, out _)!.Modo);

    [Fact] // /imprimir "%1" (verbo print do Explorer, Plano 6 Task 10) = imprimir; extensao que o app nao abre nao vira pedido
    public void Imprimir()
    {
        var p = PedidoDeAbertura.DosArgumentos(["/imprimir", @"C:\x\a.docx"], PerfilApp.Texto, out var rejeitado);
        Assert.Equal(new PedidoDeAbertura(@"C:\x\a.docx", ModoAbertura.Imprimir), p);
        Assert.Null(rejeitado);
        Assert.Null(PedidoDeAbertura.DosArgumentos(["/imprimir", @"C:\x\a.xlsx"], PerfilApp.Texto, out var rej2));
        Assert.Equal(@"C:\x\a.xlsx", rej2);
    }

    [Fact] // imprimir NAO vai pelo pipe (processo efemero): ParaPipe lanca e a linha ?imprimir|... nao existe
    public void Imprimir_NaoVaiPeloPipe()
    {
        Assert.Throws<InvalidOperationException>(() => new PedidoDeAbertura(@"C:\x\a.docx", ModoAbertura.Imprimir).ParaPipe());
        Assert.Null(PedidoDeAbertura.DoPipe(@"?imprimir|C:\x\a.docx"));
    }

    [Fact] // FMT-05: extensao que o app nao abre volta como rejeitada (a casca mostra a mensagem) em vez de sumir
    public void NaoAceita_VoltaComoRejeitado()
    {
        Assert.Null(PedidoDeAbertura.DosArgumentos([@"C:\x\a.xlsx"], PerfilApp.Texto, out var rejeitado));
        Assert.Equal(@"C:\x\a.xlsx", rejeitado);
    }

    [Theory] // ida e volta pelo pipe da instancia unica
    [InlineData(@"C:\x\a.doc", ModoAbertura.Abrir, @"C:\x\a.doc")]
    [InlineData(@"C:\x\m.dotx", ModoAbertura.NovoDoModelo, @"?modelo|C:\x\m.dotx")]
    [InlineData(@"\srv\p\s.ppsx", ModoAbertura.Apresentar, @"?apresentar|\srv\p\s.ppsx")]
    public void Pipe_IdaEVolta(string caminho, ModoAbertura modo, string linha)
    {
        var p = new PedidoDeAbertura(caminho, modo);
        Assert.Equal(linha, p.ParaPipe());
        Assert.Equal(p, PedidoDeAbertura.DoPipe(linha));
    }

    [Theory] // linha malformada do pipe: nunca vira pedido (caminho relativo, modo inventado, vazio)
    [InlineData("a.docx")] [InlineData(@"?formatar|C:\x\a.docx")] [InlineData("?modelo|a.dotx")] [InlineData("")]
    public void Pipe_Malformado_Nulo(string linha) => Assert.Null(PedidoDeAbertura.DoPipe(linha));

    [Fact] // FMT-05 na 2a instancia: a linha ?rejeitado|<abs> vai e volta; relativa/outra coisa nao vira rejeicao
    public void Pipe_Rejeicao_IdaEVolta()
    {
        Assert.Equal(@"?rejeitado|C:\x\a.pdf", PedidoDeAbertura.RejeicaoParaPipe(@"C:\x\a.pdf"));
        Assert.Equal(@"C:\x\a.pdf", PedidoDeAbertura.DoPipeRejeitado(@"?rejeitado|C:\x\a.pdf"));
        Assert.Null(PedidoDeAbertura.DoPipeRejeitado("?rejeitado|a.pdf"));
        Assert.Null(PedidoDeAbertura.DoPipeRejeitado(@"C:\x\a.pdf"));
        Assert.Null(PedidoDeAbertura.DoPipe(@"?rejeitado|C:\x\a.pdf"));   // nunca vira pedido de abrir
    }

    [Fact] // o subtitulo vem da tabela (sem modelo/macro/apresentacao)
    public void Subtitulo_VemDaTabela()
    {
        // Plano 6, Task 3: TXT e CSV abrem (com o dialogo de importacao) e entram no subtitulo
        Assert.Equal("Editor de textos · DOCX, DOC, ODT, RTF e TXT", PerfilApp.Texto.Subtitulo);
        Assert.Equal("Editor de planilhas · XLSX, XLS, XLSB, ODS e CSV", PerfilApp.Planilha.Subtitulo);
        Assert.Equal("Editor de apresentações · PPTX, PPT e ODP", PerfilApp.Slides.Subtitulo);
    }

    [Fact] // mensagem do rejeitado, montada pelo perfil
    public void MensagemNaoAbre() =>
        Assert.Equal("O mTexto não abre arquivos .XLSX: abra-o pelo aplicativo do mOffice correspondente.", PerfilApp.Texto.MensagemNaoAbre(@"C:\x\a.xlsx"));

    [Fact] // o verbo vem ANTES ou DEPOIS do caminho? O Explorer manda /modelo "%1": o verbo antes. Sem verbo antes, o modo e' Abrir
    public void VerboDepoisDoCaminho_NaoMudaOModo() =>
        Assert.Equal(ModoAbertura.Abrir, PedidoDeAbertura.DosArgumentos([@"C:\x\m.dotx", "/modelo"], PerfilApp.Texto, out _)!.Modo);

    [Fact] // o verbo e' reconhecido sem diferenciar caixa
    public void Verbo_IgnoraCaixa() =>
        Assert.Equal(ModoAbertura.NovoDoModelo, PedidoDeAbertura.DosArgumentos(["/MODELO", @"C:\x\m.dotx"], PerfilApp.Texto, out _)!.Modo);

    [Fact] // o primeiro rejeitado e' lembrado, mas um aceito depois dele ainda abre
    public void RejeitadoEAceito_AbreOAceito()
    {
        var p = PedidoDeAbertura.DosArgumentos([@"C:\x\a.xlsx", @"C:\x\b.docx"], PerfilApp.Texto, out var rejeitado);
        Assert.Equal(@"C:\x\b.docx", p!.Caminho);
        Assert.Equal(@"C:\x\a.xlsx", rejeitado);
    }

    [Theory] // as extensoes novas passam pelo gate de cada app; PDF nao abre em nenhum
    [InlineData("mtexto", ".dotx", true)] [InlineData("mtexto", ".rtf", true)] [InlineData("mtexto", ".doc", true)] [InlineData("mtexto", ".xlsm", false)]
    [InlineData("mtexto", ".pdf", false)] [InlineData("mtexto", ".html", false)] [InlineData("mtexto", ".txt", true)]
    [InlineData("mplanilha", ".xlsb", true)] [InlineData("mplanilha", ".xltx", true)] [InlineData("mplanilha", ".csv", true)] [InlineData("mplanilha", ".txt", false)] [InlineData("mtexto", ".csv", false)]
    [InlineData("mslides", ".ppsx", true)] [InlineData("mslides", ".pot", true)] [InlineData("mslides", ".potm", true)] [InlineData("mslides", ".docx", false)]
    public void Perfil_Aceita(string app, string ext, bool esperado)
    {
        var perfil = app switch { "mtexto" => PerfilApp.Texto, "mplanilha" => PerfilApp.Planilha, _ => PerfilApp.Slides };
        Assert.Equal(esperado, perfil.Aceita(@"C:\x\a" + ext));
    }

    [Fact] // filtro do Abrir: "Todos os documentos" primeiro, uma linha por extensao que abre, "Todos os arquivos" no fim
    public void FiltroAbrir_VemDaTabela()
    {
        var f = PerfilApp.Texto.FiltroAbrir;
        Assert.StartsWith("Todos os documentos do mTexto|*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.odt;*.ott;*.rtf;*.txt|", f);
        Assert.Contains("|Modelo do Word (*.dotx)|*.dotx|", f);
        Assert.EndsWith("|Todos os arquivos|*.*", f);
        // PDF so' na linha propria "abrir como documento editavel" (Task 3), nunca em "Todos os documentos"; PDF nao e' um Formato do app
        Assert.DoesNotContain("pdf", f.Split('|')[1], StringComparison.OrdinalIgnoreCase);
        Assert.Contains("|PDF (abrir como documento editável) (*.pdf)|*.pdf|Todos os arquivos|*.*", f);
        Assert.Equal([FormatoDocumento.Docx, FormatoDocumento.Docm, FormatoDocumento.Dotx, FormatoDocumento.Dotm, FormatoDocumento.Doc,
                      FormatoDocumento.Odt, FormatoDocumento.Ott, FormatoDocumento.Rtf, FormatoDocumento.Txt], PerfilApp.Texto.Formatos);
    }
}
