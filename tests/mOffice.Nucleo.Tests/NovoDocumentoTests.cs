using mOffice.Conversor;
using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

[Collection(JanelaReal.Nome)]
public class NovoDocumentoTests
{
    [Theory] // o modelo vazio do motor existe para os tres OOXML e os ODF usam o irmao OOXML
    [InlineData(FormatoDocumento.Docx, "new.docx")]
    [InlineData(FormatoDocumento.Odt, "new.docx")]
    [InlineData(FormatoDocumento.Xlsx, "new.xlsx")]
    [InlineData(FormatoDocumento.Ods, "new.xlsx")]
    [InlineData(FormatoDocumento.Pptx, "new.pptx")]
    [InlineData(FormatoDocumento.Odp, "new.pptx")]
    public void ModeloVazio_NomeDoArquivoPorFormato(FormatoDocumento formato, string esperado)
        => Assert.Equal(esperado, ModelosVazios.NomeDoModelo(formato));

    [Fact] // PDF nao tem "novo documento" (o editor de PDF fica fora deste plano)
    public void ModeloVazio_Pdf_Lanca()
        => Assert.Throws<ArgumentOutOfRangeException>(() => ModelosVazios.NomeDoModelo(FormatoDocumento.Pdf));

    [Fact] // pt-BR primeiro, default como reserva, e erro legivel quando nenhum existe
    public void ModeloVazio_Caminho_PreferePtBrDepoisDefault()
    {
        var motor = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        var semProprios = Path.Combine(motor, "sem-modelos-proprios"); // os proprios do mOffice (Task 9) vem antes e existem no exe: aqui isolados
        var def = Path.Combine(motor, "converter", "empty", "default"); Directory.CreateDirectory(def);
        File.WriteAllBytes(Path.Combine(def, "new.docx"), [1]);
        Assert.Equal(Path.Combine(def, "new.docx"), ModelosVazios.Caminho(new ConfiguracaoMotor(motor), FormatoDocumento.Docx, semProprios));
        var pt = Path.Combine(motor, "converter", "empty", "pt-BR"); Directory.CreateDirectory(pt);
        File.WriteAllBytes(Path.Combine(pt, "new.docx"), [1]);
        Assert.Equal(Path.Combine(pt, "new.docx"), ModelosVazios.Caminho(new ConfiguracaoMotor(motor), FormatoDocumento.Docx, semProprios));
        var ex = Assert.Throws<FileNotFoundException>(() => ModelosVazios.Caminho(new ConfiguracaoMotor(motor), FormatoDocumento.Xlsx, semProprios));
        Assert.Contains("new.xlsx", ex.Message);
    }

    [Fact] // titulo de documento novo por tipo; com arquivo, o nome do arquivo
    public void Sessao_Titulo_NovoPorTipo()
    {
        var s = new SessaoDocumento { ArquivoOrigem = @"C:\x\novo.docx", Formato = FormatoDocumento.Odt, PastaRecovery = @"C:\x", Novo = true };
        Assert.Equal("Documento", s.Titulo);   // sessao criada a mao: sem numero
        Assert.Equal("Documento7", new SessaoDocumento { ArquivoOrigem = @"C:\x\novo.docx", Formato = FormatoDocumento.Odt, PastaRecovery = @"C:\x", Novo = true, NumeroDoNovo = 7 }.Titulo);
        s.Repontar(@"C:\y\peticao.odt", FormatoDocumento.Odt);
        Assert.False(s.Novo);
        Assert.Equal("peticao.odt", s.Titulo);
        Assert.Equal("Pasta", new SessaoDocumento { ArquivoOrigem = "a.xlsx", Formato = FormatoDocumento.Ods, PastaRecovery = "p", Novo = true }.Titulo);
        Assert.Equal("Apresentação", new SessaoDocumento { ArquivoOrigem = "a.pptx", Formato = FormatoDocumento.Pptx, PastaRecovery = "p", Novo = true }.Titulo);
    }

    [SkippableFact] // NovoAsync abre o modelo vazio: editor pronto, titulo "Documento — mTexto", sem asterisco; salvar-como grava DOCX legivel
    public void NovoAsync_Docx_AbreEditorESalvaComo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pasta);
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { NomeApp = "mTexto" });
            j.Show();
            await j.NovoAsync(FormatoDocumento.Docx);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            Assert.True(j.SessaoAtual!.Novo);
            Assert.True(j.SessaoAtual.NumeroDoNovo >= 1);
            Assert.Equal($"Documento{j.SessaoAtual.NumeroDoNovo} — mTexto", j.Title);
            await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarAsync());
            await j.ExecutarJsAsync(EditorReal.InserirTextoWord("texto novo"));
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 50));  // espelho JS do Modificado
            Assert.Equal($"Documento{j.SessaoAtual.NumeroDoNovo}* — mTexto", j.Title);
            var destino = Path.Combine(pasta, "novo.docx");
            await j.SalvarComoAsync(destino, FormatoDocumento.Docx);
            Assert.False(j.SessaoAtual!.Novo);
            Assert.Equal("novo.docx — mTexto", j.Title);
            Assert.Contains("texto novo", FabricaDocumentos.TextoDocx(destino));
            j.Close();
        // 180s (achado 13 da revisao 1): este teste abre o editor, edita, pede o asc_Save e roda o x2t de
        // volta para DOCX — o padrao de 90s de JanelaSta.Executar e' apertado numa maquina fria.
        }, TimeSpan.FromSeconds(180));
    }

    [SkippableFact] // ODT novo: modelo OOXML por baixo, mas Formato = Odt e salvar-como padrao em ODT
    public void NovoAsync_Odt_GuardaFormatoOdt()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        JanelaSta.Executar(async () =>
        {
            var j = EditorReal.PrepararParaTeste(new JanelaEditor("mtexto") { NomeApp = "mTexto" });
            j.Show();
            await j.NovoAsync(FormatoDocumento.Odt);
            Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 150));
            Assert.Equal(FormatoDocumento.Odt, j.SessaoAtual!.Formato);
            Assert.EndsWith("new.docx", j.SessaoAtual.ArquivoOrigem, StringComparison.OrdinalIgnoreCase);
            j.Close();
        });
    }
}
