using System.IO.Compression;
using mOffice.Conversor;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Nucleo.Tests;

/// <summary>Plano 6, Task 2 — as pecas PURAS do Modo de compatibilidade, das macros e da abertura de modelos (sem janela, rodam em RDP).</summary>
public class FormatosLegadosTests
{
    private static string Pasta()
    {
        var p = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(p);
        return p;
    }

    private static SessaoDocumento Sessao(bool novo = false, bool compat = false, bool desvinculada = false) =>
        new() { ArquivoOrigem = @"C:\x\a.doc", Formato = FormatoDocumento.Docx, PastaRecovery = Pasta(), Novo = novo, Compatibilidade = compat, Desvinculada = desvinculada };

    [Theory] // o "Salvar" pergunta o destino em documento novo, recuperado desvinculado e em Modo de compatibilidade
    [InlineData(false, false, false, false)]
    [InlineData(true, false, false, true)]
    [InlineData(false, true, false, true)]
    [InlineData(false, false, true, true)]
    public void ExigeSalvarComo_PorEstadoDaSessao(bool novo, bool compat, bool desvinculada, bool esperado) =>
        Assert.Equal(esperado, Sessao(novo, compat, desvinculada).ExigeSalvarComo);

    [Fact] // depois do primeiro Salvar como, a sessao aponta para o arquivo moderno e sai do Modo de compatibilidade
    public void Repontar_ZeraACompatibilidade()
    {
        var s = Sessao(compat: true);
        Assert.True(s.Compatibilidade);
        s.Repontar(@"C:\x\a.docx", FormatoDocumento.Docx);
        Assert.False(s.Compatibilidade);
        Assert.False(s.ExigeSalvarComo);
        Assert.Equal(@"C:\x\a.docx", s.ArquivoOrigem);
    }

    [Fact] // o sessao.json de uma sessao legada leva Compatibilidade; o de antes (sem o campo) continua legivel e vale false
    public void Metadados_Compatibilidade_IdaEVolta_ECompativelComArquivoAntigo()
    {
        var p = Pasta();
        File.WriteAllText(Path.Combine(p, MetadadosSessao.NomeArquivo),
            """{"Origem":"C:\\x\\a.doc","Formato":65,"Titulo":"a.doc","Novo":false,"App":"mtexto","IdUsuario":3}""");
        Assert.False(MetadadosSessao.Ler(p)!.Compatibilidade);
        (MetadadosSessao.Ler(p)! with { Compatibilidade = true }).Gravar(p);
        Assert.True(MetadadosSessao.Ler(p)!.Compatibilidade);
        Assert.Contains("\"Compatibilidade\":true", File.ReadAllText(Path.Combine(p, MetadadosSessao.NomeArquivo)));
    }

    [Fact] // Repontar regrava o sessao.json SEM a compatibilidade (uma queda depois do Salvar como reabre no arquivo novo)
    public void Repontar_RegravaMetadadosSemCompatibilidade()
    {
        var s = Sessao(compat: true);
        s.Metadados = new MetadadosSessao(@"C:\x\a.doc", (int)FormatoDocumento.Docx, "a.doc", false, "mtexto", 1) { Compatibilidade = true };
        s.Repontar(@"C:\x\a.docx", FormatoDocumento.Docx);
        var lido = MetadadosSessao.Ler(s.PastaRecovery)!;
        Assert.False(lido.Compatibilidade);
        Assert.Equal(@"C:\x\a.docx", lido.Origem);
    }

    [Theory] // a faixa nomeia o app do tipo e a extensao do arquivo (uma linha da tabela, nao um texto por extensao)
    [InlineData(".doc", "Word", ".docx")] [InlineData(".dot", "Word", ".docx")]
    [InlineData(".xls", "Excel", ".xlsx")] [InlineData(".xlt", "Excel", ".xlsx")] [InlineData(".xlsb", "Excel", ".xlsx")]
    [InlineData(".ppt", "PowerPoint", ".pptx")] [InlineData(".pps", "PowerPoint", ".pptx")] [InlineData(".pot", "PowerPoint", ".pptx")]
    public void MensagemCompatibilidade_NomeiaAppEExtensoes(string legado, string app, string moderno) =>
        Assert.Equal($"Este documento está no formato antigo do {app} ({legado}). Ao salvar, o mOffice grava uma cópia em {moderno}; o arquivo original não é alterado.",
            JanelaEditor.MensagemCompatibilidade(TabelaFormatos.DaExtensao(legado)!));

    [Fact] // a mensagem de macros e' a do dono (texto exato)
    public void MensagemMacros_TextoDoDono() =>
        Assert.Equal("Este documento contém macros do Office (VBA). Elas são mantidas no arquivo, mas não são executadas no mOffice.", JanelaEditor.MensagemMacros);

    private static string Zip(string arquivo, params string[] entradas)
    {
        using var z = ZipFile.Open(arquivo, ZipArchiveMode.Create);
        foreach (var e in entradas) { using var w = new StreamWriter(z.CreateEntry(e).Open()); w.Write("x"); }
        return arquivo;
    }

    [Theory] // TemVba: o vbaProject.bin em word/, xl/ ou ppt/; pacote sem ele = false
    [InlineData("word/vbaProject.bin", true)] [InlineData("xl/vbaProject.bin", true)] [InlineData("ppt/vbaProject.bin", true)]
    [InlineData("word/document.xml", false)]
    public void TemVba_ProcuraOVbaProject(string entrada, bool esperado)
    {
        var a = Zip(Path.Combine(Pasta(), "a.docm"), "[Content_Types].xml", entrada);
        Assert.Equal(esperado, JanelaEditor.TemVba(a));
    }

    [Fact] // zip invalido e arquivo que nao existe = false (so' decide se a faixa de macros aparece; nunca lanca)
    public void TemVba_ArquivoRuim_Falso()
    {
        var p = Pasta();
        var lixo = Path.Combine(p, "lixo.docm"); File.WriteAllText(lixo, "nao e zip");
        Assert.False(JanelaEditor.TemVba(lixo));
        Assert.False(JanelaEditor.TemVba(Path.Combine(p, "nao-existe.docm")));
    }

    [Fact] // soltar no editor: todas as extensoes que algum app abre (a casca filtra pelo perfil) + o PDF (so' o mTexto o abre, como documento), e nada de HTML/imagem
    public void ExtensoesDocumentoSolto_VemDaTabela()
    {
        foreach (var e in new[] { ".doc", ".dotx", ".rtf", ".xlsb", ".xltm", ".ppsx", ".pot", ".odp", ".docx", ".txt", ".csv", ".pdf" })
            Assert.Contains(e, JanelaEditor.ExtensoesDocumentoSolto);
        foreach (var e in new[] { ".html", ".png" })
            Assert.DoesNotContain(e, JanelaEditor.ExtensoesDocumentoSolto);
    }

    [Theory] // MEDIDO ao vivo: o web-apps INICIA a apresentacao sozinho quando enxerga ppsx/ppsm (um .ppsx aberto pelo Abrir comum
             // entrava em tela cheia) — a URL leva o irmao de edicao (pptx/pptm) E a ponte entrega o caminho do irmao (OrigemParaOEditor);
             // a apresentacao e' pedida a parte, pela ponte, so' no modo Apresentar. Os demais formatos vao com a propria extensao.
    [InlineData(FormatoDocumento.Ppsx, "pptx")] [InlineData(FormatoDocumento.Ppsm, "pptm")]
    [InlineData(FormatoDocumento.Pptx, "pptx")] [InlineData(FormatoDocumento.Dotx, "dotx")] [InlineData(FormatoDocumento.Docm, "docm")]
    [InlineData(FormatoDocumento.Rtf, "rtf")] [InlineData(FormatoDocumento.Xltx, "xltx")] [InlineData(FormatoDocumento.Potm, "potm")]
    public void UrlDeEntrada_PpsxAbreParaEdicao(FormatoDocumento f, string filetype) =>
        Assert.Contains("&filetype=" + filetype + "&", JanelaEditor.MontarUrlDeEntrada(f, "a", TemaEditor.Claro));

    [Theory] // o 2o ponto do mesmo achado: o caminho que o editor le' (LocalFileGetSourcePath) leva o irmao de edicao; o resto, o arquivo real
    [InlineData(@"C:\x\show.ppsx", FormatoDocumento.Ppsx, @"C:\x\show.pptx")]
    [InlineData(@"C:\x\SHOW.PPSM", FormatoDocumento.Ppsm, @"C:\x\SHOW.pptm")]
    [InlineData(@"C:\x\a.docx", FormatoDocumento.Docx, @"C:\x\a.docx")]
    [InlineData(@"C:\x\m.dotx", FormatoDocumento.Dotx, @"C:\x\m.dotx")]
    [InlineData(@"C:\x\a.pptm", FormatoDocumento.Pptm, @"C:\x\a.pptm")]
    public void OrigemParaOEditor_SoOsPpsTrocam(string origem, FormatoDocumento f, string esperado)
    {
        var s = new SessaoDocumento { ArquivoOrigem = origem, Formato = f, PastaRecovery = Pasta() };
        Assert.Equal(esperado, PonteDesktop.OrigemParaOEditor(s));
        Assert.Contains(System.Text.Json.JsonSerializer.Serialize(esperado).Trim('"'), new PonteDesktop(s).SessaoJson());   // e' o que a ponte entrega
        Assert.Equal(origem, s.ArquivoOrigem);   // o arquivo REAL da sessao (o do salvar) nao muda
    }

    [Theory] // I1 da revisao final do Plano 6 (A10): documento NOVO — a origem da sessao e' a copia do modelo na recovery ("new.docx"), e o sdkjs
             // le' LocalFileGetSourcePath para o nome do cabecalho depois da 1a edicao. O editor ve' "<Titulo>" (sem extensao, como o Word) na mesma pasta.
    [InlineData("new.docx", FormatoDocumento.Docx, 1, null, "Documento1")]
    [InlineData("new.xlsx", FormatoDocumento.Xlsx, 2, null, "Pasta2")]
    [InlineData("new.pptx", FormatoDocumento.Pptx, 1, null, "Apresentação1")]
    [InlineData("Modelo Ata.dotx", FormatoDocumento.Docx, 3, null, "Documento3")]
    [InlineData("conv.docx", FormatoDocumento.Docx, 0, "relatorio", "relatorio")]
    public void OrigemParaOEditor_Novo_TituloDaSessao(string copia, FormatoDocumento f, int numero, string? tituloDoNovo, string esperado)
    {
        var rec = Pasta();
        var s = new SessaoDocumento { ArquivoOrigem = Path.Combine(rec, copia), Formato = f, PastaRecovery = rec, Novo = true, NumeroDoNovo = numero, TituloDoNovo = tituloDoNovo };
        Assert.Equal(Path.Combine(rec, esperado), PonteDesktop.OrigemParaOEditor(s));
        var json = new PonteDesktop(s).SessaoJson();
        Assert.DoesNotContain(Path.GetFileNameWithoutExtension(copia).Replace(" ", ""), json.Replace(" ", ""), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(Path.Combine(rec, copia), s.ArquivoOrigem);   // o host continua com a copia real (x2t/salvar)
    }

    [Fact] // I1: a reabertura de uma sessao RECUPERADA (LocalFileSetSourcePath antes do EndLoad) usa a mesma regra — senao o novo recuperado
           // volta a "new.docx" e o .ppsx recuperado volta a iniciar a apresentacao sozinho. Varredura da fonte (o script so' existe com janela).
    public void Recuperacao_SetSourcePath_UsaOrigemParaOEditor()
    {
        var fonte = File.ReadAllText(Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo", "JanelaEditor.xaml.cs"));
        var chamadas = System.Text.RegularExpressions.Regex.Matches(fonte, @"LocalFileSetSourcePath\(""\s*\+\s*([^;]+?)\s*\+\s*""\)");
        Assert.NotEmpty(chamadas);
        Assert.All(chamadas, m => Assert.Contains("PonteDesktop.OrigemParaOEditor(", m.Groups[1].Value));
    }

    [Fact] // guarda estrutural: o Ctrl+S do editor e o "Salvar" do fechamento decidem por ExigeSalvarComo (legado nunca grava por cima do .doc);
           // sem isto o Ctrl+S tentaria gravar o .doc e so' a guarda da extensao (Task 1) recusaria — e o seletor nunca apareceria
    public void CtrlSEFechamento_UsamExigeSalvarComo()
    {
        var fonte = File.ReadAllText(Path.Combine(FronteiraTests.RaizRepo(), "src", "mOffice.Nucleo", "JanelaEditor.xaml.cs")).Replace("\r\n", "\n");
        Assert.Contains("if (_sessao.ExigeSalvarComo) comoPedido = true;", fonte);
        Assert.Contains("DestinoSalvar? alvo = sessao.ExigeSalvarComo", fonte);
        Assert.DoesNotContain("_sessao.Novo || _sessao.Desvinculada", fonte);
        Assert.DoesNotContain("sessao.Novo || sessao.Desvinculada", fonte);
    }

    [Theory] // o destino digitado .doc/.xls/.ppt nunca combina com conteudo moderno (a guarda da Task 1 integrada com a tabela)
    [InlineData(@"C:\x\a.doc", FormatoDocumento.Docx)] [InlineData(@"C:\x\a.xls", FormatoDocumento.Xlsx)]
    [InlineData(@"C:\x\a.ppt", FormatoDocumento.Pptx)] [InlineData(@"C:\x\a.xlsb", FormatoDocumento.Xlsx)]
    [InlineData(@"C:\x\a.dotx", FormatoDocumento.Docx)]   // modelo e documento sao formatos diferentes
    public void ExtensaoLegada_NaoCombinaComConteudoModerno(string destino, FormatoDocumento f) =>
        Assert.False(JanelaEditor.ExtensaoCombina(destino, f));

    [Theory] // e as extensoes novas combinam com o proprio formato
    [InlineData(@"C:\x\a.docm", FormatoDocumento.Docm)] [InlineData(@"C:\x\a.DOTX", FormatoDocumento.Dotx)] [InlineData(@"C:\x\a.rtf", FormatoDocumento.Rtf)]
    [InlineData(@"C:\x\a.html", FormatoDocumento.Html)] [InlineData(@"C:\x\a.ppsx", FormatoDocumento.Ppsx)] [InlineData(@"C:\x\a.xlsm", FormatoDocumento.Xlsm)]
    public void ExtensaoNova_CombinaComOFormato(string destino, FormatoDocumento f) => Assert.True(JanelaEditor.ExtensaoCombina(destino, f));

    public static TheoryData<string> CondicoesDeSalvarComo => new() { "novo", "desvinculada", "somenteLeitura", "compatibilidade", "textoSemOpcoes" };

    [Theory] // M3: toda condicao de ExigeSalvarComo tem motivo pt-BR em MotivoParaExigirSalvarComo (a guarda do SalvarAsync); sessao comum: null
    [MemberData(nameof(CondicoesDeSalvarComo))]
    public void MotivoParaExigirSalvarComo_EspelhaExigeSalvarComo(string condicao)
    {
        var s = new SessaoDocumento
        {
            ArquivoOrigem = @"C:\x\a.docx", Formato = FormatoDocumento.Docx, PastaRecovery = Pasta(),
            Novo = condicao == "novo", Compatibilidade = condicao == "compatibilidade",
        };
        if (condicao == "desvinculada") s.Desvinculada = true;
        if (condicao == "somenteLeitura") s.SomenteLeitura = true;
        if (condicao == "textoSemOpcoes") s.OpcoesTextoDesconhecidas = true;
        Assert.True(s.ExigeSalvarComo);
        var motivo = JanelaEditor.MotivoParaExigirSalvarComo(s);
        Assert.False(string.IsNullOrWhiteSpace(motivo));
        Assert.NotEqual(JanelaEditor.MensagemNovoExigeSalvarComo, condicao == "novo" ? "" : motivo);   // cada condicao com o seu texto
    }

    [Fact]
    public void MotivoParaExigirSalvarComo_SessaoComum_Null() =>
        Assert.Null(JanelaEditor.MotivoParaExigirSalvarComo(new SessaoDocumento { ArquivoOrigem = @"C:\x\a.docx", Formato = FormatoDocumento.Docx, PastaRecovery = Pasta() }));
}
