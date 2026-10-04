using System.Text;
using mOffice.Conversor;
using mOffice.Conversor.Tests;
using Xunit.Abstractions;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 4, Task 2 — "Salvar como" PDF/A e "PDF assinado com mPDF" numa janela real (decisao A da S6: o PDF/A sai do
/// proprio x2t, pedido como <c>m_nFormatTo = 513</c> + <c>m_bIsPDFA = true</c>; 521 a partir do Editor.bin sai com codigo
/// 88, MEDIDO). Aqui a prova e' o que o arquivo DECLARA: o <c>pdfaid:part</c> do XMP (PDF/A-2A: part 2, conformance A),
/// que o x2t grava sem compressao (busca de texto nos bytes). Se o arquivo e' CONFORME ao perfil declarado e' a Task 1/S6
/// que valida, no veraPDF. O PDF comum exportado pela mesma janela e' o controle negativo: nao declara PDF/A.
/// </summary>
[Collection(JanelaReal.Nome)]
public class SalvarPdfATests
{
    private readonly ITestOutputHelper _saida;
    public SalvarPdfATests(ITestOutputHelper saida) { _saida = saida; }

    private static string Texto(string pdf) => Encoding.Latin1.GetString(File.ReadAllBytes(pdf));

    private static string NovaPasta(string prefixo) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, prefixo + Guid.NewGuid().ToString("N"))).FullName;

    /// <summary>O painel "Salvar como" do editor, igual ao web-apps 9.4 (documenteditor/main/app.js):
    /// <c>(a = new Asc.asc_CDownloadOptions(tipo, true)).asc_setIsSaveAs(true); api.asc_DownloadAs(a)</c> — no desktop
    /// vira <c>LocalFileSave("saveas=true;", …, fileType = tipo, …)</c>.</summary>
    private static string PainelSalvarComo(string tipo) => EditorReal.NoEditor($$"""
            var a = new w.Asc.asc_CDownloadOptions(w.Asc.c_oAscFileType.{{tipo}}, true);
            a.asc_setIsSaveAs(true);
            api.asc_DownloadAs(a);
            return "ok " + w.Asc.c_oAscFileType.{{tipo}};
        """);

    [SkippableFact] // SalvarComoAsync(PdfA): PDF com XMP pdfaid:part, EXPORTACAO (sessao no original, Exportado, nunca Salvo);
                    // o PDF comum nao declara PDF/A; o "Salvar como" do host e o painel do editor (fileType 521/513) chegam
                    // na linha certa, e so' a linha "PDF assinado" dispara AssinarPedido — depois de exportar
    public void PdfA_Exporta_ComXmp_SessaoNoOriginal_EAssinarSoNaLinhaAssinada()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-pdfa-");
        var docx = FabricaDocumentos.Docx(pasta, "Peticao inicial");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var exportados = new List<string>();
                var salvos = new List<string>();
                var assinar = new List<(JanelaEditor janela, string pdf)>();
                j.Exportado += c => exportados.Add(c);
                j.Salvo += c => salvos.Add(c);
                j.AssinarPedido += (janela, pdf) => assinar.Add((janela, pdf));
                j.Show();
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" editado-pdfa"))).Trim('"'));
                var origem = j.SessaoAtual!.ArquivoOrigem;

                // 1) API do host: PDF/A.
                var pdfa = Path.Combine(pasta, "peticao-a.pdf");
                var r = await j.SalvarComoAsync(pdfa, FormatoDocumento.PdfA);
                Assert.Null(j.UltimoErro);
                Assert.True(r.Exportou);
                var t = Texto(pdfa);
                Assert.StartsWith("%PDF-", t);
                var parte = System.Text.RegularExpressions.Regex.Match(t, @"pdfaid:part(?:>|="")\s*(\d)");
                _saida.WriteLine($"PDF/A: {t.Length} bytes, pdfaid:part = {(parte.Success ? parte.Groups[1].Value : "(ausente)")}, "
                    + "conformance = " + System.Text.RegularExpressions.Regex.Match(t, @"pdfaid:conformance(?:>|="")\s*([A-Za-z])").Groups[1].Value);
                Assert.True(parte.Success, "o PDF/A do x2t nao declara pdfaid:part no XMP");
                Assert.Equal(origem, j.SessaoAtual!.ArquivoOrigem);    // exportar nao reponta a sessao
                Assert.Equal(FormatoDocumento.Docx, j.SessaoAtual!.Formato);
                Assert.Equal([pdfa], exportados);
                Assert.Empty(salvos);
                Assert.Empty(assinar);

                // 2) Controle negativo: o PDF comum da MESMA sessao nao declara PDF/A (a busca acima discrimina).
                var pdf = Path.Combine(pasta, "peticao.pdf");
                Assert.True((await j.SalvarComoAsync(pdf, FormatoDocumento.Pdf)).Exportou);
                Assert.DoesNotContain("pdfaid:part", Texto(pdf));

                // 3) "Salvar como" do HOST com a linha PDF/A: o seam recebe o filtro de 5 linhas; sem Assinar nada e' pedido.
                exportados.Clear();
                var filtros = new List<(string filtro, FormatoDocumento sugerido)>();
                var pdfaDialogo = Path.Combine(pasta, "dialogo-a.pdf");
                j.EscolherDestinoParaTeste = (filtro, sugerido) => { filtros.Add((filtro, sugerido)); return new DestinoSalvar(pdfaDialogo, FormatoDocumento.PdfA); };
                Assert.True((await j.SalvarComoComDialogoAsync()).Exportou);
                Assert.Contains("pdfaid:part", Texto(pdfaDialogo));
                Assert.Equal([pdfaDialogo], exportados);
                Assert.Empty(assinar);
                Assert.Equal((JanelaEditor.FiltroSalvar(FormatoDocumento.Docx), FormatoDocumento.Docx), filtros.Single());

                // 4) "Salvar como" do HOST com a linha "PDF assinado com mPDF": exporta PDF e, DEPOIS, pede a assinatura.
                var pdfAssinar = Path.Combine(pasta, "dialogo-assinar.pdf");
                j.EscolherDestinoParaTeste = (_, _) =>
                {
                    Assert.Empty(assinar);   // nada pedido antes de o usuario escolher
                    return new DestinoSalvar(pdfAssinar, FormatoDocumento.Pdf, Assinar: true);
                };
                Assert.True((await j.SalvarComoComDialogoAsync()).Exportou);
                Assert.Single(assinar);
                Assert.Same(j, assinar[0].janela);
                Assert.Equal(pdfAssinar, assinar[0].pdf);
                Assert.True(File.Exists(pdfAssinar));
                Assert.DoesNotContain("pdfaid:part", Texto(pdfAssinar));   // "PDF assinado" e' PDF comum (o mPDF assina)

                // 5) Painel "Salvar como" do EDITOR: fileType 521 (c_oAscFileType.PDFA) pre-seleciona PDF/A com o filtro
                //    inteiro; 513 pre-seleciona PDF. O caminho do editor tambem respeita a linha "assinado".
                filtros.Clear();
                assinar.Clear();
                exportados.Clear();
                var doPainelA = Path.Combine(pasta, "painel-a.pdf");
                var doPainelAssinar = Path.Combine(pasta, "painel-assinar.pdf");
                j.EscolherDestinoParaTeste = (filtro, sugerido) =>
                {
                    filtros.Add((filtro, sugerido));
                    return sugerido == FormatoDocumento.PdfA
                        ? new DestinoSalvar(doPainelA, FormatoDocumento.PdfA)
                        : new DestinoSalvar(doPainelAssinar, FormatoDocumento.Pdf, Assinar: true);
                };
                Assert.Equal("ok 521", (await j.ExecutarJsAsync(PainelSalvarComo("PDFA"))).Trim('"'));
                for (var i = 0; i < 300 && exportados.Count < 1; i++) await Task.Delay(100);
                Assert.Equal("ok 513", (await j.ExecutarJsAsync(PainelSalvarComo("PDF"))).Trim('"'));
                // O Exportado sai de dentro da conversao e o AssinarPedido logo DEPOIS (JanelaEditor: PedirAssinatura apos o
                // EndSave): esperar so' o 2o exportado e ler `assinar` em seguida era corrida (falhou 1 vez em 4 no console,
                // Task 12 do Plano 6). Espera os dois sinais.
                for (var i = 0; i < 300 && (exportados.Count < 2 || assinar.Count < 1); i++) await Task.Delay(100);
                Assert.Null(j.UltimoErro);
                Assert.Equal([doPainelA, doPainelAssinar], exportados);
                Assert.Equal([(JanelaEditor.FiltroSalvar(FormatoDocumento.Docx), FormatoDocumento.PdfA),
                              (JanelaEditor.FiltroSalvar(FormatoDocumento.Docx), FormatoDocumento.Pdf)], filtros);
                Assert.Contains("pdfaid:part", Texto(doPainelA));
                Assert.Equal([doPainelAssinar], assinar.Select(a => a.pdf));
                Assert.Equal(origem, j.SessaoAtual!.ArquivoOrigem);
                Assert.Empty(salvos);
                Assert.True(j.Modificado, "exportar nao grava o documento editavel: a alteracao continua pendente");
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // documento com senha (Plano 3): o PDF/A sai SEM /Encrypt, e continua declarando PDF/A
    public void PdfA_DeDocumentoComSenha_SaiSemSenha()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var senha = $"certa-{Guid.NewGuid():N}";
        var pasta = NovaPasta("moffice-pdfa-senha-");
        var cifrado = FabricaDocumentos.DocxComSenha(pasta, "conteudo-sigiloso", senha).GetAwaiter().GetResult();
        var pdfa = Path.Combine(pasta, "protegido-a.pdf");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.Show();
                await j.AbrirAsync(cifrado);
                Assert.Equal("90", await EditorReal.Esperar(j, "String(window.__moffice && window.__moffice.ultimaOpcaoAvancada)", "90", 300));
                Assert.Equal("ok", await TelaCheiaSenhaJanelaTests.EsperarDialogoEDigitar(j, senha));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal(senha, j.SessaoAtual!.Senha);   // linha de base: a sessao TEM senha

                Assert.True((await j.SalvarComoAsync(pdfa, FormatoDocumento.PdfA)).Exportou);
                Assert.Null(j.UltimoErro);
                Assert.Equal(cifrado, j.SessaoAtual!.ArquivoOrigem);
                j.Close();
            }, TimeSpan.FromSeconds(180));
            var t = Texto(pdfa);
            Assert.StartsWith("%PDF-", t);
            Assert.DoesNotContain("/Encrypt", t);
            Assert.Contains("pdfaid:part", t);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}

/// <summary>
/// Plano 4, Task 1 — sondagem S6: o PDF/A do x2t e' CONFORME no veraPDF? Opt-in (<c>MOFFICE_VERAPDF=1</c>, exige o veraPDF
/// instalado; ver <c>tools/pdf/validar-pdfa.ps1</c>). O que conta e' o perfil DECLARADO pelo arquivo (veraPDF em modo
/// automatico: o x2t declara PDF/A-2A); o resultado no 2b vai para a saida do teste so' como registro. Dois caminhos:
/// arquivo a arquivo (<see cref="X2tConversor.ConverterAsync"/>) e o do editor (<c>SalvarComoAsync</c>: Editor.bin +
/// changes pelo <c>DeBinAsync</c>, que e' o que o usuario usa). <c>MOFFICE_VERAPDF_GUARDAR=&lt;pasta&gt;</c> (opcional)
/// copia os PDFs gerados para inspecao manual (fora do repositorio).
/// </summary>
[Collection(JanelaReal.Nome)]
public class SondaPdfATests
{
    private readonly ITestOutputHelper _saida;
    public SondaPdfATests(ITestOutputHelper saida) { _saida = saida; }

    private static bool Ligada => Environment.GetEnvironmentVariable("MOFFICE_VERAPDF") == "1";

    private static string Origem(string tipo, string pasta) => tipo switch
    {
        "docx" => FabricaDocumentos.DocxComImagem(pasta, "Texto com acentuação: ção, é, ü"),
        "xlsx" => FabricaDocumentos.Xlsx(pasta, "Célula"),
        _ => FabricaDocumentos.Pptx(pasta, "Slide"),
    };

    /// <summary>Valida no perfil declarado (auto) e registra o 2b; devolve o codigo do auto.</summary>
    private int ValidarERegistrar(string rotulo, string pdf)
    {
        var guardar = Environment.GetEnvironmentVariable("MOFFICE_VERAPDF_GUARDAR");
        if (!string.IsNullOrEmpty(guardar)) { Directory.CreateDirectory(guardar); File.Copy(pdf, Path.Combine(guardar, rotulo + ".pdf"), true); }
        var (codigo, json) = ValidadorPdfA.Rodar(pdf);
        Assert.NotEqual(2, codigo); // 2 = veraPDF ausente: a sonda ligada exige a ferramenta
        Assert.NotEqual(3, codigo); // 3 = erro ao rodar/ler o veraPDF
        var (codigo2b, json2b) = ValidadorPdfA.Rodar(pdf, "2b");
        _saida.WriteLine($"S6 {rotulo} auto: {json}");
        _saida.WriteLine($"S6 {rotulo} 2b:   {json2b} (codigo {codigo2b})");
        return codigo;
    }

    [SkippableTheory, Trait("Categoria", "Sonda")] // S6 arquivo a arquivo: x2t (513 + m_bIsPDFA) -> veraPDF, um formato por editor
    [InlineData("docx")] [InlineData("xlsx")] [InlineData("pptx")]
    public async Task S6_PdfADoX2t_NoVeraPdf(string tipo)
    {
        Skip.IfNot(Ligada, "sondagem de PDF/A: defina MOFFICE_VERAPDF=1 (exige o veraPDF instalado)");
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-s6-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            var origem = Origem(tipo, pasta);
            var pdfa = Path.Combine(pasta, "saida.pdf");
            var conv = new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar());
            await conv.ConverterAsync(origem, FormatoDocumento.PdfA, pdfa);
            Assert.Equal(0, ValidarERegistrar("arquivo-" + tipo, pdfa));

            // Controle negativo: o PDF COMUM do mesmo documento reprova (o script nao aprova tudo).
            var pdf = Path.Combine(pasta, "comum.pdf");
            await conv.ConverterAsync(origem, FormatoDocumento.Pdf, pdf);
            var guardar = Environment.GetEnvironmentVariable("MOFFICE_VERAPDF_GUARDAR");
            if (!string.IsNullOrEmpty(guardar)) File.Copy(pdf, Path.Combine(guardar, "comum-" + tipo + ".pdf"), true);
            var (codigoComum, jsonComum) = ValidadorPdfA.Rodar(pdf);
            _saida.WriteLine($"S6 comum-{tipo} auto: {jsonComum}");
            Assert.Equal(1, codigoComum);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableTheory, Trait("Categoria", "Sonda")] // S6 pelo editor: janela real, SalvarComoAsync(PdfA) = DeBinAsync (bin + changes)
    [InlineData("docx")] [InlineData("xlsx")] [InlineData("pptx")]
    public void S6_PdfAPeloEditor_NoVeraPdf(string tipo)
    {
        Skip.IfNot(Ligada, "sondagem de PDF/A: defina MOFFICE_VERAPDF=1 (exige o veraPDF instalado)");
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-s6-editor-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            var origem = Origem(tipo, pasta);
            var pdfa = Path.Combine(pasta, "editor-a.pdf");
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela(tipo == "docx" ? "mtexto" : tipo == "pptx" ? "mslides" : JanelaEditor.AppPadrao);
                j.Show();
                await j.AbrirAsync(origem);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                if (tipo == "docx") // com alteracao: o PDF sai de Editor.bin + changes
                    Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" editado-s6"))).Trim('"'));
                Assert.True((await j.SalvarComoAsync(pdfa, FormatoDocumento.PdfA)).Exportou);
                Assert.Null(j.UltimoErro);
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(0, ValidarERegistrar("editor-" + tipo, pdfa));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}

/// <summary>Helper de teste: roda <c>tools/pdf/validar-pdfa.ps1</c> e le' a linha JSON (a ultima nao vazia do stdout).
/// Codigos do script: 0 conforme, 1 nao conforme, 2 veraPDF ausente, 3 erro ao rodar/ler o veraPDF.</summary>
internal static class ValidadorPdfA
{
    public static (int codigo, string json) Rodar(string pdf, string flavour = "auto")
    {
        var script = Path.Combine(FronteiraTests.RaizRepo(), "tools", "pdf", "validar-pdfa.ps1");
        var psi = new System.Diagnostics.ProcessStartInfo("powershell.exe")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        foreach (var a in new[] { "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-Pdf", pdf, "-Flavour", flavour })
            psi.ArgumentList.Add(a);
        using var p = System.Diagnostics.Process.Start(psi)!;
        var saida = p.StandardOutput.ReadToEndAsync();
        _ = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(TimeSpan.FromMinutes(3)))
        {
            // Mata ANTES de falhar: um powershell/java pendurado sobreviveria a' suite inteira.
            try { p.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("validar-pdfa.ps1 nao terminou em 3 minutos");
        }
        var linhas = saida.GetAwaiter().GetResult().Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return (p.ExitCode, linhas.LastOrDefault() ?? "");
    }
}
