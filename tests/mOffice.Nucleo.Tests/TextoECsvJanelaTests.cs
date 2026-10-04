using System.Text;
using mOffice.Conversor;
using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Plano 6, Task 3 — TXT, CSV e PDF-como-documento pela JANELA real (o x2t e o editor de verdade). Em sessao remota (RDP) estes testes
/// sao ignorados (EditorReal) e rodam em console na Task 12. O dialogo de importacao do CSV/TXT e' do EDITOR (HTML): aqui a resposta
/// dele e' entregue pelo mesmo ponto que o dialogo usa (<c>AscDesktopEditor.SetAdvancedOptions</c>, o canal da senha do Plano 3).
/// </summary>
[Collection(JanelaReal.Nome)]
public class TextoECsvJanelaTests
{
    private static string NovaPasta(string prefixo) =>
        Directory.CreateDirectory(Path.Combine(DadosDeTeste.RaizDeTestes, prefixo + Guid.NewGuid().ToString("N"))).FullName;

    private const string UltimaOpcao = "String(window.__moffice && window.__moffice.ultimaOpcaoAvancada)";

    /// <summary>A resposta do dialogo de importacao do editor: o que ele manda em <c>AscDesktopEditor.SetAdvancedOptions</c>.</summary>
    private static string ResponderDialogo(string xml) =>
        "(function () { try { var w = window.frames[0]; w.AscDesktopEditor.SetAdvancedOptions(" + System.Text.Json.JsonSerializer.Serialize(xml) + "); return 'ok'; } catch (e) { return 'erro: ' + e; } })()";

    private static async Task AbrirRespondendoAsync(JanelaEditor j, string arquivo, string xmlResposta)
    {
        await j.AbrirAsync(arquivo);
        Assert.Equal("89", await EditorReal.Esperar(j, UltimaOpcao, "89", 300));
        Assert.Equal(X2tConversor.CodigoPrecisaDeOpcoes, j.OpcaoAvancadaPendente);
        Assert.Equal("false", (await j.ExecutarJsAsync(EditorReal.ProntoNoEditor)).Trim('"'));
        Assert.Equal("ok", (await j.ExecutarJsAsync(ResponderDialogo(xmlResposta))).Trim('"'));
        Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
        Assert.Null(j.UltimoErro);
        Assert.Null(j.OpcaoAvancadaPendente);
    }

    [SkippableFact] // CSV ANSI ';' abre no mPlanilha com acentos e colunas certos depois de responder o dialogo (opcoes injetadas)
    public void Csv_Ansi_AbreComAcentosEColunas()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-csv-abrir-");
        var csv = Path.Combine(pasta, "ptbr-ansi.csv");
        File.WriteAllBytes(csv, Encoding.Latin1.GetBytes("Nome;Cidade\nAção;São Paulo\n"));
        var xlsx = Path.Combine(pasta, "saida.xlsx");
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha");
                j.Show();
                await AbrirRespondendoAsync(j, csv, "<m_nCsvTxtEncoding>44</m_nCsvTxtEncoding><m_nCsvDelimiter>2</m_nCsvDelimiter>");
                Assert.Equal(FormatoDocumento.Csv, j.SessaoAtual!.Formato);
                Assert.Equal(OpcoesTexto.CsvExcelPtBr, j.SessaoAtual.OpcoesTexto);
                // As celulas, lidas do documento (salvar como xlsx pelo editor e abrir o pacote): A2 = "Ação", B2 = "São Paulo".
                await j.SalvarComoAsync(xlsx, FormatoDocumento.Xlsx);
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal("Ação", FabricaDocumentos.TextoXlsxCelula(xlsx, "A2"));
            Assert.Equal("São Paulo", FabricaDocumentos.TextoXlsxCelula(xlsx, "B2"));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // salvar como "CSV (separado por ponto e virgula)" grava 1252 com ';' sem BOM e reabre igual; a faixa avisa
    public void Csv_SalvarComoAnsi_SeparadorPontoEVirgula()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-csv-salvar-");
        var origem = FabricaDocumentos.Xlsx(pasta, "Ação");
        var saida = Path.Combine(pasta, "saida.csv");
        var avisos = new List<string>();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha");
                j.AvisoAoUsuario += a => avisos.Add(a);
                j.EscolherDestinoParaTeste = (filtro, _) =>
                {
                    Assert.Contains("CSV (separado por ponto e vírgula) (*.csv)", filtro);
                    return JanelaEditor.DestinoDoDialogo(saida, JanelaEditor.IndiceDoFiltro(FormatoDocumento.Csv), FormatoDocumento.Xlsx, true);
                };
                j.Show();
                await j.AbrirAsync(origem);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoCell("São Paulo", "B1"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.PodeSalvarNoEditor, "true", 50));
                var r = await j.SalvarComoComDialogoAsync();
                Assert.True(r.Gravou);
                Assert.Equal(FormatoDocumento.Csv, j.SessaoAtual!.Formato);
                Assert.Equal(OpcoesTexto.CsvExcelPtBr, j.SessaoAtual.OpcoesTexto);
                Assert.Contains(JanelaEditor.MensagemPerdeFormatacao(FormatoDocumento.Csv), avisos);
                Assert.Contains(JanelaEditor.MensagemCsvSoPlanilhaAtiva, avisos);
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            var bytes = File.ReadAllBytes(saida);
            Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "CSV ANSI nao leva BOM");
            Assert.Equal("Ação;São Paulo", Encoding.Latin1.GetString(bytes).TrimEnd('\r', '\n'));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    // ---- Fix 1 da revisao: as opcoes da sessao CSV/TXT acompanham o arquivo (Ctrl+S, Salvar como no mesmo caminho, recuperacao) ----

    private static string SemBomECrlf(byte[] b) =>
        new UTF8Encoding(false, true).GetString(b).TrimStart('\uFEFF').Replace("\r\n", "\n");

    [SkippableFact] // abrir CSV UTF-8 com virgula e dar Ctrl+S: continua UTF-8 com ',' (nao vira ANSI ';'); o ANSI ';' tambem continua ANSI ';'
    public void Csv_CtrlS_MantemCodificacaoEDelimitador()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-csv-ctrls-");
        var utf8 = Path.Combine(pasta, "utf8.csv");
        File.WriteAllBytes(utf8, new UTF8Encoding(false).GetBytes("Nome,Cidade\nAção,São Paulo\n"));   // termina em LF: o x2t acumulava uma linha vazia por ciclo (Fix 2)
        var ansi = Path.Combine(pasta, "ansi.csv");
        File.WriteAllBytes(ansi, Encoding.Latin1.GetBytes("Nome;Cidade\nAção;São Paulo\n"));
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha");
                j.Show();
                await AbrirRespondendoAsync(j, utf8, "<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding><m_nCsvDelimiter>4</m_nCsvDelimiter>");
                await j.SalvarAsync();
                await AbrirRespondendoAsync(j, ansi, "<m_nCsvTxtEncoding>44</m_nCsvTxtEncoding><m_nCsvDelimiter>2</m_nCsvDelimiter>");
                await j.SalvarAsync();
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal("Nome,Cidade\nAção,São Paulo\n", SemBomECrlf(File.ReadAllBytes(utf8)));      // UTF-8 estrito (lanca se nao for) com ','
            var b = File.ReadAllBytes(ansi);
            Assert.Equal("Nome;Cidade\nAção;São Paulo\n", Encoding.Latin1.GetString(b).Replace("\r\n", "\n"));   // bytes 1252, sem BOM, com ';'
            Assert.False(b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // Salvar como no MESMO caminho trocando ANSI -> UTF-8: o Ctrl+S seguinte continua UTF-8 (a sessao nao fica com as opcoes velhas)
    public void Csv_SalvarComoMesmoCaminho_TrocaDeLinha_CtrlSMantem()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-csv-mesmo-");
        var csv = Path.Combine(pasta, "dados.csv");
        File.WriteAllBytes(csv, Encoding.Latin1.GetBytes("Nome;Cidade\nAção;São Paulo\n"));
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha");
                j.EscolherDestinoParaTeste = (_, _) =>
                {
                    var i = JanelaEditor.LinhasSalvar(FormatoDocumento.Csv, true).ToList().FindIndex(l => l.Texto == OpcoesTexto.CsvUtf8) + 1;
                    return JanelaEditor.DestinoDoDialogo(csv, i, FormatoDocumento.Csv, true);
                };
                j.Show();
                await AbrirRespondendoAsync(j, csv, "<m_nCsvTxtEncoding>44</m_nCsvTxtEncoding><m_nCsvDelimiter>2</m_nCsvDelimiter>");
                await j.SalvarComoComDialogoAsync();
                Assert.Equal(OpcoesTexto.CsvUtf8, j.SessaoAtual!.OpcoesTexto);   // as opcoes REALMENTE usadas, mesmo sem Repontar
                await j.SalvarAsync();
                j.EscolherDestinoParaTeste = null;
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal("Nome;Cidade\nAção;São Paulo\n", SemBomECrlf(File.ReadAllBytes(csv)));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // recuperacao: o CSV UTF-8 com ',' volta com as opcoes (sessao.json) e o Ctrl+S nao o reencoda; sessao.json ANTIGO (sem opcoes)
                    // exige Salvar como e RECUSA gravar por cima do original (nunca adivinha a codificacao)
    public void Csv_Recuperado_MantemOpcoes_OuExigeSalvarComo()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-csv-recup-");
        var csv = Path.Combine(pasta, "dados.csv");
        var original = new UTF8Encoding(false).GetBytes("Nome,Cidade\nAção,São Paulo");
        File.WriteAllBytes(csv, original);
        var recovery = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha"); j.ManterRecoveryAoFecharParaTeste = true; j.Show();
                await AbrirRespondendoAsync(j, csv, "<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding><m_nCsvDelimiter>4</m_nCsvDelimiter>");
                recovery = j.SessaoAtual!.PastaRecovery;
                Assert.StartsWith("ok", (await j.ExecutarJsAsync(EditorReal.InserirTextoCell("novo", "C1"))).Trim('"'));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                await j.DescarregarChangesAsync();
                var m = mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(recovery)!;
                Assert.Equal((46, 4), (m.TextoCodificacao, m.TextoDelimitador));   // persistido no sessao.json
                j.Close();   // "queda": a pasta fica
            }, TimeSpan.FromSeconds(180));
            Assert.True(Directory.Exists(recovery));
            var copia = Path.Combine(DadosDeTeste.RaizDeTestes, "moffice-csv-recup-copia-" + Guid.NewGuid().ToString("N"));
            CopiarPasta(recovery, copia);   // a mesma queda, reaproveitada para o cenario do sessao.json antigo

            // (1) com as opcoes: volta UTF-8 ',' e o Ctrl+S grava com elas
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha"); j.Show();
                await j.RecuperarAsync(recovery);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal(new OpcoesTexto(46, 4), j.SessaoAtual!.OpcoesTexto);
                Assert.False(j.SessaoAtual.ExigeSalvarComo);
                await j.SalvarAsync();
                j.Close();
            }, TimeSpan.FromSeconds(240));
            var texto = SemBomECrlf(File.ReadAllBytes(csv));
            Assert.Contains("Nome,Cidade,novo", texto);
            Assert.DoesNotContain(";", texto);

            // (2) sessao.json antigo (sem as opcoes): exige Salvar como e nao toca no original
            File.WriteAllBytes(csv, original);
            var m2 = mOffice.Nucleo.Recuperacao.MetadadosSessao.Ler(copia)!;
            (m2 with { TextoCodificacao = null, TextoDelimitador = null, TextoDelimitadorChar = null }).ComCarimboDe(csv).Gravar(copia);
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha"); j.Show();
                await j.RecuperarAsync(copia);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Null(j.SessaoAtual!.OpcoesTexto);
                Assert.True(j.SessaoAtual.ExigeSalvarComo);
                var e = await Assert.ThrowsAsync<InvalidOperationException>(() => j.SalvarAsync());
                Assert.Equal(JanelaEditor.MensagemOpcoesDeTextoDesconhecidas, e.Message);
                // escolhendo a codificacao numa linha do dialogo, grava (e as opcoes passam a ser as escolhidas)
                var destino = Path.Combine(pasta, "novo.csv");
                var r = await j.SalvarComoAsync(destino, FormatoDocumento.Csv, OpcoesTexto.CsvUtf8);
                Assert.True(r.Gravou);
                Assert.False(j.SessaoAtual.ExigeSalvarComo);
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(original, File.ReadAllBytes(csv));   // o original ficou byte a byte
            try { Directory.Delete(copia, true); } catch { }
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    private static void CopiarPasta(string origem, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var d in Directory.GetDirectories(origem, "*", SearchOption.AllDirectories)) Directory.CreateDirectory(Path.Combine(destino, Path.GetRelativePath(origem, d)));
        foreach (var f in Directory.GetFiles(origem, "*", SearchOption.AllDirectories))
        {
            if (Path.GetFileName(f).Equals("rec.lock", StringComparison.OrdinalIgnoreCase)) continue;
            File.Copy(f, Path.Combine(destino, Path.GetRelativePath(origem, f)), true);
        }
    }

    [SkippableFact] // TXT UTF-8 abre no mTexto com o texto integral (depois do dialogo de codificacao)
    public void Txt_Utf8_Abre()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-txt-abrir-");
        var txt = Path.Combine(pasta, "nota.txt");
        File.WriteAllText(txt, "Primeira linha com Ação\nSegunda linha em São Paulo\n", new UTF8Encoding(false));
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.Show();
                await AbrirRespondendoAsync(j, txt, "<m_nCsvTxtEncoding>46</m_nCsvTxtEncoding>");
                Assert.Equal(FormatoDocumento.Txt, j.SessaoAtual!.Formato);
                var texto = await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor);
                Assert.Contains("Primeira linha com Ação", texto);
                Assert.Contains("Segunda linha em São Paulo", texto);
                j.Close();
            }, TimeSpan.FromSeconds(240));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // fechar o dialogo sem confirmar (resposta sem opcoes) fecha a janela e nao cria nada
    public void Csv_DialogoCancelado_FechaAJanela()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-csv-cancelar-");
        var csv = Path.Combine(pasta, "a.csv");
        File.WriteAllText(csv, "a;b\n1;2\n");
        var recovery = "";
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mplanilha");
                var fechou = false;
                j.Closed += (_, _) => fechou = true;
                j.Show();
                await j.AbrirAsync(csv);
                recovery = j.SessaoAtual!.PastaRecovery;
                Assert.True(Directory.Exists(recovery));
                Assert.Equal("89", await EditorReal.Esperar(j, UltimaOpcao, "89", 300));
                Assert.Equal("ok", (await j.ExecutarJsAsync(ResponderDialogo(""))).Trim('"'));
                for (var i = 0; i < 100 && !fechou; i++) await Task.Delay(100);
                Assert.True(fechou);
            }, TimeSpan.FromSeconds(240));
            Assert.False(Directory.Exists(recovery), "a abertura cancelada nao deixa pasta de recuperacao");
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // PDF gerado pelo proprio motor abre no mTexto como documento NOVO com o texto (confirmacao injetada); o PDF nao e' alterado
    public void Pdf_ComoDocumentoEditavel()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-pdf-doc-");
        var docx = FabricaDocumentos.Docx(pasta, "texto-do-pdf-original");
        var pdf = Path.Combine(pasta, "contrato.pdf");
        new X2tConversor(MotorDeTeste.Localizar()).ConverterAsync(docx, FormatoDocumento.Pdf, pdf).GetAwaiter().GetResult();
        var shaAntes = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(pdf)));
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                var perguntas = new List<string>();
                j.ConfirmarAbrirPdfParaTeste = p => { perguntas.Add(p); return true; };
                j.Show();
                await j.AbrirPdfComoDocumentoAsync(pdf);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                Assert.Equal([pdf], perguntas);
                Assert.True(j.SessaoAtual!.Novo);
                Assert.Equal(FormatoDocumento.Docx, j.SessaoAtual.Formato);
                Assert.Equal("contrato", j.SessaoAtual.Titulo);   // o titulo do documento novo e' o nome do PDF, nao "Documento1"
                Assert.Contains("texto-do-pdf-original", await j.ExecutarJsAsync(EditorReal.LerTextoDoEditor));
                j.Close();
            }, TimeSpan.FromSeconds(240));
            Assert.Equal(shaAntes, Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(pdf))));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    [SkippableFact] // PDF recusado na confirmacao: nada e' convertido nem criado; quem chamou (o gerente) fecha a janela
    public void Pdf_ConfirmacaoNegada_NaoAbre()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = NovaPasta("moffice-pdf-negado-");
        var docx = FabricaDocumentos.Docx(pasta, "x");
        var pdf = Path.Combine(pasta, "a.pdf");
        new X2tConversor(MotorDeTeste.Localizar()).ConverterAsync(docx, FormatoDocumento.Pdf, pdf).GetAwaiter().GetResult();
        try
        {
            JanelaSta.Executar(async () =>
            {
                var j = EditorReal.NovaJanela("mtexto");
                j.ConfirmarAbrirPdfParaTeste = _ => false;
                j.Show();
                await Assert.ThrowsAsync<JanelaEditor.ConversaoDePdfRecusadaException>(() => j.AbrirPdfComoDocumentoAsync(pdf));
                Assert.Null(j.SessaoAtual);
                j.Close();
            }, TimeSpan.FromSeconds(120));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }
}
