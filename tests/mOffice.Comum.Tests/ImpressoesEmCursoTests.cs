using mOffice.Comum;
using mOffice.Conversor;
using mOffice.Nucleo;
using mOffice.Nucleo.Impressao;
using mOffice.Nucleo.Tests;

namespace mOffice.Comum.Tests;

/// <summary>Fix round 1 da Task 5 do Plano 3: o encerramento espera as impressoes em curso (e o PDF de impressao e'
/// apagado no fim de cada envio).</summary>
public class ImpressoesEmCursoTests
{
    private sealed class JanelaFake : IJanelaDocumento
    {
        public string? Caminho => null;
        public event Action? Fechado;
        public void Ativar() { }
        public Task CarregarAsync() => Task.CompletedTask;
        public void Fechar() => Fechado?.Invoke();
    }

    private static string PdfFalso()
    {
        var pdf = Path.Combine(Directory.CreateTempSubdirectory("moffice-drenar-").FullName, "impressao.pdf");
        File.WriteAllText(pdf, "%PDF-1.4");
        return pdf;
    }

    private static void Limpar(string pdf) { try { Directory.Delete(Path.GetDirectoryName(pdf)!, true); } catch { } }

    [Fact] // sem impressao em curso: fechar a ultima janela encerra NA HORA (comportamento de antes)
    public async Task SemImpressao_EncerraNaHora()
    {
        var eventos = new List<string>();
        var imp = new ImpressoesEmCurso(() => eventos.Add("encerrar"), _ => { });
        var janela = new JanelaFake();
        var g = new GerenteDeJanelas(_ => janela, () => { }, () => { }, imp.PedirEncerramento);
        await g.NovoAsync();
        janela.Fechar();
        Assert.Equal(["encerrar"], eventos);
        Assert.False(imp.Drenando);
    }

    [Fact] // fechar a ultima janela com impressao em curso: o app ESPERA; a impressao termina, o PDF e' apagado e SO'
           // ENTAO o encerramento acontece (uma vez)
    public async Task ComImpressao_EsperaTerminar_ApagaOPdf_EDepoisEncerra()
    {
        var pdf = PdfFalso();
        try
        {
            var eventos = new List<string>();
            var imp = new ImpressoesEmCurso(() => eventos.Add($"encerrar (pdf existe = {File.Exists(pdf)})"), _ => { }, esperar: _ => new TaskCompletionSource().Task);
            var janela = new JanelaFake();
            var g = new GerenteDeJanelas(_ => janela, () => { }, () => { }, imp.PedirEncerramento);
            await g.NovoAsync();
            var portao = new TaskCompletionSource<ImpressaoEnviada?>();
            var avisos = new List<string>();
            var envio = imp.ImprimirAsync(pdf, () => portao.Task, avisos.Add);
            Assert.Equal(1, imp.EmCurso);

            janela.Fechar();
            Assert.Empty(eventos);
            Assert.True(imp.Drenando);
            Assert.True(File.Exists(pdf));

            portao.SetResult(new ImpressaoEnviada("Sala 2", false));
            await envio.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(["encerrar (pdf existe = False)"], eventos);
            Assert.Equal(["Enviado para Sala 2."], avisos);
            Assert.False(imp.Drenando);
            Assert.Equal(0, imp.EmCurso);
        }
        finally { Limpar(pdf); }
    }

    [Fact] // envio travado: no limite o app encerra assim mesmo; o envio que terminar depois nao encerra de novo
    public async Task EnvioTravado_EncerraNoLimite_UmaVezSo()
    {
        var pdf = PdfFalso();
        try
        {
            var eventos = new List<string>();
            var relogio = new TaskCompletionSource();
            TimeSpan? pedido = null;
            var imp = new ImpressoesEmCurso(() => eventos.Add("encerrar"), _ => { }, TimeSpan.FromMinutes(3), t => { pedido = t; return relogio.Task; });
            var portao = new TaskCompletionSource<ImpressaoEnviada?>();
            var envio = imp.ImprimirAsync(pdf, () => portao.Task, _ => { });
            imp.PedirEncerramento();
            Assert.Equal(TimeSpan.FromMinutes(3), pedido);
            Assert.Empty(eventos);
            relogio.SetResult();
            await Task.Delay(50);
            Assert.Equal(["encerrar"], eventos);
            portao.SetResult(new ImpressaoEnviada("X", false));
            await envio.WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(["encerrar"], eventos);
        }
        finally { Limpar(pdf); }
    }

    [Fact] // o usuario abriu o app de novo durante a espera: cancelar = nem o fim da impressao nem o limite encerram
    public async Task Cancelar_OAppSegue()
    {
        var pdf = PdfFalso();
        try
        {
            var eventos = new List<string>();
            var relogio = new TaskCompletionSource();
            var imp = new ImpressoesEmCurso(() => eventos.Add("encerrar"), _ => { }, esperar: _ => relogio.Task);
            var portao = new TaskCompletionSource<ImpressaoEnviada?>();
            var envio = imp.ImprimirAsync(pdf, () => portao.Task, _ => { });
            imp.PedirEncerramento();
            Assert.True(imp.Cancelar());
            Assert.False(imp.Cancelar());
            portao.SetResult(null);
            await envio.WaitAsync(TimeSpan.FromSeconds(10));
            relogio.SetResult();
            await Task.Delay(50);
            Assert.Empty(eventos);
            Assert.False(File.Exists(pdf)); // o PDF e' apagado mesmo sem encerrar (e com o dialogo cancelado)
        }
        finally { Limpar(pdf); }
    }

    [Fact] // o que chega na faixa: sucesso (com o aviso de offline), erro pt-BR do nucleo, erro inesperado SEM detalhe
           // tecnico, cancelamento = nada; o PDF e' apagado nos quatro casos
    public async Task Avisos_SemDetalheTecnico_EPdfSempreApagado()
    {
        var casos = new (Func<Task<ImpressaoEnviada?>> Envio, string[] Esperado)[]
        {
            (() => Task.FromResult<ImpressaoEnviada?>(new ImpressaoEnviada("Brother", true)),
                ["Enviado para Brother. O Windows indica que ela está desligada ou offline: o trabalho fica na fila até ela voltar."]),
            (() => throw new InvalidOperationException("A impressora \"X\" não foi encontrada no Windows."),
                ["Não foi possível imprimir: A impressora \"X\" não foi encontrada no Windows."]),
            (() => Task.FromException<ImpressaoEnviada?>(new System.Runtime.InteropServices.COMException("HRESULT 0x80004005 Unspecified error")),
                ["Não foi possível imprimir: ocorreu um erro inesperado (os detalhes ficaram no registro do mOffice)."]),
            (() => Task.FromResult<ImpressaoEnviada?>(null), []),
        };
        foreach (var (envio, esperado) in casos)
        {
            var pdf = PdfFalso();
            try
            {
                var avisos = new List<string>();
                var log = new List<string>();
                await new ImpressoesEmCurso(() => { }, log.Add).ImprimirAsync(pdf, envio, avisos.Add);
                Assert.Equal(esperado, avisos);
                Assert.False(File.Exists(pdf));
                Assert.DoesNotContain(avisos, a => a.Contains("HRESULT") || a.Contains("Unspecified"));
            }
            finally { Limpar(pdf); }
        }
    }

    /// <summary>
    /// O caso do enunciado com a impressao DE VERDADE (sem fila: a <see cref="Impressora"/> inteira com o envio
    /// trocado por <c>PrintToPdfAsync</c>): a ultima janela fecha com o envio em curso; o app nao encerra; o envio
    /// termina (a saida tem as 3 paginas), o PDF de impressao e' apagado e so' entao o encerramento e' pedido.
    /// </summary>
    [SkippableFact]
    public void UltimaJanelaFechaNoMeioDoEnvio_ImpressaoTermina_PdfApagado_DepoisEncerra()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-drenar-real-").FullName;
        try
        {
            var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, string.Join(" ", Enumerable.Range(1, 1500).Select(i => $"palavra{i}")));
            var pdf = Path.Combine(pasta, "impressao.pdf");
            new X2tConversor(mOffice.Conversor.Tests.MotorDeTeste.Localizar()).ConverterAsync(docx, FormatoDocumento.Pdf, pdf).GetAwaiter().GetResult();
            var destino = Path.Combine(pasta, "saida.pdf");
            var eventos = new List<string>();
            var avisos = new List<string>();
            JanelaSta.Executar(async () =>
            {
                var imp = new ImpressoesEmCurso(() => eventos.Add($"encerrar (pdf existe = {File.Exists(pdf)}, saida existe = {File.Exists(destino)})"), _ => { });
                var janela = new JanelaFake();
                var g = new GerenteDeJanelas(_ => janela, () => { }, () => { }, imp.PedirEncerramento);
                await g.NovoAsync();
                var noEnvio = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                var portao = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
                var impressora = ImpressaoSondaTests.SemFila(destino, antes: async () => { noEnvio.TrySetResult(); await portao.Task; });
                var envio = imp.ImprimirAsync(pdf, () => impressora.ImprimirAsync(pdf, OpcoesImpressao.Padrao), avisos.Add);
                await noEnvio.Task.WaitAsync(TimeSpan.FromSeconds(60));
                janela.Fechar();
                Assert.Empty(eventos);
                Assert.True(imp.Drenando);
                portao.TrySetResult();
                await envio.WaitAsync(TimeSpan.FromSeconds(60));
            }, TimeSpan.FromSeconds(120));
            Assert.Equal(["encerrar (pdf existe = False, saida existe = True)"], eventos);
            Assert.Equal(["Enviado para Fila de teste."], avisos);
            Assert.Equal(3, mOffice.Conversor.Tests.TextoDePdf.Paginas(destino));
        }
        finally { try { Directory.Delete(pasta, true); } catch { } }
    }

    /// <summary>Adaptador de uma <see cref="JanelaEditor"/> de verdade para a politica de janelas (o mesmo papel do
    /// <c>JanelaAdaptada</c> privado de <see cref="AplicativoOffice"/>).</summary>
    private sealed class JanelaEditorAdaptada(JanelaEditor j) : IJanelaDocumento
    {
        public string? Caminho => null;
        public event Action? Fechado { add => j.Fechado += value; remove => j.Fechado -= value; }
        public void Ativar() { }
        public Task CarregarAsync() => Task.CompletedTask;
        public void Fechar() => j.Close();
    }

    /// <summary>
    /// Fix round 4 da revisao final do Plano 3 (defeito da era do Plano 2): Ctrl+S e fechar a ULTIMA janela na hora. O
    /// app NAO encerra com o salvamento em curso (sem isto, o Shutdown matava o Dispatcher, o x2t deixava
    /// <c>&lt;doc&gt;.moffice-tmp</c> e o arquivo do usuario ficava velho, em silencio); o salvamento termina com a janela
    /// fechada, grava o arquivo, apaga a pasta de recuperacao, nao deixa temporario — e SO' ENTAO o encerramento e' pedido.
    /// A fiacao e' a de <see cref="AplicativoOffice"/> (guardada abaixo pela fonte).
    /// </summary>
    [SkippableFact]
    public void UltimaJanelaFechaNoMeioDoSalvamento_ArquivoGravado_PastaApagada_DepoisEncerra()
    {
        EditorReal.PularSeNaoDaParaAbrirJanela();
        var pasta = Directory.CreateTempSubdirectory("moffice-drenar-salvar-").FullName;
        var docx = mOffice.Conversor.Tests.FabricaDocumentos.Docx(pasta, "Texto base");
        var eventos = new List<string>();
        var recovery = "";
        try
        {
            string Estado()
            {
                bool salvo;
                try { salvo = mOffice.Conversor.Tests.FabricaDocumentos.TextoDocx(docx).Contains("gravado-antes-de-encerrar"); }
                catch (Exception) { salvo = false; }
                return $"encerrar (salvo = {salvo}, pasta = {Directory.Exists(recovery)}, tmp = {File.Exists(docx + ".moffice-tmp")})";
            }
            JanelaSta.Executar(async () =>
            {
                var imp = new ImpressoesEmCurso(() => eventos.Add(Estado()), _ => { });
                var j = EditorReal.NovaJanela("mtexto");
                j.SalvamentoIniciado += fim => _ = imp.AcompanharAsync(fim, "salvamento");   // a fiacao de AplicativoOffice
                j.Show();
                var g = new GerenteDeJanelas(_ => new JanelaEditorAdaptada(j), () => { }, () => { }, imp.PedirEncerramento);
                await g.NovoAsync();   // registra a janela na politica (a ultima janela aberta)
                await j.AbrirAsync(docx);
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.ProntoNoEditor, "true", 300));
                recovery = j.SessaoAtual!.PastaRecovery;
                await j.ExecutarJsAsync(EditorReal.InserirTextoWord(" gravado-antes-de-encerrar"));
                Assert.Equal("true", await EditorReal.Esperar(j, EditorReal.EspelhoModificado, "true", 100));
                Assert.Equal("ok", (await j.ExecutarJsAsync(EditorReal.NoEditor("api.asc_Save(false, false, true); return \"ok\";"))).Trim('"'));
                j.Close();
                Assert.Empty(eventos);           // nao encerrou com o salvamento em curso
                Assert.True(imp.Drenando);
                for (var i = 0; i < 600 && eventos.Count == 0; i++) await Task.Delay(100);
            }, TimeSpan.FromSeconds(180));
            Assert.Equal(["encerrar (salvo = True, pasta = False, tmp = False)"], eventos);
            var fonte = File.ReadAllText(Path.Combine(AplicativoOfficeTests.RaizRepo(), "src", "mOffice.Comum", "AplicativoOffice.cs"));
            Assert.Contains("j.SalvamentoIniciado += fim => _ = _impressoes.AcompanharAsync(fim, \"salvamento\");", fonte);
        }
        finally { try { Directory.Delete(pasta, true); } catch { } try { if (recovery.Length > 0) Directory.Delete(recovery, true); } catch { } }
    }
}
