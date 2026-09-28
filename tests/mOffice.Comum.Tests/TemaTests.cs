using System.Windows;
using System.Windows.Media;
using mOffice.Comum;
using mOffice.Nucleo;
using mOffice.Nucleo.Tests;

namespace mOffice.Comum.Tests;

[Collection("janela real")] // a definicao da colecao vive em mOffice.Nucleo.Tests; aqui o nome so' serializa os testes desta assembly (e xunit.runner.json ja' desliga o paralelismo)
public class TemaTests
{
    [Fact] // os dois dicionarios de tokens declaram EXATAMENTE o mesmo conjunto de chaves (troca ao vivo sem buraco)
    public void Tokens_ClaroEEscuro_MesmasChaves()
    {
        JanelaSta.Executar(() =>
        {
            var claro = JanelaWpf.Tokens(false).Keys.Cast<object>().Select(k => k.ToString()!).OrderBy(k => k).ToList();
            var escuro = JanelaWpf.Tokens(true).Keys.Cast<object>().Select(k => k.ToString()!).OrderBy(k => k).ToList();
            Assert.Equal(claro, escuro);
            Assert.Contains("Cor.Primaria", claro); Assert.Contains("Cor.FundoBoasVindasGrad", claro);
            return Task.CompletedTask;
        });
    }

    [Fact] // nenhuma chave herdada do mPDF ficou com o prefixo antigo; os tres logos existem
    public void Tema_NenhumaChaveMPdf_ETresLogos()
    {
        JanelaSta.Executar(() =>
        {
            var tema = JanelaWpf.Tema();
            var chaves = tema.Keys.Cast<object>().Select(k => k.ToString()!).ToList();
            Assert.DoesNotContain(chaves, c => c.StartsWith("mPdf.", StringComparison.Ordinal) || c == "Logo.mPDF");
            Assert.Contains("mOffice.Button.Primario", chaves);
            foreach (var logo in new[] { "Logo.mTexto", "Logo.mPlanilha", "Logo.mSlides" })
                Assert.IsType<DrawingImage>(tema[logo]);
            return Task.CompletedTask;
        });
    }

    [Fact] // TemaService troca o dicionario de tokens na posicao 0 e e' idempotente
    public void TemaService_Aplicar_TrocaTokens()
    {
        JanelaSta.Executar(() =>
        {
            var alvo = new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() };
            var s = new TemaService(alvo);
            s.Aplicar(ModoTema.Escuro);
            Assert.Equal(2, alvo.Count);
            Assert.Contains("Tokens.Escuro", alvo[0].Source.OriginalString);
            Assert.True(s.EscuroEfetivo); Assert.Equal(TemaEditor.Escuro, s.TemaDoEditor);
            s.Aplicar(ModoTema.Claro); s.Aplicar(ModoTema.Claro);
            Assert.Equal(2, alvo.Count);
            Assert.Contains("Tokens.Claro", alvo[0].Source.OriginalString);
            Assert.False(s.EscuroEfetivo); Assert.Equal(TemaEditor.Claro, s.TemaDoEditor);
            return Task.CompletedTask;
        });
    }

    [Fact] // Sistema: segue a leitura do registro (injetada) e dispara EfetivoAlterado quando o sistema muda
    public void TemaService_Sistema_SegueOWindows()
    {
        JanelaSta.Executar(() =>
        {
            var claro = true;
            var alvo = new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() };
            var s = new TemaService(alvo, lerAppsClaros: () => claro);
            var recebidos = new List<bool>();
            s.EfetivoAlterado += v => recebidos.Add(v);
            s.Aplicar(ModoTema.Sistema);
            Assert.False(s.EscuroEfetivo);
            claro = false; s.SistemaMudou(); // seam: o que SystemEvents.UserPreferenceChanged chamaria
            Assert.True(s.EscuroEfetivo);
            Assert.Contains("Tokens.Escuro", alvo[0].Source.OriginalString);
            Assert.Equal(new[] { false, true }, recebidos);
            s.Aplicar(ModoTema.Claro);
            claro = true; s.SistemaMudou(); // fora do modo Sistema, mudanca do Windows nao faz nada
            Assert.Equal(new[] { false, true, false }, recebidos);
            return Task.CompletedTask;
        });
    }

    [Fact] // SistemaMudou vindo de outra thread (como o SystemEvents faz) e' reencaminhado ao Dispatcher da UI:
           // o dicionario troca e EfetivoAlterado dispara NA thread da UI, nunca na thread de worker
    public void TemaService_SistemaMudouDeOutraThread_ReaplicaNoDispatcher()
    {
        JanelaSta.Executar(async () =>
        {
            var idUi = Environment.CurrentManagedThreadId;
            var claro = true;
            var alvo = new System.Collections.ObjectModel.Collection<ResourceDictionary> { JanelaWpf.Tema() };
            var s = new TemaService(alvo, lerAppsClaros: () => claro);
            var threadsDoEvento = new List<int>();
            s.EfetivoAlterado += _ => threadsDoEvento.Add(Environment.CurrentManagedThreadId);
            s.Aplicar(ModoTema.Sistema);
            Assert.False(s.EscuroEfetivo);
            claro = false;
            await Task.Run(() => s.SistemaMudou()); // thread de worker, como o SystemEvents
            // cede ao Dispatcher para o BeginInvoke rodar
            await System.Windows.Threading.Dispatcher.CurrentDispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Background);
            Assert.True(s.EscuroEfetivo);
            Assert.Contains("Tokens.Escuro", alvo[0].Source.OriginalString);
            Assert.Equal(2, threadsDoEvento.Count);
            Assert.All(threadsDoEvento, id => Assert.Equal(idUi, id));
        });
    }

    [Fact] // leitura real do registro nunca lanca e devolve bool
    public void TemaDoSistema_AppsClaros_NaoLanca() => _ = TemaDoSistema.AppsClaros();
}
