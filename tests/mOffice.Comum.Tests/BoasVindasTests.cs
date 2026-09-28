using mOffice.Comum;
using mOffice.Comum.Telas;
using mOffice.Conversor;
using mOffice.Nucleo.Recuperacao;

namespace mOffice.Comum.Tests;

public class BoasVindasTests
{
    private static Recentes Rec() => new(Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"), "r.json"));

    [Fact] // perfis: nome tecnico, nome de exibicao, formato do novo e filtro coerentes
    public void Perfis_Coerentes()
    {
        Assert.Equal("mtexto", PerfilApp.Texto.Nome); Assert.Equal("mTexto", PerfilApp.Texto.NomeExibicao);
        Assert.Equal(FormatoDocumento.Docx, PerfilApp.Texto.FormatoNovo); Assert.Contains("*.odt", PerfilApp.Texto.FiltroAbrir);
        Assert.Equal(FormatoDocumento.Xlsx, PerfilApp.Planilha.FormatoNovo); Assert.Equal(FormatoDocumento.Pptx, PerfilApp.Slides.FormatoNovo);
        Assert.Equal(["Logo.mTexto", "Logo.mPlanilha", "Logo.mSlides"], new[] { PerfilApp.Texto, PerfilApp.Planilha, PerfilApp.Slides }.Select(p => p.ChaveLogo));
    }

    [Fact] // Abrir: escolhe o arquivo e chama abrir; cancelar no dialogo nao chama abrir
    public async Task Abrir_ChamaEscolherDepoisAbrir()
    {
        var abertos = new List<string>(); string? escolha = @"C:\x\a.docx";
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, c => { abertos.Add(c); return Task.CompletedTask; }, () => Task.FromResult<string?>(escolha));
        await vm.AbrirCommand.ExecuteAsync(null);
        Assert.Equal([@"C:\x\a.docx"], abertos);
        escolha = null; await vm.AbrirCommand.ExecuteAsync(null);
        Assert.Single(abertos);
    }

    [Fact] // recente inexistente: sai da lista, mensagem de erro, abrir NAO e' chamado
    public async Task AbrirRecente_Inexistente_RemoveEAvisa()
    {
        var rec = Rec(); var caminho = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N") + ".docx");
        rec.Registrar(caminho);
        var abertos = 0;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => { abertos++; return Task.CompletedTask; }, () => Task.FromResult<string?>(null));
        Assert.Single(vm.Recentes);
        await vm.AbrirRecenteCommand.ExecuteAsync(vm.Recentes[0]);
        Assert.Equal(0, abertos); Assert.Empty(vm.Recentes); Assert.Contains("não existe mais", vm.MensagemErro);
    }

    [Fact] // fixar/remover refletem no arquivo e na colecao observavel
    public void FixarERemover_AtualizamListaEDisco()
    {
        var rec = Rec(); rec.Registrar(@"C:\a.docx"); rec.Registrar(@"C:\b.docx");
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, rec, () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null));
        vm.FixarCommand.Execute(vm.Recentes.Single(i => i.Caminho == @"C:\a.docx"));
        Assert.True(vm.Recentes[0].Fixado); Assert.Equal(@"C:\a.docx", vm.Recentes[0].Caminho);
        vm.RemoverCommand.Execute(vm.Recentes[1]);
        Assert.Single(vm.Recentes); Assert.Single(rec.Listar());
    }

    [Fact] // abrir que falha vira MENSAGEM na tela, nunca excecao solta (o Drop do code-behind e' async void)
    public async Task AbrirCaminho_QuandoAbrirLanca_ViraMensagemENaoPropaga()
    {
        var falhar = true;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask,
            _ => falhar ? throw new InvalidOperationException("formato corrompido") : Task.CompletedTask,
            () => Task.FromResult<string?>(null));

        await vm.AbrirCaminhoAsync(@"C:\x\a.docx"); // nao lanca: se lancasse, este await propagaria
        Assert.Contains("Não foi possível abrir o arquivo", vm.MensagemErro);
        Assert.Contains("formato corrompido", vm.MensagemErro);

        falhar = false;
        await vm.AbrirCaminhoAsync(@"C:\x\a.docx"); // a mensagem velha nao pode ficar na tela
        Assert.Null(vm.MensagemErro);
    }

    [Fact] // solto na tela: so' extensoes do perfil contam; a primeira valida abre (funcao pura no perfil, sem WPF)
    public void ArquivoSolto_FiltraPorPerfil()
    {
        Assert.Equal(@"C:\b.odt", PerfilApp.Texto.PrimeiroAceito([@"C:\a.xlsx", @"C:\b.odt", @"C:\c.docx"]));
        Assert.Null(PerfilApp.Texto.PrimeiroAceito([@"C:\a.xlsx"]));
    }

    // ---------------------------------------------------------------------------------------------------
    // Documentos recuperados (Task 3 do Plano 3). A pergunta de descarte e' SEMPRE um falso aqui: a padrao e'
    // um MessageBox do sistema, que travaria o teste.
    // ---------------------------------------------------------------------------------------------------

    [Fact] // recuperados aparecem; Recuperar chama o delegado; Descartar pergunta, chama o delegado e tira da lista
    public async Task Recuperados_RecuperarEDescartar()
    {
        var d1 = new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false, DateTime.UtcNow);
        var d2 = new DocumentoRecuperavel(@"C:\r\2", "Documento1", "", true, DateTime.UtcNow);
        var lista = new List<DocumentoRecuperavel> { d1, d2 };
        DocumentoRecuperavel? recuperado = null, perguntado = null;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask,
            () => Task.FromResult<string?>(null), null, null,
            () => lista, d => { recuperado = d; return Task.CompletedTask; }, d => lista.Remove(d),
            d => { perguntado = d; return true; });
        Assert.True(vm.TemRecuperados); Assert.Equal(2, vm.Recuperados.Count);
        await vm.RecuperarCommand.ExecuteAsync(d1); Assert.Same(d1, recuperado);
        vm.DescartarCommand.Execute(d2); Assert.Same(d2, perguntado); Assert.Equal([d1], vm.Recuperados);
        Assert.Null(vm.MensagemErro);
    }

    [Fact] // "Nao" na pergunta: nada e' apagado; a lista (e o aviso de TemRecuperados) acompanham a fonte
    public void Recuperados_DescartarSemConfirmar_NaoApaga()
    {
        var d1 = new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false, DateTime.UtcNow);
        var lista = new List<DocumentoRecuperavel> { d1 };
        var descartes = 0; var confirmar = false;
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask,
            () => Task.FromResult<string?>(null), recuperaveis: () => lista, recuperar: _ => Task.CompletedTask,
            descartar: d => { descartes++; lista.Remove(d); }, confirmarDescarte: _ => confirmar);
        var avisos = new List<string?>(); vm.PropertyChanged += (_, e) => avisos.Add(e.PropertyName);
        vm.DescartarCommand.Execute(d1);
        Assert.Equal(0, descartes); Assert.Equal([d1], vm.Recuperados);
        confirmar = true; vm.DescartarCommand.Execute(d1);
        Assert.Equal(1, descartes); Assert.Empty(vm.Recuperados); Assert.False(vm.TemRecuperados);
        Assert.Contains(nameof(TelaBoasVindasViewModel.TemRecuperados), avisos);
    }

    [Fact] // descartar que nao apagou (pasta presa) avisa em vez de sumir calado; recuperar que falhou vira mensagem
    public async Task Recuperados_Falhas_ViramMensagem()
    {
        var d1 = new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", @"C:\x\peticao.docx", false, DateTime.UtcNow);
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask,
            () => Task.FromResult<string?>(null), recuperaveis: () => [d1],
            recuperar: _ => throw new InvalidOperationException("Este documento já está aberto em outra janela."),
            descartar: _ => { }, confirmarDescarte: _ => true);
        vm.DescartarCommand.Execute(d1);
        Assert.Contains("Não foi possível descartar", vm.MensagemErro); Assert.Single(vm.Recuperados);
        await vm.RecuperarCommand.ExecuteAsync(d1);
        Assert.Contains("já está aberto em outra janela", vm.MensagemErro); Assert.Single(vm.Recuperados);
    }

    [Fact] // sem fonte de recuperados (null): secao escondida; o texto da pergunta de descarte e' o da spec
    public void Recuperados_SemFonte_EscondidoETextoDaPergunta()
    {
        var vm = new TelaBoasVindasViewModel(PerfilApp.Texto, Rec(), () => Task.CompletedTask, _ => Task.CompletedTask, () => Task.FromResult<string?>(null));
        Assert.False(vm.TemRecuperados); Assert.Empty(vm.Recuperados);
        Assert.Equal("Descartar as alterações não salvas de peticao.docx? Esta ação não pode ser desfeita.",
            TelaBoasVindasViewModel.PerguntaDescartar(new DocumentoRecuperavel(@"C:\r\1", "peticao.docx", "", false, DateTime.UtcNow)));
    }
}
