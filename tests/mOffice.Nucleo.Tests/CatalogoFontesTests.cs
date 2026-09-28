using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Task 2 (Marco 1 / Plano 1) — o catalogo de fontes do motor passa a ser GERADO pelo mOffice, com o
/// <c>allfontsgen.exe</c> que ele embarca, em vez de emprestado do ONLYOFFICE instalado na maquina.
///
/// O que estes testes prendem, e por que cada um existe:
/// 1. geracao numa pasta vazia produz os DOIS arquivos que o motor exige (AllFonts.js + font_selection.bin);
/// 2. os thumbnails (<c>fonts_thumbnail*.png</c>) nascem junto — o editor os pede por <c>ascdesktop://</c>,
///    e sem eles a lista de fontes da interface fica sem as miniaturas;
/// 3. a impressao digital evita regenerar a cada abertura (7,5 s medidos nesta maquina nao podem virar
///    pedagio de toda inicializacao);
/// 4. uma fonte NOVA na pasta forca regenerar — o controle positivo do item 3: sem ele, uma implementacao
///    que nunca regenera passaria igual;
/// 5. gerador ausente devolve mensagem em pt-BR, nao excecao — e' o caso do pacote quebrado, e o unico
///    sintoma alternativo seria o editor parado em "Carregando documento" para sempre.
///
/// Os testes que tocam o gerador sao <see cref="SkippableFactAttribute"/>: <c>third_party/</c> e' ignorado
/// pelo git, entao numa arvore recem-clonada o exe nao existe ate rodar <c>tools/onlyoffice/obter-motor.ps1</c>
/// e <c>tools/onlyoffice/obter-allfontsgen.ps1</c>. Os dois ultimos (erro legivel e impressao digital) rodam
/// SEMPRE — nao dependem do gerador.
/// </summary>
public class CatalogoFontesTests : IDisposable
{
    private readonly string _raiz = Directory.CreateTempSubdirectory("moffice-catalogo-").FullName;

    private static string Allfontsgen => Path.Combine(mOffice.Conversor.Tests.MotorDeTeste.Localizar().PastaMotor, "converter", "allfontsgen.exe");
    private static string FontesDoMotor => Path.Combine(mOffice.Conversor.Tests.MotorDeTeste.Localizar().PastaMotor, "fonts");

    private static void PularSemGerador() => Skip.If(!File.Exists(Allfontsgen),
        "allfontsgen.exe ausente em third_party/onlyoffice/converter — rode tools/onlyoffice/obter-allfontsgen.ps1");

    [SkippableFact] // gera catalogo numa pasta vazia a partir de UMA pasta de fontes (as do motor)
    public async Task GarantirAtualizado_PastaVazia_GeraOsDoisArquivos()
    {
        PularSemGerador();
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), Allfontsgen, [FontesDoMotor], _ => { });
        Assert.False(cat.Atualizado);

        var r = await cat.GarantirAtualizadoAsync();

        Assert.Null(r.Erro);
        Assert.True(r.Regenerou);
        Assert.True(File.Exists(Path.Combine(cat.PastaCatalogo, "AllFonts.js")));
        Assert.True(File.Exists(Path.Combine(cat.PastaCatalogo, "font_selection.bin")));
        Assert.True(r.Fontes >= 5, $"o motor traz pelo menos as 5 Open Sans; vieram {r.Fontes}");
        Assert.True(cat.Atualizado);
    }

    [SkippableFact] // os thumbnails que o editor pede (fonts_thumbnail*.png) saem do MESMO comando (--images)
    public async Task GarantirAtualizado_GeraOsThumbnailsQueOEditorPede()
    {
        PularSemGerador();
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), Allfontsgen, [FontesDoMotor]);

        Assert.Null((await cat.GarantirAtualizadoAsync()).Erro);

        // O sdkjs pede a base e as variantes por densidade de tela (@1.25x .. @2x) e a variante do leste
        // asiatico (_ea) — a mesma lista que o cache do app oficial tem nesta maquina.
        Assert.True(File.Exists(Path.Combine(cat.PastaCatalogo, "fonts_thumbnail.png")), "fonts_thumbnail.png não foi gerado");
        Assert.True(File.Exists(Path.Combine(cat.PastaCatalogo, "fonts_thumbnail@2x.png")), "fonts_thumbnail@2x.png não foi gerado");
        Assert.True(File.Exists(Path.Combine(cat.PastaCatalogo, "fonts_thumbnail_ea.png")), "fonts_thumbnail_ea.png não foi gerado");
    }

    [SkippableFact] // segunda chamada sem mudanca nas fontes NAO regenera (impressao digital igual)
    public async Task GarantirAtualizado_SemMudanca_NaoRegenera()
    {
        PularSemGerador();
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), Allfontsgen, [FontesDoMotor]);
        var r1 = await cat.GarantirAtualizadoAsync();

        var r2 = await cat.GarantirAtualizadoAsync();

        Assert.True(r1.Regenerou);
        Assert.False(r2.Regenerou);
        Assert.Null(r2.Erro);
        Assert.Equal(r1.Fontes, r2.Fontes);
    }

    [SkippableFact] // uma fonte nova na pasta muda a impressao digital e forca regenerar
    public async Task GarantirAtualizado_FonteNova_Regenera()
    {
        PularSemGerador();
        var fontes = Path.Combine(_raiz, "fontes");
        Directory.CreateDirectory(fontes);
        // DUAS faces DIFERENTES da mesma familia, uma de cada vez. Copiar o MESMO arquivo com outro nome nao
        // serviria: o allfontsgen deduplica por FACE (medido — as Open Sans Regular/Bold do motor sumiram do
        // nosso catalogo porque as mesmas faces ja estavam em C:\Windows\Fonts), entao a contagem nao mudaria
        // e o teste passaria sem provar nada.
        var primeira = Path.Combine(FontesDoMotor, "OpenSans-Light.ttf");
        var segunda = Path.Combine(FontesDoMotor, "OpenSans-Semibold.ttf");
        Skip.If(!File.Exists(primeira) || !File.Exists(segunda), "fontes do motor ausentes");
        File.Copy(primeira, Path.Combine(fontes, "OpenSans-Light.ttf"));

        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), Allfontsgen, [fontes]);
        var r1 = await cat.GarantirAtualizadoAsync();
        File.Copy(segunda, Path.Combine(fontes, "OpenSans-Semibold.ttf"));
        var r2 = await cat.GarantirAtualizadoAsync();

        Assert.True(r1.Regenerou);
        Assert.True(r2.Regenerou, "a fonte nova não disparou regeneração");
        Assert.Equal(r1.Fontes + 1, r2.Fontes);
    }

    [Fact] // gerador ausente: erro legivel em pt-BR, sem excecao, catalogo nao atualizado
    public async Task GarantirAtualizado_GeradorAusente_ErroLegivel()
    {
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), Path.Combine(_raiz, "nao-existe.exe"), [_raiz]);

        var r = await cat.GarantirAtualizadoAsync();

        Assert.NotNull(r.Erro);
        Assert.Contains("gerador de catálogo de fontes", r.Erro);
        Assert.False(r.Regenerou);
        Assert.False(cat.Atualizado);
    }

    [Fact] // impressao digital: determinista e sensivel ao conteudo/tamanho do arquivo
    public void ImpressaoDigital_MudaComArquivo()
    {
        var fontes = Path.Combine(_raiz, "f");
        Directory.CreateDirectory(fontes);
        File.WriteAllText(Path.Combine(fontes, "a.ttf"), "x");
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), "x", [fontes]);

        var d1 = cat.ImpressaoDigitalAtual();
        Assert.Equal(d1, cat.ImpressaoDigitalAtual());

        File.WriteAllText(Path.Combine(fontes, "a.ttf"), "xy");
        Assert.NotEqual(d1, cat.ImpressaoDigitalAtual());
    }

    [Fact] // controle negativo da impressao digital: arquivo que NAO e' fonte nao entra na conta (senao um
           // .log/.tmp escrito na pasta de fontes regeneraria o catalogo inteiro a cada abertura)
    public void ImpressaoDigital_IgnoraArquivoQueNaoEFonte()
    {
        var fontes = Path.Combine(_raiz, "g");
        Directory.CreateDirectory(fontes);
        File.WriteAllText(Path.Combine(fontes, "a.ttf"), "x");
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), "x", [fontes]);
        var antes = cat.ImpressaoDigitalAtual();

        File.WriteAllText(Path.Combine(fontes, "fonts.log"), "qualquer coisa");

        Assert.Equal(antes, cat.ImpressaoDigitalAtual());
    }

    // ---------------------------------------------------------------------------------------------------
    // Revisao 1 (fix review 1, Marco 1/Plano 1/Task 2) — itens 1, 3, 4 e o mtime do item 8.
    // ---------------------------------------------------------------------------------------------------

    [Fact] // item 3: a identidade do PROPRIO gerador (tamanho+mtime) entra na impressao digital — trocar o
           // allfontsgen.exe (nova versao pinada, correcao de bug nele) sem mudar nenhuma fonte no disco tem
           // de forcar regenerar
    public void ImpressaoDigital_MudaQuandoOGeradorMuda()
    {
        var fontes = Path.Combine(_raiz, "fg");
        Directory.CreateDirectory(fontes);
        File.WriteAllText(Path.Combine(fontes, "a.ttf"), "x");
        var gerador = Path.Combine(_raiz, "gerador-falso.exe");
        File.WriteAllText(gerador, "v1");
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), gerador, [fontes]);
        var antes = cat.ImpressaoDigitalAtual();

        File.WriteAllText(gerador, "v2 - conteudo maior, tamanho e mtime diferentes");

        Assert.NotEqual(antes, cat.ImpressaoDigitalAtual());
    }

    [Fact] // item 3: bump de VersaoFormatoCatalogo (aqui via o parametro de construtor, que existe so' para
           // este teste poder provar o efeito sem InternalsVisibleTo nem reflexao) muda a digital mesmo sem
           // nenhuma fonte ou gerador terem mudado — e' o que forca regenerar todo catalogo existente quando o
           // FORMATO esperado muda
    public void ImpressaoDigital_MudaComVersaoDoFormato()
    {
        var fontes = Path.Combine(_raiz, "fv");
        Directory.CreateDirectory(fontes);
        File.WriteAllText(Path.Combine(fontes, "a.ttf"), "x");
        var cat1 = new CatalogoFontes(Path.Combine(_raiz, "cat"), "x", [fontes], versaoFormato: 1);
        var cat2 = new CatalogoFontes(Path.Combine(_raiz, "cat"), "x", [fontes], versaoFormato: 2);

        Assert.NotEqual(cat1.ImpressaoDigitalAtual(), cat2.ImpressaoDigitalAtual());
    }

    [Fact] // item 8 (minor): so' o mtime muda (mesmo conteudo, mesmo tamanho) — prende o componente mtime da
           // impressao digital isoladamente do tamanho/conteudo, que os outros testes ja cobrem
    public void ImpressaoDigital_MudaSoComMtime()
    {
        var fontes = Path.Combine(_raiz, "fm");
        Directory.CreateDirectory(fontes);
        var arquivo = Path.Combine(fontes, "a.ttf");
        File.WriteAllText(arquivo, "mesmo-conteudo-mesmo-tamanho");
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat"), "x", [fontes]);
        var antes = cat.ImpressaoDigitalAtual();

        File.SetLastWriteTimeUtc(arquivo, DateTime.UtcNow.AddMinutes(5));

        Assert.NotEqual(antes, cat.ImpressaoDigitalAtual());
    }

    [Fact] // item 1: a mensagem distingue o cancelamento PEDIDO PELO CHAMADOR do teto INTERNO de 5 minutos —
           // funcao pura `internal`, testada direto para nao depender de um processo real pendurado 5 minutos
    public void MensagemDeCancelamento_DistingueChamadorDoTetoInterno()
    {
        var doChamador = CatalogoFontes.MensagemDeCancelamento(cancelamentoDoChamador: true);
        var doTetoInterno = CatalogoFontes.MensagemDeCancelamento(cancelamentoDoChamador: false);

        Assert.Contains("geração do catálogo cancelada", doChamador, StringComparison.Ordinal);
        Assert.Contains("5 minutos", doTetoInterno, StringComparison.Ordinal);
        Assert.NotEqual(doChamador, doTetoInterno);
    }

    [SkippableFact] // item 4: duas chamadas concorrentes na MESMA pasta (mesma instancia) — o mutex global
                     // serializa, e o recheck de Atualizado depois de adquirir garante que SO' UMA regenera; a
                     // outra reaproveita o catalogo que a primeira acabou de escrever (nunca 0 fontes/erro)
    public async Task GarantirAtualizado_DuasChamadasConcorrentes_SoUmaRegenera()
    {
        PularSemGerador();
        var cat = new CatalogoFontes(Path.Combine(_raiz, "cat-concorrente"), Allfontsgen, [FontesDoMotor]);

        var t1 = cat.GarantirAtualizadoAsync();
        var t2 = cat.GarantirAtualizadoAsync();
        var resultados = await Task.WhenAll(t1, t2);

        Assert.Single(resultados, r => r.Regenerou);
        Assert.All(resultados, r => Assert.Null(r.Erro));
        Assert.All(resultados, r => Assert.True(r.Fontes > 0, $"catalogo com 0 fontes (erro: {r.Erro})"));
    }

    // ---------------------------------------------------------------------------------------------------
    // Fix round 1 da Task 1 do Plano 5: o AllFonts.js NAO e' mais copiado para dentro do motor (instalado ele
    // fica em Program Files, so' leitura); o host o serve do catalogo (CatalogoNoEditor). Antes de navegar a
    // janela so' CONFERE que o arquivo do catalogo existe — e falha fechado se nao existir.
    // ---------------------------------------------------------------------------------------------------

    private static ConfiguracaoMotor MotorEm(string pastaMotor, string pastaFontes)
        => new(pastaMotor) { PastaFontes = pastaFontes };

    [Fact] // catalogo ausente: nao lanca, fecha (Ok = false) e a mensagem nomeia a pasta onde o arquivo era esperado
    public void ConferirParaOEditor_SemCatalogo_FalhaFechadoNomeandoAPasta()
    {
        var fontes = Path.Combine(_raiz, "fontes-que-nao-existem");
        var r = CatalogoFontes.ConferirParaOEditor(MotorEm(Path.Combine(_raiz, "motor"), fontes));

        Assert.False(r.Ok);
        Assert.Contains("AUSENTE", r.Mensagem, StringComparison.Ordinal);
        Assert.Contains("o editor nao vai carregar", r.Mensagem, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(fontes, r.Mensagem, StringComparison.OrdinalIgnoreCase);
    }

    [Fact] // com o catalogo: Ok, e NADA e' escrito no motor (a pasta do motor nem existe e continua nao existindo)
    public void ConferirParaOEditor_ComCatalogo_OkSemEscreverNoMotor()
    {
        var motor = Path.Combine(_raiz, "motor-inexistente");
        var fontes = Path.Combine(_raiz, "fontes");
        Directory.CreateDirectory(fontes);
        File.WriteAllText(Path.Combine(fontes, "AllFonts.js"), "window[\"__fonts_files\"] = [];");

        var r = CatalogoFontes.ConferirParaOEditor(MotorEm(motor, fontes));

        Assert.True(r.Ok, r.Mensagem);
        Assert.Contains("servido do catalogo pelo host", r.Mensagem, StringComparison.Ordinal);
        Assert.False(Directory.Exists(motor));
    }

    public void Dispose() { try { Directory.Delete(_raiz, true); } catch { } }
}
