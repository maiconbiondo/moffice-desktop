using System.Text;
using mOffice.Conversor;
using mOffice.Nucleo;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// SalvarChanges sem janela: e' pura manipulacao de estado + arquivo. Cobre o achado 2 da revisao 2026-09-17
/// (deleteIndex &lt; 0 nao pode truncar changes0.json — ver o comentario de <see cref="PonteDesktop.SalvarChanges"/>).
/// </summary>
public class PonteDesktopTests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("moffice-changes-").FullName;
    private readonly SessaoDocumento _sessao;
    private readonly PonteDesktop _ponte;

    public PonteDesktopTests()
    {
        _sessao = new SessaoDocumento { ArquivoOrigem = "doc.docx", Formato = FormatoDocumento.Docx, PastaRecovery = _pasta };
        _ponte = new PonteDesktop(_sessao);
    }

    public void Dispose() { try { Directory.Delete(_pasta, true); } catch { } GC.SuppressFinalize(this); }

    private string LerArquivo() => File.ReadAllText(_sessao.ArquivoChanges, Encoding.UTF8);

    [Fact]
    public void TresAppends_GravaSequenciaEmUtf8SemBom()
    {
        _ponte.SalvarChanges("c1", 0, 1);
        _ponte.SalvarChanges("c2", 1, 1);
        _ponte.SalvarChanges("c3", 2, 1);

        Assert.Equal("\"c1\",\"c2\",\"c3\",", LerArquivo());
        Assert.Equal(3, _sessao.IndiceChanges);
        // Sem BOM: os 3 primeiros bytes tem de ser o '"' inicial (0x22), nao o EF BB BF do UTF-8 com BOM.
        var bytes = File.ReadAllBytes(_sessao.ArquivoChanges);
        Assert.Equal((byte)'"', bytes[0]);
    }

    [Fact]
    public void DeleteIndexDentroDoIntervalo_TruncaEApendaAPartirDele()
    {
        _ponte.SalvarChanges("c1", 0, 1);
        _ponte.SalvarChanges("c2", 1, 1);
        _ponte.SalvarChanges("c3", 2, 1); // IndiceChanges=3, arquivo "c1","c2","c3",

        _ponte.SalvarChanges("c4", 1, 1);

        Assert.Equal("\"c1\",\"c4\",", LerArquivo());
        Assert.Equal(2, _sessao.IndiceChanges);
    }

    [Fact] // achado 2: deleteIndex negativo NAO pode truncar — so apenda a partir do indice atual
    public void DeleteIndexNegativo_NaoTruncaNadaESoApenda()
    {
        _ponte.SalvarChanges("c1", 0, 1);
        _ponte.SalvarChanges("c2", 1, 1);
        _ponte.SalvarChanges("c3", 2, 1); // IndiceChanges=3, arquivo "c1","c2","c3",

        _ponte.SalvarChanges("c4", -1, 1);

        Assert.Equal("\"c1\",\"c2\",\"c3\",\"c4\",", LerArquivo());
        Assert.Equal(4, _sessao.IndiceChanges);
    }

    [Fact] // deleteIndex 0 e' "desfazer tudo": esvazia o arquivo legitimamente (nao e' o bug do achado 2)
    public void DeleteIndexZeroComCountZero_RemoveOArquivoEZeraIndice()
    {
        _ponte.SalvarChanges("c1", 0, 1);
        _ponte.SalvarChanges("c2", 1, 1);
        _ponte.SalvarChanges("c3", 2, 1); // IndiceChanges=3

        _ponte.SalvarChanges("", 0, 0);

        Assert.False(File.Exists(_sessao.ArquivoChanges));
        Assert.Equal(0, _sessao.IndiceChanges);
    }

    [Fact]
    public void CountZeroComDeleteIndexMaiorOuIgualAoIndiceAtual_NaoFazNada()
    {
        _ponte.SalvarChanges("c1", 0, 1);
        _ponte.SalvarChanges("c2", 1, 1); // IndiceChanges=2, arquivo "c1","c2",

        _ponte.SalvarChanges("", 5, 0);

        Assert.Equal("\"c1\",\"c2\",", LerArquivo());
        Assert.Equal(2, _sessao.IndiceChanges);
    }

    // -----------------------------------------------------------------------------------------------------
    // ImagemUrl = LocalFileGetImageUrl (Task 5/S4). Contrato do desktop-sdk oficial: devolve o NOME do
    // arquivo em media/, ou "error". Aqui sem janela: a decisao de QUE dado vira midia e' pura.
    // -----------------------------------------------------------------------------------------------------

    private string[] Media() { var m = Path.Combine(_pasta, "media"); return Directory.Exists(m) ? [.. Directory.GetFiles(m).Select(Path.GetFileName)!] : []; }

    private static string DataUrlPngVermelho(int lado = 8) =>
        "data:image/png;base64," + Convert.ToBase64String(mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(lado, 255, 0, 0));

    [Fact] // colar: data: URL vira arquivo em media/ e o nome volta para o sdkjs montar a URL
    public void ImagemUrl_DataUrl_GravaEmMediaEDevolveONome()
    {
        var nome = _ponte.ImagemUrl(DataUrlPngVermelho());

        Assert.Equal("image1.png", nome);
        Assert.Equal(["image1.png"], Media());
    }

    [Fact] // o mesmo conteudo colado duas vezes nao duplica a midia (o oficial guarda o mesmo mapa, por CRC32)
    public void ImagemUrl_MesmoConteudoDuasVezes_ReaproveitaOArquivo()
    {
        var dados = DataUrlPngVermelho();

        Assert.Equal(_ponte.ImagemUrl(dados), _ponte.ImagemUrl(dados));
        Assert.Single(Media());
    }

    [Fact] // nome de midia que ja esta no documento passa direto (oficial: `if (sUrl.find(L"image") == 0) return sUrl;`)
    public void ImagemUrl_NomeDeMidiaDoDocumento_PassaDireto()
    {
        Assert.Equal("image7.png", _ponte.ImagemUrl("image7.png"));
        Assert.Empty(Media());
    }

    [Fact] // ... mas nunca como CAMINHO: "image" no comeco nao pode virar passeio para fora de media/
    public void ImagemUrl_NomeComCaminho_ERecusado()
    {
        Assert.Equal("error", _ponte.ImagemUrl("image/../../Editor.bin"));
        Assert.Equal("error", _ponte.ImagemUrl("images\\..\\..\\Editor.bin"));
    }

    [Fact] // egresso zero: o app oficial BAIXARIA a imagem (CFileDownloaderWrapper); este host nao
    public void ImagemUrl_UrlDeRede_ERecusadaSemBaixar()
    {
        Assert.Equal("error", _ponte.ImagemUrl("https://exemplo.invalido/figura.png"));
        Assert.Equal("error", _ponte.ImagemUrl("http://exemplo.invalido/figura.png"));
        Assert.Equal("error", _ponte.ImagemUrl("www.exemplo.invalido/figura.png"));
        Assert.Empty(Media());
    }

    [Fact] // conteudo web NAO escolhe arquivo do disco: sem o dialogo do usuario, nenhum caminho e' lido
    // (mesma defesa do g_pLocalResolver->Check do app oficial). Sem isto, uma pagina do editor mandaria
    // copiar QUALQUER arquivo da maquina para media/ — que e' servido de volta para ela por https.
    public void ImagemUrl_ArquivoNaoEscolhidoPeloUsuario_ERecusado()
    {
        var fora = Path.Combine(_pasta, "segredo.png");
        File.WriteAllBytes(fora, mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(8, 0, 255, 0));

        Assert.Equal("error", _ponte.ImagemUrl(fora));
        Assert.Equal("error", _ponte.ImagemUrl("file:///" + fora.Replace('\\', '/')));
        Assert.Empty(Media());
    }

    [Fact] // ... e depois que o usuario escolhe no dialogo do host, o MESMO caminho e' aceito
    public void ImagemUrl_ArquivoEscolhidoPeloUsuario_EGravadoEmMedia()
    {
        var escolhido = Path.Combine(_pasta, "escolhida.png");
        var bytes = mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(8, 255, 0, 0);
        File.WriteAllBytes(escolhido, bytes);

        _ponte.PermitirArquivoEscolhido(escolhido);

        Assert.Equal("image1.png", _ponte.ImagemUrl(escolhido));
        Assert.Equal(bytes, File.ReadAllBytes(Path.Combine(_pasta, "media", "image1.png")));
    }

    [Fact] // arquivo escolhido que NAO e' imagem nao entra em media/ (o oficial devolve "error" pelo checker)
    public void ImagemUrl_ArquivoEscolhidoQueNaoEImagem_ERecusado()
    {
        var escolhido = Path.Combine(_pasta, "escolhido.txt");
        File.WriteAllText(escolhido, "isto nao e uma imagem");
        _ponte.PermitirArquivoEscolhido(escolhido);

        Assert.Equal("error", _ponte.ImagemUrl(escolhido));
        Assert.Empty(Media());
    }

    // -----------------------------------------------------------------------------------------------------
    // Consentimento de UM USO SO' (minor 2 da revisao 1): PermitirArquivoEscolhido autoriza aquela
    // colagem/insercao, nao um acesso permanente da pagina ao arquivo.
    // -----------------------------------------------------------------------------------------------------

    [Fact] // o mesmo caminho, sem passar de novo pelo dialogo do host, e' recusado na SEGUNDA vez
    public void ImagemUrl_MesmoArquivoEscolhidoDuasVezes_SegundaVezERecusada()
    {
        var escolhido = Path.Combine(_pasta, "escolhida-unica.png");
        var bytes = mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(8, 255, 0, 0);
        File.WriteAllBytes(escolhido, bytes);
        _ponte.PermitirArquivoEscolhido(escolhido);

        Assert.Equal("image1.png", _ponte.ImagemUrl(escolhido));
        Assert.Equal("error", _ponte.ImagemUrl(escolhido));
        Assert.Equal(["image1.png"], Media());
    }

    [Fact] // um NOVO dialogo (PermitirArquivoEscolhido de novo) reautoriza o MESMO caminho
    public void ImagemUrl_ArquivoEscolhidoDeNovoAposConsumido_EAceitoOutraVez()
    {
        var escolhido = Path.Combine(_pasta, "escolhida-de-novo.png");
        var bytes = mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(8, 255, 0, 0);
        File.WriteAllBytes(escolhido, bytes);
        _ponte.PermitirArquivoEscolhido(escolhido);
        Assert.Equal("image1.png", _ponte.ImagemUrl(escolhido));

        _ponte.PermitirArquivoEscolhido(escolhido); // um novo dialogo do host

        // Mesmo conteudo, entao a dedupe por SHA-256 devolve o mesmo nome (nao grava de novo) — o que importa
        // aqui e' que a segunda chamada NAO caia em "error" so' por o caminho ja ter sido consumido antes.
        Assert.Equal("image1.png", _ponte.ImagemUrl(escolhido));
    }

    [Fact] // conteudo escolhido que NAO e' imagem: Gravar lanca, entao NAO houve "uso" — o dialogo continua valendo
    public void ImagemUrl_ArquivoEscolhidoQueNaoEImagem_NaoConsomeAEntrada()
    {
        var escolhido = Path.Combine(_pasta, "nao-e-imagem.png");
        File.WriteAllBytes(escolhido, System.Text.Encoding.UTF8.GetBytes("nao sou imagem"));
        _ponte.PermitirArquivoEscolhido(escolhido);

        Assert.Equal("error", _ponte.ImagemUrl(escolhido)); // Gravar lanca InvalidDataException, ImagemUrl vira "error"
        // Se a entrada tivesse sido consumida na primeira tentativa (mesmo falhando), esta segunda chamada
        // cairia em "error" pela recusa de "nao escolhido" em vez de repetir a recusa de "nao e' imagem" —
        // ambas devolvem "error", mas o log (nao verificado aqui) diria motivos diferentes. O que este teste
        // prova e' que o caminho AINDA esta' na lista apos a falha.
        File.WriteAllBytes(escolhido, mOffice.Conversor.Tests.FabricaDocumentos.PngSolido(8, 0, 255, 0));
        Assert.Equal("image1.png", _ponte.ImagemUrl(escolhido));
    }

    // -----------------------------------------------------------------------------------------------------
    // SVG por `data:` e' recusado DE PROPOSITO, com erro proprio (minor 3 da revisao 1)
    // -----------------------------------------------------------------------------------------------------

    [Fact]
    public void ImagemUrl_DataUrlSvg_ERecusadoComErroProprio()
    {
        var svg = "<svg xmlns='http://www.w3.org/2000/svg'><script>alert(1)</script></svg>";
        var dataUrl = "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

        Assert.Equal("formato-nao-suportado", _ponte.ImagemUrl(dataUrl));
        Assert.Empty(Media());
    }

    [Fact] // variacao de maiusculas/minusculas e ";charset=utf-8" antes do "base64," continuam sendo SVG
    public void ImagemUrl_DataUrlSvgComVariacaoDeGrafia_ERecusadoComErroProprio()
    {
        var svg = "<svg><script>alert(1)</script></svg>";
        var dataUrl = "DATA:IMAGE/SVG+XML;charset=utf-8;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

        Assert.Equal("formato-nao-suportado", _ponte.ImagemUrl(dataUrl));
        Assert.Empty(Media());
    }

    // -----------------------------------------------------------------------------------------------------
    // Farejo do CONTEUDO (fix 2 da revisao da Task 5, item 3): o MIME da `data:` URL e' declarado pela
    // pagina; mentir "image/png" com texto SVG dentro nao pode escapar da recusa propria acima. O controle
    // positivo e' ImagemUrl_DataUrl_GravaEmMediaEDevolveONome (png de verdade continua aceito).
    // -----------------------------------------------------------------------------------------------------

    [Fact] // data:image/png com <svg> dentro: recusa PROPRIA, nao o "error" generico de "nao e' imagem"
    public void ImagemUrl_DataUrlPngComConteudoSvg_ERecusadoComErroProprio()
    {
        var svg = "<svg xmlns='http://www.w3.org/2000/svg'><script>alert(1)</script></svg>";
        var dataUrl = "data:image/png;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));

        Assert.Equal("formato-nao-suportado", _ponte.ImagemUrl(dataUrl));
        Assert.Empty(Media());
    }

    [Fact] // prologo <?xml ...> + BOM UTF-8 + espacos antes: continua sendo XML para o farejo
    public void ImagemUrl_DataUrlPngComPrologoXml_ERecusadoComErroProprio()
    {
        var xml = "\uFEFF \r\n\t<?xml version='1.0' encoding='UTF-8'?><svg><script>alert(1)</script></svg>";
        var dataUrl = "data:image/png;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(xml));

        Assert.Equal("formato-nao-suportado", _ponte.ImagemUrl(dataUrl));
        Assert.Empty(Media());
    }

    // -----------------------------------------------------------------------------------------------------
    // MarcarModificado (Task 2 do Marco 1 / Plano 2). O editor chama onDocumentModifiedChanged a CADA tecla,
    // varias vezes com o mesmo valor; a casca redesenharia o asterisco a toa. O callback so' sai na MUDANCA.
    // Nao precisa de janela: e' o parametro opcional do construtor + um bool de sessao.
    // -----------------------------------------------------------------------------------------------------

    private (PonteDesktop ponte, List<bool> avisos) PonteComAviso()
    {
        var avisos = new List<bool>();
        return (new PonteDesktop(_sessao, aoMudarModificado: v => avisos.Add(v)), avisos);
    }

    [Fact] // false -> true avisa uma vez; repetir o true NAO avisa de novo
    public void MarcarModificado_SoAvisaNaMudanca()
    {
        var (ponte, avisos) = PonteComAviso();

        ponte.MarcarModificado(true);
        Assert.Equal(new[] { true }, avisos);
        Assert.True(_sessao.Modificado);

        ponte.MarcarModificado(true);   // mesmo valor: nada
        ponte.MarcarModificado(true);
        Assert.Equal(new[] { true }, avisos);

        ponte.MarcarModificado(false);  // volta: avisa
        Assert.Equal(new[] { true, false }, avisos);
        Assert.False(_sessao.Modificado);

        ponte.MarcarModificado(false);  // mesmo valor: nada
        Assert.Equal(new[] { true, false }, avisos);
    }

    [Fact] // false -> false (o estado inicial da sessao) nao e' mudanca: nao avisa
    public void MarcarModificado_MesmoValorInicial_NaoAvisa()
    {
        var (ponte, avisos) = PonteComAviso();
        ponte.MarcarModificado(false);
        Assert.Empty(avisos);
    }

    [Fact] // o valor JA esta gravado na sessao quando o callback roda (a casca le' Modificado de dentro dele)
    public void MarcarModificado_CallbackVeOEstadoNovo()
    {
        bool? vistoNaSessao = null;
        var ponte = new PonteDesktop(_sessao, aoMudarModificado: _ => vistoNaSessao = _sessao.Modificado);
        ponte.MarcarModificado(true);
        Assert.True(vistoNaSessao);
    }

    [Fact] // sem callback (o construtor antigo, 1 parametro) continua funcionando: so' grava na sessao
    public void MarcarModificado_SemCallback_SoGrava()
    {
        _ponte.MarcarModificado(true);
        Assert.True(_sessao.Modificado);
        _ponte.MarcarModificado(false);
        Assert.False(_sessao.Modificado);
    }
}
