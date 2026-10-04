using System.Text.Json;
using System.Text.RegularExpressions;
using mOffice.Conversor;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Revisao final do Plano 6, C1 (seguranca): <see cref="PonteDesktop.Avisar"/> e' publico (objeto de host COM-visivel) e entregava QUALQUER
/// tipo ao host — inclusive os pedidos que so' a propria ponte emite depois de conferir o consentimento (<c>converterArquivo</c>,
/// <c>lerArquivo</c>, <c>abrirReferenciaExterna</c>, <c>vinculoRecusado</c>, <c>liberarConversao</c>). Um script da pagina forjava o pedido com
/// qualquer caminho (local ou UNC) e <c>confirmar:false</c>, passando por cima de PedirLeitura/PedirConversao/VinculoPermitido. Duas camadas:
/// (a) o <see cref="PonteDesktop.Avisar"/> recusa esses tipos; (b) o host so' aceita pedido cujo id a ponte emitiu, e usa o pedido GUARDADO.
/// </summary>
public class PonteAvisosReservadosTests : IDisposable
{
    private readonly string _pasta = Directory.CreateTempSubdirectory("moffice-c1-").FullName;
    private readonly List<(string Tipo, string Param)> _avisos = [];
    private readonly List<string> _log = [];
    private readonly SessaoDocumento _sessao;
    private readonly PonteDesktop _ponte;

    public PonteAvisosReservadosTests()
    {
        Directory.CreateDirectory(Path.Combine(_pasta, "docs"));
        _sessao = new SessaoDocumento
        {
            ArquivoOrigem = Path.Combine(_pasta, "docs", "aberto.xlsx"), Formato = FormatoDocumento.Xlsx,
            PastaRecovery = Directory.CreateDirectory(Path.Combine(_pasta, "recovery")).FullName,
        };
        _ponte = new PonteDesktop(_sessao, (t, p) => _avisos.Add((t, p)), log: _log.Add);
    }

    public void Dispose() { try { Directory.Delete(_pasta, true); } catch { } GC.SuppressFinalize(this); }

    private string Arquivo(string nome)
    {
        var c = Path.Combine(_pasta, "docs", nome);
        File.WriteAllBytes(c, [1, 2, 3]);
        return c;
    }

    [Theory] // pedido forjado pela pagina: nenhuma entrega ao host (logo, nenhuma E/S no caminho), e a recusa fica no log
    [InlineData("lerArquivo")]
    [InlineData("converterArquivo")]
    [InlineData("abrirReferenciaExterna")]
    [InlineData("vinculoRecusado")]
    [InlineData("liberarConversao")]
    public void Avisar_TipoReservadoAoHost_Recusado(string tipo)
    {
        var segredo = Arquivo("segredo.csv");
        var forjado = JsonSerializer.Serialize(new { id = "0123456789abcdef0123456789abcdef", caminho = segredo, formato = 8194, uso = "converter", confirmar = false });

        _ponte.Avisar(tipo, forjado);

        Assert.Empty(_avisos);
        Assert.Contains("PONTE aviso reservado ao host recusado: " + tipo, _log);
    }

    [Theory] // a grafia nao contorna a recusa: a comparacao e' exata, e um tipo com outra caixa nao e' tratado pelo host (switch ordinal)
    [InlineData("pronto")] [InlineData("localFileSave")] [InlineData("abrirArquivo")] [InlineData("recentes")] [InlineData("imprimir")]
    public void Avisar_TipoDaPagina_Entregue(string tipo)
    {
        _ponte.Avisar(tipo, "x");
        Assert.Equal((tipo, "x"), Assert.Single(_avisos));
    }

    [Fact] // o caminho legitimo continua: dialogo do host -> PedirLeitura -> pedido entregue pelo canal interno, com o id emitido
    public void PedirLeitura_Legitima_ContinuaEntregue()
    {
        var c = Arquivo("dados.csv");
        _ponte.PermitirLeitura(c);
        var id = _ponte.PedirLeitura(c);
        var (tipo, param) = Assert.Single(_avisos);
        Assert.Equal("lerArquivo", tipo);
        Assert.Equal(param, JanelaEditor.PedidoEmitido(_ponte, tipo, param, _log.Add));
        using var d = JsonDocument.Parse(param);
        Assert.Equal(id, d.RootElement.GetProperty("id").GetString());
    }

    [Fact] // camada (b): id que a ponte nao emitiu -> recusado; id emitido -> o pedido GUARDADO (caminho da pagina ignorado); um uso so'
    public void PedidoEmitido_SoIdsDaPonte_UsaOGuardado()
    {
        var c = Arquivo("dados.csv");
        var outro = Arquivo("outro.csv");
        Assert.Null(JanelaEditor.PedidoEmitido(_ponte, "lerArquivo",
            JsonSerializer.Serialize(new { id = "0123456789abcdef0123456789abcdef", caminho = outro }), _log.Add));
        Assert.Contains(_log, l => l.StartsWith("PONTE pedido nao emitido pela ponte recusado: lerArquivo", StringComparison.Ordinal));

        _ponte.PermitirLeitura(c);
        var id = _ponte.PedirLeitura(c);
        var guardado = _avisos.Single().Param;
        // mesmo id com o caminho trocado: o host usa o que a ponte guardou
        var trocado = JsonSerializer.Serialize(new { id, caminho = outro });
        Assert.Equal(guardado, JanelaEditor.PedidoEmitido(_ponte, "lerArquivo", trocado, _log.Add));
        // um uso so'
        Assert.Null(JanelaEditor.PedidoEmitido(_ponte, "lerArquivo", trocado, _log.Add));
    }

    [Fact] // id emitido para um tipo nao vale para outro (leitura nao vira conversao)
    public void PedidoEmitido_TipoDiferente_Recusado()
    {
        var c = Arquivo("dados.csv");
        _ponte.PermitirLeitura(c);
        var id = _ponte.PedirLeitura(c);
        Assert.Null(JanelaEditor.PedidoEmitido(_ponte, "converterArquivo", JsonSerializer.Serialize(new { id, caminho = c, formato = 8193, uso = "converter" }), _log.Add));
    }

    [Fact] // vinculo: o confirmar guardado (true) prevalece sobre um confirmar:false da pagina
    public void PedidoEmitido_Vinculo_ConfirmarGuardado()
    {
        var xlsx = Arquivo("origem.xlsx");
        var id = _ponte.AbrirReferenciaExterna(xlsx);
        var forjado = JsonSerializer.Serialize(new { id, caminho = xlsx, confirmar = false });
        var p = JanelaEditor.PedidoEmitido(_ponte, "abrirReferenciaExterna", forjado, _log.Add);
        Assert.NotNull(p);
        using var d = JsonDocument.Parse(p);
        Assert.True(d.RootElement.GetProperty("confirmar").GetBoolean());
    }

    [Fact] // invariante de cobertura: todo tipo que a PROPRIA ponte emite pelo canal interno e' reservado; nenhum tipo que o JS manda e' reservado
    public void TiposSoDoHost_CobremOsEmitidosPelaPonteENaoOsDaPagina()
    {
        var raiz = FronteiraTests.RaizRepo();
        var cs = File.ReadAllText(Path.Combine(raiz, "src", "mOffice.Nucleo", "PonteDesktop.cs"));
        var js = File.ReadAllText(Path.Combine(raiz, "src", "mOffice.Nucleo", "Ponte", "ponte-desktop.js"));
        var daPonte = Regex.Matches(cs, @"(?:_aviso\?\.Invoke|EmitirPedido)\(""([A-Za-z]+)""").Select(m => m.Groups[1].Value).Distinct().ToList();
        var daPagina = Regex.Matches(js, @"avisar\(""([A-Za-z]+)""").Select(m => m.Groups[1].Value).Distinct().ToList();
        Assert.True(daPonte.Count >= 5, "a varredura do PonteDesktop.cs nao achou os pedidos internos: " + string.Join(",", daPonte));
        Assert.True(daPagina.Count >= 10, "a varredura do ponte-desktop.js nao achou os avisos da pagina: " + string.Join(",", daPagina));
        Assert.All(daPonte, t => Assert.Contains(t, PonteDesktop.TiposSoDoHost));
        Assert.All(daPagina, t => Assert.DoesNotContain(t, PonteDesktop.TiposSoDoHost));
    }
}
