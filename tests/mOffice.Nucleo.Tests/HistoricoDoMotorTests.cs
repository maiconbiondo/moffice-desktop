using System.Text.RegularExpressions;
using mOffice.Conversor.Tests;

namespace mOffice.Nucleo.Tests;

/// <summary>
/// Guarda de deriva do motor (PL-02, Task 1): <c>JanelaEditor.MarcarModificadoNoEditorAsync</c> acha o campo "ForceSave" do
/// <c>AscCommon.History</c> lendo o codigo da <c>Have_Changes</c> em tempo de execucao (o nome e' minificado e muda por
/// editor). Este teste aplica a MESMA expressao ao <c>sdk-all.js</c> de cada editor, sem abrir janela (vale tambem no RDP):
/// se uma atualizacao do motor mudar a forma da funcao, o vermelho aparece aqui e nao so' no teste de janela.
/// </summary>
public class HistoricoDoMotorTests
{
    /// <summary>A expressao de <c>MarcarModificadoNoEditorAsync</c> (manter igual).</summary>
    private static readonly Regex CampoForceSave = new(@"&&!1===this\.([\w$]+)\?", RegexOptions.CultureInvariant);

    [Theory]
    [InlineData("word")]
    [InlineData("cell")]
    [InlineData("slide")]
    public void HaveChanges_TemUmaSoOcorrencia_ComCampoBooleano(string editor)
    {
        var motor = MotorDeTeste.Localizar();
        var arquivo = Path.Combine(motor.PastaMotor, "editors", "sdkjs", editor, "sdk-all.js");
        Skip.IfNot(File.Exists(arquivo), $"sdk-all.js de {editor} ausente: {arquivo}");
        var js = File.ReadAllText(arquivo);

        const string ancora = "LocalFileGetOpenChangesCount";
        var pos = js.IndexOf(ancora, StringComparison.Ordinal);
        Assert.True(pos >= 0, $"{editor}: a Have_Changes (que cita {ancora}) sumiu do sdk-all.js");
        Assert.Equal(-1, js.IndexOf(ancora, pos + 1, StringComparison.Ordinal)); // exatamente uma funcao do historico

        var inicio = js.LastIndexOf("function(", pos, StringComparison.Ordinal);
        Assert.True(inicio >= 0);
        var corpo = js.Substring(inicio, Math.Min(js.Length - inicio, pos - inicio + 600));
        var achados = CampoForceSave.Matches(corpo);
        Assert.NotEmpty(achados);
        // a funcao pode citar o campo mais de uma vez (word/slide: ramo do usuario e ramo geral), sempre o MESMO campo
        Assert.Single(achados.Select(a => a.Groups[1].Value).Distinct());
        var campo = Regex.Escape(achados[0].Groups[1].Value);
        // e e' uma bandeira booleana do historico (inicializada/zerada com !1 = false)
        Assert.Matches($@"this\.{campo}=!1", js);
        // e o corpo usa o campo como a clausula ForceSave da Have_Changes ("... || ForceSave === true"), fixando o seu significado
        Assert.Matches($@"!0===this\.{campo}", corpo);
    }
}
