using System.Windows.Media;
using mOffice.Comum.Icones;
using mOffice.Nucleo.Tests;

namespace mOffice.Comum.Tests;

[Collection("janela real")]
public class PhosphorTests
{
    [Fact] public void Mapa_Tem58Icones() => Assert.Equal(58, Ph.Glifos.Count);
    [Fact] public void TodoNome_ResolveUmCaractere() { foreach (var n in Ph.Glifos.Keys) Assert.Equal(1, Ph.Glyph(n).Length); }
    [Fact] public void NomeDesconhecido_Lanca() => Assert.Throws<KeyNotFoundException>(() => Ph.Glyph("nao-existe"));

    [Fact] // todo glifo existe nas tres fontes embarcadas (GlyphTypeface real)
    public void TodoGlifo_ExisteNasFontes()
    {
        JanelaSta.Executar(() =>
        {
            foreach (var arquivo in new[] { "Phosphor.ttf", "Phosphor-Fill.ttf", "Phosphor-Bold.ttf" })
            {
                var gt = new GlyphTypeface(new Uri($"pack://application:,,,/mOffice.Comum;component/Assets/Fonts/{arquivo}"));
                foreach (var (nome, glifo) in Ph.Glifos)
                    Assert.True(gt.CharacterToGlyphMap.ContainsKey(glifo[0]), $"{nome} falta em {arquivo}");
            }
            return Task.CompletedTask;
        });
    }
}
