namespace mOffice.Comum.Dialogos;

/// <summary>Atribuicao completa mostrada na tela "Sobre": um item por componente que o instalador
/// redistribui. Os dados moram em <c>Atribuicoes.txt</c> (recurso embutido), espelho da tabela de
/// <c>docs/licencas/INVENTARIO.md</c> — o teste <c>LicencasTests.Atribuicoes_IgualAoInventario</c> confere as
/// duas direcoes e <c>DialogosTests.Sobre_ListaTodoOInventario</c> confere o texto renderizado. Ao acrescentar
/// um componente ao pacote, acrescentar nos DOIS lugares.
///
/// Por que .txt e nao uma lista em C#: sao nomes de componentes (o SDK do WebView2, bibliotecas de rede do
/// motor), ou seja DADO de atribuicao; as guardas de fronteira varrem o CODIGO e nao devem ser afrouxadas por
/// causa de um texto de licenca.</summary>
public static class Atribuicoes
{
    /// <param name="Componente">Nome IGUAL ao da primeira coluna do INVENTARIO.md.</param>
    public sealed record Item(string Componente, string Versao, string Licenca)
    {
        /// <summary>A segunda linha do item no Sobre: "versao · licenca".</summary>
        public string Detalhe => $"{Versao} · {Licenca}";
    }

    /// <param name="versaoApp">Versao do mOffice (Directory.Build.props).</param>
    /// <param name="versaoMotor">Versao do ONLYOFFICE embarcado (<c>mOffice.Nucleo.Motor.Versao</c>).</param>
    public static IReadOnlyList<Item> Itens(string versaoApp, string versaoMotor)
    {
        using var s = typeof(Atribuicoes).Assembly.GetManifestResourceStream("mOffice.Comum.Dialogos.Atribuicoes.txt")
            ?? throw new InvalidOperationException("Recurso Atribuicoes.txt ausente do mOffice.Comum.");
        using var r = new StreamReader(s);
        var itens = new List<Item>();
        while (r.ReadLine() is { } linha)
        {
            if (linha.Length == 0 || linha.StartsWith('#')) continue;
            var c = linha.Split('\t');
            if (c.Length != 3) throw new InvalidOperationException($"Linha invalida em Atribuicoes.txt: {linha}");
            string V(string x) => x.Replace("{versao}", versaoApp).Replace("{motor}", versaoMotor);
            itens.Add(new Item(c[0], V(c[1]), V(c[2])));
        }
        return itens;
    }
}
