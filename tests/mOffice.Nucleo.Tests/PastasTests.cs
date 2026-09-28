using mOffice.Nucleo;

namespace mOffice.Nucleo.Tests;

public class PastasTests
{
    [Fact] // sem a variavel de ambiente, a raiz e' %LOCALAPPDATA%\mOffice
    public void Raiz_SemVariavel_EhLocalAppDataMOffice()
    {
        var esperado = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "mOffice");
        Assert.Equal(esperado, Pastas.RaizPadrao());
    }

    [Fact] // MOFFICE_DADOS_DIR nao vazia redireciona TUDO (recuperacao, logs, config, recentes)
    public void Raiz_ComVariavel_RedirecionaTodasAsSubpastas()
    {
        var temp = Path.Combine(Path.GetTempPath(), "mOffice-testes", Guid.NewGuid().ToString("N"));
        var raiz = Pastas.Resolver(temp);
        Assert.Equal(temp, raiz);
        Assert.Equal(Path.Combine(temp, "recuperacao"), Pastas.RecuperacaoEm(temp));
        Assert.Equal(Path.Combine(temp, "logs"), Pastas.LogsEm(temp));
        Assert.Equal(Path.Combine(temp, "config.json"), Pastas.ConfigEm(temp));
        Assert.Equal(Path.Combine(temp, "recentes-mtexto.json"), Pastas.RecentesEm(temp, "mtexto"));
    }

    [Fact] // variavel em branco ou so' espacos = ignorada (nunca uma raiz vazia)
    public void Resolver_VariavelEmBranco_UsaPadrao()
    {
        Assert.Equal(Pastas.RaizPadrao(), Pastas.Resolver("   "));
        Assert.Equal(Pastas.RaizPadrao(), Pastas.Resolver(null));
    }

    [Fact] // a raiz de recovery do nucleo e' a pasta por usuario, nao %TEMP%
    public void SessaoDocumento_RaizRecovery_EstaSobPastas()
    {
        Assert.Equal(Pastas.Recuperacao, SessaoDocumento.RaizRecovery);
        // Duas afirmacoes, de proposito (revisao 2). A do caminho EFETIVO so' vale enquanto o arranjo de
        // teste mantiver MOFFICE_DADOS_DIR fora de %TEMP% (ver DadosDeTeste.PastaDados, que explica por que);
        // a da raiz PADRAO e' a que fala do PRODUTO — "nada em %TEMP%" da spec paragrafo 4 e' uma promessa
        // sobre onde o app INSTALADO grava, e essa nenhuma variavel de ambiente de teste pode afrouxar.
        Assert.DoesNotContain(Path.GetTempPath().TrimEnd('\\'), SessaoDocumento.RaizRecovery, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(Path.GetTempPath().TrimEnd('\\'), Pastas.RecuperacaoEm(Pastas.RaizPadrao()), StringComparison.OrdinalIgnoreCase);
    }
}
