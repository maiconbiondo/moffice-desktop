using System.Runtime.CompilerServices;

namespace mOffice.Comum.Tests;

/// <summary>Força o construtor estático de System.Windows.Application (PresentationFramework) ANTES de
/// qualquer teste rodar. É esse construtor que registra o esquema de URI "pack" (UriParser.Register) e o
/// WebRequest.RegisterPrefix("pack", ...) que resolve `pack://application:,,,/...` pra um stream de
/// recurso — normalmente isso acontece de graça na primeira vez que uma janela XAML-backed chama
/// InitializeComponent (Application.LoadComponent toca essa infra). TemaTests/PhosphorTests nunca
/// constroem uma janela real (só ResourceDictionary.Source/GlyphTypeface direto), então sem isto o
/// primeiro pack:// desta assembly lança UriFormatException (esquema não registrado) e, corrigido isso,
/// NotSupportedException em WebRequest.Create (prefixo "pack" não registrado) — os dois vêm do mesmo
/// construtor faltando. ModuleInitializer roda uma vez, garantidamente antes de qualquer tipo desta
/// assembly ser usado, não importa a ordem dos testes. RunClassConstructor (em vez de `new Application()`)
/// evita o custo/risco de instanciar uma Application de verdade (só pode haver uma por processo, e o
/// ModuleInitializer não roda necessariamente numa thread STA).</summary>
internal static class InicializacaoModulo
{
    [ModuleInitializer]
    internal static void RegistrarEsquemaPack() =>
        System.Runtime.CompilerServices.RuntimeHelpers.RunClassConstructor(typeof(System.Windows.Application).TypeHandle);

    /// <summary>Mesma higiene da assembly do núcleo (ver <see cref="mOffice.Nucleo.Tests.DadosDeTeste"/>):
    /// `MOFFICE_DADOS_DIR` para `%LOCALAPPDATA%\mOffice-testes\dados`, para a suíte não escrever no
    /// `%LOCALAPPDATA%\mOffice` real do desenvolvedor. Precisa estar AQUI também: um `[ModuleInitializer]`
    /// roda por MÓDULO, e o do núcleo só dispararia quando o primeiro tipo daquela assembly fosse tocado —
    /// o que nesta assembly acontece em um teste só, tarde demais.</summary>
    [ModuleInitializer]
    internal static void RedirecionarDadosDoProduto() => mOffice.Nucleo.Tests.DadosDeTeste.Preparar();
}
