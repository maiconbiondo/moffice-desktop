using System.Windows;
using System.Windows.Shell;

namespace mOffice.Comum;

/// <summary>Jump List do ícone do app na barra de tarefas e Recentes do Windows (Plano 6, Task 10 — SO-09). Os recentes continuam sendo o
/// <c>recentes-&lt;app&gt;.json</c> (fonte da tela inicial); o Windows ganha uma CÓPIA: cada arquivo registrado vai também para <c>SHAddToRecentDocs</c>
/// (<see cref="JumpList.AddToRecentCategory(string)"/>), o que o põe nos "Recentes" do Explorer e na categoria "Recentes" da Jump List — o Windows só
/// lista o que o app registrou em <c>Applications\&lt;exe&gt;\SupportedTypes</c> (o instalador faz, gerado da tabela de formatos).</summary>
public static class IntegracaoWindows
{
    public const string CategoriaFixados = "Fixados";

    /// <summary>A Jump List: categoria "Recentes" do Windows ligada, "Frequentes" desligada e uma categoria própria "Fixados" com os fixados, na ordem da lista
    /// (<paramref name="itens"/> já vem como <see cref="Recentes.Listar"/>: fixados primeiro). Não toca o Windows — só monta.</summary>
    /// <param name="exe">O executável do app (reservado: o ícone/caminho do app o Windows tira do próprio processo).</param>
    public static JumpList MontarJumpList(IReadOnlyList<ItemRecente> itens, string exe)
    {
        _ = exe;
        var jl = new JumpList { ShowRecentCategory = true, ShowFrequentCategory = false };
        foreach (var i in itens.Where(i => i.Fixado && EhLocalFixo(i.Caminho)))
            jl.JumpItems.Add(new JumpPath { Path = i.Caminho, CustomCategory = CategoriaFixados });
        return jl;
    }

    /// <summary>So' caminho em unidade LOCAL FIXA (C:\..): o <c>Apply</c> resolve cada item pelo shell na thread da tela, e um fixado em UNC/unidade de rede/removivel fora do
    /// ar travava a abertura do app pelo tempo do SMB (MEDIDO: 14 s com \\10.255.255.1\x). Esses fixados continuam na lista da tela inicial; so' ficam fora da Jump List.
    /// <c>DriveInfo.DriveType</c> nao toca a rede (GetDriveType).</summary>
    internal static bool EhLocalFixo(string caminho)
    {
        try
        {
            var raiz = Path.GetPathRoot(caminho);
            if (string.IsNullOrEmpty(raiz) || raiz.StartsWith(@"\\", StringComparison.Ordinal)) return false;
            return new DriveInfo(raiz).DriveType == DriveType.Fixed;
        }
        catch (Exception e) when (e is ArgumentException or IOException or NotSupportedException) { return false; }
    }

    /// <summary>Instala a Jump List no app. Melhor esforço: a barra de tarefas nunca derruba o app (ambiente sem shell, política que a desliga).</summary>
    public static void Aplicar(Application app, IReadOnlyList<ItemRecente> itens)
    {
        try
        {
            var jl = MontarJumpList(itens, Environment.ProcessPath ?? "");
            JumpList.SetJumpList(app, jl);
            jl.Apply();
        }
        catch (Exception e) when (e is InvalidOperationException or System.Runtime.InteropServices.COMException or IOException or UnauthorizedAccessException) { }
    }

    /// <summary>Registra o arquivo nos Recentes do Windows (e, por isso, na categoria "Recentes" da Jump List). Melhor esforço, nunca lança. Só chamar para um
    /// documento do usuário que foi aberto/salvo — nunca para modelo usado como novo nem para documento novo (Task 9).</summary>
    public static void AdicionarAosRecentesDoWindows(string caminho)
    {
        try { if (Path.IsPathRooted(caminho)) JumpList.AddToRecentCategory(caminho); }
        catch (Exception e) when (e is InvalidOperationException or System.Runtime.InteropServices.COMException or IOException or UnauthorizedAccessException or ArgumentException) { }
    }
}
