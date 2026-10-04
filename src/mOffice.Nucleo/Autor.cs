using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32;

namespace mOffice.Nucleo;

/// <summary>Autor de comentarios e alteracoes controladas (SO-03/CE-1/MvP-01, auditoria 2026-10-01): vira o
/// <c>editorConfig.user</c> pela URL de entrada (<c>userid</c>/<c>username</c>, lidos pelo index.html do motor). Antes era
/// "mOffice"/1 para todo mundo. Nada sai da maquina: so' APIs locais do Windows e o registro do usuario. O SID nunca e'
/// gravado nem logado: so' o hash (<see cref="IdDoSid"/>).</summary>
public sealed record Autor(string Id, string Nome)
{
    /// <summary>Ordem: nome configurado, "Nome de usuario" do Office desta conta, nome de exibicao do Windows,
    /// <see cref="Environment.UserName"/>.</summary>
    public static Autor DoWindows(string? nomeConfigurado = null) =>
        Limpo(nomeConfigurado) is { } nome ? new Autor(Padrao.Value.Id, nome) : Padrao.Value;

    // Calculado UMA vez por processo (registro, GetUserNameEx e SID): o nome do Office/Windows nao muda durante a sessao.
    private static readonly Lazy<Autor> Padrao = new(() =>
    {
        var sid = WindowsIdentity.GetCurrent().User?.Value ?? Environment.UserName;
        var nome = Limpo(NomeDoOffice()) ?? Limpo(NomeDeExibicao()) ?? Environment.UserName;
        return new Autor(IdDoSid(sid), nome);
    });

    /// <summary>"u" + 16 hex minusculos do SHA-256 do SID: estavel por pessoa, sem revelar o SID.</summary>
    public static string IdDoSid(string sid) =>
        "u" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sid)))[..16].ToLowerInvariant();

    /// <summary>"Nome de usuario" do Office (Arquivo &gt; Opcoes &gt; Geral) desta conta, se o Office ja' foi usado aqui.</summary>
    public static string? NomeDoOffice()
    {
        try { return Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo", "UserName", null) as string; }
        catch (Exception e) when (e is System.Security.SecurityException or IOException or UnauthorizedAccessException) { return null; }
    }

    private static string? Limpo(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    [DllImport("secur32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool GetUserNameEx(int formato, StringBuilder nome, ref uint tamanho);

    /// <summary>NameDisplay (3). Em conta local/Microsoft pode falhar (ERROR_NONE_MAPPED): devolve null.</summary>
    private static string? NomeDeExibicao()
    {
        try
        {
            uint n = 256;
            var sb = new StringBuilder((int)n);
            return GetUserNameEx(3, sb, ref n) ? sb.ToString() : null;
        }
        catch (Exception e) when (e is DllNotFoundException or EntryPointNotFoundException) { return null; }
    }
}
