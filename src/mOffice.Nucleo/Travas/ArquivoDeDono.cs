using mOffice.Conversor;

namespace mOffice.Nucleo.Travas;

/// <summary>Qual Office o arquivo de dono imita: o Word grava 162 bytes; Excel e PowerPoint, 165 (e o PowerPoint um 0x00 a mais).</summary>
public enum AppDoDono { Word, Excel, PowerPoint }

/// <summary>
/// Arquivo de dono <c>~$&lt;nome&gt;</c> ao lado do documento aberto (Plano 6, Task 5 — SO-02): e' dele que o Word/Excel/PowerPoint
/// dos colegas tiram o "em uso por &lt;nome&gt;". FORMATO (so' o formato; nada de codigo) o documentado pela implementacao compativel do
/// LibreOffice, <c>svl/source/misc/msodocumentlockfile.cxx</c> (MPL), conferido em 2026-10-02:
/// <list type="bullet">
/// <item>Word (162 bytes): [0] = N (tamanho do nome, ate' 52); [1..N] = nome em 1 byte por caractere; 0x00 ate' a posicao 54;
/// [54] = N, [55] = 0; nome em UTF-16LE; 0x00 ate' o fim.</item>
/// <item>Excel (165): [0] = N; nome em 1 byte; 0x20 ate' a posicao 55; [55] = N, [56] = 0; UTF-16LE; "20 00" ate' o fim.</item>
/// <item>PowerPoint (165): como o Excel, com um 0x00 logo depois do nome em 1 byte.</item>
/// </list>
/// O nome do arquivo (<see cref="CaminhoPara"/>) tambem vem de la': so' nos documentos de TEXTO o Word troca as primeiras letras.
/// A comparacao com um <c>~$</c> gravado pelo Office de verdade e' passo HUMANO (relatorio da Task 5).
/// </summary>
public sealed class ArquivoDeDono : IDisposable
{
    private const int MaxNome = 52;
    private const int TamanhoWord = 162;
    private const int TamanhoExcelPowerPoint = 165;

    private FileStream? _arquivo;

    private ArquivoDeDono(string caminho, FileStream arquivo) { Caminho = caminho; _arquivo = arquivo; }

    /// <summary>Onde o arquivo de dono deste documento fica.</summary>
    public string Caminho { get; }

    /// <summary>Regra do Office: documento de texto com 8 letras ou mais antes da extensao troca as 2 primeiras por "~$"; com 7, a
    /// primeira; com menos, so' prefixa. Planilha e apresentacao: sempre "~$" + o nome inteiro.</summary>
    public static string CaminhoPara(string documento)
    {
        var nome = Path.GetFileName(documento);
        var ext = Path.GetExtension(nome);
        if (App(documento) == AppDoDono.Word)
        {
            var semExt = nome.Length - ext.Length;
            if (semExt >= 8) nome = nome[2..];
            else if (semExt == 7) nome = nome[1..];
        }
        return Path.Combine(Path.GetDirectoryName(documento) ?? "", "~$" + nome);
    }

    /// <summary>O Office do tipo do documento (pela tabela de formatos): planilha = Excel, apresentacao = PowerPoint, o resto = Word.</summary>
    public static AppDoDono App(string documento) => TabelaFormatos.DaExtensao(documento)?.DocType switch
    {
        "cell" => AppDoDono.Excel,
        "slide" => AppDoDono.PowerPoint,
        _ => AppDoDono.Word,
    };

    /// <summary>Os bytes do arquivo de dono com <paramref name="nome"/> (cortado em 52 caracteres, como o Office). No campo de 1 byte,
    /// caractere fora do Latin-1 vira '?'; o UTF-16 leva o nome certo.</summary>
    public static byte[] Montar(string nome, AppDoDono app)
    {
        if (nome.Length > MaxNome) nome = nome[..MaxNome];
        var tamanho = app == AppDoDono.Word ? TamanhoWord : TamanhoExcelPowerPoint;
        var b = new byte[tamanho];
        var i = 0;
        b[i++] = (byte)nome.Length;
        foreach (var c in nome) b[i++] = c <= 0xFF ? (byte)c : (byte)'?';
        switch (app)
        {
            case AppDoDono.Word:
                while (i < MaxNome + 2) b[i++] = 0;
                break;
            case AppDoDono.PowerPoint:
                b[i++] = 0;
                while (i < MaxNome + 3) b[i++] = 0x20;
                break;
            default:
                while (i < MaxNome + 3) b[i++] = 0x20;
                break;
        }
        b[i++] = (byte)nome.Length;
        b[i++] = 0;
        foreach (var c in nome) { b[i++] = (byte)(c & 0xFF); b[i++] = (byte)(c >> 8); }
        if (app == AppDoDono.Word) while (i < tamanho) b[i++] = 0;
        else
            while (i < tamanho)
            {
                b[i++] = 0x20;
                if (i < tamanho) b[i++] = 0;
            }
        return b;
    }

    /// <summary>O nome de quem esta' com o documento, lido de um <c>~$</c> (do Office ou nosso). <c>null</c> se nao existe, nao da'
    /// para ler (lido com <c>FileShare.ReadWrite | Delete</c>: o dono pode estar com ele aberto) ou nao tem o formato.</summary>
    public static string? LerNome(string caminhoDono)
    {
        byte[] b;
        try
        {
            using var f = new FileStream(caminhoDono, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            b = new byte[Math.Min(f.Length, 512)];
            f.ReadExactly(b);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) { return null; }
        if (b.Length < 2) return null;
        int n = b[0];
        if (n is < 1 or > MaxNome) return null;
        // UTF-16: Word em 56 ([54] = N, [55] = 0); Excel/PowerPoint em 57 ([55] = N, [56] = 0).
        foreach (var inicio in new[] { 56, 57 })
            if (b.Length >= inicio + 2 * n && b[inicio - 2] == n && b[inicio - 1] == 0)
                return System.Text.Encoding.Unicode.GetString(b, inicio, 2 * n);
        // Sem o UTF-16 (arquivo cortado): o campo de 1 byte.
        return b.Length >= 1 + n ? System.Text.Encoding.Latin1.GetString(b, 1, n) : null;
    }

    /// <summary>Cria (ou toma de uma queda) o <c>~$</c> de <paramref name="documento"/>, oculto, com <paramref name="nome"/>, e o mantem
    /// aberto (<c>FileShare.Read | Delete</c>: o Office dos colegas le' o nome). Um <c>~$</c> que ja' exista sem ninguem segurando (sobra
    /// de queda, nossa ou de outra pessoa) e' substituido; um que outro processo segura fica como esta' — devolve <c>null</c>.</summary>
    public static ArquivoDeDono? Criar(string documento, string nome)
    {
        var caminho = CaminhoPara(documento);
        var bytes = Montar(nome, App(documento));
        FileStream? f = null;
        try
        {
            // ABRE primeiro (OpenOrCreate: aceita um ~$ oculto, ao contrario do CREATE_ALWAYS) e so' mexe no conteudo e nos atributos do
            // que conseguiu abrir — um ~$ preso por outra pessoa nunca e' desocultado nem truncado (fix 1 da revisao da Task 5, MINOR 4).
            f = new FileStream(caminho, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read | FileShare.Delete);
            f.SetLength(0);
            f.Write(bytes);
            f.Flush();
            File.SetAttributes(caminho, FileAttributes.Hidden);
            return new ArquivoDeDono(caminho, f);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            f?.Dispose();
            return null;
        }
    }

    /// <summary>Fecha e apaga (melhor esforco: um apagar que falhe deixa uma sobra que a proxima abertura substitui).</summary>
    public void Dispose()
    {
        if (_arquivo is null) return;
        _arquivo.Dispose();
        _arquivo = null;
        try
        {
            File.SetAttributes(Caminho, FileAttributes.Normal);
            File.Delete(Caminho);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
    }
}
