namespace mOffice.Nucleo.Ortografia;

/// <summary>Palavras que o usuario mandou "Adicionar ao dicionario": uma por linha, UTF-8, em
/// %LOCALAPPDATA%\mOffice\dicionario-usuario.txt (<see cref="Pastas.DicionarioUsuario"/>, spec §8).
/// Compartilhado pelos tres apps. Lido do disco uma vez por instancia (a primeira consulta) e depois so'
/// apendado — uma palavra adicionada por OUTRO processo aparece na proxima vez que o app abrir.</summary>
public sealed class DicionarioUsuario
{
    private readonly string _arquivo;
    private readonly object _trava = new();
    private HashSet<string>? _palavras;

    public DicionarioUsuario(string arquivo) => _arquivo = arquivo;

    private HashSet<string> Palavras()
    {
        if (_palavras is not null) return _palavras;
        _palavras = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            if (File.Exists(_arquivo))
                foreach (var l in File.ReadAllLines(_arquivo)) if (l.Length > 0) _palavras.Add(l);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        return _palavras;
    }

    public bool Contem(string palavra) { lock (_trava) return Palavras().Contains(palavra); }

    public void Adicionar(string palavra)
    {
        // Uma palavra por linha: quebra de linha dentro dela corromperia o arquivo (e nao e' palavra).
        if (string.IsNullOrWhiteSpace(palavra) || palavra.Contains('\n') || palavra.Contains('\r')) return;
        lock (_trava)
        {
            if (!Palavras().Add(palavra)) return;
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_arquivo)!);
                File.AppendAllText(_arquivo, palavra + Environment.NewLine);
            }
            // Disco cheio/sem permissao: a palavra vale nesta sessao (ja' esta' no conjunto), so' nao persiste.
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        }
    }

    public void Limpar()
    {
        lock (_trava)
        {
            Palavras().Clear();
            try { if (File.Exists(_arquivo)) File.Delete(_arquivo); }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException) { }
        }
    }
}
