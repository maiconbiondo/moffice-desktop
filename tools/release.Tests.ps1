<#
.SYNOPSIS
    Testes (Pester) das funcoes de tools/release-functions.ps1 e do contrato de tools/release.ps1.
    NUNCA dot-source-a tools/release.ps1 - ele tem o fluxo que de fato publica uma release.

.DESCRIPTION
    Roda com:  Invoke-Pester -Script tools\release.Tests.ps1
    (Pester 3.4.0, o que vem com o Windows - sintaxe "Should Be", como no mPDF.)

    Nenhum teste chama o gh de verdade nem usa rede:
      - o gh e um FALSO (.cmd em %TEMP%) que escreve no stderr, grava os argumentos recebidos e sai com o
        codigo pedido - reproduz o caso do PowerShell 5.1 (stderr de comando nativo + 'Stop' = erro
        terminante) sem depender de token nem da API do GitHub;
      - o git e um repositorio descartavel em %TEMP% (commit com data fixa e tag v9.9.9), apagado no fim.

    ASCII puro e sem BOM, como os demais scripts de tools/.
#>

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $PSScriptRoot 'release-functions.ps1')

$Rascunho = Join-Path ([System.IO.Path]::GetTempPath()) "moffice-release-tests-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $Rascunho -Force | Out-Null

function New-GhFalso {
    # Cria um gh.cmd falso que grava os argumentos em <nome>.args, escreve no stderr e sai com $Codigo.
    param([string]$Nome, [int]$Codigo)
    $cmd = Join-Path $Rascunho "$Nome.cmd"
    $log = Join-Path $Rascunho "$Nome.args"
    $conteudo = "@echo off`r`n>""$log"" echo %*`r`necho gh falso: erro simulado 1>&2`r`nexit /b $Codigo`r`n"
    [System.IO.File]::WriteAllText($cmd, $conteudo, (New-Object System.Text.ASCIIEncoding))
    return @{ Exe = $cmd; Log = $log }
}

# Repositorio git descartavel: um commit com data de commit FIXA e a tag v9.9.9 nele.
$RepoFalso = Join-Path $Rascunho 'repo'
$DataCommitFixa = [DateTimeOffset]::FromUnixTimeSeconds(1780000000).UtcDateTime
New-Item -ItemType Directory -Path $RepoFalso -Force | Out-Null
$envAntes = @{ A = $env:GIT_AUTHOR_DATE; C = $env:GIT_COMMITTER_DATE }
try {
    $env:GIT_AUTHOR_DATE = '1780000000 +0000'
    $env:GIT_COMMITTER_DATE = '1780000000 +0000'
    & git -C $RepoFalso init -q 2>$null | Out-Null
    Set-Content -LiteralPath (Join-Path $RepoFalso 'a.txt') -Value 'a' -Encoding Ascii
    & git -C $RepoFalso add a.txt | Out-Null
    & git -C $RepoFalso -c user.name=teste -c user.email=teste@example.invalid -c commit.gpgsign=false commit -q -m inicial | Out-Null
    & git -C $RepoFalso -c tag.gpgsign=false tag v9.9.9 | Out-Null
} finally {
    $env:GIT_AUTHOR_DATE = $envAntes.A
    $env:GIT_COMMITTER_DATE = $envAntes.C
}

try {

Describe 'Get-ProductVersion' {
    It 'le a <Version> de um Directory.Build.props com <Import> antes do PropertyGroup' {
        $tmp = Join-Path $Rascunho 'Directory.Build.props'
        Set-Content -LiteralPath $tmp -Encoding Ascii -Value @'
<Project>
  <Import Project="..\outro.props" />
  <PropertyGroup>
    <Version>9.9.9</Version>
  </PropertyGroup>
</Project>
'@
        (Get-ProductVersion -PropsPath $tmp) | Should Be '9.9.9'
    }

    It 'recusa versao com sufixo (System.Version do atualizador nao aceita -rc)' {
        $tmp = Join-Path $Rascunho 'rc.props'
        Set-Content -LiteralPath $tmp -Encoding Ascii -Value '<Project><PropertyGroup><Version>1.0.0-rc1</Version></PropertyGroup></Project>'
        { Get-ProductVersion -PropsPath $tmp } | Should Throw
    }

    It 'recusa props sem <Version>' {
        $tmp = Join-Path $Rascunho 'sem.props'
        Set-Content -LiteralPath $tmp -Encoding Ascii -Value '<Project><PropertyGroup><Outra>1</Outra></PropertyGroup></Project>'
        { Get-ProductVersion -PropsPath $tmp } | Should Throw
    }

    It 'lanca erro se o arquivo nao existe' {
        { Get-ProductVersion -PropsPath (Join-Path $Rascunho 'nao-existe.props') } | Should Throw
    }

    It 'reflete a versao ATUAL do produto (src\Directory.Build.props real, sem hardcode)' {
        $props = Join-Path $RepoRoot 'src\Directory.Build.props'
        # Caminho INDEPENDENTE (regex no texto cru) contra o parse [xml] - nao apodrece a cada release.
        $raw = [regex]::Match((Get-Content -LiteralPath $props -Raw), '<Version>\s*([^<\s]+)\s*</Version>').Groups[1].Value
        $raw | Should Match '^\d+\.\d+\.\d+$'
        (Get-ProductVersion -PropsPath $props) | Should Be $raw
    }
}

Describe 'Test-GhAuthenticated (gh falso, sem rede)' {
    # Regressao do mPDF: `gh auth status 2>$null` sob 'Stop' lancava NativeCommandError TERMINANTE com
    # token invalido. Aqui o gh falso escreve no stderr e sai com 1 - mesmo gatilho, sem token nem API.
    It 'devolve $false sem lancar quando o gh sai com erro e escreve no stderr (sob Stop)' {
        $gh = New-GhFalso -Nome 'gh-auth-falha' -Codigo 1
        $ErrorActionPreference = 'Stop'
        { Test-GhAuthenticated -GhExe $gh.Exe } | Should Not Throw
        (Test-GhAuthenticated -GhExe $gh.Exe) | Should Be $false
    }

    It 'devolve $true quando o gh sai com 0, e pergunta "auth status"' {
        $gh = New-GhFalso -Nome 'gh-auth-ok' -Codigo 0
        (Test-GhAuthenticated -GhExe $gh.Exe) | Should Be $true
        (Get-Content -LiteralPath $gh.Log -Raw).Trim() | Should Be 'auth status'
    }

    It 'lanca (mensagem clara) quando o gh nao existe' {
        { Test-GhAuthenticated -GhExe (Join-Path $Rascunho 'nao-existe-gh.exe') } | Should Throw 'gh CLI nao encontrado'
    }
}

Describe 'Test-ReleaseExists (gh falso, sem rede)' {
    It 'devolve $false sem lancar quando a release nao existe (gh sai 1 com stderr, sob Stop)' {
        $gh = New-GhFalso -Nome 'gh-view-404' -Codigo 1
        $ErrorActionPreference = 'Stop'
        { Test-ReleaseExists -Version '9.9.9' -Repo 'dono/repo' -GhExe $gh.Exe } | Should Not Throw
        (Test-ReleaseExists -Version '9.9.9' -Repo 'dono/repo' -GhExe $gh.Exe) | Should Be $false
    }

    It 'devolve $true quando a release existe (idempotencia: o release.ps1 aborta) e consulta a tag v<versao> no repo certo' {
        $gh = New-GhFalso -Nome 'gh-view-ok' -Codigo 0
        (Test-ReleaseExists -Version '9.9.9' -Repo 'dono/repo' -GhExe $gh.Exe) | Should Be $true
        (Get-Content -LiteralPath $gh.Log -Raw).Trim() | Should Be 'release view v9.9.9 --repo dono/repo'
    }
}

Describe 'Test-WorkingTreeClean (repositorio descartavel)' {
    It 'devolve $true com a arvore limpa' {
        (Test-WorkingTreeClean -RepoRoot $RepoFalso) | Should Be $true
    }

    It 'devolve $false com arquivo NAO rastreado e lista o arquivo' {
        $solto = Join-Path $RepoFalso 'solto.exe'
        Set-Content -LiteralPath $solto -Value 'x' -Encoding Ascii
        try {
            $sujos = @()
            (Test-WorkingTreeClean -RepoRoot $RepoFalso -Sujos ([ref]$sujos)) | Should Be $false
            ($sujos -join '|') | Should Match 'solto\.exe'
        } finally {
            Remove-Item -LiteralPath $solto -Force
        }
    }

    It 'devolve $false com arquivo rastreado modificado' {
        $a = Join-Path $RepoFalso 'a.txt'
        Set-Content -LiteralPath $a -Value 'mudou' -Encoding Ascii
        try {
            (Test-WorkingTreeClean -RepoRoot $RepoFalso) | Should Be $false
        } finally {
            & git -C $RepoFalso checkout -q -- a.txt
        }
        (Test-WorkingTreeClean -RepoRoot $RepoFalso) | Should Be $true
    }
}

Describe 'Test-VersionTagExists (repositorio descartavel)' {
    It 'devolve $true para a tag que existe (v9.9.9)' {
        (Test-VersionTagExists -Version '9.9.9' -RepoRoot $RepoFalso) | Should Be $true
    }

    It 'devolve $false para versao sem tag, inclusive prefixo parecido (9.9)' {
        (Test-VersionTagExists -Version '999.999.999' -RepoRoot $RepoFalso) | Should Be $false
        (Test-VersionTagExists -Version '9.9' -RepoRoot $RepoFalso) | Should Be $false
    }
}

Describe 'Get-TagCommitDate (repositorio descartavel)' {
    It 'le a data de commit FIXA da tag v9.9.9' {
        (Get-TagCommitDate -Version '9.9.9' -RepoRoot $RepoFalso) | Should Be $DataCommitFixa
    }

    It 'lanca erro para tag que nao existe (mesmo sob Stop, com mensagem pt-BR)' {
        $ErrorActionPreference = 'Stop'
        { Get-TagCommitDate -Version '999.999.999' -RepoRoot $RepoFalso } | Should Throw 'git log falhou'
    }
}

Describe 'Test-InstallerFresh (cinto de frescor)' {
    It 'devolve $true quando o exe e MAIS NOVO que a tag' {
        (Test-InstallerFresh -ExeLastWriteTimeUtc (Get-Date '2026-01-02T00:00:00Z') -TagCommitDateUtc (Get-Date '2026-01-01T00:00:00Z')) | Should Be $true
    }

    It 'devolve $false quando o exe e MAIS ANTIGO que a tag' {
        (Test-InstallerFresh -ExeLastWriteTimeUtc (Get-Date '2026-01-01T00:00:00Z') -TagCommitDateUtc (Get-Date '2026-01-02T00:00:00Z')) | Should Be $false
    }

    It 'devolve $true quando os dois timestamps sao IGUAIS (>=, nao >)' {
        $mesmo = Get-Date '2026-01-01T12:00:00Z'
        (Test-InstallerFresh -ExeLastWriteTimeUtc $mesmo -TagCommitDateUtc $mesmo) | Should Be $true
    }

    It 'com a data real da tag do repositorio descartavel: exe escrito antes recusa, escrito depois passa' {
        $tag = Get-TagCommitDate -Version '9.9.9' -RepoRoot $RepoFalso
        $exe = Join-Path $Rascunho 'mOffice-Setup-9.9.9.exe'
        Set-Content -LiteralPath $exe -Value 'x' -Encoding Ascii
        (Get-Item -LiteralPath $exe).LastWriteTimeUtc = $tag.AddSeconds(-1)
        (Test-InstallerFresh -ExeLastWriteTimeUtc (Get-Item -LiteralPath $exe).LastWriteTimeUtc -TagCommitDateUtc $tag) | Should Be $false
        (Get-Item -LiteralPath $exe).LastWriteTimeUtc = $tag.AddSeconds(1)
        (Test-InstallerFresh -ExeLastWriteTimeUtc (Get-Item -LiteralPath $exe).LastWriteTimeUtc -TagCommitDateUtc $tag) | Should Be $true
    }
}

Describe 'Test-AssetSize (limite de 2 GiB do GitHub)' {
    It 'aceita o instalador medido na Task 3 (464.352.457 bytes)' {
        (Test-AssetSize -Bytes 464352457) | Should Be $true
    }

    It 'aceita 1 byte abaixo de 2 GiB e recusa exatamente 2 GiB e acima' {
        (Test-AssetSize -Bytes (2GB - 1)) | Should Be $true
        (Test-AssetSize -Bytes 2GB) | Should Be $false
        (Test-AssetSize -Bytes (3GB)) | Should Be $false
    }

    It 'recusa arquivo vazio' {
        (Test-AssetSize -Bytes 0) | Should Be $false
    }
}

Describe 'Get-InstallerExePath' {
    It 'monta o caminho esperado a partir de versao + pasta' {
        (Get-InstallerExePath -Version '1.0.0' -OutputDir 'C:\saida') | Should Be 'C:\saida\mOffice-Setup-1.0.0.exe'
    }

    It 'o nome bate com o OutputBaseFilename do moffice.iss' {
        $iss = Get-Content -LiteralPath (Join-Path $RepoRoot 'tools\installer\moffice.iss') -Raw
        $iss | Should Match '(?m)^OutputBaseFilename=mOffice-Setup-\{#MyAppVersion\}\s*$'
        (Split-Path -Leaf (Get-InstallerExePath -Version '1.0.0' -OutputDir 'C:\x')) | Should Be ('mOffice-Setup-{#MyAppVersion}.exe' -replace '\{#MyAppVersion\}', '1.0.0')
    }
}

Describe 'Format-ReleaseNotes (contrato do corpo)' {
    It 'com Notas: notas + linha em branco + SHA256 como ultima linha' {
        (Format-ReleaseNotes -Notas 'Correcao: X.' -Sha256 'abc123') | Should Be "Correcao: X.`n`nSHA256: abc123"
    }

    It 'sem Notas: corpo contem SOMENTE a linha SHA256' {
        (Format-ReleaseNotes -Notas '' -Sha256 'abc123') | Should Be 'SHA256: abc123'
        (Format-ReleaseNotes -Sha256 'abc123') | Should Be 'SHA256: abc123'
    }

    It 'apara espacos em volta das notas e normaliza CRLF para LF' {
        (Format-ReleaseNotes -Notas "  Linha 1`r`nLinha 2  `r`n" -Sha256 'abc123') | Should Be "Linha 1`nLinha 2`n`nSHA256: abc123"
    }

    It 'a ultima linha bate com a regex do ServicoAtualizacao (lida do .cs, nao copiada)' {
        $cs = Get-Content -LiteralPath (Join-Path $RepoRoot 'src\mOffice.Comum\Atualizacao\ServicoAtualizacao.cs') -Raw
        $regexDoApp = [regex]::Match($cs, 'LinhaSha\s*=\s*new\(@"([^"]+)"').Groups[1].Value
        $regexDoApp | Should Be '^SHA256:\s*([0-9a-fA-F]{64})$'
        $hash = ('a' * 64)
        $ultimaLinha = ((Format-ReleaseNotes -Notas 'Nota qualquer.' -Sha256 $hash) -split "`n")[-1]
        $ultimaLinha | Should Match $regexDoApp
    }

    It 'REGRA DO LEITOR: um "SHA256: ..." decoy nas notas nao e a ultima linha que casa' {
        $hashReal = ('b' * 64)
        $decoy = "SHA256: $('c' * 64)"
        $body = Format-ReleaseNotes -Notas "Aviso.`n$decoy`nCorrecao: X." -Sha256 $hashReal
        $linhas = $body -split "`n"
        $casamentos = @($linhas | Where-Object { $_ -match '^SHA256:\s*([0-9a-fA-F]{64})$' })
        $casamentos.Count | Should Be 2
        $casamentos[-1] | Should Be "SHA256: $hashReal"
        $linhas[-1] | Should Be $casamentos[-1]
    }
}

Describe 'Write-ReleaseNotesFile (arquivo do --notes-file)' {
    It 'grava UTF-8 SEM BOM, byte a byte, com acento e aspas preservados' {
        $arq = Join-Path $Rascunho 'notas.md'
        $texto = "Corre$([char]0xE7)$([char]0xE3)o: ""aspas"" (parenteses).`n`nSHA256: $('d' * 64)"
        Write-ReleaseNotesFile -Conteudo $texto -Caminho $arq
        $bytes = [System.IO.File]::ReadAllBytes($arq)
        ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) | Should Be $false
        [System.Text.Encoding]::UTF8.GetString($bytes) | Should BeExactly $texto
    }
}

Describe 'Get-ReleaseCreateArgs' {
    $a = Get-ReleaseCreateArgs -Version '1.0.0' -ExePath 'C:\o\mOffice-Setup-1.0.0.exe' -Repo 'maiconbiondo/moffice-desktop' -NotesFile 'C:\t\n.md'

    It 'monta o gh release create com tag, asset, repo e titulo "mOffice <versao>"' {
        ($a -join ' ') | Should Be 'release create v1.0.0 C:\o\mOffice-Setup-1.0.0.exe --repo maiconbiondo/moffice-desktop --title mOffice 1.0.0 --notes-file C:\t\n.md'
        $a[7] | Should Be 'mOffice 1.0.0'
    }

    It 'corpo SO por --notes-file (nunca --notes "texto")' {
        ($a -contains '--notes') | Should Be $false
        $a[[array]::IndexOf($a, '--notes-file') + 1] | Should Be 'C:\t\n.md'
    }
}

Describe 'release.ps1 (estrutura, sem executar)' {
    $src = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'release.ps1') -Raw
    # So o CODIGO (depois do cabecalho de ajuda <# ... #>): o cabecalho cita os mesmos nomes na prosa.
    $src = $src.Substring($src.IndexOf('#>') + 2)

    It 'repositorio padrao e o mesmo que o app consulta (ServicoAtualizacao.Repositorio)' {
        $cs = Get-Content -LiteralPath (Join-Path $RepoRoot 'src\mOffice.Comum\Atualizacao\ServicoAtualizacao.cs') -Raw
        $doApp = [regex]::Match($cs, 'const string Repositorio\s*=\s*"([^"]+)"').Groups[1].Value
        $doApp | Should Be 'maiconbiondo/moffice-desktop'
        $src | Should Match ([regex]::Escape("[string]`$Repo = '$doApp'"))
    }

    It 'le a versao de src\Directory.Build.props' {
        $src | Should Match ([regex]::Escape("'src\Directory.Build.props'"))
    }

    It 'checa idempotencia e tamanho ANTES do gh release create, e usa a funcao do --notes-file' {
        $iCreate = $src.IndexOf('& gh @ghArgs')
        $iCreate | Should BeGreaterThan 0
        $src.IndexOf('Test-ReleaseExists -Version') | Should BeGreaterThan 0
        ($src.IndexOf('Test-ReleaseExists -Version') -lt $iCreate) | Should Be $true
        ($src.IndexOf('Test-AssetSize -Bytes') -lt $iCreate) | Should Be $true
        ($src.IndexOf('Get-ReleaseCreateArgs') -lt $iCreate) | Should Be $true
        $src | Should Not Match "'--notes'"
    }

    It 'ordem das precondicoes: gh auth, arvore limpa, versao, tag, instalador, hash, idempotencia' {
        $ordem = @('Test-GhAuthenticated', 'Test-WorkingTreeClean -RepoRoot', 'Get-ProductVersion -PropsPath', 'Test-VersionTagExists -Version',
                   'Test-InstallerFresh -ExeLastWriteTimeUtc', 'Get-FileHash', 'Test-ReleaseExists -Version', '& gh @ghArgs')
        $pos = @($ordem | ForEach-Object { $src.IndexOf($_) })
        ($pos | Where-Object { $_ -lt 0 }).Count | Should Be 0
        for ($i = 1; $i -lt $pos.Count; $i++) { ($pos[$i] -gt $pos[$i - 1]) | Should Be $true }
    }

    It 'scripts de release em ASCII puro e sem BOM' {
        foreach ($f in @('release.ps1', 'release-functions.ps1', 'release.Tests.ps1')) {
            $bytes = [System.IO.File]::ReadAllBytes((Join-Path $PSScriptRoot $f))
            @($bytes | Where-Object { $_ -gt 0x7F }).Count | Should Be 0
        }
    }
}

} finally {
    Remove-Item -LiteralPath $Rascunho -Recurse -Force -ErrorAction SilentlyContinue
}
