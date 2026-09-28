<#
.SYNOPSIS
    Funcoes usadas por tools/release.ps1, isoladas neste arquivo SEM efeito colateral de publicacao
    (nenhuma delas cria nem publica release) para que tools/release.Tests.ps1 possa dot-source-ar so isto.
    Dot-source-ar tools/release.ps1 rodaria o fluxo de publicacao de verdade.

.DESCRIPTION
    Marco 1 / Plano 5 / Task 6. Porte do tools/release-functions.ps1 do mPDF, com o mesmo contrato. Cada
    funcao aqui e leitura (arquivo local, `git status`, `git tag -l`, `git log`, `gh auth status`,
    `gh release view`), montagem pura de string/lista ou a gravacao do arquivo TEMPORARIO do corpo da
    release. O fluxo que publica fica exclusivamente em tools/release.ps1.

    Diferencas do exemplar (mPDF), todas para testar sem rede e sem o gh de verdade:
      - as funcoes que chamam o gh recebem -GhExe (padrao 'gh'); os testes apontam para um gh FALSO
        (um .cmd em %TEMP% que escreve no stderr e sai com o codigo pedido). O mPDF batia na API real.
      - as funcoes de git recebem -RepoRoot; os testes usam um repositorio git descartavel em %TEMP%
        (o mOffice ainda nao tem nenhuma tag de versao).
      - Test-WorkingTreeClean, Test-AssetSize, Get-ReleaseCreateArgs e Write-ReleaseNotesFile sairam do
        corpo do release.ps1 para ca, para o contrato (arvore limpa, limite de 2 GiB do GitHub, corpo por
        --notes-file, UTF-8 sem BOM) ter teste.

    ASCII puro e sem BOM, como os demais scripts de tools/.
#>

# Limite de um asset de release do GitHub: 2 GiB (o mesmo do build-installer.ps1). O instalador medido na
# Task 3 tem 464.352.457 bytes (~443 MB), ~22% do limite.
$script:LimiteAssetGitHub = 2GB

function Get-ProductVersion {
    <# Le a <Version> de src/Directory.Build.props (fonte UNICA do produto, a mesma do
       build-installer.ps1 e do export-public.ps1). #>
    param([Parameter(Mandatory)][string]$PropsPath)
    if (-not (Test-Path -LiteralPath $PropsPath)) {
        throw "Directory.Build.props nao encontrado: $PropsPath"
    }
    [xml]$xml = Get-Content -LiteralPath $PropsPath -Raw
    $version = $xml.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Nao foi possivel extrair <Version> de $PropsPath."
    }
    $version = $version.Trim()
    if ($version -notmatch '^\d+\.\d+\.\d+$') {
        throw "<Version> '$version' em $PropsPath nao tem o formato X.Y.Z (System.Version do atualizador; sem sufixo -rc)."
    }
    return $version
}

function Invoke-GhSemLancar {
    <# Chama o gh e devolve SO o codigo de saida. No Windows PowerShell 5.1, redirecionar o stderr de um
       comando nativo (mesmo para $null) embrulha cada linha numa ErrorRecord "NativeCommandError"; com
       $ErrorActionPreference = 'Stop' (o do release.ps1) isso vira erro TERMINANTE. "Nao autenticado" e
       "release nao existe" sao resultados NORMAIS (codigo != 0), nunca excecao crua do gh.exe - por isso
       'Continue' so nesta chamada, restaurado no finally. Mesmo padrao do mPDF (Test-GhAuthenticated e
       Test-ReleaseExists), fatorado. Lanca so se o gh nao puder ser chamado. #>
    param(
        [Parameter(Mandatory)][string]$GhExe,
        [Parameter(Mandatory)][string[]]$Argumentos
    )
    if (-not (Get-Command $GhExe -ErrorAction SilentlyContinue)) {
        throw "gh CLI nao encontrado ($GhExe). Instale com: winget install GitHub.cli"
    }
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $null = & $GhExe @Argumentos 2>$null
        $codigo = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevEap
    }
    return $codigo
}

function Test-GhAuthenticated {
    <# Confere via `gh auth status` se o gh CLI esta autenticado. Nao lanca quando NAO autenticado - so
       devolve $false. #>
    param([string]$GhExe = 'gh')
    return ((Invoke-GhSemLancar -GhExe $GhExe -Argumentos @('auth', 'status')) -eq 0)
}

function Test-ReleaseExists {
    <# IDEMPOTENCIA: confere via `gh release view v<Version> --repo <Repo>` se a release ja existe. "Nao
       existe" e o caminho ESPERADO (gh sai != 0 e escreve no stderr) e devolve $false, sem lancar. #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$Repo,
        [string]$GhExe = 'gh'
    )
    return ((Invoke-GhSemLancar -GhExe $GhExe -Argumentos @('release', 'view', "v$Version", '--repo', $Repo)) -eq 0)
}

function Test-WorkingTreeClean {
    <# $true se `git status --porcelain` nao lista nada (inclusive arquivos nao rastreados). Devolve a
       lista suja em -Sujos para a mensagem de erro. #>
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [ref]$Sujos
    )
    $saida = @(& git -C $RepoRoot status --porcelain)
    if ($LASTEXITCODE -ne 0) {
        throw "git status falhou (exit $LASTEXITCODE)."
    }
    $linhas = @($saida | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($Sujos) { $Sujos.Value = $linhas }
    return ($linhas.Count -eq 0)
}

function Test-VersionTagExists {
    <# Confere se existe a tag LOCAL "v<Version>" (git tag -l; sem rede, independente do branch atual). #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$RepoRoot
    )
    $tag = "v$Version"
    $found = & git -C $RepoRoot tag -l $tag
    if ($LASTEXITCODE -ne 0) {
        throw "git tag -l falhou (exit $LASTEXITCODE)."
    }
    return [bool]($found -and ($found | Where-Object { $_.Trim() -eq $tag }))
}

function Get-InstallerExePath {
    <# Caminho esperado do instalador (o mesmo OutputBaseFilename do moffice.iss e o asset que o
       atualizador procura) - so monta a string. #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$OutputDir
    )
    return Join-Path $OutputDir "mOffice-Setup-$Version.exe"
}

function Get-TagCommitDate {
    <# Data de commit (DateTime UTC) do commit que a tag LOCAL "v<Version>" aponta, via
       `git log -1 --format=%ct` (le so o repositorio local). Usada pelo cinto de frescor. #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$RepoRoot
    )
    $tag = "v$Version"
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $epoch = & git -C $RepoRoot log -1 --format=%ct "refs/tags/$tag" 2>$null
        $codigo = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevEap
    }
    if ($codigo -ne 0) {
        throw "git log falhou ao ler a data de commit da tag '$tag' (exit $codigo)."
    }
    if ([string]::IsNullOrWhiteSpace($epoch)) {
        throw "Nao foi possivel ler a data de commit da tag '$tag' (git log nao devolveu nada)."
    }
    return [DateTimeOffset]::FromUnixTimeSeconds([long]("$epoch".Trim())).UtcDateTime
}

function Test-InstallerFresh {
    <# CINTO DE FRESCOR (comparacao pura): o instalador reaproveitado e fresco se foi escrito NA ou DEPOIS
       da data de commit da tag (>=: os dois no mesmo segundo nao e falso positivo). #>
    param(
        [Parameter(Mandatory)][datetime]$ExeLastWriteTimeUtc,
        [Parameter(Mandatory)][datetime]$TagCommitDateUtc
    )
    return ($ExeLastWriteTimeUtc -ge $TagCommitDateUtc)
}

function Test-AssetSize {
    <# $true se o asset cabe no limite do GitHub (estritamente menor que 2 GiB, a mesma comparacao do
       build-installer.ps1). Arquivo vazio tambem e recusado (build quebrado). #>
    param([Parameter(Mandatory)][long]$Bytes)
    return ($Bytes -gt 0 -and $Bytes -lt $script:LimiteAssetGitHub)
}

function Format-ReleaseNotes {
    <# CONTRATO do corpo (ver o cabecalho do release.ps1 e ServicoAtualizacao.ExtrairSha256): notas livres
       aparadas, linha em branco, e "SHA256: <hash>" como ULTIMA linha. Sem notas, o corpo e SO a linha do
       hash. O leitor toma a ULTIMA linha que casa com ^SHA256:\s*([0-9a-fA-F]{64})$. #>
    param(
        [string]$Notas,
        [Parameter(Mandatory)][string]$Sha256
    )
    $shaLine = "SHA256: $Sha256"
    if ([string]::IsNullOrWhiteSpace($Notas)) {
        return $shaLine
    }
    $normalizadas = ($Notas.Trim()) -replace "`r`n", "`n"
    return "$normalizadas`n`n$shaLine"
}

function Write-ReleaseNotesFile {
    <# Grava o corpo num arquivo (UTF-8 SEM BOM, byte a byte) para o --notes-file. Um BOM viraria o
       primeiro caractere do corpo da release. #>
    param(
        [Parameter(Mandatory)][string]$Conteudo,
        [Parameter(Mandatory)][string]$Caminho
    )
    [System.IO.File]::WriteAllText($Caminho, $Conteudo, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-ReleaseCreateArgs {
    <# Argumentos do `gh release create`. O corpo vai SEMPRE por --notes-file, nunca por --notes "texto":
       passar o corpo como argumento fez o gh/PowerShell 5.1 quebrar em aspas e parenteses ("no matches
       found for ...") e abortou duas releases do mPDF (v2.8.2 e v2.9.0). #>
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$ExePath,
        [Parameter(Mandatory)][string]$Repo,
        [Parameter(Mandatory)][string]$NotesFile
    )
    return @('release', 'create', "v$Version", $ExePath, '--repo', $Repo, '--title', "mOffice $Version", '--notes-file', $NotesFile)
}
