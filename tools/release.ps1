<#
.SYNOPSIS
    Publica o instalador do mOffice como uma GitHub Release no repositorio publico
    (maiconbiondo/moffice-desktop).

.DESCRIPTION
    Marco 1 / Plano 5 / Task 6. Porte do tools/release.ps1 do mPDF, com o mesmo contrato.

    PUBLICAR E ATO HUMANO: so rode este script com o "sim" do responsavel pela release no momento. Ele cria a release
    publica (gh release create), que o botao "Verificar atualizacoes" e o /update do RMM de TODAS as
    maquinas passam a enxergar. Nunca apaga nem sobrescreve uma release.

    CONTRATO DO CORPO DA RELEASE (o atualizador do app depende dele - ServicoAtualizacao.ExtrairSha256 em
    src/mOffice.Comum/Atualizacao/):

        [notas livres, uma ou mais linhas - bloco OMITIDO se -Notas nao for informado]

        SHA256: <64 caracteres hex minusculos>

    "SHA256: <hash>" e SEMPRE a ULTIMA linha do corpo. O leitor toma a ULTIMA linha que casa com
    ^SHA256:\s*([0-9a-fA-F]{64})$ (uma linha por vez), nunca a primeira: as notas livres podem conter algo
    parecido. Nunca mover essa linha nem mudar o prefixo "SHA256:".
    O asset e "mOffice-Setup-<versao>.exe" (o OutputBaseFilename do moffice.iss); o atualizador so aceita
    release com um asset .exe E o hash no corpo.

    Passos, nesta ordem, cada um abortando se falhar:
      1. gh CLI autenticado (`gh auth status`; "nao autenticado" e precondicao normal, nunca excecao crua).
      2. Arvore de trabalho limpa (git status --porcelain vazio).
      3. Versao de src/Directory.Build.props (<Version>, fonte UNICA; formato X.Y.Z - a release candidate
         e 0.99.0, porque System.Version nao aceita "-rc" e o releases/latest ignora pre-release).
      4. Tag LOCAL v<versao> existe (git tag -l; independente do branch atual).
      5. Instalador tools/installer/Output/mOffice-Setup-<versao>.exe. Se existe, CINTO DE FRESCOR: a data
         de escrita (UTC) precisa ser >= a data de commit da tag; senao ABORTA (pode ser um build antigo com
         o nome certo). Se nao existe, roda tools/installer/build-installer.ps1 (suite completa + pacote +
         ISCC; exige sessao de console) e confere que o .exe apareceu.
         Depois, TAMANHO: um asset de release do GitHub aceita no maximo 2 GiB (o build-installer.ps1 ja
         confere; aqui de novo porque o .exe pode ser reaproveitado). Medido na Task 3: ~443 MB.
      6. SHA-256 do instalador (Get-FileHash).
      7. IDEMPOTENCIA: `gh release view v<versao> --repo <repo>`. Se a release ja existe, ABORTA sem tocar
         em nada. Depois, o corpo (contrato acima) vai para um arquivo temporario UTF-8 sem BOM e roda
         `gh release create v<versao> <exe> --repo <repo> --title "mOffice <versao>" --notes-file <arq>`.
         Nunca `--notes "texto"` (quebrou duas releases do mPDF por aspas e parenteses). O repositorio
         publico ainda nao tem a tag: o gh a cria na ponta da main publica - por isso a main publica
         precisa JA ter o codigo desta versao (docs/rollout.md, secao da release).
      8. Imprime a URL da release.

    As funcoes ficam em tools/release-functions.ps1 (testadas por tools/release.Tests.ps1, sem rede e sem
    o gh de verdade).

    COMO CHAMAR COM NOTAS (gotcha do PowerShell 5.1, visto no mPDF): use o operador de chamada na MESMA
    sessao,
        $n = 'Correcao: X (ver "Y").'
        & .\tools\release.ps1 -Notas $n
    e NUNCA `powershell -File tools\release.ps1 -Notas "..."`: o -File re-tokeniza a linha de comando e as
    aspas duplas das notas vazam para os proximos parametros (ex.: o -Repo). Com o `&`, a string chega
    inteira e vai byte a byte para o --notes-file.

    ASCII puro e sem BOM, como os demais scripts de tools/.

.PARAMETER Notas
    Texto livre (opcional), em portugues. Se omitido, o corpo contem SO a linha SHA256.

.PARAMETER Repo
    Repositorio publico de destino, "dono/repositorio". Padrao: maiconbiondo/moffice-desktop (a mesma
    constante ServicoAtualizacao.Repositorio que o app consulta).

.EXAMPLE
    $n = 'Primeira versao publica.'
    & .\tools\release.ps1 -Notas $n

.EXAMPLE
    & .\tools\release.ps1
    Corpo so com a linha SHA256.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Notas,

    [Parameter(Mandatory = $false)]
    [string]$Repo = 'maiconbiondo/moffice-desktop'
)

$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$PropsFile = Join-Path $RepoRoot 'src\Directory.Build.props'
$OutputDir = Join-Path $RepoRoot 'tools\installer\Output'
$BuildScript = Join-Path $RepoRoot 'tools\installer\build-installer.ps1'

. (Join-Path $PSScriptRoot 'release-functions.ps1')

Write-Host "=== [1/8] Confirmando autenticacao do gh CLI ===" -ForegroundColor Cyan
if (-not (Test-GhAuthenticated)) {
    throw "gh CLI nao esta autenticado neste ambiente. Rode 'gh auth login' antes de publicar a release."
}
Write-Host "  OK: gh autenticado." -ForegroundColor Green

Write-Host "=== [2/8] Confirmando arvore de trabalho limpa ===" -ForegroundColor Cyan
$sujos = @()
if (-not (Test-WorkingTreeClean -RepoRoot $RepoRoot -Sujos ([ref]$sujos))) {
    $lista = ($sujos | ForEach-Object { "    $_" }) -join "`n"
    throw "Arvore de trabalho suja - commit ou descarte as mudancas antes de publicar a release:`n$lista"
}
Write-Host "  OK: arvore limpa." -ForegroundColor Green

Write-Host "=== [3/8] Extraindo a versao do produto (src\Directory.Build.props) ===" -ForegroundColor Cyan
$version = Get-ProductVersion -PropsPath $PropsFile
Write-Host "  Versao: $version" -ForegroundColor Green

Write-Host "=== [4/8] Confirmando tag local v$version ===" -ForegroundColor Cyan
if (-not (Test-VersionTagExists -Version $version -RepoRoot $RepoRoot)) {
    throw "Tag 'v$version' nao encontrada no repositorio local (git tag -l). Marque a versao antes de publicar: git tag v$version <commit da main>."
}
Write-Host "  OK: tag v$version existe." -ForegroundColor Green

Write-Host "=== [5/8] Confirmando instalador (Output\mOffice-Setup-$version.exe) ===" -ForegroundColor Cyan
$exePath = Get-InstallerExePath -Version $version -OutputDir $OutputDir
if (Test-Path -LiteralPath $exePath) {
    $exeTime = (Get-Item -LiteralPath $exePath).LastWriteTimeUtc
    $tagTime = Get-TagCommitDate -Version $version -RepoRoot $RepoRoot
    if (-not (Test-InstallerFresh -ExeLastWriteTimeUtc $exeTime -TagCommitDateUtc $tagTime)) {
        throw "Instalador $exePath e MAIS ANTIGO (build de $exeTime UTC) que o commit da tag v$version ($tagTime UTC) - pode nao refletir o codigo que a tag marca. Rode tools\installer\build-installer.ps1 de novo antes de publicar."
    }
    Write-Host "  Reaproveitando instalador existente: $exePath (build $exeTime UTC >= tag $tagTime UTC)" -ForegroundColor Green
} else {
    Write-Host "  Nao encontrado - rodando build-installer.ps1 (suite completa + pacote + ISCC)..." -ForegroundColor Yellow
    & $BuildScript
    if (-not (Test-Path -LiteralPath $exePath)) {
        throw "build-installer.ps1 terminou mas $exePath nao existe - o build nao gerou o instalador da versao $version."
    }
    Write-Host "  OK: instalador gerado em $exePath" -ForegroundColor Green
}
$tamanho = (Get-Item -LiteralPath $exePath).Length
if (-not (Test-AssetSize -Bytes $tamanho)) {
    throw "Instalador com $tamanho bytes: fora do limite de um asset de release do GitHub (maior que zero e menor que 2 GiB)."
}
Write-Host ("  Tamanho: {0:N1} MB ({1} bytes; limite do GitHub: 2 GiB)" -f ($tamanho / 1MB), $tamanho) -ForegroundColor Green

Write-Host "=== [6/8] Calculando SHA-256 do instalador ===" -ForegroundColor Cyan
$hash = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash.ToLowerInvariant()
Write-Host "  SHA256: $hash" -ForegroundColor Green

Write-Host "=== [7/8] Conferindo se a release v$version ja existe em $Repo ===" -ForegroundColor Cyan
if (Test-ReleaseExists -Version $version -Repo $Repo) {
    throw "Release 'v$version' ja existe em $Repo - abortando (nunca sobrescreve release publicada). Confira com: gh release view v$version --repo $Repo"
}
Write-Host "  OK: release ainda nao existe." -ForegroundColor Green

Write-Host "=== [8/8] Publicando release v$version em $Repo ===" -ForegroundColor Cyan
$notes = Format-ReleaseNotes -Notas $Notas -Sha256 $hash
$notesFile = Join-Path ([System.IO.Path]::GetTempPath()) "moffice-release-notes-$version-$([guid]::NewGuid().ToString('N')).md"
Write-ReleaseNotesFile -Conteudo $notes -Caminho $notesFile
try {
    $ghArgs = Get-ReleaseCreateArgs -Version $version -ExePath $exePath -Repo $Repo -NotesFile $notesFile
    $releaseUrl = & gh @ghArgs
    if ($LASTEXITCODE -ne 0) {
        throw "gh release create falhou (exit $LASTEXITCODE)."
    }
}
finally { Remove-Item -LiteralPath $notesFile -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host "=== Release publicada com sucesso ===" -ForegroundColor Green
Write-Host "  $releaseUrl" -ForegroundColor Green
