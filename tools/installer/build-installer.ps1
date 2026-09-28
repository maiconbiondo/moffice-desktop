<#
.SYNOPSIS
    Gera o instalador do mOffice (Inno Setup): tools/installer/Output/mOffice-Setup-<versao>.exe.

.DESCRIPTION
    Marco 1 / Plano 5 / Task 3. Derivado do tools/installer/build-installer.ps1 do mPDF. Passos, nesta
    ordem, cada um abortando o build se falhar:
      0. Confere que o texto de termos docs\licencas\AVISO-mOffice.txt (InfoBeforeFile do .iss, Task 4)
         existe - ANTES da suite, para nao esperar minutos por um build que nao pode terminar;
      1. roda a suite COMPLETA (dotnet test mOffice.slnx) - nunca gera instalador com a suite vermelha;
      2. monta o pacote (tools/installer/montar-pacote.ps1) numa pasta em %TEMP%;
      3. GUARDA de caminho pessoal nos binarios do pacote (tokens derivados em runtime, como no mPDF);
      4. le a versao de src/Directory.Build.props (<Version>, fonte UNICA) e chama o ISCC com
         /DMyAppVersion e /DPacoteDir;
      5. imprime o tamanho do instalador e o tempo do ISCC.
    A pasta do pacote (~1,5 GB) e apagada no fim, com sucesso ou nao.

    NAO instala nada. Instalar nesta maquina (a de desenvolvimento) misturaria o registro do produto com o
    do tools/explorer/registrar-dev.ps1 e as regras de firewall do tools/egresso/firewall.ps1.

    ASCII puro e sem BOM, como os demais scripts de tools/.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools\installer\build-installer.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$SlnFile = Join-Path $RepoRoot 'mOffice.slnx'
$PropsFile = Join-Path $RepoRoot 'src\Directory.Build.props'
$Aviso = Join-Path $RepoRoot 'docs\licencas\AVISO-mOffice.txt'
$MontarPacote = Join-Path $PSScriptRoot 'montar-pacote.ps1'
$IssFile = Join-Path $PSScriptRoot 'moffice.iss'
$OutputDir = Join-Path $PSScriptRoot 'Output'
$PacoteDir = Join-Path ([System.IO.Path]::GetTempPath()) 'moffice-pacote-instalador'

function Find-Iscc {
    # Mesmos caminhos do mPDF: instalacao por usuario via winget (JRSoftware.InnoSetup), depois as
    # instalacoes para todos os usuarios, depois o PATH.
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
        (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe')
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path $c)) { return $c }
    }
    $onPath = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }
    throw "ISCC.exe (Inno Setup 6) nao encontrado em nenhum dos caminhos conhecidos nem no PATH. Instale com: winget install JRSoftware.InnoSetup"
}

Write-Host "=== [0/5] Conferindo o texto de termos do instalador ===" -ForegroundColor Cyan
if (-not (Test-Path $Aviso)) {
    throw "docs\licencas\AVISO-mOffice.txt nao existe. E o texto de termos que o instalador mostra antes de instalar (InfoBeforeFile; Task 4 do Plano 5) e ainda nao foi escrito. Build do instalador ABORTADO - nao gerar instalador com texto juridico provisorio."
}
Write-Host "  OK: $Aviso" -ForegroundColor Green

Write-Host "=== [1/5] Rodando a suite de testes completa (dotnet test) ===" -ForegroundColor Cyan
& dotnet test $SlnFile
if ($LASTEXITCODE -ne 0) {
    throw "Suite de testes falhou (exit code $LASTEXITCODE). Build do instalador ABORTADO - nunca gerar instalador com a suite vermelha."
}

try {
    Write-Host "=== [2/5] Montando o pacote (montar-pacote.ps1) em $PacoteDir ===" -ForegroundColor Cyan
    & $MontarPacote -Saida $PacoteDir
    if (-not (Test-Path (Join-Path $PacoteDir 'mTexto.exe'))) {
        throw "montar-pacote.ps1 terminou sem mTexto.exe em $PacoteDir."
    }

    # -----------------------------------------------------------------------------------------------------
    # GUARDA DE VAZAMENTO DE CAMINHO (pos-publish, pre-ISCC) - porte do build-installer.ps1 do mPDF.
    #
    # Os binarios podem embutir o caminho absoluto da maquina de build (simbolos, strings de depuracao), com
    # o nome do usuario/pastas pessoais entre o drive e a pasta do repositorio. O Directory.Build.props da
    # raiz (Deterministic + PathMap) existe para isso nao acontecer; a prova e varrer os bytes do que vai
    # para o instalador e ABORTAR se algum residuo aparecer.
    #
    # Os tokens NAO sao fixos (este script vai para o repositorio publico): saem do proprio $RepoRoot em
    # tempo de execucao. De "<drive>\<a>\<b>\mOffice", os segmentos entre o drive e a pasta do repo sao as
    # partes sensiveis; entram o caminho inteiro (nos dois separadores) e cada segmento.
    #
    # Varre o pacote INTEIRO (apps, motor e runtime) nos tipos que carregam texto de build: .dll, .exe, .pdb
    # (o .iss os exclui, mas a guarda nao depende disso), .json e .js (a Ponte\ e o motor). Le cada arquivo
    # como latin1 (1 byte = 1 caractere, sem erro de decodificacao).
    # -----------------------------------------------------------------------------------------------------
    Write-Host "=== [3/5] Guarda: varrendo o pacote por caminho de build pessoal ===" -ForegroundColor Cyan
    $segmentosRepo = $RepoRoot -split '[\\/]' | Where-Object { $_ -ne '' }
    $segmentosSensiveis = @()
    if ($segmentosRepo.Count -gt 2) {
        $segmentosSensiveis = $segmentosRepo[1..($segmentosRepo.Count - 2)]
    }
    $caminhosLiteral = @($RepoRoot, ($RepoRoot -replace '\\', '/'))
    # Tambem o perfil e o nome do usuario QUE ESTA GERANDO o build (lidos do ambiente agora, nunca escritos
    # aqui): um caminho de %USERPROFILE% (cache do NuGet, %TEMP%) pode vazar sem passar pela pasta do repo.
    if (-not [string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        $caminhosLiteral += @($env:USERPROFILE, ($env:USERPROFILE -replace '\\', '/'))
    }
    if (-not [string]::IsNullOrWhiteSpace($env:USERNAME) -and ($segmentosSensiveis -notcontains $env:USERNAME)) {
        $segmentosSensiveis = @($segmentosSensiveis) + $env:USERNAME
    }
    # Segmento so casa quando NAO esta colado a [A-Za-z0-9]: num caminho real ele vem entre separadores. Isso
    # evita o falso positivo conhecido do mPDF - o slug publico do repositorio embutido no app (o nome do
    # usuario colado a outras letras, em "github.com/<usuario>/..."), que nao e caminho de build.
    $regexesSegmento = foreach ($seg in $segmentosSensiveis) {
        [regex]::new('(?<![A-Za-z0-9])' + [regex]::Escape($seg) + '(?![A-Za-z0-9])', 'IgnoreCase')
    }
    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $arquivosBin = @(Get-ChildItem -Path $PacoteDir -Recurse -File -Include '*.dll', '*.exe', '*.pdb', '*.json', '*.js')
    if ($arquivosBin.Count -eq 0) {
        throw "GUARDA: nenhum .dll/.exe/.json/.js encontrado em $PacoteDir para varrer - pacote vazio nao e caminho feliz."
    }
    $vazamentos = New-Object System.Collections.Generic.List[string]
    foreach ($bin in $arquivosBin) {
        $texto = $latin1.GetString([System.IO.File]::ReadAllBytes($bin.FullName))
        foreach ($cam in $caminhosLiteral) {
            if ($texto.IndexOf($cam, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $vazamentos.Add(("{0}  ->  caminho do repositorio" -f $bin.FullName))
            }
        }
        foreach ($rx in $regexesSegmento) {
            if ($rx.IsMatch($texto)) {
                $vazamentos.Add(("{0}  ->  segmento do caminho do repositorio" -f $bin.FullName))
            }
        }
        $texto = $null
    }
    if ($vazamentos.Count -gt 0) {
        Write-Host "GUARDA DE CAMINHO RECUSOU O BUILD - caminho pessoal embutido em $($vazamentos.Count) ocorrencia(s):" -ForegroundColor Red
        foreach ($v in $vazamentos) { Write-Host "  $v" -ForegroundColor Red }
        throw "Caminho de build pessoal vazou nos binarios do pacote. Confira o Directory.Build.props (PathMap/Deterministic) na raiz do repo. Build ABORTADO."
    }
    Write-Host "  Limpo: nenhum caminho pessoal em $($arquivosBin.Count) arquivo(s)." -ForegroundColor Green

    Write-Host "=== [4/5] Versao (src/Directory.Build.props) e ISCC ===" -ForegroundColor Cyan
    [xml]$props = Get-Content -Raw $PropsFile
    $version = $props.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Nao foi possivel extrair <Version> de $PropsFile."
    }
    Write-Host "  Versao: $version" -ForegroundColor Green
    $iscc = Find-Iscc
    Write-Host "  ISCC: $iscc"
    if (-not (Test-Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
    $cronometro = [System.Diagnostics.Stopwatch]::StartNew()
    & $iscc "/DMyAppVersion=$version" "/DPacoteDir=$PacoteDir" $IssFile
    if ($LASTEXITCODE -ne 0) {
        throw "ISCC.exe falhou (exit code $LASTEXITCODE)."
    }
    $cronometro.Stop()

    Write-Host "=== [5/5] Instalador gerado ===" -ForegroundColor Green
    $instalador = Join-Path $OutputDir "mOffice-Setup-$version.exe"
    if (-not (Test-Path $instalador)) { throw "ISCC terminou sem gerar $instalador." }
    $tam = (Get-Item $instalador).Length
    Write-Host ("  {0}" -f $instalador) -ForegroundColor Green
    Write-Host ("  Tamanho: {0:N1} MB ({1} bytes); ISCC: {2:N1} min" -f ($tam / 1MB), $tam, $cronometro.Elapsed.TotalMinutes) -ForegroundColor Green
    # Limite de um asset de release do GitHub: 2 GiB.
    if ($tam -ge 2GB) { throw "O instalador passou de 2 GiB ($tam bytes): o GitHub recusa o asset da release." }
}
finally {
    if (Test-Path $PacoteDir) {
        Write-Host "Apagando o pacote temporario $PacoteDir ..."
        Remove-Item -LiteralPath $PacoteDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
