<#
.SYNOPSIS
    Monta a pasta instalavel do mOffice (mTexto, mPlanilha, mSlides + motor ONLYOFFICE enxuto + runtime
    WebView2 fixo + licencas), no layout de producao que o ConfiguracaoMotor.Localizar() procura primeiro.

.DESCRIPTION
    Marco 1 / Plano 5 / Task 1. Layout gerado (consumido pelo .iss da Task 3):

      <Saida>\                       mTexto.exe, mPlanilha.exe, mSlides.exe, mOffice.*.dll,
                                     runtime .NET self-contained (UMA copia para os tres)
      <Saida>\motor\converter\       x2t.exe, allfontsgen.exe, DLLs do core (inclui ICU), empty\, templates\
      <Saida>\motor\editors\         web-apps + sdkjs + webext, SEM sdkjs-plugins e SEM sdkjs\common\AllFonts.js
      <Saida>\motor\fonts\           fontes livres do pacote (SEM ASC.ttf)
      <Saida>\motor\dictionaries\    pt_BR e en_US
      <Saida>\WebView2Runtime\       runtime fixo enxuto (conteudo da pasta do msedgewebview2.exe)
      <Saida>\licencas\              docs\licencas\* (menos o rascunho de e-mail)
      (e na raiz, Ponte\ SEM ponte-teste.html/.js, a pagina de teste do PoC)

    Passos, cada um abortando se falhar:
      1. dotnet publish self-contained win-x64 (pasta, NAO single-file - WPF + WebView2 precisam da pasta,
         como no build-installer.ps1 do mPDF) de cada app numa pasta de preparo PROPRIA;
      2. junta as tres publicacoes na raiz de <Saida>: arquivo com o mesmo caminho relativo e conteudo
         DIFERENTE (sha256) e' conflito e ABORTA - nunca sobrescrever em silencio. .pdb nao entra;
      3. copia o subconjunto do motor de third_party\onlyoffice e confere o que nao pode estar la';
      4. copia o runtime de third_party\webview2\<versao de VERSAO.txt> e confere que esta' enxuto;
      5. copia docs\licencas e imprime os tamanhos por pasta e o total.

    NAO roda testes (quem roda e' o build-installer.ps1 da Task 3).

    ASCII puro e sem BOM, como os demais scripts de tools/ (o Windows PowerShell 5.1 le .ps1 sem BOM como
    ANSI; um acento aqui quebraria o parser).

.PARAMETER Saida
    Pasta do pacote. Se ja' existir, precisa estar vazia ou ser um pacote anterior (tem mTexto.exe e motor\);
    qualquer outra coisa ABORTA em vez de apagar.

.EXAMPLE
    powershell -File tools\installer\montar-pacote.ps1 -Saida $env:TEMP\moffice-pacote
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Saida
)

$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$Apps = @('mTexto', 'mPlanilha', 'mSlides')
$OnlyOffice = Join-Path $RepoRoot 'third_party\onlyoffice'
$WebView2 = Join-Path $RepoRoot 'third_party\webview2'
$Licencas = Join-Path $RepoRoot 'docs\licencas'
# Dicionarios que o Corretor usa (src/mOffice.Nucleo/Ortografia/Corretor.cs: 1046 -> pt_BR, 1033 -> en_US).
$Dicionarios = @('pt_BR', 'en_US')
# Executaveis que o obter-runtime.ps1 -Enxugar remove (tem rede propria e o mOffice nao usa).
$ExesNaoEnxutos = @('copilotapp.exe', 'mscopilot.exe', 'platform_experiences_helper.exe',
    'notification_helper.exe', 'elevated_tracing_service.exe', 'copilot_setup.exe')

function Copiar-Arvore([string]$De, [string]$Para, [string[]]$ExcluirPastas = @(), [string[]]$ExcluirArquivos = @()) {
    if (-not (Test-Path $De)) { throw "Origem nao encontrada: $De" }
    $argumentos = @($De, $Para, '/E', '/NFL', '/NDL', '/NJH', '/NJS', '/NP', '/R:1', '/W:1')
    if ($ExcluirPastas.Count -gt 0) { $argumentos += '/XD'; $argumentos += $ExcluirPastas }
    if ($ExcluirArquivos.Count -gt 0) { $argumentos += '/XF'; $argumentos += $ExcluirArquivos }
    & robocopy @argumentos | Out-Null
    # robocopy: 0-7 = sucesso (com ou sem copia); 8+ = falha.
    if ($LASTEXITCODE -ge 8) { throw "robocopy falhou ($LASTEXITCODE): $De -> $Para" }
    $global:LASTEXITCODE = 0
}

function Tamanho-MB([string]$Caminho) {
    $soma = (Get-ChildItem -LiteralPath $Caminho -Recurse -File -Force | Measure-Object -Property Length -Sum).Sum
    if (-not $soma) { $soma = 0 }
    return [math]::Round($soma / 1MB, 1)
}

# ---------------------------------------------------------------------------------------------------------
Write-Host "=== [0/5] Preparando $Saida ===" -ForegroundColor Cyan
foreach ($p in @($OnlyOffice, $WebView2, $Licencas)) {
    if (-not (Test-Path $p)) { throw "Nao encontrado: $p (rode tools\onlyoffice\obter-motor.ps1, obter-allfontsgen.ps1 e tools\webview2\obter-runtime.ps1)" }
}
$Saida = [System.IO.Path]::GetFullPath($Saida)
if ($Saida.TrimEnd('\') -eq $RepoRoot.TrimEnd('\') -or $Saida.StartsWith($RepoRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "-Saida nao pode ficar dentro do repositorio ($RepoRoot): o pacote tem ~1,4 GB e nao vai para o git."
}
if ([System.IO.Path]::GetPathRoot($Saida).TrimEnd('\') -eq $Saida.TrimEnd('\')) { throw "-Saida nao pode ser a raiz de um disco." }
if (Test-Path $Saida) {
    $vazia = -not (Get-ChildItem -LiteralPath $Saida -Force | Select-Object -First 1)
    $pacoteAnterior = (Test-Path (Join-Path $Saida 'mTexto.exe')) -and (Test-Path (Join-Path $Saida 'motor'))
    if (-not ($vazia -or $pacoteAnterior)) {
        throw "$Saida existe e nao e' vazia nem um pacote anterior (sem mTexto.exe + motor\). Nada foi apagado; escolha outra pasta."
    }
    if (-not $vazia) {
        Write-Host "  apagando o pacote anterior em $Saida"
        Remove-Item -LiteralPath $Saida -Recurse -Force
    }
}
New-Item -ItemType Directory -Path $Saida -Force | Out-Null
$Preparo = Join-Path ([System.IO.Path]::GetTempPath()) ("moffice-publish-" + [guid]::NewGuid().ToString('N'))

try {
    # -----------------------------------------------------------------------------------------------------
    Write-Host "=== [1/5] Publicando os tres apps (self-contained, win-x64, pasta) ===" -ForegroundColor Cyan
    foreach ($app in $Apps) {
        $proj = Join-Path $RepoRoot "src\$app\$app.csproj"
        $out = Join-Path $Preparo $app
        & dotnet publish $proj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -o $out
        if ($LASTEXITCODE -ne 0) { throw "dotnet publish de $app falhou (exit code $LASTEXITCODE)." }
        if (-not (Test-Path (Join-Path $out "$app.exe"))) { throw "$app.exe nao apareceu em $out apos o publish." }
    }

    # -----------------------------------------------------------------------------------------------------
    Write-Host "=== [2/5] Juntando as publicacoes na raiz (conflito = aborta) ===" -ForegroundColor Cyan
    $vistos = @{}      # caminho relativo -> @{ Hash; App }
    $conflitos = New-Object System.Collections.Generic.List[string]
    $iguais = 0
    foreach ($app in $Apps) {
        $raizApp = Join-Path $Preparo $app
        foreach ($f in Get-ChildItem -LiteralPath $raizApp -Recurse -File) {
            if ($f.Extension -ieq '.pdb') { continue }
            $rel = $f.FullName.Substring($raizApp.Length + 1)
            $hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
            if ($vistos.ContainsKey($rel)) {
                if ($vistos[$rel].Hash -ne $hash) {
                    $conflitos.Add("$rel ($($vistos[$rel].App) x $app)")
                } else { $iguais++ }
                continue
            }
            $vistos[$rel] = @{ Hash = $hash; App = $app }
            $destino = Join-Path $Saida $rel
            New-Item -ItemType Directory -Path (Split-Path $destino -Parent) -Force | Out-Null
            Copy-Item -LiteralPath $f.FullName -Destination $destino
        }
    }
    if ($conflitos.Count -gt 0) {
        throw ("Conflito entre as publicacoes (mesmo arquivo, conteudo diferente) - os tres apps nao podem dividir a pasta assim:`n  " + ($conflitos -join "`n  "))
    }
    foreach ($app in $Apps) {
        foreach ($f in @("$app.exe", "$app.dll", "$app.deps.json", "$app.runtimeconfig.json")) {
            if (-not (Test-Path (Join-Path $Saida $f))) { throw "Faltou $f na raiz do pacote." }
        }
    }
    # Pagina de teste do PoC (Ponte\ponte-teste.html/.js): so' o HostFumacaTests navega ate' ela; o produto
    # nunca. Sai do pacote (o mapeamento moffice-poc.local continua, apontando para uma pasta sem ela).
    Get-ChildItem -LiteralPath (Join-Path $Saida 'Ponte') -Filter 'ponte-teste.*' | Remove-Item -Force
    if (Get-ChildItem -LiteralPath $Saida -Recurse -Filter 'ponte-teste.*') { throw "ponte-teste.* ainda no pacote." }
    if (-not (Test-Path (Join-Path $Saida 'Ponte\ponte-desktop.js'))) { throw "Ponte\ponte-desktop.js ausente no pacote." }
    Write-Host ("  OK: {0} arquivos distintos no pacote dos apps, {1} copias identicas entre os apps descartadas, 0 conflitos." -f $vistos.Count, $iguais) -ForegroundColor Green

    # -----------------------------------------------------------------------------------------------------
    Write-Host "=== [3/5] Motor ONLYOFFICE (subconjunto) ===" -ForegroundColor Cyan
    $Motor = Join-Path $Saida 'motor'
    Copiar-Arvore (Join-Path $OnlyOffice 'converter') (Join-Path $Motor 'converter')
    Copiar-Arvore (Join-Path $OnlyOffice 'editors') (Join-Path $Motor 'editors') -ExcluirPastas @((Join-Path $OnlyOffice 'editors\sdkjs-plugins'))
    Copiar-Arvore (Join-Path $OnlyOffice 'fonts') (Join-Path $Motor 'fonts') -ExcluirArquivos @('ASC.ttf')
    foreach ($d in $Dicionarios) {
        Copiar-Arvore (Join-Path $OnlyOffice "dictionaries\$d") (Join-Path $Motor "dictionaries\$d")
    }
    # AllFonts.js e' o catalogo de UMA maquina (sobra de versoes antigas, que o copiavam para o motor). Hoje o
    # editor o recebe do catalogo pelo host (CatalogoNoEditor) e o motor nao precisa dele: o do build nao vai.
    $allFontsBuild = Join-Path $Motor 'editors\sdkjs\common\AllFonts.js'
    if (Test-Path $allFontsBuild) { Remove-Item -LiteralPath $allFontsBuild -Force }
    # preload.html e cache-scripts.html carregam o AllFonts.js por <script src> ESTATICO (markup do parser), que a
    # reescrita do CatalogoNoEditor nao pega (so' .src atribuido por JS). Dormentes no mOffice (o warmUp do
    # api.js exige ?placeholder=; cache-scripts nao e' referenciado): saem do pacote.
    foreach ($dormente in @('preload.html', 'cache-scripts.html')) {
        $p = Join-Path $Motor "editors\web-apps\apps\api\documents\$dormente"
        if (Test-Path $p) { Remove-Item -LiteralPath $p -Force }
    }
    $restantes = @(Get-ChildItem -LiteralPath (Join-Path $Motor 'editors') -Recurse -File | Where-Object { $_.Name -ieq 'preload.html' -or $_.Name -ieq 'cache-scripts.html' })
    if ($restantes.Count -gt 0) { throw ("Pagina com <script src> estatico do AllFonts.js ainda no pacote: " + ($restantes.FullName -join ', ')) }

    foreach ($obrigatorio in @('converter\x2t.exe', 'converter\allfontsgen.exe', 'converter\icuuc74.dll', 'converter\icudt74.dll',
                               'editors\web-apps', 'editors\sdkjs\word', 'editors\sdkjs\cell', 'editors\sdkjs\slide',
                               'dictionaries\pt_BR\pt_BR.dic', 'dictionaries\en_US\en_US.dic')) {
        if (-not (Test-Path (Join-Path $Motor $obrigatorio))) { throw "Motor incompleto: falta motor\$obrigatorio" }
    }
    if (Test-Path (Join-Path $Motor 'editors\sdkjs-plugins')) { throw "sdkjs-plugins nao deveria estar no pacote." }
    $proibidos = @(Get-ChildItem -LiteralPath $Saida -Recurse -File -Force | Where-Object { $_.Name -ieq 'ASC.ttf' -or $_.Name -ieq 'AllFonts.js' })
    if ($proibidos.Count -gt 0) { throw ("Arquivo que nao pode ir no pacote: " + ($proibidos.FullName -join ', ')) }
    $dicts = @(Get-ChildItem -LiteralPath (Join-Path $Motor 'dictionaries') -Directory | ForEach-Object Name)
    if ($dicts.Count -ne $Dicionarios.Count) { throw "Dicionarios inesperados: $($dicts -join ', ')" }
    $versaoMotor = (Get-Content (Join-Path $OnlyOffice 'VERSAO.txt') -TotalCount 1).Trim([char]0xFEFF, ' ')
    Write-Host "  OK: motor $versaoMotor sem sdkjs-plugins, sem ASC.ttf, sem AllFonts.js, dicionarios: $($dicts -join ', ')" -ForegroundColor Green

    # -----------------------------------------------------------------------------------------------------
    Write-Host "=== [4/5] Runtime WebView2 fixo ===" -ForegroundColor Cyan
    $versaoRuntime = (Get-Content (Join-Path $WebView2 'VERSAO.txt') -TotalCount 1).Trim([char]0xFEFF, ' ')
    $raizRuntime = Join-Path $WebView2 $versaoRuntime
    if (-not (Test-Path $raizRuntime)) { throw "Runtime $versaoRuntime nao extraido em $raizRuntime (rode tools\webview2\obter-runtime.ps1)." }
    $exeRuntime = Get-ChildItem -LiteralPath $raizRuntime -Recurse -Filter 'msedgewebview2.exe' | Select-Object -First 1
    if (-not $exeRuntime) { throw "msedgewebview2.exe nao encontrado em $raizRuntime." }
    $sujos = @(Get-ChildItem -LiteralPath $exeRuntime.DirectoryName -Recurse -File | Where-Object { $ExesNaoEnxutos -contains $_.Name.ToLowerInvariant() })
    if ($sujos.Count -gt 0) { throw ("Runtime nao enxuto (rode obter-runtime.ps1 -Enxugar): " + ($sujos.Name -join ', ')) }
    Copiar-Arvore $exeRuntime.DirectoryName (Join-Path $Saida 'WebView2Runtime')
    if (-not (Test-Path (Join-Path $Saida 'WebView2Runtime\msedgewebview2.exe'))) { throw "WebView2Runtime\msedgewebview2.exe ausente no pacote." }
    Write-Host "  OK: runtime $versaoRuntime (enxuto) em WebView2Runtime\" -ForegroundColor Green

    # -----------------------------------------------------------------------------------------------------
    Write-Host "=== [5/5] Licencas e tamanhos ===" -ForegroundColor Cyan
    # O rascunho do e-mail a Ascensio e' documento de trabalho (tem nome pessoal), nao licenca.
    Copiar-Arvore $Licencas (Join-Path $Saida 'licencas') -ExcluirArquivos @('email-*')
    if (-not (Test-Path (Join-Path $Saida 'licencas\LICENCA-ONLYOFFICE-AGPL-3.0.txt'))) { throw "licencas\LICENCA-ONLYOFFICE-AGPL-3.0.txt ausente." }

    $linhas = @()
    $arquivosRaiz = @(Get-ChildItem -LiteralPath $Saida -File -Force)
    $linhas += [pscustomobject]@{ Pasta = '(raiz: apps + .NET)'; MB = [math]::Round((($arquivosRaiz | Measure-Object Length -Sum).Sum) / 1MB, 1); Arquivos = $arquivosRaiz.Count }
    foreach ($sub in @(Get-ChildItem -LiteralPath $Saida -Directory | Where-Object Name -ne 'motor') + @(Get-ChildItem -LiteralPath $Motor -Directory)) {
        $rotulo = if ($sub.Parent.FullName -eq $Motor) { "motor\$($sub.Name)" } else { $sub.Name }
        $linhas += [pscustomobject]@{ Pasta = $rotulo; MB = (Tamanho-MB $sub.FullName); Arquivos = @(Get-ChildItem -LiteralPath $sub.FullName -Recurse -File -Force).Count }
    }
    $linhas | Format-Table -AutoSize | Out-String | Write-Host
    $total = Tamanho-MB $Saida
    $qtd = @(Get-ChildItem -LiteralPath $Saida -Recurse -File -Force).Count
    Write-Host ("TOTAL: {0} MB em {1} arquivos ({2})" -f $total, $qtd, $Saida) -ForegroundColor Green
}
finally {
    if (Test-Path $Preparo) { Remove-Item -LiteralPath $Preparo -Recurse -Force -ErrorAction SilentlyContinue }
}
