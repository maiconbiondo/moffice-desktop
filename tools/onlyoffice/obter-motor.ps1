<#
.SYNOPSIS
  Copia do ONLYOFFICE Desktop Editors instalado APENAS as pastas que o mOffice redistribui
  (converter, editors, fonts, dictionaries) para third_party/onlyoffice/ (ignorada pelo git),
  e grava VERSAO.txt (versao + sha256 do x2t.exe). NAO copia CEF, Qt, DesktopEditors.exe,
  editors.exe, updatesvc.exe.

  NAO copia `fonts\ASC.ttf` (familia ASCW3): licenca nao localizavel ("All rights reserved" na tabela
  `name`, nenhum LICENSE ao lado dela no pacote nem no repo core-fonts). O experimento da Task 4/S3 mediu
  que o arquivo e' DISPENSAVEL - a evidencia esta em docs\licencas\AVISO-fontes.md. Use -ComAscTtf para
  refazer o A/B com ela.
.PARAMETER Origem
  Pasta da instalacao oficial. Padrao: C:\Program Files\ONLYOFFICE\DesktopEditors
.PARAMETER ComAscTtf
  Copia tambem `fonts\ASC.ttf`. So' para repetir o experimento da Task 4/S3 - NAO use para gerar o pacote
  que vai ser distribuido.
#>
[CmdletBinding()]
param(
    [string]$Origem = "C:\Program Files\ONLYOFFICE\DesktopEditors",
    [string]$Destino = (Join-Path $PSScriptRoot "..\..\third_party\onlyoffice"),
    [switch]$ArquivarMsi,
    [switch]$ComAscTtf,
    [string]$UrlMsi = "https://download.onlyoffice.com/install/desktop/editors/windows/distrib/onlyoffice/DesktopEditors_x64.msi"
)
$ErrorActionPreference = "Stop"
$pastas = @("converter", "editors", "fonts", "dictionaries")
foreach ($p in $pastas) {
    if (-not (Test-Path (Join-Path $Origem $p))) { throw "Pasta '$p' nao encontrada em '$Origem'. Instale o ONLYOFFICE Desktop Editors ou passe -Origem." }
}
New-Item -ItemType Directory -Force $Destino | Out-Null
foreach ($p in $pastas) {
    $de = Join-Path $Origem $p; $para = Join-Path $Destino $p
    Write-Host "Copiando $p ..."
    if (Test-Path $para) { Remove-Item -Recurse -Force $para }
    # robocopy: /E subpastas (inclusive vazias), /NFL /NDL sem lista de arquivos, /NJH /NJS sem cabecalho/sumario
    # /XF ASC.ttf: a fonte sem licenca localizavel fica FORA do pacote por padrao (ver cabecalho + AVISO-fontes.md).
    $excluir = if ($p -eq 'fonts' -and -not $ComAscTtf) { @('/XF', 'ASC.ttf') } else { @() }
    & robocopy $de $para /E /NFL /NDL /NJH /NJS /R:1 /W:1 @excluir | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy falhou em '$p' (codigo $LASTEXITCODE)." }
}

# Guarda do resultado, nao da intencao: confere no que FICOU no destino. Se o /XF parar de funcionar (nome
# mudado, robocopy trocado por outra copia), isto falha em vez de deixar a fonte entrar no pacote calada.
$ascNoDestino = Join-Path $Destino "fonts\ASC.ttf"
if ($ComAscTtf) {
    if (-not (Test-Path $ascNoDestino)) { throw "-ComAscTtf pedido mas fonts\ASC.ttf nao chegou ao destino." }
    Write-Warning "fonts\ASC.ttf COPIADA (-ComAscTtf): licenca nao localizavel - so' para experimento, nao distribua este pacote."
} elseif (Test-Path $ascNoDestino) {
    throw "fonts\ASC.ttf entrou no pacote apesar da exclusao - veja docs\licencas\AVISO-fontes.md."
}
$exe = Join-Path $Origem "DesktopEditors.exe"
$versao = if (Test-Path $exe) { (Get-Item $exe).VersionInfo.ProductVersion } else { "desconhecida" }
$x2t = Join-Path $Destino "converter\x2t.exe"
if (-not (Test-Path $x2t)) { throw "x2t.exe nao encontrado apos a copia." }
$hash = (Get-FileHash -Algorithm SHA256 $x2t).Hash.ToLowerInvariant()
@($versao, $hash) | Set-Content -Encoding utf8 (Join-Path $Destino "VERSAO.txt")

if ($ArquivarMsi) {
    if ($versao -eq "desconhecida") { throw "versao do ONLYOFFICE desconhecida - nao e possivel arquivar o MSI." }
    $cache = Join-Path $Destino "cache"; New-Item -ItemType Directory -Force $cache | Out-Null
    $msi = Join-Path $cache "DesktopEditors_x64_$versao.msi"
    $msiPart = "$msi.part"

    $precisaBaixar = $true
    if (Test-Path $msi) {
        $tamanhoLocal = (Get-Item $msi).Length
        try {
            $resp = Invoke-WebRequest -Method Head -Uri $UrlMsi -UseBasicParsing
            $tamanhoRemoto = [int64]$resp.Headers["Content-Length"]
            if ($tamanhoLocal -eq $tamanhoRemoto) {
                $precisaBaixar = $false
            } else {
                Write-Host "Aviso: arquivo em cache incompleto ($tamanhoLocal de $tamanhoRemoto bytes) - baixando de novo."
            }
        } catch {
            Write-Host "Aviso: nao foi possivel verificar a integridade do arquivo em cache (sem acesso a rede) - mantendo o arquivo existente."
            $precisaBaixar = $false
        }
    }

    if ($precisaBaixar) {
        if (Test-Path $msiPart) { Remove-Item -Force $msiPart }
        Write-Host "Baixando MSI oficial para arquivo (fonte correspondente AGPL) ..."
        try {
            Invoke-WebRequest -Uri $UrlMsi -OutFile $msiPart -UseBasicParsing
        } catch {
            if (Test-Path $msiPart) { Remove-Item -Force $msiPart }
            throw
        }
        Move-Item -Force $msiPart $msi
    }

    $hashMsi = (Get-FileHash -Algorithm SHA256 $msi).Hash.ToLowerInvariant()
    # Commit PINADO, nao um ponteiro para "conferir depois": a fonte correspondente da AGPL tem de identificar
    # exatamente o que foi redistribuido. v9.4.0.129 do ONLYOFFICE/core resolve para o commit abaixo
    # (verificado por `git ls-remote --tags https://github.com/ONLYOFFICE/core.git`, 2026-09-17). Se a versao
    # baixada deixar de ser a 9.4.0.129, a linha diz isso em vez de afirmar um commit que nao corresponde.
    $commitCore = if ($versao -eq '9.4.0.129') { 'a016fc2868b9f99f59dd30aa2088e68bbe5104af' } else { 'NAO PINADO PARA ESTA VERSAO - rode git ls-remote --tags https://github.com/ONLYOFFICE/core.git' }
    Add-Content -Encoding utf8 (Join-Path $Destino "VERSAO.txt") -Value @("msi sha256 $hashMsi", "core tag: v$versao commit $commitCore")
}

Write-Host "Motor ONLYOFFICE $versao copiado para $Destino (x2t sha256 $hash); aplicando os patches do mOffice ..."

# Patches do mOffice no sdkjs (hoje: a fonte de simbolos "ASCW3" embutida em base64 nos 4 sdk-all.js e' trocada
# pela nossa, livre - docs\licencas\AVISO-fontes.md secao 1.5). Roda DEPOIS do VERSAO.txt, porque o script so'
# patcha a versao para a qual o patch foi medido; um motor recem-obtido sai daqui ja' patchado. Se o patch
# falhar, o throw dele interrompe aqui - a linha de "pronto" abaixo so' sai com o motor patchado.
& (Join-Path $PSScriptRoot "aplicar-patches.ps1") -PastaMotor $Destino
Write-Host "Motor ONLYOFFICE $versao pronto em $Destino (copiado e patchado)."
