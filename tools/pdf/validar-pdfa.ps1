<#
.SYNOPSIS
    Plano 4 / Task 1 (sondagem S6) - valida um PDF contra um perfil PDF/A no veraPDF e escreve UMA linha JSON.

.DESCRIPTION
    Ferramenta de verificacao local, fora do produto (o mOffice nao tem rede nem chama isto em tempo de uso).

    Saida (stdout, uma linha, ASCII):
        {"conforme":true|false,"flavour":"2a","regrasFalhas":N,"regras":["6.2.11.4.1-1",...],"resumo":"..."}
    - flavour: o perfil que o veraPDF USOU (tirado de profileName, ex.: "PDF/A-2A validation profile" -> "2a").
      Com -Flavour auto o veraPDF escolhe pelo XMP do arquivo (pdfaid:part/conformance); se o arquivo nao declara
      PDF/A, cai no padrao do veraPDF (1b).
    - regras: "<clausula>-<teste>" de cada regra que falhou (o id do veraPDF), na ordem do relatorio.

    Codigo de saida: 0 = conforme; 1 = nao conforme; 2 = veraPDF nao encontrado; 3 = erro ao rodar/ler o veraPDF.

    Onde procura o veraPDF (nesta ordem): $env:VERAPDF_HOME\verapdf.bat, `verapdf` no PATH,
    C:\Program Files\veraPDF\verapdf.bat e %USERPROFILE%\verapdf\verapdf.bat (instalacao por usuario).

    Opcoes do veraPDF 1.30.2 conferidas no --help: --format json, --flavour (0 = automatico, 1a..3u, 4, ua1...).
    O JSON e' report.jobs[0].validationResult[0] com compliant/profileName/details.ruleSummaries
    (validationResult e' LISTA na 1.30).

.PARAMETER Pdf
    Caminho do PDF a validar.

.PARAMETER Flavour
    auto (padrao) | 1a | 1b | 2a | 2b | 2u | 3a | 3b | 3u.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Pdf,
    [ValidateSet('auto', '1a', '1b', '2a', '2b', '2u', '3a', '3b', '3u')][string]$Flavour = 'auto'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Escrever([bool]$conforme, [string]$flavour, [int]$falhas, [string[]]$regras, [string]$resumo) {
    # ASCII na saida: tira acento e qualquer caractere fora de 0x20..0x7E do resumo (mensagens do Java/PowerShell).
    $limpo = $resumo.Normalize([Text.NormalizationForm]::FormD)
    $limpo = -join ($limpo.ToCharArray() | Where-Object { [int]$_ -ge 0x20 -and [int]$_ -le 0x7E })
    $obj = [ordered]@{
        conforme     = $conforme
        flavour      = $flavour
        regrasFalhas = $falhas
        regras       = @($regras)
        resumo       = $limpo
    }
    Write-Output ($obj | ConvertTo-Json -Compress -Depth 4)
}

function Localizar-VeraPdf {
    $candidatos = @()
    if ($env:VERAPDF_HOME) { $candidatos += (Join-Path $env:VERAPDF_HOME 'verapdf.bat') }
    $noPath = Get-Command verapdf -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($noPath) { $candidatos += $noPath.Source }
    $candidatos += 'C:\Program Files\veraPDF\verapdf.bat'
    if ($env:USERPROFILE) { $candidatos += (Join-Path $env:USERPROFILE 'verapdf\verapdf.bat') }
    foreach ($c in $candidatos) { if ($c -and (Test-Path -LiteralPath $c -PathType Leaf)) { return $c } }
    return $null
}

$vera = Localizar-VeraPdf
if (-not $vera) {
    Escrever $false $Flavour 0 @() 'veraPDF nao encontrado (VERAPDF_HOME, PATH, Program Files, %USERPROFILE%\verapdf)'
    exit 2
}
if (-not (Test-Path -LiteralPath $Pdf -PathType Leaf)) {
    Escrever $false $Flavour 0 @() "PDF nao encontrado: $Pdf"
    exit 3
}

$argFlavour = if ($Flavour -eq 'auto') { '0' } else { $Flavour }
try {
    # So' o stdout e' o relatorio; o stderr (log do Java) e' descartado para nao sujar o JSON. No PowerShell 5.1,
    # com $ErrorActionPreference = 'Stop', QUALQUER linha no stderr de um executavel nativo vira ErrorRecord e erro
    # TERMINANTE mesmo com 2>$null (o catch abaixo respondia 'exit 3' falso). So' em volta desta chamada: 'Continue'.
    $eapAntes = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $bruto = & $vera --format json --flavour $argFlavour (Resolve-Path -LiteralPath $Pdf).Path 2>$null }
    finally { $ErrorActionPreference = $eapAntes }
    $texto = ($bruto | Out-String)
    $inicio = $texto.IndexOf('{')
    if ($inicio -lt 0) { throw "saida do veraPDF sem JSON" }
    $rel = $texto.Substring($inicio) | ConvertFrom-Json
    $job = @($rel.report.jobs)[0]
    $vr = @($job.validationResult)
    if ($vr.Count -eq 0 -or $null -eq $vr[0]) { throw "veraPDF nao validou o arquivo (sem validationResult)" }
    $v = $vr[0]
    $perfil = [string]$v.profileName
    $usado = $Flavour
    $m = [regex]::Match($perfil, 'PDF/A-(\d)([ABU])', 'IgnoreCase')
    if ($m.Success) { $usado = ($m.Groups[1].Value + $m.Groups[2].Value).ToLowerInvariant() }
    $regras = @()
    $falhas = 0
    $det = $v.details
    if ($det) {
        $falhas = [int]$det.failedRules
        foreach ($r in @($det.ruleSummaries)) {
            if ($null -ne $r -and $r.status -ne 'passed') { $regras += ('{0}-{1}' -f $r.clause, $r.testNumber) }
        }
    }
    $conforme = [bool]$v.compliant
    Escrever $conforme $usado $falhas $regras ("{0}: {1}" -f $perfil, $v.statement)
    if ($conforme) { exit 0 } else { exit 1 }
}
catch {
    Escrever $false $Flavour 0 @() ("erro ao rodar o veraPDF: " + $_.Exception.Message)
    exit 3
}
