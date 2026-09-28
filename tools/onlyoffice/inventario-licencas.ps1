<# Gera third_party/onlyoffice/INVENTARIO.md: tamanho por pasta, subpastas de fontes e todo arquivo
   de licenca/aviso encontrado (LICENSE*, *.txt com "License"/"OFL"/"Apache"), para a Task 8. #>
[CmdletBinding()] param([string]$Raiz = (Join-Path $PSScriptRoot "..\..\third_party\onlyoffice"))
$ErrorActionPreference = "Stop"
# Resolve para caminho canonico: $Raiz por padrao contem "..\..\" nao resolvido, o que o torna
# mais LONGO como string que o FullName (ja normalizado) devolvido por Get-ChildItem, quebrando
# o Substring($Raiz.Length + 1) mais abaixo.
$Raiz = (Resolve-Path $Raiz).Path
$saida = Join-Path $Raiz "INVENTARIO.md"
$linhas = @("# Inventario do motor ONLYOFFICE redistribuido", "", "Gerado em $(Get-Date -Format s)", "")
$linhas += "| Pasta | MB | Arquivos |"; $linhas += "|---|---|---|"
foreach ($p in @("converter","editors\sdkjs","editors\web-apps","editors\sdkjs-plugins","fonts","dictionaries")) {
    $d = Join-Path $Raiz $p
    if (Test-Path $d) {
        $arqs = Get-ChildItem -Recurse -File $d
        $mb = [math]::Round(($arqs | Measure-Object Length -Sum).Sum / 1MB, 1)
        $linhas += "| $p | $mb | $($arqs.Count) |"
    }
}
$linhas += ""; $linhas += "## Subpastas de fontes"; $linhas += ""
Get-ChildItem -Directory (Join-Path $Raiz "fonts") | ForEach-Object { $linhas += "- $($_.Name): $((Get-ChildItem -File $_.FullName | Measure-Object).Count) arquivos" }
$soltas = Get-ChildItem -File (Join-Path $Raiz "fonts") | ForEach-Object { $_.Name }
if ($soltas) { $linhas += "- (arquivos soltos) $($soltas -join ', ')" }
$linhas += ""; $linhas += "## Arquivos de licenca/aviso encontrados"; $linhas += ""
Get-ChildItem -Recurse -File $Raiz | Where-Object {
    $_.Name -match '^(LICENSE|LICENCE|COPYING|NOTICE|OFL|3rd)' -or ($_.Extension -eq ".txt" -and $_.Length -lt 200KB -and (Select-String -Path $_.FullName -Pattern 'License|OFL|Apache|GPL|MIT' -Quiet))
} | ForEach-Object { $linhas += "- $($_.FullName.Substring($Raiz.Length + 1))" }
$linhas | Set-Content -Encoding utf8 $saida
Write-Host "Inventario gravado em $saida"
