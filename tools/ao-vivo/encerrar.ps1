<# Encerra SO' o que o lancar.ps1 abriu: o PID de <Dados>\pid.txt e os msedgewebview2.exe cuja linha de comando cita a
   pasta de dados da task. Nunca mexe em outro mOffice/Word/Excel aberto pelo usuario. #>
[CmdletBinding()] param([Parameter(Mandatory)][string]$Dados)
$ErrorActionPreference = "Stop"
$Dados = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Dados)
$arqPid = Join-Path $Dados "pid.txt"
if (Test-Path $arqPid) {
    $id = [int](Get-Content $arqPid -TotalCount 1)
    Stop-Process -Id $id -ErrorAction SilentlyContinue
    Remove-Item $arqPid
}
$alvo = [IO.Path]::GetFullPath($Dados)
Get-CimInstance Win32_Process -Filter "Name='msedgewebview2.exe'" |
    Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($alvo, [StringComparison]::OrdinalIgnoreCase) -ge 0 } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -ErrorAction SilentlyContinue }
Write-Host "encerrado: $Dados"
