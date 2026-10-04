<#
.SYNOPSIS
  Sobe UM app do mOffice (build Release) com porta CDP e pasta de dados proprias, para o roteiro ao vivo do Plano 6.
  Recusa porta ja' ocupada (incidente da auditoria de 2026-10-01: dois testadores na mesma porta). Grava o PID em
  <Dados>\pid.txt para o encerrar.ps1 matar SO' o que foi aberto aqui.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet("mTexto","mPlanilha","mSlides")][string]$App,
    [Parameter(Mandatory)][int]$Porta,
    [Parameter(Mandatory)][string]$Dados,
    [string]$Arquivo,
    # Plano 6, Task 2: o verbo que o Explorer passa antes do arquivo (/modelo = documento NOVO do modelo; /apresentar = abre apresentando; /imprimir = verbo print: converte, imprime na padrao e sai com 0/20).
    [ValidateSet("", "/modelo", "/apresentar", "/imprimir")][string]$Verbo = "",
    [switch]$Diagnostico
)
$ErrorActionPreference = "Stop"
if (Get-NetTCPConnection -LocalPort $Porta -State Listen -ErrorAction SilentlyContinue) {
    Write-Host "porta $Porta ja esta em uso: escolha a porta da sua task (tabela do plano)"; exit 1
}
$raiz = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$exe = Join-Path $raiz "src\$App\bin\Release\net10.0-windows\$App.exe"
if (-not (Test-Path $exe)) { Write-Host "nao achei $exe - rode: dotnet build mOffice.slnx -c Release"; exit 1 }
$Dados = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Dados)
New-Item -ItemType Directory -Force $Dados | Out-Null
$env:MOFFICE_DADOS_DIR = $Dados
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=$Porta"
if ($Diagnostico) { $env:MOFFICE_DIAGNOSTICO = "1" } else { Remove-Item Env:MOFFICE_DIAGNOSTICO -ErrorAction SilentlyContinue }
$argumentos = @(); if ($Verbo) { $argumentos += $Verbo }; if ($Arquivo) { $argumentos += "`"$Arquivo`"" }
$p = if ($argumentos.Count) { Start-Process -FilePath $exe -ArgumentList $argumentos -PassThru } else { Start-Process -FilePath $exe -PassThru }
Set-Content -Path (Join-Path $Dados "pid.txt") -Value $p.Id -Encoding ascii
Write-Host ("lancado {0} pid {1} porta {2} dados {3}" -f $App, $p.Id, $Porta, $Dados)
