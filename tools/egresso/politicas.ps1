<#
.SYNOPSIS
    Aplica (ou remove) as politicas de registro do WebView2 que cortam a comunicacao de fundo. CAMADA 2 do
    egresso zero (Marco 1 / Plano 1 / Task 3).

.DESCRIPTION
    Mesma lista de mOffice.Nucleo.Egresso.Politicas - e' o INSTALADOR quem aplica isto em producao; este
    script existe para a bancada de medida e para desfazer a mao. O teste
    PoliticasTests.ScriptDaBancada_NaoDivergiuDaLista falha se as duas listas divergirem.

    ExperimentationAndConfigurationServiceControl = 0 (RestrictedMode): o processo do navegador nao consulta
    o servico de experimentacao/configuracao do Edge. Esse servico e' ligado ANTES de qualquer argumento de
    linha de comando nosso chegar - por isso politica de registro, e nao mais uma flag.

    PRECISA DE ADMINISTRADOR (a chave e' HKLM). Sem elevacao o script para' com uma mensagem, nao com um
    erro de acesso negado no meio do caminho.

    ATENCAO ao editar: ASCII puro, sem BOM (o Windows PowerShell 5.1 le um .ps1 sem BOM como ANSI).

.PARAMETER Remover
    Apaga os VALORES aplicados (nunca a chave inteira: outros produtos podem ter politicas ali).

.EXAMPLE
    powershell -File tools/egresso/politicas.ps1            # aplica
    powershell -File tools/egresso/politicas.ps1 -Remover   # desfaz
#>
[CmdletBinding()]
param([switch]$Remover)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$elevado = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $elevado) { throw "este script precisa de administrador (a chave de politica fica em HKLM)" }

$chave = 'HKLM\SOFTWARE\Policies\Microsoft\Edge\WebView2'

# Chamadas ao reg.exe (executavel NATIVO) com o stderr silenciado SEM derrubar o script (fix final da revisao
# do branch): no Windows PowerShell 5.1, sob $ErrorActionPreference = 'Stop', redirecionar o stderr de um nativo
# (2>&1 ou 2>$null) embrulha cada linha num NativeCommandError e ABORTA - e o reg.exe escreve no stderr
# justamente no caso "valor/chave ja' ausente", que para -Remover e' sucesso idempotente, nao erro. Dentro do
# bloco `& { ... }` a preferencia volta a 'Continue' (escopo local); quem decide e' o codigo de saida
# ($LASTEXITCODE: 0 = achou/apagou, 1 = nao existe). O nome do valor fica LITERAL depois da opcao de valor do
# reg.exe (nunca numa variavel) porque PoliticasTests.ScriptDaBancada_NaoDivergiuDaLista le a lista deste
# arquivo por essa forma - e nenhum comentario pode imitar essa forma, senao entra na lista.
if ($Remover) {
    & { $ErrorActionPreference = 'Continue'; & reg query "$chave" /v ExperimentationAndConfigurationServiceControl 2>$null } | Out-Null
    if ($LASTEXITCODE -eq 0) {
        & { $ErrorActionPreference = 'Continue'; & reg delete "$chave" /v ExperimentationAndConfigurationServiceControl /f 2>$null } | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "reg delete falhou com codigo $LASTEXITCODE" }
        Write-Host "politica removida: ExperimentationAndConfigurationServiceControl"
    }
    else {
        Write-Host "politica ja' ausente: ExperimentationAndConfigurationServiceControl"
    }
}
else {
    & reg add "$chave" /v ExperimentationAndConfigurationServiceControl /t REG_DWORD /d 0 /f | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "reg add falhou com codigo $LASTEXITCODE" }
    Write-Host "politica aplicada: ExperimentationAndConfigurationServiceControl = 0 (RestrictedMode)"
}

# Estado final, para o relatorio da medida (a chave pode nem existir depois de -Remover numa maquina limpa).
$estado = & { $ErrorActionPreference = 'Continue'; & reg query "$chave" 2>$null }
if ($LASTEXITCODE -eq 0) { $estado | Write-Host } else { Write-Host "chave ausente: $chave" }
