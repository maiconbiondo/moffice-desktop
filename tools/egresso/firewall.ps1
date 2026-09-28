<#
.SYNOPSIS
    Uma regra de firewall de SAIDA por executavel do runtime WebView2 embarcado (nao so' o
    msedgewebview2.exe: o pacote Fixed Version tambem traz Copilot e helpers com rede propria - achado 1 da
    revisao 1). CAMADA 3 do egresso zero (Marco 1 / Plano 1 / Task 3).

.DESCRIPTION
    Defesa em profundidade: mesmo que uma versao futura do runtime volte a falar com a rede por um caminho
    que as flags e a politica nao cobrem, o Windows nega a saida. Cada regra aponta para o CAMINHO de um
    executavel especifico dentro do runtime embarcado - nunca para o nome do arquivo em geral -, entao o
    WebView2 (e o Copilot) dos OUTROS aplicativos da maquina continuam funcionando.

    Uma regra por .exe encontrado RECURSIVAMENTE dentro da pasta do runtime (hoje: msedgewebview2.exe,
    copilotapp.exe, mscopilot.exe, platform_experiences_helper.exe, notification_helper.exe,
    elevated_tracing_service.exe, Installer\copilot_setup.exe - a lista nao e' fixa no script: o que existir
    na pasta na hora ganha regra). O nome de cada regra e' "<Nome> (<exe>)"; -Remover apaga TODAS as regras
    cujo nome comeca com <Nome>, sem precisar saber a lista de exe's de antemao - se um runtime futuro
    trouxer mais binarios, -Remover ainda os limpa.

    Trafego de loopback nao passa pelo firewall do Windows, entao o editor (que e' servido de
    https://moffice.local por mapeamento de pasta e do esquema ascdesktop://) nao precisa de regra de
    liberacao.

    PRECISA DE ADMINISTRADOR para criar ou apagar regras. Reler as regras existentes (netsh ... show rule)
    NAO precisa de administrador - e' o que tests/mOffice.Nucleo.Tests/FirewallTests.cs usa para verificar,
    sem elevacao, que toda regra aponta para o runtime ATUAL.

    REPIN (procedimento para quando o runtime for atualizado): tools/webview2/obter-runtime.ps1 baixando uma
    versao nova extrai para uma pasta third_party/webview2/<versao-nova> DIFERENTE da atual, e as regras
    desta execucao continuam apontando para a pasta VELHA ate alguem agir. Rode este script DE NOVO depois
    do obter-runtime.ps1: ele apaga as regras antigas (pelo prefixo do nome, sem precisar saber quais
    existiam) e recria apontando para o caminho novo - idempotente. Ate isso acontecer,
    FirewallTests.RegraApontaParaORuntimeAtual fica VERMELHO de proposito: e' o guarda contra uma regra
    pendurada num caminho que nao existe mais (a versao velha pode ate ter sido apagada do disco).

    ATENCAO ao editar: ASCII puro, sem BOM.

.PARAMETER Pasta
    Pasta raiz do runtime embarcado (a que contem msedgewebview2.exe e os demais executaveis). Se omitido,
    procura em third_party/webview2 a versao mais alta com msedgewebview2.exe.

.PARAMETER Remover
    Apaga TODAS as regras cujo nome comeca com -Nome (nao precisa de -Pasta).

.PARAMETER Nome
    Prefixo do nome de cada regra; o nome completo de cada regra e' "<Nome> (<exe>)".

.EXAMPLE
    powershell -File tools/egresso/firewall.ps1
    powershell -File tools/egresso/firewall.ps1 -Remover
#>
[CmdletBinding()]
param(
    [string]$Pasta,
    [switch]$Remover,
    [string]$Nome = "mOffice WebView2 - bloqueio de saida"
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Le as regras JA' EXISTENTES cujo nome comeca com $Prefixo, sem exigir administrador (so' consulta) e sem
# depender do idioma do Windows (rotulo em ingles OU em portugues).
function Get-NomesDeRegrasMOffice {
    param([string]$Prefixo)
    $saida = & netsh advfirewall firewall show rule name=all
    $nomes = New-Object System.Collections.Generic.List[string]
    foreach ($linha in $saida) {
        if ($linha -match '^(Rule Name|Nome da Regra):\s*(.+?)\s*$') {
            $nome = $Matches[2]
            if ($nome.StartsWith($Prefixo)) { [void]$nomes.Add($nome) }
        }
    }
    return $nomes
}

$elevado = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($Remover) {
    if (-not $elevado) { throw "este script precisa de administrador (apagar regra de firewall)" }
    # @(...) e' obrigatorio aqui (e nao cosmetico): uma funcao que devolve uma List[string] com 0 ou 1 item e'
    # DESENROLADA pelo pipeline do PowerShell - 0 itens vira $null, 1 item vira uma STRING pura - e nenhum dos
    # dois tem propriedade .Count sob Set-StrictMode. Bug real, achado rodando este script de verdade (nao so'
    # os testes): sem o @() a chamada com zero regras restantes lancava "propriedade Count nao encontrada" em
    # vez de confirmar a remocao.
    $existentes = @(Get-NomesDeRegrasMOffice -Prefixo $Nome)
    foreach ($n in $existentes) {
        # 2>$null (em vez de 2>&1 | Out-Null): no Windows PowerShell 5.1 redirecionar o stderr de um
        # executavel NATIVO para dentro do pipeline embrulha cada linha num NativeCommandError e derruba $?
        # mesmo quando o netsh sai com codigo 0 - a mensagem so' precisa sumir da tela, nao virar erro.
        & netsh advfirewall firewall delete rule name="$n" 2>$null | Out-Null
    }
    # Verifica DE VERDADE que a remocao pegou: rele as regras depois de apagar, em vez de confiar no
    # resultado do netsh (fix da revisao 1, achado 6).
    $restantes = @(Get-NomesDeRegrasMOffice -Prefixo $Nome)
    if ($restantes.Count -gt 0) { throw "$($restantes.Count) regra(s) continuam apos -Remover: $($restantes -join ', ')" }
    Write-Host "$($existentes.Count) regra(s) removida(s), 0 restante(s):"
    foreach ($n in $existentes) { Write-Host "  - $n" }
    exit 0
}

if (-not $elevado) { throw "este script precisa de administrador (criar regra de firewall)" }

if (-not $Pasta) {
    $raiz = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'third_party\webview2'
    if (-not (Test-Path $raiz)) { throw "third_party\webview2 nao existe: rode tools/webview2/obter-runtime.ps1 ou passe -Pasta" }
    $exePrincipal = Get-ChildItem -Path $raiz -Recurse -File -Filter msedgewebview2.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $exePrincipal) { throw "msedgewebview2.exe nao encontrado em $raiz" }
    $Pasta = $exePrincipal.DirectoryName
}
if (-not (Test-Path $Pasta)) { throw "pasta nao encontrada: $Pasta" }

# TODOS os .exe da pasta, recursivamente (achado 1 da revisao 1): o pacote Fixed Version nao traz so' o
# msedgewebview2.exe, e um executavel networkavel que ninguem bloqueia e' um furo no egresso zero tao real
# quanto o proprio navegador.
$exes = @(Get-ChildItem -Path $Pasta -Recurse -File -Filter *.exe -ErrorAction SilentlyContinue)
if ($exes.Count -eq 0) { throw "nenhum .exe encontrado em $Pasta" }

$criadas = @()
foreach ($exe in $exes) {
    $nomeRegra = "$Nome ($($exe.Name))"
    # Idempotente: apaga antes de criar, para nao empilhar regras iguais a cada execucao.
    & netsh advfirewall firewall delete rule name="$nomeRegra" 2>$null | Out-Null
    & netsh advfirewall firewall add rule name="$nomeRegra" dir=out action=block program="$($exe.FullName)" enable=yes profile=any | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "netsh add rule falhou (codigo $LASTEXITCODE) para $($exe.FullName)" }
    $criadas += $nomeRegra
}

Write-Host "$($criadas.Count) regra(s) criada(s) em ${Pasta}:"
foreach ($n in $criadas) {
    Write-Host "  - $n"
    & netsh advfirewall firewall show rule name="$n" verbose | Select-String "Program|Programa|Action|Acao|Enabled|Habilitado" | ForEach-Object { Write-Host "      $_" }
}
