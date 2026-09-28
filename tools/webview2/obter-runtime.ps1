<#
.SYNOPSIS
    Baixa (ou usa um .cab ja' baixado do) WebView2 Fixed Version Runtime, confere o sha256 pinado e extrai
    para third_party/webview2/<versao>/ (fora do git).

.DESCRIPTION
    Marco 1 / Plano 1 / Task 3 (S2 - egresso zero). O runtime Evergreen do WebView2 e' atualizado pela
    Microsoft e, mesmo com --disable-background-networking, o PROCESSO DO NAVEGADOR abriu 2 conexoes TLS/443
    para IPs da Microsoft por janela na medida do Marco 0. O runtime FIXO (um .cab com a arvore completa do
    msedgewebview2.exe) e' a primeira das tres camadas de egresso zero: nada de servico de atualizacao, nada
    de componentes baixados, e a versao do motor de renderizacao passa a ser a que NOS escolhemos - o que
    tambem elimina a classe de bug "o editor parou de abrir porque o Edge atualizou sozinho".

    O .cab NAO e' commitado (>250 MB; third_party/ inteiro fica fora do git). O que o git guarda e' este
    script, o sha256 pinado e os termos de licenca em docs/licencas/LICENCA-WebView2-Runtime.txt.

    ATENCAO ao editar este arquivo: como os demais scripts de tools/, ele e' ASCII puro e sem BOM. O Windows
    PowerShell 5.1 le um .ps1 sem BOM como ANSI, e um travessao ou acento vira ruido que quebra o parser.

.PARAMETER Cab
    Caminho local de um .cab ja' baixado. Tem prioridade sobre -Url.

.PARAMETER Url
    URL direta do .cab. O padrao e' a versao pinada em $UrlPadrao (anotada na Task 3 a partir da pagina
    oficial developer.microsoft.com/microsoft-edge/webview2 - secao "Fixed Version").

.PARAMETER Sha256Esperado
    sha256 do .cab. Pinado apos o primeiro download; um pacote que nao confere ABORTA (fail-closed).

.PARAMETER Destino
    Raiz da extracao. Padrao: third_party/webview2 (ignorado pelo git).

.PARAMETER Enxugar
    Depois de extrair, apaga do pacote os executaveis nao essenciais que o Fixed Version tambem traz e o
    mOffice nao usa: copilotapp.exe, mscopilot.exe, platform_experiences_helper.exe, notification_helper.exe,
    elevated_tracing_service.exe e Installer\copilot_setup.exe (~20 MB dos ~668 MB extraidos). Fix da revisao
    1 da Task 3 (S2), achado 1: sao binarios com REDE PROPRIA que nenhuma flag/politica/regra desta task
    mira especificamente (a regra de firewall cobre todos, enxugado ou nao - isto e' so' para nao
    REDISTRIBUIR o que nao e' preciso).

    PADRAO: LIGADO ($true). Testado de verdade nesta revisao: com os 6 arquivos removidos do runtime
    PINADO desta maquina, SeisFormatosTests rodou 6/6 aprovado (ver poc/relatorios/egresso-S2-20260918.md,
    secao "Fix (review 1)", com o tamanho da pasta extraida antes/depois) - o editor nao chama nenhum desses
    binarios. Passe -Enxugar:$false para manter o pacote completo (ex.: se uma versao futura do runtime
    passar a depender de um desses arquivos e este script ainda nao tiver sido atualizado).

.EXAMPLE
    powershell -File tools/webview2/obter-runtime.ps1
    powershell -File tools/webview2/obter-runtime.ps1 -Enxugar:$false
#>
[CmdletBinding()]
param(
    [string]$Cab,
    [string]$Url,
    [string]$Sha256Esperado = "11e8240cb0bc56dcd3e4498907203c251346f65107fe35a3a13e152c7d51c79e",
    [string]$Destino = (Join-Path $PSScriptRoot "..\..\third_party\webview2"),
    [bool]$Enxugar = $true
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Versao pinada: a mais recente estavel listada na pagina oficial em 2026-09-18.
$UrlPadrao = 'https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/08cd33ee-d109-49b8-9301-9f0bea43c575/Microsoft.WebView2.FixedVersionRuntime.153.0.4234.48.x64.cab'

$Destino = [IO.Path]::GetFullPath($Destino)
$cache = Join-Path $Destino 'cache'
if (-not (Test-Path $cache)) { New-Item -ItemType Directory -Force -Path $cache | Out-Null }

if (-not $Cab) {
    if (-not $Url) { $Url = $UrlPadrao }
    $Cab = Join-Path $cache ([IO.Path]::GetFileName(([Uri]$Url).AbsolutePath))
    if (-not (Test-Path $Cab)) {
        Write-Host "baixando $Url"
        $parcial = "$Cab.part"
        # WebClient em vez de Invoke-WebRequest: o pacote passa de 250 MB e o Invoke-WebRequest do
        # PowerShell 5.1 carrega a resposta inteira em memoria antes de gravar.
        $cliente = New-Object System.Net.WebClient
        try { $cliente.DownloadFile($Url, $parcial) } finally { $cliente.Dispose() }
        Move-Item -Force $parcial $Cab
    }
    else { Write-Host "usando o cab ja' em cache: $Cab" }
}

if (-not (Test-Path $Cab)) { throw "cab nao encontrado: $Cab" }
$hash = (Get-FileHash -Algorithm SHA256 $Cab).Hash.ToLowerInvariant()
$tamanho = (Get-Item $Cab).Length
if ($Sha256Esperado) {
    if ($hash -ne $Sha256Esperado.ToLowerInvariant()) {
        throw "sha256 do runtime nao confere: obtido $hash, esperado $($Sha256Esperado.ToLowerInvariant())"
    }
    Write-Host "sha256 confere: $hash"
}
else { Write-Host "sha256 do pacote (pinar no script): $hash" }

$versao = ([regex]::Match([IO.Path]::GetFileName($Cab), '(\d+\.\d+\.\d+\.\d+)')).Groups[1].Value
if (-not $versao) { throw "nao consegui ler a versao do nome do arquivo: $Cab" }

$pasta = Join-Path $Destino $versao
if (Test-Path $pasta) { Remove-Item -Recurse -Force $pasta }
New-Item -ItemType Directory -Force -Path $pasta | Out-Null
Write-Host "extraindo para $pasta"
& expand.exe "$Cab" -F:* "$pasta" | Out-Null
if ($LASTEXITCODE -ne 0) { throw "expand.exe falhou com codigo $LASTEXITCODE" }

$exe = Get-ChildItem -Path $pasta -Recurse -File -Filter msedgewebview2.exe | Select-Object -First 1
if (-not $exe) { throw "msedgewebview2.exe nao encontrado apos a extracao" }

if ($Enxugar) {
    $pastaX64 = $exe.DirectoryName
    $naoEssenciais = @(
        'copilotapp.exe',
        'mscopilot.exe',
        'platform_experiences_helper.exe',
        'notification_helper.exe',
        'elevated_tracing_service.exe',
        'Installer\copilot_setup.exe'
    )
    $removidos = 0
    foreach ($rel in $naoEssenciais) {
        $alvo = Join-Path $pastaX64 $rel
        if (Test-Path $alvo) { Remove-Item -Force $alvo; $removidos++ }
    }
    # A pasta Installer so' tinha o copilot_setup.exe; se ficou vazia, some tambem.
    $pastaInstaller = Join-Path $pastaX64 'Installer'
    if ((Test-Path $pastaInstaller) -and -not (Get-ChildItem -Path $pastaInstaller -Force -ErrorAction SilentlyContinue)) {
        Remove-Item -Force $pastaInstaller
    }
    Write-Host "-Enxugar: $removidos executavel(is) nao essencial(is) removido(s) de $pastaX64"
}

# Windows 10 + runtime fixo: o processo do navegador roda em App Container e precisa de leitura/execucao
# pelos SIDs ALL_APP_PACKAGES / ALL_RESTRICTED_APP_PACKAGES (doc da Microsoft, "Fixed Version distribution
# mode"). No Windows 11 a concessao e' inofensiva - aplicamos sempre, para a pasta valer nos dois.
& icacls $exe.DirectoryName /grant "*S-1-15-2-2:(OI)(CI)(RX)" | Out-Null
& icacls $exe.DirectoryName /grant "*S-1-15-2-1:(OI)(CI)(RX)" | Out-Null

@(
    $versao
    "cab $([IO.Path]::GetFileName($Cab))"
    "cab bytes $tamanho"
    "cab sha256 $hash"
    "url $(if ($Url) { $Url } else { '(cab local)' })"
    "exe $($exe.DirectoryName)"
    "obtido em $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
) | Set-Content -Encoding utf8 (Join-Path $Destino 'VERSAO.txt')

Write-Host "WebView2 Fixed Version $versao em $($exe.DirectoryName)"
