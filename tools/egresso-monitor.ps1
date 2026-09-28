<#
.SYNOPSIS
    Amostra as conexoes TCP REMOTAS dos processos do PoC durante uma sessao de abrir/editar/salvar.

.DESCRIPTION
    Item de gate do Marco 0 (protocolo, secao "Auditoria de egresso"; plano Task 8 Step 2). A guarda de
    egresso do WebView2 prova o que passa PELA guarda; este script prova o que sai PELA PILHA DE REDE DO
    SISTEMA - inclusive o que a guarda nao veria (o x2t, o proprio host, os processos internos do WebView2).

    A cada intervalo o script:
      1. levanta os processos vigiados: TODO processo cujo executavel principal (MainModule) roda de DENTRO
         da pasta do runtime WebView2 embarcado (third_party/webview2/<versao>/...) - e nao so' o
         msedgewebview2.exe: o pacote Fixed Version tambem traz Copilot e helpers com rede propria (achado 1
         da revisao 1) -, mais `x2t` e `mOffice.Sonda`/`mTexto`/`mPlanilha`/`mSlides` POR NOME (fora do runtime).
         Uma lista fixa de nomes nunca pegaria um binario novo que uma versao futura do runtime trouxesse;
         o caminho pega, porque o criterio e' "roda da NOSSA copia", nao "se chama X";
      2. pergunta ao Windows (`Get-NetTCPConnection`) todas as conexoes desses PIDs;
      3. DESCARTA loopback/curinga (127.0.0.1, ::1, 0.0.0.0, ::) - essas sao internas ao WebView2 e nao sao
         egresso;
      4. acrescenta ao arquivo de saida o que sobrou.

    O criterio do gate e' ZERO conexoes remotas. O arquivo e' de APPEND e registra TODAS as amostras,
    inclusive as vazias: uma amostragem que so registra achados nao consegue provar ausencia - e' preciso
    saber quantas vezes se olhou e nao se viu nada.

    O arquivo de saida guarda apenas nomes de processo, PIDs, enderecos IP e portas - nenhum caminho de
    arquivo e nenhum dado pessoal -, por isso pode ser commitado como evidencia.

    ATENCAO ao editar este arquivo: como os demais scripts de tools/poc, ele e' ASCII puro e sem BOM. O
    Windows PowerShell 5.1 le um .ps1 sem BOM como ANSI, e um travessao ou acento vira ruido que quebra o
    parser (ja aconteceu na primeira versao deste script).

.PARAMETER Segundos
    Duracao total da amostragem. Padrao: 900 (15 min), o suficiente para uma rodada de SeisFormatosTests.

.PARAMETER IntervaloSegundos
    Intervalo entre amostras. Padrao: 10 (o que o plano pede).

.PARAMETER Saida
    Arquivo de saida. Padrao: poc/relatorios/egresso-conexoes-<aaaaMMdd>.txt

.EXAMPLE
    # Em um console, comeca a amostrar; em outro, roda a suite:
    powershell -File tools/egresso-monitor.ps1 -Segundos 900
    dotnet test tests/mOffice.Nucleo.Tests --filter SeisFormatosTests
#>
[CmdletBinding()]
param(
    [int]$Segundos = 900,
    [int]$IntervaloSegundos = 10,
    [string]$Saida
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$raiz = Split-Path -Parent $PSScriptRoot
if (-not $Saida) {
    $Saida = Join-Path $raiz "poc/relatorios/egresso-conexoes-$(Get-Date -Format 'yyyyMMdd').txt"
}
$pastaSaida = Split-Path -Parent $Saida
if (-not (Test-Path $pastaSaida)) { New-Item -ItemType Directory -Path $pastaSaida -Force | Out-Null }

# Enderecos que NAO sao egresso: loopback e curingas de escuta.
$naoEgresso = @('127.0.0.1', '::1', '0.0.0.0', '::')

# Vigiados POR NOME (fora da pasta do runtime): o conversor e os hosts da suite. Curta de proposito - tudo
# que roda de DENTRO da pasta do runtime entra pelo CAMINHO (Resolver-RaizRuntime abaixo), nao por estar
# nesta lista, entao ela nao precisa listar Copilot nem os helpers do WebView2.
$processosPorNome = @('x2t', 'mOffice.Sonda', 'mTexto', 'mPlanilha', 'mSlides')

# Acha a pasta do runtime WebView2 embarcado (a versao mais alta com msedgewebview2.exe dentro de
# third_party/webview2), do mesmo jeito que AmbienteWebView2.PastaRuntimeFixo faz em C# para o ramo de
# desenvolvimento. $null se o runtime fixo nao estiver baixado - neste caso o script so' vigia por nome
# (nenhum processo do WebView2 entra na amostra, o que e' o esperado: sem runtime fixo nao ha "nossa copia"
# nenhuma para identificar por caminho).
function Resolver-RaizRuntime {
    param([string]$RaizRepo)
    $dev = Join-Path $RaizRepo 'third_party\webview2'
    if (-not (Test-Path $dev)) { return $null }
    $versoes = Get-ChildItem -Path $dev -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+(\.\d+){1,3}$' } |
        Sort-Object { [version]$_.Name } -Descending
    foreach ($v in $versoes) {
        $exe = Get-ChildItem -Path $v.FullName -Recurse -File -Filter msedgewebview2.exe -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exe) { return $exe.DirectoryName }
    }
    return $null
}
$raizRuntime = Resolver-RaizRuntime -RaizRepo $raiz
if ($raizRuntime) { Write-Host "raiz do runtime vigiada (por caminho): $raizRuntime" }
else { Write-Host "AVISO: runtime fixo nao encontrado em third_party/webview2 - vigiando so' por nome ($($processosPorNome -join ', '))" }

# Nomes que, aparecendo na ASCENDENCIA de um processo, identificam a arvore do PoC. `testhost`/`dotnet`
# cobrem as janelas abertas pela suite de testes; `mOffice.Sonda`, a execucao avulsa do host.
$donosDoPoC = @('moffice.sonda', 'testhost', 'dotnet', 'vstest.console')

# Monta a cadeia de ancestrais de um PID, so com NOMES de processo (nunca caminhos nem linha de comando,
# que trariam dados pessoais para um arquivo commitado). E' o que permite dizer de QUEM e' a conexao:
# um `msedgewebview2` nao diz nada sozinho, ja que outros aplicativos da maquina usam o mesmo binario.
function Get-CadeiaDeAncestrais {
    param([int]$PidAlvo, [hashtable]$Tabela)
    $cadeia = @()
    $atual = $PidAlvo
    $vistos = @{}
    for ($i = 0; $i -lt 12; $i++) {
        if (-not $Tabela.ContainsKey($atual)) { break }
        if ($vistos.ContainsKey($atual)) { break }
        $vistos[$atual] = $true
        $info = $Tabela[$atual]
        $cadeia += $info.Nome
        $atual = [int]$info.Pai
        if ($atual -le 0) { break }
    }
    return $cadeia
}

$inicio = Get-Date
$fim = $inicio.AddSeconds($Segundos)
$amostras = 0
$amostrasComProcesso = 0
$conexoesRemotas = 0
$conexoesDoPoC = 0

Add-Content -Path $Saida -Encoding utf8 -Value @"

===============================================================================
SESSAO iniciada em $($inicio.ToString('yyyy-MM-dd HH:mm:ss')) - maquina: $env:COMPUTERNAME
Vigiados por NOME: $($processosPorNome -join ', ')
Vigiados por CAMINHO (roda de dentro da pasta do runtime, qualquer nome): $(if ($raizRuntime) { $raizRuntime } else { '(nenhum - runtime fixo nao encontrado)' })
Intervalo: ${IntervaloSegundos}s - duracao prevista: ${Segundos}s
Enderecos ignorados (nao sao egresso): $($naoEgresso -join ', ')
NOTA: os vigiados por CAMINHO apanham TODO processo (msedgewebview2.exe, Copilot, helpers - qualquer .exe)
que esteja rodando da NOSSA copia do runtime, nao so' uma lista fixa de nomes (achado 1 da revisao 1) - uma
versao futura do runtime que trouxesse mais binarios ainda seria vigiada sem precisar editar este script.
Os vigiados por NOME apanham TODO processo daquele nome NA MAQUINA, nao so' o do PoC (msedgewebview2.exe de
OUTRO aplicativo instalado, por exemplo, nao entraria por caminho mas poderia entrar aqui se o nome
coincidisse com a lista por nome - hoje nao coincide). Por isso cada conexao remota sai com a CADEIA DE
ANCESTRAIS do processo (so nomes, sem caminhos), e a conexao so conta contra o gate quando essa cadeia
passa pela arvore do PoC ($($donosDoPoC -join '/')).
===============================================================================
"@

while ((Get-Date) -lt $fim) {
    $amostras++
    $agora = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    # Tabela de TODOS os processos (pid -> nome + pai): e' dela que sai a cadeia de ancestrais.
    $tabela = @{}
    foreach ($w in Get-CimInstance Win32_Process -Property ProcessId, ParentProcessId, Name -ErrorAction SilentlyContinue) {
        $tabela[[int]$w.ProcessId] = @{ Nome = ($w.Name -replace '\.exe$', ''); Pai = $w.ParentProcessId }
    }

    # Vigiado por NOME (curto, so' o conversor e os hosts da suite) OU por CAMINHO (qualquer .exe rodando de
    # dentro da pasta do runtime embarcado) - nunca uma lista fixa de nomes de binarios do WebView2.
    $procs = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
        if ($processosPorNome -contains $_.ProcessName) { return $true }
        if (-not $raizRuntime) { return $false }
        try { $caminho = $_.MainModule.FileName } catch { return $false }
        return ($caminho -and $caminho.StartsWith($raizRuntime, [System.StringComparison]::OrdinalIgnoreCase))
    })
    if ($procs.Count -eq 0) {
        Add-Content -Path $Saida -Encoding utf8 -Value "[$agora] amostra $amostras : nenhum processo vigiado em execucao"
    }
    else {
        $amostrasComProcesso++
        $mapa = @{}
        foreach ($p in $procs) { $mapa[[int]$p.Id] = $p.ProcessName }
        $pids = @($mapa.Keys)
        $nomes = ($procs | Select-Object -ExpandProperty ProcessName | Sort-Object -Unique) -join '/'

        # Quantos dos processos vigiados sao da arvore do PoC (o resto e' outro app usando WebView2).
        $doPoC = 0
        foreach ($chave in $pids) {
            $cad = Get-CadeiaDeAncestrais -PidAlvo $chave -Tabela $tabela
            if (@($cad | Where-Object { $donosDoPoC -contains $_.ToLower() }).Count -gt 0) { $doPoC++ }
        }

        $conns = @(Get-NetTCPConnection -OwningProcess $pids -ErrorAction SilentlyContinue |
            Where-Object { $_.RemoteAddress -notin $naoEgresso })

        $cabecalho = "[$agora] amostra $amostras : $($procs.Count) processo(s) ($nomes), $doPoC do PoC"
        if ($conns.Count -eq 0) {
            Add-Content -Path $Saida -Encoding utf8 -Value "$cabecalho - ZERO conexoes remotas"
        }
        else {
            $conexoesRemotas += $conns.Count
            Add-Content -Path $Saida -Encoding utf8 -Value "$cabecalho - $($conns.Count) CONEXAO(OES) REMOTA(S) ***"
            foreach ($c in $conns) {
                $chave = [int]$c.OwningProcess
                $nome = if ($mapa.ContainsKey($chave)) { $mapa[$chave] } else { '?' }
                $cad = Get-CadeiaDeAncestrais -PidAlvo $chave -Tabela $tabela
                $ehDoPoC = @($cad | Where-Object { $donosDoPoC -contains $_.ToLower() }).Count -gt 0
                if ($ehDoPoC) { $conexoesDoPoC++; $dono = 'DO PoC' } else { $dono = 'de OUTRO app' }
                $det = "           $nome (pid $chave) -> $($c.RemoteAddress):$($c.RemotePort) [$($c.State)] local $($c.LocalAddress):$($c.LocalPort)"
                Add-Content -Path $Saida -Encoding utf8 -Value $det
                Add-Content -Path $Saida -Encoding utf8 -Value ("             dono: $dono - ascendencia: " + ($cad -join ' <- '))
            }
        }
    }

    $restante = ($fim - (Get-Date)).TotalSeconds
    if ($restante -le 0) { break }
    Start-Sleep -Seconds ([Math]::Min($IntervaloSegundos, [Math]::Ceiling($restante)))
}

if ($conexoesRemotas -eq 0) {
    $veredito = 'ZERO conexoes remotas em toda a sessao'
} elseif ($conexoesDoPoC -eq 0) {
    $veredito = "ZERO conexoes remotas DO PoC ($conexoesRemotas de outro(s) aplicativo(s) com WebView2 - ver a ascendencia de cada uma acima)"
} else {
    $veredito = "$conexoesDoPoC conexao(oes) remota(s) DO PoC, de $conexoesRemotas no total - VER ACIMA"
}

Add-Content -Path $Saida -Encoding utf8 -Value @"
-------------------------------------------------------------------------------
FIM em $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $amostras amostra(s), $amostrasComProcesso com algum processo vigiado vivo.
RESULTADO: $veredito
-------------------------------------------------------------------------------
"@

Write-Host "Amostras: $amostras (com processo vigiado vivo: $amostrasComProcesso). $veredito"
Write-Host "Saida: $Saida"
