<#
.SYNOPSIS
    Marco 0 / Task 7 - mede abertura do editor (a frio de verdade, 1a e 2a abertura), memoria do renderer e
    tempos do x2t, e gera poc/relatorios/medidas-<data>.md com a tabela do protocolo e o veredito de cada limiar.

.DESCRIPTION
    Para cada amostra (poc/amostras/*.* e poc/amostras/sinteticas/*.*), TRES aberturas, cada uma num processo
    NOVO do Host (mOffice.Sonda.exe <arquivo> --medir). O Host cronometra desde o OnStartup e desde o
    Process.StartTime, e grava uma linha em poc/relatorios/medidas-host.txt quando o EDITOR sinaliza documento
    pronto (asc_onDocumentContentReady, via ponte-desktop.js -> Avisar("pronto")).

      1. A FRIO (UDF limpo) - a pasta de dados de usuario do WebView2 (JanelaEditor.PastaDadosWebView) e'
         APAGADA antes, e o arquivo medido e' uma COPIA NOVA em pasta temporaria (caminho que o SO nunca viu).
         E' a condicao mais fria que da' para produzir sem reiniciar a maquina.
         RESSALVA HONESTA: isto NAO esvazia o cache de arquivos do SO. Os ~885 MB do motor copiado, as DLLs do
         WebView2 e o .NET continuam no standby cache do Windows. Uma abertura logo depois de ligar o PC pode
         ser mais lenta que este numero, e este script nao sabe medir isso - so' um reinicio mede.
      2. 1a ABERTURA (processo novo) - processo novo, UDF ja existente (o estado normal de quem abre o primeiro
         documento do dia com o app ja instalado ha' tempo).
      3. 2a ABERTURA (processo novo) - mesmo arquivo, logo em seguida. NAO e' "no mesmo processo": o PoC ainda
         nao abre um segundo documento na mesma janela, entao o que esta coluna mede e' o efeito do cache do SO
         entre duas execucoes, nao o de reaproveitar o WebView2 ja carregado.

    Com o documento aberto, tambem amostra o WorkingSet64 dos processos msedgewebview2 DESTA abertura (o maior
    = o renderer, para o veredito; a soma = o que pesa de verdade na maquina do usuario).

    Os tempos do x2t vem de poc/relatorios/medidas-x2t.txt (sintetico, escrito pela suite) e de
    poc/relatorios/medidas-x2t-grandes.txt (5/20/50 MB, escrito por AmostrasEMedidasTests). Sao arquivos de
    APPEND: o relatorio so' usa as linhas DE HOJE, e so' a ultima de cada arquivo medido - senao a tabela
    misturaria medidas de dias diferentes sem dizer.

.PARAMETER Amostras
    Caminhos especificos a medir. Sem isto, usa tudo o que houver em poc/amostras/ e poc/amostras/sinteticas/.

.PARAMETER GerarGrandes
    Roda os testes de medida do x2t (gera os arquivos de 5/20/50 MB - leva minutos e dezenas de MB em disco).

.PARAMETER EsperaSegundos
    Quanto esperar o documento ficar pronto antes de desistir (padrao 60).

.PARAMETER SemFrio
    Pula a abertura com UDF limpo (util para uma rodada rapida; a coluna sai como "sem medida", nunca como
    um veredito de uma condicao que nao foi exercitada).

.EXAMPLE
    pwsh tools/medir.ps1
    pwsh tools/medir.ps1 -GerarGrandes
#>
[CmdletBinding()]
param(
    [string[]] $Amostras,
    [switch]   $GerarGrandes,
    [int]      $EsperaSegundos = 60,
    [switch]   $SemFrio,

    # Task 7 (Marco 1 / Plano 2) - modo separado do resto do script: mede memoria do mTexto REAL (nao mais a
    # Sonda) com tres documentos abertos na MESMA janela (a mesma abertura encaminha "b" e "c" pelo pipe de
    # instancia unica). Nao participa do relatorio Markdown do Marco 0 acima - grava o proprio arquivo
    # (poc/relatorios/medidas-plano2-<data>.txt) e sai.
    [switch]   $Memoria3,
    [string[]] $Arquivos,
    [int]      $Segundos = 30
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$raiz        = Split-Path -Parent $PSScriptRoot
$relatorios  = Join-Path $raiz 'poc\relatorios'
$arqHost     = Join-Path $relatorios 'medidas-host.txt'
$arqX2t      = Join-Path $relatorios 'medidas-x2t.txt'
$arqX2tGrande= Join-Path $relatorios 'medidas-x2t-grandes.txt'
$exe         = Join-Path $raiz 'src\mOffice.Sonda\bin\Debug\net10.0-windows\mOffice.Sonda.exe'

# Pasta de dados de usuario do WebView2 - a MESMA que JanelaEditor.PastaDadosWebView monta. Apaga-la e' o
# unico "frio" que este script consegue produzir sem reiniciar a maquina (ver a ressalva no .DESCRIPTION).
$udfWebView = Join-Path $env:LOCALAPPDATA 'mOffice\PocWebView2'

# Limiares do gate (plano da Task 7; o usuario pode revisar no veredito). O plano define UM limiar de "a frio"
# (<= 6 s) e um de "a quente" (<= 3 s); aqui a abertura com UDF limpo e a 1a abertura respondem ao de 6 s (as
# duas sao processo novo) e a 2a abertura responde ao de 3 s.
$limiares = [ordered]@{
    'abertura_fria_ms'   = 6000
    'abertura_quente_ms' = 3000
    'x2t_abrir_20mb_ms'  = 10000
    'x2t_salvar_20mb_ms' = 15000
    'memoria_renderer_mb'= 600
}

function Escrever($texto) { Write-Host $texto }

# ============================================================================================================
# -Memoria3 (Task 7, criterio 6): tres documentos abertos no mTexto REAL, mesma janela/processo, soma da
# arvore inteira (mTexto.exe + msedgewebview2.exe desta abertura + x2t.exe filhos). Veredito do gate: soma
# <= 1,3 GB (1.331 MiB). Nao usa a Sonda nem o relatorio Markdown de cima -- sai antes de qualquer coisa que
# exija o Host compilado.
# ============================================================================================================
if ($Memoria3) {
    $exeMTexto     = Join-Path $raiz 'src\mTexto\bin\Release\net10.0-windows\mTexto.exe'
    $pastaAmostrasPlano2 = Join-Path $raiz 'poc\relatorios\amostras-plano2'
    $arqPlano2     = Join-Path $relatorios ('medidas-plano2-' + (Get-Date -Format 'yyyyMMdd') + '.txt')
    $limiarPlano2Mb = 1331   # 1,3 GB em MiB (1.3 * 1024)

    Escrever 'Compilando src/mTexto (Release)...'
    & dotnet build (Join-Path $raiz 'src\mTexto') -c Release --nologo | Out-Host
    if (-not (Test-Path $exeMTexto)) { throw "mTexto nao compilado: $exeMTexto" }

    if (-not $Arquivos) {
        # Sem -Arquivos: usa os tres DOCX sinteticos ricos gerados uma vez pela Task 7 (ver task-7-report.md)
        # e guardados em poc/relatorios/amostras-plano2/ -- AmostrasSinteticas.DocxRico vive no projeto de
        # TESTE (mOffice.Conversor.Tests) e nao ha' um "--gerar-amostras" de producao neste plano.
        if (Test-Path $pastaAmostrasPlano2) {
            $Arquivos = @(Get-ChildItem -Path $pastaAmostrasPlano2 -Filter '*.docx' -File |
                Sort-Object Name | Select-Object -First 3 -ExpandProperty FullName)
        }
        if (-not $Arquivos -or $Arquivos.Count -lt 3) {
            throw ("Preciso de 3 arquivos DOCX para medir (a/b/c) -- passe -Arquivos ou gere-os em " +
                   "$pastaAmostrasPlano2 (ver task-7-report.md, 'como as amostras foram geradas').")
        }
    }
    if ($Arquivos.Count -lt 3) { throw "-Arquivos precisa de pelo menos 3 caminhos (a, b, c); recebi $($Arquivos.Count)." }

    # A mesma pasta que AmbienteWebView2.PastaDadosUsuario("mtexto") monta: Pastas.Raiz (MOFFICE_DADOS_DIR,
    # senao %LOCALAPPDATA%\mOffice) + "mtexto" + "WebView2". Producao nunca define a variavel, entao o padrao
    # e' o caminho real que o usuario tera' na maquina dele.
    $dadosRaiz  = if ($env:MOFFICE_DADOS_DIR) { $env:MOFFICE_DADOS_DIR } else { Join-Path $env:LOCALAPPDATA 'mOffice' }
    $udfMTexto  = Join-Path $dadosRaiz 'mtexto\WebView2'

    # Pasta do runtime fixo (a versao mais alta achada em third_party/webview2, mesmo criterio de
    # AmbienteWebView2.CalcularPastaRuntimeFixo) -- so' o msedgewebview2.exe QUE MORA AQUI conta para a soma;
    # um WebView2 do sistema (Evergreen) de outro programa do usuario nunca deve ser contado nem morto.
    $pastaWebView2Base = Join-Path $raiz 'third_party\webview2'
    $pastaRuntimeFixo = $null
    if (Test-Path $pastaWebView2Base) {
        # So' pastas cujo NOME e' uma versao (ex.: "153.0.4234.48") -- "cache" (o cache de download do .cab,
        # tools/webview2/obter-runtime.ps1) e' uma subpasta irma, nao um runtime, e [version] lancaria nela.
        foreach ($v in Get-ChildItem -Path $pastaWebView2Base -Directory | Where-Object {
                     $verTeste = $null; [version]::TryParse($_.Name, [ref]$verTeste)
                 } | Sort-Object { [version]($_.Name) } -Descending) {
            $achado = Get-ChildItem -Path $v.FullName -Filter 'msedgewebview2.exe' -Recurse -Depth 1 -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($achado) { $pastaRuntimeFixo = $achado.DirectoryName; break }
        }
    }
    if (-not $pastaRuntimeFixo) { throw "Runtime fixo do WebView2 nao encontrado em $pastaWebView2Base -- rode tools/webview2/obter-runtime.ps1." }

    # Cadeia de pais por Win32_Process, mesmo desenho de WebView2-Desta-Abertura acima: se nao puder ser lida,
    # cai para o criterio "nao estava na foto de antes" (nunca mexe em processo alheio que ja existia).
    function Descende-De {
        param([hashtable] $Pais, [int] $IdInicial, [int] $PidAlvo)
        $id = $IdInicial
        for ($i = 0; $i -lt 12; $i++) {
            if (-not $Pais.ContainsKey($id)) { return $false }
            $pai = $Pais[$id]
            if ($pai -eq $PidAlvo) { return $true }
            if ($pai -le 0 -or $pai -eq $id) { return $false }
            $id = $pai
        }
        return $false
    }

    # Identifica a arvore NOVA desta rodada (mTexto + msedgewebview2 por user-data-dir + x2t descendente) a
    # partir das fotos de "antes". Usada duas vezes por rodada -- na medida (sucesso) e na limpeza (fix
    # revisao 1, ver Medir-UmaRodada abaixo) -- para as duas nunca divergirem: se a identificacao mudasse
    # entre medir e matar, a limpeza poderia deixar sobra ou, pior, matar um mTexto de verdade do usuario que
    # nao tem nada a ver com esta rodada.
    function Obter-ArvoreDaRodada {
        param([int[]] $MtextoAntes, [int[]] $WebviewAntes, [int[]] $X2tAntes)

        $pais = @{}
        try { foreach ($p in @(Get-WmiObject Win32_Process -ErrorAction Stop)) { $pais[[int]$p.ProcessId] = [int]$p.ParentProcessId } } catch { }

        $mtextoProcs = @(Get-Process -Name mTexto -ErrorAction SilentlyContinue | Where-Object { $MtextoAntes -notcontains $_.Id })

        $webviewCands = @(Get-Process -Name msedgewebview2 -ErrorAction SilentlyContinue | Where-Object { $WebviewAntes -notcontains $_.Id })
        $webviewProcs = @($webviewCands | Where-Object {
            $wmi = Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue
            $caminhoOk = $wmi -and $wmi.ExecutablePath -and $wmi.ExecutablePath.StartsWith($pastaRuntimeFixo, [StringComparison]::OrdinalIgnoreCase)
            $udfOk     = $wmi -and $wmi.CommandLine -and ($wmi.CommandLine -like "*--user-data-dir=*$udfMTexto*")
            $caminhoOk -and $udfOk
        })

        # x2t roda a CADA abertura (JanelaEditor.AbrirSessaoAsync chama X2tConversor.ParaBinAsync para gerar o
        # Editor.bin), nao so' em salvar/exportar -- mas e' rapido e normalmente ja terminou e saiu MUITO antes
        # dos 10 s + $Segundos desta amostragem. Um x2tProcs vazio aqui e' o ESPERADO em regime estacionario
        # (nada rodando de conversao no instante da foto), nao prova que a abertura nao passou pelo x2t.
        $x2tCands = @(Get-Process -Name x2t -ErrorAction SilentlyContinue | Where-Object { $X2tAntes -notcontains $_.Id })
        $x2tProcs = @($x2tCands | Where-Object {
            if ($pais.Count -eq 0) { return $true }   # sem arvore de processos: vale so' a foto de novos
            $idFilho = $_.Id
            $descendeDeAlgumMTexto = $false
            foreach ($m in $mtextoProcs) { if (Descende-De -Pais $pais -IdInicial $idFilho -PidAlvo $m.Id) { $descendeDeAlgumMTexto = $true; break } }
            $descendeDeAlgumMTexto
        })

        return [pscustomobject]@{ MTexto = $mtextoProcs; WebView2 = $webviewProcs; X2t = $x2tProcs }
    }

    function Medir-UmaRodada {
        param([int] $Numero)

        Escrever "=== Rodada $Numero ==="
        $mtextoAntes   = @(Get-Process -Name mTexto -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $webviewAntes  = @(Get-Process -Name msedgewebview2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $x2tAntes      = @(Get-Process -Name x2t -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

        try {
            Escrever "Abrindo $($Arquivos[0]) (processo novo)..."
            $proc = Start-Process -FilePath $exeMTexto -ArgumentList @("`"$($Arquivos[0])`"") -PassThru
            Start-Sleep -Seconds 10   # tempo do ambiente WebView2 subir e o documento carregar

            foreach ($a in $Arquivos[1..2]) {
                Escrever "Encaminhando $a a mesma instancia (pipe de instancia unica)..."
                # A SEGUNDA instancia so' encaminha o caminho pelo pipe e morre (InstanciaUnica.DeveContinuar) --
                # -Wait aqui espera esse processo transitorio sair, nao o mTexto principal.
                Start-Process -FilePath $exeMTexto -ArgumentList @("`"$a`"") -Wait
            }

            Escrever "Esperando $Segundos s para as tres janelas assentarem..."
            Start-Sleep -Seconds $Segundos

            $arvore      = Obter-ArvoreDaRodada -MtextoAntes $mtextoAntes -WebviewAntes $webviewAntes -X2tAntes $x2tAntes
            $mtextoProcs = $arvore.MTexto
            $webviewProcs= $arvore.WebView2
            $x2tProcs    = $arvore.X2t

            # Sem isto, um mTexto que morreu sozinho no meio da espera (medido na Task 7: um hiccup da sessao RDP
            # derrubou o processo entre o "doc:onready" e esta amostragem, sem excecao nem crash no Log de Eventos)
            # gravaria silenciosamente "soma=0 MiB" como se fosse uma medida valida -- o pior resultado possivel,
            # porque passa no limiar por acidente. Falha alto e claro em vez disso (o `finally` abaixo ainda
            # limpa o que sobrou antes do erro subir).
            if ($mtextoProcs.Count -eq 0) { throw "mTexto.exe nao esta rodando na hora de amostrar (rodada $Numero) -- caiu sozinho entre a abertura e a espera? Rode de novo." }

            $todos = @($mtextoProcs) + @($webviewProcs) + @($x2tProcs)
            $somaMb = if ($todos.Count -gt 0) { [math]::Round((($todos | Measure-Object WorkingSet64 -Sum).Sum) / 1MB, 1) } else { 0 }

            Escrever ("mTexto.exe: {0} | msedgewebview2.exe (desta abertura): {1} | x2t.exe filhos: {2} | soma: {3} MiB" -f `
                $mtextoProcs.Count, $webviewProcs.Count, $x2tProcs.Count, $somaMb)

            return [pscustomobject]@{ Numero = $Numero; MTexto = $mtextoProcs.Count; WebView2 = $webviewProcs.Count; X2t = $x2tProcs.Count; SomaMb = $somaMb }
        }
        finally {
            # Fix revisao 1: roda SEMPRE, sucesso ou excecao (inclusive o throw do guard acima, ou qualquer
            # falha no meio do Start-Process/Start-Sleep). ANTES desta correcao, um throw no meio da rodada
            # deixava a arvore parcial (mTexto + webview2 + x2t ja' abertos) viva: a PROXIMA rodada mediria
            # memoria ACUMULADA de duas aberturas, ou, se o throw fosse o ultimo da lista de 3, o script inteiro
            # terminava com o mTexto do bench ainda de pe' na maquina do usuario.
            #
            # Recalcula do zero em vez de reusar $mtextoProcs/$webviewProcs/$x2tProcs: se o throw veio de UM
            # Start-Process que falhou (ou de qualquer erro ANTES da linha "$arvore = Obter-ArvoreDaRodada..."),
            # essas variaveis nunca chegam a existir. Usa o MESMO criterio de identificacao da medida (Obter-
            # ArvoreDaRodada, que so' pega msedgewebview2 pelo user-data-dir do mtexto e mTexto/x2t "novos desde
            # a foto de antes") -- nunca um Stop-Process -Name generico, que mataria um mTexto de verdade do
            # usuario que por acaso estivesse aberto.
            $sobra = Obter-ArvoreDaRodada -MtextoAntes $mtextoAntes -WebviewAntes $webviewAntes -X2tAntes $x2tAntes
            @($sobra.MTexto) + @($sobra.WebView2) + @($sobra.X2t) | ForEach-Object {
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Milliseconds 500
        }
    }

    New-Item -ItemType Directory -Path $relatorios -Force | Out-Null

    # Foto de ANTES DE QUALQUER RODADA -- rede de seguranca do DRIVER (fix revisao 1), um segundo nivel alem do
    # `finally` de cada Medir-UmaRodada. Cada rodada ja limpa a propria arvore; isto cobre o caso (nao
    # observado, mas nao descartado por prudencia) de alguma sobra escapar de uma rodada anterior sem passar
    # pelo Stop-Process dela. So' pega o que sobrou do BENCH inteiro (novo desde ANTES deste script comecar) --
    # nunca um mTexto do usuario que ja estivesse rodando antes.
    $mtextoAntesGeral  = @(Get-Process -Name mTexto -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $webviewAntesGeral = @(Get-Process -Name msedgewebview2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $x2tAntesGeral     = @(Get-Process -Name x2t -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

    try {
        $rodadas = 1..3 | ForEach-Object { Medir-UmaRodada -Numero $_ }
    }
    finally {
        $sobraGeral = Obter-ArvoreDaRodada -MtextoAntes $mtextoAntesGeral -WebviewAntes $webviewAntesGeral -X2tAntes $x2tAntesGeral
        @($sobraGeral.MTexto) + @($sobraGeral.WebView2) + @($sobraGeral.X2t) | ForEach-Object {
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    }

    $minMb = ($rodadas | Measure-Object SomaMb -Minimum).Minimum
    $maxMb = ($rodadas | Measure-Object SomaMb -Maximum).Maximum
    $veredito = if ($maxMb -le $limiarPlano2Mb) { 'PASSOU' } else { '**REPROVADO**' }

    $linhasPlano2 = New-Object System.Collections.Generic.List[string]
    $linhasPlano2.Add("# Medidas de memoria - Marco 1 / Plano 2 - Task 7 (criterio 6) - $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    $linhasPlano2.Add('')
    $linhasPlano2.Add("Arquivos: $($Arquivos -join ', ')")
    $linhasPlano2.Add("Espera pos-abertura: $Segundos s. Limiar do gate: <= $limiarPlano2Mb MiB (1,3 GB).")
    $linhasPlano2.Add('')
    foreach ($r in $rodadas) {
        $linhasPlano2.Add("Rodada $($r.Numero): mTexto=$($r.MTexto) msedgewebview2=$($r.WebView2) x2t=$($r.X2t) soma=$($r.SomaMb) MiB")
    }
    $linhasPlano2.Add('')
    $linhasPlano2.Add("Minimo: $minMb MiB | Maximo: $maxMb MiB | Veredito (<= $limiarPlano2Mb MiB): $veredito")
    $linhasPlano2 | Out-File -LiteralPath $arqPlano2 -Append -Encoding utf8

    Escrever ''
    Escrever "Minimo: $minMb MiB | Maximo: $maxMb MiB | Veredito: $veredito"
    Escrever "Relatorio: $arqPlano2"
    return
}

if (-not (Test-Path $exe)) {
    throw "Host nao compilado: $exe. Rode 'dotnet build mOffice.slnx' antes."
}

if ($GerarGrandes) {
    Escrever 'Gerando amostras grandes e medindo o x2t (pode levar minutos)...'
    $env:MOFFICE_MEDIDAS = '1'
    & dotnet test (Join-Path $raiz 'tests\mOffice.Conversor.Tests\mOffice.Conversor.Tests.csproj') `
        --nologo --filter 'Categoria=Medida' | Out-Host
    Remove-Item Env:\MOFFICE_MEDIDAS -ErrorAction SilentlyContinue
}

if (-not $Amostras) {
    $pastas = @((Join-Path $raiz 'poc\amostras'), (Join-Path $raiz 'poc\amostras\sinteticas'))
    $Amostras = @()
    foreach ($p in $pastas) {
        if (Test-Path $p) {
            $Amostras += Get-ChildItem -Path $p -File |
                Where-Object { $_.Extension -in '.docx', '.odt', '.xlsx', '.ods', '.pptx', '.odp' } |
                Where-Object { $_.Name -notlike 'grande-*' } |   # os grandes sao medida de x2t, nao de abertura
                Select-Object -ExpandProperty FullName
        }
    }
}
if (-not $Amostras) { throw 'Nenhuma amostra encontrada em poc/amostras/ nem em poc/amostras/sinteticas/.' }

# --------------------------------------------------------------------------------------------------------
# Quais msedgewebview2 sao DESTA abertura (revisao final 2026-09-17). Antes o criterio era a HORA DE INICIO
# ("comecou depois do nosso Host"), e isso matava o WebView2 de QUALQUER outro aplicativo do usuario que
# tivesse subido no mesmo intervalo - o Teams, o Outlook novo, o proprio Windows. Agora e' o mesmo desenho do
# reabrir.ps1: fotografa os PIDs ANTES de subir o Host e so' considera os que nao existiam; alem disso exige
# que o processo DESCENDA do Host lancado, quando a arvore de processos puder ser lida.
#
# A cadeia de pais e' lida por Win32_Process (ParentProcessId). Se ela nao puder ser lida - permissao, WMI
# indisponivel - a funcao cai para o criterio "nao estava na foto", que ja' e' correto (nunca toca em processo
# alheio que ja' existia); o que se perde e' so' a precisao contra um WebView2 de terceiro que suba no meio.
# --------------------------------------------------------------------------------------------------------
function WebView2-Desta-Abertura {
    param([int[]] $PidsAntes, [int] $PidHost)

    $novos = @(Get-Process msedgewebview2 -ErrorAction SilentlyContinue | Where-Object { $PidsAntes -notcontains $_.Id })
    if ($novos.Count -eq 0) { return @() }

    $pais = @{}
    try {
        foreach ($p in @(Get-WmiObject Win32_Process -ErrorAction Stop)) { $pais[[int]$p.ProcessId] = [int]$p.ParentProcessId }
    } catch { return $novos }   # sem arvore de processos: vale so' a foto
    if ($pais.Count -eq 0) { return $novos }

    $desceDoHost = {
        param($idInicial)
        $id = $idInicial
        for ($i = 0; $i -lt 12; $i++) {          # teto: a arvore do WebView2 e' rasa e um ciclo travaria o laco
            if (-not $pais.ContainsKey($id)) { return $false }
            $pai = $pais[$id]
            if ($pai -eq $PidHost) { return $true }
            if ($pai -le 0 -or $pai -eq $id) { return $false }
            $id = $pai
        }
        return $false
    }
    $daArvore = @($novos | Where-Object { & $desceDoHost $_.Id })
    # Se a arvore nao explicou nenhum deles (o Host ja' saiu e o pai virou orfao), fica a foto - nunca o
    # criterio antigo de hora de inicio.
    if ($daArvore.Count -eq 0) { return $novos }
    return $daArvore
}

# --------------------------------------------------------------------------------------------------------
# Uma abertura: processo novo, espera a linha nova em medidas-host.txt, amostra a memoria, fecha.
# --------------------------------------------------------------------------------------------------------
function Medir-Abertura {
    param([string] $Arquivo, [switch] $LimparUdf)

    if ($LimparUdf) {
        # Apagar a pasta de dados do WebView2 forca reconstruir cache de codigo, de GPU e de rede do Chromium.
        # Se algum processo ainda segura um arquivo dela, a remocao falha - nesse caso a medida deixa de ser
        # "a frio" e o script TEM de dizer isso, nunca dar o numero como se fosse.
        try { if (Test-Path $udfWebView) { Remove-Item -LiteralPath $udfWebView -Recurse -Force -ErrorAction Stop } }
        catch { Escrever "  [aviso] nao deu para limpar a pasta de dados do WebView2: $($_.Exception.Message)"; return $null }
    }

    # @() obrigatorio: com uma unica linha Get-Content devolve string (sem .Count) e o StrictMode derruba.
    $antes = if (Test-Path $arqHost) { @(Get-Content -LiteralPath $arqHost).Count } else { 0 }
    # Start-Process NAO poe aspas nos argumentos: um caminho com espaco (ex.: "C:\Meus Projetos") chegaria partido
    # no Host e ele responderia "Extensao nao suportada: ''" num MessageBox modal (medido na Task 7).
    # Foto dos PIDs do WebView2 ANTES de subir o Host: so' o que aparecer depois pode ser encerrado por nos.
    $webviewAntes = @(Get-Process msedgewebview2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $proc  = Start-Process -FilePath $exe -ArgumentList @("`"$Arquivo`"", '--medir') -PassThru

    $prazo = (Get-Date).AddSeconds($EsperaSegundos)
    $linha = $null
    while ((Get-Date) -lt $prazo) {
        Start-Sleep -Milliseconds 200
        if (Test-Path $arqHost) {
            $todas = @(Get-Content -LiteralPath $arqHost)
            if ($todas.Count -gt $antes) { $linha = $todas[-1]; break }
        }
        if ($proc.HasExited) { break }
    }

    $memoriaMb = $null; $rendererMb = $null; $quantos = $null
    if ($linha) {
        # Amostrada COM o documento aberto (o Host so grava a linha depois do editor ficar pronto).
        Start-Sleep -Milliseconds 800   # deixa o pico de carga assentar
        # So os processos DESTA abertura: a maquina pode ter outros aplicativos usando WebView2, e somar todos
        # inflaria a medida com memoria que nao e' nossa (medido: 12 processos, boa parte de fora).
        $procs = @(WebView2-Desta-Abertura -PidsAntes $webviewAntes -PidHost $proc.Id)
        if ($procs.Count -gt 0) {
            # O limiar do protocolo fala do processo RENDERER; o WebView2 sobe varios processos (browser, GPU,
            # utilitarios, crashpad). Guardamos os dois: o maior (o renderer, que segura o documento) para o
            # veredito, e a SOMA como contexto, porque e' ela que pesa na maquina do usuario.
            $rendererMb = [math]::Round((($procs | Measure-Object WorkingSet64 -Maximum).Maximum) / 1MB, 1)
            $memoriaMb  = [math]::Round((($procs | Measure-Object WorkingSet64 -Sum).Sum) / 1MB, 1)
            $quantos    = $procs.Count
        }
    }

    if (-not $proc.HasExited) {
        $null = $proc.CloseMainWindow()
        if (-not $proc.WaitForExit(5000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    }
    # O WebView2 nao morre junto na hora; a proxima medida precisa da maquina limpa. So' os que NAO estavam na
    # foto de antes - nunca um WebView2 de outro aplicativo do usuario.
    Get-Process msedgewebview2 -ErrorAction SilentlyContinue |
        Where-Object { $webviewAntes -notcontains $_.Id } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500

    if (-not $linha) { return [pscustomobject]@{ ProntoMs = $null; ProcessoMs = $null; MemoriaMb = $null; RendererMb = $null; Processos = $null } }
    $pronto    = if ($linha -match 'pronto=(\d+)ms')   { [int]$Matches[1] } else { $null }
    $processo  = if ($linha -match 'processo=(-?\d+)ms') { [int]$Matches[1] } else { $null }
    return [pscustomobject]@{ ProntoMs = $pronto; ProcessoMs = $processo; MemoriaMb = $memoriaMb; RendererMb = $rendererMb; Processos = $quantos }
}

$vazio = [pscustomobject]@{ ProntoMs = $null; ProcessoMs = $null; MemoriaMb = $null; RendererMb = $null; Processos = $null }
$pastaFria = Join-Path $env:TEMP ('moffice-frio-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))

$resultados = @()
foreach ($a in $Amostras) {
    $nomeAmostra = [IO.Path]::GetFileName($a)

    $frioUdf = $vazio
    if (-not $SemFrio) {
        # Copia NOVA, em caminho que o SO nunca viu, para nao aproveitar o cache daquele arquivo. (O conteudo
        # do motor continua em cache - ver a ressalva do .DESCRIPTION; por isso a coluna se chama "UDF limpo",
        # e nao "maquina fria".)
        New-Item -ItemType Directory -Path $pastaFria -Force | Out-Null
        $copia = Join-Path $pastaFria $nomeAmostra
        Copy-Item -LiteralPath $a -Destination $copia -Force
        Escrever "Medindo $nomeAmostra (a frio: UDF do WebView2 limpo + copia nova)..."
        $r = Medir-Abertura -Arquivo $copia -LimparUdf
        if ($r) { $frioUdf = $r }
        Remove-Item -LiteralPath $copia -Force -ErrorAction SilentlyContinue
    }

    Escrever "Medindo $nomeAmostra (1a abertura, processo novo)..."
    $primeira = Medir-Abertura -Arquivo $a
    Escrever "Medindo $nomeAmostra (2a abertura, processo novo)..."
    $segunda = Medir-Abertura -Arquivo $a

    $resultados += [pscustomobject]@{
        Arquivo     = $nomeAmostra
        Editor      = switch ([IO.Path]::GetExtension($a)) { { $_ -in '.docx', '.odt' } { 'texto' } { $_ -in '.xlsx', '.ods' } { 'planilha' } default { 'apresentacao' } }
        FrioUdfMs   = $frioUdf.ProcessoMs
        FrioUdfAppMs= $frioUdf.ProntoMs
        PrimeiraMs  = $primeira.ProcessoMs
        PrimeiraAppMs = $primeira.ProntoMs
        SegundaMs   = $segunda.ProcessoMs
        SegundaAppMs= $segunda.ProntoMs
        RendererMb  = $primeira.RendererMb
        TotalMb     = $primeira.MemoriaMb
        Processos   = $primeira.Processos
    }
}
Remove-Item -LiteralPath $pastaFria -Recurse -Force -ErrorAction SilentlyContinue

# --------------------------------------------------------------------------------------------------------
# Relatorio
# --------------------------------------------------------------------------------------------------------
function Veredito($valor, $limiar) {
    if ($null -eq $valor) { return 'sem medida' }
    if ($valor -le $limiar) { return 'sim' } else { return '**nao**' }
}

$data     = Get-Date -Format 'yyyyMMdd'
$destino  = Join-Path $relatorios "medidas-$data.md"
$linhas   = New-Object System.Collections.Generic.List[string]
$linhas.Add("# Medidas do Marco 0 - $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
$linhas.Add('')
$linhas.Add('Gerado por `tools/medir.ps1`. Todos os tempos sao contados desde `Process.StartTime` do Host ate o')
$linhas.Add('editor sinalizar `asc_onDocumentContentReady`; `app=` e'' o mesmo instante contado a partir do `OnStartup`')
$linhas.Add('(sem o tempo de carga do .NET), para separar o custo do runtime do custo do editor.')
$linhas.Add('')
$linhas.Add('As tres colunas de abertura sao **tres condicoes diferentes**, e nenhuma delas e'' "maquina recem-ligada":')
$linhas.Add('')
$linhas.Add('- **A frio (UDF limpo)** - a pasta de dados de usuario do WebView2 e'' apagada antes, e o arquivo medido e''')
$linhas.Add('  uma copia nova em pasta temporaria (caminho que o SO nunca viu). E'' a condicao mais fria que da'' para')
$linhas.Add('  produzir sem reiniciar a maquina. **Ressalva:** isto NAO esvazia o cache de arquivos do SO - o motor')
$linhas.Add('  copiado, as DLLs do WebView2 e o .NET continuam no standby cache. Uma abertura logo depois de ligar o PC')
$linhas.Add('  pode ser mais lenta que este numero, e **esta condicao nao foi exercitada** nesta rodada.')
$linhas.Add('- **1a abertura (processo novo)** - processo novo, UDF ja existente: o primeiro documento do dia num app')
$linhas.Add('  instalado ha'' tempo.')
$linhas.Add('- **2a abertura (processo novo)** - mesmo arquivo, logo em seguida. NAO e'' "no mesmo processo": o PoC ainda')
$linhas.Add('  nao abre um segundo documento na mesma janela.')
$linhas.Add('')
$linhas.Add('## Abertura e memoria')
$linhas.Add('')
$linhas.Add('| Amostra | Editor | A frio, UDF limpo (ms) | dentro do limiar (<= 6000)? | 1a abertura (ms) | dentro do limiar (<= 6000)? | 2a abertura (ms) | dentro do limiar (<= 3000)? | Renderer (MB) | dentro do limiar (<= 600)? | Todos os msedgewebview2 (MB / processos) |')
$linhas.Add('|---|---|---|---|---|---|---|---|---|---|---|')
foreach ($r in $resultados) {
    $linhas.Add(("| {0} | {1} | {2} (app={3}) | {4} | {5} (app={6}) | {7} | {8} (app={9}) | {10} | {11} | {12} | {13} MB / {14} |" -f `
        $r.Arquivo, $r.Editor,
        $r.FrioUdfMs, $r.FrioUdfAppMs, (Veredito $r.FrioUdfMs $limiares.abertura_fria_ms),
        $r.PrimeiraMs, $r.PrimeiraAppMs, (Veredito $r.PrimeiraMs $limiares.abertura_fria_ms),
        $r.SegundaMs, $r.SegundaAppMs, (Veredito $r.SegundaMs $limiares.abertura_quente_ms),
        $r.RendererMb, (Veredito $r.RendererMb $limiares.memoria_renderer_mb),
        $r.TotalMb, $r.Processos))
}
$linhas.Add('')
$linhas.Add('## x2t (abrir = ParaBin, salvar = DeBin)')
$linhas.Add('')
$linhas.Add('Tres familias de arquivo grande, porque "20 MB" nao diz nada sozinho - o que custa ao x2t e'' o que tem')
$linhas.Add('dentro: `grande-NNmb.docx` = peticao com imagens (o caso do limiar; o x2t basicamente copia a midia),')
$linhas.Add('`grande-NNmb-texto.docx` = so paragrafos (o extremo caro: ~5 mil paginas num arquivo de 20 MB) e')
$linhas.Add('`grande-NNmb.xlsx` = linhas de verdade, que e'' como planilha grande cresce. Geradas por')
$linhas.Add('`AmostrasSinteticas`; `sintetico.docx` e'' a medida pequena que a suite normal ja grava.')
$linhas.Add('')
$linhas.Add('Os dois arquivos de origem (`medidas-x2t.txt`, `medidas-x2t-grandes.txt`) sao de APPEND: acumulam uma linha')
$linhas.Add('por medida, de todas as rodadas de todos os dias. A tabela abaixo usa **so as linhas de hoje** e, dentro')
$linhas.Add('de hoje, **so a ultima** de cada arquivo medido - senao ela misturaria dias sem dizer qual e qual.')
$linhas.Add('')
$linhas.Add('| Arquivo | Tamanho | abrir (ms) | salvar (ms) | limiar aplicavel | dentro do limiar? |')
$linhas.Add('|---|---|---|---|---|---|')
$hoje = Get-Date -Format 'yyyy-MM-dd'
$doDia = [ordered]@{}          # nome do arquivo medido -> ultima linha DE HOJE
$descartadas = 0
foreach ($arq in @($arqX2t, $arqX2tGrande)) {
    if (-not (Test-Path $arq)) { continue }
    foreach ($l in Get-Content -LiteralPath $arq) {
        if ($l -notmatch 'abrir=(\d+)ms\s+salvar=(\d+)ms') { continue }
        # O Host/os testes gravam "<ISO 8601> <arquivo> ...": a data e' o prefixo da primeira coluna.
        if ($l -notlike "$hoje*") { $descartadas++; continue }
        $doDia[($l -split '\s+')[1]] = $l   # a ultima de hoje vence
    }
}
foreach ($nome in $doDia.Keys) {
    $l = $doDia[$nome]
    $null = $l -match 'abrir=(\d+)ms\s+salvar=(\d+)ms'
    $abrir = [int]$Matches[1]; $salvar = [int]$Matches[2]
    # Exige o separador decimal: sem isso o '20mb' do NOME do arquivo casa antes do tamanho medido.
    $tamanho = if ($l -match '(\d+[.,]\d+MB)') { $Matches[1] } else { 'sintetico (pequeno)' }
    if ($nome -like '*20mb*') {
        $v = "abrir: $(Veredito $abrir $limiares.x2t_abrir_20mb_ms) / salvar: $(Veredito $salvar $limiares.x2t_salvar_20mb_ms)"
        $lim = 'abrir <= 10000, salvar <= 15000'
    } else {
        $v = 'n/a (o limiar e'' definido para 20 MB)'
        $lim = '-'
    }
    $linhas.Add("| $nome | $tamanho | $abrir | $salvar | $lim | $v |")
}
if ($doDia.Count -eq 0) {
    $linhas.Add('| *(nenhuma medida de x2t hoje)* | - | - | - | - | rode com `-GerarGrandes`, ou rode a suite |')
}
$linhas.Add('')
$linhas.Add("Linhas de outros dias ignoradas nesta tabela: **$descartadas**.")
$linhas.Add('')
$linhas.Add('## Limiares (plano da Task 7)')
$linhas.Add('')
foreach ($k in $limiares.Keys) { $linhas.Add("- ``$k`` <= $($limiares[$k])") }

# UTF-8 SEM BOM: o Set-Content -Encoding utf8 do Windows PowerShell 5.1 poe BOM, que aparece como lixo no
# inicio do markdown em varias ferramentas.
[IO.File]::WriteAllLines($destino, $linhas, (New-Object Text.UTF8Encoding($false)))
Escrever ''
Escrever "Relatorio: $destino"
$resultados | Format-Table | Out-Host
