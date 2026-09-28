<#
.SYNOPSIS
    Marco 0 / Task 7 - reabre, de forma reproduzivel, cada arquivo GRAVADO PELO PoC nos conferidores que tem
    linha de comando, e gera poc/relatorios/reabertura-<data>.md com uma linha por artefato.

.DESCRIPTION
    As colunas "Reabre ..." da matriz do protocolo eram, ate a revisao 1 da Task 7, numeros anotados a mao
    numa sessao que ninguem consegue repetir. Este script e' o driver que faltava: mesma lista de pares que
    tools/fidelidade.ps1 usa, e cada celula da matriz sai de um comando que qualquer um roda de novo.

    Passo 0 - PRODUZIR o artefato: fidelidade.ps1 roda o round-trip pelo editor (Host <origem> --salvar
              <destino>) e deixa <base>-salvo.<ext> na pasta de trabalho. E' esse arquivo, o que o PoC gravou,
              que os tres conferidores abaixo tentam abrir.

    (a) LibreOffice headless - soffice --headless --norestore --convert-to pdf --outdir <tmp> <arquivo>.
        "Abre" = processo saiu com codigo 0 E o PDF existe com mais de 0 bytes. Converter exige ter aberto.
        Cada artefato tem a SUA pasta de saida: rico-salvo.docx e rico-salvo.odt gerariam o mesmo rico-salvo.pdf.

    (b) Microsoft Office por automacao COM - Word (DOCX/ODT), Excel (XLSX/ODS), PowerPoint (PPTX/ODP): abre em
        somente-leitura, le uma grandeza do documento (caracteres / planilhas + A1 / slides), fecha e sai.
        Nunca deixa processo para tras: o que for aberto aqui e' fechado aqui, e o script SO encerra o
        aplicativo se ele NAO estava rodando antes (senao um Quit() fecharia o Word do usuario).

    (c) O proprio PoC (2a geracao) - o arquivo salvo e' reaberto no PoC e salvo OUTRA vez, e a 2a geracao e'
        comparada com a 1a pelo mesmo comparador de fidelidade.ps1 (contagens estruturais + texto integral em
        OOXML; LibreOffice/content.xml em ODF). Feito chamando fidelidade.ps1 com os arquivos salvos como
        amostras - a logica de comparacao nao e' duplicada aqui.

    (d) ONLYOFFICE Desktop Editors OFICIAL - NAO tem interface de linha de comando que sirva de prova (o
        executavel nao expoe "abrir e sair com codigo"). A coluna sai como "(HUMAN) pendente": e' uma
        verificacao para o usuario fazer na mao, e este script nao inventa um resultado para ela.

    (e) ODF + -Pares: alem do SALVO (item a), tambem renderiza o ORIGINAL em PDF pelo LibreOffice - antes so'
        o salvo virava PDF, entao nao dava pra comparar visualmente original x salvo (achado da revisao
        2026-09-18: a diferenca de texto do ODT/ODP reais precisa de conferencia visual, nao so' contagem de
        linhas). PDFs vao para `<Saida>\originais\<rotulo>-<ext>\<rotulo>.pdf` - a pasta DOS PARES, nunca
        `poc/relatorios/` - com o ORIGINAL copiado antes para um nome NEUTRO (o rotulo): o LibreOffice nomeia
        o PDF a partir do arquivo de entrada, entao converter a Origem direto vazaria o nome real em disco.

.PARAMETER Amostras
    Origens especificas. Sem isto (e sem -Pares), usa poc/amostras/ e poc/amostras/sinteticas/ (menos os
    `grande-*` e os `*-salvo.*`) - a MESMA lista de fidelidade.ps1.

.PARAMETER Pares
    Lista de pares { Origem; Rotulo } para amostras CONFIDENCIAIS - repassado para a 1a chamada de
    fidelidade.ps1 (ver o PARAMETER la). A partir da 1a geracao (ja' nomeada `<Rotulo>-salvo.<ext>`) tudo o
    resto do script (2a geracao, LibreOffice, MS Office, relatorio) usa esses nomes JA' redigidos por rotulo -
    nenhum nome real do arquivo entra no relatorio. Para ODT/ODS/ODP tambem dispara o item (e) acima.

.PARAMETER Saida
    Pasta de trabalho (padrao: %TEMP%\moffice-reabertura). Nada dela entra no git.

.PARAMETER Relatorio
    Markdown de saida. Padrao: poc/relatorios/reabertura-<yyyyMMdd>.md.

.PARAMETER SemMsOffice
    Pula a automacao COM (maquina sem Microsoft Office). As celulas saem "sem medida", nunca "sim".

.PARAMETER AutoTeste
    Roda o self-teste de `ConverterLinhaTabela` (o parser da tabela do fidelidade.ps1, achado desta revisao:
    quebrava quando o Detalhe continha "|") e sai - nao faz round-trip nenhum, nao precisa do Host compilado.

.EXAMPLE
    pwsh tools/reabrir.ps1
    pwsh tools/reabrir.ps1 -SemMsOffice
    pwsh tools/reabrir.ps1 -AutoTeste
#>
[CmdletBinding()]
param(
    [string[]] $Amostras,
    [object[]] $Pares,
    [string]   $Saida = (Join-Path $env:TEMP 'moffice-reabertura'),
    [string]   $Relatorio,
    [int]      $EsperaSegundos = 90,
    [switch]   $SemMsOffice,
    [switch]   $AutoTeste
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------------------------------------
# Parser da tabela markdown de fidelidade.ps1: "| <amostra> | igual|diferente|FALHOU... | <metodo> | <detalhe> |".
# O Detalhe pode conter "|" (fidelidade.ps1 junta varios trechos de diferenca com " | " - ver Diferenca-De-Texto
# e o comparador ODF), entao NAO da para usar '\|' como delimitador em toda a linha: uma regex de 4 grupos que
# supoe "o ultimo grupo nao tem pipe" corta o Detalhe no primeiro "|" dele e silenciosamente perde o resto (achado
# desta revisao). Em vez disso, o split usa LIMITE 5: os primeiros 4 pipes dividem Amostra/Veredito/Metodo, e o
# QUINTO elemento fica com o resto da linha inteiro, sem novo split - e' so' preciso tirar o "|" final da tabela.
# -AutoTeste exercita esta funcao isoladamente (linha sintetica com "|" no Detalhe) antes de qualquer round-trip.
# ---------------------------------------------------------------------------------------------------------
function ConverterLinhaTabela {
    param([string] $Linha)
    if ($Linha -notmatch '^\|') { return $null }
    $partes = $Linha -split '\|', 5
    if ($partes.Count -lt 5) { return $null }
    $amostra  = $partes[1].Trim()
    $veredito = $partes[2].Trim()
    if ($veredito -notmatch '^(igual|diferente|FALHOU)') { return $null }   # descarta cabecalho/separador da tabela
    $metodo  = $partes[3].Trim()
    $detalhe = $partes[4].Trim()
    if ($detalhe.EndsWith('|')) { $detalhe = $detalhe.Substring(0, $detalhe.Length - 1).Trim() }
    return [pscustomobject]@{ Amostra = $amostra; Veredito = $veredito; Metodo = $metodo; Detalhe = $detalhe }
}

if ($AutoTeste) {
    # Linha sintetica com "|" no Detalhe (mesmo formato que Diferenca-De-Texto/comparador ODF produzem em
    # fidelidade.ps1) - a regex antiga cortava o Detalhe no primeiro "|" e perdia tudo depois dele.
    $linhaTeste = '| real-odt | diferente | LibreOffice txt | texto por LibreOffice txt : 3 linha(s) diferente(s) - => trecho A | trecho B | <= trecho C |'
    $r = ConverterLinhaTabela $linhaTeste
    $falhas = @()
    if ($null -eq $r) { $falhas += 'linha nao reconhecida (retornou $null)' }
    else {
        if ($r.Amostra  -ne 'real-odt')          { $falhas += "Amostra = '$($r.Amostra)' (esperado 'real-odt')" }
        if ($r.Veredito -ne 'diferente')         { $falhas += "Veredito = '$($r.Veredito)' (esperado 'diferente')" }
        if ($r.Metodo   -ne 'LibreOffice txt')   { $falhas += "Metodo = '$($r.Metodo)' (esperado 'LibreOffice txt')" }
        $detalheEsperado = 'texto por LibreOffice txt : 3 linha(s) diferente(s) - => trecho A | trecho B | <= trecho C'
        if ($r.Detalhe  -ne $detalheEsperado)    { $falhas += "Detalhe = '$($r.Detalhe)' (esperado '$detalheEsperado')" }
    }
    # Linha de cabecalho/separador nao deve casar.
    if ($null -ne (ConverterLinhaTabela '| Amostra | Resultado | Metodo | O que mudou |'))  { $falhas += 'cabecalho foi reconhecido como linha de dados' }
    if ($null -ne (ConverterLinhaTabela '|---|---|---|---|'))                                { $falhas += 'separador foi reconhecido como linha de dados' }

    if ($falhas) {
        Write-Host "AutoTeste FALHOU:" -ForegroundColor Red
        $falhas | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
        exit 1
    }
    Write-Host "AutoTeste OK: ConverterLinhaTabela preserva '|' dentro do Detalhe -> '$($r.Detalhe)'"
    exit 0
}

$raiz        = Split-Path -Parent $PSScriptRoot
$relatorios  = Join-Path $raiz 'poc\relatorios'
$fidelidade  = Join-Path $PSScriptRoot 'fidelidade.ps1'
$exe         = Join-Path $raiz 'src\mOffice.Sonda\bin\Debug\net10.0-windows\mOffice.Sonda.exe'
$soffice     = 'C:\Program Files\LibreOffice\program\soffice.exe'

if (-not (Test-Path $exe))        { throw "Host nao compilado: $exe. Rode 'dotnet build mOffice.slnx' antes." }
if (-not (Test-Path $fidelidade)) { throw "tools/fidelidade.ps1 nao encontrado." }
if (-not $Relatorio) { $Relatorio = Join-Path $relatorios ("reabertura-" + (Get-Date -Format 'yyyyMMdd') + ".md") }

if (-not $Amostras -and -not $Pares) {
    $pastas = @((Join-Path $raiz 'poc\amostras'), (Join-Path $raiz 'poc\amostras\sinteticas'))
    $Amostras = @()
    foreach ($p in $pastas) {
        if (Test-Path $p) {
            $Amostras += Get-ChildItem -Path $p -File |
                Where-Object { $_.Extension -in '.docx', '.odt', '.xlsx', '.ods', '.pptx', '.odp' } |
                Where-Object { $_.Name -notlike 'grande-*' -and $_.Name -notlike '*-salvo.*' } |
                Select-Object -ExpandProperty FullName
        }
    }
}
if (-not $Amostras -and -not $Pares) { throw 'Nenhuma amostra encontrada em poc/amostras/ nem em poc/amostras/sinteticas/.' }

$g1 = Join-Path $Saida 'geracao1'
$g2 = Join-Path $Saida 'geracao2'
$lo = Join-Path $Saida 'libreoffice'
New-Item -ItemType Directory -Path $g1, $g2, $lo -Force | Out-Null

# ---------------------------------------------------------------------------------------------------------
# Passo 0 + (c): fidelidade.ps1 duas vezes.
#   1a chamada: origem -> 1a geracao (o "salvo pelo PoC" que todo mundo tenta reabrir).
#   2a chamada: 1a geracao -> 2a geracao, comparando as duas. E' exatamente "reabre no proprio PoC".
# ---------------------------------------------------------------------------------------------------------
Write-Host '== Gerando os arquivos salvos pelo PoC (1a geracao) =='
$md1 = Join-Path $Saida 'geracao1.md'
if ($Pares) {
    & $fidelidade -Pares $Pares -Saida $g1 -Markdown $md1 -EsperaSegundos $EsperaSegundos | Out-Host
} else {
    & $fidelidade -Amostras $Amostras -Saida $g1 -Markdown $md1 -EsperaSegundos $EsperaSegundos | Out-Host
}

$salvos = @(Get-ChildItem -Path $g1 -File |
            Where-Object { $_.Extension -in '.docx', '.odt', '.xlsx', '.ods', '.pptx', '.odp' } |
            Sort-Object Name | Select-Object -ExpandProperty FullName)
if (-not $salvos) { throw "O PoC nao gravou nenhum arquivo em $g1 - veja a saida acima." }

Write-Host ''
Write-Host '== Reabrindo no proprio PoC e salvando de novo (2a geracao) =='
$md2 = Join-Path $Saida 'geracao2.md'
& $fidelidade -Amostras $salvos -Saida $g2 -Markdown $md2 -EsperaSegundos $EsperaSegundos | Out-Host

# Tabela do fidelidade.ps1: "| <amostra> | igual|diferente | <metodo> | <detalhe> |" (4 colunas desde que a
# coluna Metodo entrou no relatorio - achado da revisao 2026-09-18: a regex aqui esperava as 3 colunas ANTIGAS,
# sem Metodo, entao NUNCA casava com a linha real e $poc ficava sempre vazio - "sem medida" em TODAS as linhas,
# em silencio, sem nenhum erro. So' nao apareceu antes porque a rodada de 2026-09-17 comparava contra o
# cabecalho fixo do relatorio e ninguem tinha reparado que a coluna "Reabre no PoC" saia de um parser que nunca
# casava. Corrigido usando ConverterLinhaTabela (definida no topo do script) em vez de uma regex inline: aquela
# primeira correcao trocou para uma regex de 4 grupos que ainda quebrava quando o Detalhe continha "|" (achado
# desta revisao) - ConverterLinhaTabela usa split com limite em vez de regex gulosa e tem -AutoTeste cobrindo o
# caso.
$poc = @{}
foreach ($l in Get-Content -LiteralPath $md2) {
    $linha = ConverterLinhaTabela $l
    if ($linha) { $poc[$linha.Amostra] = [pscustomobject]@{ Veredito = $linha.Veredito; Metodo = $linha.Metodo; Detalhe = $linha.Detalhe } }
}

# ---------------------------------------------------------------------------------------------------------
# (a) LibreOffice headless
# ---------------------------------------------------------------------------------------------------------
function Reabrir-LibreOffice {
    param([string] $Arquivo, [string] $PastaBase = $lo)
    if (-not (Test-Path $soffice)) { return 'sem medida (LibreOffice nao instalado)' }
    # Uma pasta por artefato: rico-salvo.docx e rico-salvo.odt produziriam o MESMO rico-salvo.pdf.
    $destino = Join-Path $PastaBase ([IO.Path]::GetFileNameWithoutExtension($Arquivo) + '-' + [IO.Path]::GetExtension($Arquivo).TrimStart('.'))
    New-Item -ItemType Directory -Path $destino -Force | Out-Null
    $pdf = Join-Path $destino ([IO.Path]::GetFileNameWithoutExtension($Arquivo) + '.pdf')
    if (Test-Path $pdf) { Remove-Item -LiteralPath $pdf -Force }
    # Start-Process (e nao o operador &): o soffice escreve avisos no stderr e, no Windows PowerShell 5.1,
    # isso vira NativeCommandError - com $ErrorActionPreference='Stop' o script inteiro morria num aviso.
    $p = Start-Process -FilePath $soffice -PassThru -WindowStyle Hidden `
        -ArgumentList @('--headless', '--norestore', '--convert-to', 'pdf', '--outdir', "`"$destino`"", "`"$Arquivo`"") `
        -RedirectStandardError (Join-Path $destino 'stderr.txt') -RedirectStandardOutput (Join-Path $destino 'stdout.txt')
    # Ler .Handle AGORA e' obrigatorio: sem isso o objeto devolvido por Start-Process nao guarda o handle do
    # processo e .ExitCode volta VAZIO depois que ele sai - e o teste "saiu com codigo 0" daria sempre "nao",
    # mesmo com o PDF gerado ao lado (foi o que aconteceu na primeira rodada deste script).
    $null = $p.Handle
    if (-not $p.WaitForExit(180000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; return 'nao (timeout de 180 s)' }
    if ($p.ExitCode -ne 0) { return "nao (soffice saiu com codigo $($p.ExitCode))" }
    if (-not (Test-Path $pdf)) { return 'nao (nenhum PDF gerado)' }
    $bytes = (Get-Item -LiteralPath $pdf).Length
    if ($bytes -le 0) { return 'nao (PDF de 0 bytes)' }
    return "sim (PDF de $bytes bytes)"
}

# ---------------------------------------------------------------------------------------------------------
# (b) Microsoft Office por COM.
#   - O aplicativo so' e' encerrado se NAO estava aberto antes: em COM, New-Object anexa a uma instancia ja
#     rodando, e um Quit() fecharia o Word/Excel/PowerPoint DO USUARIO junto com o nosso documento.
#   - PowerPoint nao aceita Visible = $false (a propriedade lanca); a forma de nao mostrar janela e' abrir com
#     WithWindow = msoFalse. E o equivalente de DisplayAlerts = 0 la' e' ppAlertsNone = 1.
#   - No fim, so' processos que NAO existiam antes desta chamada podem ser encerrados a forca.
# ---------------------------------------------------------------------------------------------------------
function Reabrir-MsOffice {
    param([string] $Arquivo)
    if ($SemMsOffice) { return 'sem medida (-SemMsOffice)' }

    $ext = [IO.Path]::GetExtension($Arquivo).ToLowerInvariant()
    $progId   = switch ($ext) { { $_ -in '.docx', '.odt' } { 'Word.Application' } { $_ -in '.xlsx', '.ods' } { 'Excel.Application' } default { 'PowerPoint.Application' } }
    $processo = switch ($progId) { 'Word.Application' { 'WINWORD' } 'Excel.Application' { 'EXCEL' } default { 'POWERPNT' } }

    $pidsAntes  = @(Get-Process -Name $processo -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $jaEstavaAberto = $pidsAntes.Count -gt 0

    $app = $null; $doc = $null; $resposta = 'nao (falhou)'
    try {
        try { $app = New-Object -ComObject $progId }
        catch { return "sem medida ($progId indisponivel: $($_.Exception.Message))" }

        if ($progId -eq 'PowerPoint.Application') {
            $app.DisplayAlerts = 1                                   # ppAlertsNone
            $doc = $app.Presentations.Open($Arquivo, $true, $false, $false)   # ReadOnly, Untitled, WithWindow=false
            $resposta = "sim ($($doc.Slides.Count) slides)"
        }
        elseif ($progId -eq 'Excel.Application') {
            $app.Visible = $false; $app.DisplayAlerts = $false
            $doc = $app.Workbooks.Open($Arquivo, 0, $true)           # UpdateLinks=0, ReadOnly=true
            $a1 = [string]$doc.Worksheets.Item(1).Range('A1').Text
            $resposta = "sim ($($doc.Worksheets.Count) planilhas, A1 = '$a1')"
        }
        else {
            $app.Visible = $false; $app.DisplayAlerts = 0
            $doc = $app.Documents.Open($Arquivo, $false, $true, $false)  # ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false
            $resposta = "sim ($($doc.Content.Text.Length) caracteres)"
        }
    }
    catch { $resposta = "nao ($($_.Exception.Message))" }
    finally {
        if ($doc) {
            try { if ($progId -eq 'PowerPoint.Application') { $doc.Close() } elseif ($progId -eq 'Excel.Application') { $doc.Close($false) } else { $doc.Close(0) } } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
        }
        if ($app) {
            # So' encerra o aplicativo se fomos nos que o subimos.
            if (-not $jaEstavaAberto) { try { $app.Quit() } catch { } }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # Rede de seguranca: so' o que NAO existia antes desta chamada.
        if (-not $jaEstavaAberto) {
            Get-Process -Name $processo -ErrorAction SilentlyContinue |
                Where-Object { $pidsAntes -notcontains $_.Id } |
                Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }
    if ($jaEstavaAberto) { $resposta += ' [instancia do usuario reaproveitada: nao foi encerrada]' }
    return $resposta
}

# ---------------------------------------------------------------------------------------------------------

# -Pares + ODF (achado 2 da revisao 2026-09-18): so' o SALVO era renderizado em PDF - nao dava pra comparar
# visualmente original x salvo, so por contagem/texto. Mapa rotulo -> origem para tambem renderizar o ORIGINAL
# das amostras ODF confidenciais. PDFs vao para $Saida\originais (a pasta DOS PARES que o chamador passou -
# NUNCA poc/relatorios, que e' commitado). Nao mistura com o fluxo padrao (sem -Pares, isto nunca dispara).
$origensPorRotulo = @{}
if ($Pares) {
    foreach ($par in $Pares) { $origensPorRotulo[[string]$par.Rotulo] = [string]$par.Origem }
}
$pastaOriginais = Join-Path $Saida 'originais'

Write-Host ''
Write-Host '== Reabrindo nos conferidores externos =='
$linhasResultado = @()
$originaisRenderizados = @()
foreach ($arquivo in $salvos) {
    $nome = [IO.Path]::GetFileName($arquivo)
    Write-Host "Reabrindo $nome ..."
    $rLo = Reabrir-LibreOffice -Arquivo $arquivo
    Write-Host "  LibreOffice : $rLo"
    $rMs = Reabrir-MsOffice -Arquivo $arquivo
    Write-Host "  MS Office   : $rMs"
    if ($poc.ContainsKey($nome)) {
        $p = $poc[$nome]
        $rPoc = if ($p.Veredito -eq 'igual') { 'sim (2a geracao identica a 1a)' } else { "nao ($($p.Veredito): $($p.Detalhe))" }
    } else {
        $rPoc = 'sem medida (fidelidade.ps1 nao reportou este arquivo)'
    }
    Write-Host "  PoC (2a ger): $rPoc"

    $extArquivo = [IO.Path]::GetExtension($arquivo).ToLowerInvariant()
    if ($Pares -and $extArquivo -in '.odt', '.ods', '.odp') {
        $rotulo = [IO.Path]::GetFileNameWithoutExtension($arquivo) -replace '-salvo$', ''
        if ($origensPorRotulo.ContainsKey($rotulo)) {
            # Copia para um nome NEUTRO (rotulo) antes de chamar o LibreOffice: soffice nomeia o PDF de saida a
            # partir do nome do arquivo de ENTRADA, entao converter a Origem diretamente vazaria o nome real do
            # arquivo real do escritorio no nome da pasta/PDF - mesma disciplina do resto do script (so' o
            # rotulo aparece em disco a partir daqui).
            $origemReal  = $origensPorRotulo[$rotulo]
            $copiaRotulo = Join-Path $pastaOriginais ($rotulo + [IO.Path]::GetExtension($origemReal))
            New-Item -ItemType Directory -Path $pastaOriginais -Force | Out-Null
            Copy-Item -LiteralPath $origemReal -Destination $copiaRotulo -Force
            $rLoOriginal = Reabrir-LibreOffice -Arquivo $copiaRotulo -PastaBase $pastaOriginais
            Write-Host "  LibreOffice (ORIGINAL, para comparacao visual) : $rLoOriginal"
            $originaisRenderizados += [pscustomobject]@{ Rotulo = $rotulo; Resultado = $rLoOriginal }
        }
    }

    $linhasResultado += [pscustomobject]@{
        Artefato    = $nome
        LibreOffice = $rLo
        MsOffice    = $rMs
        Poc         = $rPoc
        OnlyOffice  = '(HUMAN) pendente'
    }
}

# ---------------------------------------------------------------------------------------------------------
# Relatorio. So' NOMES de arquivo entram - nenhum caminho pessoal, porque este arquivo e' commitado.
# ---------------------------------------------------------------------------------------------------------
$versaoLo = if (Test-Path $soffice) { (Get-Item $soffice).VersionInfo.ProductVersion } else { 'nao instalado' }
$winword  = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
$versaoMs = if (Test-Path $winword) { (Get-Item $winword).VersionInfo.ProductVersion } else { 'nao encontrado' }

$md = New-Object System.Collections.Generic.List[string]
$md.Add("# Reabertura dos arquivos salvos pelo PoC - $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
$md.Add('')
$md.Add('Gerado por `tools/reabrir.ps1` (re-executavel). Cada linha e'' um arquivo GRAVADO PELO PoC pelo')
$md.Add('round-trip do editor (`Host <origem> --salvar <destino>`, o mesmo de `tools/fidelidade.ps1`).')
$md.Add('')
$md.Add("- **LibreOffice** $versaoLo - ``--headless --norestore --convert-to pdf``; *abre* = codigo de saida 0 **e** PDF com mais de 0 bytes.")
$md.Add("- **Microsoft Office** $versaoMs - automacao COM em somente-leitura: abre, le uma grandeza, fecha, sai.")
$md.Add('- **PoC (2a geracao)** - o arquivo salvo e'' reaberto no PoC e salvo de novo; a 2a geracao e'' comparada')
$md.Add('  com a 1a pelo comparador de `fidelidade.ps1` (contagens + texto em OOXML, LibreOffice/`content.xml` em ODF).')
$md.Add('- **ONLYOFFICE oficial** - **(HUMAN) pendente**: o Desktop Editors nao expoe linha de comando que sirva')
$md.Add('  de prova (abrir e sair com codigo). Esta coluna e'' verificacao manual do usuario; o script nao a inventa.')
$md.Add('')
$md.Add('| Artefato | Reabre no PoC (2a geracao) | Reabre LibreOffice | Reabre MS Office | Reabre ONLYOFFICE oficial |')
$md.Add('|---|---|---|---|---|')
foreach ($r in $linhasResultado) {
    $md.Add("| $($r.Artefato) | $($r.Poc) | $($r.LibreOffice) | $($r.MsOffice) | $($r.OnlyOffice) |")
}
$md.Add('')
$sim = @($linhasResultado | Where-Object { $_.LibreOffice -like 'sim*' }).Count
$simMs = @($linhasResultado | Where-Object { $_.MsOffice -like 'sim*' }).Count
$simPoc = @($linhasResultado | Where-Object { $_.Poc -like 'sim*' }).Count
$n = $linhasResultado.Count
$md.Add("Resumo: PoC (2a geracao) **$simPoc de $n** - LibreOffice **$sim de $n** - MS Office **$simMs de $n** - ONLYOFFICE oficial: **0 de $n verificados (HUMAN)**.")

if ($originaisRenderizados) {
    $md.Add('')
    $md.Add('**Originais ODF renderizados em PDF (`-Pares`)** - ao lado do salvo, para conferencia visual')
    $md.Add("original x salvo (pasta \`$pastaOriginais\`, fora do git):")
    $md.Add('')
    foreach ($o in $originaisRenderizados) { $md.Add("- $($o.Rotulo): $($o.Resultado)") }
}

# UTF-8 SEM BOM (mesma nota de medir.ps1/fidelidade.ps1).
New-Item -ItemType Directory -Path (Split-Path -Parent $Relatorio) -Force | Out-Null
[IO.File]::WriteAllLines($Relatorio, $md, (New-Object Text.UTF8Encoding($false)))

Write-Host ''
$linhasResultado | Format-Table -AutoSize | Out-Host
Write-Host "Resumo: PoC (2a geracao) $simPoc de $n | LibreOffice $sim de $n | MS Office $simMs de $n | ONLYOFFICE oficial 0 de $n (HUMAN)"
Write-Host "Relatorio: $Relatorio"
