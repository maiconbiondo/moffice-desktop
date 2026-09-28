<#
.SYNOPSIS
    Marco 0 / Task 7 - fidelidade objetiva do round-trip: compara cada amostra com o arquivo que o PoC salvou.

.DESCRIPTION
    Para cada amostra, o script faz o round-trip PELO EDITOR (mOffice.Sonda.exe <origem> --salvar <destino>,
    que abre no editor real, pede asc_Save e sai) e depois compara origem x destino:

      OOXML (comparacao estrutural, sem depender de nenhum outro programa)
        DOCX -> numero de paragrafos (w:p), de imagens (w:drawing), de tabelas (w:tbl) e o texto integral (w:t)
        XLSX -> numero de planilhas, de celulas com formula (<f>) e o texto/valores integrais
        PPTX -> numero de slides, de formas (p:sp) e o texto integral (a:t)

      ODF (o formato interno muda de nome de estilo a cada gravacao; comparar XML seria ruido)
        ODT/ODS/ODP -> LibreOffice headless converte OS DOIS para txt e compara o texto linha a linha.
        Quando o LibreOffice NAO converte (o filtro `txt:Text` e' do Writer e recusa ODS/ODP), cai para o
        `content.xml` do proprio pacote. Sao metodos DIFERENTES, com forcas diferentes, e por isso cada linha
        do relatorio traz a coluna `Metodo` dizendo qual foi usado nela (antes o cabecalho falava so' de
        LibreOffice e o relatorio afirmava mais do que o script fazia).

    Saida: uma linha por par, `igual [metodo]` ou `diferente [metodo] (o que mudou)`, e um resumo no fim.
    Com -Markdown <arquivo> tambem grava a tabela pronta para o protocolo.

    -ProvarDeteccao e' o CONTROLE do proprio script: depois do round-trip ele altera de proposito o texto do
    arquivo salvo, entao TODAS as linhas tem de sair "diferente". Uma rodada normal que da "6 de 6 iguais" so
    vale alguma coisa se esta rodada der "0 de 6 iguais" - senao o comparador poderia estar sempre calado.

.PARAMETER Amostras
    Caminhos especificos. Sem isto (e sem -Pares), usa poc/amostras/ e poc/amostras/sinteticas/ (menos os
    `grande-*`).

.PARAMETER Pares
    Lista de pares { Origem; Rotulo } para amostras CONFIDENCIAIS (ex.: arquivos reais do escritorio em
    poc/amostras/, com nome real e possivelmente acentos/espacos). Quando informado, o RELATORIO usa o Rotulo
    (nunca o nome real do arquivo) na coluna Amostra e no nome do arquivo `<Rotulo>-salvo.<ext>` gravado em
    -Saida - e' assim que um nome real nunca chega a um arquivo commitado. O round-trip em si (Host <Origem>
    --salvar <destino>) roda com o CAMINHO REAL de Origem, entao acentos/espacos no nome sao exercitados de
    ponta a ponta (URL/titulo/pasta de recovery) mesmo com o relatorio redigido por rotulo. Nao muda o
    comportamento padrao (sem -Pares, nada disto entra em jogo).

.PARAMETER Saida
    Pasta de trabalho para os arquivos salvos pelo PoC (padrao: %TEMP%\moffice-fidelidade).

.PARAMETER SemTrechos
    Redige a coluna Detalhe: so contagens/posicoes/classificacao, NUNCA um trecho do texto do documento. Ligado
    por padrao quando -Pares e' usado (amostras confidenciais); passar -SemTrechos explicitamente (true ou
    false) decide por cima disso nos dois sentidos - ver o comentario de `$redigirTrechos` no corpo do script.

.EXAMPLE
    pwsh tools/fidelidade.ps1
    pwsh tools/fidelidade.ps1 -Markdown poc/relatorios/fidelidade-20260917.md
    -ProvarDeteccao serve para provar que o comparador fala quando ha diferenca.
    pwsh tools/fidelidade.ps1 -Pares @(@{Origem='poc\amostras\real-docx.docx'; Rotulo='real-docx'}) -Saida poc\amostras\trabalho
    -Pares redige o Detalhe por padrao (-SemTrechos:$false desliga, para depuracao local).
#>
[CmdletBinding()]
param(
    [string[]] $Amostras,
    [object[]] $Pares,
    [string]   $Saida = (Join-Path $env:TEMP 'moffice-fidelidade'),
    [string]   $Markdown,
    [int]      $EsperaSegundos = 90,
    [switch]   $ProvarDeteccao,
    [switch]   $SemTrechos
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Redacao (achado da revisao 2026-09-18): com -Pares (amostras CONFIDENCIAIS), o Detalhe NAO deve carregar
# trecho nenhum de texto do documento real - nem no markdown, nem no console (os dois eram persistidos: o
# markdown vai para poc/amostras/trabalho/*.md, git-ignorado mas ainda assim um arquivo em disco com texto
# real, e o console idem quando redirecionado/logado). Regra: -Pares ativa a redacao POR PADRAO (o Detalhe vira
# so contagens/posicoes/classificacao); -SemTrechos, passado explicitamente, decide por cima disso nos dois
# sentidos (liga a redacao mesmo sem -Pares, ou desliga mesmo com -Pares, para quem precisar depurar). Sem
# -Pares e sem -SemTrechos, nada muda: as amostras sinteticas continuam com o Detalhe cheio de hoje.
$redigirTrechos = if ($PSBoundParameters.ContainsKey('SemTrechos')) { [bool]$SemTrechos } else { [bool]$Pares }

$raiz = Split-Path -Parent $PSScriptRoot
$exe  = Join-Path $raiz 'src\mOffice.Sonda\bin\Debug\net10.0-windows\mOffice.Sonda.exe'
if (-not (Test-Path $exe)) { throw "Host nao compilado: $exe. Rode 'dotnet build mOffice.slnx' antes." }

$soffice = 'C:\Program Files\LibreOffice\program\soffice.exe'

$ns = @{
    w  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    a  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    p  = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    s  = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
}

function Ler-Entrada {
    param([string] $Zip, [string] $Entrada)
    $arquivo = [IO.Compression.ZipFile]::OpenRead($Zip)
    try {
        $e = $arquivo.GetEntry($Entrada)
        if (-not $e) { return $null }
        $sr = New-Object IO.StreamReader($e.Open())
        try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
    } finally { $arquivo.Dispose() }
}

function Entradas-Que-Casam {
    param([string] $Zip, [string] $Padrao)
    $arquivo = [IO.Compression.ZipFile]::OpenRead($Zip)
    try { return @($arquivo.Entries | Where-Object { $_.FullName -like $Padrao } | Select-Object -ExpandProperty FullName) }
    finally { $arquivo.Dispose() }
}

function Contar {
    param([string] $Xml, [string] $XPath)
    if (-not $Xml) { return 0 }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    foreach ($k in $ns.Keys) { $mgr.AddNamespace($k, $ns[$k]) }
    return @($doc.SelectNodes($XPath, $mgr)).Count
}

function Texto-De {
    param([string] $Xml, [string] $XPath)
    if (-not $Xml) { return '' }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($Xml)
    $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    foreach ($k in $ns.Keys) { $mgr.AddNamespace($k, $ns[$k]) }
    return (-join (@($doc.SelectNodes($XPath, $mgr)) | ForEach-Object { $_.InnerText }))
}

# Descreve ONDE dois textos divergem. So "origem N chars, salvo N chars" nao ajuda ninguem (e fica mudo
# quando a troca mantem o tamanho, que e' justo o caso do controle -ProvarDeteccao).
# -Redigir (amostras CONFIDENCIAIS, ver $redigirTrechos no topo do script): omite os trechos citados
# (origem '...' vs salvo '...') e devolve so posicao/tamanhos/classificacao - nunca texto do documento.
function Diferenca-De-Texto {
    param([string] $A, [string] $B, [switch] $Redigir)
    $n = [Math]::Min($A.Length, $B.Length)
    $i = 0
    while ($i -lt $n -and $A[$i] -eq $B[$i]) { $i++ }
    if ($i -eq $n -and $A.Length -eq $B.Length) { return $null }
    if ($Redigir) {
        $delta = $B.Length - $A.Length
        $soEspacos = if (($A -replace '\s', '') -eq ($B -replace '\s', '')) { '; so espacos mudaram (mesmos caracteres nao-espaco, na mesma ordem)' } else { '' }
        return "texto diverge no caractere $i de $n comparados (origem $($A.Length) chars, salvo $($B.Length) chars, delta $delta)$soEspacos"
    }
    $trecho = { param($t, $p) $ini = [Math]::Max(0, $p - 10); $t.Substring($ini, [Math]::Min(40, $t.Length - $ini)) }
    return "texto diverge no caractere $i (origem $($A.Length) chars, salvo $($B.Length)): origem '$(& $trecho $A $i)' vs salvo '$(& $trecho $B $i)'"
}

# Classifica as linhas que Compare-Object aponta como diferentes entre dois ODF convertidos para texto, SEM
# citar o conteudo delas - usado em modo -Redigir. Cada linha "so em A" que, sem os espacos das PONTAS, bate
# com alguma linha "so em B" conta como "perdeu/ganhou indentacao" (mesmas palavras, so mudou o espaco em
# branco no inicio/fim); o resto fica "nao classificado" (pode ser conteudo realmente diferente).
function Classificar-DiferencaLinhas {
    param([object[]] $Dif)
    $soOrigem  = @($Dif | Where-Object SideIndicator -eq '<=' | Select-Object -ExpandProperty InputObject)
    $soDestino = @($Dif | Where-Object SideIndicator -eq '=>' | Select-Object -ExpandProperty InputObject)
    $usados = New-Object System.Collections.Generic.HashSet[int]
    $indentacao = 0
    for ($o = 0; $o -lt $soOrigem.Count; $o++) {
        $trim = $soOrigem[$o].Trim()
        for ($d = 0; $d -lt $soDestino.Count; $d++) {
            if ($usados.Contains($d)) { continue }
            if ($soDestino[$d].Trim() -eq $trim -and $soDestino[$d] -cne $soOrigem[$o]) {
                $indentacao++; [void]$usados.Add($d); break
            }
        }
    }
    $naoClassificadas = $Dif.Count - ($indentacao * 2)
    $resumo = "$($Dif.Count) linha(s) diferente(s)"
    if ($indentacao -gt 0) { $resumo += "; $indentacao com so espaco no inicio/fim mudado (mesmo texto)" }
    if ($naoClassificadas -gt 0) { $resumo += "; $naoClassificadas linha(s) nao classificada(s) (pode ser conteudo, nao so espaco)" }
    return $resumo
}

# ---------------------------------------------------------------------------------------------------------
# Perfis por formato: contagens + texto integral
# ---------------------------------------------------------------------------------------------------------
function Perfil-Docx {
    param([string] $Arquivo)
    $xml = Ler-Entrada $Arquivo 'word/document.xml'
    [ordered]@{
        paragrafos = Contar $xml '//w:p'
        imagens    = Contar $xml '//w:drawing'
        tabelas    = Contar $xml '//w:tbl'
        texto      = (Texto-De $xml '//w:t') -replace '\s+', ' '
    }
}

function Perfil-Xlsx {
    param([string] $Arquivo)
    $wb     = Ler-Entrada $Arquivo 'xl/workbook.xml'
    $folhas = Entradas-Que-Casam $Arquivo 'xl/worksheets/sheet*.xml'
    $formulas = 0; $comValor = 0; $texto = ''
    foreach ($f in $folhas) {
        $x = Ler-Entrada $Arquivo $f
        $formulas += Contar $x '//s:f'
        $comValor += Contar $x '//s:c[s:v or s:is]'
        $texto    += Texto-De $x '//s:v | //s:is'
    }
    # sharedStrings: o x2t regrava textos como shared string; sem resolver, o texto "muda" so por isso.
    $ss = Ler-Entrada $Arquivo 'xl/sharedStrings.xml'
    if ($ss) { $texto += Texto-De $ss '//s:si' }
    [ordered]@{
        planilhas = Contar $wb '//s:sheet'
        formulas  = $formulas
        celulas   = $comValor
        texto     = ($texto -replace '\s+', ' ')
    }
}

function Perfil-Pptx {
    param([string] $Arquivo)
    $slides = Entradas-Que-Casam $Arquivo 'ppt/slides/slide*.xml'
    $formas = 0; $texto = ''
    foreach ($s in ($slides | Sort-Object)) {
        $x = Ler-Entrada $Arquivo $s
        $formas += Contar $x '//p:sp'
        $texto  += Texto-De $x '//a:t'
    }
    [ordered]@{
        slides = $slides.Count
        formas = $formas
        texto  = ($texto -replace '\s+', ' ')
    }
}

function Texto-Por-LibreOffice {
    param([string] $Arquivo, [string] $Pasta)
    if (-not (Test-Path $soffice)) { return $null }
    New-Item -ItemType Directory -Path $Pasta -Force | Out-Null
    $txt = Join-Path $Pasta ([IO.Path]::GetFileNameWithoutExtension($Arquivo) + '.txt')
    if (Test-Path $txt) { Remove-Item -LiteralPath $txt -Force }
    # Start-Process (e nao o operador &): o soffice desta instalacao escreve avisos no stderr
    # ("Could not find platform independent libraries"), e no Windows PowerShell 5.1 isso vira
    # NativeCommandError -- com $ErrorActionPreference='Stop' o script inteiro morria num aviso.
    $erro = Join-Path $Pasta 'soffice-stderr.txt'
    $saidaLo = Join-Path $Pasta 'soffice-stdout.txt'
    $p = Start-Process -FilePath $soffice -PassThru -WindowStyle Hidden `
        -ArgumentList @('--headless', '--norestore', '--convert-to', 'txt:Text', '--outdir', "`"$Pasta`"", "`"$Arquivo`"") `
        -RedirectStandardError $erro -RedirectStandardOutput $saidaLo
    if (-not $p.WaitForExit(120000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; return $null }
    if (-not (Test-Path $txt)) { return $null }
    return @(Get-Content -LiteralPath $txt)
}

# Plano B do ODF quando o LibreOffice nao converte (o filtro 'txt:Text' e' do Writer; ODS/ODP podem recusar):
# le o content.xml do proprio pacote e junta o texto. Nao substitui o LibreOffice (nao roda o layout), mas
# responde a mesma pergunta objetiva -- o texto sobreviveu ao round-trip?
function Texto-De-Odf {
    param([string] $Arquivo)
    $xml = Ler-Entrada $Arquivo 'content.xml'
    if (-not $xml) { return $null }
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($xml)
    $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $mgr.AddNamespace('text', 'urn:oasis:names:tc:opendocument:xmlns:text:1.0')
    return @(@($doc.SelectNodes('//text:p | //text:h', $mgr)) | ForEach-Object { $_.InnerText } | Where-Object { $_ })
}

# ---------------------------------------------------------------------------------------------------------
# Controle do comparador (-ProvarDeteccao): altera o texto dentro do pacote salvo. Toca so as partes XML que
# contem a marca das amostras, entao vale igual para OOXML (word/document.xml, sharedStrings, slides) e para
# ODF (content.xml), sem saber nada do formato.
# ---------------------------------------------------------------------------------------------------------
function Mutar-Texto {
    param([string] $Pacote)
    $alteradas = 0
    $zip = [IO.Compression.ZipFile]::Open($Pacote, 'Update')
    try {
        foreach ($nomeEntrada in @($zip.Entries | Where-Object { $_.FullName -like '*.xml' } | Select-Object -ExpandProperty FullName)) {
            $e = $zip.GetEntry($nomeEntrada)
            $sr = New-Object IO.StreamReader($e.Open()); $conteudo = $sr.ReadToEnd(); $sr.Dispose()
            if ($conteudo -notlike '*mOffice*') { continue }
            $novo = $conteudo.Replace('mOffice', 'mXffice')
            $e.Delete()
            $nova = $zip.CreateEntry($nomeEntrada)
            $sw = New-Object IO.StreamWriter($nova.Open()); $sw.Write($novo); $sw.Dispose()
            $alteradas++
        }
    } finally { $zip.Dispose() }
    return $alteradas
}

# ---------------------------------------------------------------------------------------------------------
# Round-trip pelo editor
# ---------------------------------------------------------------------------------------------------------
function RoundTrip {
    param([string] $Origem, [string] $Destino)
    if (Test-Path $Destino) { Remove-Item -LiteralPath $Destino -Force }
    # Aspas obrigatorias: Start-Process passa os argumentos concatenados, e um caminho com espaco chegaria
    # partido no Host (medido na Task 7 - o Host abre um MessageBox modal e o processo nunca sai).
    # Foto dos PIDs do WebView2 ANTES de subir o Host (revisao final 2026-09-17). O criterio antigo era a HORA
    # DE INICIO ("comecou depois do nosso Host"), e isso matava o WebView2 de QUALQUER outro aplicativo do
    # usuario que tivesse subido no mesmo intervalo - o Teams, o Outlook novo, o proprio Windows. Mesmo desenho
    # do reabrir.ps1: so' pode ser encerrado o que NAO existia antes desta chamada.
    $webviewAntes = @(Get-Process msedgewebview2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $p = Start-Process -FilePath $exe -ArgumentList @("`"$Origem`"", '--salvar', "`"$Destino`"") -PassThru
    if (-not $p.WaitForExit($EsperaSegundos * 1000)) {
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        return $false
    }
    Get-Process msedgewebview2 -ErrorAction SilentlyContinue |
        Where-Object { $webviewAntes -notcontains $_.Id } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    return (Test-Path $Destino)
}

# ---------------------------------------------------------------------------------------------------------

# -Pares: cada { Origem; Rotulo } vira uma amostra, e o Rotulo fica registrado para render o relatorio sem o
# nome real do arquivo (ver o PARAMETER acima). Nao mistura com a descoberta padrao - quem pede -Pares sabe
# exatamente quais arquivos quer.
$rotulosPorOrigem = @{}
if ($Pares) {
    $Amostras = @()
    foreach ($par in $Pares) {
        $origemPar = [string]$par.Origem
        $rotuloPar = [string]$par.Rotulo
        if (-not (Test-Path -LiteralPath $origemPar)) { throw "Par nao encontrado: $origemPar" }
        $Amostras += (Resolve-Path -LiteralPath $origemPar).Path
        $rotulosPorOrigem[(Resolve-Path -LiteralPath $origemPar).Path] = $rotuloPar
    }
}

if (-not $Amostras) {
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
if (-not $Amostras) { throw 'Nenhuma amostra encontrada.' }

New-Item -ItemType Directory -Path $Saida -Force | Out-Null
$resultado = @()
$metodos = @()

foreach ($origem in $Amostras) {
    $ext  = [IO.Path]::GetExtension($origem).ToLowerInvariant()
    # Rotulo (sem extensao) quando -Pares deu um; senao o nome de arquivo de sempre (comportamento padrao intacto).
    $baseSemExt = if ($rotulosPorOrigem.ContainsKey($origem)) { $rotulosPorOrigem[$origem] } else { [IO.Path]::GetFileNameWithoutExtension($origem) }
    $nome    = if ($rotulosPorOrigem.ContainsKey($origem)) { $rotulosPorOrigem[$origem] } else { [IO.Path]::GetFileName($origem) }
    $destino = Join-Path $Saida ($baseSemExt + '-salvo' + $ext)

    Write-Host "Round-trip $nome ..."
    if (-not (RoundTrip -Origem $origem -Destino $destino)) {
        $resultado += [pscustomobject]@{ Amostra = $nome; Veredito = 'FALHOU (o PoC nao gravou o arquivo)'; Metodo = 'nao comparado'; Detalhe = '' }
        continue
    }

    if ($ProvarDeteccao) {
        $quantas = Mutar-Texto -Pacote $destino
        Write-Host "  [controle] $quantas parte(s) do arquivo salvo alteradas de proposito"
    }

    $diferencas = @()
    # O METODO usado nesta linha. Ele varia por formato E, no ODF, por AMOSTRA (o filtro txt:Text do
    # LibreOffice e' do Writer e recusa ODS/ODP), entao tem de sair no relatorio linha a linha - o cabecalho
    # fixo que so' falava de LibreOffice para ODF afirmava mais do que o script fazia (achado critico da
    # revisao final 2026-09-17).
    $metodo = ''
    switch ($ext) {
        { $_ -in '.docx' } {
            $metodo = 'OpenXML contagens+texto'
            $a = Perfil-Docx $origem; $b = Perfil-Docx $destino
            foreach ($k in 'paragrafos', 'imagens', 'tabelas') { if ($a[$k] -ne $b[$k]) { $diferencas += "$k $($a[$k]) -> $($b[$k])" } }
            if ($a.texto -ne $b.texto) { $diferencas += (Diferenca-De-Texto $a.texto $b.texto -Redigir:$redigirTrechos) }
        }
        { $_ -in '.xlsx' } {
            $metodo = 'OpenXML contagens+texto'
            $a = Perfil-Xlsx $origem; $b = Perfil-Xlsx $destino
            foreach ($k in 'planilhas', 'formulas', 'celulas') { if ($a[$k] -ne $b[$k]) { $diferencas += "$k $($a[$k]) -> $($b[$k])" } }
            if ($a.texto -ne $b.texto) { $diferencas += (Diferenca-De-Texto $a.texto $b.texto -Redigir:$redigirTrechos) }
        }
        { $_ -in '.pptx' } {
            $metodo = 'OpenXML contagens+texto'
            $a = Perfil-Pptx $origem; $b = Perfil-Pptx $destino
            foreach ($k in 'slides', 'formas') { if ($a[$k] -ne $b[$k]) { $diferencas += "$k $($a[$k]) -> $($b[$k])" } }
            if ($a.texto -ne $b.texto) { $diferencas += (Diferenca-De-Texto $a.texto $b.texto -Redigir:$redigirTrechos) }
        }
        default {
            # ODT / ODS / ODP
            $metodo = 'LibreOffice txt'
            $ta = Texto-Por-LibreOffice $origem  (Join-Path $Saida 'lo-origem')
            $tb = Texto-Por-LibreOffice $destino (Join-Path $Saida 'lo-salvo')
            if ($null -eq $ta -or $null -eq $tb) {
                $metodo = 'content.xml (fallback)'
                $ta = Texto-De-Odf $origem
                $tb = Texto-De-Odf $destino
            }
            if ($null -eq $ta -or $null -eq $tb) {
                $metodo = 'nao comparado'
                $diferencas += 'texto nao comparado (nem LibreOffice nem content.xml)'
            } else {
                $dif = Compare-Object -ReferenceObject $ta -DifferenceObject $tb
                if ($dif) {
                    if ($redigirTrechos) {
                        $diferencas += "texto por $metodo : $(Classificar-DiferencaLinhas $dif)"
                    } else {
                        $diferencas += "texto por $metodo : $($dif.Count) linha(s) diferente(s) - $((($dif | Select-Object -First 3 | ForEach-Object { $_.SideIndicator + ' ' + $_.InputObject }) -join ' | '))"
                    }
                }
                else { $metodos += "$nome : texto conferido por $metodo" }
            }
        }
    }

    $veredito = if ($diferencas.Count -eq 0) { 'igual' } else { 'diferente' }
    $resultado += [pscustomobject]@{ Amostra = $nome; Veredito = $veredito; Metodo = $metodo; Detalhe = ($diferencas -join '; ') }
    Write-Host ("  {0} [{1}]{2}" -f $veredito, $metodo, $(if ($diferencas) { " ($($diferencas -join '; '))" } else { '' }))
}

Write-Host ''
$resultado | Format-Table -AutoSize | Out-Host
$iguais = @($resultado | Where-Object Veredito -eq 'igual').Count
Write-Host "Resumo: $iguais de $($resultado.Count) iguais."
$metodos | ForEach-Object { Write-Host "  $_" }

if ($Markdown) {
    $linhas = New-Object System.Collections.Generic.List[string]
    $linhas.Add("# Fidelidade do round-trip - $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
    $linhas.Add('')
    $linhas.Add('Gerado por `tools/fidelidade.ps1`. O metodo NAO e'' o mesmo para todas as linhas - a coluna')
    $linhas.Add('`Metodo` diz qual foi usado em cada uma:')
    $linhas.Add('')
    $linhas.Add('- `OpenXML contagens+texto` - DOCX/XLSX/PPTX: contagens estruturais (paragrafos/imagens/tabelas,')
    $linhas.Add('  planilhas/formulas/celulas, slides/formas) + texto integral, lidos do proprio pacote. Nao depende')
    $linhas.Add('  de nenhum outro programa.')
    $linhas.Add('- `LibreOffice txt` - ODF: os DOIS arquivos convertidos para txt pelo LibreOffice headless e')
    $linhas.Add('  comparados linha a linha.')
    $linhas.Add('- `content.xml (fallback)` - ODF em que o LibreOffice nao converteu (o filtro `txt:Text` e'' do')
    $linhas.Add('  Writer e recusa ODS/ODP): o texto sai do `content.xml` do proprio pacote. Responde a mesma')
    $linhas.Add('  pergunta objetiva - o texto sobreviveu ao round-trip? -, mas NAO roda o layout.')
    $linhas.Add('')
    $linhas.Add('Ressalva das linhas ODF: as amostras `rico.odt`/`rico.ods`/`rico.odp` foram produzidas pelo PROPRIO')
    $linhas.Add('x2t a partir das OOXML sinteticas. Um `igual` aqui compara saida de x2t com saida de x2t, e por isso')
    $linhas.Add('nao diz nada sobre ODF escrito pelo LibreOffice ou pelo MS Office. ODF de escritorio real continua')
    $linhas.Add('pendente (`poc/amostras/` esta vazia).')
    $linhas.Add('')
    $linhas.Add('| Amostra | Resultado | Metodo | O que mudou |')
    $linhas.Add('|---|---|---|---|')
    foreach ($r in $resultado) { $linhas.Add("| $($r.Amostra) | $($r.Veredito) | $($r.Metodo) | $(if ($r.Detalhe) { $r.Detalhe } else { '-' }) |") }
    # UTF-8 sem BOM (ver a nota em medir.ps1).
    [IO.File]::WriteAllLines($Markdown, $linhas, (New-Object Text.UTF8Encoding($false)))
    Write-Host "Relatorio: $Markdown"
}
