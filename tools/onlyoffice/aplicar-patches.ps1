<#
.SYNOPSIS
  Aplica (ou reverte) os patches do mOffice nos bundles do sdkjs em third_party/onlyoffice.

  Hoje existe UM patch, e ele e' uma modificacao de verdade do sdkjs (registrada em
  docs/licencas/MODIFICACOES-ONLYOFFICE.md e no aviso de modificacao de docs/licencas/AVISO-ONLYOFFICE.txt):

    A fonte de simbolos internos que o motor traz embutida em base64 dentro de
    editors/sdkjs/{word,cell,slide,visio}/sdk-all.js (familia "ASCW3", 6.868 bytes decodificados, sha256
    35ff902a3b9c9254ce3c57d610dad9e57008b8a3351691c463bee775644c1e18 - a mesma fonts/ASC.ttf da Ascensio, sem
    licenca localizavel) e' TROCADA pela fonte livre do mOffice, tools/fontes/mOffice-Simbolos.ttf (SIL OFL 1.1,
    desenhada em codigo por tools/fontes/construir-simbolos.py, mesmos cmaps e metricas). Contexto e evidencia:
    docs/licencas/AVISO-fontes.md, secao 1.

  Como acha o trecho: pelo CONTEUDO, nao por offset. Varre cada bundle atras de literais base64 que comecam
  como um TrueType ("AAEAAA" = 00 01 00 00 ...), decodifica cada candidato e confere o sha256. So' o trecho
  cujo sha256 e' o da fonte original e' trocado. Depois de gravar, RELE o arquivo e confere: zero ocorrencias
  da fonte original, exatamente uma da nossa, e o tamanho mudou exatamente pela diferenca dos dois base64.

  Idempotente: rodar de novo diz "ja aplicado". Se a fonte for regenerada (bytes novos), o bundle e' refeito a
  partir do <bundle>.orig guardado ao lado na primeira aplicacao (third_party/ e' ignorado pelo git; o
  .gitignore tambem cobre *.orig por garantia).

  Guarda de versao: so' patcha se third_party/onlyoffice/VERSAO.txt disser a versao esperada (9.4.0.129). Motor
  novo = parar com mensagem: o patch precisa ser refeito e re-medido (A/B de janela real) antes de valer.

  Chamado no fim de tools/onlyoffice/obter-motor.ps1 (motor recem-obtido sai patchado) e executavel sozinho.
.PARAMETER PastaMotor
  Raiz do motor extraido. Padrao: third_party/onlyoffice.
.PARAMETER Fonte
  O .ttf que entra no lugar da original. Padrao: tools/fontes/mOffice-Simbolos.ttf. (Os testes de janela real
  usam outro .ttf, VAZIO, numa copia temporaria do motor - controle negativo; nunca no motor compartilhado.)
.PARAMETER Reverter
  Restaura os 4 bundles a partir dos .orig e apaga os .orig. Se nao houver .orig e o bundle ja' for o original,
  nao faz nada.
.PARAMETER VersaoEsperada
  Versao do motor para a qual este patch foi medido.
.PARAMETER PastaPatches
  Patches de TEXTO (Plano 6): tools/onlyoffice/patches/NN-nome.json. Formato e regras no bloco "Patches de TEXTO" abaixo.
.PARAMETER SoPatchesDeTexto
  Pula o patch da fonte de simbolos (usado pelos testes com um motor-miniatura que so' tem VERSAO.txt e os arquivos citados).
.EXAMPLE
  tools\onlyoffice\aplicar-patches.ps1
  tools\onlyoffice\aplicar-patches.ps1 -Reverter
#>
[CmdletBinding()]
param(
    [string]$PastaMotor = "",
    [string]$Fonte = "",
    [switch]$Reverter,
    [string]$VersaoEsperada = "9.4.0.129",
    [string]$PastaPatches = "",
    [switch]$SoPatchesDeTexto
)
$ErrorActionPreference = "Stop"
# Padroes resolvidos AQUI e nao no param(): $PSScriptRoot chega vazio nos valores padrao do param() em algumas sessoes do PS 5.1 (medido).
$pastaDoScript = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $PastaMotor) { $PastaMotor = Join-Path $pastaDoScript "..\..\third_party\onlyoffice" }
if (-not $Fonte) { $Fonte = Join-Path $pastaDoScript "..\fontes\mOffice-Simbolos.ttf" }
if (-not $PastaPatches) { $PastaPatches = Join-Path $pastaDoScript "patches" }

# sha256 do blob original (decodificado) - o que este script REMOVE do pacote. Igual ao de fonts/ASC.ttf.
$shaOriginal = "35ff902a3b9c9254ce3c57d610dad9e57008b8a3351691c463bee775644c1e18"
$editores = @("word", "cell", "slide", "visio")

# Latin1 (ISO-8859-1) e' 1:1 byte<->char: ler e gravar por ela nao altera NENHUM byte fora do trecho trocado.
# (Ler como UTF-8 e regravar poderia mexer em bytes invalidos ou no BOM.)
$latin1 = [Text.Encoding]::GetEncoding(28591)
$sha = [Security.Cryptography.SHA256]::Create()
function Sha256Hex([byte[]]$bytes) { ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join "" }
function Sha256Arquivo([string]$caminho) { (Get-FileHash -Algorithm SHA256 $caminho).Hash.ToLowerInvariant() }

# Candidatos: literais base64 entre aspas que comecam como TrueType (00 01 00 00). {1000,}: fontes de verdade;
# evita casar prefixos curtos por acaso.
$padrao = [regex]'"(AAEAAA[A-Za-z0-9+/]{1000,}={0,2})"'

function Classificar([string]$texto, [string]$shaNossa) {
    $r = @{ Original = @(); Nossa = @() }
    foreach ($m in $padrao.Matches($texto)) {
        try { $bytes = [Convert]::FromBase64String($m.Groups[1].Value) } catch { continue }
        $h = Sha256Hex $bytes
        if ($h -eq $shaOriginal) { $r.Original += $m } elseif ($h -eq $shaNossa) { $r.Nossa += $m }
    }
    return $r
}

function Absoluto([string]$caminho) {
    if (-not [IO.Path]::IsPathRooted($caminho)) { $caminho = Join-Path (Get-Location).Path $caminho }
    return [IO.Path]::GetFullPath($caminho)
}
$PastaMotor = Absoluto $PastaMotor
$arquivoVersao = Join-Path $PastaMotor "VERSAO.txt"
if (-not (Test-Path $arquivoVersao)) { throw "VERSAO.txt nao encontrado em '$PastaMotor' - rode tools/onlyoffice/obter-motor.ps1 primeiro." }
$versao = (Get-Content $arquivoVersao -TotalCount 1).Trim([char]0xFEFF).Trim()
if ($versao -ne $VersaoEsperada) {
    throw "Motor na versao '$versao', patch medido para '$VersaoEsperada'. NAO aplicado: refaca e re-meca o patch (docs/licencas/AVISO-fontes.md, secao 1.5) antes de liberar para a versao nova."
}

# --- Patches de TEXTO (Plano 6): tools/onlyoffice/patches/NN-nome.json ------------------------------------------------------
# Formato: { "descricao", "versaoMotor", "arquivos": [ { "arquivo": "<caminho relativo ao motor, com />",
#            "trocas": [ { "de": "<trecho EXATO>", "para": "<trecho novo>", "ocorrencias": N } ] } ] }
# Regras: (1) o arquivo e' lido como UTF-8 ESTRITO, BOM preservado; (2) cada troca tem de achar "de" exatamente N vezes e "para" 0
# vezes (aplicar) - ou "de" 0 e "para" N vezes (ja aplicado); qualquer outra contagem para SEM GRAVAR NADA (nem do patch da fonte);
# (3) cada arquivo do motor pertence a UM so' patch .json (o .orig e' por arquivo) e nenhum e' um dos sdk-all.js do patch da fonte
# (que usa o mesmo .orig); (4) o resultado e' conferido ANTES de gravar e RELIDO do disco depois; (5) .orig ao lado na primeira
# aplicacao; -Reverter restaura dos .orig. As Tasks 8/9 do Plano 6 so' acrescentam arquivos NN-*.json.
$utf8Estrito = New-Object Text.UTF8Encoding($false, $true)
function Contar([string]$texto, [string]$trecho) { $n = 0; $i = 0; while (($i = $texto.IndexOf($trecho, $i, [StringComparison]::Ordinal)) -ge 0) { $n++; $i += $trecho.Length }; $n }
function LerTextoDoMotor([string]$alvo) {
    $bytes = [IO.File]::ReadAllBytes($alvo)
    $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $off = if ($bom) { 3 } else { 0 }
    return [pscustomobject]@{ Texto = $utf8Estrito.GetString($bytes, $off, $bytes.Length - $off); Bom = $bom }
}
function GravarTextoDoMotor([string]$alvo, [string]$texto, [bool]$bom) {
    $saida = $utf8Estrito.GetBytes($texto)
    if ($bom) { $saida = [byte[]](@([byte]0xEF, [byte]0xBB, [byte]0xBF) + $saida) }
    [IO.File]::WriteAllBytes($alvo, $saida)
}
function ListarPatchesDeTexto {
    $achados = @()
    if (-not (Test-Path -LiteralPath $PastaPatches)) { return $achados }
    $vistos = @{}
    foreach ($json in (Get-ChildItem -LiteralPath $PastaPatches -Filter *.json | Sort-Object Name)) {
        $p = [IO.File]::ReadAllText($json.FullName, $utf8Estrito) | ConvertFrom-Json
        if ($p.versaoMotor -ne $VersaoEsperada) { throw "$($json.Name): patch para '$($p.versaoMotor)', motor '$VersaoEsperada'. NAO aplicado." }
        foreach ($a in $p.arquivos) {
            $chave = ([string]$a.arquivo).ToLowerInvariant().Replace('\', '/')
            if ($vistos.ContainsKey($chave)) { throw "$($json.Name): '$($a.arquivo)' ja' e' alterado por $($vistos[$chave]) - um arquivo do motor pertence a um so' patch." }
            $vistos[$chave] = $json.Name
            if ($chave -match '^editors/sdkjs/(word|cell|slide|visio)/sdk-all\.js$') { throw "$($json.Name): '$($a.arquivo)' e' do patch da fonte de simbolos (mesmo .orig) - nao pode ter patch de texto." }
            $alvo = Join-Path $PastaMotor ($a.arquivo -replace '/', '\')
            if (-not (Test-Path -LiteralPath $alvo)) { throw "$($json.Name): arquivo nao encontrado: $alvo" }
            foreach ($t in $a.trocas) {
                if ([string]::IsNullOrEmpty($t.de) -or $t.de -ceq $t.para -or $t.ocorrencias -lt 1) { throw "$($json.Name): $($a.arquivo): troca invalida (de vazio, de igual a para ou ocorrencias < 1)." }
            }
            $achados += [pscustomobject]@{ Patch = $json.Name; Arquivo = $a.arquivo; Alvo = $alvo; Trocas = $a.trocas }
        }
    }
    return $achados
}
# Planeja TUDO (sem gravar): se um trecho nao bater (motor de outra versao, arquivo ja' mexido por fora), lanca.
function PlanejarTexto {
    $planos = @()
    foreach ($a in (ListarPatchesDeTexto)) {
        $arq = LerTextoDoMotor $a.Alvo
        $texto = $arq.Texto; $novo = $texto; $jaAplicado = $true
        foreach ($t in $a.Trocas) {
            $de = Contar $texto $t.de; $para = Contar $texto $t.para
            if ($de -eq 0 -and $para -eq $t.ocorrencias) { continue }
            if ($de -ne $t.ocorrencias -or $para -ne 0) { throw ("{0}: {1}: trecho 'de' {2}x (esperado {3}), 'para' {4}x (esperado 0). Nada gravado." -f $a.Patch, $a.Arquivo, $de, $t.ocorrencias, $para) }
            $jaAplicado = $false
            $novo = $novo.Replace($t.de, $t.para)
        }
        if (-not $jaAplicado) {
            # Conferido ANTES de gravar: uma troca cujo "para" contem o "de" de outra (ou o inverso) daria um resultado que o
            # criterio de "ja aplicado" nao reconheceria - recusa agora, com o motor intacto.
            foreach ($t in $a.Trocas) {
                if ((Contar $novo $t.de) -ne 0 -or (Contar $novo $t.para) -ne $t.ocorrencias) { throw ("{0}: {1}: o resultado das trocas nao fecha a contagem (trechos que se sobrepoem?). Nada gravado." -f $a.Patch, $a.Arquivo) }
            }
        }
        $planos += [pscustomobject]@{ Patch = $a.Patch; Arquivo = $a.Arquivo; Alvo = $a.Alvo; Bom = $arq.Bom; Novo = $novo; JaAplicado = $jaAplicado; Trocas = $a.Trocas }
    }
    return $planos
}
function AplicarTexto($planos) {
    foreach ($pl in $planos) {
        if ($pl.JaAplicado) { Write-Host ("ja aplicado  {0}  {1}" -f $pl.Patch, $pl.Alvo); continue }
        if (-not (Test-Path -LiteralPath "$($pl.Alvo).orig")) { Copy-Item -LiteralPath $pl.Alvo "$($pl.Alvo).orig" }
        GravarTextoDoMotor $pl.Alvo $pl.Novo $pl.Bom
        $relido = (LerTextoDoMotor $pl.Alvo).Texto
        foreach ($t in $pl.Trocas) {
            if ((Contar $relido $t.de) -ne 0 -or (Contar $relido $t.para) -ne $t.ocorrencias) { throw "Conferencia falhou em $($pl.Alvo) - restaure com -Reverter." }
        }
        Write-Host ("aplicado     {0}  {1}" -f $pl.Patch, $pl.Alvo)
    }
}
function ReverterTexto {
    # Planeja antes de gravar: o .orig tem de conter todos os "de" (e nenhum "para"); sem .orig, o arquivo tem de ja' ser o original.
    $restaurar = @()
    foreach ($a in (ListarPatchesDeTexto)) {
        $orig = "$($a.Alvo).orig"
        $alvoDaConferencia = if (Test-Path -LiteralPath $orig) { $orig } else { $a.Alvo }
        $texto = (LerTextoDoMotor $alvoDaConferencia).Texto
        foreach ($t in $a.Trocas) {
            if ((Contar $texto $t.de) -ne $t.ocorrencias -or (Contar $texto $t.para) -ne 0) { throw "$($a.Patch): $alvoDaConferencia nao e' o original deste patch - nao vou reverter por cima. Nada gravado." }
        }
        $restaurar += [pscustomobject]@{ Alvo = $a.Alvo; Orig = $orig; TemOrig = (Test-Path -LiteralPath $orig); Patch = $a.Patch }
    }
    foreach ($r in $restaurar) {
        if ($r.TemOrig) { Copy-Item -Force -LiteralPath $r.Orig $r.Alvo; Remove-Item -Force -LiteralPath $r.Orig; Write-Host ("revertido    {0}  {1}" -f $r.Patch, $r.Alvo) }
        else { Write-Host ("ja original  {0}  {1}" -f $r.Patch, $r.Alvo) }
    }
}

$bundles = $editores | ForEach-Object { Join-Path $PastaMotor "editors\sdkjs\$_\sdk-all.js" }
if (-not $SoPatchesDeTexto) { foreach ($b in $bundles) { if (-not (Test-Path $b)) { throw "Bundle nao encontrado: $b" } } }

if ($Reverter) {
    ReverterTexto
    foreach ($b in $(if ($SoPatchesDeTexto) { @() } else { $bundles })) {
        $orig = "$b.orig"
        if (Test-Path $orig) {
            $c = Classificar ($latin1.GetString([IO.File]::ReadAllBytes($orig))) ""
            if ($c.Original.Count -ne 1) { throw "$orig nao contem a fonte original exatamente 1x ($($c.Original.Count)) - nao vou restaurar por cima do bundle." }
            Copy-Item -Force $orig $b
            Remove-Item -Force $orig
            Write-Host ("revertido  {0}  sha256 {1}" -f $b, (Sha256Arquivo $b))
        } else {
            $c = Classificar ($latin1.GetString([IO.File]::ReadAllBytes($b))) ""
            if ($c.Original.Count -eq 1) { Write-Host "ja original  $b" }
            else { throw "Sem $orig para reverter e o bundle nao e' o original: $b" }
        }
    }
    return
}

# Texto primeiro so' PLANEJA (nao grava): trecho que nao bate para tudo, inclusive o patch da fonte.
$planosTexto = PlanejarTexto
if (-not $SoPatchesDeTexto) {
# Caminho absoluto UMA vez, resolvido contra o CWD do PowerShell: Test-Path usa esse CWD e [IO.File] usa o do
# processo, que podem divergir — assim os dois olham o mesmo arquivo. (PastaMotor passa pelo mesmo caminho acima.)
$Fonte = Absoluto $Fonte
if (-not (Test-Path -LiteralPath $Fonte)) { throw "Fonte nao encontrada: $Fonte (gere com python tools/fontes/construir-simbolos.py)." }
$bytesNossa = [IO.File]::ReadAllBytes($Fonte)
$shaNossa = Sha256Hex $bytesNossa
if ($shaNossa -eq $shaOriginal) { throw "A fonte informada E' a original da Ascensio - recusando." }
$b64Nossa = [Convert]::ToBase64String($bytesNossa)
Write-Host ("fonte: {0} ({1} bytes, sha256 {2})" -f $Fonte, $bytesNossa.Length, $shaNossa)

$aplicados = 0
foreach ($b in $bundles) {
    $orig = "$b.orig"
    $antes = Sha256Arquivo $b
    $texto = $latin1.GetString([IO.File]::ReadAllBytes($b))
    $c = Classificar $texto $shaNossa

    if ($c.Nossa.Count -eq 1 -and $c.Original.Count -eq 0) {
        Write-Host ("ja aplicado  {0}  sha256 {1}" -f $b, $antes)
        continue
    }
    if ($c.Original.Count -eq 1) {
        if (-not (Test-Path $orig)) { Copy-Item $b $orig }
    } elseif (Test-Path $orig) {
        # Bundle com uma fonte NOSSA de outra geracao (ou algo inesperado): refaz a partir do original guardado.
        $texto = $latin1.GetString([IO.File]::ReadAllBytes($orig))
        $c = Classificar $texto $shaNossa
        if ($c.Original.Count -ne 1) { throw "$orig nao contem a fonte original exatamente 1x ($($c.Original.Count))." }
    } else {
        throw ("Nao achei o trecho a trocar em {0}: original {1}x, nossa {2}x. Bundle de outra versao ou ja' alterado por fora - nada gravado." -f $b, $c.Original.Count, $c.Nossa.Count)
    }

    $m = $c.Original[0]
    $b64Original = $m.Groups[1].Value
    # Troca SO' o miolo entre as aspas (Index+1 .. Length-2): as aspas ficam.
    $novo = $texto.Substring(0, $m.Index + 1) + $b64Nossa + $texto.Substring($m.Index + $m.Length - 1)
    [IO.File]::WriteAllBytes($b, $latin1.GetBytes($novo))

    # Confere o RESULTADO relendo do disco, nao a intencao.
    $relido = $latin1.GetString([IO.File]::ReadAllBytes($b))
    $c2 = Classificar $relido $shaNossa
    $esperado = $texto.Length - $b64Original.Length + $b64Nossa.Length
    if ($c2.Original.Count -ne 0 -or $c2.Nossa.Count -ne 1 -or $relido.Length -ne $esperado) {
        throw ("Patch gravado mas a conferencia falhou em {0}: original {1}x, nossa {2}x, tamanho {3} (esperado {4}). Restaure com -Reverter." -f $b, $c2.Original.Count, $c2.Nossa.Count, $relido.Length, $esperado)
    }
    $aplicados++
    Write-Host ("aplicado   {0}`n           antes  {1}`n           depois {2}" -f $b, $antes, (Sha256Arquivo $b))
}
Write-Host ("Patch da fonte de simbolos: {0} bundle(s) alterado(s) agora, {1} ja' estavam. Motor {2}." -f $aplicados, ($bundles.Count - $aplicados), $versao)
}
AplicarTexto $planosTexto
