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
.EXAMPLE
  tools\onlyoffice\aplicar-patches.ps1
  tools\onlyoffice\aplicar-patches.ps1 -Reverter
#>
[CmdletBinding()]
param(
    [string]$PastaMotor = (Join-Path $PSScriptRoot "..\..\third_party\onlyoffice"),
    [string]$Fonte = (Join-Path $PSScriptRoot "..\fontes\mOffice-Simbolos.ttf"),
    [switch]$Reverter,
    [string]$VersaoEsperada = "9.4.0.129"
)
$ErrorActionPreference = "Stop"

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

$bundles = $editores | ForEach-Object { Join-Path $PastaMotor "editors\sdkjs\$_\sdk-all.js" }
foreach ($b in $bundles) { if (-not (Test-Path $b)) { throw "Bundle nao encontrado: $b" } }

if ($Reverter) {
    foreach ($b in $bundles) {
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
