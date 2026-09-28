<#
.SYNOPSIS
  Extrai o allfontsgen.exe (gerador do catalogo de fontes do ONLYOFFICE) do pacote do ONLYOFFICE Docs
  (Document Server) para Windows, SEM instalar, e copia para third_party/onlyoffice/converter/.

.DESCRIPTION
  Rota S1a da sondagem do Marco 1 (Task 2). O pacote do Desktop Editors 9.4.0.129 que o mOffice embarca
  NAO traz o allfontsgen.exe (o app oficial gera o catalogo por dentro, com o mesmo codigo linkado no
  binario principal). O pacote do Document Server para Windows traz o gerador como executavel solto em
  {app}\server\tools\ - e, MEDIDO em 2026-09-18, na MESMA versao 9.4.0.129 do pacote do desktop.

  O gerador vai para converter/, ao lado do x2t.exe, e leva SO o .exe (32 KB). As DLLs do core que ele
  carrega (kernel, graphics, UnicodeConverter, icu*) ja' estao la'. Conferido: as 19 DLLs do pacote do
  Document Server diferem das do desktop por exatamente 8.688 bytes CADA - o desktop e' o mesmo binario
  com assinatura Authenticode anexada. Mesmo codigo, mesma versao; nada e' sobrescrito.

  O instalador do Document Server e' Inno Setup: o 7-Zip (ate' 21.07, pelo menos) NAO abre esse formato -
  so' enxerga as secoes PE do stub. Por isso a ordem aqui e' innounp primeiro, 7-Zip como alternativa.

  Revisao 1 (fix review 1, item 5): os hashes esperados do PACOTE e do GERADOR sao conferidos (pinos da
  autorizacao de 2026-09-18, registrados em third_party/onlyoffice/VERSAO.txt e
  poc/relatorios/medidas-catalogo-20260918.txt) - um pacote ou um gerador com hash diferente falha ALTO, em
  vez de seguir em frente extraindo/copiando um binario nao conferido. A versao do gerador tambem e'
  comparada com a versao do motor desktop ja embarcado (VERSAO.txt, linha 1): misturar builds do core e'
  risco de incompatibilidade com o x2t.exe que o resto do produto usa.

.PARAMETER Pacote
  Caminho do instalador ja' baixado. O download e' autorizado pelo usuario e feito a' parte, para
  third_party/onlyoffice/cache/ (ignorado pelo git):
    https://download.onlyoffice.com/install/documentserver/windows/onlyoffice-documentserver.exe

.PARAMETER Destino
  Pasta de destino do gerador. Padrao: third_party/onlyoffice/converter/.

.PARAMETER Extracao
  Pasta temporaria de extracao. Apagada no fim (inclusive se o script falhar no meio - try/finally).

.PARAMETER Sha256PacoteEsperado
  Sha256 esperado do INSTALADOR (o .exe/.msi passado em -Pacote). Padrao: o pacote autorizado em
  2026-09-18 (onlyoffice-documentserver.exe, 9.4.0.129). Vazio ("") pula a conferencia.

.PARAMETER Sha256GeradorEsperado
  Sha256 esperado do allfontsgen.exe DEPOIS de extraido do pacote. Padrao: o gerador extraido em
  2026-09-18 (9.4.0.129). Vazio ("") pula a conferencia.

.PARAMETER IgnorarVersao
  Segue em frente mesmo se a ProductVersion do allfontsgen.exe extraido for diferente da versao do motor
  desktop registrada em VERSAO.txt (linha 1). Sem isto, a divergencia de versao e' erro.

.EXAMPLE
  .\tools\onlyoffice\obter-allfontsgen.ps1 -Pacote third_party\onlyoffice\cache\onlyoffice-documentserver.exe
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Pacote,
    [string]$Destino  = (Join-Path $PSScriptRoot "..\..\third_party\onlyoffice\converter"),
    [string]$Extracao = (Join-Path $env:TEMP "moffice-ds-extracao"),
    # Pinos da autorizacao de 2026-09-18 (poc/relatorios/medidas-catalogo-20260918.txt e
    # third_party/onlyoffice/VERSAO.txt). Um download vem de fora mesmo quando autorizado - a diferenca entre
    # "o mesmo binario AGPL de sempre" e "um exe desconhecido rodando com os privilegios de quem instala o
    # mOffice" e' exatamente este hash.
    [string]$Sha256PacoteEsperado  = "7ba127f7d3e82eb6881302cb8e12f8549d1afee8d1f74aaaded231a26d10376e",
    [string]$Sha256GeradorEsperado = "ec1e88896002d8259b209ce7262d754b1d970f301b4bc28c17fe8eeb900e75bb",
    [switch]$IgnorarVersao
)
$ErrorActionPreference = "Stop"

if (-not (Test-Path $Pacote)) { throw "Pacote nao encontrado: $Pacote" }
$Pacote = (Resolve-Path $Pacote).Path
if (-not (Test-Path $Destino)) { throw "Pasta do motor nao encontrada: $Destino (rode antes tools/onlyoffice/obter-motor.ps1)" }
$Destino = (Resolve-Path $Destino).Path
$versaoTxt = Join-Path $Destino "..\VERSAO.txt"

if ($Sha256PacoteEsperado) {
    $hashPacote = (Get-FileHash -Algorithm SHA256 $Pacote).Hash.ToLowerInvariant()
    if ($hashPacote -ne $Sha256PacoteEsperado.ToLowerInvariant()) {
        throw "Sha256 do PACOTE nao confere (esperado $Sha256PacoteEsperado, obtido $hashPacote) - '$Pacote' " +
              "nao e' o instalador autorizado em 2026-09-18. Baixe de novo de " +
              "https://download.onlyoffice.com/install/documentserver/windows/onlyoffice-documentserver.exe " +
              "ou passe -Sha256PacoteEsperado explicito se a troca de versao for intencional."
    }
}

function Achar-Ferramenta([string]$nome, [string[]]$candidatos) {
    $c = Get-Command $nome -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    foreach ($p in $candidatos) { if ($p -and (Test-Path $p)) { return $p } }
    # instalada pelo winget mas sem o PATH recarregado nesta sessao
    $w = Get-ChildItem -Recurse -File -Filter $nome "$env:LOCALAPPDATA\Microsoft\WinGet\Packages" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($w) { return $w.FullName }
    return $null
}
$innounp = Achar-Ferramenta "innounp.exe" @()
$sete    = Achar-Ferramenta "7z.exe" @("C:\Program Files\7-Zip\7z.exe", "C:\Program Files (x86)\7-Zip\7z.exe")

if (Test-Path $Extracao) { Remove-Item -LiteralPath $Extracao -Recurse -Force }
New-Item -ItemType Directory -Force $Extracao | Out-Null

# try/finally (revisao 1, item 5): antes, uma falha no meio (hash que nao bate, ferramenta ausente,
# allfontsgen.exe nao encontrado) deixava a pasta de extracao (ate' ~1 GB, o pacote inteiro descompactado)
# parada em %TEMP%. Agora a limpeza acontece em QUALQUER saida, sucesso ou erro.
try {
    $ext = [IO.Path]::GetExtension($Pacote).ToLowerInvariant()
    switch ($ext) {
        ".msi" { Start-Process msiexec -ArgumentList "/a `"$Pacote`" /qn TARGETDIR=`"$Extracao`"" -Wait -NoNewWindow }
        ".exe" {
            if ($innounp) {
                # So' o que interessa: o gerador e, para diagnostico, as DLLs do converter do proprio pacote.
                & $innounp -x -d"$Extracao" "$Pacote" "{app}\server\tools\*" | Out-Null
            } elseif ($sete) {
                # 7-Zip 21.07 nao abre Inno Setup (ve so' as secoes PE do stub) - mas se um dia suportar (ou o
                # pacote virar outro formato), tenta extrair SO' o allfontsgen.exe primeiro: mais rapido e nao
                # deixa ~1 GB solto na pasta de extracao. So' cai para o pacote inteiro se isso nao achar nada.
                & $sete e -r -y -o"$Extracao" "$Pacote" "allfontsgen.exe" | Out-Null
                $achadoIsolado = Get-ChildItem -Recurse -File -Filter allfontsgen.exe $Extracao -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $achadoIsolado) {
                    Write-Host "7-Zip nao extraiu allfontsgen.exe isolado (formato provavelmente Inno Setup) - tentando o pacote inteiro."
                    & $sete x -y -o"$Extracao" "$Pacote" | Out-Null
                }
            } else {
                throw "Precisa de innounp (winget install --id JurgenRathlev.innounp) ou 7-Zip para extrair um .exe sem instalar. O instalador do Document Server e' Inno Setup - o 7-Zip 21.07 NAO abre esse formato."
            }
        }
        default { throw "Extensao nao suportada: $ext" }
    }

    $gen = Get-ChildItem -Recurse -File -Filter allfontsgen.exe $Extracao -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $gen) { throw "allfontsgen.exe nao encontrado dentro do pacote (procurado em $Extracao)." }
    Write-Host "allfontsgen encontrado em: $($gen.FullName.Substring($Extracao.Length))"

    $hashGeradorExtraido = (Get-FileHash -Algorithm SHA256 $gen.FullName).Hash.ToLowerInvariant()
    if ($Sha256GeradorEsperado -and $hashGeradorExtraido -ne $Sha256GeradorEsperado.ToLowerInvariant()) {
        throw "Sha256 do allfontsgen.exe EXTRAIDO nao confere (esperado $Sha256GeradorEsperado, obtido $hashGeradorExtraido) - " +
              "o conteudo do pacote do Document Server mudou sem o pino ser atualizado. Confira a proveniencia " +
              "antes de continuar, ou passe -Sha256GeradorEsperado explicito se a atualizacao for intencional."
    }

    $versaoGerador = (Get-Item $gen.FullName).VersionInfo.ProductVersion
    if (-not $IgnorarVersao -and (Test-Path $versaoTxt)) {
        $versaoMotor = (Get-Content -Path $versaoTxt -TotalCount 1).Trim([char]0xFEFF, ' ')
        if ($versaoMotor -and $versaoGerador -and $versaoMotor -ne $versaoGerador) {
            throw "Versao do allfontsgen.exe ($versaoGerador) DIFERENTE da versao do motor desktop em VERSAO.txt " +
                  "($versaoMotor) - misturar builds do core pode gerar um catalogo incompativel com o x2t.exe " +
                  "ja embarcado. Baixe o pacote do Document Server da MESMA versao do desktop, ou passe " +
                  "-IgnorarVersao para prosseguir mesmo assim."
        }
    }

    Copy-Item $gen.FullName (Join-Path $Destino "allfontsgen.exe") -Force
    # Rede de seguranca: se alguma DLL vizinha do gerador NAO existir em converter/, ela vai junto. Nada e'
    # sobrescrito - as que ja estao la' sao as assinadas do pacote do desktop, da mesma versao.
    Get-ChildItem -File -Filter *.dll $gen.DirectoryName -ErrorAction SilentlyContinue |
        Where-Object { -not (Test-Path (Join-Path $Destino $_.Name)) } |
        ForEach-Object { Copy-Item $_.FullName $Destino; Write-Host "DLL extra copiada: $($_.Name)" }

    $alvo   = Join-Path $Destino "allfontsgen.exe"
    $hash   = (Get-FileHash -Algorithm SHA256 $alvo).Hash.ToLowerInvariant()
    $versao = (Get-Item $alvo).VersionInfo.ProductVersion
    if (-not $versao) { $versao = "(sem ProductVersion)" }
    $linha  = "allfontsgen $versao sha256 $hash (extraido de $([IO.Path]::GetFileName($Pacote)))"

    if (-not (Select-String -Path $versaoTxt -SimpleMatch $hash -Quiet -ErrorAction SilentlyContinue)) {
        Add-Content -Encoding utf8 $versaoTxt -Value $linha
    }
    Write-Host $linha

    # O --help do allfontsgen 9.4.0.129 NAO imprime nada (sai com codigo 0). As opcoes reais estao no proprio
    # binario, como literais UTF-16 - e' de la' que esta lista sai, e cada uma foi confirmada rodando:
    #   --input (repetivel)  --allfonts  --allfonts-web  --selection  --images  --output-web
    #   --use-system=true|false  --use-system-user-fonts=true|false  --font-format
    Write-Host "`n--- opcoes encontradas no binario ---"
    $bytes = [IO.File]::ReadAllBytes($alvo)
    [regex]::Matches([Text.Encoding]::Unicode.GetString($bytes), '[\x20-\x7E]{4,}') |
        ForEach-Object { $_.Value } | Where-Object { $_ -like '--*' } | Select-Object -Unique | Sort-Object
}
finally {
    if (Test-Path $Extracao) { Remove-Item -LiteralPath $Extracao -Recurse -Force }
}
