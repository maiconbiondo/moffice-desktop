<#
.SYNOPSIS
  Registro de DESENVOLVIMENTO no Explorer: faz o duplo clique (na verdade, "Abrir com") de um documento do
  Office (.docx/.doc/.dotx/.xlsx/.xls/.pptx/.ppsx/.odt/...) chegar aos executaveis compilados de src/, sem instalador.

.DESCRIPTION
  Escreve em HKCU\Software\Classes (so' o usuario atual, nunca a maquina) os MESMOS ProgIDs que o
  instalador do Plano 5 vai escrever em HKLM:

      mOffice.Texto.<ext>     -> mTexto.exe      (docx, docm, dotx, dotm, doc, dot, odt, ott, rtf)
      mOffice.Planilha.<ext>  -> mPlanilha.exe   (xlsx, xlsm, xltx, xltm, xls, xlt, xlsb, ods, ots)
      mOffice.Slides.<ext>    -> mSlides.exe     (pptx, pptm, potx, potm, ppsx, ppsm, ppt, pps, pot, odp, otp)

  A lista vem de tools\installer\associacoes.json (gerado da TabelaFormatos). Modelos abrem com /modelo (documento novo) e
  apresentacoes de slides com /apresentar, e ganham o verbo edit (abre o proprio arquivo).

  Cada ProgID ganha DefaultIcon = "<exe>,0" e shell\open\command = "<exe>" "%1", e cada extensao ganha o
  ProgID em OpenWithProgids.

  Sem -Padrao: SO' OpenWithProgids — o efeito e' "mTexto passa a aparecer no Abrir com do .docx".

  Com -Padrao: tambem grava o valor padrao da extensao em HKCU\Software\Classes\<ext>, que tem
  prioridade sobre o de HKLM (o que o instalador do ONLYOFFICE/LibreOffice gravou para a maquina). Isso
  so' decide o duplo clique quando o usuario NAO tem escolha propria gravada (UserChoice ou, no Windows 11
  atual, UserChoiceLatest) — este script
  nunca escreve UserChoice, que o Windows 10/11 protege com hash e desfaz se for adulterado. Havendo
  UserChoice, o script avisa e o usuario troca pela interface (Abrir com > Sempre). O valor que existia
  antes em HKCU e' guardado em "mOffice.PadraoAnterior" e volta no -Remover.

.PARAMETER Configuracao
  Debug (padrao) ou Release: de qual pasta bin/ sair os .exe.

.PARAMETER Padrao
  Tambem torna os apps o padrao do duplo clique para o usuario atual (ver DESCRIPTION).

.PARAMETER Remover
  Apaga SO' o que este script cria (os ProgIDs, os valores de OpenWithProgids e, se o padrao
  do usuario ainda for o nosso ProgID, devolve o valor anterior). Nao toca em mais nada da extensao.

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File tools\explorer\registrar-dev.ps1
  powershell -NoProfile -ExecutionPolicy Bypass -File tools\explorer\registrar-dev.ps1 -Padrao
  powershell -NoProfile -ExecutionPolicy Bypass -File tools\explorer\registrar-dev.ps1 -Remover
#>
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')][string]$Configuracao = 'Debug',
    [switch]$Padrao,
    [switch]$Remover
)

$ErrorActionPreference = 'Stop'
$raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$classes = 'HKCU:\Software\Classes'
$valorAnterior = 'mOffice.PadraoAnterior'
$capRaiz = 'HKCU:\Software\mOffice.Dev'
$registrados = 'HKCU:\Software\RegisteredApplications'

# Plano 6, Task 2: a lista de ProgIDs e' UMA so' (tools\installer\associacoes.json, GERADO da TabelaFormatos pelo teste
# AssociacoesGeradas_EmDia) — o mesmo que o instalador e o RMM usam. Nada de lista literal repetida aqui.
$lista = Join-Path $raiz 'tools\installer\associacoes.json'
if (-not (Test-Path -LiteralPath $lista)) { throw "Lista de associacoes ausente: $lista" }
$registros = @((Get-Content -LiteralPath $lista -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object {
    @{ ProgId = [string]$_.ProgId; Ext = [string]$_.Ext; Projeto = [string]$_.App; Exe = [string]$_.Exe; Nome = [string]$_.Nome;
       Modelo = [bool]$_.Modelo; Apresentar = [bool]$_.Apresentar }
})

# NUNCA `New-Item -Force` numa chave de registro: se ela ja' existe, o provedor do PowerShell a RECRIA
# VAZIA (apaga valores e subchaves). Foi assim que uma versao deste script zerou
# HKCU\Software\RegisteredApplications (chave compartilhada com OneDrive, Zoom, apps da Loja).
# Esta funcao so' cria o que falta e nunca toca no conteudo de uma chave existente.
function GarantirChave([string] $Caminho) {
    if (-not (Test-Path -LiteralPath $Caminho)) { New-Item -Path $Caminho -Force | Out-Null }
}

function CaminhoDoExe([string] $Projeto, [string] $Exe) {
    Join-Path $raiz ('src\' + $Projeto + '\bin\' + $Configuracao + '\net10.0-windows\' + $Exe)
}

if ($Remover) {
    foreach ($r in $registros) {
        # Padrao do usuario: so' mexe se AINDA for o nosso ProgID (se o usuario trocou depois, e' dele).
        $chaveExt = Join-Path $classes $r.Ext
        if (Test-Path $chaveExt) {
            $props = Get-ItemProperty -Path $chaveExt
            if ($props.'(default)' -eq $r.ProgId) {
                $anterior = $props.$valorAnterior
                if ([string]::IsNullOrEmpty($anterior)) {
                    # O valor padrao nao sai com Remove-ItemProperty '(default)' (o provedor recusa); pela API
                    # .NET o nome do valor padrao e' a string vazia.
                    $k = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\' + $r.Ext, $true)
                    try { $k.DeleteValue('', $false) } finally { $k.Dispose() }
                } else {
                    Set-ItemProperty -Path $chaveExt -Name '(default)' -Value $anterior
                }
                Write-Host "restaurado $($r.Ext) padrao do usuario -> '$anterior'"
            }
            if ($null -ne $props.$valorAnterior) { Remove-ItemProperty -Path $chaveExt -Name $valorAnterior -Force }
        }
        $chave = Join-Path $classes $r.ProgId
        if (Test-Path $chave) {
            Remove-Item -Path $chave -Recurse -Force
            Write-Host "removido  $($r.ProgId)"
        } else {
            Write-Host "ausente   $($r.ProgId)"
        }
        # Apaga SO' o nosso valor de OpenWithProgids; a chave (e os ProgIDs de outros programas nela)
        # continua intacta.
        $abrirCom = Join-Path $classes ($r.Ext + '\OpenWithProgids')
        if ((Test-Path $abrirCom) -and ($null -ne (Get-ItemProperty -Path $abrirCom -Name $r.ProgId -ErrorAction SilentlyContinue))) {
            Remove-ItemProperty -Path $abrirCom -Name $r.ProgId -Force
            Write-Host "removido  $($r.Ext)\OpenWithProgids\$($r.ProgId)"
        }
        # Se a chave ficou VAZIA (nenhum valor e nenhuma subchave), ela era nossa: o Windows nao cria
        # OpenWithProgids sozinho, e nenhum outro programa registrou nada ali. Deixa-la para tras sujaria o
        # registro do usuario com uma chave orfa por extensao. Com QUALQUER valor restante (outro programa
        # tambem se registrou), a chave fica — apagar levaria junto o registro alheio.
        if (Test-Path $abrirCom) {
            $chaveAbrirCom = Get-Item -Path $abrirCom
            if ($chaveAbrirCom.ValueCount -eq 0 -and $chaveAbrirCom.SubKeyCount -eq 0) {
                Remove-Item -Path $abrirCom -Force
                Write-Host "removido  $($r.Ext)\OpenWithProgids (chave vazia)"
            }
        }
    }
    foreach ($app in ($registros | Group-Object { $_.Projeto })) {
        $primeiro = $app.Group[0]
        $appKey = Join-Path $classes ('Applications\' + $primeiro.Exe)
        if (Test-Path $appKey) { Remove-Item -Path $appKey -Recurse -Force; Write-Host "removido  Applications\$($primeiro.Exe)" }
        $nome = 'mOffice.Dev.' + $primeiro.Projeto
        if ((Test-Path $registrados) -and ($null -ne (Get-ItemProperty -Path $registrados -Name $nome -ErrorAction SilentlyContinue))) {
            Remove-ItemProperty -Path $registrados -Name $nome -Force
            Write-Host "removido  RegisteredApplications\$nome"
        }
    }
    # A arvore Software\mOffice.Dev e' so' deste script (o produto nao usa o registro para dados).
    if (Test-Path $capRaiz) { Remove-Item -Path $capRaiz -Recurse -Force; Write-Host "removido  $capRaiz" }
    Add-Type -Namespace mOfficeDev -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, uint f, System.IntPtr a, System.IntPtr b);' -ErrorAction SilentlyContinue
    [mOfficeDev.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
    Write-Host ''
    Write-Host 'Confira com: reg query HKCU\Software\Classes\mOffice.Texto.docx   (deve dizer que nao existe)'
    return
}

# Falha FECHADA: registrar um ProgID apontando para um .exe que nao existe deixaria o "Abrir com" do
# usuario com uma entrada morta, e o erro so' apareceria no duplo clique.
$faltando = @()
foreach ($r in ($registros | Group-Object { $_.Exe } | ForEach-Object { $_.Group[0] })) {
    $exe = CaminhoDoExe $r.Projeto $r.Exe
    if (-not (Test-Path $exe)) { $faltando += $exe }
}
if ($faltando.Count -gt 0) {
    throw ("Executavel nao compilado: " + ($faltando -join ', ') + "`nRode: dotnet build mOffice.slnx -c $Configuracao")
}

foreach ($r in $registros) {
    $exe = CaminhoDoExe $r.Projeto $r.Exe
    $chave = Join-Path $classes $r.ProgId

    GarantirChave $chave
    Set-ItemProperty -Path $chave -Name '(default)' -Value $r.Nome

    $icone = Join-Path $chave 'DefaultIcon'
    GarantirChave $icone
    Set-ItemProperty -Path $icone -Name '(default)' -Value ('"' + $exe + '",0')

    $comando = Join-Path $chave 'shell\open\command'
    GarantirChave $comando
    # Aspas nos DOIS: o caminho de desenvolvimento tem espaco (ex.: "C:\Meus Projetos") e o arquivo do usuario
    # tambem pode ter — sem elas o app recebe o caminho partido em varios argumentos.
    # Modelo (.dotx...) abre como documento NOVO (/modelo) e apresentacao de slides (.ppsx...) ja' apresenta (/apresentar); os dois
    # ganham o verbo edit para abrir o PROPRIO arquivo (como o instalador gera, a partir da mesma tabela).
    $verbo = ''
    if ($r.Modelo) { $verbo = '/modelo ' } elseif ($r.Apresentar) { $verbo = '/apresentar ' }
    Set-ItemProperty -Path $comando -Name '(default)' -Value ('"' + $exe + '" ' + $verbo + '"%1"')
    if ($r.Modelo -or $r.Apresentar) {
        $editar = Join-Path $chave 'shell\edit'
        GarantirChave (Join-Path $editar 'command')
        Set-ItemProperty -Path $editar -Name '(default)' -Value 'Editar'
        Set-ItemProperty -Path (Join-Path $editar 'command') -Name '(default)' -Value ('"' + $exe + '" "%1"')
    }

    $abrirCom = Join-Path $classes ($r.Ext + '\OpenWithProgids')
    GarantirChave $abrirCom
    # Valor VAZIO, como manda a documentacao de OpenWithProgids: o que importa e' o NOME do valor.
    New-ItemProperty -Path $abrirCom -Name $r.ProgId -Value '' -PropertyType String -Force | Out-Null

    Write-Host ("registrado  {0,-22} {1,-6} -> {2}" -f $r.ProgId, $r.Ext, $exe)

    if ($Padrao) {
        $chaveExt = Join-Path $classes $r.Ext
        $atual = (Get-ItemProperty -Path $chaveExt).'(default)'
        if ($atual -ne $r.ProgId) {
            # Guarda o que havia (mesmo vazio) UMA vez, para o -Remover devolver exatamente.
            if ($null -eq (Get-ItemProperty -Path $chaveExt).$valorAnterior) {
                New-ItemProperty -Path $chaveExt -Name $valorAnterior -Value ([string]$atual) -PropertyType String -Force | Out-Null
            }
            Set-ItemProperty -Path $chaveExt -Name '(default)' -Value $r.ProgId
        }
        # Escolha propria do usuario, protegida por hash: UserChoice (Windows 10/11 antigo) ou
        # UserChoiceLatest\ProgId (Windows 11 atual). Se existir, ela vence o padrao gravado acima.
        $fe = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$($r.Ext)"
        $escolhido = (Get-ItemProperty -Path "$fe\UserChoiceLatest\ProgId" -ErrorAction SilentlyContinue).ProgId
        if (-not $escolhido) { $escolhido = (Get-ItemProperty -Path "$fe\UserChoice" -ErrorAction SilentlyContinue).ProgId }
        if ($escolhido -and $escolhido -ne $r.ProgId) {
            Write-Host "  ATENCAO: $($r.Ext) tem escolha propria do usuario ($escolhido) - troque em Abrir com > Sempre."
        } else {
            Write-Host "  padrao do usuario: $($r.Ext) -> $($r.ProgId)"
        }
    }
}

# Aplicativo REGISTRADO (Capabilities + RegisteredApplications + Applications\<exe>): e' disso que o
# Windows 11 monta a lista "Escolher outro aplicativo" e a pagina Configuracoes > Aplicativos padrao.
# So' ProgID + OpenWithProgids nao basta para o app aparecer nessas telas.
foreach ($app in ($registros | Group-Object { $_.Projeto })) {
    $primeiro = $app.Group[0]
    $exe = CaminhoDoExe $primeiro.Projeto $primeiro.Exe

    $appKey = Join-Path $classes ('Applications\' + $primeiro.Exe)
    GarantirChave $appKey
    Set-ItemProperty -Path $appKey -Name 'FriendlyAppName' -Value $primeiro.Projeto
    GarantirChave (Join-Path $appKey 'DefaultIcon')
    Set-ItemProperty -Path (Join-Path $appKey 'DefaultIcon') -Name '(default)' -Value ('"' + $exe + '",0')
    GarantirChave (Join-Path $appKey 'shell\open\command')
    Set-ItemProperty -Path (Join-Path $appKey 'shell\open\command') -Name '(default)' -Value ('"' + $exe + '" "%1"')
    GarantirChave (Join-Path $appKey 'SupportedTypes')

    $cap = Join-Path $capRaiz ($primeiro.Projeto + '\Capabilities')
    GarantirChave (Join-Path $cap 'FileAssociations')
    Set-ItemProperty -Path $cap -Name 'ApplicationName' -Value $primeiro.Projeto
    Set-ItemProperty -Path $cap -Name 'ApplicationDescription' -Value ($primeiro.Projeto + ' (mOffice, build de desenvolvimento)')
    Set-ItemProperty -Path $cap -Name 'ApplicationIcon' -Value ('"' + $exe + '",0')
    foreach ($r in $app.Group) {
        New-ItemProperty -Path (Join-Path $appKey 'SupportedTypes') -Name $r.Ext -Value '' -PropertyType String -Force | Out-Null
        Set-ItemProperty -Path (Join-Path $cap 'FileAssociations') -Name $r.Ext -Value $r.ProgId
    }
    GarantirChave $registrados
    Set-ItemProperty -Path $registrados -Name ('mOffice.Dev.' + $primeiro.Projeto) -Value ('Software\mOffice.Dev\' + $primeiro.Projeto + '\Capabilities')
    Write-Host ("aplicativo registrado  {0} ({1})" -f $primeiro.Projeto, (($app.Group | ForEach-Object Ext) -join ' '))
}

# Avisa o Explorer: sem isto os icones e o duplo clique so' mudam depois de reiniciar a sessao.
Add-Type -Namespace mOfficeDev -Name Shell -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, uint f, System.IntPtr a, System.IntPtr b);' -ErrorAction SilentlyContinue
[mOfficeDev.Shell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)

Write-Host ''
Write-Host 'Pronto. No Explorer, o menu "Abrir com" de um .docx passa a listar mTexto (idem .xlsx/.pptx).'
if (-not $Padrao) { Write-Host 'O aplicativo PADRAO nao foi alterado. Para o duplo clique abrir nos apps: -Padrao.' }
Write-Host 'Para desfazer: tools\explorer\registrar-dev.ps1 -Remover'
