<#
.SYNOPSIS
    Desfaz a politica que forca mTexto/mPlanilha/mSlides como padrao (par de aplicar-padrao-office.ps1).
    Porte do remover-padrao-pdf.ps1 do mPDF.

.DESCRIPTION
    Roda pelo mass-command do RMM (LocalSystem). Tira do XML do mOffice as associacoes do mOffice (as de ProgID mOffice.*, que
    sao as 17 do padrao e as seis de antes):
      - se sobrar associacao de outro app (o aplicar juntou, ex.: .pdf -> mPDF), o XML fica so com elas e a
        politica continua apontando para ele (o padrao de .pdf nao cai junto);
      - se nao sobrar nada, remove o VALOR DefaultAssociationsConfiguration (nunca a chave, que pode ter
        outras politicas) e apaga o XML.
    Se a politica apontar para OUTRO arquivo (ex.: o aplicar do mPDF rodou depois e sobrescreveu), ela nao
    e do mOffice e nao e tocada; so o XML orfao do mOffice e apagado.
    Efetivo no proximo logon. Nao desinstala o mOffice nem mexe em outras associacoes.

    Idempotente: sem politica, sai 0 sem erro. ASCII puro, saida ASCII.

.PARAMETER ChavePolitica
    So para teste: a chave do valor DefaultAssociationsConfiguration.
.PARAMETER Pasta
    So para teste: onde fica o XML (padrao: %ProgramData%\mOffice).
#>
[CmdletBinding()]
param(
    [string]$ChavePolitica = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System',
    [string]$Pasta = (Join-Path $env:ProgramData 'mOffice')
)
$ErrorActionPreference = 'Stop'

$valor = 'DefaultAssociationsConfiguration'
$xmlPath = Join-Path $Pasta 'moffice-padrao-office.xml'

# Link (juncao/symlink) na pasta ou no XML: nao le, nao grava e nao apaga atraves dele (mesma regra do aplicar).
foreach ($c in @($Pasta, $xmlPath)) {
    $item = Get-Item -LiteralPath $c -Force -ErrorAction SilentlyContinue
    if ($item -and ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
        Write-Output "mOffice: $c e um link (juncao ou link simbolico). Por seguranca, nada foi alterado."
        exit 1
    }
}

$atual = $null
if (Test-Path -LiteralPath $ChavePolitica) {
    $atual = (Get-ItemProperty -LiteralPath $ChavePolitica -Name $valor -ErrorAction SilentlyContinue).$valor
}

if (-not $atual -or -not [string]::Equals($atual, $xmlPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue
    if ($atual) {
        Write-Output "mOffice: a politica aponta para outro arquivo ($atual), nao e do mOffice - nada alterado nela."
    } else {
        Write-Output "mOffice: nenhuma politica de associacao ativa - nada a remover."
    }
    exit 0
}

$restantes = 0
if (Test-Path -LiteralPath $xmlPath) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($xmlPath)
    foreach ($a in @($doc.SelectNodes('/DefaultAssociations/Association'))) {
        if ($a.GetAttribute('ProgId') -like 'mOffice.*') {
            [void]$a.ParentNode.RemoveChild($a)
        } else {
            $restantes++
        }
    }
}

if ($restantes -gt 0) {
    $config = New-Object System.Xml.XmlWriterSettings
    $config.Indent = $true
    $config.Encoding = New-Object System.Text.UTF8Encoding($false)
    $w = [System.Xml.XmlWriter]::Create($xmlPath, $config)
    try { $doc.Save($w) } finally { $w.Dispose() }
    Write-Output "mOffice: associacoes do mOffice removidas; $restantes de outros apps mantida(s) na politica ($xmlPath). Efetivo no proximo logon."
} else {
    Remove-ItemProperty -LiteralPath $ChavePolitica -Name $valor -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue
    Write-Output "mOffice: politica de associacao removida. Efetivo no proximo logon."
}
exit 0
