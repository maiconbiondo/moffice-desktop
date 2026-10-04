<#
.SYNOPSIS
    Define mTexto, mPlanilha e mSlides como apps PADRAO dos documentos do Office (as 17 extensoes marcadas "Rmm" em
    tools/installer/associacoes.json: .docx/.docm/.doc/.odt/.rtf/.dotx, .xlsx/.xlsm/.xls/.ods/.xltx, .pptx/.pptm/.ppsx/.ppt/.odp/.potx)
    na maquina, pela politica de associacoes do Windows. Porte do aplicar-padrao-pdf.ps1 do mPDF.

.DESCRIPTION
    Roda pelo mass-command do RMM (agente LocalSystem). O Windows 10/11 NAO deixa um app se tornar padrao
    sozinho (protecao UserChoice); o caminho suportado e a politica de associacoes:
      1. grava o XML de associacoes em %ProgramData%\mOffice\moffice-padrao-office.xml;
      2. aponta a politica HKLM\SOFTWARE\Policies\Microsoft\Windows\System\DefaultAssociationsConfiguration
         para esse XML.
    O Windows entao FORCA essas extensoes a cada logon, para TODOS os usuarios (sobrepoe a escolha
    manual). Vale no PROXIMO logon.

    A politica e UM valor so, que aponta para UM arquivo. Se ela ja apontar para outro XML (ex.: o do mPDF,
    com .pdf -> mPDF), as associacoes de la que NAO sao das extensoes do mOffice sao COPIADAS para o
    XML do mOffice - sobrescrever o valor sem isso apagaria o padrao de .pdf em silencio. Ordem recomendada:
    aplicar o do mPDF ANTES deste (o do mPDF sobrescreve, nao junta).

    A lista de extensoes vem de tools/installer/associacoes.json (GERADO da TabelaFormatos pelo teste AssociacoesGeradas_EmDia): o
    script le' esse arquivo (-Lista), nada de lista literal repetida aqui. Rodando o script fora do repositorio, leve o JSON junto.

    Pre-requisito: mOffice instalado com a task "associar" (registra os ProgIDs). Sem eles o script
    sai com 1 e nao mexe em nada: apontar a politica para ProgIDs inexistentes faz o Windows desfazer o
    padrao do usuario.

    Idempotente. Nao mexe em outras associacoes. Desfazer: remover-padrao-office.ps1.
    ASCII puro, saida ASCII (o agente do RMM le a saida sem saber a codificacao).

.PARAMETER ChavePolitica
    So para teste: a chave do valor DefaultAssociationsConfiguration (padrao: a da politica do Windows).
.PARAMETER Pasta
    So para teste: onde gravar o XML (padrao: %ProgramData%\mOffice).
.PARAMETER ChaveClasses
    So para teste: onde procurar os ProgIDs (padrao: HKLM:\SOFTWARE\Classes).
.PARAMETER Lista
    A lista de associacoes (padrao: tools/installer/associacoes.json, ao lado deste script no repositorio).
#>
[CmdletBinding()]
param(
    [string]$ChavePolitica = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System',
    [string]$Pasta = (Join-Path $env:ProgramData 'mOffice'),
    [string]$ChaveClasses = 'HKLM:\SOFTWARE\Classes',
    [string]$Lista = ''
)
$ErrorActionPreference = 'Stop'

$valor = 'DefaultAssociationsConfiguration'
$xmlPath = Join-Path $Pasta 'moffice-padrao-office.xml'
# As MESMAS linhas de tools/rmm/moffice-padrao-office.xml e dos ProgIDs do moffice.iss: as que o JSON gerado da TabelaFormatos marca
# com Rmm (InstaladorTests.AssociacoesGeradas_EmDia).
# $PSScriptRoot no corpo (nao no padrao do param, que o PowerShell 5.1 avalia vazio com -File).
if (-not $Lista) { $Lista = Join-Path $PSScriptRoot '..\installer\associacoes.json' }
if (-not (Test-Path -LiteralPath $Lista)) {
    Write-Output "mOffice: a lista de associacoes ($Lista) nao foi encontrada. Nada foi alterado."
    exit 1
}
$nossas = @((Get-Content -LiteralPath $Lista -Raw | ConvertFrom-Json) | Where-Object { $_.Rmm } | ForEach-Object {
    @{ Ext = [string]$_.Ext; ProgId = [string]$_.ProgId; App = [string]$_.App }
})
if ($nossas.Count -eq 0) {
    Write-Output "mOffice: a lista de associacoes ($Lista) nao traz nenhuma extensao para o padrao. Nada foi alterado."
    exit 1
}
$nossasExt = @($nossas | ForEach-Object { $_.Ext })

$faltando = @($nossas | Where-Object { -not (Test-Path -LiteralPath (Join-Path $ChaveClasses $_.ProgId)) } | ForEach-Object { $_.ProgId })
if ($faltando.Count -gt 0) {
    Write-Output ("mOffice: ProgIDs ausentes (" + ($faltando -join ', ') + "). Instale o mOffice com a opcao de associacao de arquivos marcada. Nada foi alterado.")
    exit 1
}

# Associacoes de OUTROS apps que a politica atual ja traz (inclusive de uma execucao anterior deste script,
# que ja juntou as do mPDF no nosso XML).
$outras = New-Object System.Collections.Generic.List[System.Xml.XmlElement]
$atual = $null
if (Test-Path -LiteralPath $ChavePolitica) {
    $atual = (Get-ItemProperty -LiteralPath $ChavePolitica -Name $valor -ErrorAction SilentlyContinue).$valor
}
if ($atual -and (Test-Path -LiteralPath $atual)) {
    try {
        $existente = New-Object System.Xml.XmlDocument
        $existente.Load($atual)
        foreach ($a in @($existente.SelectNodes('/DefaultAssociations/Association'))) {
            # Nossa = extensao da lista OU ProgID mOffice.* (uma politica de versao antiga do mOffice nao deixa lixo para tras).
            if (($nossasExt -notcontains $a.GetAttribute('Identifier').ToLowerInvariant()) -and ($a.GetAttribute('ProgId') -notlike 'mOffice.*')) { $outras.Add($a) }
        }
    }
    catch {
        Write-Output "mOffice: a politica atual aponta para um XML ilegivel ($atual). Nada foi alterado."
        exit 1
    }
}

$doc = New-Object System.Xml.XmlDocument
[void]$doc.AppendChild($doc.CreateXmlDeclaration('1.0', 'UTF-8', $null))
$raiz = $doc.AppendChild($doc.CreateElement('DefaultAssociations'))
foreach ($a in $outras) { [void]$raiz.AppendChild($doc.ImportNode($a, $true)) }
foreach ($n in $nossas) {
    $e = $doc.CreateElement('Association')
    $e.SetAttribute('Identifier', $n.Ext)
    $e.SetAttribute('ProgId', $n.ProgId)
    $e.SetAttribute('ApplicationName', $n.App)
    [void]$raiz.AppendChild($e)
}

# Link (juncao/symlink) na pasta ou no XML: um usuario comum poderia apontar para outro lugar e este script
# (LocalSystem) gravaria la. Recusa sem mexer em nada.
foreach ($c in @($Pasta, $xmlPath)) {
    $item = Get-Item -LiteralPath $c -Force -ErrorAction SilentlyContinue
    if ($item -and ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
        Write-Output "mOffice: $c e um link (juncao ou link simbolico). Por seguranca, nada foi alterado."
        exit 1
    }
}
New-Item -ItemType Directory -Force -Path $Pasta | Out-Null
# Arquivo NOVO, nunca sobrescrito: um XML pre-criado por um usuario comum manteria o dono (e a ACL) dele, e
# esse usuario poderia trocar o padrao de todos depois. Criado aqui (LocalSystem), herda a ACL da pasta.
Remove-Item -LiteralPath $xmlPath -Force -ErrorAction SilentlyContinue
if (Test-Path -LiteralPath $xmlPath) {
    Write-Output "mOffice: nao foi possivel substituir $xmlPath. Nada foi alterado."
    exit 1
}
$config = New-Object System.Xml.XmlWriterSettings
$config.Indent = $true
# UTF-8 sem BOM (o parser de politica do Windows nao gosta de BOM - licao do mPDF).
$config.Encoding = New-Object System.Text.UTF8Encoding($false)
$w = [System.Xml.XmlWriter]::Create($xmlPath, $config)
try { $doc.Save($w) } finally { $w.Dispose() }

# NUNCA New-Item -Force numa chave que ja existe: o provedor a recria VAZIA (licao do registrar-dev.ps1).
if (-not (Test-Path -LiteralPath $ChavePolitica)) { New-Item -Path $ChavePolitica -Force | Out-Null }
Set-ItemProperty -LiteralPath $ChavePolitica -Name $valor -Value $xmlPath -Type String

Write-Output ("mOffice: padrao de .docx .odt .xlsx .ods .pptx .odp aplicado via politica ($xmlPath; " + $outras.Count + " associacao(oes) de outros apps mantida(s)). Efetivo no proximo logon.")
exit 0
