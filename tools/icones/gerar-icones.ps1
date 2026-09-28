<#
.SYNOPSIS
  Regera os .ico dos tres aplicativos (mTexto, mPlanilha, mSlides) a partir dos logos vetoriais do tema.

.DESCRIPTION
  Fonte da verdade: src/mOffice.Comum/Temas/mOfficeTema.xaml (chaves Logo.mTexto/mPlanilha/mSlides).
  Dois passos, os dois em Python e sem GPU nenhuma:
    1. tools/icones/render-logos.py  — faz o PARSE daquele XAML e rasteriza cada logo em 256/48/32/16 px;
    2. tools/icones/gerar-ico.py     — empacota os quatro PNG num .ico de 4 quadros (formato do mpdf.ico).

  O desenho original (plano da Task 6) rasterizava pelo WPF, com `mOffice.Sonda --gerar-icones` +
  RenderTargetBitmap. Trocado porque RenderTargetBitmap devolve bitmap TODO ZERADO em sessao RDP (medido:
  ate um DrawRectangle solido sai em branco, tambem em PowerShell puro) — um gerador que escreve icone em
  branco em silencio nao serve. Ver o cabecalho de render-logos.py e o relatorio da Task 6.

  Os .ico sao COMMITADOS: este script so' roda quando o logo do tema muda. Requer Python com
  **Pillow >= 10** (`pip install "pillow>=10"`); o render-logos.py verifica a versao e para se for menor.

.PARAMETER Saida
  Pasta onde gravar os .ico. Sem este parametro, grava em src/<App>/Assets/<app>.ico (regeneracao de
  verdade); com ele, grava na pasta dada — e' o que o teste de reproducao (IconesTests) usa para comparar
  com o que esta' commitado sem sobrescrever nada.
#>
[CmdletBinding()]
param([string]$Saida)

$ErrorActionPreference = 'Stop'
$raiz = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$temp = Join-Path ([System.IO.Path]::GetTempPath()) ("mOffice-icones-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
try {
    Write-Host "Rasterizando os logos do tema em $temp ..."
    # Sem 2>&1 em nativo (licao do Plano 1): no Windows PowerShell 5.1 isso transforma cada linha de stderr
    # num ErrorRecord e derruba $? mesmo com codigo de saida 0.
    python (Join-Path $raiz 'tools\icones\render-logos.py') $temp
    if ($LASTEXITCODE -ne 0) { throw "render-logos.py falhou ($LASTEXITCODE) - Pillow >= 10 instalado? (pip install 'pillow>=10')" }

    $apps = @(
        @{ Projeto = 'mTexto';    App = 'mtexto' },
        @{ Projeto = 'mPlanilha'; App = 'mplanilha' },
        @{ Projeto = 'mSlides';   App = 'mslides' }
    )
    foreach ($a in $apps) {
        if ($Saida) {
            if (-not (Test-Path $Saida)) { New-Item -ItemType Directory -Path $Saida | Out-Null }
            $destino = Join-Path (Resolve-Path $Saida).Path ($a.App + '.ico')
        } else {
            $destino = Join-Path $raiz ('src\' + $a.Projeto + '\Assets\' + $a.App + '.ico')
        }
        python (Join-Path $raiz 'tools\icones\gerar-ico.py') $temp $a.App $destino
        if ($LASTEXITCODE -ne 0) { throw "gerar-ico.py falhou para $($a.App) ($LASTEXITCODE)" }
    }
} finally {
    Remove-Item -Recurse -Force $temp -ErrorAction SilentlyContinue
}
Write-Host "Pronto."
