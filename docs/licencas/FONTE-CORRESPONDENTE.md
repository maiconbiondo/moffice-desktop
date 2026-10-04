# Fonte correspondente do mOffice

Este arquivo cumpre a AGPL-3.0 (§6, "Corresponding Source"): aponta, com tag e commit exatos, a fonte de
tudo o que o instalador do mOffice redistribui sob a AGPL — o código do próprio mOffice e o motor de
edição do ONLYOFFICE — e diz como remontar o pacote a partir dela. Ele vai junto de cada versão (pasta
`licencas` da instalação) e está no repositório público
[github.com/maiconbiondo/moffice-desktop](https://github.com/maiconbiondo/moffice-desktop).

## 1. Código do mOffice

Todo o código escrito pelo mOffice — os três aplicativos (mTexto, mPlanilha, mSlides), a casca WPF, a
ponte com o editor (`src/mOffice.Nucleo/Ponte/`), o conversor que chama o `x2t.exe`, os scripts de
build e de obtenção do motor — está no repositório público, sob AGPL-3.0 (`LICENSE`). A fonte de uma
versão é a tag `v<versão>` desse repositório (a versão está em `src/Directory.Build.props`).

## 2. Motor ONLYOFFICE 9.4.0.129

O mOffice redistribui o motor do **ONLYOFFICE Desktop Editors 9.4.0.129** (Ascensio System SIA,
AGPL-3.0 com termos adicionais — `LICENCA-ONLYOFFICE-AGPL-3.0.txt` e `AVISO-ONLYOFFICE.txt`): o
conversor `x2t` e as bibliotecas do `core` (`motor\converter`), os editores `sdkjs` e `web-apps`
(`motor\editors`), as fontes livres (`motor\fonts`) e os dicionários `pt_BR` e `en_US`
(`motor\dictionaries`).

Tags e commits oficiais (repositórios em `github.com/ONLYOFFICE`, conferidos em 2026-09-28 pela API do
GitHub, `repos/ONLYOFFICE/<repositório>/git/ref/tags/<tag>`; todas são tags leves que apontam direto
para o commit):

| Repositório | O que vem dele no pacote | Tag | Commit |
|---|---|---|---|
| [`core`](https://github.com/ONLYOFFICE/core/tree/v9.4.0.129) | `x2t.exe`, DLLs do conversor, `allfontsgen.exe`, os módulos WebAssembly do `sdkjs` (fontes, zlib, corretor) | `v9.4.0.129` | `a016fc2868b9f99f59dd30aa2088e68bbe5104af` |
| [`sdkjs`](https://github.com/ONLYOFFICE/sdkjs/tree/v9.4.0.129) | `motor\editors\sdkjs` (os `sdk-all*.js` de texto, planilha, apresentação e diagrama) | `v9.4.0.129` | `75b53d504bfd81bfa35d59bb57dd9e56f7c9235d` |
| [`web-apps`](https://github.com/ONLYOFFICE/web-apps/tree/v9.4.0.129) | `motor\editors\web-apps` (interface dos editores) | `v9.4.0.129` | `1993a6d89731af0fcb5129ebf0ee380076fea5a9` |
| [`DesktopEditors`](https://github.com/ONLYOFFICE/DesktopEditors/releases/tag/v9.4.0) | o instalador oficial cujo programa instalado forneceu os arquivos do motor (seção 3). Este repositório só publica os instaladores: não tem tag `v9.4.0.129`; a release correspondente é `v9.4.0` | `v9.4.0` | `7b9cd379f428ce23c7ef9c93cbb93f1e3c75f613` |
| [`dictionaries`](https://github.com/ONLYOFFICE/dictionaries/tree/v9.4.0.129) | `motor\dictionaries\pt_BR` e `en_US` | `v9.4.0.129` | `d3223bbb777883db66ac3cd249f71c6ebdc992c7` |
| [`core-fonts`](https://github.com/ONLYOFFICE/core-fonts/tree/v9.4.0.129) | `motor\fonts` | `v9.4.0.129` | `7030c6681fb5bbed560675cb42422f91df15d5c9` |
| [`build_tools`](https://github.com/ONLYOFFICE/build_tools/tree/v9.4.0.129) | nada (são os scripts oficiais que compilam os repositórios acima) | `v9.4.0.129` | `cf4cac0846d03e2e00a6fec5d1ecd4e9463086d5` |

A casca oficial do Desktop Editors (`desktop-apps` e `desktop-sdk`: `DesktopEditors.exe`, `editors.exe`,
CEF, Qt, serviço de atualização) **não** é redistribuída — o mOffice usa a própria casca (seção 1). Para
referência, na mesma versão: `desktop-apps` `v9.4.0.129` = `3ad4e294ae5e25902ccc56088a760ad21a9d123a` e
`desktop-sdk` `v9.4.0.129` = `8f6d24bfe9a5e7ea9405266d341ed4187cf5823e`.

## 3. Pacotes oficiais de onde vêm os arquivos do motor

| Pacote oficial | Link | Tamanho | sha256 |
|---|---|---|---|
| ONLYOFFICE Desktop Editors 9.4.0.129, Windows x64, MSI | https://github.com/ONLYOFFICE/DesktopEditors/releases/download/v9.4.0/DesktopEditors_x64.msi | 615.632.896 bytes | `1db1d01eaeea7c494aca9bc5b853d1aaa376792e05408faf2533e72a5c67eae9` |
| ONLYOFFICE Docs (Document Server) 9.4.0.129, Windows — só o `allfontsgen.exe` sai dele | https://download.onlyoffice.com/install/documentserver/windows/onlyoffice-documentserver.exe (baixado em 2026-09-18; o link oficial aponta sempre para a versão mais nova) | 1.043.099.928 bytes | `7ba127f7d3e82eb6881302cb8e12f8549d1afee8d1f74aaaded231a26d10376e` |

- Os arquivos do motor foram **copiados do programa instalado** por esse MSI oficial (versão 9.4.0.129,
  conferida no `DesktopEditors.exe` da instalação), não extraídos do MSI; `tools/onlyoffice/obter-motor.ps1`
  copia só as pastas que o mOffice usa. O sha256 do MSI é o mesmo que o GitHub publica para esse arquivo na
  release `v9.4.0` e o mesmo da cópia arquivada junto da cópia do motor (`obter-motor.ps1 -ArquivarMsi`).
  O `x2t.exe` copiado tem sha256 `deb80914b5c8743be2373a31fe54c08e0422051546a9d8fea6a08688e913cffa`.
- O `allfontsgen.exe` (32.752 bytes, versão 9.4.0.129, sha256
  `ec1e88896002d8259b209ce7262d754b1d970f301b4bc28c17fe8eeb900e75bb`) é do mesmo `core` e da mesma
  versão; o pacote do Desktop Editors não o traz. `tools/onlyoffice/obter-allfontsgen.ps1` confere os dois
  sha256 acima antes de copiar. A release correspondente do repositório `DocumentServer` é `v9.4.0`
  (commit `fb73d33c85f59d1a5d2e5a5ed05388ebc2418337`).
- O runtime Microsoft Edge WebView2 (Fixed Version 153.0.4234.48) também vai no instalador, mas não é
  AGPL: é software da Microsoft, sob os termos em `LICENCA-WebView2-Runtime.txt`, que proíbem
  re-hospedá-lo. `tools/webview2/obter-runtime.ps1` baixa o `.cab` da própria Microsoft e confere o
  sha256 pinado no script.

## 4. O que o mOffice modifica no motor

Tudo o que o mOffice altera, acrescenta ou deixa de fora da árvore do ONLYOFFICE está em
[`MODIFICACOES-ONLYOFFICE.md`](MODIFICACOES-ONLYOFFICE.md), arquivo a arquivo e com data. Em resumo:

- **Única modificação de conteúdo:** nos quatro `editors/sdkjs/{word,cell,slide,visio}/sdk-all.js`, a
  fonte de símbolos embutida em base64 (`ASCW3`, sem licença localizável) é trocada pela fonte livre do
  mOffice, `tools/fontes/mOffice-Simbolos.ttf` (SIL OFL 1.1, gerada por
  `tools/fontes/construir-simbolos.py`). Quem aplica é
  [`tools/onlyoffice/aplicar-patches.ps1`](../../tools/onlyoffice/aplicar-patches.ps1), por conteúdo
  (sha256 do trecho), com verificação depois de gravar; `-Reverter` desfaz.
- **Acréscimo:** `converter/allfontsgen.exe`, binário AGPL sem modificação (seção 3).
- **Omissões:** `fonts/ASC.ttf`, `editors/sdkjs-plugins/`, `sdkjs/common/AllFonts.js`, duas páginas do
  `web-apps` e os dicionários fora `pt_BR`/`en_US` — a lista e os motivos estão em
  `MODIFICACOES-ONLYOFFICE.md`, seção "Arquivos não redistribuídos".

## 5. Como remontar o pacote

Pré-requisitos: Windows x64, .NET 10 SDK, Windows PowerShell 5.1 e, para o instalador, Inno Setup 6.
Os binários do motor e do runtime **não** ficam no git: os scripts abaixo os obtêm e os conferem.

1. **Compilar e testar o código** (não precisa do motor; os testes que precisam dele são pulados):
   `dotnet build mOffice.slnx` e `dotnet test mOffice.slnx`.
2. **Motor:** instalar o ONLYOFFICE Desktop Editors 9.4.0.129 pelo MSI da seção 3 (numa máquina ou VM de
   build) e rodar `tools/onlyoffice/obter-motor.ps1` (com `-ArquivarMsi` para arquivar e registrar o
   sha256 do MSI). Ele copia `converter`, `editors`, `fonts` e `dictionaries` para
   `third_party/onlyoffice/`, grava `VERSAO.txt` e roda `tools/onlyoffice/aplicar-patches.ps1`.
3. **Gerador do catálogo de fontes:** baixar o pacote do Document Server da seção 3 e rodar
   `tools/onlyoffice/obter-allfontsgen.ps1 -Pacote <caminho do .exe>` (extrai sem instalar, com
   `innounp` ou 7-Zip).
4. **Runtime WebView2:** `tools/webview2/obter-runtime.ps1`.
5. **Pacote e instalador:** `tools/installer/montar-pacote.ps1 -Saida <pasta>` monta a pasta
   instalável; `tools/installer/build-installer.ps1` roda a suíte, monta o pacote, confere que nenhum
   caminho da máquina de build ficou nos binários e gera `tools/installer/Output/mOffice-Setup-<versão>.exe`.
