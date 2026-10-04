# Inventário de licenças do que o mOffice redistribui

Este inventário lista **tudo o que viaja no instalador** — o pacote que `tools/installer/montar-pacote.ps1`
monta (layout: `{app}\`, `{app}\motor\`, `{app}\WebView2Runtime\`, `{app}\licencas\`) — com a versão, a
licença e o arquivo de licença que acompanha o pacote em `{app}\licencas\` (no repositório,
`docs/licencas/`).

Levantado em 2026-09-27 sobre o pacote montado de verdade (3228 arquivos): nomes de DLL da raiz, do
`motor\converter`, as pastas do `motor`, o `deps.json` dos apps e os pacotes NuGet do build.

**Regra do teste:** cada linha da tabela abaixo (a primeira coluna, sem a formatação) tem de aparecer no
texto da tela "Sobre" (`DialogosTests.Sobre_ListaTodoOInventario`). Ao acrescentar um componente ao
pacote, acrescente a linha aqui **e** em `src/mOffice.Comum/Dialogos/Atribuicoes.txt`.

## Componentes

| Componente | Versão | Licença | Arquivo(s) em `licencas\` | Onde no pacote |
|---|---|---|---|---|
| mOffice (mTexto, mPlanilha, mSlides) | a do produto (`<Version>` de `src/Directory.Build.props`) | AGPL-3.0 | `LICENCA-mOffice-AGPL-3.0.txt` (cópia exata do `LICENSE` da raiz), `AVISO-mOffice.txt` | raiz: `mTexto.exe`, `mPlanilha.exe`, `mSlides.exe`, `mOffice.Comum/Nucleo/Conversor.dll`, `Ponte\` |
| ONLYOFFICE Desktop Editors | 9.4.0.129 (core `a016fc28…`; `allfontsgen.exe` do ONLYOFFICE Docs 9.4.0.129) | AGPL-3.0 com termos adicionais; marca da Ascensio System SIA; ilustrações e ícones CC BY-SA 4.0 | `LICENCA-ONLYOFFICE-AGPL-3.0.txt`, `AVISO-ONLYOFFICE.txt`, `MODIFICACOES-ONLYOFFICE.md` | `motor\converter` (x2t, DLLs do core, allfontsgen), `motor\editors` (sdkjs, web-apps, webext) |
| Bibliotecas JavaScript do ONLYOFFICE | jQuery 4.0.0, jQuery Browser 0.1.0, Backbone 1.6.0, Underscore 1.13.6, RequireJS 2.3.8, requirejs-text 2.0.16, Socket.IO 4.5.3, perfect-scrollbar, jquery-mousewheel 3.1.9, XRegExp 3.0.0, es6-promise, whatwg-fetch, Monaco Editor 0.52.2 | MIT (Bootstrap dropdown: Apache-2.0) | `AVISO-ONLYOFFICE-3DPARTYLICENSE.txt`; o aviso de cada biblioteca no cabeçalho do próprio `.js`; Monaco: `LICENSE` e `ThirdPartyNotices.txt` na própria pasta; `LICENCA-es6-promise-MIT.txt`, `LICENCA-whatwg-fetch-MIT.txt` (os únicos dois sem cabeçalho no `.js`) | `motor\editors\web-apps\vendor`, bundles do `web-apps` |
| Bibliotecas nativas do conversor ONLYOFFICE | as do core 9.4.0.129 (V8, Boost, FreeType, HarfBuzz, Crypto++, Gumbo, Katana e CxImage confirmados em `Common/3dParty`/`3DPARTY.md` do core na tag `v9.4.0.129`; libxml2, zlib, libpng, libjpeg identificados por texto nos binários mas sem pasta dedicada nesse repositório — ver nota abaixo) | várias (BSD-3-Clause, BSL-1.0, FTL, MIT/"Old MIT", Apache-2.0, licença própria do CxImage, zlib, PNG Reference Library License, IJG) | `LICENCA-3P-V8.txt`, `LICENCA-3P-Boost.txt`, `LICENCA-3P-FreeType.txt`, `LICENCA-3P-HarfBuzz.txt`, `LICENCA-3P-CryptoPP.txt`, `LICENCA-3P-Gumbo.txt`, `LICENCA-3P-Katana.txt`, `LICENCA-3P-CxImage.txt`, `LICENCA-3P-libxml2.txt`, `LICENCA-3P-zlib.txt`, `LICENCA-3P-libpng.txt`, `LICENCA-3P-libjpeg.txt` | `motor\converter\*.dll`, `x2t.exe`; FreeType/HarfBuzz/zlib também em `sdkjs\common\libfont` e `zlib` (WebAssembly) |
| ICU | 74.2 (`icuuc74.dll`, `icudt74.dll`) | Unicode-3.0 (Unicode License V3) | `LICENCA-ICU-Unicode.txt` (agora com o `LICENSE` completo do icu4c 74.2, tag `release-74-2`) | `motor\converter` (+ `icudtl.dat`, `icudtl_extra.dat`) |
| Hunspell | o do core ONLYOFFICE 9.4.0.129 (compilado em WebAssembly) | MPL-1.1 / GPL-2.0+ / LGPL-2.1+ (o mOffice usa a LGPL-2.1+) | `LICENCA-WeCantSpell-Hunspell-MPL-GPL-LGPL.txt`, `LICENCA-LGPL-2.1.txt` | `motor\editors\sdkjs\common\spell` |
| WeCantSpell.Hunspell | 7.0.1 | MPL-1.1 / GPL-2.0+ / LGPL-2.1+ (o mOffice usa a LGPL-2.1+) | `LICENCA-WeCantSpell-Hunspell-MPL-GPL-LGPL.txt`, `LICENCA-LGPL-2.1.txt` | raiz: `WeCantSpell.Hunspell.dll` |
| Dicionário pt_BR (VERO) | o do ONLYOFFICE 9.4.0.129 (dicionário VERO de Raimundo Moura; hifenização DivSilab 1.0.8) | dicionário LGPL-2.1; hifenização LGPL-3.0 ou MPL (o mOffice usa a LGPL-3.0) | `LICENCA-Dicionarios-pt_BR.txt`, `LICENCA-LGPL-2.1.txt`, `LICENCA-LGPL-3.0.txt`, `LICENCA-GPL-3.0.txt` (e os README originais na própria pasta) | `motor\dictionaries\pt_BR` |
| Dicionário en_US (SCOWL) | 2020.12.07 | licença do SCOWL (permissiva, com aviso WordNet); hifenização estilo BSD | `LICENCA-Dicionarios-en_US.txt` (e os originais na própria pasta) | `motor\dictionaries\en_US` |
| Open Sans | a do ONLYOFFICE 9.4.0.129 | Apache-2.0 | `LICENCA-FONTE-OPENSANS-APACHE-2.0.txt` | `motor\fonts` |
| Asana Math | a do ONLYOFFICE 9.4.0.129 | OFL-1.1 | `LICENCA-FONTE-ASANA-MATH-OFL-1.1.txt` | `motor\fonts\asana` |
| Caladea | a do ONLYOFFICE 9.4.0.129 | OFL-1.1 | `LICENCA-FONTE-CALADEA-OFL-1.1.txt` | `motor\fonts\caladea` |
| Carlito | a do ONLYOFFICE 9.4.0.129 | OFL-1.1 | `LICENCA-FONTE-CARLITO-OFL-1.1.txt` | `motor\fonts\crosextra` |
| OpenSymbol | a do ONLYOFFICE 9.4.0.129 | OFL-1.0 | `LICENCA-FONTE-OPENSYMBOL-OFL-1.0.txt` | `motor\fonts\openoffice` |
| mOffice Simbolos | 2026-09-19 (`tools/fontes/mOffice-Simbolos.ttf`) | OFL-1.1 | `LICENCA-mOffice-Simbolos-OFL.txt`; decisões e evidência sobre as fontes do motor (e por que o `ASC.ttf` fica fora): `AVISO-fontes.md` | em base64 nos `sdkjs\{word,cell,slide,visio}\sdk-all.js` |
| Inter | 4.001 | OFL-1.1 | `LICENCA-Inter-OFL.txt` | embutida em `mOffice.Comum.dll` |
| Phosphor Icons | 2.1 | MIT | `LICENCA-Phosphor-MIT.txt` | embutida em `mOffice.Comum.dll` |
| Microsoft Edge WebView2 Runtime | 153.0.4234.48 (Fixed Version, x64) | Microsoft Software License Terms | `LICENCA-WebView2-Runtime.txt` (os avisos de terceiros do Chromium vêm dentro do próprio runtime: `WebView2Runtime\show_third_party_software_licenses.bat`) | `WebView2Runtime\` |
| Microsoft.Web.WebView2 (SDK) | 1.0.2903.40 | licença estilo BSD da Microsoft | `LICENCA-WebView2-SDK.txt` | raiz: `Microsoft.Web.WebView2.*.dll`, `WebView2Loader.dll`, `runtimes\win-*\native` |
| .NET 10 (runtime) | 10.0.10 (Microsoft.NETCore.App + Microsoft.WindowsDesktop.App) | MIT | `LICENCA-dotnet-runtime-MIT.txt`, `AVISO-dotnet-THIRD-PARTY-NOTICES.txt` (Microsoft.NETCore.App), `LICENCA-WindowsDesktop-THIRD-PARTY-NOTICES.txt` (Microsoft.WindowsDesktop.App, dotnet/wpf + dotnet/winforms tag `v10.0.10`) | raiz (self-contained): `coreclr.dll`, `System.*.dll`, WPF, `msquic.dll`, `createdump.exe`, satélites de idioma |
| CommunityToolkit.Mvvm | 8.4.2 | MIT | `LICENCA-CommunityToolkit.Mvvm-MIT.txt` | raiz: `CommunityToolkit.Mvvm.dll` |

## Fonte correspondente (AGPL-3.0 §6)

`FONTE-CORRESPONDENTE.md` aponta a fonte exata do que vai sob a AGPL: o repositório público do mOffice e,
do ONLYOFFICE 9.4.0.129, as tags e os commits de `core`, `sdkjs`, `web-apps` e demais repositórios, os
pacotes oficiais de onde vêm os arquivos do motor (link e sha256), as modificações (`MODIFICACOES-ONLYOFFICE.md`)
e como remontar o pacote.

## Textos integrais de licença (referenciados pela tabela)

`LICENCA-LGPL-2.1.txt`, `LICENCA-LGPL-3.0.txt` e `LICENCA-GPL-3.0.txt` são os textos da FSF copiados
**verbatim** das pastas de dicionários do próprio ONLYOFFICE 9.4.0.129 (`de_AT/COPYING_LGPL_v2.1.txt`,
`es_ES/LGPLv3.txt`, `es_ES/GPLv3.txt`). A GPL-3.0 acompanha a LGPL-3.0 porque a própria LGPL-3.0
(§4(b)) exige.

## Origem de cada texto

| Arquivo | De onde veio |
|---|---|
| `LICENCA-ICU-Unicode.txt` | `icu4c/LICENSE` da tag `release-74-2` em github.com/unicode-org/icu, verbatim (https://raw.githubusercontent.com/unicode-org/icu/release-74-2/LICENSE, obtido em 2026-09-27) |
| `LICENCA-Dicionarios-pt_BR.txt` | `README_pt_BR.TXT`, `README_hyph_pt_BR.txt` (ISO-8859-1 → UTF-8, texto intacto) e `license-thes.readme` da pasta `pt_BR` do motor |
| `LICENCA-Dicionarios-en_US.txt` | `README_en_US.txt`, `README_hyph_en_US.txt`, `WordNet_license.txt` e `license.txt` da pasta `en_US` do motor |
| `LICENCA-dotnet-runtime-MIT.txt` | `LICENSE.TXT` / `LICENSE` dos pacotes de runtime NuGet 10.0.10 |
| `AVISO-dotnet-THIRD-PARTY-NOTICES.txt` | `THIRD-PARTY-NOTICES.TXT` do `microsoft.netcore.app.runtime.win-x64` 10.0.10 |
| `LICENCA-WindowsDesktop-THIRD-PARTY-NOTICES.txt` | `THIRD-PARTY-NOTICES.TXT` de github.com/dotnet/wpf e github.com/dotnet/winforms, ambos na tag `v10.0.10` (a mesma versão do runtime .NET 10 embarcado), obtidos em 2026-09-27 |
| `LICENCA-CommunityToolkit.Mvvm-MIT.txt` | `License.md` e `ThirdPartyNotices.txt` do pacote NuGet 8.4.2 |
| `LICENCA-WebView2-SDK.txt` | `LICENSE.txt` e `NOTICE.txt` do pacote NuGet 1.0.2903.40 |
| `AVISO-ONLYOFFICE-3DPARTYLICENSE.txt` | `3DPARTYLICENSE.txt` da instalação oficial do ONLYOFFICE Desktop Editors 9.4.0.129 (versão conferida no `DesktopEditors.exe`; sha256 no cabeçalho do arquivo) |
| `LICENCA-WebView2-Runtime.txt` | caixa de termos da página de download do runtime Fixed Version, obtida em 2026-09-18 para a versão pinada 153.0.4234.48 (a mesma de `third_party/webview2/VERSAO.txt` e do `.cab` com sha256 conferido) |
| `LICENCA-mOffice-AGPL-3.0.txt` | cópia exata do `LICENSE` da raiz do repositório (um teste confere) |
| `LICENCA-3P-V8.txt` | `LICENSE` de github.com/v8/v8, branch `main` (o `Common/3dParty/v8/v8.pri` do core, tag `v9.4.0.129`, aponta para esse repositório; `3DPARTY.md` do core confirma "3-clause BSD License"), obtido em 2026-09-27 |
| `LICENCA-3P-Boost.txt` | `LICENSE_1_0.txt` de github.com/boostorg/boost, tag `boost-1.72.0` (versão pinada em `Common/3dParty/boost/boost.pri` do core, tag `v9.4.0.129`: `-1_72`), obtido em 2026-09-27 |
| `LICENCA-3P-FreeType.txt` | `docs/FTL.TXT` de github.com/freetype/freetype, branch `master` (referência do `3DPARTY.md` do core, tag `v9.4.0.129`; versão exata do FreeType não fixada em `Common/3dParty`), obtido em 2026-09-27 |
| `LICENCA-3P-HarfBuzz.txt` | `COPYING` de github.com/harfbuzz/harfbuzz no commit `894a1f72ee93a1fd8dc1d9218cb3fd8f048be29a` (pinado por `Common/3dParty/harfbuzz/make.py` do core, tag `v9.4.0.129`), obtido em 2026-09-27 |
| `LICENCA-3P-CryptoPP.txt` | `Common/3dParty/cryptopp/License.txt` do próprio repositório ONLYOFFICE/core, tag `v9.4.0.129` (código vendorizado por inteiro nessa tag; versão 8.7.0 lida de `config_ver.h`), obtido em 2026-09-27 |
| `LICENCA-3P-Gumbo.txt` | `COPYING` de github.com/google/gumbo-parser no commit `aa91b27b02c0c80c482e24348a457ed7c3c088e0` (pinado por `Common/3dParty/html/fetch.py` do core, tag `v9.4.0.129`), obtido em 2026-09-27 |
| `LICENCA-3P-Katana.txt` | `LICENSE` de github.com/jasenhuang/katana-parser no commit `be6df458d4540eee375c513958dcb862a391cdd1` (pinado por `Common/3dParty/html/fetch.py` do core, tag `v9.4.0.129`), obtido em 2026-09-27 |
| `LICENCA-3P-CxImage.txt` | `license.txt` de github.com/movableink/cximage, branch `master` (referência do `3DPARTY.md` do core, tag `v9.4.0.129`), obtido em 2026-09-27 |
| `LICENCA-3P-libxml2.txt` | `Copyright` de github.com/GNOME/libxml2, branch `master` — versão não determinada (libxml2 não tem pasta dedicada em `Common/3dParty` nem entrada no `3DPARTY.md` do core; texto do tipo de licença estável usado, ver nota no arquivo), obtido em 2026-09-27 |
| `LICENCA-3P-zlib.txt` | `LICENSE` de github.com/madler/zlib, branch `master` — versão não determinada (mesma nota de libxml2), obtido em 2026-09-27 |
| `LICENCA-3P-libpng.txt` | `LICENSE` de github.com/glennrp/libpng, branch `libpng16` — versão não determinada (mesma nota), obtido em 2026-09-27 |
| `LICENCA-3P-libjpeg.txt` | seção "LEGAL ISSUES" de `README.ijg` de github.com/libjpeg-turbo/libjpeg-turbo, branch `main` (texto da licença IJG clássica) — versão não determinada (mesma nota), obtido em 2026-09-27 |
| `LICENCA-es6-promise-MIT.txt` | `LICENSE` de github.com/stefanpenner/es6-promise, branch `master`, obtido em 2026-09-27 |
| `LICENCA-whatwg-fetch-MIT.txt` | `LICENSE` de github.com/github/fetch (pacote npm `whatwg-fetch`), branch `main`, obtido em 2026-09-27 |

## NEEDS-SOURCE resolvido (2026-09-27/28)

Os 4 itens NEEDS-SOURCE da Task 4 foram resolvidos com os textos oficiais acima. Nota sobre o item 2:
`Common/3dParty` e `3DPARTY.md` do repositório ONLYOFFICE/core (tag `v9.4.0.129`) — a fonte autoritativa
citada pela Task 4 — confirmam V8, Boost, FreeType, HarfBuzz, Crypto++, Gumbo, Katana e CxImage como
componentes de terceiros do core, mas **não** listam libxml2, zlib, libpng ou libjpeg como componentes
próprios; essas quatro strings provavelmente vêm embutidas dentro do CxImage (que historicamente traz
cópias internas de codecs jpeg/png/zlib) ou de outra dependência transitiva não isolada nesta tarefa.
Ainda assim, cada uma tem seu próprio arquivo de licença (upstream oficial, versão mais próxima possível,
com a ressalva registrada no arquivo), para não deixar nenhuma string encontrada nos binários sem
cobertura. A origem exata de cada texto (repositório, tag ou commit, data) está na tabela "Origem de cada
texto", acima.
