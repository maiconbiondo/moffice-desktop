# Modificações na árvore redistribuída do ONLYOFFICE

Este arquivo lista, por exigência da AGPLv3 (§5 "Conveying Modified Source Versions" e a nota
da §6 sobre "Corresponding Source"), toda modificação que o PoC do mOffice faz dentro da árvore
extraída do ONLYOFFICE (`third_party/onlyoffice/`) — arquivo, data e motivo. Modificação aqui
significa: qualquer arquivo dentro de `third_party/onlyoffice/` que o mOffice cria, sobrescreve
ou altera, seja em tempo de build ou em tempo de execução. Isto é bookkeeping, não licença: a
licença do próprio ONLYOFFICE está em `third_party/onlyoffice/` (ou em `tools/onlyoffice/`, ver
o script de obtenção do motor).

**Desde 2026-09-19 o mOffice redistribui uma VERSÃO MODIFICADA do `sdkjs`:** os quatro
`editors/sdkjs/{word,cell,slide,visio}/sdk-all.js` são *patcheados* (editados no lugar, por script
reprodutível) para trocar a fonte de símbolos embutida — ver a linha de 2026-09-19 da tabela. Até essa
data nenhum `.js`/`.cpp`/`.xml` do motor tinha o conteúdo alterado; existiam só a cópia do `AllFonts.js`,
a omissão do `ASC.ttf` e o acréscimo do `allfontsgen.exe`. O aviso de modificação com data que os termos
adicionais da AGPL exigem está em `docs/licencas/AVISO-ONLYOFFICE.txt`.

**Resumo em 2026-09-27 (Plano 5, release):** a **única modificação de conteúdo** em arquivo do motor que o
mOffice redistribui é a troca da fonte de símbolos nos quatro `sdk-all.js` (linha de 2026-09-19). O resto
da tabela é acréscimo (`allfontsgen.exe`, binário AGPL sem modificação), omissão (`ASC.ttf`) ou histórico
(a cópia do `AllFonts.js`, encerrada em 2026-09-27). O que fica **fora** do pacote está em "Arquivos não
redistribuídos", e o que é código do mOffice (não do ONLYOFFICE) em "Código do mOffice", abaixo.

## Lista

| Data | Arquivo afetado (dentro de `third_party/onlyoffice/`) | O que o mOffice faz | Por quê |
|---|---|---|---|
| 2026-09-17 (**encerrada em 2026-09-27**) | `editors/sdkjs/common/AllFonts.js` | **Encerrada em 2026-09-27 (Plano 5, Task 1):** o mOffice **não copia mais nada** para dentro do motor. O editor recebe o catálogo do host, pela URL `https://moffice-fontes.local/AllFonts.js` (`CatalogoNoEditor`: o `src` do `<script>` que o requirejs cria é reescrito para esse host, que o servidor local da janela responde lendo `<PastaFontes>\AllFonts.js`), e o arquivo **não existe** no motor do pacote (o `montar-pacote.ps1` confere). O texto a seguir é o histórico. Cópia (não patch): o mOffice copia `AllFonts.js` de `<PastaFontes>/AllFonts.js` para dentro do motor copiado, em `editors/sdkjs/common/AllFonts.js`, toda vez que a janela do host inicializa (`CatalogoFontes.GarantirNoMotor`, chamado de `JanelaEditor.InicializarAsync`). **Atualizado em 2026-09-18 (Task 2 do Marco 1): a ORIGEM mudou** — até então era o catálogo que o *instalador oficial do ONLYOFFICE Desktop Editors* gera em `%LOCALAPPDATA%\ONLYOFFICE\DesktopEditors\data\fonts`; agora é o catálogo que o **próprio mOffice gera** em `%ProgramData%\mOffice\fontes` (ou `%LOCALAPPDATA%\mOffice\fontes`). O destino e o mecanismo são os mesmos. | O `sdkjs/common/AllFonts.js` não vem no motor extraído desta versão (9.4.0.129) — sem ele o `requirejs` do editor morre no módulo `"allfonts"` antes de qualquer coisa carregar (medido na sonda da Task 4). O arquivo não pode ser servido por interceptação (o que `SetVirtualHostNameToFolderMapping` mapeia não passa pelo `WebResourceRequested`, achado 2 da mesma sonda), então ele precisa existir de fato no disco, dentro da pasta mapeada. O conteúdo copiado é gerado pelo mOffice (era gerado pelo app oficial até 2026-09-18) — o mOffice só move o arquivo de um lugar para o outro. |
| 2026-09-18 | `fonts/ASC.ttf` (**omissão**, não modificação de arquivo existente) | O `tools/onlyoffice/obter-motor.ps1` passou a **não copiar** `fonts\ASC.ttf` (família `ASCW3`, 6.868 B, sha256 `35ff902a…4c1e18`) da instalação oficial — `robocopy /XF ASC.ttf` na pasta `fonts`, com guarda que confere o resultado da cópia e falha se o arquivo aparecer no destino. `-ComAscTtf` recopia a fonte, só para repetir o experimento. Nenhum arquivo do motor foi editado. | A fonte é o único arquivo do pacote **sem licença localizável** (`All rights reserved` na própria tabela `name`). O experimento A/B da Task 4/S3 mediu que ela é **dispensável**: capturas da janela real byte a byte idênticas com e sem o arquivo (0 pixels de 961.904, com e sem o ¶ ligado), PDF do `x2t` idêntico exceto os 8 dígitos do carimbo de data, e a suíte sem regressão. Motivo mecânico: os **mesmos bytes** da fonte já vêm embutidos em `base64` dentro de `sdkjs/{word,cell,slide,visio}/sdk-all.js`, e é de lá que o editor a carrega. Evidência completa e a pendência que isso deixa aberta: `docs/licencas/AVISO-fontes.md` §1. |
| 2026-09-19 | `editors/sdkjs/word/sdk-all.js`, `editors/sdkjs/cell/sdk-all.js`, `editors/sdkjs/slide/sdk-all.js`, `editors/sdkjs/visio/sdk-all.js` (**PATCH de verdade** — arquivos do motor editados no lugar) | `tools/onlyoffice/aplicar-patches.ps1` (chamado no fim de `obter-motor.ps1`, e executável sozinho) localiza em cada bundle, **pelo conteúdo** (decodifica os literais base64 que começam como TrueType e confere o sha256), o literal `"AAEAAAAPAIAAAwBwRFNJRw…"` de 9.160 caracteres — a fonte de símbolos internos **`ASCW3`** da Ascensio (6.868 B decodificados, sha256 `35ff902a…4c1e18`, idêntica a `fonts/ASC.ttf`) — e o **troca** pelo base64 (4.028 caracteres) de **`tools/fontes/mOffice-Simbolos.ttf`** (3.020 B, sha256 `3c5066cc9d89f5df74494d0bfe7d4dce354b4a090e6c8259dd79d6bc1f54631e`), fonte **desenhada em código** pelo mOffice (`tools/fontes/construir-simbolos.py`, saída determinística) e licenciada **SIL OFL 1.1** (`docs/licencas/LICENCA-mOffice-Simbolos-OFL.txt`). Nada mais muda no arquivo (conferido byte a byte: só o trecho entre as aspas; −5.132 bytes por bundle). O original fica em `<bundle>.orig` ao lado (ignorado pelo git) e `-Reverter` restaura. Guarda de versão: só patcha o motor `9.4.0.129`. sha256 dos bundles antes → depois: word `53ae764d…2109` → `73067860…f0f7`; cell `5be774a7…b1b9` → `a058035a…239f`; slide `5331132a…0e83` → `139e3ab7…f811`; visio `ef4301a0…1523` → `5adb16ac…e8e`. | Era a **única** forma de a fonte sem licença localizável sair do que distribuímos (tirar o arquivo solto removeu só uma das cinco cópias — `AVISO-fontes.md` §1.3/§1.7). A substituta tem os **mesmos** `cmap` (3/0 simbólica e 1/0 formato 6, por onde o editor chega nos glifos com `String.fromCharCode(33)` etc.), os mesmos 12 nomes de glifo, `unitsPerEm`, ascent/descent e larguras de avanço — o layout das marcas não muda; só o **desenho** é nosso. A família na tabela `name` é `mOffice Simbolos`: medido na janela real que o motor liga o binário pelo **índice** (`n.ASCW3=w`), não pelo nome — a string `"ASCW3"` no código do motor fica como está (identificador interno + nome de reserva que documentos salvos pelo ONLYOFFICE, e também pelo mOffice pela mesma serialização, carregam: no entendimento do projeto, compatibilidade de arquivo, não uso de marca). Evidência (A/B com controle negativo de fonte vazia, capturas, testes): `AVISO-fontes.md` §1.5 e `tests/mOffice.Nucleo.Tests/FonteSimbolosTests.cs`. |
| 2026-09-18 | `converter/allfontsgen.exe` (**acréscimo**, não modificação de arquivo existente) | O `tools/onlyoffice/obter-allfontsgen.ps1` extrai o `allfontsgen.exe` (32.752 bytes, versão **9.4.0.129**, sha256 `ec1e8889…75bb`) do instalador do **ONLYOFFICE Docs (Document Server) para Windows** — `onlyoffice-documentserver.exe`, também 9.4.0.129 — **sem instalar nada**, e o copia para `third_party/onlyoffice/converter/`, ao lado do `x2t.exe`. Só o `.exe` é copiado; nenhuma DLL do motor é sobrescrita. Proveniência (versão + sha256 do gerador e do pacote) registrada em `third_party/onlyoffice/VERSAO.txt`. | O pacote do Desktop Editors 9.4.0.129 **não traz** o `allfontsgen.exe` (o app oficial gera o catálogo por dentro, com o mesmo código linkado no binário principal), e sem ele o mOffice não conseguia gerar `AllFonts.js`/`font_selection.bin`/`fonts_thumbnail*.png` — ou seja, **dependia do ONLYOFFICE estar instalado na máquina**. O binário do Document Server é do **mesmo `core` AGPL-3.0 e da mesma versão**: as 19 DLLs do core dos dois pacotes diferem por exatamente 8.688 bytes cada (a assinatura Authenticode do pacote do desktop), ou seja é o mesmo código. É redistribuição de binário AGPL sem modificação — o `core` correspondente é o mesmo já apontado em `VERSAO.txt` (tag `v9.4.0.129`). |

## O que **não** está nesta lista (e por quê)

- **O catálogo de fontes gerado pelo mOffice** (`%ProgramData%\mOffice\fontes` ou
  `%LOCALAPPDATA%\mOffice\fontes`: `AllFonts.js`, `font_selection.bin`, `fonts_thumbnail*.png`):
  fica **fora** de `third_party/onlyoffice/`, é escrito pelo `allfontsgen` em tempo de execução e
  descreve as fontes **da máquina do usuário** — não é arquivo da árvore redistribuída. Só a cópia
  do `AllFonts.js` para dentro do motor está na tabela acima.
- **As pastas de fontes do sistema** (`C:\Windows\Fonts`, fontes do usuário, fontes da nuvem do
  Microsoft 365) e `third_party/onlyoffice/fonts`: são lidas, nunca escritas, pelo
  `ServidorAscDesktop`/`CatalogoFontes` — não é modificação da árvore redistribuída.
- **Contingência do brief da Task 5A** (`--allow-file-access-from-files` + patch textual trocando
  `ascdesktop://fonts/` por `file:///` em `sdk-all.js`): **não foi necessária** — o esquema nativo
  `ascdesktop://` funcionou sem patch nenhum (ver `task-5A-report.md`, §4). Se algum dia essa
  contingência for adotada, ela entra nesta lista como patch de verdade, com data e revisão do
  código afetado.
- **A troca do `ASC.ttf` embutido em `base64` dentro dos `sdk-all.js`**: **feita em 2026-09-19** —
  está na tabela acima como patch de verdade (a única forma de a fonte sem licença realmente sair do
  que distribuímos). O que continua **fora** desta lista a respeito dela: os `.orig` que o script
  deixa ao lado dos bundles (cópias intactas do original, ignoradas pelo git) e a fonte substituta em
  si (`tools/fontes/mOffice-Simbolos.ttf`), que é código do mOffice, não da árvore do ONLYOFFICE.
- **A string `"ASCW3"` no código do motor** (17 ocorrências em cada `sdk-all.js`): **não é alterada**.
  É o identificador interno pelo qual o `sdkjs` liga o binário embutido (por índice) e o nome de
  reserva que documentos salvos pelo ONLYOFFICE — e também pelo mOffice, pela mesma serialização —
  carregam. No entendimento do projeto, mantê-la é compatibilidade de arquivo, não uso de marca
  (`AVISO-fontes.md` §1.5).
- Qualquer coisa fora de `third_party/onlyoffice/` (o código do próprio mOffice, hoje em
  `src/mOffice.Nucleo` e `src/mOffice.Conversor` — no Marco 0 viviam em `poc/mOffice.Poc.Host` e
  `poc/mOffice.Poc.Converter` —, etc.) não é redistribuição do ONLYOFFICE e
  não entra aqui.

## Arquivos não redistribuídos (remoções, não modificações)

O instalador leva um **subconjunto** do motor extraído. Tirar um arquivo inteiro não altera nenhum
arquivo que vai junto; está listado aqui para quem comparar o pacote com o ONLYOFFICE original. Quem
monta e confere o subconjunto é `tools/installer/montar-pacote.ps1` (aborta se algum deles aparecer).

| Arquivo ou pasta (dentro do motor) | Por que fica de fora |
|---|---|
| `editors/sdkjs-plugins/` (a pasta inteira) | O mOffice não usa plugins do editor. |
| `editors/web-apps/apps/api/documents/preload.html` e `cache-scripts.html` | Páginas dormentes no mOffice que carregam o `AllFonts.js` por `<script src>` fixo no HTML, caminho que a entrega do catálogo pelo host não cobre. |
| `editors/sdkjs/common/AllFonts.js` | É o catálogo de fontes de **uma** máquina. O editor recebe o da máquina do usuário, gerado pelo mOffice, pelo host (linha de 2026-09-17, encerrada). |
| `fonts/ASC.ttf` | Sem licença localizável (linha de 2026-09-18 e `AVISO-fontes.md` §1). |
| `dictionaries/*` exceto `pt_BR` e `en_US` | Só os dois idiomas que o corretor do mOffice atende vão no pacote. Os dois vão com os arquivos de licença originais da pasta. |

## Código do mOffice (não é modificação do ONLYOFFICE)

A **ponte** (`src/mOffice.Nucleo/Ponte/`, os scripts injetados na página do editor e o objeto que responde
às chamadas `AscDesktopEditor` que o sdkjs faz ao hospedeiro) e o **host** (as janelas WPF, o
`WebView2`, o `ServidorAscDesktop`, o `CatalogoNoEditor`, o conversor que chama o `x2t.exe`, o corretor e o
resto de `src/`) são código escrito do zero pelo mOffice. Eles conversam com o motor pela interface que o
próprio motor expõe, mas não alteram arquivo algum dele. São distribuídos sob a mesma AGPL-3.0
(`LICENCA-mOffice-AGPL-3.0.txt`), com a fonte completa no repositório público do mOffice.

Os ajustes que o host faz **em tempo de execução**, dentro da página, também são código do mOffice e não
tocam em arquivo do motor: a reescrita do `src` do `AllFonts.js`, os scripts da ponte e as flags do
WebView2. Por isso eles não entram na tabela acima.

## Como manter isto atualizado

Toda vez que uma task tocar em algo dentro de `third_party/onlyoffice/` (cópia, patch, geração de
arquivo em tempo de execução), acrescentar uma linha na tabela acima **antes** do commit — não
depois, e não em lote no fim do marco.
