# Fontes que o mOffice redistribui — decisões e evidência

Este arquivo registra as decisões sobre as fontes que vêm dentro da árvore do ONLYOFFICE
(`third_party/onlyoffice/fonts/`). O inventário completo, gerado da árvore, está em
`third_party/onlyoffice/INVENTARIO.md`; a licença de cada fonte que vai no pacote está em
`INVENTARIO.md` (nesta pasta). Aqui ficam só as fontes que exigiram **decisão**.

---

## 1. `fonts/ASC.ttf` (família `ASCW3`) — **fora do pacote** (2026-09-18, Marco 1 / S3) **e substituída nos bundles** (2026-09-19, §1.5)

### 1.1 Por que a fonte é um problema

Único arquivo de fonte do pacote sem licença localizável. A tabela `name` da própria fonte diz:

| nameID | Conteúdo |
|---|---|
| 0 (copyright) | `Copyright (c) Ascensio System SIA 2012-2014. All rights reserved` |
| 1 / 4 (família) | `ASCW3` |
| 7 (marca) | `ASCW3 is a trademark of Ascensio System SIA.` |

*All rights reserved* é o oposto de uma concessão. Não há `LICENSE` ao lado dela no pacote (o
`fonts/LICENSE.txt` é a Apache-2.0 do Open Sans); no repositório oficial `ONLYOFFICE/core-fonts`
ela está solta na raiz, enquanto toda outra família fica numa subpasta com o próprio `LICENSE.txt`,
e o repositório não tem `LICENSE` na raiz. Marcada como **NÃO REDISTRIBUIR** no Marco 0.

- sha256: `35ff902a3b9c9254ce3c57d610dad9e57008b8a3351691c463bee775644c1e18`
- 6.868 bytes, 12 glifos.

### 1.2 Para que o motor usa a família `ASCW3`

A fonte é **minúscula e de uso interno**: cobre só os símbolos que o editor desenha por conta
própria, nunca texto do documento. São 12 glifos, dos quais **9 alcançáveis** — `.notdef`, `.null`
e `uni000D` existem na `glyf` mas **não estão em nenhum `cmap`**. Mapeamentos reais:

| Subtabela `cmap` | Mapeamento | Onde aparece |
|---|---|---|
| 3/0 (simbólica) | `U+2610`→`uni2610`, `U+2611`→`uni2611` (caixa vazia / com visto) | Caixas de seleção de formulário |
| 3/0 | `U+25C9`→`uni25C9`, `U+25CB`→`circle` (olho-de-boi / círculo vazio) | Botões de opção (*radio*) de formulário |
| 3/0 | `U+F020`–`U+F022`, `U+F038`–`U+F039` → `uniF020`…`uniF039` | Marcas de formatação não imprimíveis |
| 1/0 (formato 6) | `U+0020`–`U+0022` e `U+0038`–`U+0039` → os mesmos `uniF0xx` | É por aqui que o código chega neles: ele desenha `String.fromCharCode(33)` etc., não o ponto de código da área privada |

**A marca de parágrafo `¶` (U+00B6) NÃO vem da ASCW3** — não está no `cmap` dela. O editor a
desenha com outra fonte. Da ASCW3 vêm, nas marcas de formatação, as **setas de tabulação** e as
**marcas de quebra**.

Referências no código do motor (offsets em bytes, versão 9.4.0.129):

- `editors/sdkjs/word/sdk-all.js` @ 4 764 473 e 4 780 023 — as marcas de **quebra** e de
  **tabulação**: o editor desenha `String.fromCharCode(56)`/`(57)` e `(33)`/`(34)` com
  `sC({wg:{xa:"ASCW3",…},Ub:10,…})`, **dentro de um `if (editor.Go)`** (`Go` é o
  `ShowParaMarks` deste build minificado — o botão ¶ da barra de ferramentas).
- `editors/sdkjs/word/sdk-all.js` @ 764 325 — serialização de controle de conteúdo para o
  conversor: `9745===ka&&9744===Y` (`U+2611`/`U+2610`) e `9673===ka&&9675===Y`
  (`U+25C9`/`U+25CB`); se a fonte pedida **não tiver** o glifo, o nome da fonte é trocado por
  `"ASCW3"`. Ou seja, ASCW3 é o **substituto de último recurso** do glifo da caixa/bolinha.
- `editors/sdkjs/word/sdk-all.js` @ 3 140 045 — a lista de fontes que o editor manda carregar
  **exclui** ASCW3 explicitamente (`"ASCW3"!==p.xa && …`): a família nunca é pedida ao servidor
  de fontes.
- `editors/sdkjs/word/sdk-all.js` @ 679 300 — ASCW3 entra na lista de famílias "de sistema" do
  seletor interno.

### 1.3 Achado central: **o arquivo é redundante — a fonte já vem embutida nos *bundles***

Na função que monta o catálogo de fontes do editor (`sdk-all.js` @ ~3 123 153), o objeto da
família ASCW3 é criado assim:

```js
for (var x = 0, D = "ASC.ttf", A = 0; A < w; A++) {
  var N = u[A];
  "ASCW3" === N[0] ? D = v[N[1]].Va : (/* demais famílias */);
}
u = new b(D);            // D = caminho vindo do AllFonts.js, ou o literal "ASC.ttf"
u.UB = 0;                //    estado = JÁ CARREGADA
p[w] = AscFonts.VLf("AAEAAAAPAIAAAwBwRFNJRw…");   // binário embutido em base64
u.wOb(w);                //    fluxo da fonte = esse binário
```

O caminho `D` serve só de identificador (`GetID`). O **binário** usado é sempre o `base64`
literal do bundle, e como `UB = 0` o carregador nunca vai ao disco. Decodificando esse base64:

| Bundle | Tamanho decodificado | sha256 |
|---|---|---|
| `editors/sdkjs/word/sdk-all.js` | 6.868 B | `35ff902a…4c1e18` |
| `editors/sdkjs/cell/sdk-all.js` | 6.868 B | `35ff902a…4c1e18` |
| `editors/sdkjs/slide/sdk-all.js` | 6.868 B | `35ff902a…4c1e18` |
| `editors/sdkjs/visio/sdk-all.js` | 6.868 B | `35ff902a…4c1e18` |

**Os mesmos bytes do `ASC.ttf`, em quatro bundles que também redistribuímos.**

> **Consequência, dita sem rodeio:** tirar `fonts/ASC.ttf` do pacote **não** encerra a pendência
> de licença — apenas remove **uma** das cinco cópias. Ver §1.7 — e §1.5, onde as outras quatro
> foram trocadas.

### 1.4 Experimento A/B (2026-09-18, sessão de console local)

Amostra sintética gerada pelo Word via COM (`poc/amostras/sinteticas/simbolos.docx`, fora do git):
lista com marcadores (• e dois marcadores Wingdings `F0D8`/`F0FC`), lista numerada, símbolos
`√ ∑ ≠ ≤ → ©`, a linha `☐ ☑ ◉ ○`, uma linha em Wingdings e outra em Symbol, duas tabulações, uma
equação (`OMaths`) e **dois controles de conteúdo de caixa de seleção** (um marcado, um não).

Procedimento, para cada metade: `third_party/onlyoffice/fonts/ASC.ttf` presente/ausente →
apagar `%ProgramData%\mOffice\fontes\impressao-digital.txt` para forçar a regeneração do catálogo
→ abrir o documento na janela real → capturar a página → ligar o ¶ pela barra de ferramentas →
capturar de novo.

**Resultado 1 — o que o motor enxerga.** Sonda em JS dentro do editor
(`AscFonts.ccb` / `AscFonts.iEc`):

| | índice da família | caminho declarado | estado | fluxo |
|---|---|---|---|---|
| com `ASC.ttf` | 385 de 386 | `…/third_party/onlyoffice/fonts/ASC.ttf` | `0` (carregada) | `0` |
| sem `ASC.ttf` | 385 de 386 | `ASC.ttf` (literal, sem pasta) | `0` (carregada) | `0` |

A família **continua existindo e carregada** sem o arquivo; só o caminho declarado vira o literal
de reserva. O catálogo gerado confirma: `AllFonts.js` passa de 1 para **0** menções a `ASCW3`,
e de 639 para 638 arquivos de fonte.

**Resultado 2 — o que aparece na tela.** Capturas em `poc/relatorios/`:

| Par | sha256 igual? | Pixels diferentes |
|---|---|---|
| `asc-com.png` × `asc-sem.png` (página normal) | **sim** | **0** de 961.904 |
| `asc-com-marcas.png` × `asc-sem-marcas.png` (com o ¶ ligado) | **sim** | **0** de 961.904 |

As capturas são **byte a byte idênticas**. Com o ¶ ligado veem-se os pilcrows `¶`, os pontos de
espaço `·` e as **setas de tabulação `→`**; as setas são os glifos `uniF021`/`uniF022` da ASCW3
(os pilcrows vêm de outra fonte, §1.2), e estão nas duas metades. As caixas de seleção (`☐`/`☒`) e a linha `☐ ☑ ◉ ○` também são idênticas.

**Resultado 3 — exportação para PDF pelo x2t.** É o caminho de risco de verdade: o `x2t` não tem
o `base64` do `sdkjs`, só conhece as fontes que o `AllFonts.js` gerado nomeia.

| | bytes | páginas | fontes embutidas |
|---|---|---|---|
| com `ASC.ttf` | 120.555 | 1 | Aptos ×2, Symbol, Wingdings, Cambria Math, Calibri, MS Gothic |
| sem `ASC.ttf` | 120.555 | 1 | (as mesmas 7) |

Os dois PDFs diferem em **8 bytes**, todos dígitos de minuto/segundo em `/CreationDate`,
`/ModDate` e nos campos XMP equivalentes. O conteúdo é idêntico, e **ASCW3 não é embutida em
nenhuma das duas** — o `x2t` resolve caixa e bolinha por Wingdings/Symbol antes de chegar ao
substituto.

**Resultado 4 — suíte.** `SeisFormatosTests` + `AmostrasReaisTests` sem o arquivo: 11 aprovados,
1 reprovado — o `AmostrasReaisTests(real-odp)`, que já falhava antes desta task. Nenhuma diferença
atribuível à fonte. **Atualização (Task 6, commit `b9cf819`):** a falha do `real-odp` não era regressão
do produto, e sim artefato do extrator ODF do próprio teste (apagava `<text:s/>`); com o extrator
corrigido, `AmostrasReaisTests` fecha **6/6**.

### 1.5 Substituída em 2026-09-19: a fonte `mOffice Simbolos` no lugar do blob embutido

Decisão do usuário (2026-09-19, `marco1-gate.md` §7 decisão 1): executar a opção (a) agora — o e-mail à
Ascensio (b) segue em paralelo, por cortesia e registro, não como condição. Branch `feat/marco1-ascw3-livre`.

**O que é a nossa fonte.** `tools/fontes/mOffice-Simbolos.ttf` — 3.020 bytes, sha256
`3c5066cc9d89f5df74494d0bfe7d4dce354b4a090e6c8259dd79d6bc1f54631e`, família `mOffice Simbolos`, SIL OFL 1.1
com *Reserved Font Name* "mOffice Simbolos" (`docs/licencas/LICENCA-mOffice-Simbolos-OFL.txt`), copyright
"Copyright (c) 2026 mOffice", sem nenhuma marca de terceiros na tabela `name` nem em byte algum do arquivo
(o teste de unidade procura `ASCW3`/`Ascensio` em ASCII e UTF-16). **Nenhum byte ou contorno do `ASC.ttf`
entra no repositório**: a fonte original foi apenas *lida* (fontTools) para saber quais glifos existem, o que
cada um representa e as métricas.

| O que é igual ao original (números que o motor pressupõe) | O que é nosso |
|---|---|
| `cmap` 3/0 simbólica (`U+2610/2611`, `U+25C9/25CB`, `U+F020–F022`, `U+F038–F039`) e `cmap` 1/0 formato 6 (`U+0020–0022`, `U+0038–0039` → os mesmos `uniF0xx`) — é por este último que o editor chega nos glifos (`String.fromCharCode(33)` etc.) | todos os contornos: setas de tabulação (cabeça triangular, haste fina), marcas de quebra (seta com barra vertical), caixa de seleção e o desenho do visto, círculo e olho-de-boi (arcos quadráticos nossos) |
| ordem e nomes dos 12 glifos (`.notdef`, `.null`, `uni000D` sem `cmap`, como no original) | tabela `name` inteira (família, copyright, OFL nos nameID 13/14) |
| `unitsPerEm` 2048, ascent 1491 / descent −431 / lineGap 307, e a largura de avanço de cada glifo (setas 1825, quebras 1135, caixa/círculo 1764, espaço 512) — o layout das marcas não muda | sem hinting (`cvt`/`fpgm`/`prep`) e sem `DSIG`: a fonte fica em 3 KB e é toda auditável |

**Como é gerada.** `python tools/fontes/construir-simbolos.py` (Python 3 + fontTools; `TTGlyphPen`, formas
geométricas simples desenhadas no próprio script). Saída **determinística** — `created`/`modified` do `head`
fixos, `recalcTimestamp` desligado — e o teste `Fonte_ReproduzivelPeloScript` regenera a fonte numa pasta
temporária e compara **byte a byte** com a commitada (idem para a variante `--vazia`, o controle negativo
abaixo). A fonte vive em `tools/fontes/`, ao lado do gerador e do patch que a consome, e **não** em
`src/mOffice.Nucleo/`: nenhum código C# a carrega em tempo de execução — o que o produto usa é o bundle já
patchado; ela é insumo de ferramenta, não recurso do assembly.

**Como é aplicada.** `tools/onlyoffice/aplicar-patches.ps1` (chamado no fim de `obter-motor.ps1`, e executável
sozinho): em cada um dos quatro `sdk-all.js`, acha **pelo conteúdo** os literais base64 que começam como
TrueType, decodifica, confere o sha256 `35ff902a…` e troca **só o miolo entre as aspas** pelo base64 da nossa
fonte; relê o arquivo e confere (0× a original, 1× a nossa, tamanho exato). Idempotente ("já aplicado"),
`-Reverter` restaura do `<bundle>.orig` guardado ao lado (fora do git — `third_party/` inteiro e `*.orig`),
guarda de versão: só patcha `VERSAO.txt` = 9.4.0.129 (motor novo → parar e re-medir). Conferido byte a byte
que nada fora do trecho muda (−5.132 bytes por bundle):

| Bundle | sha256 antes (motor 9.4.0.129 intacto) | sha256 depois |
|---|---|---|
| `word/sdk-all.js` | `53ae764d670ce687bc8422604511fd28f5ab28fab711a140a1ffa668afa12109` | `730678603b331d9b4a294bd3fe580cf71fabc867b1d85d891dfa34b6c9baf0f7` |
| `cell/sdk-all.js` | `5be774a78b87cb3f33d23fc945ef20b6989e422c84ac3b9ae66a6d9e2e51b1b9` | `a058035a3baefed5b7cc0f002e54ff0d2db069423b048d18c9f36f53cccf239f` |
| `slide/sdk-all.js` | `5331132ae920fd02d8485244ad308d05030e9eae27bd60d8f53c2f10e0ac0e83` | `139e3ab722c50dc76c7edc0886d966f65a0032d26b241b94d062f9ccac3cf811` |
| `visio/sdk-all.js` | `ef4301a0196c078be01b2e3a70dc73c1bb339b00460a22326c42a896d25e1523` | `5adb16ac0d2b503b49f944a6aa19a6e1cf7c7c47e2db0509cd7fd277bd184e8e` |

Isto torna o `sdkjs` redistribuído uma **versão modificada**: linha própria em `MODIFICACOES-ONLYOFFICE.md` e o
aviso de modificação com data em `AVISO-ONLYOFFICE.txt` (termos adicionais da AGPL).

**A/B novo (2026-09-19, sessão de console, `FonteSimbolosJanelaRealTests`).** Documento sintético
(`FabricaDocumentos.DocxMarcasDeFormatacao`: duas tabulações, uma quebra de linha, duas caixas de seleção
`w14:checkbox` em MS Gothic — uma marcada, uma não), aberto na janela real, zoom 300 % (as marcas são
desenhadas a 10 pt fixos pelo motor, `Ub:10`; o zoom é o que as engorda na captura), ¶ desligado → captura,
¶ ligado (`put_ShowParaMarks`) → captura. **Controle negativo:** a mesma sequência sobre um **motor
temporário** cujo único bundle diferente é o `word/sdk-all.js` com a fonte `--vazia` (mesmos cmaps, mesmas
métricas, **contornos vazios**) — junções para `converter`/`fonts`/`dictionaries`, cópia real de
`web-apps` + `sdkjs/common` + `sdkjs/word` (o mapeamento de pasta do WebView2 não segue junções), perfil
WebView2 próprio (mesma URL do bundle nos dois motores; um cache compartilhado contaminaria o controle).

| Medida (só a página, abaixo dos 20 % de cima) | Valor |
|---|---|
| Pixels diferentes entre ¶ ligado e desligado (nossa fonte) | **1.036** — as marcas ligaram |
| Pixels diferentes entre nossa fonte e fonte vazia, ¶ ligado | **296** de 769.776 |
| … dos quais com tinta (luma < 128) na **nossa** captura | **263** |
| … dos quais com tinta na captura da fonte **vazia** | **0** |

A máscara da diferença (`poc/relatorios/ascw3-livre-marcas-diff.png`) é **exatamente** a seta de tabulação
`→` e a marca de quebra `↵`, nada mais: os pixels das marcas vêm dos **nossos** contornos. Capturas:
`poc/relatorios/ascw3-livre-marcas.png` (¶ ligado), `ascw3-livre-formulario.png` (¶ desligado, caixas),
`ascw3-livre-marcas-controle-vazio.png` (controle). Mutação: retirar o glifo `uni2611` do script → a fonte
sai com 11 glifos e `Fonte_TemOsCmapsEGlifosEsperados`, a guarda do pacote e a reprodução falham; restaurado.

**Nome da família (item 5 da task): não precisa ser "ASCW3".** Medido: com a família `mOffice Simbolos` na
tabela `name`, as marcas aparecem. O motor cria a entrada por nome fixo no código (`"ASCW3"===N[0]`,
`n.ASCW3=w`) e liga o binário embutido pelo **índice** — a tabela `name` da fonte não entra nessa ligação. A
string `"ASCW3"` no código do motor **fica como está**: é o identificador interno do `sdkjs` e o nome de
reserva que documentos salvos pelo ONLYOFFICE — **e também pelo mOffice**, que usa a mesma serialização —
carregam (a serialização de §1.2 grava `"ASCW3"` como substituto de último recurso). No entendimento do
projeto, mantê-la é compatibilidade de arquivo, não uso de marca (ponto a confirmar na resposta da Ascensio,
se vier). Achado lateral: pedir
`ASCW3` como fonte de um *run* do documento **não** chega à fonte embutida (o texto passa pelo catálogo de
fontes, onde ela não existe; os símbolos saem de outra fonte) — a embutida serve só às marcas e ao substituto
da serialização, como §1.2 já dizia.

### 1.6 Decisão de 2026-09-18 (arquivo solto)

**`fonts/ASC.ttf` não entra no pacote.** `tools/onlyoffice/obter-motor.ps1` passa a excluí-la
(`robocopy /XF ASC.ttf` na pasta `fonts`), com uma guarda que confere o **resultado** da cópia e
falha se o arquivo aparecer no destino mesmo assim. O parâmetro `-ComAscTtf` recopia a fonte —
só para repetir o experimento; o script avisa em voz alta que o pacote resultante não deve ser
distribuído.

**Nenhuma fonte substituta foi necessária**, porque nada deixou de ser desenhado: a família ASCW3
continua funcionando a partir do binário embutido nos bundles. Colocar uma substituta livre em
`third_party/onlyoffice/fonts/livres/` — o plano original — **não teria efeito nenhum**: o motor
não lê o disco para essa família (§1.3).

### 1.7 Pendência nomeada em 2026-09-18 — **resolvida pela opção 1 em 2026-09-19 (§1.5)**

> **Estado atual:** a opção **1** abaixo foi executada (§1.5): nenhum bundle contém mais a fonte da
> Ascensio. A opção **2** (e-mail) segue em paralelo, por cortesia e registro — não é mais condição para
> distribuir. O texto abaixo fica como registro do raciocínio da época.

A fonte **continuava sendo redistribuída** dentro de `editors/sdkjs/{word,cell,slide,visio}/sdk-all.js`.
Excluir o arquivo solto é necessário mas **não suficiente**. As saídas possíveis, para decisão
humana:

1. **Trocar o `base64` nos quatro bundles** por uma fonte livre (OFL/Apache) com os 12 glifos nos
   mesmos 9 códigos alcançáveis (`U+F020`–`U+F022`, `U+F038`–`U+F039`, `U+25C9`, `U+25CB`,
   `U+2610`, `U+2611`) **e com a subtabela `cmap` 1/0**, sem a qual o `String.fromCharCode(33)` do
   editor não acha glifo nenhum. Renomeada para a família `ASCW3` com `fonttools`. Resolve de vez,
   mas passa a ser um **patch de verdade** no `sdkjs` (entra em `MODIFICACOES-ONLYOFFICE.md` como
   tal, com a obrigação de fonte correspondente da AGPL), muda o desenho das marcas `→`/`☐`, e
   precisa de um A/B novo. É trabalho de uma task própria.
2. **Perguntar à Ascensio System SIA** sob que licença está o `ASC.ttf` — a alternativa que o
   usuário já havia autorizado como último recurso, e que cabe no mesmo e-mail da redação dos
   termos adicionais da AGPL (`AVISO-ONLYOFFICE.txt`).

Enquanto nenhuma das duas acontecer, o pacote **não pode ser distribuído** alegando que a
pendência do `ASC.ttf` está resolvida. O que está resolvido é: *o arquivo é dispensável e saiu do
pacote*, com a evidência acima.

### 1.8 Como reproduzir

```powershell
# a) provar que o binário embutido ERA o mesmo arquivo (no bundle original, guardado como .orig pelo patch)
python -c "import base64,hashlib; d=open('third_party/onlyoffice/editors/sdkjs/word/sdk-all.js.orig','rb').read(); i=d.find(b'AAEAAAAPAIAAAwBwRFNJRw'); j=d.find(b'\"',i); print(hashlib.sha256(base64.b64decode(d[i:j])).hexdigest())"

# b) refazer o A/B de 2026-09-18 (arquivo solto; sessão de console local, não RDP)
tools\onlyoffice\obter-motor.ps1 -ComAscTtf          # metade "com"
Remove-Item $env:ProgramData\mOffice\fontes\impressao-digital.txt
tools\onlyoffice\obter-motor.ps1                     # metade "sem"
Remove-Item $env:ProgramData\mOffice\fontes\impressao-digital.txt

# c) a fonte nossa e o patch (2026-09-19)
python tools\fontes\construir-simbolos.py            # regenera tools\fontes\mOffice-Simbolos.ttf (mesmos bytes)
tools\onlyoffice\aplicar-patches.ps1                 # troca o base64 nos 4 bundles (idempotente); -Reverter desfaz
dotnet test tests\mOffice.Nucleo.Tests --filter FonteSimbolos   # unidade + guarda + janela real (console) com controle negativo
```

---

## 2. Demais fontes do pacote

Sem decisão pendente — licença localizada e arquivo de licença em `docs/licencas/`:

| Fonte | Licença |
|---|---|
| Open Sans (`OpenSans-*.ttf`, `fonts/LICENSE.txt`) | Apache-2.0 (`LICENCA-FONTE-OPENSANS-APACHE-2.0.txt`) |
| Caladea (`fonts/caladea/`) | OFL-1.1 (`LICENCA-FONTE-CALADEA-OFL-1.1.txt`) |
| Carlito (`fonts/crosextra/`) | OFL-1.1 (`LICENCA-FONTE-CARLITO-OFL-1.1.txt`) |
| Asana Math (`fonts/asana/`) | OFL-1.1 (`LICENCA-FONTE-ASANA-MATH-OFL-1.1.txt`) |
| OpenSymbol (`fonts/openoffice/`) | OFL-1.0 (`LICENCA-FONTE-OPENSYMBOL-OFL-1.0.txt`) |
