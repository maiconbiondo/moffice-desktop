// ponte-desktop.js — implementa window.AscDesktopEditor (subconjunto) para os builds desktop do sdkjs/web-apps.
// Injetado ANTES do grampo.js (que continua embrulhando por cima e logando cada chamada).
//
// Fonte da verdade: o motor instalado (9.4.0.129). Cada valor abaixo foi lido de
//   sdkjs/word/sdk-all.js       — DesktopOfflineAppDocumentEndLoad, Z$, fe.vdc/RS, UpdateInstallPlugins
//   sdkjs/word/sdk-all-min.js   — CreateEditorApi(this), GetEncryptedHeader, GetFontThumbnailHeight
//   web-apps/apps/*/main/app.js — getFontsSprite, isSupportBinaryFontsSprite, isActive/isFeatureAvailable
// ou da sonda da Task 4 (poc/relatorios/sonda-20260916.md). Se algo aqui discordar deles, ELES vencem.
(function () {
  if (window.AscDesktopEditor && window.AscDesktopEditor.__moffice) return;

  // O objeto de estado pode JA EXISTIR quando este script roda: desde a Task 2 do Marco 1 / Plano 2 a janela
  // injeta `window.__moffice.tema` num script de DocumentCreated registrado ANTES deste arquivo. Por isso os
  // padroes sao MESCLADOS no objeto existente (e nao `window.__moffice || {…}`, que os descartaria se o
  // objeto ja estivesse la sem eles) — e sem sobrescrever o que ja veio do host.
  var estado = window.__moffice = window.__moffice || {};
  if (estado.pronto === undefined) estado.pronto = false;
  if (estado.sessao === undefined) estado.sessao = null;
  if (estado.indice === undefined) estado.indice = 0;

  // Configuracao do editor vinda do host (Plano 6, Task 6; patch 01 no index.html): mescla PROFUNDA de
  // window.__moffice.configEditor (ConfigEditor.Json() do host) no editorConfig que o motor monta. Os campos da URL (usuario,
  // idioma, tema) continuam do index.html; aqui so' entra o que o host decide (region, ajuda/feedback/macros/dicas desligados).
  // "arquivo=0" na URL (documento novo/desvinculado, sem arquivo em disco) tira o "Abrir Local do Arquivo": goback sem url
  // faz canBack falso no app.js (canBack = !isEmpty(goback.url)). Sem arquivo=, nada muda.
  estado.mesclarConfig = function (c) {
    function mesclar(destino, origem) {
      for (var k in origem) {
        if (!Object.prototype.hasOwnProperty.call(origem, k)) continue;
        var v = origem[k];
        if (v && typeof v === "object" && !Array.isArray(v) && destino[k] && typeof destino[k] === "object") mesclar(destino[k], v);
        else destino[k] = v;
      }
      return destino;
    }
    try {
      if (estado.configEditor) mesclar(c, JSON.parse(JSON.stringify(estado.configEditor)));
      if (/[?&]arquivo=0(&|$)/.test(location.search)) {
        c.customization = c.customization || {};
        c.customization.goback = { url: "" };
      }
    } catch (e) { logar("mesclarConfig: " + e); }
    return c;
  };

  function host() { return chrome.webview.hostObjects.sync.ponte; }

  // "log" (PonteLog) e' o objeto de host para logar do quadro do editor (achado 1 da revisao 2026-09-17:
  // "arquivos" nao vai mais para o quadro). Usado so pelos proprios erros internos da ponte.
  function logar(linha) {
    try { chrome.webview.hostObjects.sync.log.Log(linha); } catch (e) { console.log("[ponte] " + linha); }
  }

  // Listener PERMANENTE (preocupacao 1 do relatorio da Task 6, revisao 2026-09-17 fix review 1): foi uma
  // rejeicao de Promise NAO tratada (window.RendererProcessVariable ausente, ver abaixo) que travou o editor
  // em somente-leitura, em SILENCIO — sem excecao visivel, sem "STUB" no grampo, o texto so parava de mudar.
  // window.onerror nao pega rejeicao de Promise; sem isto, a proxima degradacao assim so aparece no log do
  // grampo/egresso se alguem instrumentar na mao de novo. Registrado uma vez por documento (janela de cima E
  // quadro do editor recebem seu proprio __moffice, ver o guard do topo do arquivo).
  window.addEventListener("unhandledrejection", function (e) {
    logar("UNHANDLED_REJECTION " + (e.reason && (e.reason.stack || e.reason.message || e.reason)));
  });
  window.addEventListener("error", function (e) {
    var texto = (e.error && (e.error.stack || e.error.message)) || e.message;
    logar("ERROR " + texto);
    // Contador legivel pelos testes de janela real (window.frames[0].__moffice.errosJs): um erro de JS no
    // quadro do editor e' exatamente o que o trava em "Ocorreu um erro" sem o C# ficar sabendo.
    (window.__moffice.errosJs = window.__moffice.errosJs || []).push(String(texto).slice(0, 300));
  });

  // Uma capacidade que o host nao tem foi pedida: registra UMA vez por nome (a roda do mouse, por exemplo,
  // chamaria dezenas de vezes por segundo) e segue sem lancar.
  var jaAvisados = {};
  // Fila do corretor (ver SpellCheck): cada pedido so' comeca depois do anterior responder.
  var filaOrtografia = Promise.resolve();
  function semSuporte(nome, detalhe) {
    if (jaAvisados[nome]) return;
    jaAvisados[nome] = true;
    logar("SEM SUPORTE " + nome + (detalhe === undefined ? "" : " " + detalhe));
  }

  // A sessao e' lida sob demanda: o objeto de host do quadro so existe depois do FrameCreated, e este script
  // roda no DocumentCreated (que pode vir antes em documentos sem sessao, como a janela vazia).
  function sessao() {
    if (!estado.sessao) {
      try { estado.sessao = JSON.parse(host().SessaoJson()); } catch (e) { return null; }
    }
    return estado.sessao;
  }

  function avisar(tipo, param) {
    try { host().Avisar(tipo, param || ""); }
    catch (e) { logar("falhou Avisar " + tipo + ": " + e); }
  }

  // window.RendererProcessVariable: variavel GLOBAL que o renderer do app oficial (CEF) injeta na pagina
  // antes de qualquer script dela. web-apps/apps/*/main/index.html le' `window.RendererProcessVariable.theme`
  // para montar `window.uitheme`, e Common.Controllers.Desktop guarda o objeto inteiro em `r`, usado SEM
  // guarda em `systemThemeSupported()` (`return r.theme && "disabled" !== r.theme.system`).
  //
  // Sem ela (MEDIDO na Task 6, 2026-09-17): `systemThemeSupported` estoura
  // "Cannot read properties of undefined (reading 'theme')" DENTRO de uma Promise; ninguem trata a rejeicao;
  // o proprio sdkjs escuta `onunhandledrejection` e reage com `tAc()` -> `asc_setViewMode(true)` + lock
  // global do CollaborativeEditing. Resultado: o documento abre, o texto aparece, NADA falha visivelmente —
  // e o editor fica em SOMENTE LEITURA, com toda edicao pela API (asc_AddText, asc_PasteData...) descartada
  // em silencio. Foi assim que a Task 6 comecou vermelha. A pilha que provou isso:
  //   Object.systemThemeSupported -> Object.map -> n.a -> (Promise) -> onunhandledrejection -> tAc -> eRb(true)
  //
  // Campos REALMENTE lidos pelo web-apps (medido nesta janela): theme{id,type,system}, localthemes (mapa
  // id -> {colors}), rtl — so' esses tres. os/defaultPrinterName/keyboard.langs abaixo existem so' porque o
  // objeto real do renderer os define (completude com a forma oficial), mas nenhum uso deles foi observado
  // nesta versao; helpUrl NAO entra aqui por isso mesmo (nunca foi visto sendo lido, e o objeto tambem nao o
  // define). theme.system: "disabled" (nao "light"/"dark") porque o PoC nao acompanha tema do SO nenhum —
  // "light" deixaria o item "usar tema do sistema" vivo na UI sem que o host jamais o atualize.
  //
  // Task 2 do Marco 1 / Plano 2: o TEMA vem do host. A janela injeta `window.__moffice.tema = {id, type}` num
  // script de DocumentCreated registrado ANTES deste arquivo (JanelaEditor.InicializarAsync), entao aqui ele
  // ja existe. Sem ele (pagina de teste, host antigo) fica o claro de sempre. `system` continua "disabled"
  // nos dois casos: o mOffice nao acompanha o tema do SO — quem escolhe e' a configuracao do produto.
  var temaDoHost = (estado && estado.tema) || null;
  window.RendererProcessVariable = window.RendererProcessVariable || {
    theme: temaDoHost
      ? { id: temaDoHost.id, type: temaDoHost.type, system: "disabled" }
      : { id: "theme-classic-light", type: "light", system: "disabled" },
    localthemes: {},
    rtl: false,
    os: "win",
    defaultPrinterName: "",
    keyboard: { langs: [] }
  };

  // --- midia do documento (Task 5/S4) -----------------------------------------------------------------
  // DesktopOfflineAppDocumentEndLoad(pasta, ...) faz `AscCommon.Ss.tia = pasta` e, como a pasta nao comeca
  // com "file:", prefixa "file://" (sdkjs/word/sdk-all.js). Cada imagem sai de `tia + "/media/" + nome`
  // (fe.RS/fe.vdc). Num app CEF a pagina do editor TAMBEM e' file://, entao funciona; aqui a pagina e'
  // https://moffice.local e o Chromium RECUSA subrecurso file:// (medido na Task 5A: fetch -> TypeError,
  // <img> -> onerror, sem chegar ao disco). A pasta de recovery ja' e' servida por https://moffice-doc.local,
  // e a condicao do EndLoad so' aceita argumento comecando com "file:" — entao a base e' trocada AQUI, logo
  // depois do EndLoad, direto na propriedade. Nenhum arquivo do motor e' alterado.
  //
  // A propriedade NAO e' procurada pelo nome: cada editor tem o seu bundle minificado em separado e o nome e'
  // DIFERENTE nos tres (MEDIDO em third_party/onlyoffice/editors/sdkjs/*/sdk-all.js, 9.4.0.129):
  //   word  -> AscCommon.Ss.tia      slide -> AscCommon.ru.Eda      cell -> AscCommon.Cw.$pa
  // (foi assim que a primeira versao desta troca, que procurava "AscCommon.Ss.tia", funcionou no DOCX e nao
  // fez NADA no PPTX — o log dizia "sem AscCommon.Ss" e as figuras do slide continuavam invisiveis.) Entao a
  // busca e' pelo VALOR: o EndLoad acabou de gravar exatamente `file://` + a pasta de recovery, string que o
  // host conhece e passa aqui. Vale para os tres editores e sobrevive a uma renomeacao do minificador.
  estado.trocarBaseMidia = function (base, antigo) {
    if (!base) return "sem troca (EndLoad de erro)";
    var ac = window.AscCommon;
    if (!ac) return "sem AscCommon";
    var trocados = [];
    for (var k in ac) {
      var o;
      try { o = ac[k]; } catch (e) { continue; }
      if (!o || typeof o !== "object") continue;
      for (var p in o) {
        var v;
        try { v = o[p]; } catch (e) { continue; }
        if (v === antigo) { try { o[p] = base; trocados.push(k + "." + p); } catch (e) { } }
      }
    }
    return trocados.length ? "trocado " + trocados.join(",") + " -> " + base : "nao achei a base " + antigo;
  };

  // OpenFilenameDialog(filtro, multi, callback): o app oficial guarda o callback em
  // window.on_native_open_filename_dialog, manda a mensagem para o processo do navegador e o chama de volta
  // com o CAMINHO escolhido (desktop-sdk, client_renderer_wrapper.cpp:2567 e :6840). Mesmo desenho aqui: o
  // host abre o dialogo fora desta chamada sincrona e responde por `responderDialogo`. Um callback pendente
  // por vez, como no oficial (`if (window.on_native_open_filename_dialog) return;`).
  estado.responderDialogo = function (caminho) {
    var cb = estado.dialogoPendente;
    estado.dialogoPendente = null;
    if (!cb) return "sem callback pendente";
    try { cb(caminho || ""); } catch (e) { logar("callback do dialogo de arquivo: " + e); return "erro no callback"; }
    return caminho ? "callback chamado" : "callback chamado (cancelado)";
  };

  // Comparar/combinar (Plano 6, Task 8): o resultado chega por window.onDocumentCompare/onDocumentMerge (o host chama no quadro).
  function pedirComparacao(caminho, uso) {
    var id = "";
    try { id = host().PedirConversao(String(caminho || ""), 8192, uso); } catch (e) { logar(uso + ": " + e); }
    if (!id) logar(uso + ": caminho sem consentimento (nao veio do dialogo)");
  }
  function base64ParaBuffer(b64) {
    var s = atob(b64), n = s.length, u = new Uint8Array(n);
    for (var i = 0; i < n; i++) u[i] = s.charCodeAt(i);
    return u.buffer;
  }
  // Entrega da conversao pedida por convertFile: json = {"ok":true,"base64","tamanho","pasta","imagens"} ou {"ok":false}.
  estado.entregarConversao = function (id, json) {
    var cb = estado.conversoes && estado.conversoes[id];
    if (!cb) return "sem pedido";
    delete estado.conversoes[id];
    var r;
    try { r = JSON.parse(json); } catch (e) { r = { ok: false }; }
    var obj = r.ok ? {
      get: function () { return base64ParaBuffer(r.base64); },
      getImages: function () { return r.imagens || {}; },
      close: function () { try { host().LiberarConversao(id); } catch (e) { } }
    } : null;
    try { cb(obj); } catch (e) { logar("entregarConversao cb: " + e); }
    return "ok";
  };
  // Entrega do "Abrir origem" (openExternalReference): 1 = abriu, 0 = nao (o sdkjs mostra o aviso dele).
  estado.entregarReferencia = function (id, codigo) {
    var cb = estado.referencias && estado.referencias[id];
    if (!cb) return "sem pedido";
    delete estado.referencias[id];
    if (typeof cb === "function") { try { cb(codigo | 0); } catch (e) { logar("entregarReferencia cb: " + e); } }
    return "ok";
  };
  // Entrega da leitura pedida por loadLocalFile: base64 dos bytes, ou "" (falhou) -> cb(null).
  estado.entregarLeitura = function (id, b64) {
    var cb = estado.leituras && estado.leituras[id];
    if (!cb) return "sem pedido";
    delete estado.leituras[id];
    try { cb(b64 ? new Uint8Array(base64ParaBuffer(b64)) : null); } catch (e) { logar("entregarLeitura cb: " + e); }
    return "ok";
  };

  // O editor vive num iframe; o teste (e quem olha de fora) le o documento de cima. Marca nos dois.
  function marcarPronto() {
    if (estado.pronto) return;   // asc_onDocumentContentReady e AscDesktopEditor.onDocumentContentReady chamam os dois
    estado.pronto = true;
    try { if (window.top !== window && window.top.__moffice) window.top.__moffice.pronto = true; } catch (e) { }
    // Task 7: e' este instante que a medida "abertura do editor" cronometra no lado C# (o host nao tem como
    // saber sozinho quando o documento ficou pronto — quem sabe e' o editor).
    avisar("pronto", "");
  }

  function fontesSprite(escala) {
    var s = sessao();
    if (!s || !s.fontes) return "";
    // O app oficial gera os sprites junto com o AllFonts.js; eles NAO estao no motor copiado. Servidos pelo
    // esquema nativo (ServidorAscDesktop). Sufixo "_ea" (zh/ja/ko) nao tratado: o PoC e' pt-BR.
    return "ascdesktop://fonts/" + s.fontes + "/fonts_thumbnail" + (escala || "") + ".png";
  }

  var ponte = {
    __moffice: true,

    // Macros JavaScript do ONLYOFFICE: DESLIGADAS por decisao do dono (Plano 6). Defesa em profundidade junto de
    // customization.macros:false da config injetada; o VBA dos arquivos so' e' preservado, nunca executado.
    isSupportMacroses: function () { return false; },

    // --- ciclo de vida do editor
    CreateEditorApi: function (api) {
      estado.api = api;
      // "Salvar como" no menu Arquivo, igual ao app oficial (Task 2b do Plano 4). O web-apps 9.4 so' troca
      // "Baixar em" (painel de download da versao web) por "Salvar como" (acao "save-desktop" ->
      // api.asc_DownloadAs() -> LocalFileSave("saveas=true;", ..., fileType 0) -> dialogo do host) quando
      // `isDesktopApp && isOffline`, e `isOffline = api.asc_isOffline()` e' lido UMA vez no onEditorPermissions
      // dos tres editores. O asc_isOffline do sdkjs testa `location.protocol` comecando com "file": o oficial abre
      // de file://, o mOffice de https://moffice.local — dava false. O sdk-all.js desktop troca o prototipo por
      // `return true` depois, tarde demais (MEDIDO sem esta linha: asc_isOffline() ja' da' true com o documento
      // aberto, mas appOptions.isOffline ficou false e o menu mostrava "Baixar em"). A sobreposicao e' NA
      // INSTANCIA e so' do nome publico: o sdkjs interno chama a funcao minificada direto (ex. cell
      // `Y.asc_isOffline=Y.d1e`), entao so' o web-apps enxerga o true. CreateEditorApi e' chamado no construtor
      // do api, antes das permissoes. Efeitos colaterais (os do app oficial, lidos no bundle): sem chat,
      // historico, favorito e forcesave; menu Arquivo abre em "Informacoes"; sem botao "Baixar" no cabecalho;
      // o plugins.json deixa de ser pedido. O Ctrl+S continua salvando no lugar (canSaveToFile -> asc_Save) e
      // documento novo continua perguntando o destino ao host. Teste: SalvarComoDesktopTests.
      if (api) api.asc_isOffline = function () { return true; };
      if (api && api.asc_registerCallback) api.asc_registerCallback("asc_onDocumentContentReady", marcarPronto);
    },
    InitJSContext: function () { },
    // Notificacao pura (sdk-all-min.js: `if (AscDesktopEditor.onDocumentContentReady) ...()`), sem retorno lido.
    onDocumentContentReady: function () { marcarPronto(); },
    // app.js chama depois do doc:onready para a lista de recentes; o retorno nao e' lido. Plano 6, Task 8: o host responde FORA desta
    // chamada com window.onupdaterecents(<lista da casca>) neste quadro — o "Abrir recente" do menu Arquivo aparece preenchido.
    LocalFileRecents: function () { avisar("recentes", ""); },
    ConsoleLog: function (s) { console.log(s); },
    SetDocumentName: function (nome) { try { document.title = nome; } catch (e) { } },

    // --- valores que o bootstrap le' SEM guarda de tipo (ver sonda: um neutro errado derruba o carregamento)
    GetSupportedScaleValues: function () { return [1, 1.25, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3, 3.5, 4, 4.5, 5]; },
    GetFontThumbnailHeight: function () { return 28; },      // padrao do proprio sdkjs quando nao ha nativo
    GetEncryptedHeader: function () { return "ENCRYPTED;"; }, // idem (e' lido .length)
    getViewportSettings: function () { return {}; },          // app.js le .widgetType
    // UpdateSystemPlugins/UpdateInstallPlugins exigem JSON com EXATAMENTE dois elementos (sistema e usuario)
    GetInstallPlugins: function () { return '[{"url":"","pluginsData":[]},{"url":"","pluginsData":[]}]'; },
    getFontsSprite: fontesSprite,
    isSupportBinaryFontsSprite: false, // propriedade, nao funcao: com false o app usa o .png (e nao o .png.bin)

    // --- capacidades que este PoC nao tem
    isSupportNetworkFunctionality: function () { return false; },
    // Task 7 do Plano 3: "Proteger com senha" (Arquivo -> Proteger) LIGADO, como no app oficial (desktop-sdk
    // client_renderer_wrapper.cpp: m_bIsSupportProtect = flag "onlypass" & 2, padrao 3 -> true). O web-apps so'
    // mostra o painel com `isPasswordSupport = isEdit && api.asc_isProtectionSupport()`. Com ele, trocar/remover a
    // senha chega ao host pelo 2o argumento de LocalFileSave (a senha atual do editor, ver abaixo).
    IsProtectionSupport: function () { return true; },
    // PROPRIEDADE (nao funcao), como no oficial (Object.defineProperty(..., 'CryptoMode') -> m_nCryptoMode, padrao
    // 0). Obrigatoria para a senha (Task 7): window.asc_initAdvancedOptions (sdkjs/*/sdk-all-min.js) faz
    // `if (AscDesktopEditor && 0 !== AscDesktopEditor.CryptoMode && !api.<carregado>) { enfileira; return }` — com
    // `undefined` o pedido de senha ficava ENFILEIRADO para o modo de criptografia de plugin e o dialogo nunca
    // aparecia. 0 = sem criptografia ponta a ponta.
    CryptoMode: 0,
    IsSupportMedia: function () { return false; },
    isBlockchainSupport: function () { return false; },
    IsSupportNativePrint: function () { return false; },

    // --- chamadas SEM teste de existencia no sdkjs/web-apps 9.4.0.129 (achado de 2026-09-25: a roda do
    // mouse chamava CheckNeedWheel, que nao existia — TypeError, e o editor travava o documento com "Ocorreu
    // um erro. Use a opcao 'Baixar como'…"). Ate' o Plano 2 o grampo.js (sempre ligado) respondia qualquer
    // nome faltando com um stub e escondia estas lacunas; ele virou opcional e elas apareceram. A guarda
    // PonteCompletaTests varre os bundles e falha se surgir outra chamada sem implementacao aqui.
    // Retorno != false = o PROPRIO editor rola com a roda (o que queremos: nao ha "rolagem nativa" no host).
    CheckNeedWheel: function () { return true; },
    // Apresentacao de slides em tela cheia (Task 7 do Plano 3). MEDIDO em sdkjs/slide/sdk-all.js: o editor chama
    // SetFullscreen(true) ao iniciar a demonstracao e SetFullscreen(false) ao sair (Esc, fim, fechar) — sem
    // retorno lido. O oficial manda "onfullscreenenter"/"onfullscreenleave" ao processo do navegador; aqui o host
    // (JanelaEditor) deixa a janela sem borda, do tamanho do monitor, e restaura depois.
    SetFullscreen: function (b) {
      estado.telaCheia = !!b;
      try { if (window.top !== window && window.top.__moffice) window.top.__moffice.telaCheia = !!b; } catch (e) { }
      avisar("telaCheia", b ? "1" : "0");
    },
    // "Modo apresentador" (janela do apresentador) — Task 7 do Plano 3. MEDIDO em sdkjs/slide/sdk-all-min.js: o FIM de
    // TODA demonstracao passa por `DAf`, que com AscDesktopEditor presente chama `AscDesktopEditor.endReporter()` SEM
    // testar se existe (o PonteCompletaTests nao varre os sdk-all-min.js). Sem ela o Esc lancava TypeError dentro do
    // encerramento, o SetFullscreen(false) nunca vinha e a janela ficava presa em tela cheia (achado na janela real
    // desta task). O apresentador em janela separada nao existe neste host: encerrar e as mensagens entre as janelas
    // sao no-op (nao ha' outra janela).
    // "Mostrar modo de apresentador" (fix round 1 da Task 7): MEDIDO em sdkjs/slide/sdk-all-min.js (`kTc`): com
    // apresentador pedido o editor chama startReporter e depois `X9h`, que deixa a tela preta com "Loading" ESPERANDO a
    // janela do apresentador mandar {reporter_command:"start_show"} — que aqui nunca viria (e o Esc nao sai dessa espera:
    // o teclado da demonstracao so' age depois de ela comecar). Reserva: logo depois (setTimeout 0, quando o X9h ja'
    // rodou), a propria ponte entrega esse "start_show" pela API publica `DemonstrationReporterMessages` — a
    // apresentacao normal comeca na janela, em tela cheia, e sai com Esc como qualquer outra.
    // Em modo apresentador a janela principal tambem NAO trata teclado/clique/roda: manda tudo ao apresentador
    // (sdkjs/slide `P1`/`sj`/`IVg`: `if (YW) a_a({main_command:true, keyCode|mouseUp|mouseWhell|resize})` ->
    // sendToReporter), que decide e devolve reporter_command. Por isso, enquanto a reserva vale, sendToReporter faz o
    // papel do apresentador: traduz o main_command no reporter_command equivalente (as mesmas teclas de `GVg`) e o
    // devolve pela mesma API publica. Sem isto o Esc nao encerrava (medido: so' a rede de seguranca do host tirava a
    // janela da tela cheia, com a apresentacao ainda rodando por baixo).
    startReporter: function () {
      semSuporte("startReporter", "(modo apresentador: sem janela propria, vira apresentacao normal)");
      estado.apresentadorEmulado = true;
      setTimeout(function () { comandoDoApresentador({ reporter_command: "start_show" }); }, 0);
    },
    endReporter: function () { estado.apresentadorEmulado = false; },
    sendToReporter: function (json) {
      if (!estado.apresentadorEmulado) return;
      var m;
      try { m = JSON.parse(json); } catch (e) { return; }
      if (!m || !m.main_command) return;
      var c = null;
      if (m.keyCode !== undefined) {
        var k = m.keyCode;
        if (k === 27) c = { reporter_command: "end" };
        else if (k === 13 || k === 32 || k === 34 || k === 39 || k === 40) c = { reporter_command: "next" };
        else if (k === 33 || k === 37 || k === 38) c = { reporter_command: "prev" };
        else if (k === 36) c = { reporter_command: "go_to_slide", slide: 0 };
        else if (k === 35) {
          var api = apiDoEditor();
          var n = api && typeof api.getCountPages === "function" ? api.getCountPages() : 0;
          if (n > 0) c = { reporter_command: "go_to_slide", slide: n - 1 };
        }
      }
      else if (m.mouseUp && m.isMainMouseDown) c = { reporter_command: "next" };
      else if (m.mouseWhell !== undefined) c = { reporter_command: m.mouseWhell > 0 ? "next" : "prev" };
      else if (m.resize) c = { reporter_command: "resize" };
      // Assincrono, como a resposta de uma janela de verdade: nunca reentra no tratador que acabou de mandar.
      if (c) setTimeout(function () { comandoDoApresentador(c); }, 0);
    },
    sendFromReporter: function () { },
    // Opcoes de abertura (Task 7 do Plano 3): a resposta do dialogo de senha que o editor mostra depois de o host
    // chamar window.asc_initAdvancedOptions(90|91). MEDIDO nos tres bundles: "<m_sPassword>" + escapeXml(senha) +
    // "</m_sPassword>" (CSV/TXT mandaria <m_nCsvTxtEncoding>..., fora deste plano). O conteudo leva a SENHA: nunca
    // vai para log (o host omite o parametro deste aviso; o grampo tambem).
    SetAdvancedOptions: function (x) { avisar("opcoesAvancadas", String(x || "")); },
    // Conversao de arquivo para o binario do editor (Plano 6, Task 8). Contrato MEDIDO em sdkjs/word/sdk-all.js 9.4: o chamador
    // espera cb(obj) com obj.get() -> ArrayBuffer, obj.getImages() -> mapa de midias, obj.close(). O host converte com o x2t fora
    // da chamada sincrona e devolve por __moffice.entregarConversao. Sem consentimento (o caminho nao veio do dialogo): cb(null).
    // Chamadores: Inserir > Texto do arquivo (word, 8193), vinculo externo e grafico ligado (8194), Dados > XML (cell, 8194).
    convertFile: function (caminho, formato, cb) {
      var id = "";
      try { id = host().PedirConversao(String(caminho || ""), formato | 0, "converter"); } catch (e) { logar("convertFile: " + e); }
      if (!id) { if (typeof cb === "function") setTimeout(function () { try { cb(null); } catch (e) { logar("convertFile cb: " + e); } }, 0); return; }
      (estado.conversoes = estado.conversoes || {})[id] = cb;
    },
    // Dicionarios de hifenizacao (Plano 6, Task 8): o sdkjs pede <base><lang>/hyph_<lang>.dic. Sem isto a base era
    // "../../../../dictionaries/" a partir do editor = https://moffice.local/dictionaries/, que nao existe (os dicionarios ficam em
    // motor/dictionaries, ao lado de editors). O host mapeia moffice-dic.local para essa pasta (so' leitura, local).
    getDictionariesPath: function () { return "https://moffice-dic.local/"; },
    // Chamadas COM DIGITO no nome que a guarda PonteCompletaTests nao enxergava antes da Task 8 : neutras. GetImageBase64 e' a imagem da
    // assinatura OOXML (Plano 7) e a imagem de criptografia; LoadFontBase64 = fonte do sistema para o editor; localSaveToDrawingFormat2 = salvar
    // como imagem do nativo; Crypto_GetLocalImageBase64 = modo de criptografia (CryptoMode = 0: nunca ligado).
    GetImageBase64: function () { semSuporte("GetImageBase64", "(assinatura OOXML: Plano 7)"); return ""; },
    LoadFontBase64: function () { semSuporte("LoadFontBase64"); return ""; },
    localSaveToDrawingFormat2: function () { semSuporte("localSaveToDrawingFormat2"); },
    Crypto_GetLocalImageBase64: function () { semSuporte("Crypto_GetLocalImageBase64"); return ""; },
    // Criptografia, certificados, midia e visualizador nativo: capacidades que este host declara NAO ter
    // (IsProtectionSupport/IsSupportMedia/... = false). Se o editor chamar mesmo assim, registra e segue.
    NativeViewerOpen: function () { semSuporte("NativeViewerOpen"); },
    PreloadCryptoImage: function () { semSuporte("PreloadCryptoImage"); },
    buildCryptedStart: function () { semSuporte("buildCryptedStart"); },
    buildCryptedEnd: function () { semSuporte("buildCryptedEnd"); },
    ViewCertificate: function () { semSuporte("ViewCertificate"); },
    AddAudio: function () { semSuporte("AddAudio"); },
    AddVideo: function () { semSuporte("AddVideo"); },
    // Baixar da internet (imagem por URL, comparar/combinar por URL): o mOffice nao tem rede. O chamador ja' abriu uma
    // "acao longa" e so' a fecha dentro do callback — sem resposta a mascara de carregamento ficava para sempre
    // (revisao final do Plano 4). O callback recebe o mapa url -> arquivo baixado: {} = nada baixado (o sdkjs fecha a
    // acao; no sdk-all.js desktop ainda mostra o aviso NAO critico dele). A casca mostra a faixa.
    DownloadFiles: function (urls, destinos, cb) {
      semSuporte("DownloadFiles");
      avisar("semSuporte", "baixarArquivos");
      if (typeof cb === "function") setTimeout(function () { try { cb({}); } catch (e) { logar("DownloadFiles callback: " + e); } }, 0);
    },
    SaveQuestion: function () { semSuporte("SaveQuestion"); },
    startExternalConvertation: function () { semSuporte("startExternalConvertation"); },

    // --- chamadas SEM teste de existencia nos sdk-all-min.js (revisao final do Plano 4). Com isOffline = true (Task 2b)
    // o web-apps passou a mostrar caminhos do app oficial que chegam aqui — ex. planilha: "Mover ou copiar planilha" ->
    // "(Criar nova planilha)" chama OpenWorkbook, e "Vinculos externos" -> "Abrir origem" chama openExternalReference.
    // Sem a funcao: TypeError -> window.onerror do sdkjs -> asc_onError CRITICO e o documento travado (mesma classe do
    // CheckNeedWheel). O PonteCompletaTests agora varre tambem os sdk-all-min.js. Retornos: so' onde o chamador LE
    // (sites MEDIDOS no 9.4.0.129 — ver o relatorio da Task 2b, secao "Final review fix").
    // Planilha: copia/move de planilha para uma pasta de trabalho NOVA (o oficial abre outra janela com o binario).
    // Sem janela nova aqui: avisa a casca (faixa) e nada mais — o sdkjs nao espera resposta (e' a ultima instrucao).
    OpenWorkbook: function () {
      semSuporte("OpenWorkbook", "(mover/copiar planilha para pasta de trabalho nova)");
      avisar("semSuporte", "novaPastaDeTrabalho");
    },
    // Vinculo externo -> "Abrir origem" (Plano 6, Task 8). MEDIDO em cell/sdk-all-min.js 9.4: `ref` e' o caminho do vinculo sem "file:///" e o
    // callback trata 0 como "nao abriu" (aviso NAO critico do sdkjs) e qualquer outro valor como sucesso. O host pergunta ao usuario (o caminho
    // vem do documento; fix 1 da revisao), confere a existencia so' depois do sim, pede a abertura em janela nova e responde 1/0 por
    // __moffice.entregarReferencia. Recusado na hora (dispositivo, relativo, nao-planilha, ja' negado): cb(0), como antes.
    openExternalReference: function (ref, cb) {
      var id = "";
      try { id = host().AbrirReferenciaExterna(typeof ref === "string" ? ref : String((ref && (ref.path || ref.Path)) || "")); }
      catch (e) { logar("openExternalReference: " + e); }
      if (!id) { if (typeof cb === "function") setTimeout(function () { try { cb(0); } catch (e) { logar("openExternalReference callback: " + e); } }, 0); return; }
      (estado.referencias = estado.referencias || {})[id] = cb;
    },
    // Importar CSV/TXT de arquivo (planilha; Plano 6, Task 8): cb(Uint8Array) com os bytes do arquivo escolhido no dialogo "csv/txt" (o sdkjs
    // detecta a codificacao pelo BOM: `AscCommon.Vbf`). Caminho sem consentimento do host: cb(null) — o chamador trata o nulo (`ta && G(ta)`).
    loadLocalFile: function (caminho, cb) {
      var id = "";
      try { id = host().PedirLeitura(String(caminho || "")); } catch (e) { logar("loadLocalFile: " + e); }
      if (!id) { if (typeof cb === "function") setTimeout(function () { try { cb(null); } catch (e) { logar("loadLocalFile cb: " + e); } }, 0); return; }
      (estado.leituras = estado.leituras || {})[id] = cb;
    },
    // Salvar como imagem (formatos de desenho): o editor fica em "acao longa" ate' o nativo chamar
    // api.onLocalSaveToDrawingFormat(codigo); != 0 -> o aviso NAO critico do proprio sdkjs.
    localSaveToDrawingFormat: function () {
      semSuporte("localSaveToDrawingFormat");
      setTimeout(function () {
        var api = apiDoEditor();
        try { if (api && typeof api.onLocalSaveToDrawingFormat === "function") api.onLocalSaveToDrawingFormat(1); }
        catch (e) { logar("onLocalSaveToDrawingFormat: " + e); }
      }, 0);
    },
    emulateCloudPrinting: function () { },
    // So' e' consultado depois de SetPdfCloudPrintFileInfo existir (nao existe): "ja' em cache" = nada a mandar.
    IsCachedPdfCloudPrintFileInfo: function () { return true; },
    // Comparar/combinar com documento de arquivo (texto; Plano 6, Task 8). MEDIDO em word/sdk-all-min.js 9.4: o editor chama
    // CompareDocumentFile(caminho, opcoes) depois do dialogo "word", e o nativo responde window.onDocumentCompare(pasta, base64, tamanho,
    // imagens) / window.onDocumentMerge(...) (base64 "" = erro: o proprio editor mostra o aviso). O host converte e chama de volta neste quadro.
    CompareDocumentFile: function (caminho) { pedirComparacao(caminho, "comparar"); },
    CompareDocumentUrl: function () { semSuporte("CompareDocumentUrl"); },
    MergeDocumentFile: function (caminho) { pedirComparacao(caminho, "combinar"); },
    MergeDocumentUrl: function () { semSuporte("MergeDocumentUrl"); },
    // Imagens de arquivos baixados (DownloadFiles, que ja' e' sem suporte): o formato vai para uma lista.
    // RemoveFile NUNCA apaga nada aqui (o caminho viria da pagina).
    GetImageFormat: function () { semSuporte("GetImageFormat"); return ""; },
    RemoveFile: function () { semSuporte("RemoveFile"); },
    // Carregamento de script pelo nativo: 0 = "nao carreguei" -> o sdkjs carrega pelo caminho normal (XHR).
    LoadJS: function () { return 0; },
    // Assinatura digital do oficial (IsSignaturesSupport ausente = sem suporte; os pontos de entrada chegam mesmo assim).
    Sign: function () { semSuporte("Sign"); },
    SelectCertificate: function () { semSuporte("SelectCertificate"); },
    GetDefaultCertificate: function () { semSuporte("GetDefaultCertificate"); },
    RemoveSignature: function () { semSuporte("RemoveSignature"); },
    RemoveAllSignatures: function () { semSuporte("RemoveAllSignatures"); },
    // Criptografia ponta a ponta (CryptoMode = 0: nunca ligada) e nuvem/compartilhamento: nao alcancaveis aqui.
    CryptoDownloadAs: function () { semSuporte("CryptoDownloadAs"); },
    OpenFileCrypt: function () { semSuporte("OpenFileCrypt"); },
    ResaveFile: function () { semSuporte("ResaveFile"); },
    sendSystemMessage: function () { semSuporte("sendSystemMessage"); },
    CallInAllWindows: function () { semSuporte("CallInAllWindows"); },
    // Plugins: so' chamados com location.protocol "file" (o mOffice e' https). "Nao instalou/removeu".
    PluginInstall: function () { semSuporte("PluginInstall"); return false; },
    PluginUninstall: function () { semSuporte("PluginUninstall"); return false; },

    // --- identidade e estado local (sincronos, respondidos do estado que o C# mandou)
    // CheckUserId: o DesktopOfflineAppDocumentEndLoad grava o retorno em AscCommon.Hj.FJ (sdkjs/word/sdk-all.js),
    // o prefixo dos ids de elementos criados na sessao. Tem de AUMENTAR a cada reabertura da mesma pasta de
    // recuperacao (1 na primeira, 2 na primeira recuperacao...): os changes reaplicados vieram da sessao que
    // caiu, e reusar o id dela faria elementos novos colidirem com ids que ja' existem (Task 2 do Plano 3).
    CheckUserId: function () { var s = sessao(); return String((s && s.idUsuario) || 1); },
    IsLocalFile: function () { return true; },
    IsLocalFileExist: function () { return false; },
    LocalFileGetSourcePath: function () { var s = sessao(); return s ? s.origem : ""; },
    LocalFileSetSourcePath: function (p) { var s = sessao(); if (s) s.origem = p; },
    // Recuperacao (Task 2 do Plano 3): o "modificado" do editor no desktop (AscCommon.W0b.prototype.l5, sdk-all.js)
    // e' `LocalFileGetOpenChangesCount() != 0 || !LocalFileGetSaved()` enquanto o historico desta sessao esta'
    // vazio. Na reabertura de uma pasta recuperada o host chama LocalFileSetSaved(false) e
    // LocalFileSetOpenChangesCount(n) logo depois de DesktopOfflineAppDocumentApplyChanges, antes do EndLoad —
    // e o documento aparece modificado mesmo sem nenhuma edicao nova. Fora disso, os padroes de sempre (salvo, 0).
    // Plano 6, Task 5: restricao local do editor (Asc.c_oAscLocalRestrictionType). O host aplica ReadOnly com asc_setLocalRestrictions(t, true)
    // (o `true` nao chama de volta); quem chama aqui e' o web-apps, que ao ver ReadOnly marca o cabecalho "Somente leitura" e volta a
    // restricao para None (o documento fica editavel; salvar vira "Salvar como" pelo host). So' guarda e avisa (o host registra no log).
    SetLocalRestrictions: function (t) { estado.restricao = t; avisar("restricao", String(t)); },
    LocalFileGetSaved: function () { return estado.salvo !== false; },
    LocalFileSetSaved: function (b) { estado.salvo = !!b; },
    LocalFileRecoverFolder: function (pasta) { estado.recovery = pasta; },
    LocalFileGetOpenChangesCount: function () { return estado.changesAbertos || 0; },
    LocalFileSetOpenChangesCount: function (n) { estado.changesAbertos = n | 0; },
    LocalFileGetCurrentChangesIndex: function () { return estado.indice || 0; },
    LocalFileGetRelativePath: function (p) { return p; },
    // LocalFileGetImageUrl(dados) -> NOME do arquivo em media/, ou "error" (contrato do desktop-sdk oficial,
    // client_renderer_wrapper.cpp:5269 GetLocalImageUrl). `dados` e' uma data: URL (colar), um caminho local
    // (inserir do arquivo / arrastar) ou um nome de midia que ja esta no documento ("image…"/"display…").
    // Quem monta a URL final e' o sdkjs: "<AscCommon.Ss.tia>/media/" + isto (fe.RS/fe.vdc) — devolver URL
    // completa aqui geraria ".../media/https://...". O erro tem de ser "error" (e NAO o proprio argumento,
    // como estava antes): devolver o argumento faria o sdkjs montar uma URL com o caminho local/data: dentro,
    // escondendo a falha atras de uma figura quebrada sem nenhum sinal.
    LocalFileGetImageUrl: function (dados) {
      try { return host().ImagemUrl(dados || ""); }
      catch (e) { logar("ImagemUrl: " + e); return "error"; }
    },
    LocalFileGetImageUrlCorrect: function (url) { return url; },
    // so usado com "binary_content://" (aqui o bin vai em base64) e, desde o Plano 6 (Task 3), pelo dialogo de importacao de CSV/TXT:
    // asc_initAdvancedOptions(89, undefined, docinfo, 1) le' por aqui o COMECO do arquivo (previa e deteccao do delimitador). O host o
    // entrega em __moffice.dadosAbertos (Uint8Array) SO durante essa chamada; fora dela segue null.
    GetOpenedFile: function () { return (window.__moffice && window.__moffice.dadosAbertos) || null; },
    GetDropFiles: function () { return []; },
    IsImageFile: function () { return false; },

    // --- abertura: o editor avisa que esta pronto; o C# responde com DesktopOfflineAppDocumentEndLoad(...)
    LocalStartOpen: function () { avisar("localStartOpen", ""); },
    LocalFileOpen: function (p) { avisar("localFileOpen", p || ""); },

    // --- edicao: cada lote de changes vai para changes0.json (Task 6 exercita)
    // O indice tem de seguir EXATAMENTE a mesma regra do C# (PonteDesktop.SalvarChanges), porque
    // LocalFileGetCurrentChangesIndex responde daqui e o arquivo de changes e' escrito la'. Antes era
    // `estado.indice = deleteIndex + count` sempre, e no caso `deleteIndex < 0` (o "sem delecao, so' apenda"
    // do desktop-sdk) o JS dizia `-1 + count` enquanto o C# somava `count` ao indice corrente: o editor lia
    // um indice menor que o numero de changes realmente gravados e o proximo lote podia pedir uma delecao a
    // partir do lugar errado. Tres casos, iguais aos do C#:
    LocalFileSaveChanges: function (param, deleteIndex, count) {
      // Congelada quando a reaplicacao de uma sessao recuperada falhou (o host liga no catch do script do
      // EndLoad): o changes0.json tem de ficar como estava — a janela vai fechar e a pasta fica para recuperar.
      if (estado.congelado) { semSuporte("LocalFileSaveChanges (congelada: a recuperacao falhou)"); return; }
      var s = Array.isArray(param) ? param.join('","') : param;
      // O indice so' avanca DEPOIS de o host confirmar que o changes0.json bate com o lote (fix round 2 da revisao
      // final do Plano 3; invariante em PonteDesktop.SalvarChanges). Qualquer coisa diferente de "ok" (o host nao
      // conseguiu gravar, ja' estava congelado, ou a chamada lancou) CONGELA a ponte deste quadro na hora: nenhum
      // lote seguinte e' mandado sobre um arquivo que nao bate, e o host fecha a janela mantendo a pasta.
      var r;
      try { r = host().SalvarChanges(s, deleteIndex, count); }
      catch (e) {
        // A chamada ao host LANCOU (fix round 3): o host nao sabe que o lote se perdeu — avisa, para ele falhar fechado
        // do mesmo jeito que num "falhou" (congelar, manter a pasta, fechar a janela).
        logar("SalvarChanges: " + e); r = "erro";
        avisar("falhaGravacao", String((e && e.message) || e));
      }
      if (r !== "ok") { estado.congelado = true; logar("LocalFileSaveChanges nao gravado (" + r + "): ponte congelada"); return; }
      var i = estado.indice || 0;
      if (deleteIndex < 0) estado.indice = i + count;              // apenda a partir do indice atual
      else if (deleteIndex < i) estado.indice = deleteIndex + count; // trunca e apenda
      else if (count > 0) estado.indice = deleteIndex + count;      // nada a truncar; so' avanca
      // else: deleteIndex >= indice e count == 0 — nada muda.
    },
    // O espelho em window.__moffice.sessaoModificada e' o que os testes de janela real leem para esperar de
    // forma DETERMINISTICA a edicao chegar ao modelo, sem tocar no host (Task 2 do Marco 1 / Plano 2). Marcado
    // tambem no documento de CIMA, como marcarPronto: o editor vive num iframe e quem le' de fora
    // (ExecutarJsAsync) esta' no topo.
    onDocumentModifiedChanged: function (m) {
      estado.sessaoModificada = !!m;
      try { if (window.top !== window && window.top.__moffice) window.top.__moffice.sessaoModificada = !!m; } catch (e) { }
      try { host().MarcarModificado(!!m); } catch (e) { }
    },

    // --- salvar: o C# roda o x2t e chama window.DesktopOfflineAppDocumentEndSave(erro, hash, senha).
    // Assinatura MEDIDA em sdkjs/word/sdk-all.js (DesktopOfflineAppDocumentStartSave):
    //   LocalFileSave(param, senha, docinfo, fileType, jsonOptions, senhaAntiga)
    // param e' "" no salvar normal e "saveas=true;" no salvar como; fileType e' 0 a menos que o chamador tenha
    // passado options.xV (asc_DownloadAs). O jsonOptions vai so' para o log — quem decide o destino e o formato e'
    // o host. A SENHA (Task 7 do Plano 3) e' a senha ATUAL do documento no editor (`void 0 === b ? editor.$ta : b`
    // nos tres bundles: a digitada ao abrir, ou a definida/removida em Arquivo -> Proteger; "" = sem senha). Vai
    // num aviso PROPRIO, ANTES do localFileSave (JS e' de uma thread so': o host a recebe primeiro), e o host nunca
    // registra o parametro desse aviso — o do localFileSave continua indo para o log, entao ela nao pode ir nele.
    LocalFileSave: function (param, senha, docinfo, fileType, jsonOptions, senhaAntiga) {
      if (estado.congelado) { semSuporte("LocalFileSave (congelada: a recuperacao falhou)"); return; }
      avisar("senhaSalvar", typeof senha === "string" ? senha : "");
      avisar("localFileSave", JSON.stringify({ param: param || "", fileType: fileType || 0, jsonOptions: jsonOptions || "" }));
    },
    // Notificacao pura, achada pelo grampo na Task 6 (aparecia como STUB no log do salvar): o sdkjs chama
    // `void 0 !== AscDesktopEditor && AscDesktopEditor.OnSave()` quando o lote de changes termina de ser
    // enviado (sdk-all-min.js, Ukg). O retorno nao e' lido.
    OnSave: function () { },

    // --- web-apps (Desktop.js): comandos de UI. A sonda mostrou a interface inteira subindo sem resposta
    // nenhuma a eles; ficam registrados no grampo para a Task 6 decidir o que responder.
    execCommand: function (cmd, json) { avisar("execCommand", JSON.stringify({ cmd: cmd || "", json: json || "" })); },

    // --- dialogo de arquivo, um por filtro do motor (Plano 6, Task 8; MvP-15/PL-08: FiltrosDoEditor no host)
    // O callback recebe o CAMINHO escolhido (string; array quando multi) — e' o sdkjs que decide o que fazer com ele (imagem:
    // LocalFileGetImageUrl(caminho), sdk-all.js:9919046; documento: convertFile/CompareDocumentFile; CSV: loadLocalFile). O host da' o
    // consentimento de UM uso que o filtro autoriza. Responder cb("") na hora, como antes, fazia o recurso nao fazer NADA em silencio.
    OpenFilenameDialog: function (filtro, multi, cb) {
      // Imagem solta do Explorer (ver estado.inserirImagemSolta): responde na hora, sem dialogo.
      if (estado.imagemSolta && filtro === "images") {
        var solta = estado.imagemSolta;
        estado.imagemSolta = null;
        if (cb) cb(solta);
        return;
      }
      if (estado.dialogoPendente) { logar("OpenFilenameDialog: ja ha um dialogo pendente"); return; }
      estado.dialogoPendente = cb || null;
      avisar("abrirArquivo", JSON.stringify({ filtro: filtro || "", multi: !!multi }));
    },
    SaveFilenameDialog: function (filtro, cb) { if (cb) cb(""); },
    // Impressao (Task 4 do Plano 3). O editor chama Print(json) ao confirmar o painel Arquivo -> Imprimir ou o
    // botao de impressao rapida (sdkjs/word NIe, slide W9f, cell OGg: `AscDesktopEditor.Print(JSON.stringify(...))`)
    // e nao espera retorno nem callback. O host gera o PDF de impressao pelo x2t (JanelaEditor, tipo "imprimir").
    // Print_Start/Page/End sao o renderizador nativo do app oficial (comandos de desenho): nao reproduzivel aqui.
    Print_Start: function () { },
    Print_Page: function () { },
    Print_End: function () { },
    Print: function (json) { avisar("imprimir", String(json || "")); },
    // Corretor (Task 1 do Plano 3). Protocolo do sdkjs 9.4 (MEDIDO em sdkjs/word/sdk-all-min.js): o editor
    // manda a string "clear" ao reiniciar o corretor e ignora toda resposta ate' recebe-la de volta; o resto e'
    // JSON ({"type":"spell"|"suggest"|"add"|"clear",...}) e o host responde com o MESMO objeto + usrCorrect/
    // usrSuggest (PedidoOrtografia no C#). A resposta vai como OBJETO (o tratador le' w.ParagraphId), sempre
    // assincrona e na ORDEM dos pedidos — a fila garante isso mesmo com o proxy assincrono do objeto de host
    // (hostObjects.ponte sem .sync: o renderer nao espera o C#). A resposta vai para o asc_nativeOnSpellCheck
    // DESTE quadro (quem pediu).
    SpellCheck: function (s) {
      var quadro = window;
      filaOrtografia = filaOrtografia.then(function () {
        if (s === "clear") { if (quadro.asc_nativeOnSpellCheck) quadro.asc_nativeOnSpellCheck("clear"); return; }
        return chrome.webview.hostObjects.ponte.VerificarOrtografia(String(s)).then(function (r) {
          if (!r) return;
          // Espelho para os testes de janela real (lido do documento de cima, como pronto/sessaoModificada).
          estado.ultimaOrtografia = r;
          try { if (window.top !== window && window.top.__moffice) window.top.__moffice.ultimaOrtografia = r; } catch (e) { }
          if (quadro.asc_nativeOnSpellCheck) quadro.asc_nativeOnSpellCheck(JSON.parse(r));
        });
      // O .catch fica DENTRO da atribuicao: uma falha num pedido nao pode deixar a fila rejeitada (todo
      // pedido seguinte seria pulado em silencio e o corretor morreria ate' recarregar o documento).
      }).catch(function (e) { logar("SpellCheck: " + e); });
    },
    // Area de transferencia pelos BOTOES e pelo menu de contexto (Task 6 do Plano 3). O teclado nao passa por
    // aqui: Ctrl+C/V/X sao eventos nativos do navegador. Quem chama e' `window.asc_desktop_copypaste(api, nome)`
    // (sdkjs/*/sdk-all.js), que antes devolve o foco a' area de texto escondida do editor (`api.kDa(true)`) e
    // depois chama AscDesktopEditor[nome]() — o retorno NAO e' lido (o api.Copy/Cut/Paste do sdkjs devolve
    // `true` sempre, entao o web-apps nunca mostra o aviso "use o teclado").
    // Copy/Cut: a chamada acontece DENTRO do clique (ativacao do usuario), entao `execCommand` dispara o evento
    // copy/cut na area de texto com foco, e quem preenche a area de transferencia e' o proprio tratador do
    // sdkjs. Sem foco do documento (ver areaDeTransferencia), o host manda o atalho real ao quadro (CDP).
    // Paste: `execCommand("paste")` e' bloqueado em conteudo web; o host manda um Ctrl+V real ao quadro com
    // foco (CDP Input.dispatchKeyEvent com o comando "paste"), exatamente o caminho do teclado.
    Copy: function () { return areaDeTransferencia("copy", "copiar"); },
    Cut: function () { return areaDeTransferencia("cut", "recortar"); },
    Paste: function () {
      estado.ultimaAreaTransferencia = "paste:host";
      avisar("colar", "");
      return true;
    }
  };

  // MEDIDO na sonda S7 (Task 6 do Plano 3): com o documento SEM FOCO (document.hasFocus() == false) o
  // execCommand("copy"/"cut") devolve TRUE, o tratador do sdkjs roda (no cut ele APAGA a selecao) e nada chega
  // a' area de transferencia do Windows — o texto recortado se perderia. Nessa medida a ativacao do usuario
  // (navigator.userActivation.isActive) era TRUE: o que discrimina e' o FOCO do documento, nao a ativacao nem o
  // retorno do execCommand. A ativacao continua na condicao (o clique do usuario da' as duas coisas; nao foi
  // medido um caso com foco e sem ativacao), mas a guarda contra a perda de dado e' o document.hasFocus().
  // Fora disso, quem copia/recorta e' o host, com o atalho real (CDP), que o navegador trata como teclado — e
  // nada e' apagado aqui.
  function areaDeTransferencia(comando, aviso) {
    var ok = false;
    var ativo = !!(navigator.userActivation && navigator.userActivation.isActive) && document.hasFocus();
    if (ativo) {
      try { ok = document.execCommand(comando); } catch (e) { logar(comando + ": " + e); }
    }
    estado.ultimaAreaTransferencia = comando + ":" + (ok ? "execCommand" : "host");
    if (!ok) avisar(aviso, "");
    return true;
  }

  // Soltar arquivos do Explorer (Task 6 do Plano 3). O sdkjs do desktop pede os caminhos a GetDropFiles(), que o
  // WebView2 nao tem, e o File do navegador nao traz caminho. `postMessageWithAdditionalObjects` entrega os File
  // ao host COM o caminho real (CoreWebView2File.Path, preenchido pelo NAVEGADOR — a pagina nao consegue forjar
  // um caminho: um File montado em JS nao tem caminho nenhum), e o host decide: imagem entra no documento pelo
  // mesmo caminho do "Imagem do arquivo", documento abre em janela nova (JanelaEditor.ArquivosSoltos).
  // Fase de CAPTURA no documento: roda antes do ondrop do sdkjs e o impede de ver o soltar de arquivo; o
  // preventDefault tambem impede o navegador de ABRIR o arquivo solto (sair do editor). Conteudo que nao e'
  // arquivo (texto arrastado de outro programa) segue o caminho normal do sdkjs.
  function temArquivos(dt) {
    return !!dt && Array.prototype.indexOf.call(dt.types || [], "Files") >= 0;
  }
  document.addEventListener("dragover", function (ev) {
    if (temArquivos(ev.dataTransfer)) ev.preventDefault();
  }, true);
  document.addEventListener("drop", function (ev) {
    var dt = ev.dataTransfer;
    if (!temArquivos(dt) || !dt.files || dt.files.length === 0) return;
    ev.preventDefault(); ev.stopPropagation();
    estado.soltos = (estado.soltos || 0) + 1;   // lido pelos testes de janela real
    try { chrome.webview.postMessageWithAdditionalObjects("moffice:soltar", dt.files); }
    catch (e) { logar("soltar: " + e); }
  }, true);

  // Insercao de uma imagem SOLTA pelo MESMO caminho do "Imagem do arquivo" (asc_addImage, publico nos tres
  // editores): o editor chama OpenFilenameDialog("images", …) e o callback recebe o caminho — aqui, na hora,
  // sem dialogo. O caminho so' e' lido porque o host ja' deu o consentimento de um uso para ELE
  // (PermitirArquivoEscolhido); a pagina escrever em `imagemSolta` sozinha nao abre arquivo nenhum.
  // Tudo sincrono: se o editor nao chamar o dialogo, a pendencia e' desfeita no fim e um "Imagem do arquivo"
  // posterior abre o dialogo normal.
  estado.inserirImagemSolta = function (caminho) {
    var api = window.Asc && window.Asc.editor ? window.Asc.editor : window.editor;
    if (!api || typeof api.asc_addImage !== "function") return "sem asc_addImage";
    estado.imagemSolta = caminho;
    try { api.asc_addImage(); }
    catch (e) { estado.imagemSolta = null; return "erro: " + e; }
    var sobrou = estado.imagemSolta;
    estado.imagemSolta = null;
    return sobrou ? "dialogo nao pedido" : "ok";
  };

  function apiDoEditor() { return estado.api || (window.Asc && window.Asc.editor) || window.editor; }

  // Iniciar a apresentacao (Plano 6: .ppsx pelo Explorer e, na Task 7, F5/Shift+F5). MEDIDO em
  // presentationeditor/main/app.js 9.4: o atalho Ctrl+F5 do editor faz Common.NotificationCenter.trigger("preview:start", 0),
  // e o botao "Do slide atual" chama onPreview(api.getCurrentPage(), false, true) -> trigger("preview:start", n, false, false, true).
  // E' o caminho da interface (abre o painel de apresentacao); api.StartDemonstration sozinho nao abre o painel.
  // Devolve "ok" ou o motivo (o host registra no log).
  estado.iniciarApresentacao = function (doSlideAtual) {
    var api = apiDoEditor();
    if (!window.Common || !Common.NotificationCenter || !api) return "sem editor de apresentacao";
    try {
      if (doSlideAtual) Common.NotificationCenter.trigger("preview:start", api.getCurrentPage(), false, false, true);
      else Common.NotificationCenter.trigger("preview:start", 0);
    } catch (e) { return "erro: " + e; }
    return "ok";
  };

  // Entrega um reporter_command ao editor como se viesse da janela do apresentador (API publica
  // DemonstrationReporterMessages = o tratador `rJc` de sdkjs/slide, que le' `event.data` como JSON).
  function comandoDoApresentador(c) {
    try {
      var api = apiDoEditor();
      if (api && typeof api.DemonstrationReporterMessages === "function")
        api.DemonstrationReporterMessages({ data: JSON.stringify(c) });
      else logar("apresentador emulado: sem DemonstrationReporterMessages");
    } catch (e) { logar("apresentador emulado (" + c.reporter_command + "): " + e); }
  }

  // Rede de seguranca da tela cheia (fix round 1 da Task 7): todo Esc com a apresentacao em tela cheia avisa o host
  // ANTES do editor (captura no window), e o host restaura a janela sozinho se o editor nao pedir SetFullscreen(false)
  // em ~1 s — um editor que falhe no encerramento (como o endReporter ausente fazia) nao prende mais o usuario.
  window.addEventListener("keydown", function (ev) {
    if (ev.key !== "Escape" && ev.keyCode !== 27) return;
    var cheia = estado.telaCheia;
    try { if (!cheia && window.top !== window && window.top.__moffice) cheia = window.top.__moffice.telaCheia; } catch (e) { }
    if (cheia) avisar("escTelaCheia", "");
  }, true);

  // Atalhos do Office que o editor web nao tem (Plano 6, Task 7). Fase de CAPTURA: roda antes dos ouvintes do editor; o que for
  // tratado aqui para (preventDefault + stopPropagation). O resto segue para o editor como sempre (Ctrl+F, Ctrl+P, Ctrl+S, Ctrl+Z/Y, Ctrl+N).
  // Canal = a PAGINA: com as aceleradoras do navegador desligadas (Task 6) e o WebView2 focado, o WPF nao ve' estas teclas.
  window.addEventListener("keydown", function (ev) {
    var k = ev.key, ctrl = ev.ctrlKey && !ev.altKey && !ev.metaKey, shift = ev.shiftKey, nada = !ev.ctrlKey && !ev.altKey && !ev.metaKey;
    var tipo = window.PE ? "slide" : window.SSE ? "cell" : window.DE ? "word" : "";
    var acao = null;
    if (ctrl && !shift && (k === "o" || k === "O")) acao = "abrir";
    else if (ctrl && !shift && (k === "w" || k === "W" || k === "F4")) acao = "fechar";
    else if (nada && !shift && k === "F12") acao = "salvarComo";
    else if (nada && shift && k === "F12") acao = "salvar";
    else if (ctrl && !shift && k === "F6") acao = "proximaJanela";
    else if (tipo === "slide" && nada && k === "F5") {
      // F5 = do inicio; Shift+F5 = do slide atual. Durante a apresentacao (tela cheia) a tecla e' do apresentador: nao reinicia.
      var apresentando = estado.telaCheia;
      try { if (!apresentando && window.top !== window && window.top.__moffice) apresentando = window.top.__moffice.telaCheia; } catch (e) { }
      if (apresentando) return;
      ev.preventDefault(); ev.stopPropagation();
      if (ev.repeat) return;   // tecla segurada: nao reinicia a apresentacao
      logar("atalho F5: " + estado.iniciarApresentacao(shift));
      return;
    }
    else if (tipo === "cell" && ((ctrl && !shift && (k === "g" || k === "G")) || (nada && !shift && k === "F5"))) {
      ev.preventDefault(); ev.stopPropagation();
      // "Ir para" do Excel: no motor nao ha' dialogo, so' a Caixa de Nome.
      if (ev.repeat) return;
      var caixa = document.querySelector("#ce-cell-name");
      if (caixa) { caixa.focus(); if (caixa.select) caixa.select(); } else logar("atalho Ir para: sem #ce-cell-name");
      return;
    }
    if (!acao) return;
    ev.preventDefault(); ev.stopPropagation();
    if (ev.repeat) return;   // tecla segurada: o autorepeat nao reabre dialogo nem refecha janela (a tecla continua engolida)
    avisar("atalho", acao);
  }, true);

  window.AscDesktopEditor = ponte;
  // Desktop.js usa window.desktop em alguns pontos (window.desktop || window.AscDesktopEditor).
  window.desktop = window.desktop || ponte;
})();
