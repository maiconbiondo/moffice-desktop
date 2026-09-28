// grampo.js — injetado via AddScriptToExecuteOnDocumentCreatedAsync em TODA pagina/iframe.
// Se window.AscDesktopEditor ja existir (definido por ponte-desktop.js, Task 5A) embrulha; se nao, cria um stub
// que loga e devolve valores neutros — assim a sonda (Task 4) mostra o que o editor pede sem travar.
(function () {
  // "log" (PonteLog) e o unico objeto de host que o quadro do editor recebe alem de "ponte" (achado 1 da
  // revisao 2026-09-17: "arquivos"/PonteArquivos tem LerBase64/GravarBase64/dialogos = I/O de arquivo
  // arbitrario e NAO vai para o quadro, que roda documento de terceiros). O grampo roda tanto na janela
  // (onde "arquivos" ainda existe, para a pagina de teste) quanto no quadro — usar "log" funciona nos dois.
  var log = function (linha) {
    try { chrome.webview.hostObjects.sync.log.Log(linha); } catch (e) { console.log("[grampo] " + linha); }
  };
  // O grampo roda DENTRO da pilha de chamada do editor: se ele lanca, quem quebra e o editor. Observado na
  // sonda (Task 4): d.YCa do sdk-all-min passa um objeto ciclico (editor -> KXb -> editor) e o JSON.stringify
  // direto derrubava o carregamento com "Converting circular structure to JSON".
  var descrever = function (args) {
    try { return JSON.stringify(args).slice(0, 300); }
    catch (e) {
      var partes = [];
      for (var i = 0; i < args.length; i++) {
        try { partes.push(args[i] === null ? "null" : typeof args[i] === "object" ? "<" + ((args[i].constructor && args[i].constructor.name) || "object") + ">" : String(args[i]).slice(0, 60)); }
        catch (e2) { partes.push("<?>"); }
      }
      return "[" + partes.join(",") + "] (nao serializavel)";
    }
  };
  // Argumentos que levam SENHA de documento (Task 7 do Plano 3): nunca vao para o log, nem com o grampo ligado.
  // Indices MEDIDOS nos bundles 9.4.0.129 / desktop-sdk: SetAdvancedOptions(xml com <m_sPassword>),
  // LocalFileSave(param, SENHA, docinfo, fileType, jsonOptions, SENHA_ANTIGA), NativeViewerOpen(SENHA),
  // buildCryptedStart(dados, docinfo, SENHA, ...), _SetCryptoMode/SetCryptoMode(SENHA, ...).
  var sigilosos = { SetAdvancedOptions: [0], LocalFileSave: [1, 5], NativeViewerOpen: [0], buildCryptedStart: [2],
                    SetCryptoMode: [0], _SetCryptoMode: [0] };
  var mascarar = function (nome, args) {
    var idx = sigilosos[nome];
    if (!idx) return args;
    var copia = args.slice();
    for (var i = 0; i < idx.length; i++) if (idx[i] < copia.length) copia[idx[i]] = "<omitido>";
    return copia;
  };
  var alvo = window.AscDesktopEditor || {};
  // Cuidado: a armadilha "has" devolve true para tudo e "get" devolve uma funcao para qualquer nome, entao
  // guardas do tipo `if (window.AscDesktopEditor.Foo)` SEMPRE passam — um valor neutro errado vira um
  // TypeError no editor. Observado na sonda (Task 4): index.html do editor faz
  // `supportedScaleValues = AscDesktopEditor.GetSupportedScaleValues()` e logo le `.length` — com undefined,
  // o bootstrap morre antes de qualquer outra chamada de ponte.
  var neutros = { CheckUserId: "1", LocalFileGetSourcePath: "", LocalFileGetSaved: true, LocalFileGetOpenChangesCount: 0,
                  LocalFileGetCurrentChangesIndex: 0, IsLocalFile: true, LocalFileGetImageUrl: "",
                  IsSupportNativePrint: false, isSupportBinaryFontsSprite: false,
                  GetSupportedScaleValues: [1, 1.25, 1.5, 1.75, 2, 2.25, 2.5, 2.75, 3, 3.5, 4, 4.5, 5],
                  // Valores iguais ao padrao que o proprio sdkjs usa quando NAO ha AscDesktopEditor
                  // (sdkjs/word/sdk-all-min.js: `...GetEncryptedHeader?...():"ENCRYPTED;"` e
                  // `...GetFontThumbnailHeight?...():28`) — o stub nao pode devolver algo "mais neutro" que isso.
                  GetEncryptedHeader: "ENCRYPTED;", GetFontThumbnailHeight: 28, getFontsSprite: "",
                  // Formatos exigidos, descobertos na sonda: UpdateSystemPlugins (sdk-all-min) faz
                  // JSON.parse(GetInstallPlugins()) e le a[0].url / a[1].url / a[].pluginsData — precisa ser
                  // um JSON com DOIS elementos (sistema e usuario); on_native_message (app.js) le
                  // getViewportSettings().widgetType.
                  GetInstallPlugins: '[{"url":"","pluginsData":[]},{"url":"","pluginsData":[]}]',
                  getViewportSettings: {} };
  window.AscDesktopEditor = new Proxy(alvo, {
    get: function (obj, prop) {
      if (prop in obj) {
        var v = obj[prop];
        if (typeof v !== "function") return v;
        return function () { var args = Array.prototype.slice.call(arguments); log("ponte " + String(prop) + " " + descrever(mascarar(String(prop), args))); return v.apply(obj, args); };
      }
      if (typeof prop === "symbol") return undefined;
      return function () { var args = Array.prototype.slice.call(arguments); log("STUB " + String(prop) + " " + descrever(mascarar(String(prop), args))); return neutros[prop]; };
    },
    has: function (obj, prop) { return true; }
  });
  log("grampo instalado em " + location.href);
})();
