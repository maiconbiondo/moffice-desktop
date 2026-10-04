// Cliente CDP minimo do roteiro ao vivo do Plano 6. So' fala com a PAGINA do mOffice (alvo cuja URL contem
// moffice.local/web-apps/apps/api/documents/index.html). Nunca captura a area de trabalho.
//   node tools/ao-vivo/cdp.mjs <porta> eval "<js com w e api>"
//   node tools/ao-vivo/cdp.mjs <porta> teclas "Control+o" | "F5" | "Shift+F12"
//   node tools/ao-vivo/cdp.mjs <porta> captura <arquivo.png>
import { writeFileSync } from "node:fs";
const [porta, cmd, arg] = process.argv.slice(2);
const alvos = await (await fetch(`http://127.0.0.1:${porta}/json`)).json();
const alvo = alvos.find(a => a.type === "page" && a.url.includes("moffice.local/web-apps/apps/api/documents/index.html"));
if (!alvo) { console.log("sem pagina do editor nesta porta"); process.exit(1); }
const ws = new WebSocket(alvo.webSocketDebuggerUrl);
let id = 0; const pendentes = new Map();
ws.onmessage = m => { const r = JSON.parse(m.data); if (r.id && pendentes.has(r.id)) { pendentes.get(r.id)(r); pendentes.delete(r.id); } };
await new Promise(r => ws.onopen = r);
const chamar = (method, params = {}) => new Promise(r => { const i = ++id; pendentes.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
const TECLAS = { F5: [116, "F5"], F12: [123, "F12"], Escape: [27, "Escape"], o: [79, "KeyO"], w: [87, "KeyW"], g: [71, "KeyG"], F4: [115, "F4"], F6: [117, "F6"] };
if (cmd === "eval") {
  const js = `(function(){var w=window.frames[0];var api=w&&((w.Asc&&w.Asc.editor)||w.editor);return JSON.stringify((function(){${arg}})());})()`;
  const r = await chamar("Runtime.evaluate", { expression: js, awaitPromise: true, returnByValue: true });
  console.log(r.result?.result?.value ?? JSON.stringify(r.result));
} else if (cmd === "teclas") {
  const partes = arg.split("+"); const tecla = partes.pop();
  const mods = (partes.includes("Alt") ? 1 : 0) | (partes.includes("Control") ? 2 : 0) | (partes.includes("Shift") ? 8 : 0);
  const [vk, code] = TECLAS[tecla] ?? [tecla.toUpperCase().charCodeAt(0), "Key" + tecla.toUpperCase()];
  for (const type of ["rawKeyDown", "keyUp"])
    await chamar("Input.dispatchKeyEvent", { type, modifiers: mods, windowsVirtualKeyCode: vk, code, key: tecla });
  console.log("ok");
} else if (cmd === "captura") {
  const r = await chamar("Page.captureScreenshot", { format: "png" });
  writeFileSync(arg, Buffer.from(r.result.data, "base64")); console.log("ok " + arg);
} else { console.log("comando desconhecido"); }
ws.close();
