/* ===== Rifa Navideña · Grupo Misionero Espíritu Santo — ventas de números (v1, 23/11/2026) =====
   Cómo instalarlo (una sola vez):
   1. En la planilla Rifa_Navideña: Extensiones > Apps Script. Borrar todo y pegar este archivo entero. Guardar.
   2. Arriba, en la lista de funciones, elegir "instalar" y tocar Ejecutar (si pide permisos, aceptarlos).
      Eso arma la planilla: Léeme, Ventas, Misioneros, Resumen y Config.
   3. Implementar > Nueva implementación > tipo "Aplicación web" > Ejecutar como: "Yo" > Quién tiene acceso: "Cualquier usuario"
      > Implementar. Copiar el link que termina en /exec y pasárselo a Claude (va en la página).

   Pestañas:
   - Ventas: una fila por venta (comprador, teléfono, cantidad y números asignados). La llena la página.
   - Misioneros: un renglón por misionero. Los nombres se agregan solos desde 'Asistencia ' de la planilla
     maestra. Los coordinadores escriben "Cupo" (si alguien vende distinto) y "Pagó ($)" (lo que entregó hasta ahora).
   - Resumen: totales. Config: precio, cupo, fechas, claves. */

const TZ = "America/Argentina/Buenos_Aires";
const MAESTRA = "1ff9dWtfo3Pg8UeySRfsGUJ7FyAz53AEWu0VNmQFzmWY"; /* ¡San Pedro 2027! */
const ASISTENCIA = "Asistencia ";
const HOJA_PREMIOS = "rifas"; /* pestaña de la maestra con los premios (columna A desde la fila 5) */
const PAGO_TXT = "Podés pagar en efectivo a un coordinador (en los encuentros de los domingos) o por transferencia con el QR del grupo y mandarle el comprobante a un coordinador.";
const AMARILLO = "#fff2cc";
const MAX_POR_VENTA = 100;
/* Coordinadores que venden a voluntad (cupo 0: sin mínimo, entregan lo que venden). No están en Asistencia: se agregan al instalar. */
const COORDINADORES = ["Agustín Crivaro", "Carmela Martín y Herrera", "Guadalupe Lagomarsino", "Ignacio Zavalía", "Iñaki Gallo", "Juan Cruz Rolón", "Matilde Goñi", "Pedro Gallardo", "Pedro Alsina"];

function SS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  try { preparar(); sincronizar(false); return out(datos()); }
  catch (err) { return out({ ok: false, error: "Error del sistema: " + err.message }); }
}
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: "Pedido inválido." }); }
  try { preparar(); return out(accion(b || {})); }
  catch (err) { return out({ ok: false, error: "Error del sistema: " + err.message }); }
}

/* ---------- utilidades ---------- */
function num(v) { if (typeof v === "number") return v; const d = String(v == null ? "" : v).replace(/[^\d]/g, ""); return d ? Number(d) : 0; }
function plata(v) { if (typeof v === "number") return v; return Number(String(v == null ? "" : v).replace(/[^\d-]/g, "")) || 0; }
function norm(s) { return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim(); }
function clean(s, max) { let t = String(s == null ? "" : s).replace(/\s+/g, " ").trim().slice(0, max || 80); if (/^[=+\-@]/.test(t)) t = "'" + t; return t; }
function sinComilla(s) { return String(s == null ? "" : s).replace(/^'/, ""); }
function fmt(v, f) { return v instanceof Date ? Utilities.formatDate(v, TZ, f) : String(v == null ? "" : v); }
function conLock(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
function siNo(v) { return !/^(no|n|false|0|cerrada)$/i.test(String(v == null ? "" : v).trim()); }

/* Teléfono argentino → "549" + 10 dígitos (o "" si no se entiende). Acepta 11 2345-6789, 15 2345-6789, +54 9 11..., 0221 15 ... */
function telAR(s) {
  let d = String(s == null ? "" : s).replace(/\D/g, "");
  if (d.indexOf("00") === 0) d = d.slice(2);
  if (d.indexOf("54") === 0 && d.length >= 12) { d = d.slice(2); if (d.charAt(0) === "9") d = d.slice(1); }
  if (d.charAt(0) === "0") d = d.slice(1);
  if (d.length === 10 && d.indexOf("15") === 0) d = "11" + d.slice(2);
  if (d.length === 12) { for (const a of [2, 3, 4]) if (d.substr(a, 2) === "15") { d = d.slice(0, a) + d.slice(a + 2); break; } }
  if (d.length === 8) d = "11" + d;
  return d.length === 10 ? "549" + d : "";
}
function telLindo(d) { if (!/^549\d{10}$/.test(d)) return d; const r = d.slice(3), a = r.indexOf("11") === 0 ? 2 : 3; return "+54 9 " + r.slice(0, a) + " " + r.slice(a, 6) + "-" + r.slice(6); }

function cfg() {
  const sh = SS().getSheetByName("Config"), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (r[0] !== "") m[String(r[0]).trim()] = r[1]; });
  const s = k => String(m[k] == null ? "" : m[k]).trim();
  return {
    price: num(m["Precio del número ($)"]) || 5000,
    per: m["Números por misionero"] === "" || m["Números por misionero"] == null ? 40 : num(m["Números por misionero"]),
    meta: num(m["Meta de la rifa ($)"]),
    sorteo: s("Fecha del sorteo"), como: s("Cómo es el sorteo"), ig: s("Instagram"),
    e1: s("Primera entrega: fecha"), e1p: num(m["Primera entrega: % del total"]) || 50, e2: s("Segunda entrega: fecha"),
    abierta: siNo(m["Venta abierta"]),
    pagoTxt: s("Cómo pagan los misioneros") || PAGO_TXT,
    wa: s("WhatsApp de contacto").replace(/\D/g, ""), mail: s("Mail de contacto"),
    claveM: s("Clave de misioneros"), claveA: s("Clave de administración")
  };
}
function pub(c) {
  return { price: c.price, per: c.per, meta: c.meta, sorteo: c.sorteo, como: c.como, ig: c.ig, e1: c.e1, e1p: c.e1p, e2: c.e2,
    abierta: c.abierta, wa: c.wa, mail: c.mail, pagoTxt: c.pagoTxt, clave: !!c.claveM };
}

/* Misioneros: A Misionero · B Cupo · C Vendidos (fórmula) · D Tiene que entregar (fórmula) · E Pagó ($) · F Falta pagar (fórmula) · G Estado */
function misioneros(c) {
  const sh = SS().getSheetByName("Misioneros"), n = sh.getLastRow(), L = [];
  if (n < 2) return L;
  sh.getRange(2, 1, n - 1, 7).getValues().forEach((r, i) => {
    const k = String(r[0]).trim();
    if (k) L.push({ row: i + 2, n: k, c: (r[1] === "" || isNaN(Number(r[1]))) ? c.per : Math.max(0, Math.round(Number(r[1]))), pag: plata(r[4]), est: String(r[6] || "") });
  });
  return L;
}
/* Ventas: A Comprador · B Teléfono · C Cantidad · D Del N° · E Al N° · F Misionero · G Cargada el · H ID */
function ventas() {
  const sh = SS().getSheetByName("Ventas"), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 8).getValues().map((r, i) => ({
    row: i + 2, n: sinComilla(r[0]), tel: sinComilla(r[1]), q: Math.max(0, Math.round(Number(r[2]) || 0)), del: Number(r[3]) || 0, al: Number(r[4]) || 0,
    kid: String(r[5]).trim(), f: fmt(r[6], "dd/MM HH:mm"), id: String(r[7])
  })).filter(x => x.id && x.n);
}
function nuevoId(L, pre) { const s = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789", usados = {}; L.forEach(x => usados[x.id] = 1); let id; do { id = (pre || "V") + "-"; for (let i = 0; i < 5; i++) id += s[Math.floor(Math.random() * s.length)]; } while (usados[id]); return id; }
/* Próximo número libre: nunca se reusan números borrados (así nadie queda con un número que ya le mandaron a otro). */
function proximo(L) {
  const props = PropertiesService.getScriptProperties();
  let ult = Number(props.getProperty("ult")) || 0;
  L.forEach(x => { if (x.al > ult) ult = x.al; });
  return ult + 1;
}
function usar(hasta) { PropertiesService.getScriptProperties().setProperty("ult", String(hasta)); }
function agregarVenta(fila) {
  const sh = SS().getSheetByName("Ventas"); sh.appendRow(fila);
  const r = sh.getLastRow(); sh.getRange(r, 4, 1, 2).setNumberFormat("0000"); sh.getRange(r, 7).setNumberFormat("dd/MM HH:mm");
}
/* Filas cargadas a mano en Ventas (comprador, cantidad y misionero): les pone números e ID */
function completarAMano() {
  const sh = SS().getSheetByName("Ventas"), n = sh.getLastRow();
  if (n < 2) return;
  const v = sh.getRange(2, 1, n - 1, 8).getValues();
  if (!v.some(r => String(r[0]).trim() && !String(r[7]).trim())) return;
  conLock(() => {
    const w = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues(), L = ventas();
    let next = proximo(L);
    w.forEach((r, i) => {
      if (!String(r[0]).trim() || String(r[7]).trim()) return;
      const q = Math.max(1, Math.round(Number(r[2]) || 1)), id = nuevoId(L, "V"); L.push({ id });
      sh.getRange(i + 2, 3, 1, 3).setValues([[q, next, next + q - 1]]); sh.getRange(i + 2, 4, 1, 2).setNumberFormat("0000");
      if (r[6] === "") sh.getRange(i + 2, 7).setValue(new Date());
      sh.getRange(i + 2, 8).setValue(id); next += q; usar(next - 1);
    });
  });
}

/* ---------- premios (pestaña "rifas" de la maestra) ---------- */
function premios() {
  const cache = CacheService.getScriptCache(), c = cache.get("premios");
  if (c) return JSON.parse(c);
  let L = [];
  try {
    const sh = SpreadsheetApp.openById(MAESTRA).getSheetByName(HOJA_PREMIOS);
    if (sh) L = sh.getRange("A5:A80").getDisplayValues().map(r => String(r[0]).replace(/\s+/g, " ").trim()).filter(x => x && !/^premio$/i.test(x));
  } catch (e) { L = []; }
  cache.put("premios", JSON.stringify(L), 600);
  return L;
}

/* ---------- datos para la portada ---------- */
function totales(c, L, M) {
  const vend = L.filter(x => x.id.indexOf("C-") !== 0).reduce((s, x) => s + x.q, 0);
  const pag = M.reduce((s, x) => s + x.pag, 0);
  return { vend, pag };
}
function datos() {
  completarAMano();
  const c = cfg(), L = ventas(), M = misioneros(c), cnt = {};
  L.forEach(x => { if (x.id.indexOf("C-") !== 0) cnt[x.kid] = (cnt[x.kid] || 0) + x.q; });
  return { ok: true, v: 1, config: pub(c), premios: premios(), tot: totales(c, L, M),
    chicos: M.map(x => ({ n: x.n, c: x.c, k: cnt[x.n] || 0 })) };
}

/* ---------- acciones de misioneros ---------- */
function accion(b) {
  const c = cfg();
  if (b.action === "admin") return admin(b, c);
  const kid = String(b.kid || "").trim(), yo = misioneros(c).filter(x => x.n === kid)[0];
  if (!yo) return { ok: false, error: "No encontramos tu nombre en la lista de misioneros. Avisale a un coordinador." };
  if (c.claveM && String(b.clave || "").trim() !== c.claveM) return { ok: false, clave: true, error: "La clave de misioneros no es correcta." };
  if (b.action === "mis") return mis(yo, c);
  if (b.action === "cargar") return conLock(() => {
    if (!c.abierta) return { ok: false, error: "La venta de la rifa ya cerró. Si te faltó anotar algo, hablá con un coordinador." };
    const name = clean(b.name, 70), tel = telAR(b.tel), q = Math.round(Number(b.q));
    if (!/\S{2,}\s+\S{2,}/.test(sinComilla(name))) return { ok: false, error: "Escribí nombre y apellido de quien compró." };
    if (!tel) return { ok: false, error: "Revisá el celular: escribilo con código de área, por ejemplo 11 2345-6789." };
    if (!(q >= 1 && q <= MAX_POR_VENTA)) return { ok: false, error: "La cantidad de números tiene que ser entre 1 y " + MAX_POR_VENTA + "." };
    const L = ventas(), tl = telLindo(tel);
    const d = L.filter(x => x.kid === kid && norm(x.n) === norm(sinComilla(name)) && x.tel.replace(/\D/g, "") === tel)[0];
    if (d && !b.force) return { ok: false, dup: { n: d.n, q: d.q, f: d.f } };
    const del = proximo(L), al = del + q - 1, id = nuevoId(L, "V");
    agregarVenta([name, "'" + tl, q, del, al, kid, new Date(), id]);
    usar(al);
    const r = mis(yo, c); r.nuevo = { id, n: sinComilla(name), tel: tl, q, del, al }; return r;
  });
  if (b.action === "borrar") return conLock(() => {
    const x = ventas().filter(r => r.id === String(b.id) && r.kid === kid)[0];
    if (!x) return { ok: false, error: "No encontramos esa venta (¿ya la borraste?)." };
    if (!c.abierta || x.id.indexOf("C-") === 0) return { ok: false, error: "La venta ya cerró: para corregir algo, hablá con un coordinador." };
    SS().getSheetByName("Ventas").deleteRow(x.row);
    return mis(yo, c);
  });
  return { ok: false, error: "Acción desconocida." };
}
function mis(yo, c) {
  const L = ventas(), items = L.filter(x => x.kid === yo.n).map(x => ({ id: x.id, n: x.n, tel: x.tel, q: x.q, del: x.del, al: x.al, cierre: x.id.indexOf("C-") === 0 }));
  const vend = items.filter(x => !x.cierre).reduce((s, x) => s + x.q, 0);
  return { ok: true, kid: yo.n, cupo: yo.c, items, vend, pagado: yo.pag, debe: Math.max(yo.c, vend) * c.price, price: c.price, pagoTxt: c.pagoTxt,
    abierta: c.abierta, tot: totales(c, L, misioneros(c)) };
}

/* ---------- administración ---------- */
function admin(b, c) {
  if (!c.claveA || String(b.clave || "") !== c.claveA) return { ok: false, error: "Clave incorrecta." };
  const op = b.op;
  if (op === "login") return { ok: true };
  if (op === "todo") { sincronizar(false); completarAMano(); return { ok: true, config: pub(c), premios: premios(),
    ventas: ventas().map(x => ({ id: x.id, f: x.f, kid: x.kid, n: x.n, tel: x.tel, q: x.q, del: x.del, al: x.al })),
    chicos: misioneros(c).map(x => ({ n: x.n, c: x.c, pag: x.pag, est: x.est })) }; }
  if (op === "borrar") return conLock(() => {
    const x = ventas().filter(r => r.id === String(b.id))[0];
    if (!x) return { ok: false, error: "No encontramos esa venta." };
    SS().getSheetByName("Ventas").deleteRow(x.row); return { ok: true };
  });
  if (op === "pagado") return conLock(() => {
    const x = misioneros(c).filter(m => m.n === String(b.kid || "").trim())[0], t = (b.total === null || b.total === "" || b.total === undefined) ? NaN : Math.round(Number(b.total));
    if (!x) return { ok: false, error: "No encontramos a ese misionero." };
    if (!(t >= 0) || t > 20000000) return { ok: false, error: "Escribí un monto válido." };
    SS().getSheetByName("Misioneros").getRange(x.row, 5).setValue(t || "");
    return { ok: true, pag: t };
  });
  if (op === "cierre") return conLock(() => {
    if (!c.abierta) return { ok: false, error: "La venta ya estaba cerrada." };
    const L = ventas(), M = misioneros(c), cnt = {}, d = new Date();
    L.forEach(x => { if (x.id.indexOf("C-") !== 0) cnt[x.kid] = (cnt[x.kid] || 0) + x.q; });
    let next = proximo(L), filas = 0, nums = 0;
    M.forEach(m => {
      if (m.est) return; /* bajas: no se les asignan números */
      const falta = m.c - (cnt[m.n] || 0);
      if (falta <= 0) return;
      const id = nuevoId(L, "C"); L.push({ id });
      agregarVenta([m.n + " (sin vender)", "", falta, next, next + falta - 1, m.n, d, id]);
      next += falta; filas++; nums += falta;
    });
    if (next > 1) usar(next - 1);
    const sh = SS().getSheetByName("Config"), v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    v.forEach((r, i) => { if (String(r[0]).trim() === "Venta abierta") sh.getRange(i + 2, 2).setValue("No"); });
    return { ok: true, filas, nums };
  });
  return { ok: false, error: "Operación desconocida." };
}

/* ---------- nombres de misioneros desde la planilla maestra ---------- */
function nombresMaestra() {
  const sh = SpreadsheetApp.openById(MAESTRA).getSheetByName(ASISTENCIA);
  if (!sh) return null;
  const vistos = {}, L = [];
  sh.getRange("A2:A400").getDisplayValues().forEach(r => {
    const n = String(r[0]).replace(/\s+/g, " ").trim();
    if (n && !/:$/.test(n) && isNaN(Number(n)) && !vistos[n]) { vistos[n] = 1; L.push(n); }
  });
  return L;
}
function sincronizar(forzar) {
  const cache = CacheService.getScriptCache();
  if (!forzar && cache.get("sync")) return;
  let maestra;
  try { maestra = nombresMaestra(); } catch (e) { maestra = null; }
  if (!maestra || !maestra.length) { cache.put("sync", "1", 120); return; }
  conLock(() => {
    const props = PropertiesService.getScriptProperties(), auto = JSON.parse(props.getProperty("auto") || "{}");
    const sh = SS().getSheetByName("Misioneros"), n = sh.getLastRow(), c = cfg();
    const filas = n >= 2 ? sh.getRange(2, 1, n - 1, 7).getValues().filter(r => String(r[0]).trim()) : [];
    const cnt = {}; ventas().forEach(x => cnt[x.kid] = (cnt[x.kid] || 0) + x.q);
    const enMaestra = {}; maestra.forEach(k => enMaestra[k] = 1);
    const tengo = {}; let cambio = false;
    const nuevas = [];
    filas.forEach(r => {
      const k = String(r[0]).trim(); tengo[k] = 1;
      let est = String(r[6] || "");
      if (enMaestra[k]) est = "";
      else if (auto[k]) {
        if (!cnt[k] && (r[4] === "" || plata(r[4]) === 0)) { delete auto[k]; cambio = true; return; } /* baja sin datos: se va */
        est = "Ya no está en Asistencia";
      }
      if (est !== String(r[6] || "")) cambio = true;
      nuevas.push([k, r[1], r[4], est]);
    });
    maestra.forEach(k => { if (!tengo[k]) { nuevas.push([k, "", "", ""]); auto[k] = 1; cambio = true; } });
    const orden = nuevas.slice().sort((a, b) => a[0].localeCompare(b[0], "es"));
    if (cambio) {
      const m = orden.length, viejo = Math.max(0, n - 1);
      if (m) {
        sh.getRange(2, 1, m, 2).setValues(orden.map(r => [r[0], r[1]]));
        sh.getRange(2, 5, m, 1).setValues(orden.map(r => [r[2]]));
        sh.getRange(2, 7, m, 1).setValues(orden.map(r => [r[3]]));
      }
      if (viejo > m) { sh.getRange(m + 2, 1, viejo - m, 2).clearContent(); sh.getRange(m + 2, 5, viejo - m, 1).clearContent(); sh.getRange(m + 2, 7, viejo - m, 1).clearContent(); }
      props.setProperty("auto", JSON.stringify(auto));
    }
  });
  cache.put("sync", "1", 300);
}

/* ---------- instalación (corre una sola vez) ---------- */
function instalar() { preparar(true); sincronizar(true); agregarCoordinadores(); }
function agregarCoordinadores() {
  const sh = SS().getSheetByName("Misioneros"), n = sh.getLastRow();
  const ya = n >= 2 ? sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  COORDINADORES.filter(k => ya.indexOf(k) < 0).forEach((k, i) => { sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setValues([[k, 0]]); });
}
const F_VEND = '={"Vendidos";ARRAYFORMULA(IF(A2:A="","",SUMIFS(Ventas!C2:C,Ventas!F2:F,A2:A,Ventas!H2:H,"V-*")))}';
const F_DEBE = '={"Tiene que entregar ($)";ARRAYFORMULA(IF(A2:A="","",IF(C2:C>IF(B2:B="",Config!B3,B2:B),C2:C,IF(B2:B="",Config!B3,B2:B))*Config!B2))}';
const F_FALTA = '={"Falta pagar ($)";ARRAYFORMULA(IF(A2:A="","",D2:D-IF(E2:E="",0,E2:E)))}';
function claveNueva() { const s = "abcdefghjkmnpqrstuvwxyz23456789"; let k = "rifa-"; for (let i = 0; i < 6; i++) k += s[Math.floor(Math.random() * s.length)]; return k; }
function preparar(forzar) {
  if (SS().getSheetByName("Ventas")) return; /* ya instalado */
  conLock(() => {
    const ss = SS();
    if (ss.getSheetByName("Ventas")) return;
    const sfx = "_" + Date.now();
    const le = ss.insertSheet("_leeme" + sfx), ve = ss.insertSheet("_ventas" + sfx), mi = ss.insertSheet("_misioneros" + sfx), rs = ss.insertSheet("_resumen" + sfx), co = ss.insertSheet("_config" + sfx);
    const head = (sh, n) => sh.getRange(1, 1, 1, n).setFontWeight("bold").setBackground("#553a1e").setFontColor("#ffffff");

    /* Ventas */
    ve.getRange(1, 1, 1, 8).setValues([["Comprador", "Teléfono", "Cantidad", "Del N°", "Al N°", "Misionero", "Cargada el", "ID"]]); head(ve, 8);
    ve.setFrozenRows(1); ve.getRange("B2:B6000").setNumberFormat("@"); ve.getRange("D2:E6000").setNumberFormat("0000"); ve.getRange("G2:G6000").setNumberFormat("dd/MM HH:mm");
    ve.getRange("H1:H6000").setFontColor("#999999");
    [240, 170, 80, 80, 80, 220, 110, 80].forEach((w, i) => ve.setColumnWidth(i + 1, w));

    /* Config */
    const conf = [
      ["Parámetro", "Valor", "Para qué sirve"],
      ["Precio del número ($)", 5000, "Lo cobra el misionero y se lo entrega a los coordinadores."],
      ["Números por misionero", 40, "Cupo de cada misionero (mínimo que paga aunque no lo venda). En la pestaña Misioneros se puede cambiar el de cada uno."],
      ["Meta de la rifa ($)", 22000000, "Lo que tiene que dejar la rifa. Se muestra como barra de avance en la página."],
      ["Fecha del sorteo", "sábado 26/12", "Se muestra en la página y en el mensaje de WhatsApp."],
      ["Cómo es el sorteo", "En vivo por las historias de Instagram", "Se muestra en la página y en el mensaje de WhatsApp."],
      ["Instagram", "@mision_espiritusanto", "Cuenta donde se hace el sorteo."],
      ["Primera entrega: fecha", "domingo 13/12", "Para esa fecha cada misionero entrega una parte de la plata."],
      ["Primera entrega: % del total", 50, "Qué parte entrega en la primera fecha (50 = la mitad)."],
      ["Segunda entrega: fecha", "domingo 20/12", "Para esa fecha tiene que estar todo pagado."],
      ["Venta abierta", "Sí", "\"No\" = los misioneros ya no pueden anotar ni borrar ventas. Se pone sola en \"No\" con el botón Cerrar la venta del admin."],
      ["Cómo pagan los misioneros", PAGO_TXT, "Texto que ve cada misionero debajo de lo que tiene que entregar."],
      ["WhatsApp de contacto", "5491170544319", "Número para dudas (solo dígitos, con 549)."],
      ["Mail de contacto", "ignaciozavalia1@gmail.com", "Mail para consultas sobre datos (privacidad)."],
      ["Clave de misioneros", "", "Opcional. Vacío = sin clave."],
      ["Clave de administración", claveNueva(), "Clave de la pantalla de coordinadores (.../rifa/#admin). Se generó sola: compartila solo con coordinadores."]];
    co.getRange(1, 1, conf.length, 3).setValues(conf); head(co, 3);
    co.getRange(2, 2, conf.length - 1, 1).setBackground(AMARILLO); co.getRange("B2").setNumberFormat("$#,##0"); co.getRange("B4").setNumberFormat("$#,##0"); co.setFrozenRows(1);
    co.setColumnWidth(1, 230); co.setColumnWidth(2, 330); co.setColumnWidth(3, 560);

    /* Misioneros */
    mi.getRange(1, 1, 1, 7).setValues([["Misionero", "Cupo", "", "", "Pagó ($)", "", "Estado"]]); head(mi, 7);
    mi.setFrozenRows(1); mi.getRange("D2:F400").setNumberFormat("$#,##0"); mi.getRange("B2:B400").setBackground(AMARILLO); mi.getRange("E2:E400").setBackground(AMARILLO);
    mi.setColumnWidth(1, 240); mi.setColumnWidth(4, 150); mi.setColumnWidth(6, 130); mi.setColumnWidth(7, 200);

    /* Resumen (las fórmulas se escriben al final, cuando las pestañas ya tienen su nombre) */
    /* Léeme */
    const t = [
      ["Rifa Navideña · Grupo Misionero Espíritu Santo — Ventas de números"],
      ["Cada misionero vende sus números, cobra la plata y anota en la página cada venta (nombre, celular y cantidad). La página le asigna los números al comprador y le arma el mensaje de WhatsApp para mandárselos."],
      [""],
      ["Ventas → una fila por venta. La llena la página (no hace falta tocarla). Si alguien da una venta en papel, se puede escribir a mano: Comprador, Teléfono, Cantidad y Misionero; la página le pone los números y el ID sola."],
      ["Misioneros → un renglón por misionero. Los nombres se agregan solos desde la planilla maestra (Asistencia)."],
      ["      Columnas amarillas = las completan los coordinadores:"],
      ["      · Pagó ($): lo que entregó el misionero HASTA AHORA (el total, no lo de hoy). Ejemplo: pagó 100.000 y después 100.000 → escribir 200000."],
      ["      · Cupo: vacío = el de Config. Cambiarlo solo si alguien vende un cupo distinto. Cupo 0 = vende a voluntad (los coordinadores): entrega lo que vende."],
      ["      'Tiene que entregar' = el cupo × precio, o lo que vendió si vendió más. Lo demás se calcula solo."],
      ["      ¿Alguien que no está en Asistencia? Escribir su nombre en una fila vacía al final."],
      ["Resumen → totales de números y plata."],
      ["Config → precio, cupo, fechas, meta y claves (celdas amarillas)."],
      [""],
      ["Al terminar la venta: Admin > Sorteo > Cerrar la venta. Los números que cada misionero no vendió quedan a su nombre (los pagó) y participan del sorteo."],
      ["Los números borrados no se vuelven a usar (por eso puede haber saltos en la numeración)."],
      [""],
      ["Página para misioneros: https://misionespiritusanto.github.io/Pena/rifa/"],
      ["Página de coordinadores: https://misionespiritusanto.github.io/Pena/rifa/#admin  (clave: 'Clave de administración' de Config)"],
      ["No cambiar nombres de pestañas ni títulos de columnas. No ordenar ni insertar columnas en Misioneros: los nombres se corrigen en Asistencia."]];
    le.getRange(1, 1, t.length, 1).setValues(t); le.getRange("A1").setFontWeight("bold").setFontSize(14); le.setColumnWidth(1, 980);

    /* Borrar lo demás y dejar los nombres definitivos */
    const nuevas = [le, ve, mi, rs, co].map(s => s.getSheetId());
    ss.getSheets().forEach(s => { if (nuevas.indexOf(s.getSheetId()) < 0) ss.deleteSheet(s); });
    SpreadsheetApp.flush();
    le.setName("Léeme"); ve.setName("Ventas"); mi.setName("Misioneros"); rs.setName("Resumen"); co.setName("Config");
    mi.getRange(1, 3).setFormula(F_VEND); mi.getRange(1, 4).setFormula(F_DEBE); mi.getRange(1, 6).setFormula(F_FALTA);
    const CUPOS = 'ARRAYFORMULA(SUMPRODUCT((Misioneros!A2:A<>"")*(Misioneros!G2:G="")*IF(Misioneros!B2:B="",Config!B3,Misioneros!B2:B)))';
    const res = [
      ["Números", ""],
      ["Números vendidos (anotados por misioneros)", '=SUMIF(Ventas!H2:H,"V-*",Ventas!C2:C)'],
      ["Ventas anotadas", '=COUNTIF(Ventas!H2:H,"V-*")'],
      ["Números de los cupos", "=" + CUPOS],
      ["Números a nombre de misioneros (cierre)", '=SUMIF(Ventas!H2:H,"C-*",Ventas!C2:C)'],
      ["Total de números en el sorteo", "=SUM(Ventas!C2:C)"],
      ["", ""],
      ["Plata", ""],
      ["Tienen que entregar los misioneros", '=SUMIF(Misioneros!G2:G,"",Misioneros!D2:D)'],
      ["Ya entregaron", "=SUM(Misioneros!E2:E)"],
      ["Falta que entreguen", "=B9-B10"],
      ["Meta de la rifa", "=Config!B4"],
      ["Avance (entregado / meta)", '=IF(B12=0,"",B10/B12)']];
    rs.getRange(1, 1, res.length, 2).setValues(res);
    rs.getRange("A1").setFontWeight("bold"); rs.getRange("A8").setFontWeight("bold");
    rs.getRange("B9:B12").setNumberFormat("$#,##0"); rs.getRange("B13").setNumberFormat("0%"); rs.setColumnWidth(1, 340); rs.setColumnWidth(2, 140);
    [le, ve, mi, rs, co].forEach((s, i) => { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); });
    ss.setActiveSheet(le);
    PropertiesService.getScriptProperties().deleteProperty("auto");
    PropertiesService.getScriptProperties().deleteProperty("ult");
  });
  CacheService.getScriptCache().remove("sync");
}
