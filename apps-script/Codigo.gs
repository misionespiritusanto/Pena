/* ===== Peña · Grupo Misionero Espíritu Santo — lista de entradas (v2, noviembre 2026) =====
   Cómo instalarlo: en la planilla Pena_Entradas, Extensiones > Apps Script, borrar todo el código
   que haya y pegar este archivo entero. Guardar, y en Implementar > Gestionar implementaciones >
   lápiz > Versión: "Nueva versión" > Implementar. El link de la web app no cambia.

   La primera vez que alguien abre la página, el script prepara la planilla solo:
   crea las pestañas "Lista" y "Pagos misioneros", pone 4 entradas por chico, rehace las
   columnas de "Chicos" y el "Resumen", reescribe el "Léeme" y renombra "Entradas" a
   "Entradas (sistema viejo)". */

const TZ = "America/Argentina/Buenos_Aires";
const PUERTA = "Venta en puerta";
const H_LISTA = ["ID", "Cargada", "Misionero", "Nombre y apellido", "Teléfono", "Ingresó (hora)", "Notas"];
const H_PAGOS = ["Fecha", "Misionero", "Monto ($)", "Medio", "Nota"];
const PAGO_TXT = "Entregale la plata a Nacho: en efectivo en los encuentros de los domingos o como te digan los coordinadores.";

function SS() { return SpreadsheetApp.getActiveSpreadsheet(); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  try { preparar(); return out(datos()); }
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
function norm(s) { return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim(); }
function clean(s, max) { let t = String(s == null ? "" : s).replace(/\s+/g, " ").trim().slice(0, max || 80); if (/^[=+\-@]/.test(t)) t = "'" + t; return t; }
function fmt(v, f) { return v instanceof Date ? Utilities.formatDate(v, TZ, f) : String(v == null ? "" : v); }
function conLock(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

function cfg() {
  const sh = SS().getSheetByName("Config"), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (r[0] !== "") m[String(r[0]).trim()] = r[1]; });
  return {
    price: num(m["Precio de la entrada ($)"]) || 15000,
    per: num(m["Entradas por chico"]) || 4,
    date: String(m["Fecha y hora"] || ""),
    where: String(m["Lugar"] || ""),
    wa: String(m["WhatsApp de contacto"] || "").replace(/\D/g, ""),
    mail: String(m["Mail de contacto"] || ""),
    pagoTxt: String(m["Cómo pagan los misioneros"] || PAGO_TXT),
    cap: num(m["Capacidad máxima del salón"]) || 0,
    claveM: String(m["Clave de misioneros"] == null ? "" : m["Clave de misioneros"]).trim(),
    claveA: String(m["Clave de administración"] == null ? "" : m["Clave de administración"]).trim()
  };
}
function chicos(c) {
  const sh = SS().getSheetByName("Chicos"), n = sh.getLastRow(), L = [];
  if (n < 2) return L;
  sh.getRange(2, 1, n - 1, 2).getDisplayValues().forEach(r => { const k = String(r[0]).trim(); if (k) L.push({ n: k, c: num(r[1]) || c.per }); });
  return L;
}
function lista() {
  const sh = SS().getSheetByName("Lista"), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 7).getValues().map((r, i) => ({
    row: i + 2, id: String(r[0]), f: fmt(r[1], "dd/MM HH:mm"), kid: String(r[2]), n: String(r[3]).replace(/^'/, ""),
    tel: String(r[4]).replace(/^'/, ""), ing: fmt(r[5], "HH:mm"), nota: String(r[6] || "")
  })).filter(x => x.id);
}
function pagos() {
  const sh = SS().getSheetByName("Pagos misioneros"), n = sh.getLastRow();
  if (n < 2) return [];
  const monto = v => typeof v === "number" ? v : (Number(String(v == null ? "" : v).replace(/[^\d-]/g, "")) || 0);
  return sh.getRange(2, 1, n - 1, 5).getValues().filter(r => String(r[1]).trim() !== "" && monto(r[2]) !== 0).map(r => ({
    f: fmt(r[0], "dd/MM HH:mm"), kid: String(r[1]).trim(), m: monto(r[2]), medio: String(r[3] || ""), nota: String(r[4] || "")
  }));
}
function nuevoId(L) { const s = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789", usados = {}; L.forEach(x => usados[x.id] = 1); let id; do { id = "L-"; for (let i = 0; i < 5; i++) id += s[Math.floor(Math.random() * s.length)]; } while (usados[id]); return id; }

/* ---------- datos públicos (para elegir el nombre) ---------- */
function datos() {
  const c = cfg(), L = lista(), cnt = {};
  L.forEach(x => cnt[x.kid] = (cnt[x.kid] || 0) + 1);
  return { ok: true, v: 2, config: { price: c.price, per: c.per, date: c.date, where: c.where, wa: c.wa, mail: c.mail, pagoTxt: c.pagoTxt, clave: !!c.claveM },
    chicos: chicos(c).map(x => ({ n: x.n, c: x.c, k: cnt[x.n] || 0 })) };
}

/* ---------- acciones ---------- */
function accion(b) {
  const c = cfg();
  if (b.action === "admin") return admin(b, c);
  /* acciones de misioneros */
  const kid = String(b.kid || "").trim(), yo = chicos(c).filter(x => x.n === kid)[0];
  if (!yo) return { ok: false, error: "No encontramos tu nombre en la lista de misioneros. Avisale a un coordinador." };
  if (c.claveM && String(b.clave || "").trim() !== c.claveM) return { ok: false, clave: true, error: "La clave de misioneros no es correcta." };
  if (b.action === "mis") return mis(yo, c);
  if (b.action === "cargar") return conLock(() => {
    const name = clean(b.name, 70), tel = clean(b.tel, 25), L = lista(), mias = L.filter(x => x.kid === kid);
    if (!/\S{2,}\s+\S{2,}/.test(name.replace(/^'/, ""))) return { ok: false, error: "Escribí nombre y apellido de la persona." };
    if (tel && tel.replace(/\D/g, "").length < 8) return { ok: false, error: "El teléfono parece incompleto (o dejalo vacío)." };
    if (mias.length >= yo.c) return { ok: false, lleno: true, error: "Ya cargaste tus " + yo.c + " entradas. Si vendiste más, pedile a un coordinador que te suba el cupo." };
    const d = L.filter(x => norm(x.n) === norm(name))[0];
    if (d && !b.force) return { ok: false, dup: { n: d.n, kid: d.kid === kid ? "vos" : d.kid } };
    SS().getSheetByName("Lista").appendRow([nuevoId(L), new Date(), kid, name, tel ? "'" + tel.replace(/^'/, "") : "", "", ""]);
    return mis(yo, c);
  });
  if (b.action === "borrar") return conLock(() => {
    const x = lista().filter(r => r.id === String(b.id) && r.kid === kid)[0];
    if (!x) return { ok: false, error: "No encontramos esa entrada (¿ya la borraste?)." };
    if (x.ing) return { ok: false, error: "Esa persona ya entró a la peña: no se puede borrar." };
    SS().getSheetByName("Lista").deleteRow(x.row);
    return mis(yo, c);
  });
  return { ok: false, error: "Acción desconocida." };
}
function mis(yo, c) {
  const items = lista().filter(x => x.kid === yo.n).map(x => ({ id: x.id, n: x.n, tel: x.tel, ing: x.ing }));
  const pag = pagos().filter(p => p.kid === yo.n).reduce((s, p) => s + p.m, 0);
  return { ok: true, kid: yo.n, cupo: yo.c, items, pagado: pag, debe: Math.max(yo.c, items.length) * c.price, price: c.price, pagoTxt: c.pagoTxt };
}

/* ---------- administración ---------- */
function admin(b, c) {
  if (!c.claveA || String(b.clave || "") !== c.claveA) return { ok: false, error: "Clave incorrecta." };
  const op = b.op;
  if (op === "login") return { ok: true };
  if (op === "todo") return todo(c);
  if (op === "ingreso") return conLock(() => {
    const x = lista().filter(r => r.id === String(b.id))[0];
    if (!x) return { ok: false, error: "No encontramos a esa persona en la lista (¿la borraron?)." };
    const sh = SS().getSheetByName("Lista");
    if (b.undo) { sh.getRange(x.row, 6).setValue(""); return { ok: true, ing: "" }; }
    if (x.ing) return { ok: false, ya: x.ing, error: x.n + " ya había entrado a las " + x.ing + "." };
    const d = new Date(); sh.getRange(x.row, 6).setValue(d); return { ok: true, ing: fmt(d, "HH:mm") };
  });
  if (op === "puerta") return conLock(() => {
    const name = clean(b.name, 70);
    if (name.replace(/^'/, "").length < 3) return { ok: false, error: "Escribí el nombre de la persona." };
    const L = lista(), d = new Date(), id = nuevoId(L);
    SS().getSheetByName("Lista").appendRow([id, d, PUERTA, name, "", d, "Pagó en la puerta"]);
    return { ok: true, id, ing: fmt(d, "HH:mm") };
  });
  if (op === "borrar") return conLock(() => {
    const x = lista().filter(r => r.id === String(b.id))[0];
    if (!x) return { ok: false, error: "No encontramos esa entrada." };
    SS().getSheetByName("Lista").deleteRow(x.row); return { ok: true };
  });
  if (op === "pago") return conLock(() => {
    const kid = String(b.kid || "").trim(), m = Math.round(Number(b.monto) || 0);
    if (!kid) return { ok: false, error: "Falta el misionero." };
    if (!m || Math.abs(m) > 10000000) return { ok: false, error: "Escribí un monto válido." };
    SS().getSheetByName("Pagos misioneros").appendRow([new Date(), kid, m, b.medio === "Transferencia" ? "Transferencia" : "Efectivo", clean(b.nota, 120)]);
    return { ok: true };
  });
  return { ok: false, error: "Operación desconocida." };
}
function todo(c) {
  return { ok: true, price: c.price, per: c.per, cap: c.cap, puerta: PUERTA,
    lista: lista().map(x => ({ id: x.id, f: x.f, kid: x.kid, n: x.n, tel: x.tel, ing: x.ing })),
    chicos: chicos(c).map(x => ({ n: x.n, c: x.c })), pagos: pagos() };
}

/* ---------- preparación de la planilla (corre una sola vez) ---------- */
function preparar() {
  if (SS().getSheetByName("Lista")) return;
  conLock(() => {
    const ss = SS();
    if (ss.getSheetByName("Lista")) return;
    /* Pagos misioneros */
    let p = ss.getSheetByName("Pagos misioneros");
    if (!p) { p = ss.insertSheet("Pagos misioneros"); p.getRange(1, 1, 1, 5).setValues([H_PAGOS]).setFontWeight("bold"); p.setFrozenRows(1);
      p.getRange("A2:A2000").setNumberFormat("dd/MM/yyyy HH:mm"); p.getRange("C2:C2000").setNumberFormat("$#,##0"); p.setColumnWidth(2, 220); p.setColumnWidth(5, 260); }
    /* Config */
    const cf = ss.getSheetByName("Config"), n = cf.getLastRow(), keys = cf.getRange(1, 1, Math.max(n, 1), 1).getValues().map(r => String(r[0]).trim());
    const fila = k => keys.indexOf(k) + 1;
    if (fila("Entradas por chico")) cf.getRange(fila("Entradas por chico"), 2, 1, 2).setValues([[4, "Entradas que vende cada misionero. Las paga aunque no las venda. Para un cupo distinto: pestaña 'Extras y cupos'."]]);
    if (fila("Precio de la entrada ($)")) cf.getRange(fila("Precio de la entrada ($)"), 3).setValue("Precio por persona. Lo cobra el misionero y se lo entrega a los coordinadores.");
    const extra = [];
    if (!fila("Clave de misioneros")) extra.push(["Clave de misioneros", "", "Opcional. Si escribís una clave, los misioneros la tienen que poner para cargar entradas (pasala por #ACHEPA). Vacío = sin clave."]);
    if (!fila("Cómo pagan los misioneros")) extra.push(["Cómo pagan los misioneros", PAGO_TXT, "Texto que ve cada misionero debajo de lo que tiene que pagar."]);
    if (extra.length) cf.getRange(cf.getLastRow() + 2, 1, extra.length, 3).setValues(extra);
    /* Chicos */
    const ch = ss.getSheetByName("Chicos");
    ch.getRange(1, 3, 1, 6).setValues([["Cargadas", "Le faltan cargar", "Tiene que pagar ($)", "Pagó ($)", "Saldo ($)",
      "No editar esta pestaña: se completa sola. Bajas: borrar la fila en Asistencia. Altas o cupos especiales: 'Extras y cupos'. Pagos: 'Pagos misioneros'."]]);
    ch.getRange(2, 3, 1, 5).setFormulas([[
      '=ARRAYFORMULA(IF(A2:A400="","",COUNTIF(Lista!C2:C3000,A2:A400)))',
      '=ARRAYFORMULA(IF(A2:A400="","",IF(B2:B400-C2:C400>0,B2:B400-C2:C400,0)))',
      '=ARRAYFORMULA(IF(A2:A400="","",IF(C2:C400>B2:B400,C2:C400,B2:B400)*Config!B2))',
      "=ARRAYFORMULA(IF(A2:A400=\"\",\"\",SUMIF('Pagos misioneros'!B2:B3000,A2:A400,'Pagos misioneros'!C2:C3000)))",
      '=ARRAYFORMULA(IF(A2:A400="","",E2:E400-F2:F400))']]);
    ch.getRange("E2:G400").setNumberFormat("$#,##0");
    /* Lista (se crea al final: su existencia marca que la preparación ya se hizo) */
    const li = ss.insertSheet("Lista", 0);
    li.getRange(1, 1, 1, 7).setValues([H_LISTA]).setFontWeight("bold"); li.setFrozenRows(1);
    li.getRange("B2:B3000").setNumberFormat("dd/MM HH:mm"); li.getRange("E2:E3000").setNumberFormat("@"); li.getRange("F2:F3000").setNumberFormat("HH:mm");
    li.setColumnWidth(3, 200); li.setColumnWidth(4, 240); li.setColumnWidth(7, 200);
    /* Resumen */
    const rs = ss.getSheetByName("Resumen");
    if (rs) { rs.clear(); rs.getRange(1, 1, 15, 2).setValues([
      ["Indicador", "Valor"],
      ["Entradas en la lista", "=COUNTA(Lista!A2:A3000)"],
      ["  cargadas por misioneros", '=B2-B4'],
      ["  vendidas en la puerta", '=COUNTIF(Lista!C2:C3000,"' + PUERTA + '")'],
      ["Ya entraron", "=COUNTA(Lista!F2:F3000)"],
      ["Faltan entrar", "=B2-B5"],
      ["Misioneros que todavía no cargaron todas", '=COUNTIF(Chicos!D2:D400,">0")'],
      ["Entradas sin cargar (de los cupos)", "=SUM(Chicos!D2:D400)"],
      ["Plata que tienen que entregar los misioneros ($)", "=SUM(Chicos!E2:E400)"],
      ["Ya entregaron ($)", "=SUM('Pagos misioneros'!C2:C3000)"],
      ["  en efectivo ($)", "=SUMIF('Pagos misioneros'!D2:D3000,\"Efectivo\",'Pagos misioneros'!C2:C3000)"],
      ["  por transferencia ($)", "=SUMIF('Pagos misioneros'!D2:D3000,\"Transferencia\",'Pagos misioneros'!C2:C3000)"],
      ["Falta que entreguen ($)", "=B9-B10"],
      ["Ventas en la puerta ($)", "=B4*Config!B2"],
      ["TOTAL ENTRADAS COBRADO ($)", "=B10+B14"]]);
      rs.getRange(1, 1, 1, 2).setFontWeight("bold"); rs.getRange("B9:B15").setNumberFormat("$#,##0"); rs.setColumnWidth(1, 340); }
    /* Léeme */
    const le = ss.getSheetByName("Léeme");
    if (le) { le.clear(); const t = [
      ["Peña · Parroquia Espíritu Santo — Lista de entradas (sistema nuevo, nov 2026)"],
      ["Cada misionero vende sus entradas, cobra la plata y carga en la página el nombre y apellido de cada persona. El día de la peña, en la puerta se busca el nombre en la lista."],
      [""],
      ["QUÉ HAY EN CADA PESTAÑA"],
      ["Lista: una fila por persona que va a la peña. La página agrega y borra las filas sola. 'Ingresó (hora)' se completa en la puerta."],
      ["Pagos misioneros: cada vez que un misionero entrega plata, se registra acá (desde #admin > Misioneros, o a mano: fecha, nombre EXACTO, monto, Efectivo/Transferencia)."],
      ["Chicos: lista de misioneros (sale sola de la planilla maestra) con cuántas cargó cada uno, cuánto tiene que pagar, cuánto pagó y el saldo. No editar."],
      ["Extras y cupos: misioneros que no están en la planilla maestra, o cupo distinto de 4 para alguien."],
      ["Config: precio, entradas por chico, fecha, lugar, claves y texto de pago."],
      ["Resumen: totales automáticos."],
      ["Entradas (sistema viejo), Bandas, Menú: del sistema anterior. La página nueva no los usa."],
      [""],
      ["REGLAS"],
      ["Cada misionero paga sus 4 entradas aunque no las venda. Si vendió más, primero que le compren a otro misionero; si no, se le sube el cupo en 'Extras y cupos'."],
      ["No cambies los nombres de las pestañas ni los títulos de las columnas, y no insertes ni muevas columnas."],
      ["Página: https://misionespiritusanto.github.io/Pena/ · Administración: .../Pena/#admin con la 'Clave de administración' de Config."]];
      le.getRange(1, 1, t.length, 1).setValues(t); le.getRange(1, 1).setFontWeight("bold"); }
    /* La pestaña vieja queda guardada con otro nombre */
    const en = ss.getSheetByName("Entradas");
    if (en) en.setName("Entradas (sistema viejo)");
  });
}
