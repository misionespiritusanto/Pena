/* ===== Peña · Grupo Misionero Espíritu Santo — lista de entradas (v3, 15/11/2026) =====
   Cómo instalarlo (una sola vez):
   1. En la planilla Pena_Entradas: Extensiones > Apps Script. Borrar todo y pegar este archivo entero. Guardar.
   2. Arriba, en la lista de funciones, elegir "instalar" y tocar Ejecutar (si pide permisos, aceptarlos).
      Eso deja la planilla limpia: Léeme, Lista, Misioneros, Resumen y Config.
   3. Implementar > Gestionar implementaciones > lápiz > Versión: "Nueva versión" > Implementar.
      El link de la web app no cambia.

   Pestañas:
   - Lista: una fila por persona que va a la peña (la llena la página).
   - Misioneros: un renglón por misionero. Los nombres se agregan solos desde 'Asistencia ' de la
     planilla maestra. Los coordinadores escriben "Cupo" y "Pagó ($)" (lo que entregó hasta ahora).
   - Resumen: totales. Config: precio, fecha, claves. */

const TZ = "America/Argentina/Buenos_Aires";
const MAESTRA = "1ff9dWtfo3Pg8UeySRfsGUJ7FyAz53AEWu0VNmQFzmWY"; /* ¡San Pedro 2027! */
const ASISTENCIA = "Asistencia ";
const PUERTA = "Venta en puerta";
const PAGO_TXT = "Entregale la plata a Nacho: en efectivo en los encuentros de los domingos o como te digan los coordinadores.";
const AMARILLO = "#fff2cc";

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
function conLock(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

function cfg() {
  const sh = SS().getSheetByName("Config"), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (r[0] !== "") m[String(r[0]).trim()] = r[1]; });
  const s = k => String(m[k] == null ? "" : m[k]).trim();
  return {
    price: num(m["Precio de la entrada ($)"]) || 15000,
    per: num(m["Entradas por misionero"]) || num(m["Entradas por chico"]) || 4,
    date: s("Fecha y hora"), where: s("Lugar"), cap: num(m["Capacidad del salón"]) || 0,
    pagoTxt: s("Cómo pagan los misioneros") || PAGO_TXT,
    wa: s("WhatsApp de contacto").replace(/\D/g, ""), mail: s("Mail de contacto"),
    claveM: s("Clave de misioneros"), claveA: s("Clave de administración")
  };
}

/* Misioneros: A Misionero · B Cupo · C Cargadas (fórmula) · D Pagó ($) · E Falta pagar (fórmula) · F Estado */
function misioneros(c) {
  const sh = SS().getSheetByName("Misioneros"), n = sh.getLastRow(), L = [];
  if (n < 2) return L;
  sh.getRange(2, 1, n - 1, 6).getValues().forEach((r, i) => {
    const k = String(r[0]).trim();
    if (k) L.push({ row: i + 2, n: k, c: r[1] === "" ? c.per : num(r[1]), pag: plata(r[3]), est: String(r[5] || "") });
  });
  return L;
}
/* Lista: A Nombre y apellido · B Misionero · C Teléfono · D Entró · E Cargada el · F ID */
function lista() {
  const sh = SS().getSheetByName("Lista"), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 6).getValues().map((r, i) => ({
    row: i + 2, n: sinComilla(r[0]), kid: String(r[1]).trim(), tel: sinComilla(r[2]), ing: fmt(r[3], "HH:mm"), f: fmt(r[4], "dd/MM HH:mm"), id: String(r[5])
  })).filter(x => x.id && x.n);
}
function nuevoId(L) { const s = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789", usados = {}; L.forEach(x => usados[x.id] = 1); let id; do { id = "L-"; for (let i = 0; i < 5; i++) id += s[Math.floor(Math.random() * s.length)]; } while (usados[id]); return id; }

/* ---------- datos para elegir el nombre ---------- */
function datos() {
  const c = cfg(), L = lista(), cnt = {};
  L.forEach(x => cnt[x.kid] = (cnt[x.kid] || 0) + 1);
  return { ok: true, v: 3, config: { price: c.price, per: c.per, date: c.date, where: c.where, wa: c.wa, mail: c.mail, pagoTxt: c.pagoTxt, clave: !!c.claveM },
    chicos: misioneros(c).map(x => ({ n: x.n, c: x.c, k: cnt[x.n] || 0 })) };
}

/* ---------- acciones ---------- */
function accion(b) {
  const c = cfg();
  if (b.action === "admin") return admin(b, c);
  const kid = String(b.kid || "").trim(), yo = misioneros(c).filter(x => x.n === kid)[0];
  if (!yo) return { ok: false, error: "No encontramos tu nombre en la lista de misioneros. Avisale a un coordinador." };
  if (c.claveM && String(b.clave || "").trim() !== c.claveM) return { ok: false, clave: true, error: "La clave de misioneros no es correcta." };
  if (b.action === "mis") return mis(yo, c);
  if (b.action === "cargar") return conLock(() => {
    const name = clean(b.name, 70), tel = clean(b.tel, 25), L = lista(), mias = L.filter(x => x.kid === kid);
    if (!/\S{2,}\s+\S{2,}/.test(sinComilla(name))) return { ok: false, error: "Escribí nombre y apellido de la persona." };
    if (tel && tel.replace(/\D/g, "").length < 8) return { ok: false, error: "El teléfono parece incompleto (o dejalo vacío)." };
    if (mias.length >= yo.c) return { ok: false, lleno: true, error: "Ya cargaste tus " + yo.c + " entradas. Si vendiste más, pedile a un coordinador que te suba el cupo." };
    const d = L.filter(x => norm(x.n) === norm(sinComilla(name)))[0];
    if (d && !b.force) return { ok: false, dup: { n: d.n, kid: d.kid === kid ? "vos" : d.kid } };
    SS().getSheetByName("Lista").appendRow([name, kid, tel ? "'" + sinComilla(tel) : "", "", new Date(), nuevoId(L)]);
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
  return { ok: true, kid: yo.n, cupo: yo.c, items, pagado: yo.pag, debe: yo.c * c.price, price: c.price, pagoTxt: c.pagoTxt };
}

/* ---------- administración ---------- */
function admin(b, c) {
  if (!c.claveA || String(b.clave || "") !== c.claveA) return { ok: false, error: "Clave incorrecta." };
  const op = b.op;
  if (op === "login") return { ok: true };
  if (op === "todo") { sincronizar(false); return { ok: true, price: c.price, per: c.per, cap: c.cap, puerta: PUERTA,
    lista: lista().map(x => ({ id: x.id, f: x.f, kid: x.kid, n: x.n, tel: x.tel, ing: x.ing })),
    chicos: misioneros(c).map(x => ({ n: x.n, c: x.c, pag: x.pag, est: x.est })) }; }
  if (op === "ingreso") return conLock(() => {
    const x = lista().filter(r => r.id === String(b.id))[0];
    if (!x) return { ok: false, error: "No encontramos a esa persona en la lista (¿la borraron?)." };
    const sh = SS().getSheetByName("Lista");
    if (b.undo) { sh.getRange(x.row, 4).setValue(""); return { ok: true, ing: "" }; }
    if (x.ing) return { ok: false, ya: x.ing, error: x.n + " ya había entrado a las " + x.ing + "." };
    const d = new Date(); sh.getRange(x.row, 4).setValue(d); return { ok: true, ing: fmt(d, "HH:mm") };
  });
  if (op === "puerta") return conLock(() => {
    const name = clean(b.name, 70);
    if (sinComilla(name).length < 3) return { ok: false, error: "Escribí el nombre de la persona." };
    const L = lista(), d = new Date(), id = nuevoId(L);
    SS().getSheetByName("Lista").appendRow([name, PUERTA, "", d, d, id]);
    return { ok: true, id, ing: fmt(d, "HH:mm") };
  });
  if (op === "borrar") return conLock(() => {
    const x = lista().filter(r => r.id === String(b.id))[0];
    if (!x) return { ok: false, error: "No encontramos esa entrada." };
    SS().getSheetByName("Lista").deleteRow(x.row); return { ok: true };
  });
  if (op === "pagado") return conLock(() => {
    const x = misioneros(c).filter(m => m.n === String(b.kid || "").trim())[0], t = Math.round(Number(b.total));
    if (!x) return { ok: false, error: "No encontramos a ese misionero." };
    if (!(t >= 0) || t > 10000000) return { ok: false, error: "Escribí un monto válido." };
    SS().getSheetByName("Misioneros").getRange(x.row, 4).setValue(t || "");
    return { ok: true, pag: t };
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
    const filas = n >= 2 ? sh.getRange(2, 1, n - 1, 6).getValues().filter(r => String(r[0]).trim()) : [];
    const cnt = {}; lista().forEach(x => cnt[x.kid] = (cnt[x.kid] || 0) + 1);
    const enMaestra = {}; maestra.forEach(k => enMaestra[k] = 1);
    const tengo = {}; let cambio = false;
    const nuevas = [];
    filas.forEach(r => {
      const k = String(r[0]).trim(); tengo[k] = 1;
      let est = String(r[5] || "");
      if (enMaestra[k]) est = "";
      else if (auto[k]) {
        if (!cnt[k] && (r[3] === "" || plata(r[3]) === 0)) { delete auto[k]; cambio = true; return; } /* baja sin datos: se va */
        est = "Ya no está en Asistencia";
      }
      if (est !== String(r[5] || "")) cambio = true;
      nuevas.push([k, r[1], r[3], est]);
    });
    maestra.forEach(k => { if (!tengo[k]) { nuevas.push([k, c.per, "", ""]); auto[k] = 1; cambio = true; } });
    const orden = nuevas.slice().sort((a, b) => a[0].localeCompare(b[0], "es"));
    if (!cambio && orden.some((r, i) => r[0] !== String(filas[i] && filas[i][0]).trim())) cambio = true;
    if (cambio) {
      const m = orden.length, viejo = Math.max(0, n - 1);
      if (m) {
        sh.getRange(2, 1, m, 2).setValues(orden.map(r => [r[0], r[1]]));
        sh.getRange(2, 4, m, 1).setValues(orden.map(r => [r[2]]));
        sh.getRange(2, 6, m, 1).setValues(orden.map(r => [r[3]]));
      }
      if (viejo > m) { sh.getRange(m + 2, 1, viejo - m, 2).clearContent(); sh.getRange(m + 2, 4, viejo - m, 1).clearContent(); sh.getRange(m + 2, 6, viejo - m, 1).clearContent(); }
      props.setProperty("auto", JSON.stringify(auto));
    }
  });
  cache.put("sync", "1", 300);
}

/* ---------- instalación: deja la planilla limpia (corre una sola vez) ---------- */
function instalar() { preparar(true); sincronizar(true); }
function preparar(forzar) {
  if (!forzar && SS().getSheetByName("Misioneros")) return;
  conLock(() => {
    const ss = SS();
    if (!forzar && ss.getSheetByName("Misioneros")) return;
    if (ss.getSheetByName("Misioneros")) return; /* ya instalado */
    /* 1. Guardar lo que sirva de antes */
    const viejo = {}; const cf = ss.getSheetByName("Config");
    if (cf && cf.getLastRow() >= 2) cf.getRange(2, 1, cf.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== "") viejo[String(r[0]).trim()] = r[1]; });
    const filasLista = [], li = ss.getSheetByName("Lista");
    if (li && li.getLastRow() >= 2) li.getRange(2, 1, li.getLastRow() - 1, 7).getValues().forEach(r => { if (r[0] && r[3]) filasLista.push([r[3], r[2], r[4], r[5], r[1], r[0]]); }); /* v2: ID,Cargada,Misionero,Nombre,Tel,Ingresó */
    const pagos = {}, pg = ss.getSheetByName("Pagos misioneros");
    if (pg && pg.getLastRow() >= 2) pg.getRange(2, 1, pg.getLastRow() - 1, 3).getValues().forEach(r => { const k = String(r[1]).trim(); if (k) pagos[k] = (pagos[k] || 0) + plata(r[2]); });
    const extras = [], ex = ss.getSheetByName("Extras y cupos");
    if (ex && ex.getLastRow() >= 2) ex.getRange(2, 1, ex.getLastRow() - 1, 2).getValues().forEach(r => { const k = String(r[0]).trim(); if (k) extras.push([k, r[1]]); });
    const v = k => viejo[k] == null ? "" : viejo[k];

    /* 2. Hojas nuevas (con nombre provisorio) */
    const sfx = "_" + Date.now();
    const tmp = ss.insertSheet("_nueva_lista" + sfx), mi = ss.insertSheet("_nueva_misioneros" + sfx), rs = ss.insertSheet("_nuevo_resumen" + sfx), le = ss.insertSheet("_nuevo_leeme" + sfx), co = ss.insertSheet("_nueva_config" + sfx);

    /* Lista */
    tmp.getRange(1, 1, 1, 6).setValues([["Nombre y apellido", "Misionero", "Teléfono", "Entró (hora)", "Cargada el", "ID"]]).setFontWeight("bold").setBackground("#553a1e").setFontColor("#ffffff");
    tmp.setFrozenRows(1); tmp.getRange("C2:C3000").setNumberFormat("@"); tmp.getRange("D2:D3000").setNumberFormat("HH:mm"); tmp.getRange("E2:E3000").setNumberFormat("dd/MM HH:mm");
    tmp.getRange("F1:F3000").setFontColor("#999999");
    tmp.setColumnWidth(1, 240); tmp.setColumnWidth(2, 220); tmp.setColumnWidth(3, 130); tmp.setColumnWidth(4, 100); tmp.setColumnWidth(5, 110); tmp.setColumnWidth(6, 80);
    if (filasLista.length) tmp.getRange(2, 1, filasLista.length, 6).setValues(filasLista);

    /* Misioneros */
    mi.getRange(1, 1, 1, 6).setValues([["Misionero", "Cupo", "", "Pagó ($)", "", "Estado"]]);
    mi.getRange(1, 3).setFormula('={"Cargadas";ARRAYFORMULA(IF(A2:A="","",COUNTIF(Lista!B2:B,A2:A)))}');
    mi.getRange(1, 5).setFormula('={"Falta pagar ($)";ARRAYFORMULA(IF(A2:A="","",B2:B*Config!B2-IF(D2:D="",0,D2:D)))}');
    mi.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#553a1e").setFontColor("#ffffff");
    mi.setFrozenRows(1); mi.getRange("D2:E400").setNumberFormat("$#,##0"); mi.getRange("B2:B400").setBackground(AMARILLO); mi.getRange("D2:D400").setBackground(AMARILLO);
    mi.setColumnWidth(1, 240); mi.setColumnWidth(6, 200);
    if (extras.length) mi.getRange(2, 1, extras.length, 2).setValues(extras.map(r => [r[0], r[1] === "" ? (num(v("Entradas por chico")) || 4) : r[1]]));
    Object.keys(pagos).forEach(k => { const i = extras.findIndex(r => r[0] === k); if (i < 0) { extras.push([k, ""]); mi.getRange(extras.length + 1, 1, 1, 2).setValues([[k, 4]]); mi.getRange(extras.length + 1, 4).setValue(pagos[k]); } else mi.getRange(i + 2, 4).setValue(pagos[k]); });

    /* Config */
    const conf = [
      ["Parámetro", "Valor", "Para qué sirve"],
      ["Precio de la entrada ($)", num(v("Precio de la entrada ($)")) || 15000, "Precio por persona. Lo cobra el misionero y se lo entrega a los coordinadores."],
      ["Entradas por misionero", 4, "Cupo que se le pone a cada misionero nuevo en la pestaña Misioneros (ahí se puede cambiar el de cada uno)."],
      ["Fecha y hora", v("Fecha y hora") || "Sábado 28/11 · horario a confirmar", "Se muestra en la página."],
      ["Lugar", v("Lugar") || "Parroquia Espíritu Santo · Juan Clark 774", "Se muestra en la página."],
      ["Capacidad del salón", v("Capacidad máxima del salón"), "Opcional. Si la cargás, el Resumen muestra los lugares libres."],
      ["Cómo pagan los misioneros", v("Cómo pagan los misioneros") || PAGO_TXT, "Texto que ve cada misionero debajo de lo que tiene que entregar."],
      ["WhatsApp de contacto", v("WhatsApp de contacto") || "5491170544319", "Número para dudas (solo dígitos, con 549)."],
      ["Mail de contacto", v("Mail de contacto") || "ignaciozavalia1@gmail.com", "Mail para consultas sobre datos (privacidad)."],
      ["Clave de misioneros", v("Clave de misioneros"), "Opcional. Si escribís una, los misioneros la tienen que poner para cargar (pasala por #ACHEPA). Vacío = sin clave."],
      ["Clave de administración", v("Clave de administración") || "paguenme", "Clave de la pantalla de coordinadores (.../Pena/#admin). Cambiala y no la compartas."]];
    co.getRange(1, 1, conf.length, 3).setValues(conf); co.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#553a1e").setFontColor("#ffffff");
    co.getRange(2, 2, conf.length - 1, 1).setBackground(AMARILLO); co.getRange("B2").setNumberFormat("$#,##0"); co.setFrozenRows(1);
    co.setColumnWidth(1, 220); co.setColumnWidth(2, 330); co.setColumnWidth(3, 520);

    /* Resumen */
    const res = [
      ["Entradas", ""],
      ["Personas en la lista", "=COUNTA(Lista!A2:A)"],
      ["   cargadas por misioneros", "=B2-B4"],
      ["   vendidas en la puerta", '=COUNTIF(Lista!B2:B,"' + PUERTA + '")'],
      ["Entradas que faltan cargar (de los cupos)", "=SUM(Misioneros!B2:B)-B3"],
      ["Ya entraron a la peña", "=COUNTA(Lista!D2:D)"],
      ["Lugares libres en el salón", '=IF(Config!B6="","",Config!B6-B2)'],
      ["", ""],
      ["Plata", ""],
      ["Tienen que entregar los misioneros", "=SUM(Misioneros!B2:B)*Config!B2"],
      ["Ya entregaron", "=SUM(Misioneros!D2:D)"],
      ["Falta que entreguen", "=B10-B11"],
      ["Ventas en la puerta", "=B4*Config!B2"],
      ["TOTAL COBRADO EN ENTRADAS", "=B11+B13"]];
    rs.getRange(1, 1, res.length, 2).setValues(res);
    rs.getRange("A1").setFontWeight("bold"); rs.getRange("A9").setFontWeight("bold"); rs.getRange("A14:B14").setFontWeight("bold");
    rs.getRange("B10:B14").setNumberFormat("$#,##0"); rs.setColumnWidth(1, 330); rs.setColumnWidth(2, 140);

    /* Léeme */
    const t = [
      ["Peña · Grupo Misionero Espíritu Santo — Lista de entradas"],
      ["Cada misionero vende " + 4 + " entradas, cobra la plata y carga en la página el nombre y apellido de cada persona. El día de la peña, en la puerta se busca el nombre en la lista."],
      [""],
      ["Lista → todas las personas que van a la peña. La llena la página (no hace falta tocarla). 'Entró' se completa en la puerta."],
      ["Misioneros → un renglón por misionero. Los nombres se agregan solos desde la planilla maestra (Asistencia)."],
      ["      Columnas amarillas = las completan los coordinadores:"],
      ["      · Pagó ($): lo que entregó el misionero HASTA AHORA (el total, no lo de hoy). Ejemplo: pagó 30.000 y después 15.000 → escribir 45000."],
      ["      · Cupo: cuántas entradas vende (4). Cambiarlo solo si alguien vende más o menos."],
      ["      Lo demás se calcula solo. El misionero ve en la página 'Entregaste $X de $60.000' con lo que ustedes escriben acá."],
      ["      ¿Alguien que no está en Asistencia? Escribir su nombre en una fila vacía al final, con su cupo."],
      ["Resumen → totales de entradas y plata."],
      ["Config → precio, fecha, lugar y claves (celdas amarillas)."],
      [""],
      ["Página para misioneros: https://misionespiritusanto.github.io/Pena/"],
      ["Página de coordinadores (puerta, pagos, resumen): https://misionespiritusanto.github.io/Pena/#admin  (clave: 'Clave de administración' de Config)"],
      ["No cambiar nombres de pestañas ni títulos de columnas."]];
    le.getRange(1, 1, t.length, 1).setValues(t); le.getRange("A1").setFontWeight("bold").setFontSize(14); le.setColumnWidth(1, 900);

    /* 3. Borrar todo lo viejo y dejar los nombres definitivos */
    /* se borra TODA hoja que no sea una de las 5 nuevas (incluye restos de un intento anterior) */
    const nuevas = [le, tmp, mi, rs, co].map(s => s.getSheetId());
    ss.getSheets().forEach(s => { if (nuevas.indexOf(s.getSheetId()) < 0) ss.deleteSheet(s); });
    SpreadsheetApp.flush();
    le.setName("Léeme"); tmp.setName("Lista"); mi.setName("Misioneros"); rs.setName("Resumen"); co.setName("Config");
    /* las fórmulas apuntan a Lista/Config por nombre: se reescriben ahora que existen */
    mi.getRange(1, 3).setFormula('={"Cargadas";ARRAYFORMULA(IF(A2:A="","",COUNTIF(Lista!B2:B,A2:A)))}');
    mi.getRange(1, 5).setFormula('={"Falta pagar ($)";ARRAYFORMULA(IF(A2:A="","",B2:B*Config!B2-IF(D2:D="",0,D2:D)))}');
    rs.getRange(1, 1, res.length, 2).setValues(res);
    [le, tmp, mi, rs, co].forEach((s, i) => { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); });
    ss.setActiveSheet(le);
    PropertiesService.getScriptProperties().deleteProperty("auto");
  });
  CacheService.getScriptCache().remove("sync");
}
