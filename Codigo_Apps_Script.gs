/**
 * Encuesta Final del Rally · Grupo 19 · Equipo 4
 * Guarda las respuestas en la hoja de cálculo de Google, en una pestaña nueva ("Respuestas Finales").
 * La pestaña de la primera encuesta no se toca.
 *
 * Cómo usarlo: pega este código en el proyecto de Apps Script de la primera encuesta (reemplaza el
 * código anterior), escribe en CLAVE la misma contraseña de resultados que ya usabas y entra a
 * Implementar → Administrar implementaciones → Editar → Versión nueva → Implementar.
 * Así la dirección /exec no cambia.
 */
var CLAVE = "CAMBIA-ESTA-CLAVE";
var ID_HOJA = "";                 // solo si el script NO está dentro de la hoja: pega aquí el ID de la hoja de cálculo
var SAL = "rally-g19-e4";          // texto fijo para proteger los correos; no hace falta cambiarlo
var HOJA = "Respuestas Finales";

function hoja_() {
  var ss = ID_HOJA ? SpreadsheetApp.openById(ID_HOJA) : SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName(HOJA);
  if (!h) {
    h = ss.insertSheet(HOJA);
    h.appendRow(["Fecha", "Huella del correo", "Respuestas (JSON)"]);
    h.setFrozenRows(1);
  }
  return h;
}

function huella_(correo) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SAL + "|" + correo);
  return bytes.map(function (b) { return ("0" + (b & 255).toString(16)).slice(-2); }).join("");
}

function entero_(x, a, b) { return typeof x === "number" && x === Math.floor(x) && x >= a && x <= b; }

function limpiar_(d) {
  if (!d || typeof d !== "object") return null;
  if (!Array.isArray(d.e) || d.e.length !== 6) return null;
  for (var i = 0; i < 6; i++) if (!entero_(d.e[i], 1, 10)) return null;
  if (!entero_(d.s, 1, 10) || !entero_(d.o, 1, 10) || !entero_(d.h, 0, 2) || !entero_(d.r, 0, 4)) return null;
  if (!Array.isArray(d.ob) || !d.ob.length || d.ob.length > 5) return null;
  if (!Array.isArray(d.ac) || !d.ac.length || d.ac.length > 4) return null;
  for (var j = 0; j < d.ob.length; j++) if (!entero_(d.ob[j], 0, 4)) return null;
  for (var k = 0; k < d.ac.length; k++) if (!entero_(d.ac[k], 0, 3)) return null;
  return {
    e: d.e.slice(), s: d.s, o: d.o, h: d.h, r: d.r,
    m: String(d.m || "").slice(0, 400), x: String(d.x || "").slice(0, 400),
    ob: d.ob.slice(), ac: d.ac.slice()
  };
}

function correo_(c) {
  c = String(c || "").trim().toLowerCase();
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(c) && c.length <= 120 ? c : "";
}

function yaRespondio_(h, huella) {
  var n = h.getLastRow();
  if (n < 2) return false;
  var col = h.getRange(2, 2, n - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) if (col[i][0] === huella) return true;
  return false;
}

function enviar_(doc) {
  var c = correo_(doc && doc.correo), d = limpiar_(doc && doc.datos);
  if (!c || !d) return { error: "datos" };
  var h = hoja_(), hu = huella_(c);
  if (yaRespondio_(h, hu)) return { error: "duplicado" };
  h.appendRow([new Date(), hu, JSON.stringify(d)]);
  return { ok: true };
}

function existe_(correo) {
  var c = correo_(correo);
  if (!c) return { error: "datos" };
  return { ok: yaRespondio_(hoja_(), huella_(c)) };
}

function resultados_(clave) {
  if (String(clave) !== CLAVE || CLAVE === "CAMBIA-ESTA-CLAVE") return { error: "clave" };
  var h = hoja_(), n = h.getLastRow(), salida = [];
  if (n < 2) return { ok: [] };
  var filas = h.getRange(2, 1, n - 1, 3).getValues();
  for (var i = 0; i < filas.length; i++) {
    try {
      var d = JSON.parse(filas[i][2]);
      d.t = filas[i][0] instanceof Date ? filas[i][0].toISOString() : String(filas[i][0]);
      salida.push(d);
    } catch (e) { /* fila dañada: se omite */ }
  }
  return { ok: salida };
}

function doPost(e) {
  var lock = LockService.getScriptLock(), salida, bloqueado = false;
  try {
    var req = JSON.parse(e.postData.contents), a = req.args || [];
    lock.waitLock(20000); bloqueado = true;
    if (req.fn === "enviar") salida = enviar_(a[0]);
    else if (req.fn === "resultados") salida = resultados_(a[0]);
    else if (req.fn === "existe") salida = existe_(a[0]);
    else salida = { error: "datos" };
  } catch (err) {
    salida = { error: "servidor" };
  } finally {
    if (bloqueado) lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput("Encuesta final del rally: servicio activo.");
}
