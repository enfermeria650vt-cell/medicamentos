/**
 * SISTEMA GERIATRICO - APPS SCRIPT v2 (SYNC MULTI-DISPOSITIVO + WHATSAPP)
 * =======================================================================
 * Este script convierte una Google Sheet en una base de datos compartida
 * para que varios celulares/PCs vean los mismos datos.
 *
 * MANTIENE las funciones de WhatsApp del v1.
 * AGREGA endpoints para sincronización completa de datos.
 *
 * CÓMO INSTALARLO:
 * 1. Abrí la Google Sheet
 * 2. Extensiones -> Apps Script
 * 3. Pegá TODO este código (reemplazando el v1 si lo tenías)
 * 4. Guardá (Ctrl+S)
 * 5. Ejecutá la función "inicializarHojas" UNA VEZ para crear las pestañas necesarias
 * 6. Deploy -> Nueva implementación -> Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Copiá la URL y pegala en el HTML (botón Configurar sync)
 *
 * ENDPOINTS:
 *   GET  ?action=pull                       -> Trae todos los datos actuales
 *   POST action=push&key=...&value=...      -> Guarda datos (residentes, semanal o historial)
 *   GET  ?action=ingresos&since=ISO         -> Trae solo ingresos de WhatsApp (compatibilidad v1)
 *   POST (Twilio webhook, sin action)       -> Recibe mensaje de WhatsApp
 */

// ============ CONFIGURACIÓN ============
var TAB_STATE = "STATE_SYNC";        // Almacenamiento principal (JSON)
var TAB_INGRESOS = "📥 INGRESOS";    // Compatibilidad WhatsApp v1
var TAB_STOCK = "📦 STOCK_GENERAL";  // Compatibilidad WhatsApp v1
var TAB_LOG = "LOG_SYNC";            // Log de cambios

// ID de la Google Sheet a la que se conecta este script (standalone)
var SHEET_ID = "1SR_xi-488kLIC_yO7DOoho9WYT2Wee4Ru8NEkBSxSAk";

// ============ CHUNKING (celda de Sheets = máx 50.000 chars) ============
// El value se parte: chunk 0 -> col 2 (value), chunks 1..N -> cols 5..(4+N).
// Con CHUNK=45000 y 10 cols extra soporta ~495.000 chars. Backward-compatible:
// un value viejo guardado entero en col 2 se lee igual (cols extra vacías).
var CHUNK_SIZE = 45000;          // margen seguro bajo el tope de 50.000
var MAX_OVERFLOW_COLS = 10;      // cols 5..14
var STATE_NCOLS = 4 + MAX_OVERFLOW_COLS; // 14 columnas en total

// Asegura que la hoja tenga al menos STATE_NCOLS columnas
function ensureStateCols_(hoja) {
  var faltan = STATE_NCOLS - hoja.getMaxColumns();
  if (faltan > 0) hoja.insertColumnsAfter(hoja.getMaxColumns(), faltan);
}

// Reensambla el value a partir de una fila (array 0-based: [key, val0, at, by, c1, c2, ...])
function reassembleValue_(rowVals) {
  var v = String(rowVals[1] == null ? "" : rowVals[1]);
  for (var c = 4; c < rowVals.length; c++) {
    v += String(rowVals[c] == null ? "" : rowVals[c]);
  }
  return v;
}

// Escribe el value partido en col 2 + cols 5..14, limpiando overflow sobrante.
// Devuelve {ok:true} o {ok:false,error} si excede la capacidad.
function writeValueChunked_(hoja, rowIdx, value) {
  ensureStateCols_(hoja);
  var chunks = [];
  for (var i = 0; i < value.length; i += CHUNK_SIZE) {
    chunks.push(value.substr(i, CHUNK_SIZE));
  }
  if (chunks.length === 0) chunks = [""];
  if (chunks.length > MAX_OVERFLOW_COLS + 1) {
    return { ok: false, error: "value demasiado grande (" + value.length + " chars; máx " + ((MAX_OVERFLOW_COLS + 1) * CHUNK_SIZE) + ")" };
  }
  hoja.getRange(rowIdx, 2).setValue(chunks[0]);
  for (var k = 1; k <= MAX_OVERFLOW_COLS; k++) {
    hoja.getRange(rowIdx, 4 + k).setValue(chunks[k] || "");
  }
  return { ok: true };
}

// ============ INICIALIZACIÓN (correr 1 vez desde el editor) ============
function inicializarHojas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hojaState = ss.getSheetByName(TAB_STATE);
  if (!hojaState) {
    hojaState = ss.insertSheet(TAB_STATE);
    hojaState.appendRow(["key", "value", "updated_at", "updated_by"]);
    hojaState.appendRow(["residentes", "[]", new Date().toISOString(), "sistema"]);
    hojaState.appendRow(["semanal", "{}", new Date().toISOString(), "sistema"]);
    hojaState.appendRow(["historial", "[]", new Date().toISOString(), "sistema"]);
    hojaState.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#1F4E78").setFontColor("white");
    hojaState.setColumnWidth(1, 120);
    hojaState.setColumnWidth(2, 600);
    hojaState.setColumnWidth(3, 180);
    hojaState.setColumnWidth(4, 140);
    Logger.log("✅ Hoja STATE_SYNC creada");
  } else {
    Logger.log("ℹ️ STATE_SYNC ya existía");
  }
  var hojaLog = ss.getSheetByName(TAB_LOG);
  if (!hojaLog) {
    hojaLog = ss.insertSheet(TAB_LOG);
    hojaLog.appendRow(["fecha", "accion", "key", "usuario", "tamaño_bytes"]);
    hojaLog.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1F4E78").setFontColor("white");
    Logger.log("✅ Hoja LOG_SYNC creada");
  }
  var hojaIng = ss.getSheetByName(TAB_INGRESOS);
  if (!hojaIng) {
    hojaIng = ss.insertSheet(TAB_INGRESOS);
    hojaIng.appendRow(["INGRESOS DE MEDICAMENTOS"]);
    hojaIng.appendRow([""]);
    hojaIng.appendRow(["Fecha", "Residente", "Medicamento", "Cantidad", "Recibido por", "Observaciones"]);
    hojaIng.getRange(3, 1, 1, 6).setFontWeight("bold").setBackground("#1F4E78").setFontColor("white");
    Logger.log("✅ Hoja INGRESOS creada");
  }
  Logger.log("✅ Inicialización completa");
}

// ============ DO GET ============
function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || "pull";

    if (action === "pull") return respJSON_(handlePull_(params));
    if (action === "ingresos") return respJSON_(handleIngresosLegacy_(params));
    if (action === "ping") return respJSON_({ ok: true, ahora: new Date().toISOString(), version: "v2" });

    return respJSON_({ ok: false, error: "Acción GET no reconocida: " + action });
  } catch (err) {
    logError_("doGet", err);
    return respJSON_({ ok: false, error: String(err) });
  }
}

// ============ DO POST ============
function doPost(e) {
  try {
    var params = e.parameter || {};

    // Twilio manda Body y From -> WhatsApp
    if (params.Body !== undefined && params.From !== undefined) {
      return handleWhatsapp_(params);
    }

    // Nuevo: acción push
    if (params.action === "push") {
      return respJSON_(handlePush_(params));
    }

    return respJSON_({ ok: false, error: "POST sin acción válida" });
  } catch (err) {
    logError_("doPost", err);
    return respJSON_({ ok: false, error: String(err) });
  }
}

// ============ PULL: traer todos los datos ============
function handlePull_(params) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = ss.getSheetByName(TAB_STATE);
  if (!hoja) {
    return { ok: false, error: "Hoja " + TAB_STATE + " no existe. Ejecutá inicializarHojas() desde el editor." };
  }
  var ncol = Math.min(STATE_NCOLS, hoja.getMaxColumns());
  var data = hoja.getRange(2, 1, Math.max(1, hoja.getLastRow() - 1), ncol).getValues();
  var resp = { ok: true, ahora: new Date().toISOString(), data: {}, meta: {} };
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0] || "").trim();
    if (!key) continue;
    var value = reassembleValue_(data[i]);
    try {
      resp.data[key] = JSON.parse(value || "null");
    } catch (e) {
      resp.data[key] = null;
    }
    resp.meta[key] = {
      updatedAt: data[i][2] ? new Date(data[i][2]).toISOString() : null,
      updatedBy: String(data[i][3] || "")
    };
  }
  return resp;
}

// ============ PUSH: guardar datos ============
function handlePush_(params) {
  var key = String(params.key || "").trim();
  var value = String(params.value || "");
  var user = String(params.user || "anónimo");
  var lastSeen = params.lastSeenAt ? new Date(params.lastSeenAt) : null;

  if (!key || !value) return { ok: false, error: "Faltan key o value" };
  var keysOK = ["residentes", "semanal", "historial", "recetas", "controlSignos", "geminiKey"];
  if (keysOK.indexOf(key) === -1) return { ok: false, error: "Key inválida. Usá: " + keysOK.join(", ") };

  // Validar JSON
  try { JSON.parse(value); } catch (e) {
    return { ok: false, error: "value no es JSON válido: " + e };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hoja = ss.getSheetByName(TAB_STATE);
  if (!hoja) return { ok: false, error: "Hoja " + TAB_STATE + " no existe" };

  // Lock para evitar escrituras simultáneas
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: "Otro dispositivo está escribiendo, intentá de nuevo" }; }

  try {
    var data = hoja.getRange(2, 1, Math.max(1, hoja.getLastRow() - 1), 4).getValues();
    var rowIdx = -1;
    var currentUpdatedAt = null;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || "").trim() === key) {
        rowIdx = i + 2; // 1-based + header
        currentUpdatedAt = data[i][2] ? new Date(data[i][2]) : null;
        break;
      }
    }

    // Detectar conflicto: si lastSeenAt < currentUpdatedAt, alguien escribió después
    var hayConflicto = false;
    if (lastSeen && currentUpdatedAt && lastSeen < currentUpdatedAt) {
      hayConflicto = true;
    }

    var ahora = new Date();
    if (rowIdx === -1) {
      hoja.appendRow([key, "", ahora.toISOString(), user]);
      rowIdx = hoja.getLastRow();
    }
    var w = writeValueChunked_(hoja, rowIdx, value);
    if (!w.ok) return { ok: false, error: w.error };
    hoja.getRange(rowIdx, 3).setValue(ahora.toISOString());
    hoja.getRange(rowIdx, 4).setValue(user);

    // Log
    var hojaLog = ss.getSheetByName(TAB_LOG);
    if (hojaLog) {
      hojaLog.appendRow([ahora, "PUSH" + (hayConflicto ? " (sobrescribió)" : ""), key, user, value.length]);
    }

    return {
      ok: true,
      updatedAt: ahora.toISOString(),
      hayConflicto: hayConflicto,
      bytes: value.length
    };
  } finally {
    lock.releaseLock();
  }
}

// ============ INGRESOS LEGACY (compatibilidad con sync WhatsApp v1) ============
function handleIngresosLegacy_(params) {
  var since = params.since || "";
  var sinceDate = since ? new Date(since) : new Date(0);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hojaIng = ss.getSheetByName(TAB_INGRESOS);
  var ingresos = [];
  if (hojaIng && hojaIng.getLastRow() > 3) {
    var dataIng = hojaIng.getRange(4, 1, hojaIng.getLastRow() - 3, 6).getValues();
    for (var i = 0; i < dataIng.length; i++) {
      var fila = dataIng[i];
      if (!fila[0]) continue;
      var fecha = fila[0] instanceof Date ? fila[0] : new Date(fila[0]);
      if (fecha <= sinceDate) continue;
      ingresos.push({
        fecha: fecha.toISOString(),
        residente: String(fila[1] || ""),
        medicamento: String(fila[2] || ""),
        cantidad: Number(fila[3]) || 0,
        recibido: String(fila[4] || ""),
        obs: String(fila[5] || "")
      });
    }
  }
  return { ok: true, ahora: new Date().toISOString(), ingresos: ingresos, total: ingresos.length };
}

// ============ WHATSAPP (Twilio) ============
function handleWhatsapp_(params) {
  var from = params.From || "desconocido";
  var body = (params.Body || "").trim();
  log_("WSP RECIBIDO de " + from + ": " + body);

  try {
    var upper = body.toUpperCase();
    var respuesta;

    if (upper.startsWith("INGRESO:") || upper.startsWith("INGRESO ")) {
      respuesta = procesarIngresoWsp_(body, from);
    } else if (upper.startsWith("CONSULTA:") || upper.startsWith("CONSULTA ")) {
      respuesta = procesarConsultaWsp_(body);
    } else if (upper === "ALERTAS" || upper === "ALERTA") {
      respuesta = obtenerAlertasGlobalesWsp_();
    } else if (upper === "AYUDA" || upper === "HELP" || upper === "?") {
      respuesta = mensajeAyudaWsp_();
    } else if (body.indexOf("|") !== -1) {
      respuesta = procesarIngresoWsp_("INGRESO: " + body, from);
    } else {
      respuesta = "🤔 No entendí.\n\n" + mensajeAyudaWsp_();
    }
    return responderTwilio_(respuesta);
  } catch (err) {
    log_("ERROR WSP: " + err);
    return responderTwilio_("❌ Error procesando el mensaje. Mandá AYUDA para ver los comandos.");
  }
}

function procesarIngresoWsp_(body, from) {
  var contenido = body.replace(/^INGRESO\s*:?/i, "").trim();
  var partes = contenido.split("|").map(function (s) { return s.trim(); });
  if (partes.length < 3) {
    return "❌ Formato:\nINGRESO: residente | medicamento | cantidad\n\nEj: INGRESO: Irma Peduzzi | ATENOLOL | 50";
  }
  var residenteInput = partes[0];
  var medInput = partes[1];
  var cantidad = parseFloat(partes[2].replace(",", "."));
  if (isNaN(cantidad) || cantidad <= 0) return "❌ Cantidad inválida: " + partes[2];

  // Leer residentes desde STATE_SYNC
  var pullResp = handlePull_({});
  var residentes = (pullResp.data && pullResp.data.residentes) || [];
  if (!Array.isArray(residentes) || residentes.length === 0) {
    return "❌ No hay residentes cargados en el sistema.";
  }

  // Buscar match flexible
  var match = null;
  for (var i = 0; i < residentes.length; i++) {
    if (sim_(residentes[i].nombre, residenteInput)) {
      for (var j = 0; j < residentes[i].meds.length; j++) {
        if (sim_(residentes[i].meds[j].med, medInput)) {
          match = { rIdx: i, mIdx: j, r: residentes[i], m: residentes[i].meds[j] };
          break;
        }
      }
      if (match) break;
    }
  }

  if (!match) {
    // Sugerir medicamentos del residente si existe
    var resMatch = residentes.find(function(r) { return sim_(r.nombre, residenteInput); });
    if (resMatch) {
      return "❌ No encontré \"" + medInput + "\" para " + residenteInput + ".\n\nMedicamentos:\n• " +
        resMatch.meds.map(function(m){return m.med;}).slice(0, 12).join("\n• ");
    }
    return "❌ No encontré " + residenteInput + " + " + medInput + ".\nRevisá los nombres o mandá CONSULTA: residente";
  }

  // Actualizar stock y fechaIngreso
  match.m.stock = (Number(match.m.stock) || 0) + cantidad;
  match.m.fechaIngreso = new Date().toISOString().slice(0, 10);

  // Guardar residentes
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return "❌ Otro proceso está escribiendo, probá de nuevo en 5 seg"; }
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hoja = ss.getSheetByName(TAB_STATE);
    if (hoja) {
      // Buscar la fila de residentes
      var data = hoja.getRange(2, 1, hoja.getLastRow() - 1, 4).getValues();
      for (var k = 0; k < data.length; k++) {
        if (String(data[k][0]).trim() === "residentes") {
          writeValueChunked_(hoja, k + 2, JSON.stringify(residentes));
          hoja.getRange(k + 2, 3).setValue(new Date().toISOString());
          hoja.getRange(k + 2, 4).setValue("WSP " + from);
          break;
        }
      }
    }

    // Registrar en hoja INGRESOS (para el log y para la sincronización legacy)
    var hojaIng = ss.getSheetByName(TAB_INGRESOS);
    if (hojaIng) {
      hojaIng.appendRow([new Date(), match.r.nombre, match.m.med, cantidad, "WhatsApp (" + from + ")", ""]);
    }
  } finally {
    lock.releaseLock();
  }

  return "✅ Registrado:\n\n👤 " + match.r.nombre + "\n💊 " + match.m.med + "\n➕ +" + cantidad + " unidades\n📦 Stock: " + match.m.stock;
}

function procesarConsultaWsp_(body) {
  var residenteInput = body.replace(/^CONSULTA\s*:?/i, "").trim();
  if (!residenteInput) return "❌ Indicá el residente. Ej: CONSULTA: Fernanda";

  var pull = handlePull_({});
  var residentes = (pull.data && pull.data.residentes) || [];
  var r = residentes.find(function(x) { return sim_(x.nombre, residenteInput); });
  if (!r) return "❌ No encontré el residente \"" + residenteInput + "\"";

  var lineas = r.meds.map(function(m) {
    var dias = (typeof m.stock === "number" && typeof m.consumo === "number" && m.consumo > 0)
      ? Math.floor(m.stock / m.consumo) : null;
    var estado = dias === null ? "?" : dias === 0 ? "SIN STOCK" : dias <= 7 ? "CRÍTICO" : dias <= 14 ? "ATENCIÓN" : "OK";
    return "• " + m.med + ": " + (m.stock !== null ? m.stock : "?") + " (" + estado + ")";
  });
  return "📋 " + r.nombre + ":\n\n" + lineas.join("\n");
}

function obtenerAlertasGlobalesWsp_() {
  var pull = handlePull_({});
  var residentes = (pull.data && pull.data.residentes) || [];
  var criticos = [], sinStock = [];
  residentes.forEach(function(r) {
    r.meds.forEach(function(m) {
      if (typeof m.stock !== "number" || typeof m.consumo !== "number" || m.consumo <= 0) return;
      var dias = Math.floor(m.stock / m.consumo);
      if (dias === 0) sinStock.push("• " + r.nombre + " - " + m.med);
      else if (dias <= 7) criticos.push("• " + r.nombre + " - " + m.med + " (" + dias + " días)");
    });
  });
  var resp = "🚨 ALERTAS\n\n";
  if (sinStock.length > 0) resp += "⚫ SIN STOCK (" + sinStock.length + "):\n" + sinStock.slice(0, 15).join("\n") + "\n\n";
  if (criticos.length > 0) resp += "🔴 CRÍTICOS (" + criticos.length + "):\n" + criticos.slice(0, 15).join("\n");
  if (criticos.length === 0 && sinStock.length === 0) resp = "🟢 Todo bajo control.";
  return resp;
}

function mensajeAyudaWsp_() {
  return "🤖 COMANDOS:\n\n" +
    "📥 INGRESO: residente | med | cantidad\n" +
    "🔍 CONSULTA: residente\n" +
    "🚨 ALERTAS\n" +
    "❓ AYUDA";
}

// ============ HELPERS ============
function sim_(a, b) {
  if (!a || !b) return false;
  a = quitarAcentos_(String(a).toLowerCase()).replace(/[^a-z0-9]+/g, " ").trim();
  b = quitarAcentos_(String(b).toLowerCase()).replace(/[^a-z0-9]+/g, " ").trim();
  if (!a || !b) return false;
  return a.indexOf(b) !== -1 || b.indexOf(a) !== -1;
}
function quitarAcentos_(s) { return s.normalize("NFD").replace(/[̀-ͯ]/g, ""); }

function respJSON_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function responderTwilio_(texto) {
  var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Response><Message>" + escapeXml_(texto) + "</Message></Response>";
  return ContentService.createTextOutput(xml).setMimeType(ContentService.MimeType.XML);
}

function escapeXml_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function log_(msg) {
  console.log(msg);
}

function logError_(donde, err) {
  console.log("ERROR " + donde + ": " + err);
}

// ============ TESTS (correr desde el editor) ============
function testPull() {
  var resp = handlePull_({});
  Logger.log(JSON.stringify(resp).substring(0, 500));
}
function testInicializar() {
  inicializarHojas();
}
function testIngresoWsp() {
  var resp = handleWhatsapp_({ Body: "INGRESO: Irma Peduzzi | ATENOLOL | 5", From: "whatsapp:+test" });
  Logger.log(resp.getContent());
}
