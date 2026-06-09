// =====================================================
// SISTEMA INDICACIONES MEDICAS - GERIATRICO VILLA TERESA
// Pestaña INDICACIONES sincronizada desde la carpeta Drive
// + Respaldo semanal + Alerta de modificacion por seccion
// =====================================================
// IMPORTANTE: Antes de correr, habilitar "Drive API" en
//   Apps Script editor -> Services (icono +) -> Drive API v2
// =====================================================

const FOLDER_INDICACIONES_ID = '1worGuGb8fm79t-4A_3AztA-zxymcpyZf';
const AUTHORIZED_EMAIL = 'enfermeria650vt@gmail.com';
const SHEET_INDICACIONES = 'INDICACIONES';
const SHEET_RESPALDO = 'INDICACIONES_RESPALDO';
const SHEET_HISTORIAL = 'INDICACIONES_HISTORIAL';
const NCOL = 14;
const HEADERS = ['Residente', 'FileID', 'Modif. Drive', 'Última Sync', 'Hash',
                 'Diagnóstico', 'HTML', 'Texto', 'HashMed', 'HashHor', 'HashEvo',
                 'Modificado', 'FechaModif', 'HtmlAnterior'];

// --- INSUMOS (uso diario, por turno: M=mañana, T=tarde, N=noche, P=propio) ---
const SHEET_INSUMOS = 'INSUMOS';
// 16 campos por turno, en el MISMO orden que el frontend (INSUMOS_CAMPOS_TURNOS)
const INSUMOS_FIELDS = [
  'pañal_m','pañal_t','pañal_n','pañal_p',
  'zalea_m','zalea_t','zalea_n','zalea_p',
  'aposito_m','aposito_t','aposito_n','aposito_p',
  'bombacha_m','bombacha_t','bombacha_n','bombacha_p'
];
const INSUMOS_HEADERS = ['Residente', 'AñoMes', 'Dia',
  'Pañal_M','Pañal_T','Pañal_N','Pañal_P',
  'Zalea_M','Zalea_T','Zalea_N','Zalea_P',
  'Aposito_M','Aposito_T','Aposito_N','Aposito_P',
  'Bombacha_M','Bombacha_T','Bombacha_N','Bombacha_P'];
const INSUMOS_NCOLS = 3 + INSUMOS_FIELDS.length; // 19

// --- STOCK DE INSUMOS ---
const SHEET_STOCK = 'STOCK_INSUMOS';
const SHEET_INGRESOS = 'INGRESOS_INSUMOS';
const STOCK_HEADERS = ['Residente', 'MinPañal', 'MinZalea', 'MinAposito', 'MinBombacha'];
const INGRESOS_HEADERS = ['Fecha', 'Residente', 'Pañal', 'Zalea', 'Aposito', 'Bombacha', 'Observacion'];

// --- HISTORIAL DE CAMBIOS ---
const SHEET_HISTORIAL_CAMBIOS = 'HISTORIAL_CAMBIOS';
const HISTORIAL_HEADERS = ['Fecha', 'Usuario', 'Rol', 'Acción', 'Residente', 'Detalle'];

// =====================================================
// SETUP
// =====================================================

function setupIndicaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_INDICACIONES);
  if (!sheet) sheet = ss.insertSheet(SHEET_INDICACIONES);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(6, 260);
  sheet.setColumnWidth(12, 200);
  sheet.setColumnWidth(13, 140);
  sheet.setColumnWidth(14, 60);
  sheet.hideColumns(2);
  sheet.hideColumns(5);
  sheet.hideColumns(7);
  sheet.hideColumns(8);
  sheet.hideColumns(9, 3);
  sheet.hideColumns(14);
}

function setupInsumos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_INSUMOS);
  if (!sheet) sheet = ss.insertSheet(SHEET_INSUMOS);
  sheet.clear();
  sheet.getRange(1, 1, 1, INSUMOS_HEADERS.length).setValues([INSUMOS_HEADERS]);
  sheet.getRange(1, 1, 1, INSUMOS_HEADERS.length)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 50);
  SpreadsheetApp.flush();
}

function setupStockInsumos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hoja de mínimos por residente
  let stock = ss.getSheetByName(SHEET_STOCK);
  if (!stock) stock = ss.insertSheet(SHEET_STOCK);
  stock.clear();
  stock.getRange(1, 1, 1, STOCK_HEADERS.length).setValues([STOCK_HEADERS]);
  stock.getRange(1, 1, 1, STOCK_HEADERS.length)
    .setFontWeight('bold').setBackground('#0f9d58').setFontColor('white');
  stock.setFrozenRows(1);
  stock.setColumnWidth(1, 220);
  for (let c = 2; c <= 5; c++) stock.setColumnWidth(c, 90);

  // Hoja de ingresos (historial de entregas)
  let ingresos = ss.getSheetByName(SHEET_INGRESOS);
  if (!ingresos) ingresos = ss.insertSheet(SHEET_INGRESOS);
  ingresos.clear();
  ingresos.getRange(1, 1, 1, INGRESOS_HEADERS.length).setValues([INGRESOS_HEADERS]);
  ingresos.getRange(1, 1, 1, INGRESOS_HEADERS.length)
    .setFontWeight('bold').setBackground('#0f9d58').setFontColor('white');
  ingresos.setFrozenRows(1);
  ingresos.setColumnWidth(1, 160);
  ingresos.setColumnWidth(2, 220);
  for (let c = 3; c <= 6; c++) ingresos.setColumnWidth(c, 80);
  ingresos.setColumnWidth(7, 200);

  SpreadsheetApp.flush();
  return { ok: true };
}

// =====================================================
// UTILIDADES
// =====================================================

function md5(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text || '');
  return bytes.map(function(b){ return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function isAuthorized(email) {
  return email === AUTHORIZED_EMAIL || Session.getActiveUser().getEmail() === AUTHORIZED_EMAIL;
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function etiquetaSemana(d) {
  const date = new Date(d.getTime());
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() + 3 - ((date.getDay() + 6) % 7));
  const week1 = new Date(date.getFullYear(), 0, 4);
  const semana = 1 + Math.round(((date - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
  return date.getFullYear() + '-S' + String(semana).padStart(2, '0');
}

// =====================================================
// SYNC INDICACIONES
// =====================================================

function extraerSecciones(doc) {
  const body = doc.getBody();
  let evo = '', med = '', hor = '';
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    const type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const t = el.asParagraph().getText().trim();
      if (t) evo += t + '\n';
    } else if (type === DocumentApp.ElementType.TABLE) {
      const tbl = el.asTable();
      for (let r = 1; r < tbl.getNumRows(); r++) {
        const cols = tbl.getRow(r).getNumCells();
        for (let c = 0; c < cols; c++) {
          const txt = tbl.getCell(r, c).getText().trim();
          if (c <= 1) med += txt + '\n';
          else hor += txt + '\n';
        }
      }
    }
  }
  return { med: med, hor: hor, evo: evo };
}

function syncIndicaciones() {
  const folder = DriveApp.getFolderById(FOLDER_INDICACIONES_ID);
  const files = folder.getFiles();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_INDICACIONES);
  if (!sheet) { setupIndicaciones(); sheet = ss.getSheetByName(SHEET_INDICACIONES); }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  const lastRow = sheet.getLastRow();
  const existing = {};
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, NCOL).getValues();
    data.forEach(function(row, idx){
      if (row[1]) existing[row[1]] = {
        rowIndex: idx + 2, hash: row[4], modifDrive: row[2],
        hashMed: row[8], hashHor: row[9], hashEvo: row[10],
        modificado: row[11], fechaModif: row[12],
        html: row[6], texto: row[7],
        htmlAnterior: row[13]
      };
    });
  }

  const seenIds = new Set();
  const now = new Date();
  let updates = 0, news = 0;

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.MICROSOFT_WORD) continue;
    const fileId = file.getId();
    seenIds.add(fileId);
    const fileName = file.getName().replace(/\.docx$/i, '');
    const modifTime = file.getLastUpdated();
    const prev = existing[fileId];

    if (prev && prev.modifDrive instanceof Date && modifTime instanceof Date &&
        modifTime.getTime() <= prev.modifDrive.getTime() + 5000 &&
        prev.hash && prev.hashMed && prev.hashHor && prev.hashEvo) {
      continue;
    }

    let docId;
    try {
      const blob = file.getBlob();
      const resource = { title: '_tmp_indic_' + fileId, mimeType: MimeType.GOOGLE_DOCS };
      const tmp = Drive.Files.insert(resource, blob, { convert: true });
      docId = tmp.id;
    } catch (e) {
      console.error('Error convirtiendo ' + fileName + ': ' + e);
      continue;
    }

    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();
    const hash = md5(text);

    if (prev && prev.hash === hash && prev.hashMed && prev.hashHor && prev.hashEvo) {
      DriveApp.getFileById(docId).setTrashed(true);
      continue;
    }

    const sec = extraerSecciones(doc);
    const hMed = md5(sec.med), hHor = md5(sec.hor), hEvo = md5(sec.evo);

    let modificado = prev ? (prev.modificado || '') : '';
    let fechaModif = prev ? (prev.fechaModif || '') : '';
    let htmlAnterior = prev ? (prev.htmlAnterior || '') : '';

    const html = renderDocToHTML(doc);

    if (prev && prev.hash !== hash) {
      if (prev.hashMed || prev.hashHor || prev.hashEvo) {
        const cambios = [];
        if (prev.hashMed !== hMed) cambios.push('medicamento');
        if (prev.hashHor !== hHor) cambios.push('horario');
        if (prev.hashEvo !== hEvo) cambios.push('evolución');
        modificado = cambios.length ? cambios.join(', ') : 'actualizado';
      } else {
        modificado = 'actualizado';
      }
      fechaModif = now;
      // Archivar la versión que estamos por reemplazar (historial completo)
      archivarVersionIndicacion(fileName, fileId, modificado, prev.html, prev.texto, now);
      htmlAnterior = prev.html || '';
    }

    const diagMatch = text.match(/Diagn[óo]stico:\s*([^\n]+)/i);
    const diagnostico = diagMatch ? diagMatch[1].trim() : '';

    const row = [fileName, fileId, modifTime, now, hash, diagnostico, html, text,
                 hMed, hHor, hEvo, modificado, fechaModif, htmlAnterior];
    if (prev) {
      sheet.getRange(prev.rowIndex, 1, 1, NCOL).setValues([row]);
      updates++;
    } else {
      sheet.appendRow(row);
      news++;
    }
    DriveApp.getFileById(docId).setTrashed(true);
  }

  const newLast = sheet.getLastRow();
  if (newLast > 1) {
    const data = sheet.getRange(2, 1, newLast - 1, NCOL).getValues();
    const toDel = [];
    data.forEach(function(row, idx){ if (!seenIds.has(row[1])) toDel.push(idx + 2); });
    toDel.reverse().forEach(function(r){ sheet.deleteRow(r); });
  }

  const fr = sheet.getLastRow();
  if (fr > 2) sheet.getRange(2, 1, fr - 1, NCOL).sort([{ column: 1, ascending: true }]);
  SpreadsheetApp.flush();
  return { updates: updates, news: news, total: seenIds.size };
}

function respaldoSemanalIndicaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SHEET_INDICACIONES);
  if (!src) return { error: 'sin hoja INDICACIONES' };
  const last = src.getLastRow();
  if (last < 2) return { filas: 0 };

  let bak = ss.getSheetByName(SHEET_RESPALDO);
  if (!bak) {
    bak = ss.insertSheet(SHEET_RESPALDO);
    bak.getRange(1, 1, 1, 6).setValues([['Semana', 'Fecha respaldo', 'Residente', 'Diagnóstico', 'HTML', 'Texto']]);
    bak.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    bak.setFrozenRows(1);
    bak.setColumnWidth(1, 90);
    bak.setColumnWidth(2, 150);
    bak.setColumnWidth(3, 220);
  }
  const now = new Date();
  const semana = etiquetaSemana(now);
  const data = src.getRange(2, 1, last - 1, 8).getValues();
  const filas = data.map(function(r){ return [semana, now, r[0], r[5], r[6], r[7]]; });
  if (filas.length) bak.getRange(bak.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
  SpreadsheetApp.flush();
  return { semana: semana, filas: filas.length };
}

// Guarda en INDICACIONES_HISTORIAL la versión anterior de una indicación
// (se llama desde syncIndicaciones cada vez que se detecta un cambio).
function archivarVersionIndicacion(fileName, fileId, modificado, html, texto, fecha) {
  if (!html && !texto) return; // nada que archivar
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(SHEET_HISTORIAL);
  if (!h) {
    h = ss.insertSheet(SHEET_HISTORIAL);
    h.getRange(1, 1, 1, 6).setValues([['FechaArchivado', 'FileID', 'Residente', 'Modificado', 'HTML', 'Texto']]);
    h.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    h.setFrozenRows(1);
    h.setColumnWidth(1, 150);
    h.setColumnWidth(3, 220);
    h.hideColumns(5);
    h.hideColumns(6);
  }
  h.appendRow([fecha, fileId, fileName, modificado || '', html || '', texto || '']);
}

function apiGetIndicacionHistorial(fileId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HISTORIAL);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data.filter(function(r){ return r[1] === fileId; }).map(function(r){
    return {
      fecha: r[0] instanceof Date ? r[0].toISOString() : r[0],
      residente: r[2],
      modificado: r[3] || '',
      html: r[4] || '',
      texto: r[5] || ''
    };
  }).reverse(); // más reciente primero
}

function renderDocToHTML(doc) {
  const body = doc.getBody();
  let html = '';
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    const type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const t = escapeHtml(el.asParagraph().getText());
      if (t.trim()) html += '<p>' + t + '</p>';
    } else if (type === DocumentApp.ElementType.TABLE) {
      const tbl = el.asTable();
      html += '<table class="ind-tabla">';
      for (let r = 0; r < tbl.getNumRows(); r++) {
        html += '<tr>';
        const cols = tbl.getRow(r).getNumCells();
        for (let c = 0; c < cols; c++) {
          const txt = escapeHtml(tbl.getCell(r, c).getText()).replace(/\n/g, '<br>');
          const tag = r === 0 ? 'th' : 'td';
          html += '<' + tag + '>' + txt + '</' + tag + '>';
        }
        html += '</tr>';
      }
      html += '</table>';
    }
  }
  return html;
}

// =====================================================
// API doGet
// =====================================================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    let payload;
    switch (action) {
      case 'getIndicaciones': payload = apiGetIndicaciones(); break;
      case 'getIndicacion':   payload = apiGetIndicacion(e.parameter.id); break;
      case 'getIndicacionHistorial': payload = apiGetIndicacionHistorial(e.parameter.id); break;
      case 'syncNow':
        if (!isAuthorized(e.parameter.email)) return jsonError('No autorizado');
        payload = { sync: syncIndicaciones() };
        break;
      case 'respaldoNow':
        if (!isAuthorized(e.parameter.email)) return jsonError('No autorizado');
        payload = { respaldo: respaldoSemanalIndicaciones() };
        break;
      // INSUMOS - uso diario
      case 'getInsumos':
        payload = apiGetInsumos(e.parameter.residente, e.parameter.mes);
        break;
      case 'setInsumo':
        payload = apiSetInsumo(
          e.parameter.residente, e.parameter.mes, e.parameter.dia,
          e.parameter
        );
        break;
      // STOCK de insumos
      case 'getStockResumen':
        payload = apiGetStockResumen(e.parameter.residente);
        break;
      case 'getAllStock':
        payload = apiGetAllStock();
        break;
      case 'addIngreso':
        payload = apiAddIngreso(
          e.parameter.residente,
          e.parameter.pañal, e.parameter.zalea, e.parameter.aposito, e.parameter.bombacha,
          e.parameter.obs || ''
        );
        break;
      case 'getIngresos':
        payload = apiGetIngresos(e.parameter.residente);
        break;
      // HISTORIAL DE CAMBIOS
      case 'registrarCambio':
        payload = apiRegistrarCambio(
          e.parameter.usuario, e.parameter.rol,
          e.parameter.accion, e.parameter.residente || '',
          e.parameter.detalle || ''
        );
        break;
      case 'getHistorial':
        payload = apiGetHistorial(e.parameter.residente || '', Number(e.parameter.limit) || 200);
        break;
      case 'setMinimos':
        payload = apiSetMinimos(
          e.parameter.residente,
          e.parameter.minPañal, e.parameter.minZalea,
          e.parameter.minAposito, e.parameter.minBombacha
        );
        break;

      // NUTRICIÓN
      case 'getNutricionDieta':
        payload = apiGetNutricionDieta(e.parameter.residente);
        break;
      case 'setNutricionDieta':
        payload = apiSetNutricionDieta(e.parameter.residente, e.parameter);
        break;
      case 'getNutricionPesos':
        payload = apiGetNutricionPesos(e.parameter.residente);
        break;
      case 'addNutricionPeso':
        payload = apiAddNutricionPeso(e.parameter.residente, e.parameter);
        break;
      case 'getNutricionRegistros':
        payload = apiGetNutricionRegistros(e.parameter.residente, e.parameter.mes);
        break;
      case 'setNutricionRegistro':
        payload = apiSetNutricionRegistro(e.parameter.residente, e.parameter.mes, e.parameter.dia, e.parameter);
        break;
      case 'getNutricionNotas':
        payload = apiGetNutricionNotas(e.parameter.residente);
        break;
      case 'addNutricionNota':
        payload = apiAddNutricionNota(e.parameter.residente, e.parameter.texto, e.parameter.usuario, e.parameter.fecha);
        break;

        case 'getNutricionMenu': payload = apiGetNutricionMenu(e.parameter.semana); break;
        case 'setNutricionMenu': payload = apiSetNutricionMenu(e.parameter.semana, e.parameter.data, e.parameter.usuario); break;
      default: payload = { ok: true, action: 'noop' };
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: payload }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return jsonError(String(err));
  }
}

function jsonError(msg) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// API INDICACIONES
// =====================================================

function apiGetIndicaciones() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INDICACIONES);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, NCOL).getValues();
  return data.map(function(row){
    return {
      residente: row[0], fileId: row[1],
      modifDrive: row[2] instanceof Date ? row[2].toISOString() : row[2],
      ultimaSync: row[3] instanceof Date ? row[3].toISOString() : row[3],
      diagnostico: row[5],
      modificado: row[11] || '',
      fechaModif: row[12] instanceof Date ? row[12].toISOString() : (row[12] || '')
    };
  });
}

function apiGetIndicacion(fileId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INDICACIONES);
  if (!sheet) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const data = sheet.getRange(2, 1, last - 1, NCOL).getValues();
  const row = data.find(function(r){ return r[1] === fileId; });
  if (!row) return null;
  return {
    residente: row[0], fileId: row[1],
    modifDrive: row[2] instanceof Date ? row[2].toISOString() : row[2],
    ultimaSync: row[3] instanceof Date ? row[3].toISOString() : row[3],
    hash: row[4], diagnostico: row[5],
    html: row[6], texto: row[7],
    hashMed: row[8], hashHor: row[9], hashEvo: row[10],
    modificado: row[11] || '',
    fechaModif: row[12] instanceof Date ? row[12].toISOString() : (row[12] || ''),
    htmlAnterior: row[13] || ''
  };
}

// =====================================================
// API INSUMOS (uso diario)
// =====================================================

// Devuelve la hoja INSUMOS, creándola con encabezados si no existe.
function ensureInsumosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_INSUMOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_INSUMOS);
    sheet.getRange(1, 1, 1, INSUMOS_HEADERS.length).setValues([INSUMOS_HEADERS]);
    sheet.getRange(1, 1, 1, INSUMOS_HEADERS.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
  }
  return sheet;
}

function apiGetInsumos(residente, mes) {
  const sheet = ensureInsumosSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, INSUMOS_NCOLS).getValues();
  const result = [];
  data.forEach(function(row) {
    if (String(row[0]).toLowerCase() === String(residente).toLowerCase() &&
        String(row[1]) === String(mes)) {
      const o = { dia: row[2] };
      INSUMOS_FIELDS.forEach(function(f, i) { o[f] = Number(row[3 + i]) || 0; });
      result.push(o);
    }
  });
  return result;
}

// p = e.parameter (objeto con los 16 campos por turno)
function apiSetInsumo(residente, mes, dia, p) {
  const sheet = ensureInsumosSheet_();
  const last = sheet.getLastRow();
  const diaNum = Number(dia);
  p = p || {};
  const valores = INSUMOS_FIELDS.map(function(f) { return Number(p[f]) || 0; });
  // Buscar fila existente (Residente + AñoMes + Dia)
  if (last >= 2) {
    const data = sheet.getRange(2, 1, last - 1, 3).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(residente).toLowerCase() &&
          String(data[i][1]) === String(mes) &&
          Number(data[i][2]) === diaNum) {
        sheet.getRange(i + 2, 4, 1, INSUMOS_FIELDS.length).setValues([valores]);
        SpreadsheetApp.flush();
        return { ok: true, action: 'updated' };
      }
    }
  }
  // Nueva fila
  sheet.appendRow([residente, mes, diaNum].concat(valores));
  SpreadsheetApp.flush();
  return { ok: true, action: 'inserted' };
}

// =====================================================
// API STOCK DE INSUMOS
// =====================================================

function apiGetStockResumen(residente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Depósito general: el "usado" suma el consumo de TODOS los residentes.
  const ES_GENERAL = String(residente) === '__DEPOSITO_GENERAL__';

  // Obtener mínimos
  const stockSheet = ss.getSheetByName(SHEET_STOCK);
  let minimos = { minPañal: 0, minZalea: 0, minAposito: 0, minBombacha: 0 };
  if (stockSheet && stockSheet.getLastRow() >= 2) {
    const rows = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 5).getValues();
    const r = rows.find(function(x){ return String(x[0]).toLowerCase() === String(residente).toLowerCase(); });
    if (r) minimos = { minPañal: r[1]||0, minZalea: r[2]||0, minAposito: r[3]||0, minBombacha: r[4]||0 };
  }

  // Sumar ingresos totales
  const ingSheet = ss.getSheetByName(SHEET_INGRESOS);
  let totIng = { pañal: 0, zalea: 0, aposito: 0, bombacha: 0 };
  if (ingSheet && ingSheet.getLastRow() >= 2) {
    const rows = ingSheet.getRange(2, 1, ingSheet.getLastRow() - 1, 7).getValues();
    rows.forEach(function(r) {
      if (String(r[1]).toLowerCase() === String(residente).toLowerCase()) {
        totIng.pañal    += Number(r[2]) || 0;
        totIng.zalea    += Number(r[3]) || 0;
        totIng.aposito  += Number(r[4]) || 0;
        totIng.bombacha += Number(r[5]) || 0;
      }
    });
  }

  // Sumar uso total desde hoja INSUMOS (16 columnas por turno -> 4 totales)
  const insSheet = ss.getSheetByName(SHEET_INSUMOS);
  let totUso = { pañal: 0, zalea: 0, aposito: 0, bombacha: 0 };
  if (insSheet && insSheet.getLastRow() >= 2) {
    const rows = insSheet.getRange(2, 1, insSheet.getLastRow() - 1, INSUMOS_NCOLS).getValues();
    rows.forEach(function(r) {
      if (ES_GENERAL || String(r[0]).toLowerCase() === String(residente).toLowerCase()) {
        // cols: 3..6 pañal, 7..10 zalea, 11..14 aposito, 15..18 bombacha
        for (let i = 0; i < 4; i++)  totUso.pañal    += Number(r[3 + i])  || 0;
        for (let i = 0; i < 4; i++)  totUso.zalea    += Number(r[7 + i])  || 0;
        for (let i = 0; i < 4; i++)  totUso.aposito  += Number(r[11 + i]) || 0;
        for (let i = 0; i < 4; i++)  totUso.bombacha += Number(r[15 + i]) || 0;
      }
    });
  }

  return {
    pañal:    { ingresado: totIng.pañal,    usado: totUso.pañal,    disponible: totIng.pañal    - totUso.pañal,    minimo: minimos.minPañal },
    zalea:    { ingresado: totIng.zalea,    usado: totUso.zalea,    disponible: totIng.zalea    - totUso.zalea,    minimo: minimos.minZalea },
    aposito:  { ingresado: totIng.aposito,  usado: totUso.aposito,  disponible: totIng.aposito  - totUso.aposito,  minimo: minimos.minAposito },
    bombacha: { ingresado: totIng.bombacha, usado: totUso.bombacha, disponible: totIng.bombacha - totUso.bombacha, minimo: minimos.minBombacha }
  };
}

function apiGetAllStock() {
  // Devuelve resumen de stock para todos los residentes que tienen ingresos registrados
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ingSheet = ss.getSheetByName(SHEET_INGRESOS);
  const insSheet = ss.getSheetByName(SHEET_INSUMOS);
  const stockSheet = ss.getSheetByName(SHEET_STOCK);

  // Recopilar lista de residentes únicos (de INGRESOS + STOCK)
  const residentes = new Set();
  if (ingSheet && ingSheet.getLastRow() >= 2) {
    ingSheet.getRange(2, 1, ingSheet.getLastRow() - 1, 2).getValues()
      .forEach(function(r){ if (r[1]) residentes.add(String(r[1])); });
  }
  if (stockSheet && stockSheet.getLastRow() >= 2) {
    stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 1).getValues()
      .forEach(function(r){ if (r[0]) residentes.add(String(r[0])); });
  }

  const result = [];
  residentes.forEach(function(res) {
    const r = apiGetStockResumen(res);
    result.push({ residente: res, stock: r });
  });
  result.sort(function(a, b){ return a.residente.localeCompare(b.residente); });
  return result;
}

function apiAddIngreso(residente, pañal, zalea, aposito, bombacha, obs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_INGRESOS);
  if (!sheet) return { ok: false, error: 'Hoja INGRESOS_INSUMOS no encontrada. Ejecutar setupStockInsumos().' };
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sheet.appendRow([
    fecha, residente,
    Number(pañal) || 0, Number(zalea) || 0,
    Number(aposito) || 0, Number(bombacha) || 0,
    obs || ''
  ]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function apiGetIngresos(residente) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INGRESOS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const result = [];
  data.forEach(function(row) {
    if (!residente || String(row[1]).toLowerCase() === String(residente).toLowerCase()) {
      result.push({
        fecha: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(row[0]),
        residente: row[1],
        pañal: row[2], zalea: row[3], aposito: row[4], bombacha: row[5],
        obs: row[6]
      });
    }
  });
  return result.reverse(); // más reciente primero
}

// =====================================================
// API HISTORIAL DE CAMBIOS
// =====================================================

function ensureHistorialCambiosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_HISTORIAL_CAMBIOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HISTORIAL_CAMBIOS);
    sheet.getRange(1, 1, 1, HISTORIAL_HEADERS.length).setValues([HISTORIAL_HEADERS]);
    sheet.getRange(1, 1, 1, HISTORIAL_HEADERS.length)
      .setFontWeight('bold').setBackground('#4a4a8a').setFontColor('white');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150); // Fecha
    sheet.setColumnWidth(2, 180); // Usuario
    sheet.setColumnWidth(3, 110); // Rol
    sheet.setColumnWidth(4, 140); // Acción
    sheet.setColumnWidth(5, 200); // Residente
    sheet.setColumnWidth(6, 340); // Detalle
    SpreadsheetApp.flush();
  }
  return sheet;
}

function apiRegistrarCambio(usuario, rol, accion, residente, detalle) {
  if (!usuario || !accion) return { ok: false, error: 'Faltan campos' };
  const sheet = ensureHistorialCambiosSheet_();
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  sheet.appendRow([fecha, usuario, rol || '', accion, residente || '', detalle || '']);
  SpreadsheetApp.flush();
  return { ok: true };
}

function apiGetHistorial(residente, limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HISTORIAL_CAMBIOS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const last = sheet.getLastRow();
  const data = sheet.getRange(2, 1, last - 1, HISTORIAL_HEADERS.length).getValues();
  let result = data.map(function(row) {
    return {
      fecha: row[0] instanceof Date
        ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
        : String(row[0]),
      usuario: row[1], rol: row[2], accion: row[3],
      residente: row[4], detalle: row[5]
    };
  });
  if (residente) {
    result = result.filter(function(r){ return String(r.residente).toLowerCase() === String(residente).toLowerCase(); });
  }
  result.reverse(); // más reciente primero
  return result.slice(0, limit || 200);
}

function apiSetMinimos(residente, minPañal, minZalea, minAposito, minBombacha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_STOCK);
  if (!sheet) return { ok: false, error: 'Hoja STOCK_INSUMOS no encontrada. Ejecutar setupStockInsumos().' };
  const last = sheet.getLastRow();
  if (last >= 2) {
    const data = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(residente).toLowerCase()) {
        sheet.getRange(i + 2, 2, 1, 4).setValues([[
          Number(minPañal) || 0, Number(minZalea) || 0,
          Number(minAposito) || 0, Number(minBombacha) || 0
        ]]);
        SpreadsheetApp.flush();
        return { ok: true, action: 'updated' };
      }
    }
  }
  // No existe — agregar nueva fila
  sheet.appendRow([residente, Number(minPañal)||0, Number(minZalea)||0, Number(minAposito)||0, Number(minBombacha)||0]);
  SpreadsheetApp.flush();
  return { ok: true, action: 'inserted' };
}


// =====================================================
// API NUTRICIÓN
// =====================================================

var SHEET_NUT_DIETA     = 'NUTRICION_DIETA';
var SHEET_NUT_PESO      = 'NUTRICION_PESO';
var SHEET_NUT_REGISTROS = 'NUTRICION_REGISTROS';
var SHEET_NUT_NOTAS     = 'NUTRICION_NOTAS';

function ensureNutDietaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NUT_DIETA);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NUT_DIETA);
    sh.getRange(1,1,1,7).setValues([['Residente','TipoDieta','Restricciones','Observaciones','Indicaciones','MenuSemanal','MenuMensual','Updated','UpdatedBy']]);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#4A148C').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureNutPesoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NUT_PESO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NUT_PESO);
    sh.getRange(1,1,1,7).setValues([['Residente','Fecha','Semana','Peso_kg','Talla_cm','IMC','Usuario']]);
    sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#4A148C').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureNutRegistrosSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NUT_REGISTROS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NUT_REGISTROS);
    sh.getRange(1,1,1,10).setValues([['Residente','Mes','Dia','Desayuno','Almuerzo','Merienda','Cena','Suplemento','Obs','Usuario']]);
    sh.getRange(1,1,1,10).setFontWeight('bold').setBackground('#4A148C').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureNutNotasSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NUT_NOTAS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NUT_NOTAS);
    sh.getRange(1,1,1,4).setValues([['Residente','Fecha','Texto','Usuario']]);
    sh.getRange(1,1,1,4).setFontWeight('bold').setBackground('#4A148C').setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

function apiGetNutricionDieta(residente) {
  var sh = ensureNutDietaSheet_();
  if (sh.getLastRow() < 2) return {};
  var data = sh.getRange(2,1,sh.getLastRow()-1,7).getValues();
  for (var i=0; i<data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(residente).toLowerCase()) {
      return { tipo:data[i][1], restricciones:data[i][2], obs:data[i][3], indicaciones:data[i][4], updated:data[i][5]?String(data[i][5]):'', updatedBy:data[i][6] };
    }
  }
  return {};
}

function apiSetNutricionDieta(residente, p) {
  var sh = ensureNutDietaSheet_();
  var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  var vals = [residente, p.tipo||'', p.restricciones||'', p.obs||'', p.indicaciones||'', fecha, p.usuario||''];
  if (sh.getLastRow() >= 2) {
    var data = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (var i=0; i<data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(residente).toLowerCase()) {
        sh.getRange(i+2,1,1,7).setValues([vals]);
        SpreadsheetApp.flush();
        return { ok:true };
      }
    }
  }
  sh.appendRow(vals);
  SpreadsheetApp.flush();
  return { ok:true };
}

function apiGetNutricionPesos(residente) {
  var sh = ensureNutPesoSheet_();
  if (sh.getLastRow() < 2) return [];
  var data = sh.getRange(2,1,sh.getLastRow()-1,7).getValues();
  var result = [];
  data.forEach(function(row) {
    if (String(row[0]).toLowerCase() === String(residente).toLowerCase()) {
      result.push({ residente:row[0], fecha:row[1] instanceof Date ? Utilities.formatDate(row[1],Session.getScriptTimeZone(),'yyyy-MM-dd') : String(row[1]), semana:row[2], peso:row[3], talla:row[4], imc:row[5], usuario:row[6] });
    }
  });
  return result;
}

function apiAddNutricionPeso(residente, p) {
  var sh = ensureNutPesoSheet_();
  sh.appendRow([residente, p.fecha||'', p.semana||'', Number(p.peso)||0, Number(p.talla)||0, Number(p.imc)||0, p.usuario||'']);
  SpreadsheetApp.flush();
  return { ok:true };
}

function apiGetNutricionRegistros(residente, mes) {
  var sh = ensureNutRegistrosSheet_();
  if (sh.getLastRow() < 2) return [];
  var data = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
  var result = [];
  data.forEach(function(row) {
    if (String(row[0]).toLowerCase() === String(residente).toLowerCase() && String(row[1]) === String(mes)) {
      result.push({ residente:row[0], mes:row[1], dia:Number(row[2]), des:row[3], alm:row[4], mer:row[5], cen:row[6], sup:row[7]||'', obs:row[8]||'', usuario:row[9]||'' });
    }
  });
  return result;
}

function apiSetNutricionRegistro(residente, mes, dia, p) {
  var sh = ensureNutRegistrosSheet_();
  var diaNum = Number(dia);
  var vals = [residente, String(mes), diaNum, p.des||'', p.alm||'', p.mer||'', p.cen||'', p.sup||'', p.obs||'', p.usuario||''];
  if (sh.getLastRow() >= 2) {
    var data = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
    for (var i=0; i<data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(residente).toLowerCase() && String(data[i][1]) === String(mes) && Number(data[i][2]) === diaNum) {
        sh.getRange(i+2,1,1,10).setValues([vals]);
        SpreadsheetApp.flush();
        return { ok:true };
      }
    }
  }
  sh.appendRow(vals);
  SpreadsheetApp.flush();
  return { ok:true };
}

function apiGetNutricionNotas(residente) {
  var sh = ensureNutNotasSheet_();
  if (sh.getLastRow() < 2) return [];
  var data = sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  var result = [];
  data.forEach(function(row) {
    if (String(row[0]).toLowerCase() === String(residente).toLowerCase()) {
      result.push({ residente:row[0], fecha:String(row[1]), texto:row[2], usuario:row[3] });
    }
  });
  return result.reverse();
}

function apiAddNutricionNota(residente, texto, usuario, fecha) {
  var sh = ensureNutNotasSheet_();
  sh.appendRow([residente, fecha||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy'), texto||'', usuario||'']);
  SpreadsheetApp.flush();
  return { ok:true };
}

const SHEET_NUT_MENU = 'NUTRICION_MENU';

function ensureNutMenuSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NUT_MENU);
  if (!sh) { sh = ss.insertSheet(SHEET_NUT_MENU); sh.appendRow(['Semana','Data_JSON','Updated','UpdatedBy']); }
  return sh;
}

function apiGetNutricionMenu(semana) {
  if (!semana) return { ok:false, error:'semana requerida' };
  const sh = ensureNutMenuSheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(semana)) {
      try { return { ok:true, data: JSON.parse(rows[i][1]||'{}') }; } 
      catch(e) { return { ok:true, data:{} }; }
    }
  }
  return { ok:true, data:{} };
}

function apiSetNutricionMenu(semana, dataJson, usuario) {
  if (!semana || !dataJson) return { ok:false, error:'faltan datos' };
  const sh = ensureNutMenuSheet_();
  const rows = sh.getDataRange().getValues();
  const updated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(semana)) {
      sh.getRange(i+1,1,1,4).setValues([[semana, dataJson, updated, usuario||'']]);
      SpreadsheetApp.flush();
      return { ok:true };
    }
  }
  sh.appendRow([semana, dataJson, updated, usuario||'']);
  SpreadsheetApp.flush();
  return { ok:true };
}
