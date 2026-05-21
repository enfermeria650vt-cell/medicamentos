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
const NCOL = 13;
const HEADERS = ['Residente', 'FileID', 'Modif. Drive', 'Última Sync', 'Hash',
                 'Diagnóstico', 'HTML', 'Texto', 'HashMed', 'HashHor', 'HashEvo',
                 'Modificado', 'FechaModif'];

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
  sheet.hideColumns(2);
  sheet.hideColumns(5);
  sheet.hideColumns(7);
  sheet.hideColumns(8);
  sheet.hideColumns(9, 3);
}

function md5(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text || '');
  return bytes.map(function(b){ return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

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
        modificado: row[11], fechaModif: row[12]
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

    // Optimización: si el archivo NO cambió en Drive (misma fecha de modificación)
    // y la fila ya tiene todos los hashes, saltar sin convertir. La conversión a
    // Google Doc es lo lento — así las corridas sin cambios son de segundos y el
    // activador puede correr cada pocos minutos sin sobrecargarse.
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
    }

    const html = renderDocToHTML(doc);
    const diagMatch = text.match(/Diagn[óo]stico:\s*([^\n]+)/i);
    const diagnostico = diagMatch ? diagMatch[1].trim() : '';

    const row = [fileName, fileId, modifTime, now, hash, diagnostico, html, text,
                 hMed, hHor, hEvo, modificado, fechaModif];
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

function etiquetaSemana(d) {
  const date = new Date(d.getTime());
  date.setHours(0, 0, 0, 0);
  date.setDate(date.getDate() + 3 - ((date.getDay() + 6) % 7));
  const week1 = new Date(date.getFullYear(), 0, 4);
  const semana = 1 + Math.round(((date - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
  return date.getFullYear() + '-S' + String(semana).padStart(2, '0');
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

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  try {
    let payload;
    switch (action) {
      case 'getIndicaciones': payload = apiGetIndicaciones(); break;
      case 'getIndicacion':   payload = apiGetIndicacion(e.parameter.id); break;
      case 'syncNow':
        if (!isAuthorized(e.parameter.email)) return jsonError('No autorizado');
        payload = { sync: syncIndicaciones() };
        break;
      case 'respaldoNow':
        if (!isAuthorized(e.parameter.email)) return jsonError('No autorizado');
        payload = { respaldo: respaldoSemanalIndicaciones() };
        break;
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
    diagnostico: row[5], html: row[6], texto: row[7],
    modificado: row[11] || '',
    fechaModif: row[12] instanceof Date ? row[12].toISOString() : (row[12] || '')
  };
}

function isAuthorized(email) {
  return email && String(email).toLowerCase() === AUTHORIZED_EMAIL.toLowerCase();
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    const f = t.getHandlerFunction();
    if (f === 'syncIndicaciones' || f === 'respaldoSemanalIndicaciones') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncIndicaciones').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('respaldoSemanalIndicaciones').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();
}
