/**
 * EBD - Backend de chamada (Google Apps Script)
 *
 * Como funciona o CORS:
 *  - GET é "simple request", o navegador não faz preflight.
 *  - POST com Content-Type: text/plain também é "simple request".
 *  - Por isso, o frontend envia JSON.stringify(payload) no corpo do POST com
 *    Content-Type: text/plain. Aqui dentro fazemos JSON.parse(e.postData.contents).
 *  - NÃO use application/json no frontend: isso obrigaria preflight OPTIONS,
 *    e Apps Script não responde OPTIONS -> erro de CORS.
 */

const SHEET_ID = '1-SvEEVvBPRvBXspTRMA5BxhPJEbmNZA3MsDYR_BCtLU';
const STUDENTS_TAB = 'Alunos';
const VALID_STATUS = ['P', 'A', 'J', 'F', 'FH', ''];

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Servidor ocupado, tente novamente.' });
  }

  try {
    const params = readParams_(e);
    const action = String(params.action || '').trim();
    let data;

    switch (action) {
      case 'ping':
        data = { pong: true, time: new Date().toISOString() };
        break;
      case 'getStudents':
        data = { students: getStudents_() };
        break;
      case 'getMonth':
        data = getMonth_(String(params.month || ''));
        break;
      case 'saveAttendance':
        data = saveAttendance_(
          String(params.month || ''),
          String(params.date || ''),
          params.attendance || {}
        );
        break;
      case 'addStudent':
        data = addStudent_(String(params.name || ''));
        break;
      case 'sortStudents':
        data = sortAllStudentLists_();
        break;
      default:
        return jsonOut_({ ok: false, error: 'Ação desconhecida: ' + action });
    }

    return jsonOut_({ ok: true, data: data });
  } catch (err) {
    return jsonOut_({ ok: false, error: err && err.message ? err.message : String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function readParams_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (_) {
      // fallback se vier urlencoded ou algo diferente
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 *  Lógica de negócio
 * ============================================================ */

function ss_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function compareNames_(a, b) {
  return String(a).localeCompare(String(b), 'pt-BR', { sensitivity: 'base' });
}

/** Lista de alunos a partir de uma instância já aberta (evita openById extra). */
function getStudentsFromSpreadsheet_(ss) {
  const sheet = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) throw new Error('Aba "' + STUDENTS_TAB + '" não encontrada.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values
    .map(function (r) { return String(r[0] || '').trim(); })
    .filter(function (n) { return n.length > 0; })
    .sort(compareNames_);
}

function getStudents_() {
  return getStudentsFromSpreadsheet_(ss_());
}

/**
 * Reescreve a aba Alunos em ordem alfabética (pt-BR).
 * @return {string[]} lista ordenada
 */
function sortStudentsSheet_(ss) {
  const sheet = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) throw new Error('Aba "' + STUDENTS_TAB + '" não encontrada.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const names = values
    .map(function (r) { return String(r[0] || '').trim(); })
    .filter(function (n) { return n.length > 0; })
    .sort(compareNames_);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }
  if (names.length > 0) {
    sheet.getRange(2, 1, names.length, 1).setValues(names.map(function (n) { return [n]; }));
  }
  return names;
}

/**
 * Reordena as linhas da aba do mês pela coluna Aluno, preservando as colunas de data.
 */
function sortMonthSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return 0;

  const width = Math.max(lastCol, 1);
  const body = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  const rows = body.filter(function (row) {
    return String(row[0] || '').trim().length > 0;
  });
  rows.sort(function (a, b) {
    return compareNames_(String(a[0] || '').trim(), String(b[0] || '').trim());
  });

  sheet.getRange(2, 1, lastRow - 1, width).clearContent();
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, width).setValues(rows);
  }
  return rows.length;
}

/** Ordena aba Alunos e todas as abas de mês (AAAA-MM). */
function sortAllStudentLists_() {
  const ss = ss_();
  const students = sortStudentsSheet_(ss);
  const sheets = ss.getSheets();
  let monthsSorted = 0;
  sheets.forEach(function (sheet) {
    const name = sheet.getName();
    if (/^\d{4}-\d{2}$/.test(name)) {
      sortMonthSheet_(sheet);
      monthsSorted++;
    }
  });
  return { students: students, monthsSorted: monthsSorted };
}

function ensureMonthSheetForSpreadsheet_(ss, month) {
  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error('Mês inválido. Use o formato AAAA-MM, ex: 2026-05.');
  }
  let sheet = ss.getSheetByName(month);
  if (!sheet) {
    sheet = ss.insertSheet(month);
    sheet.getRange(1, 1).setValue('Aluno').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }
  return sheet;
}

function ensureMonthSheet_(month) {
  return ensureMonthSheetForSpreadsheet_(ss_(), month);
}

/**
 * Garante que todos os alunos da aba "Alunos" existem na aba do mês,
 * sem remover quem já está lá.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} students lista já lida da aba Alunos (reuso em getMonth_)
 */
function syncStudentsToMonthWithList_(sheet, students) {
  const lastRow = sheet.getLastRow();
  let existing = [];
  if (lastRow >= 2) {
    existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(function (r) { return String(r[0] || '').trim(); });
  }
  const existingSet = {};
  existing.forEach(function (n) { existingSet[n.toLowerCase()] = true; });

  const toAdd = students.filter(function (s) {
    return !existingSet[s.toLowerCase()];
  });
  if (toAdd.length > 0) {
    sheet.getRange(existing.length + 2, 1, toAdd.length, 1)
      .setValues(toAdd.map(function (s) { return [s]; }));
  }
}

function syncStudentsToMonth_(sheet) {
  syncStudentsToMonthWithList_(sheet, getStudents_());
}

function headerToDateStr_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function getMonth_(month) {
  const ss = ss_();
  const sheet = ensureMonthSheetForSpreadsheet_(ss, month);
  const studentsList = getStudentsFromSpreadsheet_(ss);
  syncStudentsToMonthWithList_(sheet, studentsList);
  sortMonthSheet_(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return { month: month, dates: [], students: [], grid: {} };
  }

  const header = lastCol >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  const dateCols = []; // { col: 1-based index, date: 'YYYY-MM-DD' }
  for (let i = 1; i < header.length; i++) {
    const ds = headerToDateStr_(header[i]);
    if (ds) dateCols.push({ col: i + 1, date: ds });
  }

  const body = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 1)).getValues();
  const students = [];
  const grid = {};

  body.forEach(function (row) {
    const name = String(row[0] || '').trim();
    if (!name) return;
    students.push(name);
    const rec = {};
    dateCols.forEach(function (d) {
      rec[d.date] = String(row[d.col - 1] || '').trim();
    });
    grid[name] = rec;
  });

  return {
    month: month,
    dates: dateCols.map(function (d) { return d.date; }),
    students: students,
    grid: grid
  };
}

function saveAttendance_(month, date, attendance) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    throw new Error('Data inválida. Use AAAA-MM-DD.');
  }
  if (date.substring(0, 7) !== month) {
    throw new Error('A data não pertence ao mês informado.');
  }
  if (!attendance || typeof attendance !== 'object') {
    throw new Error('Payload de chamada inválido.');
  }

  const sheet = ensureMonthSheet_(month);
  syncStudentsToMonth_(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // Localiza ou cria a coluna da data
  let dateCol = -1;
  if (lastCol >= 2) {
    const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let i = 1; i < header.length; i++) {
      if (headerToDateStr_(header[i]) === date) {
        dateCol = i + 1;
        break;
      }
    }
  }
  if (dateCol === -1) {
    dateCol = Math.max(lastCol, 1) + 1;
    sheet.getRange(1, dateCol).setValue(date).setFontWeight('bold');
  }

  // Monta a coluna inteira, na ordem dos alunos da própria aba do mês
  const namesCol = sheet.getRange(2, 1, Math.max(lastRow - 1, 0), 1).getValues();
  const values = namesCol.map(function (r) {
    const name = String(r[0] || '').trim();
    if (!name) return [''];
    const raw = attendance[name];
    const status = raw == null ? '' : String(raw).trim().toUpperCase();
    if (VALID_STATUS.indexOf(status) === -1) {
      throw new Error('Status inválido "' + status + '" para o aluno ' + name);
    }
    return [status];
  });

  if (values.length > 0) {
    sheet.getRange(2, dateCol, values.length, 1).setValues(values);
  }

  return {
    saved: true,
    month: month,
    date: date,
    rows: values.length,
    column: dateCol
  };
}

function addStudent_(name) {
  const clean = String(name || '').trim();
  if (!clean) throw new Error('Nome vazio.');
  if (clean.length > 80) throw new Error('Nome muito longo.');

  const ss = ss_();
  const sheet = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) throw new Error('Aba "' + STUDENTS_TAB + '" não encontrada.');

  const lastRow = sheet.getLastRow();
  let existing = [];
  if (lastRow >= 2) {
    existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(function (r) { return String(r[0] || '').trim(); })
      .filter(function (n) { return n.length > 0; });
  }

  const lower = clean.toLowerCase();
  for (let i = 0; i < existing.length; i++) {
    if (existing[i].toLowerCase() === lower) {
      throw new Error('Já existe um aluno com esse nome.');
    }
  }

  existing.push(clean);
  existing.sort(compareNames_);
  const row = existing.indexOf(clean) + 2; // +1 header, +1 1-based

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }
  sheet.getRange(2, 1, existing.length, 1).setValues(existing.map(function (n) { return [n]; }));

  return { added: clean, row: row };
}
