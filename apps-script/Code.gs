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

function getStudents_() {
  const sheet = ss_().getSheetByName(STUDENTS_TAB);
  if (!sheet) throw new Error('Aba "' + STUDENTS_TAB + '" não encontrada.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values
    .map(function (r) { return String(r[0] || '').trim(); })
    .filter(function (n) { return n.length > 0; });
}

function ensureMonthSheet_(month) {
  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error('Mês inválido. Use o formato AAAA-MM, ex: 2026-05.');
  }
  const ss = ss_();
  let sheet = ss.getSheetByName(month);
  if (!sheet) {
    sheet = ss.insertSheet(month);
    sheet.getRange(1, 1).setValue('Aluno').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);
  }
  return sheet;
}

/**
 * Garante que todos os alunos da aba "Alunos" existem na aba do mês,
 * sem remover quem já está lá. Mantém a ordem da aba "Alunos" para novos.
 */
function syncStudentsToMonth_(sheet) {
  const students = getStudents_();
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

function headerToDateStr_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function getMonth_(month) {
  const sheet = ensureMonthSheet_(month);
  syncStudentsToMonth_(sheet);

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

  const sheet = ss_().getSheetByName(STUDENTS_TAB);
  if (!sheet) throw new Error('Aba "' + STUDENTS_TAB + '" não encontrada.');

  const lastRow = sheet.getLastRow();
  let existing = [];
  if (lastRow >= 2) {
    existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
      .map(function (r) { return String(r[0] || '').trim().toLowerCase(); });
  }
  if (existing.indexOf(clean.toLowerCase()) !== -1) {
    throw new Error('Já existe um aluno com esse nome.');
  }

  sheet.getRange(lastRow + 1, 1).setValue(clean);
  return { added: clean, row: lastRow + 1 };
}
