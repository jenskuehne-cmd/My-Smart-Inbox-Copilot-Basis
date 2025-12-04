/*** Archive.gs ***/

// Verschiebt alte/erledigte Zeilen von current activities → activity archive
function archiveOldRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!src) return;

  let dst = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);
  if (!dst) dst = ss.insertSheet(CONFIG.ARCHIVE_SHEET);

  const lastRow = src.getLastRow();
  if (lastRow <= 1) return;

  const lastCol = src.getLastColumn();
  const range = src.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();

  // Ziel: Header fuer Archiv sicherstellen, falls leer
  if (dst.getLastRow() < 1) {
    const header = src.getRange(1, 1, 1, lastCol).getValues();
    dst.getRange(1, 1, 1, lastCol).setValues(header);
  }

  const now = new Date();
  const cutoff = new Date(now.getTime() - CONFIG.ARCHIVE_OLDER_THAN_DAYS * 24 * 3600 * 1000);
  const statusSet = new Set(CONFIG.STATUS_ARCHIVE_LIST);

  const UPDATED_COL = 11; // K = Last Updated
  const STATUS_COL = 3;   // C = Status

  const toArchive = [];
  const toDeleteIdx = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const status = (row[STATUS_COL - 1] || '').toString().trim();
    const updated = row[UPDATED_COL - 1];

    const isOld = updated && updated instanceof Date && updated < cutoff;
    const matchStatus = statusSet.has(status);

    if (isOld || matchStatus) {
      toArchive.push(row);
      toDeleteIdx.push(i);
    }
  }

  if (toArchive.length === 0) {
    Logger.log('Nichts zu archivieren.');
    return;
  }

  dst.getRange(dst.getLastRow() + 1, 1, toArchive.length, lastCol).setValues(toArchive);

  // Von unten nach oben loeschen
  for (let k = toDeleteIdx.length - 1; k >= 0; k--) {
    const rowIndex = toDeleteIdx[k] + 2; // +2 wegen Header
    src.deleteRow(rowIndex);
  }

  Logger.log(`Archiviert: ${toArchive.length} Zeilen → "${CONFIG.ARCHIVE_SHEET}".`);
}
function archiveSelectedRows() {
  const SRC = CONFIG.SHEET_NAME || 'current activities';
  const DST = CONFIG.ARCHIVE_SHEET || 'activity archive';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== SRC) {
    ui.alert(`Bitte im Tab "${SRC}" die zu archivierenden Zeilen markieren.`);
    return;
  }

  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('Bitte mindestens eine Zeile markieren.');
    return;
  }

  // Zielblatt vorbereiten
  let dst = ss.getSheetByName(DST);
  if (!dst) dst = ss.insertSheet(DST);

  const lastCol = sheet.getLastColumn();

  // Header ins Archiv uebernehmen, falls leer
  if (dst.getLastRow() < 1) {
    const header = sheet.getRange(1, 1, 1, lastCol).getValues();
    dst.getRange(1, 1, 1, lastCol).setValues(header);
  }

  // Alle markierten Bereiche einsammeln → zu archivierendes Zeilen-Set (ohne Header-Zeile 1)
  const rowsToArchive = new Set();
  rangeList.getRanges().forEach(r => {
    const start = r.getRow();
    const end = start + r.getNumRows() - 1;
    for (let row = start; row <= end; row++) {
      if (row > 1) rowsToArchive.add(row); // Header auslassen
    }
  });

  if (rowsToArchive.size === 0) {
    ui.alert('Nur die Kopfzeile ist markiert – bitte Datenzeilen waehlen.');
    return;
  }

  // Zeileninhalte holen und ins Archiv schreiben (in stabiler Reihenfolge)
  const sortedRows = Array.from(rowsToArchive).sort((a, b) => a - b);
  const dataToArchive = [];
  sortedRows.forEach(row => {
    const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    dataToArchive.push(values);
  });

  // ans Archiv anhaengen
  if (dataToArchive.length > 0) {
    dst.getRange(dst.getLastRow() + 1, 1, dataToArchive.length, lastCol).setValues(dataToArchive);
  }

  // Jetzt im Source löschen – von unten nach oben, damit sich Indizes nicht verschieben
  for (let i = sortedRows.length - 1; i >= 0; i--) {
    sheet.deleteRow(sortedRows[i]);
  }

  ui.alert(`${dataToArchive.length} Zeile(n) nach "${DST}" verschoben und aus "${SRC}" entfernt.`);
}
function moveRowToArchive_(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!src || rowIndex <= 1) return;

  let dst = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);
  if (!dst) dst = ss.insertSheet(CONFIG.ARCHIVE_SHEET);

  const lastCol = src.getLastColumn();
  if (dst.getLastRow() < 1) {
    const header = src.getRange(1, 1, 1, lastCol).getValues();
    dst.getRange(1, 1, 1, lastCol).setValues(header);
  }
  const values = src.getRange(rowIndex, 1, 1, lastCol).getValues();
  dst.getRange(dst.getLastRow() + 1, 1, 1, lastCol).setValues(values);
  src.deleteRow(rowIndex);
}

function onSheetEdit_(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SHEET_NAME) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1) return;

    // 1) Last Updated (K)
    const touchCols = new Set([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    if (touchCols.has(col)) {
      sh.getRange(row, 11).setValue(new Date());
    }

    // 2) Korrekturen in Spalte A → Lernen
    if (col === 1) {
      recordCorrection_(sh, row, col, e.oldValue, e.value);
      learnFromProjectCorrection_(sh, row);
    }

    // 3) Status-Wechsel (Spalte C)
    if (col === 3) {
      const newStatus = (e.range.getValue() || '').toString().trim();

      if (newStatus === 'Done (sync)') {
        applyGmailLabelFromRow_(sh, row);
      }

      if ((CONFIG.STATUS_ARCHIVE_LIST || []).includes(newStatus)) {
        moveRowToArchive_(row);
        return;
      }

      recordCorrection_(sh, row, col, e.oldValue, e.value);
    }

    // 4) Priority Logging + Google Tasks Sync
    if (col === 4) { // Priority
      recordCorrection_(sh, row, col, e.oldValue, e.value);
      syncGoogleTaskForRow_(sh, row); // Prio-Aenderung kann Auswirkung auf Task haben
    }

    // 5) Task for Me (Spalte I) Logging + KI-Task-Suggestion + Google Tasks Sync
    if (col === 9) { // I = Task for Me
      recordCorrection_(sh, row, col, e.oldValue, e.value);

      const newVal = (e.range.getValue() || '').toString().trim();

      // Wenn jetzt "Yes": zuerst KI-Suggestion erzeugen (falls E leer)
      if (newVal === 'Yes') {
        ensureTaskSuggestionForRow_(sh, row);
      }

      // Danach Google Task anlegen / updaten / loeschen
      syncGoogleTaskForRow_(sh, row);
    }


  } catch (err) {
    Logger.log('onSheetEdit_ error: ' + err);
  }
}


function createOnEditTrigger() {
  ScriptApp.newTrigger('onSheetEdit_')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

