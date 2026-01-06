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

  // Prüfe, ob die Zeile wirklich Daten enthält (nicht leer)
  const firstColValue = src.getRange(rowIndex, 1).getValue();
  if (!firstColValue || firstColValue.toString().trim() === '') {
    Logger.log(`moveRowToArchive_: Zeile ${rowIndex} ist leer, überspringe Archivierung`);
    // Leere Zeile einfach löschen, ohne zu archivieren
    try {
      src.deleteRow(rowIndex);
    } catch (e) {
      Logger.log('Fehler beim Löschen leerer Zeile: ' + e);
    }
    return;
  }

  let dst = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);
  if (!dst) dst = ss.insertSheet(CONFIG.ARCHIVE_SHEET);

  const lastCol = src.getLastColumn();
  if (dst.getLastRow() < 1) {
    const header = src.getRange(1, 1, 1, lastCol).getValues();
    dst.getRange(1, 1, 1, lastCol).setValues(header);
  }
  
  // Hole alle Werte der Zeile
  const values = src.getRange(rowIndex, 1, 1, lastCol).getValues();
  
  // Prüfe, ob die Zeile wirklich Daten enthält (nicht nur leere Zellen)
  const hasData = values[0].some(cell => cell && cell.toString().trim() !== '');
  if (!hasData) {
    Logger.log(`moveRowToArchive_: Zeile ${rowIndex} enthält keine Daten, überspringe Archivierung`);
    try {
      src.deleteRow(rowIndex);
    } catch (e) {
      Logger.log('Fehler beim Löschen leerer Zeile: ' + e);
    }
    return;
  }
  
  // Schreibe in Archiv
  dst.getRange(dst.getLastRow() + 1, 1, 1, lastCol).setValues(values);
  
  // Lösche Zeile aus Quelle
  try {
    src.deleteRow(rowIndex);
    Logger.log(`moveRowToArchive_: Zeile ${rowIndex} erfolgreich archiviert und gelöscht`);
  } catch (e) {
    Logger.log(`moveRowToArchive_: Fehler beim Löschen der Zeile ${rowIndex}: ${e}`);
    // Versuche es nochmal mit clearContent statt deleteRow
    try {
      src.getRange(rowIndex, 1, 1, lastCol).clearContent();
      Logger.log(`moveRowToArchive_: Zeile ${rowIndex} Inhalt gelöscht (als Fallback)`);
    } catch (e2) {
      Logger.log(`moveRowToArchive_: Auch Fallback fehlgeschlagen: ${e2}`);
    }
  }
}

function onSheetEdit_(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SHEET_NAME) return;

    const startRow = e.range.getRow();
    const endRow = e.range.getLastRow();
    const col = e.range.getColumn();
    if (startRow <= 1) return;

    // Prüfe, ob es eine Bulk-Änderung ist (mehrere Zeilen)
    const isBulkEdit = endRow > startRow;
    
    // Wenn Bulk-Edit in Spalte C (Status) → verarbeite alle Zeilen
    if (isBulkEdit && col === 3) {
      const values = e.range.getValues();
      const rowsToArchive = [];
      
      for (let i = 0; i < values.length; i++) {
        const row = startRow + i;
        if (row <= 1) continue; // Header überspringen
        
        const newStatus = (values[i][0] || '').toString().trim();
        if (!newStatus) continue;
        
        // WICHTIG: Gmail-Operationen ZUERST ausführen, bevor die Zeile archiviert wird!
        if (newStatus === 'Done (sync)') {
          // Label setzen + INBOX entfernen
          applyGmailLabelFromRow_(sh, row);
        } else if (newStatus === 'Done (train only)') {
          // Nur INBOX entfernen (in Papierkorb verschieben), aber kein Label setzen
          removeFromInboxOnly_(sh, row);
        } else if (newStatus === 'Action Required') {
          // Action-Vorschlag generieren (falls noch nicht vorhanden)
          generateActionSuggestionForRow_(sh, row);
        }
        
        // Archivierung NACH Gmail-Operationen
        if ((CONFIG.STATUS_ARCHIVE_LIST || []).includes(newStatus)) {
          rowsToArchive.push(row);
        }
      }
      
      // Archiviere alle Zeilen (von hinten nach vorne, damit Indizes stimmen)
      for (let i = rowsToArchive.length - 1; i >= 0; i--) {
        moveRowToArchive_(rowsToArchive[i]);
      }
      
      Logger.log(`Bulk-Edit: ${values.length} Zeilen verarbeitet, ${rowsToArchive.length} archiviert`);
      return; // Bulk-Edit abgeschlossen
    }

    // Einzelne Zeile verarbeiten (wie bisher)
    const row = startRow;

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

    // 3) Status-Wechsel (Spalte C) - einzelne Zeile
    if (col === 3) {
      const newStatus = (e.range.getValue() || '').toString().trim();

      // WICHTIG: Gmail-Operationen ZUERST ausführen, bevor die Zeile archiviert wird!
      if (newStatus === 'Done (sync)') {
        // Label setzen + INBOX entfernen
        applyGmailLabelFromRow_(sh, row);
      } else if (newStatus === 'Done (train only)') {
        // Nur INBOX entfernen (in Papierkorb verschieben), aber kein Label setzen
        // Learning passiert bereits durch Spalte A-Änderung
        removeFromInboxOnly_(sh, row);
      } else if (newStatus === 'Action Required') {
        // Action-Vorschlag generieren (falls noch nicht vorhanden)
        generateActionSuggestionForRow_(sh, row);
      }

      // Korrektur aufzeichnen (vor Archivierung)
      recordCorrection_(sh, row, col, e.oldValue, e.value);

      // Archivierung NACH Gmail-Operationen
      if ((CONFIG.STATUS_ARCHIVE_LIST || []).includes(newStatus)) {
        moveRowToArchive_(row);
        return;
      }
    }

    // 4) Priority Logging + Google Tasks Sync + Learning
    if (col === 4) { // Priority
      recordCorrection_(sh, row, col, e.oldValue, e.value);
      
      // Learning: Lerne aus Prioritätsänderungen
      const newPriority = (e.value || '').toString().trim();
      if (newPriority && ['Low', 'Medium', 'High', 'Urgent'].includes(newPriority)) {
        learnFromPriorityCorrection_(sh, row, newPriority);
      }
      
      syncGoogleTaskForRow_(sh, row); // Prio-Aenderung kann Auswirkung auf Task haben
    }

    // 5) Task for Me (Spalte I) Logging + KI-Task-Suggestion + Google Tasks Sync + Learning
    if (col === 9) { // I = Task for Me
      recordCorrection_(sh, row, col, e.oldValue, e.value);

      const newVal = (e.range.getValue() || '').toString().trim();
      const oldVal = (e.oldValue || '').toString().trim();

      // Learning: Wenn von "Unsure" auf "No" geändert → lerne, dass solche Mails keine Tasks sind
      if ((oldVal === 'Unsure' || oldVal === '') && newVal === 'No') {
        learnFromTaskForMeCorrection_(sh, row, 'No');
      }
      // Learning: Wenn von "Unsure" auf "Yes" geändert → lerne, dass solche Mails Tasks sind
      if ((oldVal === 'Unsure' || oldVal === '') && newVal === 'Yes') {
        learnFromTaskForMeCorrection_(sh, row, 'Yes');
      }

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

