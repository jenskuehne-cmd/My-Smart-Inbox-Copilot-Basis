/*** Menu.gs ***/
// Baut ein Menue "Mail Assistant" im Sheet.

function onOpen() {
  // Menü erstellen
  SpreadsheetApp.getUi()
    .createMenu('Mail Assistant')
    .addItem('Verarbeite: Heute', 'menu_processToday')
    .addItem('Verarbeite: Letzte 7 Tage', 'menu_processLast7')
    .addItem('Verarbeite: Datumsbereich…', 'menu_processDateRange')
    .addSeparator()
    .addItem('Test: Gemini API-Key', 'menu_testAPIKey')
    .addItem('API-Key Status anzeigen', 'menu_showAPIKeyStatus')
    .addItem('📋 Logs anzeigen (Hilfe)', 'menu_showLogsHelp')
    .addSeparator()
    .addItem('Setup: Dropdowns/Validierung', 'menu_setupValidations')
    .addItem('Archiviere markierte Zeile(n)', 'menu_archiveSelection')
    .addItem('Archivieren: Alte Zeilen', 'menu_archiveOldRows')
    .addSeparator()
    .addItem('AI: Response für markierte Zeile erstellen', 'menu_generateResponseForRow')
    .addItem('AI: Summary für markierte Zeile erstellen', 'menu_generateSummaryForRow')
    .addSeparator()
    .addItem('Vorlage: Project Map erzeugen', 'menu_createProjectMapTemplate')
    .addItem('Vorlage: Konfiguration erzeugen', 'menu_createConfigurationTemplate')
    .addItem('ML: Lernen aus Korrekturen', 'menu_learnFromCorrections')
    .addSeparator()
    .addItem('🔧 Fix: "Done (sync)" Mails aus Posteingang entfernen', 'menu_fixDoneSyncInbox')
    .addItem('🔄 Verarbeite alle "Done (sync)" Zeilen', 'menu_processDoneSyncRows')
    .addItem('🔄 Verarbeite alle "Done (train only)" Zeilen', 'menu_processDoneTrainOnlyRows')
    .addToUi();
  
  // Automatisch zur letzten Zeile im "current activities" Tab springen
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // Aktiviere das Sheet und springe zur letzten Zeile
        ss.setActiveSheet(sheet);
        // Aktiviere die letzte Zeile (Spalte A, damit die ganze Zeile sichtbar ist)
        sheet.setActiveRange(sheet.getRange(lastRow, 1));
        // Optional: Scroll zur letzten Zeile (setActiveRange macht das normalerweise automatisch)
        sheet.setActiveRange(sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()));
      }
    }
  } catch (e) {
    Logger.log('onOpen: Fehler beim Springen zur letzten Zeile: ' + e);
    // Fehler ignorieren, damit das Menü trotzdem erstellt wird
  }
}

function menu_setupValidations(){ ensureValidations_(); }

/**
 * Erstellt eine AI-Response für die markierte Zeile basierend auf dem gesamten Thread-Kontext
 */
function menu_generateResponseForRow() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Fehler', 'Sheet "current activities" nicht gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const selection = sheet.getActiveRange();
    if (!selection) {
      SpreadsheetApp.getUi().alert('Fehler', 'Bitte markieren Sie eine Zeile in der Tabelle.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const row = selection.getRow();
    if (row < 2) {
      SpreadsheetApp.getUi().alert('Fehler', 'Bitte markieren Sie eine Datenzeile (nicht die Kopfzeile).', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    if (!messageId) {
      SpreadsheetApp.getUi().alert('Fehler', 'Keine Message-ID gefunden in dieser Zeile.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // API-Key prüfen
    const apiTest = testGeminiAPIKey();
    if (!apiTest.valid || apiTest.quotaExceeded) {
      SpreadsheetApp.getUi().alert('API-Key Problem', apiTest.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Response mit Thread-Kontext generieren
    const response = generateResponseWithThreadContext_(messageId, row);
    
    if (response && response.trim()) {
      // In Spalte J (Reply-Suggestion) schreiben
      sheet.getRange(row, 10).setValue(response);
      Logger.log(`AI-Response erfolgreich generiert und in Spalte J geschrieben`);
      // Keine Alert-Box bei Erfolg - der Text ist in Spalte J sichtbar
    } else {
      SpreadsheetApp.getUi().alert('Fehler', 'Konnte keine Response generieren.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    Logger.log('menu_generateResponseForRow error: ' + e);
    SpreadsheetApp.getUi().alert('Fehler', 'Fehler beim Generieren der Response: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generiert eine AI Summary für die markierte Zeile nachträglich
 */
function menu_generateSummaryForRow() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Fehler', 'Sheet "current activities" nicht gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const selection = sheet.getActiveRange();
    if (!selection) {
      SpreadsheetApp.getUi().alert('Fehler', 'Bitte markieren Sie eine Zeile in der Tabelle.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const row = selection.getRow();
    if (row < 2) {
      SpreadsheetApp.getUi().alert('Fehler', 'Bitte markieren Sie eine Datenzeile (nicht die Kopfzeile).', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    if (!messageId) {
      SpreadsheetApp.getUi().alert('Fehler', 'Keine Message-ID gefunden in dieser Zeile.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // API-Key prüfen
    const apiTest = testGeminiAPIKey();
    if (!apiTest.valid || apiTest.quotaExceeded) {
      SpreadsheetApp.getUi().alert('API-Key Problem', apiTest.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Hole E-Mail-Nachricht und Body-Text
    let bodyText = '';
    try {
      const msg = GmailApp.getMessageById(messageId);
      if (!msg) {
        SpreadsheetApp.getUi().alert('Fehler', 'E-Mail-Nachricht nicht gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      bodyText = getBestBody_(msg);
      if (!bodyText || bodyText.trim().length === 0) {
        SpreadsheetApp.getUi().alert('Fehler', 'E-Mail-Body ist leer.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
    } catch (e) {
      Logger.log('Fehler beim Laden der E-Mail: ' + e);
      SpreadsheetApp.getUi().alert('Fehler', 'Konnte E-Mail nicht laden: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // AI Summary generieren
    Logger.log(`Generiere AI Summary für Zeile ${row}, Message-ID: ${messageId}`);
    Logger.log(`Body-Text Länge: ${bodyText.length} Zeichen`);
    
    const summary = getAISummarySmart(bodyText);
    Logger.log(`Summary-Ergebnis (${summary ? summary.length : 0} Zeichen): "${summary ? summary.substring(0, 200) : 'null'}"`);
    
    // Prüfe auf Fehler-Strings
    const errorPatterns = [
      'Keine AI-Summary',
      'Could not generate',
      'Error generating',
      'verfügbar (Error)',
      'verfügbar (Netzwerk',
      'verfügbar (Limit)'
    ];
    
    const isError = errorPatterns.some(pattern => summary && summary.includes(pattern));
    
    if (summary && summary.trim() && !isError) {
      // In Spalte H (AI Summary) schreiben
      sheet.getRange(row, 8).setValue(summary);
      Logger.log(`AI Summary erfolgreich generiert: ${summary.substring(0, 100)}...`);
      // Keine Alert-Box bei Erfolg - der Text ist in Spalte H sichtbar
    } else {
      // Detaillierte Fehlermeldung
      let errorMsg = 'Konnte keine Summary generieren.\n\n';
      if (isError) {
        errorMsg += `API-Antwort: "${summary}"\n\n`;
      } else if (!summary || !summary.trim()) {
        errorMsg += 'API hat leere Antwort zurückgegeben.\n\n';
      }
      errorMsg += 'Bitte:\n';
      errorMsg += '1. Prüfen Sie die API-Konfiguration\n';
      errorMsg += '2. Versuchen Sie es erneut\n';
      errorMsg += '3. Prüfen Sie die Logs für Details';
      
      SpreadsheetApp.getUi().alert('Fehler', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
      Logger.log(`AI Summary-Generierung fehlgeschlagen für Message-ID: ${messageId}, Summary: "${summary}"`);
    }
  } catch (e) {
    Logger.log('menu_generateSummaryForRow error: ' + e);
    SpreadsheetApp.getUi().alert('Fehler', 'Fehler beim Generieren der Summary: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menu_testAPIKey() {
  const ui = SpreadsheetApp.getUi();
  const result = testGeminiAPIKey();
  
  if (result.valid && !result.quotaExceeded) {
    ui.alert('✅ API-Key Test erfolgreich', result.message, ui.ButtonSet.OK);
  } else if (result.quotaExceeded) {
    ui.alert('⚠️ API-Limit erreicht', 
      'Der API-Key ist gültig, aber das Free-Tier-Limit wurde erreicht.\n\n' + 
      result.message + '\n\n' +
      'Sie können die E-Mail-Verarbeitung trotzdem starten, aber ohne KI-Funktionen.', 
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ API-Key Test fehlgeschlagen', result.message, ui.ButtonSet.OK);
  }
}

function menu_showAPIKeyStatus() {
  const ui = SpreadsheetApp.getUi();
  const usePortkey = getProp_('USE_PORTKEY', 'false').toLowerCase() === 'true';
  
  let message = '';
  
  // Erkläre Config-ID Warnung
  const portkeyConfigId = getProp_('PORTKEY_CONFIG_ID', '');
  const portkeyConfigJson = getProp_('PORTKEY_CONFIG_JSON', '');
  let configIdNote = '';
  if (usePortkey && !portkeyConfigId && !portkeyConfigJson) {
    configIdNote = '\n\n⚠️ HINWEIS zur Config-ID:\n' +
                   'Die Warnung "Keine Config-ID gesetzt" bedeutet, dass Sie keine Portkey-Konfiguration verwenden.\n' +
                   'Das ist OK, wenn Sie direkt mit Portkey arbeiten.\n' +
                   'Falls Sie eine Portkey-Konfiguration haben (z.B. mit Fallback-Strategien), können Sie:\n' +
                   '- PORTKEY_CONFIG_ID: Ihre Config-ID (z.B. "pc-defaul-6969bd")\n' +
                   '- ODER PORTKEY_CONFIG_JSON: Ihre Config als JSON-String\n' +
                   'in Script Properties setzen.\n' +
                   'Aktuell funktioniert es auch ohne Config-ID.';
  }
  
  if (usePortkey) {
    const portkeyKey = getProp_('PORTKEY_API_KEY', '');
    const baseUrl = getProp_('PORTKEY_BASE_URL', 'https://api.portkey.ai/v1');
    const provider = getProp_('PORTKEY_PROVIDER', 'google');
    const model = getProp_('PORTKEY_MODEL', 'gemini-2.0-flash');
    
    const configId = getProp_('PORTKEY_CONFIG_ID', '');
    const configJson = getProp_('PORTKEY_CONFIG_JSON', '');
    
    if (!portkeyKey) {
      message = '❌ Portkey ist aktiviert, aber PORTKEY_API_KEY fehlt\n\n' +
                'Bitte setzen Sie in Script Properties:\n' +
                '- PORTKEY_API_KEY: Ihr Portkey API-Key (aus https://app.portkey.ai/api-keys)\n' +
                '- PORTKEY_BASE_URL (optional): ' + baseUrl + '\n\n' +
                'UND eine der folgenden Optionen:\n' +
                'Option 1: Config-ID (EMPFOHLEN - wenn Sie eine Config in Portkey haben)\n' +
                '  - PORTKEY_CONFIG_ID: Ihre Config-ID aus Portkey\n' +
                '    (z.B. pc-defaul-6969bd)\n\n' +
                'Option 2: Config als JSON\n' +
                '  - PORTKEY_CONFIG_JSON: Ihre Config als JSON-String\n' +
                '    (nur wenn keine Config-ID verfügbar)\n\n' +
                'Option 3: Vertex AI direkt (mit separatem Vertex API-Key)\n' +
                '  - PORTKEY_PROVIDER: vertex-ai\n' +
                '  - PORTKEY_VERTEX_API_KEY: Ihr Vertex API-Key aus Portkey\n' +
                '    ODER PORTKEY_VERTEX_VIRTUAL_KEY: Vertex Virtual Key ID\n' +
                '  - PORTKEY_VERTEX_PROJECT_ID (optional): GCP Project ID\n' +
                '  - PORTKEY_VERTEX_REGION (optional): us-central1 (Standard)\n' +
                '  - PORTKEY_MODEL: gemini-pro (empfohlen für Vertex AI, besonders in EU-Regionen)\n\n' +
                'Option 4: Google Gemini direkt\n' +
                '  - PORTKEY_PROVIDER: google\n' +
                '  - GEMINI_API_KEY: Ihr Gemini API-Key\n' +
                '  - PORTKEY_MODEL: gemini-2.0-flash (Standard)';
    } else {
      const masked = portkeyKey.length > 10 
        ? portkeyKey.substring(0, 4) + '...' + portkeyKey.substring(portkeyKey.length - 4)
        : '****';
      message = '✅ Portkey-Modus aktiviert\n\n' +
                `Portkey Key (maskiert): ${masked}\n` +
                `Base URL: ${baseUrl}\n`;
      
      if (configId) {
        message += `✅ Config-ID: ${configId}\n` +
                   `(Verwendet Ihre Portkey-Config mit Fallback-Strategie)\n`;
      } else if (configJson) {
        message += `Config: JSON konfiguriert (${configJson.length} Zeichen)\n`;
      } else {
        message += `Provider: ${provider}\n` +
                   `Modell: ${model}\n` +
                   `⚠️ Keine Config-ID gesetzt - verwenden Sie PORTKEY_CONFIG_ID für beste Ergebnisse\n`;
      }
      
      // Erkläre Config-ID Warnung
      if (!configId && !configJson) {
        message += '\n\n💡 HINWEIS zur Config-ID:\n' +
                   'Die Warnung "Keine Config-ID gesetzt" ist normal, wenn Sie direkt mit Portkey arbeiten.\n' +
                   'Eine Config-ID ist nur nötig, wenn Sie:\n' +
                   '- Fallback-Strategien verwenden möchten\n' +
                   '- Verschiedene Provider automatisch wechseln möchten\n' +
                   '- Retry-Logik konfigurieren möchten\n\n' +
                   'Aktuell funktioniert es auch ohne Config-ID.\n' +
                   'Falls Sie eine Config-ID haben, setzen Sie PORTKEY_CONFIG_ID in Script Properties.';
      }
      
      message += '\nUm zu direktem Gemini-Modus zu wechseln:\n' +
                'Setzen Sie USE_PORTKEY auf "false" in Script Properties';
    }
  } else {
    const apiKey = getProp_('GEMINI_API_KEY', '');
    
    if (!apiKey || apiKey.trim() === '') {
      message = '❌ Kein API-Key hinterlegt\n\n' +
                'Bitte setzen Sie den GEMINI_API_KEY in den Script Properties:\n\n' +
                '1. Extensions → Apps Script\n' +
                '2. ⚙️ Projekt-Einstellungen (Zahnrad)\n' +
                '3. Script Properties\n' +
                '4. Eigenschaft hinzufügen:\n' +
                '   - Eigenschaft: GEMINI_API_KEY\n' +
                '   - Wert: Ihr Gemini API-Key\n\n' +
                'ODER aktivieren Sie Portkey:\n' +
                '   - USE_PORTKEY: true\n' +
                '   - PORTKEY_API_KEY: Ihr Portkey Key';
    } else {
      // Zeige nur die ersten und letzten Zeichen aus Sicherheitsgründen
      const masked = apiKey.length > 10 
        ? apiKey.substring(0, 4) + '...' + apiKey.substring(apiKey.length - 4)
        : '****';
      message = '✅ Direkter Gemini-Modus aktiviert\n\n' +
                `Key (maskiert): ${masked}\n` +
                `Länge: ${apiKey.length} Zeichen\n\n` +
                'Um zu Portkey-Modus zu wechseln:\n' +
                'Setzen Sie USE_PORTKEY auf "true" in Script Properties';
    }
  }
  
  ui.alert('API-Konfiguration Status', message, ui.ButtonSet.OK);
}


function menu_archiveSelection() {
  archiveSelectedRows(); // ruft die Funktion unten auf
}


function menu_processToday() {
  processEmailsToday();
}

function menu_processLast7() {
  processEmailsLastNDays(7);
}

function menu_processDateRange() {
  const ui = SpreadsheetApp.getUi();
  const start = ui.prompt('Startdatum', 'Bitte eingeben (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (start.getSelectedButton() !== ui.Button.OK) return;
  const end = ui.prompt('Enddatum', 'Bitte eingeben (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (end.getSelectedButton() !== ui.Button.OK) return;

  const startDate = parseISODate_(start.getResponseText());
  const endDate = parseISODate_(end.getResponseText());
  if (!startDate || !endDate || startDate > endDate) {
    ui.alert('Ungueltiges Datum. Bitte im Format YYYY-MM-DD und Start <= Ende.');
    return;
  }
  processEmailsDateRange(startDate, endDate);
}

function menu_archiveOldRows() {
  archiveOldRows();
}

function menu_createProjectMapTemplate() {
  ensureProjectMapTemplate_();
}

function menu_createConfigurationTemplate() {
  ensureConfigurationTemplate_();
  SpreadsheetApp.getUi().alert('Erfolg', 'Konfigurationstabelle wurde erstellt/aktualisiert.\n\nBitte passen Sie die Werte in der Tabelle "Configuration" an.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Zeigt Hilfe zum Anzeigen der Logs in Google Apps Script
 */
/**
 * Entfernt Mails mit "Done (sync)" Status aus dem Posteingang
 * Nützlich, wenn Mails bereits archiviert wurden, aber noch im Posteingang sind
 */
function menu_fixDoneSyncInbox() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archiveSheet = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);
    if (!archiveSheet) {
      SpreadsheetApp.getUi().alert('Info', 'Kein Archiv-Sheet gefunden. Nichts zu tun.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const lastRow = archiveSheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('Info', 'Archiv ist leer. Nichts zu tun.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const STATUS_COL = 3; // C = Status
    const MESSAGE_ID_COL = 13; // M = Message ID
    const LABEL_COL = 1; // A = Label

    let processed = 0;
    let errors = 0;
    let notInInbox = 0;

    // Durchsuche Archiv nach "Done (sync)" Einträgen
    for (let row = 2; row <= lastRow; row++) {
      const status = (archiveSheet.getRange(row, STATUS_COL).getValue() || '').toString().trim();
      if (status !== 'Done (sync)') continue;

      const messageId = (archiveSheet.getRange(row, MESSAGE_ID_COL).getValue() || '').toString().trim();
      const labelName = (archiveSheet.getRange(row, LABEL_COL).getValue() || '').toString().trim();

      if (!messageId) continue;

      try {
        const msg = GmailApp.getMessageById(messageId);
        const thread = msg.getThread();
        if (!thread) {
          notInInbox++;
          continue;
        }

        // Prüfe, ob Mail noch im Posteingang ist
        const labels = thread.getLabels();
        const isInInbox = labels.some(l => l.getName() === 'INBOX');

        if (isInInbox) {
          // Mail aus Posteingang entfernen (archivieren)
          thread.moveToArchive();
          processed++;
          Logger.log(`Mail ${messageId} (Label: ${labelName}) aus Posteingang entfernt`);
        } else {
          notInInbox++;
        }
      } catch (e) {
        errors++;
        Logger.log(`Fehler bei Message-ID ${messageId}: ${e}`);
      }
    }

    const message = `Abgeschlossen!\n\n` +
      `✅ Aus Posteingang entfernt: ${processed}\n` +
      `ℹ️ Bereits nicht im Posteingang: ${notInInbox}\n` +
      `❌ Fehler: ${errors}`;

    SpreadsheetApp.getUi().alert('Fertig', message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`menu_fixDoneSyncInbox: ${processed} Mails entfernt, ${notInInbox} bereits entfernt, ${errors} Fehler`);
  } catch (e) {
    Logger.log('menu_fixDoneSyncInbox error: ' + e);
    SpreadsheetApp.getUi().alert('Fehler', 'Fehler beim Entfernen der Mails: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Verarbeitet alle Zeilen mit "Done (sync)" Status
 * Nützlich, wenn mehrere Zeilen per Copy-Paste geändert wurden
 */
function menu_processDoneSyncRows() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Fehler', 'Sheet "current activities" nicht gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('Info', 'Keine Datenzeilen gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const STATUS_COL = 3; // C = Status
    let processed = 0;
    let archived = 0;
    let errors = 0;

    // Durchsuche alle Zeilen nach "Done (sync)" Status
    // Von hinten nach vorne, damit Indizes beim Löschen stimmen
    for (let row = lastRow; row >= 2; row--) {
      const status = (sheet.getRange(row, STATUS_COL).getValue() || '').toString().trim();
      if (status !== 'Done (sync)') continue;

      try {
        // Gmail-Operation: Label setzen + INBOX entfernen
        applyGmailLabelFromRow_(sheet, row);
        processed++;
        
        // Archivierung
        moveRowToArchive_(row);
        archived++;
      } catch (e) {
        errors++;
        Logger.log(`Fehler bei Zeile ${row}: ${e}`);
      }
    }

    const message = `Abgeschlossen!\n\n` +
      `✅ Verarbeitet: ${processed}\n` +
      `📦 Archiviert: ${archived}\n` +
      `❌ Fehler: ${errors}`;

    SpreadsheetApp.getUi().alert('Fertig', message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`menu_processDoneSyncRows: ${processed} Zeilen verarbeitet, ${archived} archiviert, ${errors} Fehler`);
  } catch (e) {
    Logger.log('menu_processDoneSyncRows error: ' + e);
    SpreadsheetApp.getUi().alert('Fehler', 'Fehler beim Verarbeiten der Zeilen: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Verarbeitet alle Zeilen mit "Done (train only)" Status
 * Nützlich, wenn mehrere Zeilen per Copy-Paste geändert wurden
 */
function menu_processDoneTrainOnlyRows() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Fehler', 'Sheet "current activities" nicht gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('Info', 'Keine Datenzeilen gefunden.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const STATUS_COL = 3; // C = Status
    let processed = 0;
    let archived = 0;
    let errors = 0;

    // Durchsuche alle Zeilen nach "Done (train only)" Status
    // Von hinten nach vorne, damit Indizes beim Löschen stimmen
    for (let row = lastRow; row >= 2; row--) {
      const status = (sheet.getRange(row, STATUS_COL).getValue() || '').toString().trim();
      if (status !== 'Done (train only)') continue;

      try {
        // Gmail-Operation: Nur INBOX entfernen (kein Label setzen)
        removeFromInboxOnly_(sheet, row);
        processed++;
        
        // Archivierung
        moveRowToArchive_(row);
        archived++;
      } catch (e) {
        errors++;
        Logger.log(`Fehler bei Zeile ${row}: ${e}`);
      }
    }

    const message = `Abgeschlossen!\n\n` +
      `✅ Verarbeitet: ${processed}\n` +
      `📦 Archiviert: ${archived}\n` +
      `❌ Fehler: ${errors}`;

    SpreadsheetApp.getUi().alert('Fertig', message, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`menu_processDoneTrainOnlyRows: ${processed} Zeilen verarbeitet, ${archived} archiviert, ${errors} Fehler`);
  } catch (e) {
    Logger.log('menu_processDoneTrainOnlyRows error: ' + e);
    SpreadsheetApp.getUi().alert('Fehler', 'Fehler beim Verarbeiten der Zeilen: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menu_showLogsHelp() {
  const ui = SpreadsheetApp.getUi();
  const message = 
    '📋 So sehen Sie die Logs in Google Apps Script:\n\n' +
    '1. Öffnen Sie Google Apps Script:\n' +
    '   - Im Google Sheet: Erweiterungen → Apps Script\n' +
    '   - Oder direkt: script.google.com\n\n' +
    '2. Im Apps Script Editor:\n' +
    '   - Klicken Sie auf "Ausführen" (▶️) in der Toolbar\n' +
    '   - ODER: Menü "Ausführen" → "Ausführen"\n' +
    '   - Wählen Sie eine Funktion (z.B. "menu_generateSummaryForRow")\n\n' +
    '3. Logs anzeigen:\n' +
    '   - Nach der Ausführung: "Ausführen" → "Protokoll anzeigen"\n' +
    '   - ODER: Strg+Enter (Windows) / Cmd+Enter (Mac)\n' +
    '   - ODER: Klicken Sie auf das "Protokoll"-Icon (📋) in der Toolbar\n\n' +
    '4. Logs filtern:\n' +
    '   - Im Protokoll-Fenster können Sie nach Text suchen\n' +
    '   - Verwenden Sie "Summary" oder "Error" als Suchbegriff\n\n' +
    '💡 Tipp: Die Logs zeigen detaillierte Informationen über API-Aufrufe, Fehler und Response-Strukturen.';
  
  ui.alert('📋 Logs anzeigen - Anleitung', message, ui.ButtonSet.OK);
}
