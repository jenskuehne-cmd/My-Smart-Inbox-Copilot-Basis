/*** Menu.gs ***/
// Baut ein Menue "Mail Assistant" im Sheet.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Assistant')
    .addItem('Verarbeite: Heute', 'menu_processToday')
    .addItem('Verarbeite: Letzte 7 Tage', 'menu_processLast7')
    .addItem('Verarbeite: Datumsbereich…', 'menu_processDateRange')
    .addSeparator()
    .addItem('Test: Gemini API-Key', 'menu_testAPIKey')
    .addItem('API-Key Status anzeigen', 'menu_showAPIKeyStatus')
    .addSeparator()
    .addItem('Setup: Dropdowns/Validierung', 'menu_setupValidations')
    .addItem('Archiviere markierte Zeile(n)', 'menu_archiveSelection')
    .addItem('Archivieren: Alte Zeilen', 'menu_archiveOldRows')
    .addSeparator()
    .addItem('Vorlage: Project Map erzeugen', 'menu_createProjectMapTemplate')
    .addItem('ML: Lernen aus Korrekturen', 'menu_learnFromCorrections')
    .addToUi();
}

function menu_setupValidations(){ ensureValidations_(); }

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
