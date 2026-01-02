/*** Utils.gs ***/
// noch ein Test für github und clasp


// Optional: Team-/Rollen-Keywords, die helfen, implizite Aufgaben zu erkennen
CONFIG.MY_TEAM_KEYWORDS = ['BPM', 'Asset Maintenance', 'Aspire', 'S4/HANA'];

// Script Property lesen (mit Fallback)
function getProp_(key, fallback) {
  try {
    const val = PropertiesService.getScriptProperties().getProperty(key);
    return val !== null && val !== undefined ? val : fallback;
  } catch (_) {
    return fallback;
  }
}

// Eigene E-Mail-Adressen (aus Property + Aliasse)
function getMyEmails_() {
  const prop = getProp_('MY_EMAILS', '');
  const list = prop.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  try {
    const aliases = GmailApp.getAliases() || [];
    aliases.forEach(a => list.push(a.toLowerCase()));
  } catch (_) { }
  return Array.from(new Set(list));
}

// Prueft, ob nach letzter eingehender Nachricht eine Mail von mir kam
function getReplyInfoForThread_(messages, myEmails) {
  let lastInbound = null; // {date, message}
  for (const m of messages) {
    const from = (m.getFrom() || '').toLowerCase();
    const isFromMe = myEmails.some(me => from.includes(me));
    if (!isFromMe) {
      if (!lastInbound || m.getDate() > lastInbound.date) {
        lastInbound = { date: m.getDate(), message: m };
      }
    }
  }
  let myLastReply = null; // {date, message}
  if (lastInbound) {
    for (const m of messages) {
      const from = (m.getFrom() || '').toLowerCase();
      const isFromMe = myEmails.some(me => from.includes(me));
      if (isFromMe && m.getDate() > lastInbound.date) {
        if (!myLastReply || m.getDate() > myLastReply.date) {
          myLastReply = { date: m.getDate(), message: m };
        }
      }
    }
  }
  return {
    lastInboundDate: lastInbound?.date || null,
    myLastReplyDate: myLastReply?.date || null,
    hasReplied: !!myLastReply
  };
}

// E-Mail aus "Name <mail@domain>" extrahieren
function extractEmail_(s) {
  if (!s) return '';
  const m = s.match(/<(.+?)>/);
  return (m ? m[1] : s).trim();
}

// ISO-Datum YYYY-MM-DD parsen
function parseISODate_(s) {
  const m = (s || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

// Project Map Vorlage erzeugen (falls noch nicht vorhanden)
/**
 * Lädt Stammdaten aus der Konfigurationstabelle
 * @return {Object} Konfigurationswerte (z.B. {packageStorageDays: 5, ...})
 */
function loadConfiguration_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  
  if (!sh) {
    // Fallback auf Standardwerte, wenn Sheet nicht existiert
    Logger.log('Konfigurations-Sheet nicht gefunden, verwende Standardwerte');
    return {
      packageStorageDays: 5, // Standard: 5 Tage Lagerfrist für Pakete
      defaultTaskPriority: 'Medium',
      enableAutoDueDate: true
    };
  }
  
  const values = sh.getDataRange().getValues();
  const config = {
    packageStorageDays: 5, // Standard
    defaultTaskPriority: 'Medium',
    enableAutoDueDate: true
  };
  
  // Lese Konfigurationswerte aus Sheet (Format: Key | Value | Beschreibung)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const key = (row[0] || '').toString().trim().toLowerCase();
    const value = row[1];
    
    if (!key) continue;
    
    switch (key) {
      case 'package_storage_days':
      case 'paket_lagerfrist':
      case 'lagerfrist':
        config.packageStorageDays = parseInt(value) || 5;
        break;
      case 'default_task_priority':
      case 'standard_task_priorität':
        config.defaultTaskPriority = (value || '').toString().trim() || 'Medium';
        break;
      case 'enable_auto_due_date':
      case 'automatisches_fälligkeitsdatum':
        config.enableAutoDueDate = value === true || value === 'true' || value === 1 || value === '1';
        break;
    }
  }
  
  Logger.log(`Konfiguration geladen: ${JSON.stringify(config)}`);
  return config;
}

/**
 * Erstellt die Konfigurationstabelle-Vorlage
 */
function ensureConfigurationTemplate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.CONFIG_SHEET);
  }
  
  // Header
  sh.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Beschreibung']]);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  
  // Standard-Konfigurationswerte
  const defaultConfig = [
    ['package_storage_days', 5, 'Lagerfrist für Pakete in Tagen (Standard: 5)'],
    ['default_task_priority', 'Medium', 'Standard-Priorität für neue Tasks (Low/Medium/High/Urgent)'],
    ['enable_auto_due_date', true, 'Automatisches Fälligkeitsdatum aktivieren (true/false)']
  ];
  
  // Prüfe, ob bereits Daten vorhanden sind
  if (sh.getLastRow() <= 1) {
    sh.getRange(2, 1, defaultConfig.length, 3).setValues(defaultConfig);
  }
  
  // Formatierung
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 400);
  
  // Validierung für Value-Spalte (je nach Key)
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    for (let i = 2; i <= lastRow; i++) {
      const key = (sh.getRange(i, 1).getValue() || '').toString().trim().toLowerCase();
      if (key === 'default_task_priority' || key === 'standard_task_priorität') {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Low', 'Medium', 'High', 'Urgent'], true)
          .setAllowInvalid(false)
          .build();
        sh.getRange(i, 2).setDataValidation(rule);
      } else if (key === 'enable_auto_due_date' || key === 'automatisches_fälligkeitsdatum') {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['true', 'false'], true)
          .setAllowInvalid(false)
          .build();
        sh.getRange(i, 2).setDataValidation(rule);
      }
    }
  }
  
  Logger.log('Konfigurationstabelle erstellt/aktualisiert');
}

function ensureProjectMapTemplate_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.PROJECT_MAP_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.PROJECT_MAP_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['Project', 'Keywords (comma)', 'From Domains (comma)', 'Regex (optional)']]);
  sh.getRange(2, 1, 1, 4).setValues([['Linden', 'linden,lnd', 'client-a.com', '(LND-\\d+)']]);
  sh.getRange(3, 1, 1, 4).setValues([['Malibu', 'malibu,mb', 'client-b.com', '(MB-\\d+)']]);
  sh.getRange(4, 1, 1, 4).setValues([['LPACX', 'lpacx,lpac', '', '']]);
  SpreadsheetApp.getUi().alert('Project Map Vorlage wurde erstellt/ueberschrieben.');
}
function ensureValidations_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) return;

  const lastRow = Math.max(sh.getLastRow(), 2);

  //
  // 1) Gmail-Label-Liste aktualisieren (falls Fehler → ohne Abbruch weiter)
  //
  let listRange = null;
  try {
    const labSheet = refreshGmailLabelsSheet_();
    const dataRows = Math.max(0, labSheet.getLastRow() - 1); // Zeile 2+
    if (dataRows > 0) {
      listRange = labSheet.getRange(2, 1, dataRows, 1);
    }
  } catch (err) {
    Logger.log("refreshGmailLabelsSheet_ failed: " + err);
  }

  //
  // 2) Validierungen bauen
  //
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      CONFIG.STATUS_OPTIONS || [
        'New',
        'Open',
        'Action Required',
        'Waiting',
        'Replied',
        'Done (sync)',
        'Done (train only)',
        'Closed',
        'Archived'
      ],
      true
    )
    .setAllowInvalid(false)
    .build();

  const prioRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      CONFIG.PRIORITY_OPTIONS || ['Low', 'Medium', 'High', 'Urgent'],
      true
    )
    .setAllowInvalid(false)
    .build();

  const projRule = listRange
    ? SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange, true)
      .setAllowInvalid(false)
      .build()
    : null;

  const taskRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      CONFIG.TASK_FOR_ME_OPTIONS || ['Unsure', 'Yes', 'No'],
      true
    )
    .setAllowInvalid(false)
    .build();

  //
  // 3) Validierungen setzen
  //
  if (lastRow >= 2) {

    // ---- Spalte A: Projekt/Label ----
    if (projRule) {
      const projRange = sh.getRange(2, 1, lastRow - 1, 1);
      projRange.clearDataValidations();
      projRange.setDataValidation(projRule);
    }

    // ---- Spalte C: Status ----
    const statusRange = sh.getRange(2, 3, lastRow - 1, 1);
    statusRange.clearDataValidations();
    statusRange.setDataValidation(statusRule);

    // ---- Spalte D: Priority ----
    const prioRange = sh.getRange(2, 4, lastRow - 1, 1);
    prioRange.clearDataValidations();
    prioRange.setDataValidation(prioRule);

    // ---- Spalte I: Task for Me ----
    const taskRange = sh.getRange(2, 9, lastRow - 1, 1);
    taskRange.clearDataValidations();
    taskRange.setDataValidation(taskRule);
  }
}


// Baut/aktualisiert ein Hilfssheet "Gmail Labels" mit allen User-Labels (ohne IGNORE_LABELS)
function refreshGmailLabelsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Gmail Labels') || ss.insertSheet('Gmail Labels');

  sh.clear();
  sh.getRange(1, 1).setValue('Labels');

  const all = GmailApp.getUserLabels()
    .map(l => l.getName())
    .filter(n => !(CONFIG.IGNORE_LABELS && CONFIG.IGNORE_LABELS.has(n)))
    .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

  // Optionaler erster Eintrag: "Uncategorized"
  const values = [['Labels']].concat([['Uncategorized']], all.map(n => [n]));
  sh.getRange(1, 1, values.length, 1).setValues(values);

  return sh;
}

function computePriority_(subject, body, fromEmail, toEmails, actionAnalysis) {
  const txt = `${subject}\n${body}`.toLowerCase();
  const rules = CONFIG.PRIORITY_RULES || {};
  const singleRecipient = Array.isArray(toEmails) && toEmails.filter(Boolean).length <= 1;

  // 1) Prüfe zuerst gelernte Patterns (höchste Priorität)
  const fromDom = ((fromEmail || '').split('@')[1] || '').toLowerCase();
  const learnedPriority = getLearnedPriority_(fromDom, subject);
  if (learnedPriority) {
    Logger.log(`Priority Learning: Verwende gelernte Priorität "${learnedPriority}" für "${subject}"`);
    return learnedPriority;
  }

  // 2) AI-Analyse (falls aktiviert)
  if (rules.AI_ENABLE && actionAnalysis && Array.isArray(actionAnalysis.tasks) && actionAnalysis.tasks.length) {
    const p = (actionAnalysis.tasks[0].priority || '').toLowerCase();
    if (p.startsWith('high') || p.startsWith('urgent')) return 'High';
    if (p.startsWith('low')) return 'Low';
  }

  // 3) Statische Keyword-Regeln
  if ((rules.HIGH_KEYWORDS || []).some(k => txt.includes(k.toLowerCase()))) return 'High';
  if ((rules.LOW_KEYWORDS || []).some(k => txt.includes(k.toLowerCase()))) return 'Low';

  // 4) Domain-Regeln
  if ((rules.DOMAIN_HIGH || []).some(d => fromDom.includes(d.toLowerCase()))) return 'High';

  // 5) Single-Recipient-Boost
  if (rules.SINGLE_RECIPIENT_BOOST && singleRecipient) return 'High';

  // 6) Default
  return (CONFIG.PRIORITY_DEFAULT || 'Medium');
}

function isCalendarInvite_(msg) {
  const from = (msg.getFrom() || '').toLowerCase();
  const subject = (msg.getSubject() || '').toLowerCase();
  const body = getBestBody_(msg).toLowerCase();

  // typische Absender/Pattern fuer Google Calendar
  const fromCal = from.includes('calendar-notification@') ||
    from.includes('calendar-noreply@') ||
    from.includes('@calendar.google.com');

  const subjCal = subject.startsWith('invitation:') ||
    subject.startsWith('updated invitation:') ||
    subject.startsWith('canceled event:') ||
    subject.includes('event invitation');

  // du kannst hier gern weitere firmeninterne Patterns ergaenzen
  return fromCal || subjCal;
}

/**
 * Prüft, ob es sich um eine Terminbestätigung handelt (bereits bestätigt/abgelehnt)
 * @param {GmailMessage} msg - Die E-Mail-Nachricht
 * @return {Object} {isConfirmed: boolean, status: string} - Status: 'accepted', 'declined', 'tentative', 'none'
 */
function getCalendarInviteStatus_(msg) {
  try {
    const subject = (msg.getSubject() || '').toLowerCase();
    const body = getBestBody_(msg).toLowerCase();
    const combined = `${subject} ${body}`;
    
    // Prüfe auf Bestätigungs-Patterns
    const acceptedPatterns = [
      'hat zugesagt', 'hat angenommen', 'accepted', 'zugesagt', 'angenommen',
      'you accepted', 'you are attending', 'teilnahme bestätigt',
      'confirmed', 'bestätigt', 'wird teilnehmen'
    ];
    
    const declinedPatterns = [
      'hat abgelehnt', 'hat abgesagt', 'declined', 'abgelehnt', 'abgesagt',
      'you declined', 'you are not attending', 'teilnahme abgelehnt',
      'nicht teilnehmen', 'kann nicht teilnehmen'
    ];
    
    const tentativePatterns = [
      'hat vorläufig zugesagt', 'tentative', 'vorläufig', 'maybe',
      'you are tentatively attending', 'möglicherweise teilnehmen'
    ];
    
    // Prüfe Subject und Body
    if (acceptedPatterns.some(pattern => combined.includes(pattern))) {
      return { isConfirmed: true, status: 'accepted' };
    }
    
    if (declinedPatterns.some(pattern => combined.includes(pattern))) {
      return { isConfirmed: true, status: 'declined' };
    }
    
    if (tentativePatterns.some(pattern => combined.includes(pattern))) {
      return { isConfirmed: true, status: 'tentative' };
    }
    
    // Prüfe auf typische Bestätigungs-Subjects
    if (subject.includes('bestätigung') || subject.includes('confirmation') ||
        subject.includes('zugesagt') || subject.includes('abgelehnt') ||
        subject.includes('accepted') || subject.includes('declined')) {
      return { isConfirmed: true, status: 'accepted' }; // Default zu accepted wenn unklar
    }
    
    return { isConfirmed: false, status: 'none' };
  } catch (e) {
    Logger.log('getCalendarInviteStatus_ error: ' + e);
    return { isConfirmed: false, status: 'none' };
  }
}

function recordCorrection_(sheet, row, col, oldVal, newVal) {
  try {
    const ss = sheet.getParent();
    const log = ss.getSheetByName('Learning Log') || ss.insertSheet('Learning Log');
    if (log.getLastRow() < 1) {
      log.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Message ID', 'Field', 'Old', 'New', 'Subject']]);
    }
    const msgId = sheet.getRange(row, 13).getValue(); // M
    const subject = sheet.getRange(row, 2).getValue(); // B
    const field = col === 1 ? 'Project/Label' : (col === 3 ? 'Status' : (col === 4 ? 'Priority' : `Col${col}`));
    log.appendRow([new Date(), msgId || '', field, oldVal || '', newVal || '', subject || '']);
  } catch (_) { }
}

function learnFromProjectCorrection_(sheet, row) {
  try {
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    const selectedRaw = (sheet.getRange(row, 1).getValue() || '').toString().trim(); // A (Projekt/Label)
    if (!messageId || !selectedRaw) return;

    const msg = GmailApp.getMessageById(messageId);
    const from = (msg.getFrom && msg.getFrom()) || '';
    const dom = ((from.split('@')[1] || '').split('>')[0] || '').toLowerCase().trim();
    if (!dom) return;

    // 1) Versuch: ueber vorhandene Mapping-Logik normalisieren
    // (z.B. "0 Aspire" → "Aspire", "Newsfeed" → "Private", wenn so hinterlegt)
    let mapped = null;
    try {
      mapped = identifyProjectByLabels_([{ getName: () => selectedRaw }]) || null;
    } catch (_) {
      mapped = null;
    }

    // 2) Fallback: wenn nichts gemappt wurde → genau den Wert aus Spalte A verwenden
    const project = (mapped && mapped !== 'Uncategorized') ? mapped : selectedRaw;

    // 3) Domain→Project-Historie erhoehen (wirkt direkt ins Scoring)
    bumpHistory_(dom, project);

  } catch (e) {
    Logger.log('learnFromProjectCorrection_ error: ' + e);
  }
}


function stripHtml_(html) {
  if (!html) return '';
  try {
    // Schneller Fallback: Tags weg
    return html.replace(/<[^>]*>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim();
  } catch (e) {
    return html;
  }
}

function getBestBody_(msg) {
  // 1) Plain Body
  let txt = (msg.getPlainBody && msg.getPlainBody()) || '';
  if (txt && txt.trim().length >= 20) return txt;

  // 2) HTML Body -> strippen
  try {
    const html = (msg.getBody && msg.getBody()) || '';
    const stripped = stripHtml_(html);
    if (stripped && stripped.trim().length >= 20) return stripped;
  } catch (e) {
    // ignore
  }

  // 3) Fallback: Subject + kleine Info
  const subj = (msg.getSubject && msg.getSubject()) || '(no subject)';
  return `Subject: ${subj}\n(Kein lesbarer Mail-Body gefunden.)`;
}
/**
 * Entfernt nur das Posteingang-Label (INBOX), ohne ein neues Label zu setzen
 * Verschiebt die Mail damit in den Papierkorb (wenn keine anderen Labels vorhanden sind)
 * Learning passiert bereits durch Spalte A-Änderung
 */
function removeFromInboxOnly_(sheet, row) {
  try {
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    Logger.log(`removeFromInboxOnly_: row=${row}, msgId="${messageId}"`);
    if (!messageId) return;

    const msg = GmailApp.getMessageById(messageId);
    const thread = msg.getThread();
    if (!thread) return;

    // Mail aus Posteingang entfernen (in Papierkorb verschieben)
    // INBOX ist ein System-Label und kann nicht direkt entfernt werden
    // moveToTrash() verschiebt die Mail in den Papierkorb
    try {
      thread.moveToTrash();
      Logger.log('Mail aus Posteingang in Papierkorb verschoben');
    } catch (e) {
      Logger.log('Fehler beim Verschieben in Papierkorb: ' + e);
      // Fallback: Versuche zu archivieren (entfernt aus INBOX, aber behält Mail)
      try {
        thread.moveToArchive();
        Logger.log('Mail archiviert (Fallback, da Papierkorb fehlgeschlagen)');
      } catch (e2) {
        Logger.log('Auch Archivierung fehlgeschlagen: ' + e2);
      }
    }
  } catch (e) {
    Logger.log('removeFromInboxOnly_ error: ' + e);
  }
}

// Entfernt vorhandene "Projekt-Labels" am Thread und setzt das neue (aus Spalte A)
// Entfernt auch das Posteingang-Label (INBOX), damit die Mail nur noch unter dem neuen Label erscheint
function applyGmailLabelFromRow_(sheet, row) {

  try {
    const labelName = (sheet.getRange(row, 1).getValue() || '').toString().trim(); // A
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    Logger.log(`applyGmailLabelFromRow_: row=${row}, label="${labelName}", msgId="${messageId}"`);
    if (!messageId || !labelName) return;

    const msg = GmailApp.getMessageById(messageId);
    const thread = msg.getThread();
    if (!thread) return;

    // Ziel-Label setzen (zuerst, damit die Mail das Label hat)
    const target = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
    thread.addLabel(target);
    Logger.log(`Label "${labelName}" hinzugefügt`);

    // Mail aus Posteingang entfernen (archivieren)
    // INBOX ist ein System-Label und kann nicht direkt entfernt werden
    // moveToArchive() entfernt die Mail aus INBOX, behält aber alle Labels
    try {
      thread.moveToArchive();
      Logger.log('Mail aus Posteingang archiviert (bleibt unter Label "' + labelName + '" sichtbar)');
    } catch (e) {
      Logger.log('Fehler beim Archivieren der Mail: ' + e);
      // Fallback: Versuche INBOX-Label zu entfernen (funktioniert möglicherweise nicht)
      try {
        const inboxLabel = GmailApp.getUserLabelByName('INBOX');
        if (inboxLabel) {
          thread.removeLabel(inboxLabel);
          Logger.log('Posteingang-Label (INBOX) entfernt (Fallback)');
        }
      } catch (e2) {
        Logger.log('Auch Fallback fehlgeschlagen: ' + e2);
      }
    }

    // Optionale: Andere Projekt-Labels entfernen (falls gewünscht)
    // Aktuell auskommentiert, um alle Labels zu behalten
    // Wenn Sie möchten, dass nur das neue Label bleibt, können Sie dies aktivieren:
    /*
    const tLabels = thread.getLabels();
    for (const lab of tLabels) {
      const name = lab.getName();
      // Ignoriere System-Labels und das neue Label
      if (name === labelName) continue;
      if (name === 'INBOX') continue; // Bereits entfernt
      if (CONFIG.IGNORE_LABELS && CONFIG.IGNORE_LABELS.has(name)) continue;
      // Entferne andere Projekt-Labels (optional)
      // thread.removeLabel(lab);
    }
    */
  } catch (e) {
    Logger.log('applyGmailLabelFromRow_ error: ' + e);
  }
}


// Heuristik: ist ein Label mutmasslich ein Projekt-Label?
function isProjectishLabel_(name) {
  if (!name) return false;
  const cleaned = name.toString().trim();

  // 1) Regex-Map
  const rx = CONFIG.LABEL_REGEX_MAP || [];
  for (const r of rx) {
    try { if (new RegExp(r.re, 'i').test(cleaned)) return true; } catch (_) { }
  }

  // 2) Segment-Map (exakte oder Teil-Treffer)
  const seg = CONFIG.LABEL_SEGMENT_MAP || {};
  const lower = cleaned.toLowerCase();
  if (seg[lower]) return true;
  for (const k in seg) { if (k && lower.includes(k)) return true; }

  // 3) Exakte Map
  const exact = CONFIG.LABEL_TO_PROJECT || {};
  if (exact[cleaned]) return true;

  return false;
}

// Manuelle Projekt-/Labelkorrektur "lernen": Domain → normiertes Projekt boosten
function learnFromProjectCorrection_(sheet, row) {
  try {
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    const selected = (sheet.getRange(row, 1).getValue() || '').toString().trim();   // A (Label)
    if (!messageId || !selected) return;

    const msg = GmailApp.getMessageById(messageId);
    const from = (msg.getFrom && msg.getFrom()) || '';
    const dom = ((from.split('@')[1] || '').split('>')[0] || '').toLowerCase().trim();

    // Label → Projekt-normalisiert (nutzt deine Mapping-Logik)
    // const mapped = identifyProjectByLabels_([selected]) || 'Uncategorized';
    const mapped = identifyProjectByLabels_([{ getName: () => selected }]) || 'Uncategorized';
    if (dom && mapped && mapped !== 'Uncategorized') {
      bumpHistory_(dom, mapped); // Project History erhöhen → wirkt ins Scoring hinein
    }
  } catch (e) {
    Logger.log('learnFromProjectCorrection_ error: ' + e);
  }
}
// Synchronisiert eine Zeile mit Google Tasks (basierend auf "Task for Me" + Priority)
function syncGoogleTaskForRow_(sheet, row) {
  try {
    const taskForMe = (sheet.getRange(row, 9).getValue() || '').toString().trim(); // I
    const priority = (sheet.getRange(row, 4).getValue() || '').toString().trim();  // D
    const subject = (sheet.getRange(row, 2).getValue() || '').toString().trim();  // B
    const notes = (sheet.getRange(row, 7).getValue() || '').toString().trim();  // G
    const link = (sheet.getRange(row, 12).getFormula() || sheet.getRange(row, 12).getValue() || '').toString(); // L
    const msgId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    // KI Task-Vorschlag aus Spalte E
    const aiTaskTitle = (sheet.getRange(row, 5).getValue() || '').toString().trim(); // E


    // Spalte F: Google Task ID
    const taskIdRange = sheet.getRange(row, 6);
    const existingTaskId = (taskIdRange.getValue() || '').toString().trim();

    // Task wird erstellt wenn "Task for Me" = Yes (unabhängig von Priorität)
    const isTaskYes = ['Yes', 'Ja', 'J', 'Y'].includes(taskForMe);

    const listId = getProp_('GOOGLE_TASKS_LIST_ID', '@default');

    // Fall 1: Kein Task mehr gewuenscht → ggf. vorhandenen Task loeschen
    if (!isTaskYes) {
      if (existingTaskId) {
        try {
          Tasks.Tasks.remove(listId, existingTaskId);
        } catch (e) {
          Logger.log('Task remove failed (ignoriert): ' + e);
        }
        taskIdRange.clearContent();
      }
      return;
    }

    // Prüfe, ob es sich um eine wiederholte Paket-Erinnerung handelt
    // Wenn ja, verwende den bestehenden Task statt einen neuen zu erstellen
    if (!existingTaskId) {
      // Hole Body-Text für Sendungsnummer-Extraktion (falls verfügbar)
      let bodyText = '';
      try {
        if (msgId) {
          const msg = GmailApp.getMessageById(msgId);
          if (msg) {
            bodyText = getBestBody_(msg);
          }
        }
      } catch (e) {
        Logger.log('Konnte Body nicht laden für Paket-Erkennung: ' + e);
      }
      
      const existingTaskIdForPackage = findExistingTaskForPackage_(sheet, row, aiTaskTitle, subject, bodyText);
      if (existingTaskIdForPackage) {
        Logger.log(`Wiederholte Paket-Erinnerung erkannt - verwende bestehenden Task: ${existingTaskIdForPackage}`);
        taskIdRange.setValue(existingTaskIdForPackage);
        // Aktualisiere den bestehenden Task mit neuesten Infos
        try {
          const existing = Tasks.Tasks.get(listId, existingTaskIdForPackage);
          const taskTitle = aiTaskTitle || subject || 'Mail Task';
          existing.title = taskTitle;
          
          // Berechne Due-Date: Lagerfrist ab Erhalt (falls nicht bereits gesetzt)
          let dueDate = existing.due;
          if (!dueDate) {
            try {
              // Versuche Empfangsdatum aus Notes zu extrahieren
              const receivedMatch = notes.match(/Empfangen:\s*(\d{2}\.\d{2}\.\d{4})/);
              if (receivedMatch) {
                const config = loadConfiguration_();
                const storageDays = config.packageStorageDays || 5; // Fallback auf 5 Tage
                
                const [day, month, year] = receivedMatch[1].split('.');
                const receivedDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                dueDate = new Date(receivedDate);
                dueDate.setDate(dueDate.getDate() + storageDays);
                existing.due = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
                Logger.log(`Due-Date für wiederholte Erinnerung berechnet: ${dueDate.toLocaleDateString('de-DE')} (${storageDays} Tage)`);
              }
            } catch (e) {
              Logger.log('Due-Date-Berechnung fehlgeschlagen: ' + e);
            }
          }
          
          // Vereinfachter Task-Text für Update
          const updatedDueDateText = dueDate 
            ? Utilities.formatDate(new Date(dueDate), Session.getScriptTimeZone(), 'dd.MM.yyyy')
            : 'nicht gefunden';
          
          const updatedNotes = [
            subject || '',
            '',
            '📅 Fällig bis: ' + updatedDueDateText + (dueDateSource ? ` (${dueDateSource})` : ''),
            dueDate ? '' : '⚠️ Bitte manuell prüfen und Fälligkeitsdatum setzen',
            '',
            '🔗 Mail: ' + link.replace(/=HYPERLINK\(|"|\)/g, '').split(',')[0] || link,
            '',
            `⏰ Letzte Erinnerung: ${new Date().toLocaleDateString('de-DE')}`
          ].filter(line => line !== '').join('\n');
          
          existing.notes = updatedNotes;
          Tasks.Tasks.update(existing, listId, existingTaskIdForPackage);
        } catch (e) {
          Logger.log('Task-Update für wiederholte Erinnerung fehlgeschlagen: ' + e);
        }
        return; // Kein neuer Task nötig
      }
    }

    // Fall 2: Task gewuenscht → neu anlegen oder aktualisieren (auch für Low/Medium Priority)
    const taskTitle = aiTaskTitle || subject || 'Mail Task';

    // Extrahiere Due-Date aus verschiedenen Quellen
    let dueDate = null;
    let dueDateSource = '';
    try {
      // 1) Versuche Due-Date aus Notes zu extrahieren (wird von AI-Analyse gesetzt)
      const dueDateMatch = notes.match(/due_date[:\s]+(\d{4}-\d{2}-\d{2})/i);
      if (dueDateMatch) {
        dueDate = new Date(dueDateMatch[1]);
        dueDateSource = 'AI-Analyse';
      }
      
      // 2) Aus AI-Task-Titel (z.B. "Paket abholen bis zum 15.03.2025" oder "Termin am 20.01.2026")
      if (!dueDate) {
        const dateMatch = taskTitle.match(/(?:bis|until|by|am|zum|on)\s+(?:zum|zum|am|on)?\s*(\d{1,2})\.(\d{1,2})\.(\d{4})/i);
        if (dateMatch) {
          const [, day, month, year] = dateMatch;
          dueDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
          dueDateSource = 'Task-Titel';
        }
      }
      
      // 3) Aus Body-Text (falls verfügbar) - hole Body aus Message-ID
      if (!dueDate && msgId) {
        try {
          const msg = GmailApp.getMessageById(msgId);
          if (msg) {
            const body = getBestBody_(msg);
            // Suche nach Datumsangaben im Body (verschiedene Formate)
            const bodyDatePatterns = [
              /(?:bis|until|by|deadline|fällig|abholen bis|abholbar bis)[:\s]+(\d{1,2})\.(\d{1,2})\.(\d{4})/i,
              /(?:bis|until|by|deadline|fällig)[:\s]+(\d{4})-(\d{2})-(\d{2})/i,
              /(?:am|on)\s+(\d{1,2})\.(\d{1,2})\.(\d{4})/i
            ];
            
            for (const pattern of bodyDatePatterns) {
              const match = body.match(pattern);
              if (match) {
                if (match.length === 4) {
                  // Format: DD.MM.YYYY
                  const [, day, month, year] = match;
                  dueDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                } else if (match.length === 4 && match[0].includes('-')) {
                  // Format: YYYY-MM-DD
                  const [, year, month, day] = match;
                  dueDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
                }
                if (dueDate) {
                  dueDateSource = 'E-Mail-Body';
                  break;
                }
              }
            }
          }
        } catch (e) {
          Logger.log('Body-Extraktion für Due-Date fehlgeschlagen: ' + e);
        }
      }
      
      // 4) Für Pakete: Lagerfrist ab Empfangsdatum (falls Paket erkannt und noch kein Due-Date)
      if (!dueDate && (taskTitle.toLowerCase().includes('paket') || taskTitle.toLowerCase().includes('packstation'))) {
        const receivedMatch = notes.match(/Empfangen:\s*(\d{2}\.\d{2}\.\d{4})/);
        if (receivedMatch) {
          const config = loadConfiguration_();
          const storageDays = config.packageStorageDays || 5; // Fallback auf 5 Tage
          
          const [day, month, year] = receivedMatch[1].split('.');
          const receivedDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
          dueDate = new Date(receivedDate);
          dueDate.setDate(dueDate.getDate() + storageDays);
          dueDateSource = `Paket-Standard (${storageDays} Tage)`;
          Logger.log(`Due-Date für Paket berechnet: ${dueDate.toLocaleDateString('de-DE')} (${storageDays} Tage ab Empfang)`);
        }
      }
      
      if (dueDate) {
        Logger.log(`Due-Date gefunden (${dueDateSource}): ${dueDate.toLocaleDateString('de-DE')}`);
      }
    } catch (e) {
      Logger.log('Due-Date-Extraktion fehlgeschlagen: ' + e);
    }

    // Vereinfachter Task-Text (nur wichtige Infos)
    const dueDateText = dueDate 
      ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd.MM.yyyy')
      : 'nicht gefunden';
    
    const taskNotes = [
      subject || '',
      '',
      '📅 Fällig bis: ' + dueDateText + (dueDateSource ? ` (${dueDateSource})` : ''),
      dueDate ? '' : '⚠️ Bitte manuell prüfen und Fälligkeitsdatum setzen',
      '',
      '🔗 Mail: ' + link.replace(/=HYPERLINK\(|"|\)/g, '').split(',')[0] || link
    ].filter(line => line !== '').join('\n');

    if (existingTaskId) {
      // Update bestehenden Task
      try {
        const existing = Tasks.Tasks.get(listId, existingTaskId);
        existing.title = taskTitle;
        existing.notes = taskNotes;
        if (dueDate) {
          // Due-Date im ISO-Format für Google Tasks
          const dueDateISO = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
          existing.due = dueDateISO;
          Logger.log(`Due-Date gesetzt: ${dueDateISO} (${Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd.MM.yyyy')})`);
        }
        Tasks.Tasks.update(existing, listId, existingTaskId);
        Logger.log(`Task aktualisiert: ${existingTaskId}`);
      } catch (e) {
        Logger.log('Task-Update fehlgeschlagen: ' + e);
      }
      return;
    }

    // Neuer Task
    const newTask = {
      title: taskTitle,
      notes: taskNotes
    };
    
    // Due-Date hinzufügen falls vorhanden
    if (dueDate) {
      const dueDateISO = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
      newTask.due = dueDateISO;
      Logger.log(`Due-Date für neuen Task gesetzt: ${dueDateISO} (${Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd.MM.yyyy')})`);
    }

    try {
      const created = Tasks.Tasks.insert(newTask, listId);
      if (created && created.id) {
        taskIdRange.setValue(created.id);
        Logger.log(`Neuer Task erstellt: ${created.id} mit Due-Date: ${dueDate ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd.MM.yyyy') : 'kein'}`);
      }
    } catch (e) {
      Logger.log('Task-Erstellung fehlgeschlagen: ' + e);
    }

  } catch (e) {
    Logger.log('syncGoogleTaskForRow_ error: ' + e);
  }
}

// Legt einen Google Task fuer eine Zeile an (einfache Version ohne Rueck-Sync)
function createGoogleTaskFromRow_(sheet, row) {
  try {
    const subject = (sheet.getRange(row, 2).getValue() || '').toString().trim(); // B
    const aiTask = (sheet.getRange(row, 5).getValue() || '').toString().trim();  // E: KI-Task-Vorschlag
    const notesCell = (sheet.getRange(row, 7).getValue() || '').toString().trim(); // G: Notes (inkl. AI-Task-Grund)
    const summary = (sheet.getRange(row, 8).getValue() || '').toString().trim();   // H: AI Summary
    const link = (sheet.getRange(row, 12).getFormula() || sheet.getRange(row, 12).getValue() || '').toString(); // L
    const msgId = (sheet.getRange(row, 13).getValue() || '').toString().trim();    // M

    const listId = getProp_('GOOGLE_TASKS_LIST_ID', '@default');

    const title = aiTask || subject || 'Mail Task';
    const notes = [
      summary,
      '',
      notesCell,
      '',
      'Mail-Link:',
      link,
      '',
      'Message-ID: ' + msgId
    ].join('\n');

    const task = { title: title, notes: notes };

    Tasks.Tasks.insert(task, listId);

  } catch (e) {
    Logger.log('createGoogleTaskFromRow_ error: ' + e);
  }
}
// Erzeugt bei Bedarf eine KI-Task-Suggestion fuer eine Zeile
function ensureTaskSuggestionForRow_(sheet, row) {
  try {
    // Task for Me (I)
    const taskForMe = (sheet.getRange(row, 9).getValue() || '').toString().trim();
    if (taskForMe !== 'Yes') {
      return; // Nur fuer echte Tasks
    }

    // Wenn in E schon etwas steht, nichts tun
    const currentSuggestion = (sheet.getRange(row, 5).getValue() || '').toString().trim();
    if (currentSuggestion) {
      return;
    }

    const subject = (sheet.getRange(row, 2).getValue() || '').toString().trim();  // B
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    const summary = (sheet.getRange(row, 8).getValue() || '').toString().trim();  // H: AI Summary
    const myProfile = CONFIG.MY_PROFILE || '';

    // Versuche vollständigen Body-Text aus der Message-ID zu holen (besserer Kontext)
    let bodyText = summary; // Fallback: Summary
    if (messageId) {
      try {
        const msg = GmailApp.getMessageById(messageId);
        if (msg) {
          const fullBody = getBestBody_(msg);
          // Verwende vollständigen Body, aber begrenzt auf 5000 Zeichen für API
          bodyText = fullBody.length > 5000 ? fullBody.slice(0, 5000) + '...' : fullBody;
          Logger.log(`ensureTaskSuggestionForRow_: Verwende vollständigen Body-Text (${bodyText.length} Zeichen)`);
        }
      } catch (e) {
        Logger.log('Konnte Message nicht laden, verwende Summary: ' + e);
        bodyText = summary; // Fallback auf Summary
      }
    }

    // AI-Analyse mit vollständigem Body-Text
    const action = getAIActionabilityAnalysis(subject, bodyText, myProfile);
    let aiTaskTitle = '';

    if (action && Array.isArray(action.tasks) && action.tasks.length) {
      aiTaskTitle = (action.tasks[0].title || '').toString().trim();
    }

    // Wenn AI keinen aktionsorientierten Task generiert hat, versuche es mit spezifischem Prompt
    if (!aiTaskTitle || aiTaskTitle === subject || aiTaskTitle.toLowerCase().includes('sendung liegt')) {
      Logger.log(`Task-Titel ist nicht aktionsorientiert ("${aiTaskTitle}"), generiere neuen...`);
      aiTaskTitle = generateActionOrientedTaskTitle_(subject, bodyText);
    }

    if (!aiTaskTitle) {
      // Letzter Fallback: versuche aus Subject eine aktionsorientierte Task zu machen
      aiTaskTitle = convertSubjectToActionTask_(subject);
    }

    sheet.getRange(row, 5).setValue(aiTaskTitle); // E: Task Suggestion setzen
    Logger.log(`Task-Suggestion generiert: "${aiTaskTitle}"`);

  } catch (e) {
    Logger.log('ensureTaskSuggestionForRow_ error: ' + e);
  }
}

/**
 * Generiert einen aktionsorientierten Task-Titel basierend auf Subject und Body
 */
function generateActionOrientedTaskTitle_(subject, body) {
  try {
    const prompt = `You are a task management assistant. Convert the following email notification into a clear, action-oriented task title.

IMPORTANT RULES:
- Task titles MUST start with an ACTION VERB (e.g., "Abholen", "Bezahlen", "Bestätigen", "Prüfen")
- Include SPECIFIC details (e.g., Packstation number, dates, amounts)
- Keep it concise (max 50 characters)
- Use German language

Examples:
- "Ihre Sendung liegt in der Packstation 117" → "Paket aus Packstation 117 abholen"
- "Rechnung erhalten: 123,45 EUR" → "Rechnung bezahlen (123,45 EUR)"
- "Terminbestätigung: Meeting am 15.03.2025" → "Termin am 15.03.2025 bestätigen"
- "Paket-Benachrichtigung: Packstation 42" → "Paket aus Packstation 42 abholen"

Email:
Subject: ${subject}
Body: ${body.substring(0, 1000)}

Return ONLY the task title, no explanation, no JSON, just the title.`;

    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 100 }
    };

    const res = callGeminiAPI_(payload);
    const data = JSON.parse(res.getContentText() || '{}');
    const taskTitle = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    return taskTitle.trim();
  } catch (e) {
    Logger.log('generateActionOrientedTaskTitle_ error: ' + e);
    return '';
  }
}

/**
 * Findet einen bestehenden Task für das gleiche Paket
 * Verwendet Sendungsnummer (falls vorhanden) oder Packstation-Nummer als Fallback
 * Verhindert, dass tägliche Erinnerungen zu demselben Paket neue Tasks erstellen
 * 
 * @param {Sheet} sheet - Das Sheet
 * @param {number} currentRow - Aktuelle Zeile
 * @param {string} taskTitle - Task-Titel
 * @param {string} subject - E-Mail Subject
 * @param {string} body - E-Mail Body (optional, für Sendungsnummer-Extraktion)
 * @return {string|null} Bestehende Task-ID oder null
 */
function findExistingTaskForPackage_(sheet, currentRow, taskTitle, subject, body) {
  try {
    const combinedText = (taskTitle + ' ' + subject + ' ' + (body || '')).toLowerCase();
    
    // 1) Versuche Sendungsnummer zu extrahieren (wichtigste Identifikation)
    // Format: "Sendung 0130E0A7500200014149" oder "Sendungsnummer: 0130E0A7500200014149"
    const sendungsnummerMatch = combinedText.match(/(?:sendung|sendungsnummer|tracking)[:\s]+([a-z0-9]{10,20})/i);
    let sendungsnummer = null;
    if (sendungsnummerMatch) {
      sendungsnummer = sendungsnummerMatch[1].toUpperCase();
      Logger.log(`Sendungsnummer gefunden: ${sendungsnummer}`);
    }
    
    // 2) Extrahiere Packstation-Nummer (Fallback, wenn keine Sendungsnummer)
    const packstationMatch = combinedText.match(/packstation\s*(\d+)/i);
    const packstationNum = packstationMatch ? packstationMatch[1] : null;
    
    if (!sendungsnummer && !packstationNum) {
      return null; // Keine Identifikation gefunden → kein Paket-Task
    }
    
    // 3) Suche in allen Zeilen nach bestehenden Tasks
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    
    const listId = getProp_('GOOGLE_TASKS_LIST_ID', '@default');
    
    for (let r = 2; r <= lastRow; r++) {
      if (r === currentRow) continue; // Aktuelle Zeile überspringen
      
      const taskForMe = (sheet.getRange(r, 9).getValue() || '').toString().trim();
      if (taskForMe !== 'Yes') continue; // Nur Tasks prüfen
      
      const existingTaskId = (sheet.getRange(r, 6).getValue() || '').toString().trim();
      if (!existingTaskId) continue; // Kein Task vorhanden
      
      // Prüfe Task-Titel, Subject und Body (falls verfügbar)
      const existingTaskTitle = (sheet.getRange(r, 5).getValue() || '').toString().trim();
      const existingSubject = (sheet.getRange(r, 2).getValue() || '').toString().trim();
      const existingSummary = (sheet.getRange(r, 8).getValue() || '').toString().trim();
      const existingCombined = (existingTaskTitle + ' ' + existingSubject + ' ' + existingSummary).toLowerCase();
      
      let matchFound = false;
      
      // Priorität 1: Sendungsnummer-Match (genaueste Identifikation)
      if (sendungsnummer) {
        const existingSendungsnummerMatch = existingCombined.match(/(?:sendung|sendungsnummer|tracking)[:\s]+([a-z0-9]{10,20})/i);
        if (existingSendungsnummerMatch) {
          const existingSendungsnummer = existingSendungsnummerMatch[1].toUpperCase();
          if (existingSendungsnummer === sendungsnummer) {
            matchFound = true;
            Logger.log(`Gleiche Sendungsnummer gefunden: ${sendungsnummer}`);
          }
        }
      }
      
      // Priorität 2: Packstation-Nummer-Match (nur wenn keine Sendungsnummer vorhanden)
      if (!matchFound && packstationNum && !sendungsnummer) {
        const existingPackstationMatch = existingCombined.match(/packstation\s*(\d+)/i);
        if (existingPackstationMatch && existingPackstationMatch[1] === packstationNum) {
          // Zusätzliche Prüfung: Wenn beide keine Sendungsnummer haben, dann Match
          const existingHasSendungsnummer = /(?:sendung|sendungsnummer|tracking)[:\s]+[a-z0-9]{10,20}/i.test(existingCombined);
          if (!existingHasSendungsnummer) {
            matchFound = true;
            Logger.log(`Gleiche Packstation gefunden (ohne Sendungsnummer): ${packstationNum}`);
          }
        }
      }
      
      if (matchFound) {
        // Prüfe, ob Task noch existiert und nicht completed ist
        try {
          const task = Tasks.Tasks.get(listId, existingTaskId);
          if (task && task.status !== 'completed') {
            Logger.log(`Bestehender Task gefunden: ${existingTaskId}`);
            return existingTaskId;
          }
        } catch (e) {
          // Task existiert nicht mehr → ignoriere
          Logger.log(`Task ${existingTaskId} existiert nicht mehr: ${e}`);
        }
      }
    }
    
    return null; // Kein bestehender Task gefunden
  } catch (e) {
    Logger.log('findExistingTaskForPackage_ error: ' + e);
    return null;
  }
}

/**
 * Konvertiert einen Subject-Text in eine aktionsorientierte Task (Fallback)
 */
function convertSubjectToActionTask_(subject) {
  const subjectLower = subject.toLowerCase();
  
  // Spezielle Patterns für häufige Fälle
  if (subjectLower.includes('packstation') || subjectLower.includes('paket')) {
    const packstationMatch = subject.match(/packstation\s*(\d+)/i);
    if (packstationMatch) {
      return `Paket aus Packstation ${packstationMatch[1]} abholen`;
    }
    return 'Paket abholen';
  }
  
  if (subjectLower.includes('rechnung') || subjectLower.includes('invoice')) {
    return 'Rechnung bezahlen';
  }
  
  if (subjectLower.includes('termin') || subjectLower.includes('meeting') || subjectLower.includes('einladung')) {
    return 'Termin bestätigen';
  }
  
  // Generischer Fallback
  return `Follow-up: ${subject.substring(0, 40)}`;
}

/**
 * Lernt aus Task-for-Me-Korrekturen (z.B. Unsure → No)
 * Speichert Patterns (Domain, Subject-Keywords) für zukünftige Klassifizierung
 */
function learnFromTaskForMeCorrection_(sheet, row, decision) {
  try {
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    const subject = (sheet.getRange(row, 2).getValue() || '').toString().trim(); // B
    if (!messageId) return;

    const msg = GmailApp.getMessageById(messageId);
    const from = (msg.getFrom && msg.getFrom()) || '';
    const domain = ((from.split('@')[1] || '').split('>')[0] || '').toLowerCase().trim();
    
    if (!domain) return;

    const ss = sheet.getParent();
    let sh = ss.getSheetByName('Task Learning');
    if (!sh) {
      sh = ss.insertSheet('Task Learning');
      sh.getRange(1, 1, 1, 5).setValues([['Domain', 'Subject Pattern', 'Decision', 'Count', 'Last Updated']]);
    }

    // Extrahiere wichtige Keywords aus Subject (z.B. "Paket", "Packstation", "Benachrichtigung")
    const subjectLower = subject.toLowerCase();
    const keywords = extractTaskKeywords_(subjectLower);

    // Speichere Domain + Decision
    if (domain) {
      bumpTaskLearning_(sh, domain, '', decision);
    }

    // Speichere Subject-Keywords + Decision
    keywords.forEach(keyword => {
      if (keyword.length > 3) { // Nur relevante Keywords
        bumpTaskLearning_(sh, '', keyword, decision);
      }
    });

    Logger.log(`Task Learning: ${decision} für Domain "${domain}" und Keywords "${keywords.join(', ')}"`);
  } catch (e) {
    Logger.log('learnFromTaskForMeCorrection_ error: ' + e);
  }
}

/**
 * Extrahiert wichtige Keywords aus einem Subject für Task-Learning
 */
function extractTaskKeywords_(subjectLower) {
  // Entferne häufige Stoppwörter
  const stopWords = ['der', 'die', 'das', 'ein', 'eine', 'ist', 'sind', 'für', 'mit', 'von', 'zu', 'in', 'an', 'auf', 'bei', 'über', 'unter', 'durch', 'nach', 'vor', 'seit', 'während', 'wegen', 'trotz', 'ohne', 'gegen', 'um', 'bis', 'ab', 'aus', 'bei', 'hinter', 'neben', 'zwischen', 'innerhalb', 'außerhalb', 'während', 'wegen', 'trotz', 'ohne', 'gegen', 'um', 'bis', 'ab', 'aus', 'bei', 'hinter', 'neben', 'zwischen', 'innerhalb', 'außerhalb'];
  
  // Extrahiere Wörter (mindestens 4 Zeichen)
  const words = subjectLower.split(/\s+/).filter(w => w.length >= 4 && !stopWords.includes(w));
  
  // Spezielle Patterns für häufige Infomails
  const patterns = [];
  if (subjectLower.includes('paket') || subjectLower.includes('packstation')) patterns.push('paket');
  if (subjectLower.includes('benachrichtigung') || subjectLower.includes('notification')) patterns.push('benachrichtigung');
  if (subjectLower.includes('werbung') || subjectLower.includes('newsletter') || subjectLower.includes('angebot')) patterns.push('werbung');
  if (subjectLower.includes('bestätigung') || subjectLower.includes('confirmation')) patterns.push('bestätigung');
  if (subjectLower.includes('rechnung') || subjectLower.includes('invoice')) patterns.push('rechnung');
  
  return [...new Set([...words.slice(0, 3), ...patterns])]; // Max 3 Wörter + Patterns
}

/**
 * Erhöht die Häufigkeit eines Task-Learning-Eintrags
 */
function bumpTaskLearning_(sheet, domain, keyword, decision) {
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const d = (vals[i][0] || '').toString().toLowerCase().trim();
    const k = (vals[i][1] || '').toString().toLowerCase().trim();
    const dec = (vals[i][2] || '').toString().trim();
    if (d === domain.toLowerCase().trim() && k === keyword.toLowerCase().trim() && dec === decision) {
      const n = Number(vals[i][3] || 0) + 1;
      sheet.getRange(i + 1, 3).setValue(decision);
      sheet.getRange(i + 1, 4).setValue(n);
      sheet.getRange(i + 1, 5).setValue(new Date());
      return;
    }
  }
  sheet.appendRow([domain.toLowerCase().trim(), keyword.toLowerCase().trim(), decision, 1, new Date()]);
}

/**
 * Lernt aus Prioritäts-Korrekturen (z.B. Medium → High)
 * Speichert Patterns (Domain, Subject-Keywords) für zukünftige Prioritätsbestimmung
 */
function learnFromPriorityCorrection_(sheet, row, priority) {
  try {
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    const subject = (sheet.getRange(row, 2).getValue() || '').toString().trim(); // B
    if (!messageId) return;

    const msg = GmailApp.getMessageById(messageId);
    const from = (msg.getFrom && msg.getFrom()) || '';
    const domain = ((from.split('@')[1] || '').split('>')[0] || '').toLowerCase().trim();
    
    if (!domain) return;

    const ss = sheet.getParent();
    let sh = ss.getSheetByName('Priority Learning');
    if (!sh) {
      sh = ss.insertSheet('Priority Learning');
      sh.getRange(1, 1, 1, 5).setValues([['Domain', 'Subject Pattern', 'Priority', 'Count', 'Last Updated']]);
    }

    // Extrahiere wichtige Keywords aus Subject
    const subjectLower = subject.toLowerCase();
    const keywords = extractTaskKeywords_(subjectLower); // Wiederverwendung der Keyword-Extraktion

    // Speichere Domain + Priority
    if (domain) {
      bumpPriorityLearning_(sh, domain, '', priority);
    }

    // Speichere Subject-Keywords + Priority
    keywords.forEach(keyword => {
      if (keyword.length > 3) { // Nur relevante Keywords
        bumpPriorityLearning_(sh, '', keyword, priority);
      }
    });

    Logger.log(`Priority Learning: ${priority} für Domain "${domain}" und Keywords "${keywords.join(', ')}"`);
  } catch (e) {
    Logger.log('learnFromPriorityCorrection_ error: ' + e);
  }
}

/**
 * Erhöht die Häufigkeit eines Priority-Learning-Eintrags
 */
function bumpPriorityLearning_(sheet, domain, keyword, priority) {
  const vals = sheet.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const d = (vals[i][0] || '').toString().toLowerCase().trim();
    const k = (vals[i][1] || '').toString().toLowerCase().trim();
    const p = (vals[i][2] || '').toString().trim();
    if (d === domain.toLowerCase().trim() && k === keyword.toLowerCase().trim() && p === priority) {
      const n = Number(vals[i][3] || 0) + 1;
      sheet.getRange(i + 1, 3).setValue(priority);
      sheet.getRange(i + 1, 4).setValue(n);
      sheet.getRange(i + 1, 5).setValue(new Date());
      return;
    }
  }
  sheet.appendRow([domain.toLowerCase().trim(), keyword.toLowerCase().trim(), priority, 1, new Date()]);
}

/**
 * Prüft basierend auf gelernten Patterns, welche Priorität eine E-Mail wahrscheinlich haben sollte
 * @return {string|null} Priorität ('Low', 'Medium', 'High', 'Urgent') oder null wenn unentschieden
 */
function getLearnedPriority_(domain, subject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Priority Learning');
    if (!sh || sh.getLastRow() < 2) return null; // Keine Daten → keine Entscheidung

    const vals = sh.getDataRange().getValues();
    const domainLower = (domain || '').toLowerCase().trim();
    const subjectLower = (subject || '').toLowerCase();
    const keywords = extractTaskKeywords_(subjectLower);

    // Zähle Prioritäten für Domain
    const priorityCounts = { 'Low': 0, 'Medium': 0, 'High': 0, 'Urgent': 0 };
    
    if (domainLower) {
      for (let i = 1; i < vals.length; i++) {
        const d = (vals[i][0] || '').toString().toLowerCase().trim();
        const p = (vals[i][2] || '').toString().trim();
        const count = Number(vals[i][3] || 0);
        if (d === domainLower && priorityCounts.hasOwnProperty(p)) {
          priorityCounts[p] += count;
        }
      }
    }

    // Zähle Prioritäten für Keywords
    keywords.forEach(keyword => {
      for (let i = 1; i < vals.length; i++) {
        const k = (vals[i][1] || '').toString().toLowerCase().trim();
        const p = (vals[i][2] || '').toString().trim();
        const count = Number(vals[i][3] || 0);
        if (k === keyword.toLowerCase() && priorityCounts.hasOwnProperty(p)) {
          priorityCounts[p] += count;
        }
      }
    });

    // Finde die häufigste Priorität (mit Mindestanzahl)
    let maxPriority = null;
    let maxCount = 0;
    for (const [prio, count] of Object.entries(priorityCounts)) {
      if (count > maxCount && count >= 2) { // Mindestens 2 Vorkommen
        maxCount = count;
        maxPriority = prio;
      }
    }

    // Wenn eine Priorität deutlich häufiger ist (mindestens 2x häufiger als die zweithäufigste)
    if (maxPriority) {
      const sorted = Object.entries(priorityCounts)
        .filter(([_, count]) => count > 0)
        .sort(([_, a], [__, b]) => b - a);
      
      if (sorted.length > 1) {
        const [firstPrio, firstCount] = sorted[0];
        const [secondPrio, secondCount] = sorted[1];
        if (firstCount >= secondCount * 2 && firstCount >= 2) {
          return firstPrio;
        }
      } else if (maxCount >= 2) {
        return maxPriority;
      }
    }
    
    return null; // Unentschieden
  } catch (e) {
    Logger.log('getLearnedPriority_ error: ' + e);
    return null;
  }
}

/**
 * Prüft basierend auf gelernten Patterns, ob eine E-Mail wahrscheinlich "No" (kein Task) ist
 */
function shouldBeTaskForMe_(domain, subject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('Task Learning');
    if (!sh || sh.getLastRow() < 2) return null; // Keine Daten → keine Entscheidung

    const vals = sh.getDataRange().getValues();
    const domainLower = (domain || '').toLowerCase().trim();
    const subjectLower = (subject || '').toLowerCase();
    const keywords = extractTaskKeywords_(subjectLower);

    let noCount = 0;
    let yesCount = 0;

    // Zähle "No" und "Yes" für Domain
    if (domainLower) {
      for (let i = 1; i < vals.length; i++) {
        const d = (vals[i][0] || '').toString().toLowerCase().trim();
        const dec = (vals[i][2] || '').toString().trim();
        const count = Number(vals[i][3] || 0);
        if (d === domainLower && dec === 'No') noCount += count;
        if (d === domainLower && dec === 'Yes') yesCount += count;
      }
    }

    // Zähle "No" und "Yes" für Keywords
    keywords.forEach(keyword => {
      for (let i = 1; i < vals.length; i++) {
        const k = (vals[i][1] || '').toString().toLowerCase().trim();
        const dec = (vals[i][2] || '').toString().trim();
        const count = Number(vals[i][3] || 0);
        if (k === keyword.toLowerCase() && dec === 'No') noCount += count;
        if (k === keyword.toLowerCase() && dec === 'Yes') yesCount += count;
      }
    });

    // Wenn deutlich mehr "No" als "Yes": return "No"
    if (noCount > yesCount * 2 && noCount >= 2) return 'No';
    // Wenn deutlich mehr "Yes" als "No": return "Yes"
    if (yesCount > noCount * 2 && yesCount >= 2) return 'Yes';
    
    return null; // Unentschieden
  } catch (e) {
    Logger.log('shouldBeTaskForMe_ error: ' + e);
    return null;
  }
}


syncGoogleTaskForRow_()

