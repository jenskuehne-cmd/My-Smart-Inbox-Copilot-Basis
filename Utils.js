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

  if (rules.AI_ENABLE && actionAnalysis && Array.isArray(actionAnalysis.tasks) && actionAnalysis.tasks.length) {
    const p = (actionAnalysis.tasks[0].priority || '').toLowerCase();
    if (p.startsWith('high') || p.startsWith('urgent')) return 'High';
    if (p.startsWith('low')) return 'Low';
  }

  if ((rules.HIGH_KEYWORDS || []).some(k => txt.includes(k.toLowerCase()))) return 'High';
  if ((rules.LOW_KEYWORDS || []).some(k => txt.includes(k.toLowerCase()))) return 'Low';

  const fromDom = ((fromEmail || '').split('@')[1] || '').toLowerCase();
  if ((rules.DOMAIN_HIGH || []).some(d => fromDom.includes(d.toLowerCase()))) return 'High';

  if (rules.SINGLE_RECIPIENT_BOOST && singleRecipient) return 'High';

  return (CONFIG.PRIORITY_DEFAULT || 'Medium');
}

function isCalendarInvite_(msg) {
  const from = (msg.getFrom() || '').toLowerCase();
  const subject = (msg.getSubject() || '').toLowerCase();

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
// Entfernt vorhandene "Projekt-Labels" am Thread und setzt das neue (aus Spalte A)
function applyGmailLabelFromRow_(sheet, row) {

  try {
    const labelName = (sheet.getRange(row, 1).getValue() || '').toString().trim(); // A
    const messageId = (sheet.getRange(row, 13).getValue() || '').toString().trim(); // M
    Logger.log(`applyGmailLabelFromRow_: row=${row}, label="${labelName}", msgId="${messageId}"`);
    if (!messageId || !labelName) return;

    const msg = GmailApp.getMessageById(messageId);
    const thread = msg.getThread();
    if (!thread) return;

    // vorhandene "Projekt-Labels" entfernen (du kannst die Heuristik spaeter schaerfen)
    const tLabels = thread.getLabels();
    for (const lab of tLabels) {
      const name = lab.getName();
      // processed-by-apps-script etc. erstmal in Ruhe lassen
      if (name === labelName) continue;
      if (CONFIG.IGNORE_LABELS && CONFIG.IGNORE_LABELS.has(name)) continue;
      // alles andere darf weg, wenn du "hart" umlabeln willst:
      // thread.removeLabel(lab);
    }

    // Ziel-Label setzen
    const target = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
    thread.addLabel(target);
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

    // Nur wenn als Task fuer mich markiert UND hohe Prio
    const isTaskYes = ['Yes', 'Ja', 'J', 'Y'].includes(taskForMe);
    const isHighPrio = (priority === 'High' || priority === 'Urgent');

    const listId = getProp_('GOOGLE_TASKS_LIST_ID', '@default');

    // Fall 1: Kein Task mehr gewuenscht → ggf. vorhandenen Task loeschen
    if (!isTaskYes || !isHighPrio) {
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

    // Fall 2: Task gewuenscht & hohe Prio → neu anlegen oder aktualisieren
    const taskTitle = aiTaskTitle || subject || 'Mail Task';

    const taskNotes = [
      notes || '',
      '',
      'Mail-Link:',
      link,
      '',
      'Message-ID: ' + msgId
    ].join('\n');

    if (existingTaskId) {
      // Update bestehenden Task
      const existing = Tasks.Tasks.get(listId, existingTaskId);
      existing.title = taskTitle;
      existing.notes = taskNotes;
      Tasks.Tasks.update(existing, listId, existingTaskId);
      return;
    }

    // Neuer Task
    const newTask = {
      title: taskTitle,
      notes: taskNotes
    };

    const created = Tasks.Tasks.insert(newTask, listId);
    if (created && created.id) {
      taskIdRange.setValue(created.id);
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
    const summary = (sheet.getRange(row, 8).getValue() || '').toString().trim();  // H: AI Summary
    const myProfile = CONFIG.MY_PROFILE || '';

    // Wir verwenden die Summary als "Body-Ersatz"
    const action = getAIActionabilityAnalysis(subject, summary, myProfile);
    let aiTaskTitle = '';

    if (action && Array.isArray(action.tasks) && action.tasks.length) {
      aiTaskTitle = (action.tasks[0].title || '').toString().trim();
    }

    if (!aiTaskTitle) {
      // Fallback: kurzer Titel aus Betreff bauen
      aiTaskTitle = subject || 'Follow up on this email';
    }

    sheet.getRange(row, 5).setValue(aiTaskTitle); // E: Task Suggestion setzen

  } catch (e) {
    Logger.log('ensureTaskSuggestionForRow_ error: ' + e);
  }
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

