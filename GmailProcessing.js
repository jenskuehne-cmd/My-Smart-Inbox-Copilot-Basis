/*** GmailProcessing.gs ***/
/*
 * Oeffentliche Starter:
 * - processEmailsToday()       → verarbeitet ab heutigem Tagesbeginn
 * - processEmailsLastNDays(n)  → verarbeitet die letzten n Tage
 * - processEmailsDateRange(a,b)→ verarbeitet einen Datumsbereich (inkl. beider Tage)
 *
 * Kernlogik:
 * - processThreads_(threads)   → verarbeitet die gefundenen Threads/Mails
 *
 * Voraussetzungen:
 * - Config.gs (CONFIG, getProcessedLabel_())
 * - ProjectMapping.gs (identifyProjectByLabels_, loadProjectMap_, identifyProjectSmart_, identifyByKeywordsFallback_)
 * - AI.gs (getAISummary, getAIResponseSuggestionWithContext)
 * - Utils.gs (getMyEmails_, getReplyInfoForThread_, extractEmail_, parseISODate_ falls im Menue genutzt)
 */

// Verarbeite: Heute
function processEmailsToday() {
  const tz = CONFIG.TIMEZONE;
  const now = new Date();
  const startOfToday = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd') + ' 00:00:00');
  const afterStr = Utilities.formatDate(startOfToday, tz, 'yyyy/MM/dd');

  const query = `after:${afterStr}`; // Basis: ab heute 00:00
  const threads = GmailApp.search(query, 0, CONFIG.PROCESS_LIMIT_THREADS);

  processThreads_(threads);
}

// Verarbeite: Letzte n Tage (Backfill light)
function processEmailsLastNDays(n) {
  const today = new Date();
  const start = new Date(today.getTime() - (n * 24 * 3600 * 1000));
  processEmailsDateRange(start, today);
}

// Verarbeite: Datumsbereich (inkl. beider Tage)
function processEmailsDateRange(startDate, endDate) {
  const tz = CONFIG.TIMEZONE;
  const afterStr = Utilities.formatDate(startDate, tz, 'yyyy/MM/dd');
  // Gmail before: exklusiv → Tag nach Enddatum nehmen
  const nextDay = new Date(endDate.getTime() + 24 * 3600 * 1000);
  const beforeStr = Utilities.formatDate(nextDay, tz, 'yyyy/MM/dd');

  // Beispiel: Inbox, Promotions/Social raus
  const query = `after:${afterStr} before:${beforeStr} label:inbox -category:promotions -category:social`;
  const threads = GmailApp.search(query, 0, CONFIG.BACKFILL_LIMIT_THREADS);

  processThreads_(threads);
}

// Kernverarbeitung fuer alle Varianten
function processThreads_(threads) {
  // API-Key Validierung am Anfang
  const apiTest = testGeminiAPIKey();
  if (!apiTest.valid || apiTest.quotaExceeded) {
    const ui = SpreadsheetApp.getUi();
    let title = '⚠️ Gemini API-Key Problem';
    let message = apiTest.message;
    
    if (apiTest.quotaExceeded) {
      title = '⚠️ API-Limit erreicht';
      message = `Der Gemini API-Key ist gültig, aber das Free-Tier-Limit wurde erreicht.\n\n${apiTest.message}\n\nMöchten Sie trotzdem fortfahren? (Die E-Mails werden ohne KI-Funktionen verarbeitet.)`;
    } else {
      message = `Der Gemini API-Key funktioniert nicht:\n\n${apiTest.message}\n\nMöchten Sie trotzdem fortfahren? (Die E-Mails werden ohne KI-Funktionen verarbeitet.)`;
    }
    
    const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.NO) {
      Logger.log('Verarbeitung abgebrochen: API-Key-Problem');
      return;
    }
    // Wenn YES: weiter mit Warnung, aber ohne KI-Funktionen
    if (apiTest.quotaExceeded) {
      Logger.log('WARNUNG: Verarbeitung ohne KI-Funktionen (API-Limit erreicht)');
    } else {
      Logger.log('WARNUNG: Verarbeitung ohne funktionierenden API-Key');
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log(`Sheet "${CONFIG.SHEET_NAME}" nicht gefunden.`);
    return;
  }

  // Spalten-Setup (A–N = 1–14)
  const COL_MESSAGE_ID = 13; // M
  const COL_COST = 14;        // N: API-Kosten
  const COLS_TOTAL = 14;      // bis N

  // Bereits vorhandene Message-IDs (Duplikatcheck)
  const existingIds = new Set();

  // 1) IDs aus current activities
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, COL_MESSAGE_ID, lastRow - 1, 1).getValues();
    ids.forEach(r => {
      const id = (r[0] || '').toString().trim();
      if (id) existingIds.add(id);
    });
  }

  // 2) IDs aus dem Archiv-Sheet auch beruecksichtigen
  try {
    const arch = ss.getSheetByName(CONFIG.ARCHIVE_SHEET);
    if (arch) {
      const archLastRow = arch.getLastRow();
      if (archLastRow > 1) {
        const archIds = arch.getRange(2, COL_MESSAGE_ID, archLastRow - 1, 1).getValues();
        archIds.forEach(r => {
          const id = (r[0] || '').toString().trim();
          if (id) existingIds.add(id);
        });
      }
    }
  } catch (e) {
    Logger.log('Fehler beim Einlesen der Archiv-IDs: ' + e);
  }


  // --- ALT-Status "New" auf "Open" setzen, wenn nicht heute ---
  try {
    if (lastRow > 1) {
      const allRows = sheet.getRange(2, 1, lastRow - 1, COLS_TOTAL).getValues();
      const tz = CONFIG.TIMEZONE;
      const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

      for (let r = 0; r < allRows.length; r++) {
        const status = (allRows[r][2] || '').toString().trim(); // Spalte C
        const isoStr = (allRows[r][6] || '').toString();        // G (Notes, enthaelt ISO)
        if (status === 'New') {
          const match = isoStr.match(/\d{4}-\d{2}-\d{2}/);
          if (match) {
            const mailDate = match[0];
            if (mailDate !== todayStr) {
              sheet.getRange(r + 2, 3).setValue(CONFIG.STATUS_DEFAULT_OLD || 'Open'); // C = Status
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Status-Update-Check: ' + e);
  }

  // Optionales "processed"-Label vorbereiten (aus Config/Property)
  const processedLabelName = getProcessedLabel_();
  const setProcessedLabel = processedLabelName && processedLabelName.length > 0;
  const processedLabel = setProcessedLabel
    ? (GmailApp.getUserLabelByName(processedLabelName) || GmailApp.createLabel(processedLabelName))
    : null;

  // Meine Absender (fuer Reply-Erkennung)
  const myEmails = CONFIG.ENABLE_REPLY_TRACKING ? getMyEmails_() : [];

  // Project Map (Konfigblatt) einmal laden
  const projMap = loadProjectMap_();

  // Chronologische Reihenfolge erzwingen (aeltestes zuerst)
  try {
    threads.sort((a, b) => a.getLastMessageDate() - b.getLastMessageDate());
  } catch (e) {
    Logger.log('Sort threads by date failed: ' + e);
  }

  // Durch Threads laufen
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];

    // Labels einmal pro Thread holen (Thread-Ebene!)
    const labels = thread.getLabels();

    // Alle Messages im Thread
    const messages = thread.getMessages();

    // Habe ich nach der letzten eingehenden Mail geantwortet?
    const replyInfo = CONFIG.ENABLE_REPLY_TRACKING
      ? getReplyInfoForThread_(messages, myEmails)
      : { hasReplied: false, myLastReplyDate: null };

    // Projekt aus Labels (wenn moeglich) EINMAL pro Thread bestimmen
    const labelProject = identifyProjectByLabels_(labels);

    // Jede Nachricht des Threads verarbeiten
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const messageId = msg.getId();

      // Duplikate ueberspringen
      if (existingIds.has(messageId)) continue;

      // Nachrichtendaten
      const subject = (msg.getSubject() || '').trim();

      // Best-Body: nimmt Plain wenn vorhanden, sonst HTML-strip
      const rawBody = getBestBody_(msg);
      const body = rawBody.length > CONFIG.MAX_BODY_CHARS
        ? rawBody.slice(0, CONFIG.MAX_BODY_CHARS)
        : rawBody;

      const emailLink = `https://mail.google.com/mail/u/0/#inbox/${messageId}`;
      const receivedDate = msg.getDate();
      const receivedISO = receivedDate ? receivedDate.toISOString() : new Date().toISOString();

      // Absender/Empfaenger frueh extrahieren
      const fromEmail = extractEmail_(msg.getFrom());
      const toEmails = (msg.getTo() || '').split(',').map(s => extractEmail_(s));

      // Heute-Check fuer Status-Default
      const tz = CONFIG.TIMEZONE;
      const isToday = receivedDate &&
        Utilities.formatDate(receivedDate, tz, 'yyyy-MM-dd') ===
        Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

      const calendarInvite = isCalendarInvite_(msg);

      const receivedCH = receivedDate
        ? Utilities.formatDate(receivedDate, CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm z')
        : Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm z');

      // --- Projektbestimmung (Scoring) als Fallback / Lernsignal ---
      let identifiedProject = identifyProjectScored_(labels, projMap, subject, body, fromEmail, toEmails);

      // Optional: wenn immer noch "Uncategorized", letzter Fallback Keywords
      if (identifiedProject === 'Uncategorized') {
        identifiedProject = identifyByKeywordsFallback_(CONFIG.PROJECT_KEYWORDS, subject, body) || 'Uncategorized';
      }

      // --- NEU: KI-Labelvorschlag basierend auf Subject & Body ---
      let aiLabel = '';
      if (apiTest.valid && !apiTest.quotaExceeded) {
        try {
          aiLabel = getAILabelSuggestion(subject, body) || '';
        } catch (e) {
          Logger.log('AI label suggestion failed: ' + e);
          aiLabel = '';
        }
      }

      // Finales Label, das in Spalte A landet (und spaeter nach Gmail geht)
      const finalLabel = aiLabel || identifiedProject || 'Uncategorized';

      // Historie updaten: Domain → Projekt/Label
      const fromDomain = (fromEmail.split('@')[1] || '').toLowerCase();
      // Wenn du Labels als Projekte verstehst, kannst du hier finalLabel nehmen:
      const projectForHistory =
        identifiedProject && identifiedProject !== 'Uncategorized'
          ? identifiedProject
          : finalLabel;

      if (fromDomain && projectForHistory && projectForHistory !== 'Uncategorized') {
        bumpHistory_(fromDomain, projectForHistory);
      }

      // --- AI Actionability (kontextuelle Aufgaben-Erkennung) ---
      const myProfile = {
        name: getProp_('MY_NAME', '') || (CONFIG.MY_NAME || 'User'),
        emails: (CONFIG.ENABLE_REPLY_TRACKING ? getMyEmails_() : []),
        teamKeywords: (getProp_('MY_TEAM_KEYWORDS', '')
          .split(',')
          .map(s => s.trim())
          .filter(Boolean))
          || (CONFIG.MY_TEAM_KEYWORDS || [])
      };

      let action = { is_task_for_me: 'Unsure', reasons: '', suggested_owner: 'unknown', tasks: [] };
      if (apiTest.valid && !apiTest.quotaExceeded) {
        // Prüfe zuerst gelernte Patterns (falls vorhanden)
        const fromDomain = (fromEmail.split('@')[1] || '').toLowerCase();
        const learnedDecision = shouldBeTaskForMe_(fromDomain, subject);
        
        if (learnedDecision === 'No') {
          // Gelernt: Diese Art von Mail ist kein Task → überspringe AI-Analyse
          action = { is_task_for_me: 'No', reasons: 'Basierend auf gelernten Patterns (Infomail)', suggested_owner: 'unknown', tasks: [] };
          Logger.log(`Task Learning: Überspringe AI-Analyse für "${subject}" (gelernt: No)`);
        } else {
          action = getAIActionabilityAnalysis(subject, body, myProfile); // { is_task_for_me, reasons, ... }
          
          // Wenn AI "Unsure" sagt, aber gelernte Patterns "Yes" → verwende "Yes"
          if (action.is_task_for_me === 'Unsure' && learnedDecision === 'Yes') {
            action.is_task_for_me = 'Yes';
            action.reasons = (action.reasons || '') + ' (bestätigt durch gelernte Patterns)';
            Logger.log(`Task Learning: AI sagte "Unsure", aber gelernte Patterns sagen "Yes" → verwende "Yes"`);
          }
        }
      } else {
        action.reasons = apiTest.quotaExceeded ? 'API-Limit erreicht' : 'API-Key nicht verfügbar';
      }

      const taskForMe = action.is_task_for_me || 'Unsure';
      const notesExtra = action.reasons ? ` | AI-Task: ${action.reasons}` : '';
      const notes = `Empfangen: ${receivedCH} | ISO: ${receivedISO}${notesExtra}`;

      // Erster KI-Task-Titel (falls vorhanden)
      const aiTaskTitle =
        action && Array.isArray(action.tasks) && action.tasks.length
          ? (action.tasks[0].title || '')
          : '';


      // Regelbasierte Prioritaet
      const priority = computePriority_(subject, body, fromEmail, toEmails, action);

      // Sprachhinweis fuer Reply bestimmen
      const detectedLang = detectLang_(body);
      let replyLangHint = null;
      switch (CONFIG.AI_LANG?.REPLY_LANGUAGE_MODE) {
        case 'fixed':
          replyLangHint = CONFIG.AI_LANG?.REPLY_FIXED_LANG || null;
          break;
        case 'sender':
          replyLangHint = detectedLang || null;
          break;
        case 'none':
        default:
          replyLangHint = null;
          break;
      }

      // AI Summary / Reply + Status-Default
      let aiSummary = '';
      let aiReply = '';
      let status = isToday ? (CONFIG.STATUS_DEFAULT || 'New') : (CONFIG.STATUS_DEFAULT_OLD || 'Open');
      let totalCost = ''; // Kosten für diese E-Mail

      if (calendarInvite) {
        aiSummary = 'Kalendereinladung erkannt – bitte im Kalender bestaetigen/ablehnen.';
        status = 'Calendar Invite';
      } else {
        if (apiTest.valid && !apiTest.quotaExceeded) {
          // Kosten-Objekt zurücksetzen
          globalApiCosts = {};
          
          aiSummary = getAISummarySmart(body);
          const lastReplyISO = replyInfo.myLastReplyDate ? replyInfo.myLastReplyDate.toISOString() : '';
          
          // Prüfe, ob es eine Infomail ist → dann kein Reply-Vorschlag
          const isInfoMail = shouldGenerateReply_(subject, body, action);
          if (isInfoMail) {
            aiReply = ''; // Kein Reply-Vorschlag für Infomails
            Logger.log(`Reply-Suggestion übersprungen: Infomail erkannt ("${subject}")`);
          } else {
            aiReply = getAIResponseSuggestionWithLang(
              body,
              !!replyInfo.hasReplied,
              lastReplyISO,
              replyLangHint
            );
          }
          
          // Kosten sammeln und berechnen
          const config = getGeminiAPIConfig_();
          const model = config.model || 'gemini-2.5-pro';
          let totalUsage = { prompt_tokens: 0, completion_tokens: 0, reasoning_tokens: 0, total_tokens: 0 };
          
          if (globalApiCosts['summary']) {
            totalUsage.prompt_tokens += globalApiCosts['summary'].prompt_tokens || 0;
            totalUsage.completion_tokens += globalApiCosts['summary'].completion_tokens || 0;
            totalUsage.reasoning_tokens += globalApiCosts['summary'].reasoning_tokens || 0;
            totalUsage.total_tokens += globalApiCosts['summary'].total_tokens || 0;
          }
          if (globalApiCosts['reply']) {
            totalUsage.prompt_tokens += globalApiCosts['reply'].prompt_tokens || 0;
            totalUsage.completion_tokens += globalApiCosts['reply'].completion_tokens || 0;
            totalUsage.reasoning_tokens += globalApiCosts['reply'].reasoning_tokens || 0;
            totalUsage.total_tokens += globalApiCosts['reply'].total_tokens || 0;
          }
          if (globalApiCosts['actionability']) {
            totalUsage.prompt_tokens += globalApiCosts['actionability'].prompt_tokens || 0;
            totalUsage.completion_tokens += globalApiCosts['actionability'].completion_tokens || 0;
            totalUsage.reasoning_tokens += globalApiCosts['actionability'].reasoning_tokens || 0;
            totalUsage.total_tokens += globalApiCosts['actionability'].total_tokens || 0;
          }
          
          if (totalUsage.total_tokens > 0) {
            totalCost = calculateAPICost_(totalUsage, model);
          }
        } else {
          const reason = apiTest.quotaExceeded ? 'API-Limit erreicht' : 'API-Key-Problem';
          aiSummary = `Keine AI-Summary verfügbar (${reason})`;
          aiReply = `Keine AI-Antwortvorschläge verfügbar (${reason})`;
        }

        if (CONFIG.ENABLE_REPLY_TRACKING && replyInfo.hasReplied) {
          status = 'Replied';
        }
      }

      // Zeile (A–N) schreiben
      const row = [
        finalLabel,                               // A
        subject || '(no subject)',                // B
        status,                                   // C
        priority,                                 // D
        aiTaskTitle,                              // E: KI-Task-Vorschlag
        '',                                       // F (Reserve)
        notes,                                    // G
        aiSummary,                                // H
        taskForMe,                                // I
        aiReply,                                  // J
        new Date(),                               // K
        `=HYPERLINK("${emailLink}","Email Link")`,// L
        messageId,                                // M
        totalCost                                 // N: API-Kosten
      ];


      sheet.appendRow(row);
      existingIds.add(messageId);

      if (processedLabel) {
        thread.addLabel(processedLabel);
      }
      // msg.markRead(); // optional
    }
  }

  Logger.log('Verarbeitung abgeschlossen.');
}

