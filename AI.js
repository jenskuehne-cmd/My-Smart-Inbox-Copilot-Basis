/*** AI.gs ***/

function getAISummary(emailText) {
  const apiKey = getProp_('GEMINI_API_KEY', '');
  if (!apiKey) return 'API Key Missing';

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  const prompt = `Summarize the following email body concisely in one or two sentences:\n\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 100 }
  };

  try {
    const res = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText() || '{}');
    const txt = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    return txt || 'Could not generate summary.';
  } catch (e) {
    Logger.log('Gemini Summary Error: ' + e);
    return 'Error generating summary.';
  }
}

function getAIResponseSuggestionWithContext(emailText, hasReplied, lastReplyISO) {
  const apiKey = getProp_('GEMINI_API_KEY', '');
  if (!apiKey) return 'API Key Missing';

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  const base = hasReplied
    ? `You already replied on ${lastReplyISO}. Draft a short, polite follow-up if no response has been received since then. Keep it under 80 words.`
    : `Suggest a concise, professional first reply. Keep it under 100 words.`;

  const prompt = `${base}\n\nEmail thread context (latest inbound message):\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.5, maxOutputTokens: 200 }
  };

  try {
    const res = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText() || '{}');
    const txt = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    return txt || 'Could not generate response suggestion.';
  } catch (e) {
    Logger.log('Gemini Reply Error: ' + e);
    return 'Error generating response suggestion.';
  }
}

function getAIActionabilityAnalysis(subject, body, myProfile) {
  const apiKey = getProp_('GEMINI_API_KEY', '');
  if (!apiKey) return { is_task_for_me: 'Unsure', reasons: 'API Key Missing', tasks: [] };

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  // Profil fuers Prompt (Name, E-Mails, Teamkeywords)
  const name = (myProfile?.name || CONFIG.MY_NAME || 'User').toString();
  const emails = (myProfile?.emails || []).join(', ');
  const team = (myProfile?.teamKeywords || CONFIG.MY_TEAM_KEYWORDS || []).join(', ');

  const prompt =
    `You are an assistant that classifies whether an email implies an actionable task for the user, even if not explicitly assigned.
    User identity:
    - name: ${name}
    - emails: ${emails}
    - team_keywords: ${team}

    Given the email subject and body, decide:
    - is_task_for_me: "Yes" | "No" | "Unsure"
    - reasons: short rationale (max 200 chars)
    - suggested_owner: "me" | "someone_else" | "team" | "unknown"
    - tasks: list of {title, owner, due_date, priority}

    Return ONLY strict JSON, no prose.

    Email:
    Subject: ${subject}
    Body:
    ${body}

    JSON schema:
    {
      "is_task_for_me": "Yes|No|Unsure",
      "reasons": "string",
      "suggested_owner": "me|someone_else|team|unknown",
      "tasks": [
        {"title":"string","owner":"string","due_date":"YYYY-MM-DD or empty","priority":"Low|Medium|High"}
      ]
    }`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 256 }
  };

  try {
    const res = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const txt = JSON.parse(res.getContentText() || '{}')?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';

    // JSON robust parsen
    let data;
    try { data = JSON.parse(txt); } catch (_) { data = {}; }

    return {
      is_task_for_me: data.is_task_for_me || 'Unsure',
      reasons: data.reasons || '',
      suggested_owner: data.suggested_owner || 'unknown',
      tasks: Array.isArray(data.tasks) ? data.tasks : []
    };
  } catch (e) {
    Logger.log('Gemini Actionability Error: ' + e);
    return { is_task_for_me: 'Unsure', reasons: 'Error during analysis', tasks: [] };
  }
}

function detectLang_(text) {
  try { return LanguageApp.detectLanguage(text || '') || ''; } catch (_) { return ''; }
}

function getAISummarySmart(emailText) {
  const src = detectLang_(emailText);
  const rule = (CONFIG.AI_LANG?.SUMMARY_IF_SOURCE || []).find(r => r.source === src);
  const target = rule ? rule.target : null;

  const apiKey = getProp_('GEMINI_API_KEY', '');
  if (!apiKey) return 'API Key Missing';

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  const langLine = target ? `Respond in ${target}.` : 'Respond in the same language as the input.';
  const prompt = `${langLine}\nSummarize the following email body concisely in one or two sentences:\n\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 100 }
  };
  try {
    var lastErr = null;
    for (var attempt = 1; attempt <= 3; attempt++) {
      try {
        const res = UrlFetchApp.fetch(apiUrl, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });
        const code = res.getResponseCode();
        const txt = res.getContentText() || '';
        if (code >= 200 && code < 300) {
          const data = JSON.parse(txt || '{}');
          const out = data?.candidates?.[0]?.content?.parts?.[0]?.text;
          if (out && out.trim()) return out.trim();
        } else {
          lastErr = 'HTTP ' + code + ' ' + txt.slice(0, 200);
        }
      } catch (e) {
        lastErr = e && e.message ? e.message : '' + e;
      }
      Utilities.sleep(300 + Math.floor(Math.random() * 300)); // leichter Backoff
    }
    Logger.log('Gemini Summary Error: ' + (lastErr || 'unknown'));
    // Fallback-Text, falls alle Versuche fehlschlagen:
    return 'Keine AI-Summary verfügbar (Netzwerk/Limit).';
  } catch (e) {
    Logger.log('Gemini Summary Fatal: ' + e);
    return 'Keine AI-Summary verfügbar (Error).';
  }
}
function getAIResponseSuggestionWithLang(emailText, hasReplied, lastReplyISO, replyLangHint) {
  const apiKey = getProp_('GEMINI_API_KEY', '');
  if (!apiKey) return 'API Key Missing';

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  const base = hasReplied
    ? `You already replied on ${lastReplyISO}. Draft a short, polite follow-up if no response has been received since then. Keep it under 80 words.`
    : `Suggest a concise, professional first reply. Keep it under 100 words.`;

  const langLine = replyLangHint ? `Write the reply in ${replyLangHint}.` : 'Write the reply in the same language as the latest inbound message.';
  const prompt = `${base}\n${langLine}\n\nEmail thread context (latest inbound message):\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.5, maxOutputTokens: 200 }
  };
  try {
    const res = UrlFetchApp.fetch(apiUrl, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText() || '{}');
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || 'Could not generate response suggestion.';
  } catch (e) {
    Logger.log('Gemini Reply Error: ' + e);
    return 'Error generating response suggestion.';
  }

} 

function getAvailableLabels_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Gmail Labels');
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  return vals
    .map(r => (r[0] || '').toString().trim())
    .filter(Boolean); // keine leeren
}

function getAILabelSuggestion(subject, body) {
  const apiKey = getProp_('GEMINI_API_KEY', '');
  if (!apiKey) return '';

  const labels = getAvailableLabels_();
  if (!labels.length) return '';

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  const labelList = labels.join(' | ');
  const prompt =
    `Du bist ein Assistent, der eingehende E-Mails einem einzigen Label zuordnet. ` +
    `Du bekommst eine Liste moeglicher Labels und sollst GENAU EIN Label daraus waehlen, das am besten passt. ` +
    `Wenn nichts gut passt, waehle "Uncategorized". ` +
    `Antworte NUR mit dem exakten Labelnamen (ohne Erklaerung).\n\n` +
    `Verfuegbare Labels:\n${labelList}\n\n` +
    `E-Mail:\n` +
    `Betreff: ${subject || '(kein Betreff)'}\n` +
    `Body:\n${body || ''}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.1, maxOutputTokens: 20 }
  };

  try {
    const res = UrlFetchApp.fetch(apiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const txt = res.getContentText() || '';

    if (code < 200 || code >= 300) {
      Logger.log('Gemini Label Suggestion HTTP ' + code + ' ' + txt.slice(0, 200));
      return '';
    }

    const data = JSON.parse(txt || '{}');
    let out = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    if (!out) return '';

    out = out.trim();

    // Falls das Model doch Text drumherum liefert -> auf bekannten Labelnamen matchen
    const lowerOut = out.toLowerCase();
    const exact = labels.find(l => l.toLowerCase() === lowerOut);
    if (exact) return exact;

    for (const l of labels) {
      if (out.includes(l)) return l;
      if (lowerOut.includes(l.toLowerCase())) return l;
    }

    // Fallback: wenn "Uncategorized" in der Liste ist, nimm das
    const unc = labels.find(l => l.toLowerCase() === 'uncategorized');
    return unc || '';
  } catch (e) {
    Logger.log('Gemini Label Suggestion Error: ' + e);
    return '';
  }
}


// end of file AI.gs
