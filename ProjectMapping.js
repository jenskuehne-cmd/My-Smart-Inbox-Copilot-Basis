/*** ProjectMapping.gs ***/
/*
 * Zuständig fuer die Projektzuordnung:
 * 1) Labels (mit Prioritaet)
 * 2) Project Map (Regex → Domains → Keywords)
 * 3) Fallback-Keywords
 *
 * WICHTIG:
 * - Keine Ignore-Labels hier. Wenn du nichts ignorieren willst, brauchst du nichts weiter zu tun.
 * - Passe CONFIG.LABEL_PRIORITY und CONFIG.LABEL_TO_PROJECT in Config.gs an deine Labelnamen an.
 */

// 1) Projekt aus Labels bestimmen (mit Prioritaet)
function identifyProjectByLabels_(threadLabels) {
  if (!CONFIG.ENABLE_LABEL_MAPPING) return null;
  if (!threadLabels || threadLabels.length === 0) return null;

  const ignore = CONFIG.IGNORE_LABELS || new Set();
  const rawNames = [];
  for (const lab of threadLabels) {
    const name = lab.getName();
    if (ignore.has(name)) continue;
    rawNames.push(name);
  }
  if (rawNames.length === 0) return null;

  const mode = CONFIG.LABEL_MATCH_MODE || 'regex_first';
  const candidates = new Set();

  const norm = s => (s || '')
    .toLowerCase()
    .replace(/[._]/g, ' ')
    .replace(/[\s]*-[\s]*/g, '-')     // z. B. "6 - 1" → "6-1"
    .replace(/\s+/g, ' ')
    .trim();

  const stripNums = s => s.replace(/(^|\s)[0-9]+([.-][0-9]+)*\s*/g, '');
  const fixTypos = s => s.replace(/\bbpm\b/gi, 'bpm');

  function matchByRegex(names) {
    const list = CONFIG.LABEL_REGEX_MAP || [];
    for (const n of names) {
      for (const r of list) {
        try {
          const re = new RegExp(r.re, 'i');
          if (re.test(n)) candidates.add(r.project);
        } catch (e) {
          Logger.log('Bad LABEL_REGEX_MAP entry: ' + r.re);
        }
      }
    }
  }

  function matchBySegments(names) {
    const segMap = CONFIG.LABEL_SEGMENT_MAP || {};
    for (const n of names) {
      const cleaned = norm(fixTypos(stripNums(n)));
      const segs = cleaned.split('/').map(s => norm(stripNums(s)));
      for (const seg of segs) {
        if (segMap[seg]) candidates.add(segMap[seg]);
        Object.keys(segMap).forEach(k => { if (seg.includes(k) && !segMap[seg]) candidates.add(segMap[k]); });
      }
      const parts = cleaned.split('/').map(p => p.trim());
      for (let i = 1; i <= parts.length; i++) {
        const prefix = parts.slice(0, i).join(' / ');
        Object.keys(segMap).forEach(k => { if (prefix.includes(k)) candidates.add(segMap[k]); });
      }
    }
  }

  function matchByExact(names) {
    const exact = CONFIG.LABEL_TO_PROJECT || {};
    names.forEach(n => { if (exact[n]) candidates.add(exact[n]); });
  }

  const names = rawNames;

  if (mode === 'regex_first') {
    matchByRegex(names); matchBySegments(names); matchByExact(names);
  } else if (mode === 'segment_first') {
    matchBySegments(names); matchByRegex(names); matchByExact(names);
  } else {
    matchByExact(names);
  }

  if (candidates.size === 0) return null;
  const prio = CONFIG.PROJECT_PRIORITY || [];
  for (const p of prio) if (candidates.has(p)) return p;
  return Array.from(candidates)[0];
}


// 2) Project Map aus dem Sheet laden (Tab "Project Map")
function loadProjectMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.PROJECT_MAP_SHEET);
  if (!sh) return [];
  const values = sh.getDataRange().getValues(); // inkl. Header
  const rows = values.slice(1).filter(r => r[0]); // ab Zeile 2, nur mit Project
  return rows.map(r => ({
    project: (r[0] || '').toString().trim(),
    keywords: toList_(r[1]),
    domains: toList_(r[2]),
    regex: (r[3] || '').toString().trim()
  }));
}

// 3) Smarte Zuordnung: Regex → Domains → Keywords (aus Project Map)
function identifyProjectSmart_(cfg, subject, body, fromEmail, toEmails) {
  const lowerSub = (subject || '').toLowerCase();
  const lowerBody = (body || '').toLowerCase();
  const allText = lowerSub + '\n' + lowerBody;

  const fromDomain = (fromEmail.split('@')[1] || '').toLowerCase();
  const toDomains = (toEmails || []).map(a => (a.split('@')[1] || '').toLowerCase());

  // 3a) Regex (z. B. Ticketpraefixe "ASPIRE-123")
  for (const c of cfg) {
    if (c.regex) {
      const re = safeRegex_(c.regex);
      if (re && re.test(allText)) return c.project;

    }
  }
  // 3b) Absender/Empfaenger-Domains
  for (const c of cfg) {
    if (c.domains.some(d => d && (fromDomain === d || toDomains.includes(d)))) {
      return c.project;
    }
  }
  // 3c) Keywords (Wortgrenzen)
  for (const c of cfg) {
    if (c.keywords.some(k => k && new RegExp(`\\b${escapeRegExp_(k)}\\b`, 'i').test(allText))) {
      return c.project;
    }
  }
  return null;
}

// 4) Fallback: statische Keywordliste (CONFIG.PROJECT_KEYWORDS)
function identifyByKeywordsFallback_(keywords, subject, body) {
  const text = `${subject}\n${body}`;
  for (const k of keywords) {
    const re = new RegExp(`\\b${escapeRegExp_(k)}\\b`, 'i');
    if (re.test(text)) return k;
  }
  return null;
}

// Helpers
function toList_(cell) {
  return (cell || '')
    .toString()
    .toLowerCase()
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}
function escapeRegExp_(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function safeRegex_(pattern) {
  if (!pattern) return null;
  const pat = String(pattern).trim();
  if (!pat) return null;
  try {
    return new RegExp(pat, 'i'); // Flag i kommt hier rein, NICHT im Sheet
  } catch (e) {
    Logger.log('Invalid regex in Project Map: "' + pat + '" → ' + e);
    return null; // einfach ueberspringen statt Skript abbrechen
  }
}
// === Historie laden (optional) — Sheet "Project History": Header: Domain | Project | Count
function loadHistoryMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Project History');
  if (!sh) return new Map();
  const vals = sh.getDataRange().getValues();
  const m = new Map();
  for (let i = 1; i < vals.length; i++) {
    const d = (vals[i][0] || '').toString().toLowerCase().trim();
    const p = (vals[i][1] || '').toString().trim();
    const c = Number(vals[i][2] || 0);
    if (d && p) m.set(`${d}::${p}`, c);
  }
  return m;
}

// Historie erhoehen (lernt: Domain → Project)
function bumpHistory_(domain, project) {
  if (!domain || !project) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Project History');
  if (!sh) {
    sh = ss.insertSheet('Project History');
    sh.getRange(1, 1, 1, 3).setValues([['Domain', 'Project', 'Count']]);
  }
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const d = (vals[i][0] || '').toString().toLowerCase().trim();
    const p = (vals[i][1] || '').toString().trim();
    if (d === domain.toLowerCase().trim() && p === project) {
      const n = Number(vals[i][2] || 0) + 1;
      sh.getRange(i + 1, 3).setValue(n);
      return;
    }
  }
  sh.appendRow([domain.toLowerCase().trim(), project, 1]);
}

// === Scoring: Labels + Project Map + Hints + Historie
function identifyProjectScored_(threadLabels, projMap, subject, body, fromEmail, toEmails) {
  // Defaults, falls CONFIG.PROJECT_SCORING/HINTS noch nicht gesetzt sind
  const PS = CONFIG.PROJECT_SCORING || {
    WEIGHTS: {
      label_priority: 120, label_mapped: 90, regex: 60, domain: 35,
      keyword_subject: 18, keyword_body: 8, boost_subject: 22, boost_body: 12,
      negative_subject: -35, negative_body: -20, history_per_hit: 3
    },
    MIN_SCORE: 25
  };
  const W = PS.WEIGHTS;
  const H = CONFIG.PROJECT_HINTS || {};

  const textSub = (subject || '').toLowerCase();
  const textBody = (body || '').toLowerCase();
  const fromDomain = (fromEmail.split('@')[1] || '').toLowerCase();
  const toDomains = (toEmails || []).map(a => (a.split('@')[1] || '').toLowerCase());
  const history = loadHistoryMap_();

  // Kandidaten aus Project Map + Hints + gelernten Projekten (History)
  const candidates = new Set();

  // 1) Projekte aus der Project Map
  (projMap || []).forEach(c => candidates.add(c.project));

  // 2) Projekte, fuer die es Hints gibt
  Object.keys(H).forEach(p => candidates.add(p));

  // 3) NEU: alle in "Project History" gelernten Projektnamen auch als Kandidaten zulassen
  if (history && typeof history.forEach === 'function') {
    history.forEach((count, key) => {
      // key = "domain::project"
      const parts = (key || '').split('::');
      if (parts[1]) {
        candidates.add(parts[1]);
      }
    });
  }

  // Fallback, falls du mal gar keine Konfig definiert hast
  if (candidates.size === 0) {
    ['Aspire', 'Digitalisierung', 'Doc SPOC'].forEach(p => candidates.add(p));
  }


  // schnelle Map: project → projectMapRow
  const byProject = new Map();
  (projMap || []).forEach(c => byProject.set(c.project, c));

  // Label-Mapping (wie bisher): prioritaet → erstes gemapptes
  let labelProject = null;
  if (threadLabels && threadLabels.length) {
    if (CONFIG.LABEL_PRIORITY && CONFIG.LABEL_PRIORITY.length > 0) {
      for (const wanted of CONFIG.LABEL_PRIORITY) {
        if (threadLabels.some(l => l.getName() === wanted && CONFIG.LABEL_TO_PROJECT[wanted])) {
          labelProject = CONFIG.LABEL_TO_PROJECT[wanted]; break;
        }
      }
    }
    if (!labelProject) {
      for (const lab of threadLabels) {
        const name = lab.getName();
        if (CONFIG.LABEL_TO_PROJECT[name]) { labelProject = CONFIG.LABEL_TO_PROJECT[name]; break; }
      }
    }
  }

  // Scoring pro Kandidat
  let best = { project: 'Uncategorized', score: -1 };

  candidates.forEach(project => {
    let score = 0;

    // Label-Bonus
    if (labelProject && labelProject === project) {
      const isPriority = CONFIG.LABEL_PRIORITY && CONFIG.LABEL_PRIORITY.includes(project);
      score += isPriority ? (W.label_priority || 120) : (W.label_mapped || 90);
    }

    // Project Map: regex / domains / keywords
    const cfg = byProject.get(project);
    if (cfg) {
      const re = safeRegex_(cfg.regex);
      if (re && (re.test(textSub) || re.test(textBody))) score += (W.regex || 60);

      if (cfg.domains && cfg.domains.length > 0) {
        if (cfg.domains.some(d => d && (fromDomain === d || toDomains.includes(d)))) score += (W.domain || 35);
      }

      (cfg.keywords || []).forEach(k => {
        if (!k) return;
        const r = new RegExp(`\\b${escapeRegExp_(k)}\\b`, 'i');
        if (r.test(textSub)) score += (W.keyword_subject || 18);
        if (r.test(textBody)) score += (W.keyword_body || 8);
      });
    }

    // Hints (Boost/Negative)
    const hint = H[project] || { boost: [], negative: [] };
    (hint.boost || []).forEach(b => {
      if (!b) return;
      if (textSub.includes(b)) score += (W.boost_subject || 22);
      if (textBody.includes(b)) score += (W.boost_body || 12);
    });
    (hint.negative || []).forEach(n => {
      if (!n) return;
      if (textSub.includes(n)) score += (W.negative_subject || -35);
      if (textBody.includes(n)) score += (W.negative_body || -20);
    });

    // Historie: Domain→Project
    const hkey = `${fromDomain}::${project}`;
    const hits = history.get(hkey) || 0;
    if (hits > 0) score += hits * (W.history_per_hit || 3);

    if (score > best.score) best = { project, score };
  });

  return (best.score >= (PS.MIN_SCORE || 25)) ? best.project : 'Uncategorized';
}
