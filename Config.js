/*** Config.gs ***/
// Zentrale Konfiguration. Hier passt du am einfachsten vieles an.

const CONFIG = {
  SHEET_NAME: 'current activities',
  ARCHIVE_SHEET: 'activity archive',
  PROJECT_MAP_SHEET: 'Project Map',

  // Standardverhalten
  PROCESS_LIMIT_THREADS: 100,   // taeglich
  BACKFILL_LIMIT_THREADS: 500,  // Backfill/Zeitraum
  MAX_BODY_CHARS: 10000,
  TIMEZONE: Session.getScriptTimeZone() || 'Europe/Zurich',

  // Defaults fuer neue Zeilen
  STATUS_DEFAULT: 'New',
  PRIORITY_DEFAULT: 'Medium',
  STATUS_DEFAULT_OLD: 'Open', // wenn Mail nicht heute eingegangen ist





  // Reply-Tracking ermoeglichen (setzt Status „Replied“, schreibt Meta optional in weitere Spalten)
  ENABLE_REPLY_TRACKING: true,

  // Label → Projekt Mapping (empfohlen: Gmail-Filter vergeben)
  ENABLE_LABEL_MAPPING: true,


  // Wenn mehrere Projekt-Labels auf einem Thread sind: diese Reihenfolge bestimmt die Wahl
  LABEL_PRIORITY: [
    'Aspire',
    'Reliability Engineering',
    'Digitalisierung',
    'Capa',                // Labelname exakt wie in Gmail
    'Doc Spoc',
    'RAAD Challenge',
    'Leitung',
    'Ausbildung Coaching',
    'BPm'                  // falls so geschrieben
  ],

  // Fallback-Keywords, wenn weder Label noch Project Map greifen
  PROJECT_KEYWORDS: [
    'Aspire',
    'Reliability Engineering',
    'CAPA',
    'Doc SPOC',
    'RAAD Challenge',
    'Leitung',
    'Ausbildung Coaching',
    'BPM'
  ],

  // Projekt-Scoring (Gewichte)
  PROJECT_SCORING: {
    WEIGHTS: {
      label_priority: 120,   // Label steht in LABEL_PRIORITY
      label_mapped: 90,      // irgendein gemapptes Label
      regex: 60,             // Treffer in "Regex (optional)" der Project Map
      domain: 35,            // Absender/Empfaenger-Domain matcht Project Map
      keyword_subject: 18,   // Keyword-Treffer im Betreff (Project Map)
      keyword_body: 8,       // Keyword-Treffer im Body (Project Map)
      boost_subject: 22,     // Boost-Keyword (CONFIG.PROJECT_HINTS) im Betreff
      boost_body: 12,        // Boost-Keyword im Body
      negative_subject: -35, // Negativ-Keyword im Betreff
      negative_body: -20,    // Negativ-Keyword im Body
      history_per_hit: 8     // pro Treffer aus Domain-Historie (optional)
    },
    MIN_SCORE: 2            // Mindestscore, sonst "Uncategorized"
  },

  // Projekt-Hints (synonyme/Signals + Negatives pro Projekt)
  PROJECT_HINTS: {
    'Aspire': {
      boost: ['aspire', 'sap', 's/4', 's4', 's4hana', 's4/h', 'dms', 'genehmigungsworkflow', 'fit-to-template'],
      negative: []
    },
    'Digitalisierung': {
      boost: ['digitalisierung', 'hyperautomation', 'raadchallenge', 'raad'],
      negative: ['newsletter']
    },
    'Doc SPOC': {
      boost: ['spoc', 'doc spoc', 'documentation', 'dokumentation', 'docs', 'dokumente'],
      negative: ['google ai studio', 'gemini', 'sap', 's/4', 's4hana', 'aspire']
    }
  },

  AI_LANG: {
    SUMMARY_IF_SOURCE: [{ source: 'en', target: 'de' }], // falls Mail englisch -> Summary auf Deutsch
    REPLY_LANGUAGE_MODE: 'sender', // 'sender' | 'fixed' | 'none'
    REPLY_FIXED_LANG: 'de'         // nur wenn MODE='fixed'
  },

  STATUS_OPTIONS: [
    'New',
    'Open',
    'Waiting',
    'Replied',
    'Done (sync)',
    'Done (train only)',
    'Closed',
    'Archived'
  ],
  // Archivierung
  ARCHIVE_OLDER_THAN_DAYS: 14,                  // TODO: anpassen
  STATUS_ARCHIVE_LIST: ['Replied', 'Done (sync)', 'Done (train only)', 'Closed', 'Archived'],// TODO: anpassen

  PRIORITY_OPTIONS: ['Low', 'Medium', 'High', 'Urgent'],
  TASK_FOR_ME_OPTIONS: ['Unsure', 'Yes', 'No'],
  LEARNING: {
    LOW_KEYWORD_MIN_COUNT: 3 // wie oft ein Wort in Korrekturen auftauchen muss, um vorgeschlagen zu werden
  },

  // --- Label → Projekt Mapping (robuster) ---
  IGNORE_LABELS: new Set(['processed-by-apps-script', 'Notes', 'Unwichtig']),
  LABEL_MATCH_MODE: 'regex_first', // 'regex_first' | 'segment_first' | 'exact_only'
  PROJECT_PRIORITY: [
    'BPM Maintenance', 'Reliability Engineering', 'Digitalisierung', 'Aspire',
    'Globale Projekte', 'Projekte lokal', 'Budget', 'Betriebe', 'Leitung',
    'Ausbildung Coaching', 'Private'
  ],
  LABEL_REGEX_MAP: [
    { re: '^0\\s*Aspire', project: 'Aspire' },
    { re: 'Aspire\\s*/\\s*Nexus', project: 'Aspire' },

    { re: '^2\\s*BPM?\\s*Maintenance', project: 'BPM Maintenance' },
    { re: '^2\\s*BPm(?!\\s*Maintenance)', project: 'BPM Maintenance' },

    { re: '^2\\s*Reliability\\s*Engineering', project: 'Reliability Engineering' },
    { re: 'Reliability\\s*Engineering\\s*/\\s*2\\s*Data\\s*Science', project: 'Reliability Engineering' },
    { re: 'OEE\\s*KPI', project: 'Reliability Engineering' },
    { re: 'Master\\s*Data|GRC|DOI', project: 'Reliability Engineering' },

    { re: '^6-?1\\s*Digitalisierung', project: 'Digitalisierung' },
    { re: 'RAAD', project: 'Digitalisierung' },

    { re: '^3\\s*Globale\\s*Projekte', project: 'Globale Projekte' },
    { re: '^4\\s*Projekte\\s*lokal', project: 'Projekte lokal' },
    { re: '^5\\s*Budget', project: 'Budget' },
    { re: '^6\\s*Betriebe', project: 'Betriebe' },
    { re: '^7\\s*Leitung', project: 'Leitung' },
    { re: '^8\\s*Ausbildung\\s*Coaching', project: 'Ausbildung Coaching' },
    { re: '^9\\s*Private', project: 'Private' }
  ],
  LABEL_SEGMENT_MAP: {
    'aspire': 'Aspire', 'nexus': 'Aspire', 'rolemapping': 'Aspire', 'unterbruch': 'Aspire',
    'bpm maintenance': 'BPM Maintenance', 'bpm': 'BPM Maintenance', 'capa': 'BPM Maintenance', 'doc spoc': 'BPM Maintenance',
    'reliability engineering': 'Reliability Engineering', 'data science': 'Reliability Engineering', 'oee': 'Reliability Engineering',
    'master data': 'Reliability Engineering', 'grc': 'Reliability Engineering', 'doi': 'Reliability Engineering',
    'digitalisierung': 'Digitalisierung', 'raadchallenge': 'Digitalisierung', 'raad': 'Digitalisierung',
    'globale projekte': 'Globale Projekte', 'ind. 4.0': 'Globale Projekte',
    'projekte lokal': 'Projekte lokal', 'break through': 'Projekte lokal', 'visual factory': 'Projekte lokal', 'qrb': 'Projekte lokal', 'transformation': 'Projekte lokal',
    'budget': 'Budget', 'betriebe': 'Betriebe', 'leitung': 'Leitung', 'gruppe audit': 'Leitung', 'ziele': 'Leitung',
    'ausbildung coaching': 'Ausbildung Coaching', 'academy x': 'Ausbildung Coaching', 'agil': 'Ausbildung Coaching', 'medium': 'Ausbildung Coaching',
    'prosci': 'Ausbildung Coaching', 'scrum': 'Ausbildung Coaching', 'way we lead': 'Ausbildung Coaching',
    'private': 'Private', 'anlage': 'Private', 'hobby': 'Private', 'quora': 'Private', 'newsfeed': 'Private'
  },
  LABEL_TO_PROJECT: {}

};

// Optional: Script-Property fuer Labelname (ueberschreibt unten, wenn gesetzt)
function getProcessedLabel_() {
  const fromProp = getProp_('PROCESSED_LABEL', '');
  return fromProp !== '' ? fromProp : 'processed-by-apps-script'; // TODO: anpassen/leer lassen
}
