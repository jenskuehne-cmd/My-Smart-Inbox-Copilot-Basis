/*** Menu.gs ***/
// Baut ein Menue "Mail Assistant" im Sheet.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Assistant')
    .addItem('Verarbeite: Heute', 'menu_processToday')
    .addItem('Verarbeite: Letzte 7 Tage', 'menu_processLast7')
    .addItem('Verarbeite: Datumsbereich…', 'menu_processDateRange')
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
