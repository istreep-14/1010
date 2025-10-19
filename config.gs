// ===== DYNAMIC CONFIGURATION SYSTEM =====
// Uses Properties Service - no hardcoded spreadsheet IDs

const CONFIG = {
  USERNAME: 'frankscobey',
  MONTHS_TO_FETCH: 2, // 0 = all history
  CALLBACK_BATCH_SIZE: 11,
  DUPLICATE_CHECK_ROWS: 200
};

// ===== PROPERTY KEYS =====
const PROP_KEYS = {
  CONTROL_SHEET_ID: 'CONTROL_SPREADSHEET_ID',
  ENRICHMENT_SHEET_ID: 'ENRICHMENT_SPREADSHEET_ID',
  LAST_ARCHIVE_CHECK: 'LAST_ARCHIVE_CHECK',
  TOTAL_GAMES: 'TOTAL_GAMES'
};

// ===== GET/SET PROPERTIES =====
function getProperty(key, defaultValue = null) {
  const props = PropertiesService.getScriptProperties();
  const value = props.getProperty(key);
  return value || defaultValue;
}

function setProperty(key, value) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(key, String(value));
}

// ===== GET SPREADSHEETS (WITH AUTO-CREATION) =====
function getControlSpreadsheet() {
  let id = getProperty(PROP_KEYS.CONTROL_SHEET_ID);
  
  if (!id) {
    // Auto-create if doesn't exist
    const ss = SpreadsheetApp.create('Chess Control - ' + CONFIG.USERNAME);
    id = ss.getId();
    setProperty(PROP_KEYS.CONTROL_SHEET_ID, id);
    Logger.log('Created Control Spreadsheet: ' + id);
  }
  
  return SpreadsheetApp.openById(id);
}

function getEnrichmentSpreadsheet() {
  let id = getProperty(PROP_KEYS.ENRICHMENT_SHEET_ID);
  
  if (!id) {
    const ss = SpreadsheetApp.create('Chess Enrichment - ' + CONFIG.USERNAME);
    id = ss.getId();
    setProperty(PROP_KEYS.ENRICHMENT_SHEET_ID, id);
    Logger.log('Created Enrichment Spreadsheet: ' + id);
  }
  
  return SpreadsheetApp.openById(id);
}

function getMainSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ===== SHEET GETTERS =====
function getGamesSheet() {
  return getMainSpreadsheet().getSheetByName('Games');
}

function getArchivesSheet() {
  return getControlSpreadsheet().getSheetByName('Archives');
}

function getRegistrySheet() {
  return getControlSpreadsheet().getSheetByName('Registry');
}

function getCallbackSheet() {
  return getEnrichmentSpreadsheet().getSheetByName('Callback Data');
}

function getLichessSheet() {
  return getEnrichmentSpreadsheet().getSheetByName('Lichess Analysis');
}

function getSummarySheet() {
  return getMainSpreadsheet().getSheetByName('Summary');
}

// ===== RESET ALL (FOR CLEAN START) =====
function resetAllConnections() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset All Connections?',
    'This will clear all stored spreadsheet IDs. The system will create new Control and Enrichment sheets on next run.\n\nYour Games sheet will NOT be deleted.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteAllProperties();
    ui.alert('✅ Reset complete! Run any function to auto-create new sheets.');
  }
}

// ===== SHOW CURRENT CONFIGURATION =====
function showConfiguration() {
  const controlId = getProperty(PROP_KEYS.CONTROL_SHEET_ID, 'Not created yet');
  const enrichmentId = getProperty(PROP_KEYS.ENRICHMENT_SHEET_ID, 'Not created yet');
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .info { margin: 10px 0; font-size: 14px; }
      .id { font-family: monospace; background: #f5f5f5; padding: 5px; }
    </style>
    <h2>📊 Chess Tracker Configuration</h2>
    <div class="info"><strong>Username:</strong> ${CONFIG.USERNAME}</div>
    <div class="info"><strong>Main Spreadsheet:</strong> ${getMainSpreadsheet().getId()}</div>
    <div class="info"><strong>Control Spreadsheet:</strong><br><span class="id">${controlId}</span></div>
    <div class="info"><strong>Enrichment Spreadsheet:</strong><br><span class="id">${enrichmentId}</span></div>
    <hr>
    <p>All spreadsheet IDs are stored automatically using Properties Service.</p>
    <p>Use "Reset All Connections" from the menu to start fresh.</p>
  `)
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Configuration');
}
