// ===== HELPER FUNCTIONS =====

// Get spreadsheets
function getControlSpreadsheet() {
  if (CONFIG.CONTROL_SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(CONFIG.CONTROL_SPREADSHEET_ID);
    } catch (e) {
      throw new Error('Could not open Control spreadsheet. Check CONFIG.CONTROL_SPREADSHEET_ID');
    }
  } else {
    throw new Error('CONFIG.CONTROL_SPREADSHEET_ID not set. Run setupControlSpreadsheet() first.');
  }
}

function getCallbackDataSpreadsheet() {
  if (CONFIG.CALLBACK_DATA_SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(CONFIG.CALLBACK_DATA_SPREADSHEET_ID);
    } catch (e) {
      throw new Error('Could not open Callback Data spreadsheet. Check CONFIG.CALLBACK_DATA_SPREADSHEET_ID');
    }
  } else {
    throw new Error('CONFIG.CALLBACK_DATA_SPREADSHEET_ID not set. Run setupCallbackDataSpreadsheet() first.');
  }
}

// Get specific sheets
function getArchivesSheet() {
  return getControlSpreadsheet().getSheetByName(SHEETS.CONTROL.ARCHIVES);
}

function getRegistrySheet() {
  return getControlSpreadsheet().getSheetByName(SHEETS.CONTROL.REGISTRY);
}

function getRatingsSheet() {
  return getControlSpreadsheet().getSheetByName(SHEETS.CONTROL.RATINGS);
}

function getConfigSheet() {
  return getControlSpreadsheet().getSheetByName(SHEETS.CONTROL.CONFIG);
}

function getCallbackDataSheet() {
  return getCallbackDataSpreadsheet().getSheetByName('Callback Data');
}

// Config helpers
function getConfig(key) {
  const configSheet = getConfigSheet();
  const data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  
  for (const [setting, value] of data) {
    if (setting === key) {
      return value;
    }
  }
  
  return null;
}

function setConfig(key, value) {
  const configSheet = getConfigSheet();
  const data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      configSheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  
  // If key not found, append it
  const lastRow = configSheet.getLastRow();
  configSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[key, value]]);
}

function updateConfigAfterFetch(newGameCount) {
  const now = new Date();
  
  setConfig('Last Full Fetch Date', now);
  
  const currentTotal = getConfig('Total Games') || 0;
  setConfig('Total Games', currentTotal + newGameCount);
  
  const currentPending = getConfig('Callbacks Pending') || 0;
  setConfig('Callbacks Pending', currentPending + newGameCount);
}

function logToConfig(message, level = 'INFO') {
  const configSheet = getConfigSheet();
  const timestamp = new Date();
  const logMessage = `[${timestamp.toLocaleString()}] ${level}: ${message}`;
  
  const data = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'Recent Errors') {
      const currentLog = data[i][1] || '';
      const logs = currentLog.split('\n').filter(l => l.trim());
      
      logs.unshift(logMessage);
      const newLog = logs.slice(0, 20).join('\n');
      
      configSheet.getRange(i + 2, 2).setValue(newLog);
      return;
    }
  }
}

// Archive helpers
function findArchiveRow(archiveUrl) {
  const archivesSheet = getArchivesSheet();
  const lastRow = archivesSheet.getLastRow();
  
  if (lastRow <= 1) return -1;
  
  const archiveColumn = archivesSheet.getRange(2, ARCHIVE_COLS.ARCHIVE, lastRow - 1, 1).getValues();
  
  for (let i = 0; i < archiveColumn.length; i++) {
    if (archiveColumn[i][0] === archiveUrl) {
      return i + 2;
    }
  }
  
  return -1;
}

function getArchiveETag(archiveUrl) {
  const archivesSheet = getArchivesSheet();
  const rowIndex = findArchiveRow(archiveUrl);
  
  if (rowIndex === -1) return '';
  
  return archivesSheet.getRange(rowIndex, ARCHIVE_COLS.ETAG).getValue() || '';
}

function updateArchiveStatus(archiveUrl, updates) {
  const archivesSheet = getArchivesSheet();
  const rowIndex = findArchiveRow(archiveUrl);
  
  if (rowIndex === -1) {
    Logger.log(`Archive not found in sheet: ${archiveUrl}`);
    return;
  }
  
  if (updates.status !== undefined) {
    archivesSheet.getRange(rowIndex, ARCHIVE_COLS.STATUS).setValue(updates.status);
  }
  
  if (updates.gameCount !== undefined) {
    archivesSheet.getRange(rowIndex, ARCHIVE_COLS.GAME_COUNT).setValue(updates.gameCount);
  }
  
  if (updates.lastChecked !== undefined) {
    archivesSheet.getRange(rowIndex, ARCHIVE_COLS.LAST_CHECKED).setValue(updates.lastChecked);
  }
  
  if (updates.lastFetched !== undefined) {
    archivesSheet.getRange(rowIndex, ARCHIVE_COLS.LAST_FETCHED).setValue(updates.lastFetched);
  }
  
  if (updates.etag !== undefined) {
    archivesSheet.getRange(rowIndex, ARCHIVE_COLS.ETAG).setValue(updates.etag);
  }
  
  if (updates.notes !== undefined) {
    archivesSheet.getRange(rowIndex, ARCHIVE_COLS.NOTES).setValue(updates.notes);
  }
}

function updateArchiveStatuses(archives, newGameCount) {
  const now = new Date();
  
  for (const archiveUrl of archives) {
    const parts = archiveUrl.split('/');
    const year = parseInt(parts[parts.length - 2]);
    const month = parseInt(parts[parts.length - 1]);
    
    const archiveEndDate = new Date(year, month, 0);
    const isComplete = now > archiveEndDate;
    
    const updates = {
      lastChecked: now,
      lastFetched: now
    };
    
    if (isComplete) {
      updates.status = 'complete';
    }
    
    updateArchiveStatus(archiveUrl, updates);
  }
}

// Game index cache
function getRecentGameIndex(gamesSheet) {
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) return {};
  
  const cacheSize = CONFIG.GAME_INDEX_CACHE_SIZE || 500;
  const startRow = Math.max(2, lastRow - cacheSize + 1);
  const numRows = lastRow - startRow + 1;
  
  const gameIds = gamesSheet.getRange(startRow, GAMES_COLS.GAME_ID, numRows, 1).getValues();
  
  const index = {};
  for (let i = 0; i < gameIds.length; i++) {
    index[gameIds[i][0]] = startRow + i;
  }
  
  return index;
}

// Find game rows
// ===== FIND GAME ROW IN RATINGS (OPTIMIZED) =====
function findGameRowInRatings(gameId) {
  const ratingsSheet = getRatingsSheet();
  const lastRow = ratingsSheet.getLastRow();
  
  if (lastRow <= 1) return -1;
  
  // Read all game IDs at once
  const gameIds = ratingsSheet.getRange(2, RATINGS_COLS.GAME_ID, lastRow - 1, 1).getValues();
  
  // Search from end to beginning (recent games more likely)
  for (let i = gameIds.length - 1; i >= 0; i--) {
    if (String(gameIds[i][0]) === String(gameId)) {
      return i + 2; // +2 because array is 0-indexed and sheet starts at row 2
    }
  }
  
  return -1;
}

// ===== FIND GAME ROW IN REGISTRY (OPTIMIZED) =====
function findGameRowInRegistry(gameId) {
  const registrySheet = getRegistrySheet();
  const lastRow = registrySheet.getLastRow();
  
  if (lastRow <= 1) return -1;
  
  // Read all game IDs at once
  const gameIds = registrySheet.getRange(2, REGISTRY_COLS.GAME_ID, lastRow - 1, 1).getValues();
  
  // Search from end to beginning (recent games more likely)
  for (let i = gameIds.length - 1; i >= 0; i--) {
    if (String(gameIds[i][0]) === String(gameId)) {
      return i + 2; // +2 because array is 0-indexed and sheet starts at row 2
    }
  }
  
  return -1;
}
