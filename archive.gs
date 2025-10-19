// ===== SMART ARCHIVE MANAGEMENT =====
// Auto-detects archives based on account join date and maintains list

function setupArchivesSheet() {
  const ss = getControlSpreadsheet();
  const sheet = ss.getSheetByName('Archives') || ss.insertSheet('Archives');
  
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Archive URL', 'Year', 'Month', 'Status', 'Game Count',
      'ETag', 'Last Game ID', 'Last Checked', 'Last Fetched'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    sheet.setFrozenRows(1);
    sheet.getRange('H:I').setNumberFormat('M/D/YY h:mm');
  }
  
  return sheet;
}

// ===== GET ACCOUNT JOIN DATE =====
function getAccountJoinDate() {
  try {
    const url = `https://api.chess.com/pub/player/${CONFIG.USERNAME}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    
    if (data.joined) {
      const joinDate = new Date(data.joined * 1000);
      Logger.log(`Account joined: ${joinDate}`);
      return joinDate;
    }
  } catch (error) {
    Logger.log(`Could not fetch join date: ${error.message}`);
  }
  
  return null;
}

// ===== BUILD COMPLETE ARCHIVES LIST =====
function buildCompleteArchivesList() {
  // Get account join date
  const joinDate = getAccountJoinDate();
  
  if (!joinDate) {
    // Fallback to API list if we can't get join date
    return getAllArchivesFromAPI();
  }
  
  const archives = [];
  const now = new Date();
  
  // Start from join month
  let currentDate = new Date(joinDate.getFullYear(), joinDate.getMonth(), 1);
  
  // Generate all months from join date to current month
  while (currentDate <= now) {
    const year = currentDate.getFullYear();
    const month = String(currentDate.getMonth() + 1).padStart(2, '0');
    const archiveUrl = `https://api.chess.com/pub/player/${CONFIG.USERNAME}/games/${year}/${month}`;
    
    archives.push(archiveUrl);
    
    // Move to next month
    currentDate.setMonth(currentDate.getMonth() + 1);
  }
  
  Logger.log(`Built ${archives.length} archives from join date`);
  return archives;
}

// ===== GET ALL ARCHIVES FROM API (FALLBACK) =====
function getAllArchivesFromAPI() {
  const url = `https://api.chess.com/pub/player/${CONFIG.USERNAME}/games/archives`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  return data.archives || [];
}

// ===== AUTO-SYNC ARCHIVES (CALLED BEFORE EACH FETCH) =====
function syncArchivesList() {
  const archivesSheet = setupArchivesSheet();
  
  // Build complete list based on join date or API
  const completeArchives = buildCompleteArchivesList();
  
  if (!completeArchives.length) {
    throw new Error('No archives found for user: ' + CONFIG.USERNAME);
  }
  
  // Get existing archives from sheet
  const lastRow = archivesSheet.getLastRow();
  const existingArchives = new Set();
  
  if (lastRow > 1) {
    const existingData = archivesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingData.forEach(row => existingArchives.add(row[0]));
  }
  
  // Find new archives
  const newArchives = completeArchives.filter(archive => !existingArchives.has(archive));
  
  if (newArchives.length > 0) {
    const now = new Date();
    const newRows = newArchives.map(archiveUrl => {
      const parts = archiveUrl.split('/');
      const year = parts[parts.length - 2];
      const month = parts[parts.length - 1];
      
      return [
        archiveUrl,
        year,
        month,
        'pending',
        0,
        '',
        '',  // Last Game ID
        now,
        null
      ];
    });
    
    const startRow = archivesSheet.getLastRow() + 1;
    archivesSheet.getRange(startRow, 1, newRows.length, 9).setValues(newRows);
    
    Logger.log(`Added ${newArchives.length} new archives`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Found ${newArchives.length} new archive(s)!`, 
      '✨', 
      3
    );
  }
  
  setProperty(PROP_KEYS.LAST_ARCHIVE_CHECK, new Date().toISOString());
  
  return completeArchives.length;
}

// ===== GET ARCHIVES TO FETCH =====
function getArchivesToFetch(options = {}) {
  const { monthsToFetch = CONFIG.MONTHS_TO_FETCH, specificArchive = null } = options;
  
  // Always sync archives list first
  syncArchivesList();
  
  if (specificArchive) {
    return [specificArchive];
  }
  
  const archivesSheet = getArchivesSheet();
  const lastRow = archivesSheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  const data = archivesSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  if (monthsToFetch === 0) {
    // Fetch all pending archives
    return data
      .filter(row => row[3] !== 'complete')
      .map(row => row[0]);
  } else {
    // Fetch recent N months (pending or complete to check for updates)
    const now = new Date();
    const cutoffDate = new Date(now.getFullYear(), now.getMonth() - monthsToFetch + 1, 1);
    
    return data
      .filter(row => {
        const year = parseInt(row[1]);
        const month = parseInt(row[2]);
        const archiveDate = new Date(year, month - 1, 1);
        return archiveDate >= cutoffDate;
      })
      .map(row => row[0]);
  }
}

// ===== UPDATE ARCHIVE STATUS =====
function updateArchiveStatus(archiveUrl, updates) {
  const archivesSheet = getArchivesSheet();
  const lastRow = archivesSheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const urls = archivesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  for (let i = 0; i < urls.length; i++) {
    if (urls[i][0] === archiveUrl) {
      const rowNum = i + 2;
      
      if (updates.status !== undefined) {
        archivesSheet.getRange(rowNum, 4).setValue(updates.status);
      }
      if (updates.gameCount !== undefined) {
        archivesSheet.getRange(rowNum, 5).setValue(updates.gameCount);
      }
      if (updates.etag !== undefined) {
        archivesSheet.getRange(rowNum, 6).setValue(updates.etag);
      }
      if (updates.lastGameId !== undefined) {
        archivesSheet.getRange(rowNum, 7).setValue(updates.lastGameId);
      }
      if (updates.lastChecked !== undefined) {
        archivesSheet.getRange(rowNum, 8).setValue(updates.lastChecked);
      }
      if (updates.lastFetched !== undefined) {
        archivesSheet.getRange(rowNum, 9).setValue(updates.lastFetched);
      }
      
      return;
    }
  }
}

// ===== GET LAST GAME ID FOR ARCHIVE =====
function getLastGameIdForArchive(archiveUrl) {
  const archivesSheet = getArchivesSheet();
  const lastRow = archivesSheet.getLastRow();
  
  if (lastRow <= 1) return null;
  
  const data = archivesSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  for (const row of data) {
    if (row[0] === archiveUrl) {
      return row[6] || null;  // Last Game ID column
    }
  }
  
  return null;
}

// ===== QUICK FETCH NEW GAMES =====
function fetchNewGames() {
  // Smart fetch: only checks recent 2 months by default
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Fetch New Games',
    `This will check the last ${CONFIG.MONTHS_TO_FETCH} months for new games.\n\n` +
    'Archives are auto-synced before each fetch.',
    ui.ButtonSet.OK
  );
  
  fetchChesscomGames();
}

// ===== FETCH ALL HISTORY =====
function fetchAllHistory() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Fetch All History?',
    'This will fetch ALL games from your entire Chess.com history.\n\n' +
    'This may take several minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    fetchChesscomGames({ monthsToFetch: 0 });
  }
}
