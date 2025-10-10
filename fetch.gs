// ===== FETCH FUNCTIONS =====

function fetchAllGames() {
  fetchChesscomGames({ fetchAll: true });
}

function fetchSpecificMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Fetch Specific Month',
    'Enter archive in format YYYY-MM (e.g., 2025-10):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const archiveMonth = response.getResponseText().trim();
    
    if (!/^\d{4}-\d{2}$/.test(archiveMonth)) {
      ui.alert('Invalid format. Please use YYYY-MM (e.g., 2025-10)');
      return;
    }
    
    const [year, month] = archiveMonth.split('-');
    const archiveUrl = `https://api.chess.com/pub/player/${CONFIG.USERNAME}/games/${year}/${month}`;
    
    fetchChesscomGames({ specificArchives: [archiveUrl] });
  }
}

function fetchChesscomGames(options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!gamesSheet) {
    SpreadsheetApp.getUi().alert('❌ Run "Setup Games Sheet" first!');
    return;
  }
  
  if (!CONFIG.CONTROL_SPREADSHEET_ID) {
    SpreadsheetApp.getUi().alert('❌ Run "Setup Control Spreadsheet" first and add ID to CONFIG!');
    return;
  }
  
  try {
    const archives = getArchivesToFetch(options);
    
    if (!archives.length) {
      ss.toast('No archives to fetch', 'ℹ️', 3);
      return;
    }
    
    ss.toast(`Fetching ${archives.length} archive(s)...`, '⏳', -1);
    
    const allGames = fetchGamesFromArchives(archives);
    
    if (!allGames.length) {
      ss.toast('No games found in archives', 'ℹ️', 3);
      updateArchiveStatuses(archives, 0);
      return;
    }
    
    const newGames = filterNewGames(allGames, gamesSheet);
    
    if (!newGames.length) {
      ss.toast('No new games found', 'ℹ️', 3);
      updateArchiveStatuses(archives, 0);
      return;
    }
    
    const ledger = getLastLedger(gamesSheet);
    Logger.log('Starting ledger loaded: ' + JSON.stringify(ledger));
    
    ss.toast(`Processing ${newGames.length} new games...`, '⏳', -1);
    
    const gamesRows = processGames(newGames, CONFIG.USERNAME, ledger);
    writeGamesToSheet(gamesSheet, gamesRows);
    
    writeToControlSheets(newGames, ledger);
    
    updateArchiveStatuses(archives, newGames.length);
    
    updateConfigAfterFetch(newGames.length);
    
    ss.toast(`✅ Added ${newGames.length} new games!`, '✅', 5);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== GET ARCHIVES TO FETCH =====
function getArchivesToFetch(options) {
  const {
    fetchAll = false,
    forceComplete = false,
    specificArchives = [],
    updateExisting = false
  } = options;
  
  if (specificArchives.length > 0) {
    return specificArchives;
  }
  
  if (fetchAll || CONFIG.MONTHS_TO_FETCH === 0) {
    return getAllArchives(CONFIG.USERNAME);
  }
  
  const archivesSheet = getArchivesSheet();
  const lastRow = archivesSheet.getLastRow();
  
  if (lastRow <= 1) {
    return getRecentArchives(CONFIG.USERNAME, CONFIG.MONTHS_TO_FETCH);
  }
  
  const data = archivesSheet.getRange(2, 1, lastRow - 1, ARCHIVE_COLS.NOTES).getValues();
  
  const archives = [];
  
  for (const row of data) {
    const archiveUrl = row[ARCHIVE_COLS.ARCHIVE - 1];
    const status = row[ARCHIVE_COLS.STATUS - 1];
    
    if (status !== 'complete' || forceComplete) {
      archives.push(archiveUrl);
    }
  }
  
  return archives;
}

// ===== ARCHIVE FETCHING =====
function getAllArchives(username) {
  const url = `https://api.chess.com/pub/player/${username}/games/archives`;
  const response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText()).archives;
}

function getRecentArchives(username, months) {
  const archives = [];
  const now = new Date();
  
  for (let i = 0; i < months; i++) {
    const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    archives.push(`https://api.chess.com/pub/player/${username}/games/${year}/${month}`);
  }
  
  return archives;
}

// ===== GAME FETCHING =====
function fetchGamesFromArchives(archiveUrls) {
  const allGames = [];
  const archivesSheet = getArchivesSheet();
  
  for (const url of archiveUrls) {
    try {
      updateArchiveStatus(url, { status: 'fetching' });
      
      const storedETag = getArchiveETag(url);
      
      const fetchOptions = {
        muteHttpExceptions: true
      };
      
      if (storedETag) {
        fetchOptions.headers = {
          'If-None-Match': storedETag
        };
      }
      
      const response = UrlFetchApp.fetch(url, fetchOptions);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 304) {
        Logger.log(`Archive not modified: ${url}`);
        updateArchiveStatus(url, { 
          status: 'complete',
          lastChecked: new Date(),
          notes: 'No changes (304)'
        });
        Utilities.sleep(300);
        continue;
      }
      
      const newETag = response.getHeaders()['ETag'] || response.getHeaders()['etag'] || '';
      
      const data = JSON.parse(response.getContentText());
      if (data.games) {
        allGames.push(...data.games);
        
        updateArchiveStatus(url, {
          etag: newETag,
          lastChecked: new Date(),
          lastFetched: new Date()
        });
      }
      
      Utilities.sleep(300);
      
    } catch (e) {
      Logger.log(`Failed to fetch ${url}: ${e.message}`);
      updateArchiveStatus(url, { 
        status: 'error',
        notes: e.message 
      });
    }
  }
  
  return allGames.sort((a, b) => a.end_time - b.end_time);
}

// ===== NEW GAME DETECTION =====
function filterNewGames(games, sheet) {
  const existingGames = new Set();
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const checkRows = Math.min(CONFIG.DUPLICATE_CHECK_ROWS, lastRow - 1);
    const startRow = lastRow - checkRows + 1;
    
    Logger.log(`Checking last ${checkRows} rows for duplicates (rows ${startRow} to ${lastRow})`);
    
    const gameIds = sheet.getRange(startRow, GAMES_COLS.GAME_ID, checkRows, 2).getValues();
    
    for (const [gameId, gameType] of gameIds) {
      existingGames.add(`${gameId}_${gameType}`);
    }
  }
  
  return games.filter(game => {
    const gameId = game.url.split('/').pop();
    const gameType = (game.time_class || '').toLowerCase() === 'daily' ? 'daily' : 'live';
    return !existingGames.has(`${gameId}_${gameType}`);
  });
}

// ===== LEDGER =====
function getLastLedger(sheet) {
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return {};
  
  try {
    const lastLedgerCell = sheet.getRange(lastRow, GAMES_COLS.RATINGS_LEDGER).getValue();
    
    if (!lastLedgerCell || lastLedgerCell === '') {
      Logger.log('No ledger found in last row, returning empty ledger');
      return {};
    }
    
    const ledger = JSON.parse(lastLedgerCell);
    Logger.log('Loaded ledger from last row: ' + JSON.stringify(ledger));
    return ledger;
    
  } catch (e) {
    Logger.log('Could not parse ledger: ' + e.message);
    return {};
  }
}

// ===== WRITING =====
function writeGamesToSheet(sheet, rows) {
  if (!rows.length) return;
  
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

// ===== WRITE TO CONTROL SHEETS =====
function writeToControlSheets(games, startingLedger) {
  const now = new Date();
  
  const registryRows = [];
  const ratingsRows = [];
  
  let currentLedger = JSON.parse(JSON.stringify(startingLedger));
  
  for (const game of games) {
    if (!game || !game.url || !game.end_time) continue;
    
    const gameId = game.url.split('/').pop();
    const gameType = (game.time_class || '').toLowerCase() === 'daily' ? 'daily' : 'live';
    const format = getGameFormat(game).toLowerCase();
    const endDate = new Date(game.end_time * 1000);
    const archive = `${endDate.getFullYear()}-${String(endDate.getMonth() + 1).padStart(2, '0')}`;
    
    const isWhite = game.white?.username.toLowerCase() === CONFIG.USERNAME.toLowerCase();
    const myRating = isWhite ? game.white?.rating : game.black?.rating;
    const oppRating = isWhite ? game.black?.rating : game.white?.rating;
    
    const myRatingLast = currentLedger[format] || null;
    const myRatingDelta = (myRatingLast !== null && myRating !== null) ? (myRating - myRatingLast) : null;
    
    if (myRating !== null) {
      currentLedger[format] = myRating;
    }
    
    const oppRatingDelta = myRatingDelta !== null ? myRatingDelta * -1 : null;
    const oppRatingLast = (oppRatingDelta !== null && oppRating !== null) ? oppRating - oppRatingDelta : null;
    
    registryRows.push([
      gameId,
      archive,
      endDate,
      format,
      now,
      null,
      null,
      null,
      null,
      'Games'
    ]);
    
    ratingsRows.push([
      gameId,
      game.url,
      archive,
      endDate,
      format,
      myRating,
      myRatingLast,
      myRatingDelta,
      oppRating,
      oppRatingDelta,
      oppRatingLast,
      null,
      null,
      null,
      null,
      null,
      null,
      myRatingLast,
      myRatingDelta,
      oppRatingLast,
      oppRatingDelta
    ]);
  }
  
  if (registryRows.length > 0) {
    const registrySheet = getRegistrySheet();
    const startRow = registrySheet.getLastRow() + 1;
    registrySheet.getRange(startRow, 1, registryRows.length, registryRows[0].length).setValues(registryRows);
    Logger.log(`Wrote ${registryRows.length} games to Registry`);
  }
  
  if (ratingsRows.length > 0) {
    const ratingsSheet = getRatingsSheet();
    const startRow = ratingsSheet.getLastRow() + 1;
    ratingsSheet.getRange(startRow, 1, ratingsRows.length, ratingsRows[0].length).setValues(ratingsRows);
    Logger.log(`Wrote ${ratingsRows.length} games to Ratings`);
  }
}

// ===== POPULATE ARCHIVES =====
function populateArchives() {
  try {
    const archivesSheet = getArchivesSheet();
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Fetching archives list from Chess.com...', '⏳', -1);
    
    const url = `https://api.chess.com/pub/player/${CONFIG.USERNAME}/games/archives`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    const archives = data.archives;
    
    if (!archives || !archives.length) {
      SpreadsheetApp.getUi().alert('No archives found for user: ' + CONFIG.USERNAME);
      return;
    }
    
    const rows = [];
    const now = new Date();
    
    for (const archiveUrl of archives) {
      const parts = archiveUrl.split('/');
      const year = parts[parts.length - 2];
      const month = parts[parts.length - 1];
      
      rows.push([
        archiveUrl,
        year,
        month,
        'pending',
        0,
        now,
        null,
        null,
        '',
        'Games',
        ''
      ]);
    }
    
    const startRow = archivesSheet.getLastRow() + 1;
    archivesSheet.getRange(startRow, 1, rows.length, ARCHIVE_COLS.NOTES).setValues(rows);
    
    setConfig('Total Archives', archives.length);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Added ${archives.length} archives!`, 
      '✅', 
      5
    );
    
    Logger.log(`Populated ${archives.length} archives`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error populating archives: ' + error.message);
    Logger.log(error);
  }
}
