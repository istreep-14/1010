// ===== MAINTENANCE FUNCTIONS =====

// ===== VALIDATE COMPLETE ARCHIVES =====
function validateCompleteArchives() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Checking complete archives for changes...', '⏳', -1);
    
    const archivesSheet = getArchivesSheet();
    const lastRow = archivesSheet.getLastRow();
    
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('No archives found. Run "Populate Archives" first.');
      return;
    }
    
    const data = archivesSheet.getRange(2, 1, lastRow - 1, ARCHIVE_COLS.NOTES).getValues();
    const completeArchives = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const status = row[ARCHIVE_COLS.STATUS - 1];
      
      if (status === 'complete') {
        completeArchives.push({
          url: row[ARCHIVE_COLS.ARCHIVE - 1],
          etag: row[ARCHIVE_COLS.ETAG - 1],
          rowIndex: i + 2
        });
      }
    }
    
    if (!completeArchives.length) {
      ss.toast('No complete archives to validate', 'ℹ️', 3);
      return;
    }
    
    ss.toast(`Checking ${completeArchives.length} complete archives...`, '⏳', -1);
    
    let changedCount = 0;
    const changedArchives = [];
    
    for (const archive of completeArchives) {
      try {
        const response = UrlFetchApp.fetch(archive.url, {
          method: 'head',
          muteHttpExceptions: true
        });
        
        if (response.getResponseCode() === 200) {
          const currentETag = response.getHeaders()['ETag'] || response.getHeaders()['etag'] || '';
          
          if (currentETag && archive.etag && currentETag !== archive.etag) {
            changedCount++;
            changedArchives.push(archive.url);
            
            updateArchiveStatus(archive.url, {
              status: 'pending',
              lastChecked: new Date(),
              notes: `ETag changed (was: ${archive.etag.substring(0, 20)}...)`
            });
            
            Logger.log(`Archive changed: ${archive.url}`);
          } else {
            updateArchiveStatus(archive.url, {
              lastChecked: new Date()
            });
          }
        }
        
        Utilities.sleep(300);
        
      } catch (e) {
        Logger.log(`Error checking archive ${archive.url}: ${e.message}`);
      }
    }
    
    setConfig('Last Validation Run Date', new Date());
    
    if (changedCount > 0) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        `Found ${changedCount} changed archive(s)`,
        `The following archives have been updated by Chess.com:\n\n${changedArchives.join('\n')}\n\nWould you like to re-fetch them now?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        fetchChesscomGames({ specificArchives: changedArchives });
      } else {
        ss.toast(`${changedCount} archives marked as pending`, '✅', 5);
      }
    } else {
      ss.toast('✅ All complete archives are up to date!', '✅', 5);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== UPDATE CHANGED ARCHIVES =====
function updateChangedArchives() {
  const archivesSheet = getArchivesSheet();
  const lastRow = archivesSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No archives found.');
    return;
  }
  
  const data = archivesSheet.getRange(2, 1, lastRow - 1, ARCHIVE_COLS.STATUS).getValues();
  const pendingArchives = [];
  
  for (const row of data) {
    const url = row[ARCHIVE_COLS.ARCHIVE - 1];
    const status = row[ARCHIVE_COLS.STATUS - 1];
    
    if (status === 'pending' || status === 'error') {
      pendingArchives.push(url);
    }
  }
  
  if (!pendingArchives.length) {
    SpreadsheetApp.getUi().alert('No pending archives to update.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    `Update ${pendingArchives.length} pending archive(s)?`,
    'This will re-fetch games from archives marked as pending or error.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    fetchChesscomGames({ specificArchives: pendingArchives, updateExisting: true });
  }
}

// ===== REBUILD RATINGS FROM SCRATCH =====
function rebuildRatingsFromScratch() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Rebuild Ratings Sheet?',
    'This will delete all data in the Ratings sheet and rebuild it from the Games sheet.\n\nThis is useful if rating calculations got out of sync.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Rebuilding Ratings sheet...', '⏳', -1);
    
    const ratingsSheet = getRatingsSheet();
    const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
    
    if (!ratingsSheet || !gamesSheet) {
      throw new Error('Required sheets not found');
    }
    
    const lastRow = ratingsSheet.getLastRow();
    if (lastRow > 1) {
      ratingsSheet.getRange(2, 1, lastRow - 1, RATINGS_COLS.OPP_DELTA_EFFECTIVE).clearContent();
    }
    
    const gamesLastRow = gamesSheet.getLastRow();
    if (gamesLastRow <= 1) {
      ss.toast('No games to rebuild from', 'ℹ️', 3);
      return;
    }
    
    const gamesData = gamesSheet.getRange(2, 1, gamesLastRow - 1, GAMES_COLS.RATINGS_LEDGER).getValues();
    
    const ratingsRows = [];
    let currentLedger = {};
    
    for (const gameRow of gamesData) {
      const gameId = gameRow[GAMES_COLS.GAME_ID - 1];
      const gameUrl = gameRow[GAMES_COLS.GAME_URL - 1];
      const archive = gameRow[GAMES_COLS.ARCHIVE - 1];
      const endDate = gameRow[GAMES_COLS.END_DATE - 1];
      const format = gameRow[GAMES_COLS.FORMAT - 1];
      const myRating = gameRow[GAMES_COLS.MY_RATING - 1];
      const oppRating = gameRow[GAMES_COLS.OPP_RATING - 1];
      
      if (!gameId || !format) continue;
      
      const myRatingLast = currentLedger[format] || null;
      const myRatingDelta = (myRatingLast !== null && myRating !== null) ? (myRating - myRatingLast) : null;
      
      if (myRating !== null) {
        currentLedger[format] = myRating;
      }
      
      const oppRatingDelta = myRatingDelta !== null ? myRatingDelta * -1 : null;
      const oppRatingLast = (oppRatingDelta !== null && oppRating !== null) ? oppRating - oppRatingDelta : null;
      
      ratingsRows.push([
        gameId,
        gameUrl,
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
    
    if (ratingsRows.length > 0) {
      ratingsSheet.getRange(2, 1, ratingsRows.length, ratingsRows[0].length).setValues(ratingsRows);
    }
    
    ss.toast(`✅ Rebuilt ${ratingsRows.length} ratings!`, '✅', 5);
    
    ui.alert(
      'Ratings Rebuilt Successfully!',
      `Rebuilt ${ratingsRows.length} game ratings.\n\nNote: Callback data was cleared. Run callback enrichment again if needed.`
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== FIX MISSING GAMES IN CONTROL =====
function fixMissingGamesInControl() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Sync Games to Control Sheets?',
    'This will check for games in the Games sheet that are missing from Control sheets and add them.\n\nUseful if Control sheets got out of sync.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Checking for missing games...', '⏳', -1);
    
    const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
    const registrySheet = getRegistrySheet();
    
    if (!gamesSheet || !registrySheet) {
      throw new Error('Required sheets not found');
    }
    
    const gamesLastRow = gamesSheet.getLastRow();
    if (gamesLastRow <= 1) {
      ss.toast('No games found', 'ℹ️', 3);
      return;
    }
    
    const gamesData = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, gamesLastRow - 1, 1).getValues();
    const gamesIds = new Set(gamesData.map(row => row[0]));
    
    const registryLastRow = registrySheet.getLastRow();
    const registryIds = new Set();
    
    if (registryLastRow > 1) {
      const registryData = registrySheet.getRange(2, REGISTRY_COLS.GAME_ID, registryLastRow - 1, 1).getValues();
      registryData.forEach(row => registryIds.add(row[0]));
    }
    
    const missingIds = [...gamesIds].filter(id => !registryIds.has(id));
    
    if (!missingIds.length) {
      ss.toast('✅ All games are in sync!', '✅', 3);
      return;
    }
    
    ss.toast(`Adding ${missingIds.length} missing games...`, '⏳', -1);
    
    const missingGames = [];
    
    for (const gameId of missingIds) {
      for (let i = 0; i < gamesData.length; i++) {
        if (gamesData[i][0] === gameId) {
          const rowNum = i + 2;
          const gameRow = gamesSheet.getRange(rowNum, 1, 1, GAMES_COLS.RATINGS_LEDGER).getValues()[0];
          
          missingGames.push({
            url: gameRow[GAMES_COLS.GAME_URL - 1],
            end_time: gameRow[GAMES_COLS.END_EPOCH - 1],
            time_class: gameRow[GAMES_COLS.TIME_CLASS - 1],
            white: { 
              username: gameRow[GAMES_COLS.COLOR - 1] === 'white' ? CONFIG.USERNAME : gameRow[GAMES_COLS.OPPONENT - 1],
              rating: gameRow[GAMES_COLS.COLOR - 1] === 'white' ? gameRow[GAMES_COLS.MY_RATING - 1] : gameRow[GAMES_COLS.OPP_RATING - 1]
            },
            black: {
              username: gameRow[GAMES_COLS.COLOR - 1] === 'black' ? CONFIG.USERNAME : gameRow[GAMES_COLS.OPPONENT - 1],
              rating: gameRow[GAMES_COLS.COLOR - 1] === 'black' ? gameRow[GAMES_COLS.MY_RATING - 1] : gameRow[GAMES_COLS.OPP_RATING - 1]
            },
            rated: gameRow[GAMES_COLS.RATED - 1]
          });
          break;
        }
      }
    }
    
    const ratingsSheet = getRatingsSheet();
    let startingLedger = {};
    
    if (ratingsSheet.getLastRow() > 1) {
      const lastRatings = ratingsSheet.getRange(2, 1, ratingsSheet.getLastRow() - 1, RATINGS_COLS.MY_RATING).getValues();
      
      for (let i = lastRatings.length - 1; i >= 0; i--) {
        const format = lastRatings[i][RATINGS_COLS.FORMAT - 1];
        const rating = lastRatings[i][RATINGS_COLS.MY_RATING - 1];
        
        if (format && rating && !startingLedger[format]) {
          startingLedger[format] = rating;
        }
      }
    }
    
    writeToControlSheets(missingGames, startingLedger);
    
    ss.toast(`✅ Added ${missingIds.length} missing games!`, '✅', 5);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== RESET CALLBACK ERRORS =====
function resetCallbackErrors() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset Callback Errors?',
    'This will change all callback statuses from "error" back to null so they can be retried.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const registrySheet = getRegistrySheet();
    const lastRow = registrySheet.getLastRow();
    
    if (lastRow <= 1) {
      ss.toast('No games in registry', 'ℹ️', 3);
      return;
    }
    
    const statuses = registrySheet.getRange(2, REGISTRY_COLS.CALLBACK_STATUS, lastRow - 1, 1).getValues();
    
    let resetCount = 0;
    
    for (let i = 0; i < statuses.length; i++) {
      if (statuses[i][0] === 'error') {
        registrySheet.getRange(i + 2, REGISTRY_COLS.CALLBACK_STATUS).setValue(null);
        registrySheet.getRange(i + 2, REGISTRY_COLS.CALLBACK_DATE).setValue(null);
        resetCount++;
      }
    }
    
    setConfig('Last Registry Row Processed', 1);
    
    updateEnrichmentProgress();
    
    ss.toast(`✅ Reset ${resetCount} error statuses`, '✅', 5);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== CHECK FOR DUPLICATES =====
function checkForDuplicateGames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Checking for duplicate games...', '⏳', -1);
  
  try {
    const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
    const lastRow = gamesSheet.getLastRow();
    
    if (lastRow <= 1) {
      ss.toast('No games found', 'ℹ️', 3);
      return;
    }
    
    const data = gamesSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    const seen = new Set();
    const duplicates = [];
    
    for (let i = 0; i < data.length; i++) {
      const gameId = data[i][0];
      const gameType = data[i][1];
      const key = `${gameId}_${gameType}`;
      
      if (seen.has(key)) {
        duplicates.push({
          gameId: gameId,
          gameType: gameType,
          row: i + 2
        });
      } else {
        seen.add(key);
      }
    }
    
    if (duplicates.length === 0) {
      ss.toast('✅ No duplicates found!', '✅', 3);
      return;
    }
    
    const dupList = duplicates.map(d => `Row ${d.row}: ${d.gameId} (${d.gameType})`).join('\n');
    
    SpreadsheetApp.getUi().alert(
      `Found ${duplicates.length} duplicate game(s)`,
      `The following games appear to be duplicates:\n\n${dupList}\n\nYou may want to manually review and delete them.`
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== SYNC EFFECTIVE RATINGS TO GAMES SHEET =====
function syncEffectiveRatingsToGames() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Sync Effective Ratings to Games Sheet?',
    'This will update "Rating Before" and "Rating Δ" in the Games sheet with the Effective values from the Ratings sheet (which use callback data when available).\n\nThis overwrites the "Last method" estimates with accurate data.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Syncing effective ratings to Games sheet...', '⏳', -1);
    
    const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
    const ratingsSheet = getRatingsSheet();
    
    if (!gamesSheet || !ratingsSheet) {
      throw new Error('Required sheets not found');
    }
    
    const ratingsLastRow = ratingsSheet.getLastRow();
    if (ratingsLastRow <= 1) {
      ss.toast('No ratings data found', 'ℹ️', 3);
      return;
    }
    
    // Read all ratings data
    const ratingsData = ratingsSheet.getRange(2, 1, ratingsLastRow - 1, RATINGS_COLS.OPP_DELTA_EFFECTIVE).getValues();
    
    // Build index of effective ratings by game ID
    const effectiveRatings = {};
    let validCount = 0;
    
    for (const row of ratingsData) {
      const gameId = String(row[RATINGS_COLS.GAME_ID - 1]);
      const myPregameEffective = row[RATINGS_COLS.MY_PREGAME_EFFECTIVE - 1];
      const myDeltaEffective = row[RATINGS_COLS.MY_DELTA_EFFECTIVE - 1];
      
      // Only include if effective values exist (not null/empty)
      if (myPregameEffective !== null && myPregameEffective !== '' &&
          myDeltaEffective !== null && myDeltaEffective !== '') {
        effectiveRatings[gameId] = {
          ratingBefore: myPregameEffective,
          ratingDelta: myDeltaEffective
        };
        validCount++;
      }
    }
    
    Logger.log(`Found ${validCount} games with effective ratings`);
    
    if (validCount === 0) {
      ss.toast('No effective ratings to sync (run callback enrichment first)', 'ℹ️', 5);
      return;
    }
    
    // Read all game IDs from Games sheet
    const gamesLastRow = gamesSheet.getLastRow();
    if (gamesLastRow <= 1) {
      ss.toast('No games found', 'ℹ️', 3);
      return;
    }
    
    const gameIds = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, gamesLastRow - 1, 1).getValues();
    
    // Prepare updates
    const updates = [];
    let updateCount = 0;
    
    for (let i = 0; i < gameIds.length; i++) {
      const gameId = String(gameIds[i][0]);
      const effective = effectiveRatings[gameId];
      
      if (effective) {
        updates.push({
          row: i + 2,
          ratingBefore: effective.ratingBefore,
          ratingDelta: effective.ratingDelta
        });
        updateCount++;
      }
    }
    
    Logger.log(`Preparing to update ${updateCount} games`);
    
    if (updateCount === 0) {
      ss.toast('No matching games to update', 'ℹ️', 3);
      return;
    }
    
    // Batch write updates
    ss.toast(`Updating ${updateCount} games...`, '⏳', -1);
    
    // Update in chunks to avoid timeouts
    const chunkSize = 500;
    for (let i = 0; i < updates.length; i += chunkSize) {
      const chunk = updates.slice(i, Math.min(i + chunkSize, updates.length));
      
      for (const update of chunk) {
        gamesSheet.getRange(update.row, GAMES_COLS.RATING_BEFORE, 1, 2).setValues([[
          update.ratingBefore,
          update.ratingDelta
        ]]);
      }
      
      ss.toast(`Updated ${Math.min(i + chunkSize, updates.length)} / ${updateCount} games...`, '⏳', -1);
    }
    
    ss.toast(`✅ Updated ${updateCount} games with effective ratings!`, '✅', 5);
    
    ui.alert(
      'Sync Complete!',
      `Updated ${updateCount} games in the Games sheet.\n\n` +
      `"Rating Before" and "Rating Δ" now reflect accurate callback data where available.\n\n` +
      `Games without callback data still use the "Last method" estimates.`
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== SYNC SPECIFIC GAME =====
function syncSingleGameRating(gameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const ratingsSheet = getRatingsSheet();
  
  if (!gamesSheet || !ratingsSheet) {
    throw new Error('Required sheets not found');
  }
  
  // Find game in Ratings sheet
  const ratingsLastRow = ratingsSheet.getLastRow();
  const ratingsData = ratingsSheet.getRange(2, 1, ratingsLastRow - 1, RATINGS_COLS.OPP_DELTA_EFFECTIVE).getValues();
  
  let effectiveRating = null;
  for (const row of ratingsData) {
    if (String(row[RATINGS_COLS.GAME_ID - 1]) === String(gameId)) {
      effectiveRating = {
        ratingBefore: row[RATINGS_COLS.MY_PREGAME_EFFECTIVE - 1],
        ratingDelta: row[RATINGS_COLS.MY_DELTA_EFFECTIVE - 1]
      };
      break;
    }
  }
  
  if (!effectiveRating) {
    Logger.log(`No effective rating found for game ${gameId}`);
    return false;
  }
  
  // Find game in Games sheet
  const gamesLastRow = gamesSheet.getLastRow();
  const gameIds = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, gamesLastRow - 1, 1).getValues();
  
  for (let i = 0; i < gameIds.length; i++) {
    if (String(gameIds[i][0]) === String(gameId)) {
      const rowNum = i + 2;
      gamesSheet.getRange(rowNum, GAMES_COLS.RATING_BEFORE, 1, 2).setValues([[
        effectiveRating.ratingBefore,
        effectiveRating.ratingDelta
      ]]);
      Logger.log(`Updated game ${gameId} in Games sheet`);
      return true;
    }
  }
  
  Logger.log(`Game ${gameId} not found in Games sheet`);
  return false;
}

// ===== AUTO-SYNC AFTER CALLBACK BATCH =====
function syncRecentCallbacksToGames(count = 50) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName(SHEETS.GAMES);
  const ratingsSheet = getRatingsSheet();
  
  if (!gamesSheet || !ratingsSheet) {
    return;
  }
  
  const ratingsLastRow = ratingsSheet.getLastRow();
  if (ratingsLastRow <= 1) return;
  
  // Get last N rows from Ratings (recently updated)
  const startRow = Math.max(2, ratingsLastRow - count + 1);
  const numRows = ratingsLastRow - startRow + 1;
  
  const recentRatings = ratingsSheet.getRange(startRow, 1, numRows, RATINGS_COLS.OPP_DELTA_EFFECTIVE).getValues();
  
  let syncedCount = 0;
  
  for (const row of recentRatings) {
    const gameId = String(row[RATINGS_COLS.GAME_ID - 1]);
    const status = row[RATINGS_COLS.CALLBACK_STATUS - 1];
    const myPregame = row[RATINGS_COLS.MY_PREGAME_EFFECTIVE - 1];
    const myDelta = row[RATINGS_COLS.MY_DELTA_EFFECTIVE - 1];
    
    // Only sync if callback was fetched and has valid data
    if (status === CALLBACK_STATUS.FETCHED && myPregame !== null && myDelta !== null) {
      if (syncSingleGameRating(gameId)) {
        syncedCount++;
      }
    }
  }
  
  if (syncedCount > 0) {
    Logger.log(`Auto-synced ${syncedCount} games to Games sheet`);
  }
}
