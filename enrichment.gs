// ===== CALLBACK ENRICHMENT (SELF-HEALING) =====

function enrichWithCallbacks() {
  enrichWithCallbacksBatch(CONFIG.CALLBACK_BATCH_SIZE);
}

function enrichWithCallbacksLarge() {
  enrichWithCallbacksBatch(250);
}

function enrichWithCallbacksBatch(batchSize) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!CONFIG.CONTROL_SPREADSHEET_ID) {
    SpreadsheetApp.getUi().alert('❌ Control Spreadsheet not set up! Run "Setup Control Spreadsheet" first.');
    return;
  }
  
  if (!CONFIG.CALLBACK_DATA_SPREADSHEET_ID) {
    SpreadsheetApp.getUi().alert('❌ Callback Data Spreadsheet not set up! Run "Setup Callback Data Spreadsheet" first.');
    return;
  }
  
  // Define counters OUTSIDE try block
  let successCount = 0;
  let errorCount = 0;
  let noRatingCount = 0;
  let invalidCount = 0;
  let skippedCount = 0;
  
  try {
    ss.toast('Finding games needing callbacks...', '⏳', -1);
    
    const gamesToEnrich = getCallbackQueueSelfHealing(batchSize);
    
    if (!gamesToEnrich.length) {
      ss.toast('✅ No games need callback enrichment!', '✅', 3);
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      `Enrich ${gamesToEnrich.length} game(s) with callback data?`,
      `This will fetch detailed game data from Chess.com.\n\nContinue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    ss.toast('Building indexes...', '⏳', -1);
    const ratingsIndex = buildRatingsIndex();
    const registryIndex = buildRegistryIndex();
    const callbackIndex = buildCallbackIndex();
    
    ss.toast(`Fetching callbacks for ${gamesToEnrich.length} games...`, '⏳', -1);
    
    for (let i = 0; i < gamesToEnrich.length; i++) {
      const game = gamesToEnrich[i];
      
      try {
        if (i % 10 === 0) {
          ss.toast(`Fetching callback ${i + 1} of ${gamesToEnrich.length}...`, '⏳', -1);
        }
        
        if (callbackIndex[game.gameId]) {
          Logger.log(`Game ${game.gameId} already in Callback Data, skipping fetch`);
          
          const registryRow = registryIndex[game.gameId];
          const ratingsRow = ratingsIndex[game.gameId];
          
          if (registryRow && ratingsRow) {
            const currentStatus = getRatingCallbackStatus(ratingsRow);
            if (currentStatus === null || currentStatus === '') {
              const callbackData = getCallbackDataById(game.gameId);
              if (callbackData) {
                const ratingDataValid = hasValidRatingData(callbackData);
                const status = ratingDataValid ? CALLBACK_STATUS.FETCHED : CALLBACK_STATUS.NO_RATING;
                
                updateRatingWithCallbackDirect(ratingsRow, callbackData, ratingDataValid);
                updateRegistryCallbackStatusDirect(registryRow, status);
                
                Logger.log(`Repaired status for game ${game.gameId}`);
              }
            }
          }
          
          skippedCount++;
          continue;
        }
        
        const callbackData = fetchCallbackData(game);
        
        if (!callbackData) {
          updateRegistryCallbackStatusDirect(registryIndex[game.gameId], CALLBACK_STATUS.ERROR);
          errorCount++;
          Utilities.sleep(300);
          continue;
        }
        
        const ratingsRow = ratingsIndex[game.gameId];
        const registryRow = registryIndex[game.gameId];
        
        if (!ratingsRow || !registryRow) {
          Logger.log(`Game ${game.gameId} not found in Ratings or Registry`);
          invalidCount++;
          Utilities.sleep(300);
          continue;
        }
        
        const ratingDataValid = hasValidRatingData(callbackData);
        
        try {
          saveCallbackDataDirect(callbackData);
          
          if (ratingDataValid) {
            updateRatingWithCallbackDirect(ratingsRow, callbackData, true);
            updateRegistryCallbackStatusDirect(registryRow, CALLBACK_STATUS.FETCHED);
            successCount++;
          } else {
            updateRatingWithCallbackDirect(ratingsRow, callbackData, false);
            updateRegistryCallbackStatusDirect(registryRow, CALLBACK_STATUS.NO_RATING);
            noRatingCount++;
            Logger.log(`Game ${game.gameId}: No rating change (unrated or provisional)`);
          }
          
        } catch (writeError) {
          Logger.log(`Error writing game ${game.gameId}: ${writeError.message}`);
          try {
            updateRegistryCallbackStatusDirect(registryRow, CALLBACK_STATUS.ERROR);
          } catch (e) {
            Logger.log(`Could not update registry status: ${e.message}`);
          }
          errorCount++;
        }
        
        Utilities.sleep(300);
        
      } catch (error) {
        Logger.log(`Error processing callback for game ${game.gameId}: ${error.message}`);
        const registryRow = registryIndex[game.gameId];
        if (registryRow) {
          try {
            updateRegistryCallbackStatusDirect(registryRow, CALLBACK_STATUS.ERROR);
          } catch (e) {
            Logger.log(`Could not update registry: ${e.message}`);
          }
        }
        logToConfig(`Callback error for ${game.gameId}: ${error.message}`, 'ERROR');
        errorCount++;
      }
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  } finally {
    // This runs whether try succeeded or failed
    
    // Update progress based on actual data
    try {
      updateEnrichmentProgressFromData();
      setConfig('Last Callback Run Date', new Date());
    } catch (e) {
      Logger.log(`Error updating progress: ${e.message}`);
    }
    
    // AUTO-SYNC: Update Games sheet with newly fetched effective ratings
    if (successCount > 0) {
      ss.toast('Syncing effective ratings to Games sheet...', '⏳', -1);
      try {
        syncRecentCallbacksToGames(successCount + noRatingCount);
        Logger.log(`Auto-synced ${successCount} effective ratings to Games sheet`);
      } catch (syncError) {
        Logger.log(`Auto-sync error: ${syncError.message}`);
      }
    }
    
    const message = `✅ Callbacks:\n${successCount} valid\n${noRatingCount} no rating change\n${skippedCount} already done\n${errorCount} errors\n${invalidCount} not found`;
    ss.toast(message, '✅', 6);
  }
}

// ===== GET CALLBACK QUEUE (SELF-HEALING) =====
function getCallbackQueueSelfHealing(batchSize) {
  const registrySheet = getRegistrySheet();
  
  if (!registrySheet) {
    throw new Error('Registry sheet not found');
  }
  
  const lastRow = registrySheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  // Read ALL registry data (not just from last processed position)
  const registryData = registrySheet.getRange(2, 1, lastRow - 1, REGISTRY_COLS.DATA_LOCATION).getValues();
  
  // Build list of games needing callbacks
  const gameIdsNeeded = [];
  
  for (let i = 0; i < registryData.length; i++) {
    const row = registryData[i];
    const gameId = String(row[REGISTRY_COLS.GAME_ID - 1]);
    const callbackStatus = row[REGISTRY_COLS.CALLBACK_STATUS - 1];
    
    // Need callback if: null, empty, or error
    if (callbackStatus === null || callbackStatus === '' || callbackStatus === CALLBACK_STATUS.ERROR) {
      gameIdsNeeded.push(gameId);
      
      if (gameIdsNeeded.length >= batchSize) break;
    }
  }
  
  if (gameIdsNeeded.length === 0) {
    return [];
  }
  
  // Get game info from Games sheet
  const gamesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.GAMES);
  if (!gamesSheet) {
    throw new Error('Games sheet not found');
  }
  
  const gamesLastRow = gamesSheet.getLastRow();
  if (gamesLastRow <= 1) {
    return [];
  }
  
  const gamesData = gamesSheet.getRange(2, 1, gamesLastRow - 1, GAMES_COLS.OPPONENT).getValues();
  
  const gameIndex = {};
  for (const row of gamesData) {
    const gameId = String(row[GAMES_COLS.GAME_ID - 1]);
    gameIndex[gameId] = {
      gameUrl: row[GAMES_COLS.GAME_URL - 1],
      timeClass: row[GAMES_COLS.TIME_CLASS - 1],
      opponent: row[GAMES_COLS.OPPONENT - 1]
    };
  }
  
  const queue = [];
  
  for (const gameId of gameIdsNeeded) {
    const gameInfo = gameIndex[gameId];
    
    if (gameInfo && gameInfo.gameUrl && gameInfo.timeClass) {
      queue.push({
        gameId: gameId,
        gameUrl: gameInfo.gameUrl,
        timeClass: gameInfo.timeClass,
        white: CONFIG.USERNAME,
        black: gameInfo.opponent
      });
    }
  }
  
  Logger.log(`Built queue of ${queue.length} games needing callbacks`);
  
  return queue;
}

// ===== BUILD CALLBACK INDEX =====
function buildCallbackIndex() {
  const callbackSheet = getCallbackDataSheet();
  const lastRow = callbackSheet.getLastRow();
  
  if (lastRow <= 1) return {};
  
  const gameIds = callbackSheet.getRange(2, CALLBACK_COLS.GAME_ID, lastRow - 1, 1).getValues();
  
  const index = {};
  for (let i = 0; i < gameIds.length; i++) {
    const gameId = String(gameIds[i][0]);
    if (gameId) {
      index[gameId] = i + 2;
    }
  }
  
  Logger.log(`Built callback index with ${Object.keys(index).length} games`);
  return index;
}

// ===== GET CALLBACK DATA BY ID =====
function getCallbackDataById(gameId) {
  const callbackSheet = getCallbackDataSheet();
  const lastRow = callbackSheet.getLastRow();
  
  if (lastRow <= 1) return null;
  
  const gameIds = callbackSheet.getRange(2, CALLBACK_COLS.GAME_ID, lastRow - 1, 1).getValues();
  
  for (let i = 0; i < gameIds.length; i++) {
    if (String(gameIds[i][0]) === String(gameId)) {
      const rowNum = i + 2;
      const rowData = callbackSheet.getRange(rowNum, 1, 1, CALLBACK_COLS.DATE_FETCHED).getValues()[0];
      
      return {
        gameId: rowData[CALLBACK_COLS.GAME_ID - 1],
        gameUrl: rowData[CALLBACK_COLS.GAME_URL - 1],
        callbackUrl: rowData[CALLBACK_COLS.CALLBACK_URL - 1],
        myRatingChange: rowData[CALLBACK_COLS.MY_RATING_CHANGE - 1],
        oppRatingChange: rowData[CALLBACK_COLS.OPP_RATING_CHANGE - 1],
        myRatingBefore: rowData[CALLBACK_COLS.MY_RATING_BEFORE - 1],
        oppRatingBefore: rowData[CALLBACK_COLS.OPP_RATING_BEFORE - 1]
      };
    }
  }
  
  return null;
}

// ===== GET RATING CALLBACK STATUS =====
function getRatingCallbackStatus(ratingsRow) {
  const ratingsSheet = getRatingsSheet();
  return ratingsSheet.getRange(ratingsRow, RATINGS_COLS.CALLBACK_STATUS).getValue();
}

// ===== CHECK IF RATING DATA IS VALID =====
function hasValidRatingData(callbackData) {
  const myDelta = callbackData.myRatingChange;
  const oppDelta = callbackData.oppRatingChange;
  
  // Valid if at least one has a non-zero rating change
  return !((myDelta === null || myDelta === 0) && (oppDelta === null || oppDelta === 0));
}

// ===== UPDATE ENRICHMENT PROGRESS FROM DATA =====
function updateEnrichmentProgressFromData() {
  const registrySheet = getRegistrySheet();
  const lastRow = registrySheet.getLastRow();
  
  if (lastRow <= 1) {
    setConfig('Total Games', 0);
    setConfig('Callbacks Complete', 0);
    setConfig('Callbacks Pending', 0);
    setConfig('Callbacks No Rating', 0);
    setConfig('Callbacks Error', 0);
    return;
  }
  
  const statuses = registrySheet.getRange(2, REGISTRY_COLS.CALLBACK_STATUS, lastRow - 1, 1).getValues();
  
  let complete = 0;
  let pending = 0;
  let noRating = 0;
  let errors = 0;
  
  for (const [status] of statuses) {
    if (status === CALLBACK_STATUS.FETCHED) {
      complete++;
    } else if (status === CALLBACK_STATUS.NO_RATING) {
      noRating++;
    } else if (status === CALLBACK_STATUS.ERROR) {
      errors++;
    } else {
      pending++;
    }
  }
  
  setConfig('Total Games', lastRow - 1);
  setConfig('Callbacks Complete', complete);
  setConfig('Callbacks Pending', pending);
  setConfig('Callbacks No Rating', noRating);
  setConfig('Callbacks Error', errors);
}

// ===== DIRECT WRITE FUNCTIONS =====
function saveCallbackDataDirect(callbackData) {
  const callbackSheet = getCallbackDataSheet();
  
  const row = [
    callbackData.gameId,
    callbackData.gameUrl,
    callbackData.callbackUrl,
    callbackData.endTime,
    callbackData.myColor,
    callbackData.timeClass,
    callbackData.myRating,
    callbackData.oppRating,
    callbackData.myRatingChange,
    callbackData.oppRatingChange,
    callbackData.myRatingBefore,
    callbackData.oppRatingBefore,
    callbackData.baseTime,
    callbackData.timeIncrement,
    callbackData.moveTimestamps,
    callbackData.myUsername,
    callbackData.myCountry,
    callbackData.myMembership,
    callbackData.myMemberSince,
    callbackData.myDefaultTab,
    callbackData.myPostMoveAction,
    callbackData.myLocation,
    callbackData.oppUsername,
    callbackData.oppCountry,
    callbackData.oppMembership,
    callbackData.oppMemberSince,
    callbackData.oppDefaultTab,
    callbackData.oppPostMoveAction,
    callbackData.oppLocation,
    new Date()
  ];
  
  const lastRow = callbackSheet.getLastRow();
  callbackSheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function updateRatingWithCallbackDirect(ratingsRow, callbackData, updateEffective) {
  const ratingsSheet = getRatingsSheet();
  const now = new Date();
  
  if (updateEffective) {
    const data = [
      CALLBACK_STATUS.FETCHED,
      now,
      callbackData.myRatingBefore,
      callbackData.myRatingChange,
      callbackData.oppRatingBefore,
      callbackData.oppRatingChange,
      callbackData.myRatingBefore,
      callbackData.myRatingChange,
      callbackData.oppRatingBefore,
      callbackData.oppRatingChange
    ];
    
    ratingsSheet.getRange(ratingsRow, RATINGS_COLS.CALLBACK_STATUS, 1, 10).setValues([data]);
    
  } else {
    const data = [
      CALLBACK_STATUS.NO_RATING,
      now,
      callbackData.myRatingBefore,
      callbackData.myRatingChange,
      callbackData.oppRatingBefore,
      callbackData.oppRatingChange
    ];
    
    ratingsSheet.getRange(ratingsRow, RATINGS_COLS.CALLBACK_STATUS, 1, 6).setValues([data]);
  }
}

function updateRegistryCallbackStatusDirect(registryRow, status) {
  if (!registryRow) return;
  
  const registrySheet = getRegistrySheet();
  const now = new Date();
  
  registrySheet.getRange(registryRow, REGISTRY_COLS.CALLBACK_STATUS, 1, 2).setValues([[status, now]]);
}

// Keep other functions (buildRatingsIndex, buildRegistryIndex, fetchCallbackData, etc.) as before

function showEnrichmentProgress() {
  updateEnrichmentProgressFromData();
  
  const total = getConfig('Total Games') || 0;
  const complete = getConfig('Callbacks Complete') || 0;
  const pending = getConfig('Callbacks Pending') || 0;
  const noRating = getConfig('Callbacks No Rating') || 0;
  const errors = getConfig('Callbacks Error') || 0;
  const lastRun = getConfig('Last Callback Run Date') || 'Never';
  
  const completePercent = total > 0 ? ((complete / total) * 100).toFixed(1) : 0;
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { color: #4285f4; }
      .stat { margin: 10px 0; font-size: 14px; }
      .complete { color: #0f9d58; font-weight: bold; }
      .pending { color: #f4b400; }
      .norating { color: #9e9e9e; }
      .error { color: #db4437; }
    </style>
    <h2>📊 Callback Enrichment Progress</h2>
    <div class="stat"><strong>Total Games:</strong> ${total}</div>
    <div class="stat complete">✅ Fetched (Valid Ratings): ${complete} (${completePercent}%)</div>
    <div class="stat norating">📊 Fetched (No Rating Change): ${noRating}</div>
    <div class="stat pending">⏳ Pending: ${pending}</div>
    <div class="stat error">⚠️ Errors: ${errors}</div>
    <hr>
    <div class="stat"><strong>Last Run:</strong> ${lastRun}</div>
    <hr>
    <p><strong>Note:</strong> "No Rating Change" means callback was fetched successfully,
    but the game had no rating impact (unrated or provisional). All other data is still saved.</p>
    <hr>
    <p>Run "Callback Enrichment" from menu to continue.</p>
  `)
    .setWidth(450)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enrichment Progress');
}

function buildRatingsIndex() {
  const ratingsSheet = getRatingsSheet();
  const lastRow = ratingsSheet.getLastRow();
  
  if (lastRow <= 1) return {};
  
  const gameIds = ratingsSheet.getRange(2, RATINGS_COLS.GAME_ID, lastRow - 1, 1).getValues();
  
  const index = {};
  for (let i = 0; i < gameIds.length; i++) {
    const gameId = String(gameIds[i][0]);
    if (gameId) {
      index[gameId] = i + 2;
    }
  }
  
  Logger.log(`Built ratings index with ${Object.keys(index).length} games`);
  return index;
}

// ===== BUILD REGISTRY INDEX =====
function buildRegistryIndex() {
  const registrySheet = getRegistrySheet();
  const lastRow = registrySheet.getLastRow();
  
  if (lastRow <= 1) return {};
  
  const gameIds = registrySheet.getRange(2, REGISTRY_COLS.GAME_ID, lastRow - 1, 1).getValues();
  
  const index = {};
  for (let i = 0; i < gameIds.length; i++) {
    const gameId = String(gameIds[i][0]);
    if (gameId) {
      index[gameId] = i + 2;
    }
  }
  
  Logger.log(`Built registry index with ${Object.keys(index).length} games`);
  return index;
}

// ===== FETCH CALLBACK DATA =====
function fetchCallbackData(game) {
  if (!game || !game.gameId || !game.timeClass) {
    Logger.log(`Skipping callback fetch - incomplete game data: ${JSON.stringify(game)}`);
    return null;
  }
  
  const gameId = game.gameId;
  const timeClass = game.timeClass.toLowerCase();
  const gameType = timeClass === 'daily' ? 'daily' : 'live';
  const callbackUrl = `https://www.chess.com/callback/${gameType}/game/${gameId}`;
  
  try {
    const response = UrlFetchApp.fetch(callbackUrl, {muteHttpExceptions: true});
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`Callback API error for game ${gameId}: ${response.getResponseCode()}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data || !data.game) {
      Logger.log(`Invalid callback data for game ${gameId}`);
      return null;
    }
    
    const gameData = data.game;
    const players = data.players || {};
    const topPlayer = players.top || {};
    const bottomPlayer = players.bottom || {};
    
    let whitePlayer, blackPlayer;
    if (topPlayer.color === 'white') {
      whitePlayer = topPlayer;
      blackPlayer = bottomPlayer;
    } else {
      whitePlayer = bottomPlayer;
      blackPlayer = topPlayer;
    }
    
    let isWhite = false;
    if (whitePlayer.username && whitePlayer.username.toLowerCase() === CONFIG.USERNAME.toLowerCase()) {
      isWhite = true;
    }
    
    const myColor = isWhite ? 'white' : 'black';
    
    let myRatingChange = isWhite ? gameData.ratingChangeWhite : gameData.ratingChangeBlack;
    let oppRatingChange = isWhite ? gameData.ratingChangeBlack : gameData.ratingChangeWhite;
    
    const myPlayer = isWhite ? whitePlayer : blackPlayer;
    const oppPlayer = isWhite ? blackPlayer : whitePlayer;
    
    const myRating = myPlayer.rating || null;
    const oppRating = oppPlayer.rating || null;
    
    let myRatingBefore = null;
    let oppRatingBefore = null;
    
    if (myRating !== null && myRatingChange !== null && myRatingChange !== undefined) {
      myRatingBefore = myRating - myRatingChange;
    }
    if (oppRating !== null && oppRatingChange !== null && oppRatingChange !== undefined) {
      oppRatingBefore = oppRating - oppRatingChange;
    }
    
    return {
      gameId: gameId,
      gameUrl: game.gameUrl,
      callbackUrl: callbackUrl,
      endTime: gameData.endTime,
      myColor: myColor,
      timeClass: game.timeClass,
      myRating: myRating,
      oppRating: oppRating,
      myRatingChange: myRatingChange,
      oppRatingChange: oppRatingChange,
      myRatingBefore: myRatingBefore,
      oppRatingBefore: oppRatingBefore,
      baseTime: gameData.baseTime1 || 0,
      timeIncrement: gameData.timeIncrement1 || 0,
      moveTimestamps: gameData.moveTimestamps ? String(gameData.moveTimestamps) : '',
      myUsername: myPlayer.username || '',
      myCountry: myPlayer.countryName || '',
      myMembership: myPlayer.membershipCode || '',
      myMemberSince: myPlayer.memberSince || 0,
      myDefaultTab: myPlayer.defaultTab || null,
      myPostMoveAction: myPlayer.postMoveAction || '',
      myLocation: myPlayer.location || '',
      oppUsername: oppPlayer.username || '',
      oppCountry: oppPlayer.countryName || '',
      oppMembership: oppPlayer.membershipCode || '',
      oppMemberSince: oppPlayer.memberSince || 0,
      oppDefaultTab: oppPlayer.defaultTab || null,
      oppPostMoveAction: oppPlayer.postMoveAction || '',
      oppLocation: oppPlayer.location || ''
    };
    
  } catch (error) {
    Logger.log(`Error fetching callback data for game ${gameId}: ${error.message}`);
    return null;
  }
}

// ===== UPDATE ENRICHMENT PROGRESS =====
function updateEnrichmentProgress() {
  const registrySheet = getRegistrySheet();
  const lastRow = registrySheet.getLastRow();
  
  if (lastRow <= 1) {
    setConfig('Total Games', 0);
    setConfig('Callbacks Complete', 0);
    setConfig('Callbacks Pending', 0);
    setConfig('Callbacks Invalid', 0);
    setConfig('Callbacks Error', 0);
    return;
  }
  
  const statuses = registrySheet.getRange(2, REGISTRY_COLS.CALLBACK_STATUS, lastRow - 1, 1).getValues();
  
  let complete = 0;
  let pending = 0;
  let invalid = 0;
  let errors = 0;
  
  for (const [status] of statuses) {
    if (status === 'fetched') {
      complete++;
    } else if (status === 'invalid') {
      invalid++;
    } else if (status === 'error') {
      errors++;
    } else {
      pending++;
    }
  }
  
  setConfig('Total Games', lastRow - 1);
  setConfig('Callbacks Complete', complete);
  setConfig('Callbacks Pending', pending);
  setConfig('Callbacks Invalid', invalid);
  setConfig('Callbacks Error', errors);
}
