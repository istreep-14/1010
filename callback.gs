// ===== COMPLETE CALLBACK SYSTEM WITH ALL DATA FIELDS =====
// Based on your old working code - gets detailed player info for both players

// ===== IMPROVED RATINGS TRACKER =====
class ImprovedRatingsTracker {
  constructor() {
    this.ledger = {};
    this.loaded = false;
  }
  
  loadFromGamesSheet() {
    if (this.loaded) return;
    
    const gamesSheet = getGamesSheet();
    const lastRow = gamesSheet.getLastRow();
    
    if (lastRow <= 1) {
      this.loaded = true;
      Logger.log('No existing games found - starting fresh');
      return;
    }
    
    Logger.log(`Loading ratings from ${lastRow - 1} existing games...`);
    
    const formatCol = GAMES_COLS.FORMAT;
    const myRatingCol = GAMES_COLS.MY_RATING;
    
    const formatData = gamesSheet.getRange(2, formatCol, lastRow - 1, 1).getValues();
    const ratingData = gamesSheet.getRange(2, myRatingCol, lastRow - 1, 1).getValues();
    
    const formatLastRating = {};
    
    for (let i = 0; i < formatData.length; i++) {
      const format = formatData[i][0];
      const rating = ratingData[i][0];
      
      if (format && rating) {
        formatLastRating[format] = rating;
      }
    }
    
    this.ledger = formatLastRating;
    this.loaded = true;
    
    Logger.log('Loaded ratings ledger: ' + JSON.stringify(this.ledger));
  }
  
  calculateRating(format, currentRating) {
    const ratingBefore = this.ledger[format] || null;
    const ratingDelta = (ratingBefore !== null && currentRating !== null) 
      ? (currentRating - ratingBefore) 
      : null;
    
    if (currentRating !== null) {
      this.ledger[format] = currentRating;
    }
    
    return {
      before: ratingBefore,
      delta: ratingDelta
    };
  }
}

// ===== SETUP CALLBACK SHEET WITH ALL COLUMNS =====
function setupCallbackSheet() {
  const ss = getEnrichmentSpreadsheet();
  const sheet = ss.getSheetByName('Callback Data') || ss.insertSheet('Callback Data');
  
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Game ID',
      'Game URL',
      'Callback URL',
      'End Time',
      'My Color',
      'Time Class',
      'My Rating',
      'Opp Rating',
      'My Rating Change',
      'Opp Rating Change',
      'My Rating Before',
      'Opp Rating Before',
      'Base Time',
      'Time Increment',
      'Move Timestamps',
      'My Username',
      'My Country',
      'My Membership',
      'My Member Since',
      'My Default Tab',
      'My Post Move Action',
      'My Location',
      'Opp Username',
      'Opp Country',
      'Opp Membership',
      'Opp Member Since',
      'Opp Default Tab',
      'Opp Post Move Action',
      'Opp Location',
      'Date Fetched'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    sheet.setFrozenRows(1);
    
    // Format date columns
    sheet.getRange('D:D').setNumberFormat('M/D/YY h:mm');  // End Time
    sheet.getRange('S:S').setNumberFormat('M/D/YY');        // My Member Since
    sheet.getRange('Z:Z').setNumberFormat('M/D/YY');        // Opp Member Since
    sheet.getRange('AD:AD').setNumberFormat('M/D/YY h:mm'); // Date Fetched
    
    // Format rating columns
    sheet.getRange('G:L').setNumberFormat('0');  // All rating columns
    
    // Set column widths
    sheet.setColumnWidth(1, 90);   // Game ID
    sheet.setColumnWidth(2, 250);  // Game URL
    sheet.setColumnWidth(3, 300);  // Callback URL
    sheet.setColumnWidth(15, 300); // Move Timestamps (wide for data)
  }
  
  return sheet;
}

// ===== FETCH CALLBACK DATA (YOUR OLD WORKING CODE) =====
function fetchCallbackData(game) {
  if (!game || !game.gameId || !game.timeClass) {
    Logger.log(`Skipping callback fetch - incomplete game data: ${JSON.stringify(game)}`);
    return null;
  }
  
  const gameId = game.gameId;
  const timeClass = game.timeClass.toLowerCase();
  const gameType = timeClass === 'daily' ? 'daily' : 'live';
  const callbackUrl = `https://www.chess.com/callback/${gameType}/game/${gameId}`;
  
  Logger.log(`Fetching callback: ${callbackUrl}`);
  
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
    
    // Determine which player is white/black
    let whitePlayer, blackPlayer;
    if (topPlayer.color === 'white') {
      whitePlayer = topPlayer;
      blackPlayer = bottomPlayer;
    } else {
      whitePlayer = bottomPlayer;
      blackPlayer = topPlayer;
    }
    
    // Determine if we're white or black
    let isWhite = false;
    if (whitePlayer.username && whitePlayer.username.toLowerCase() === CONFIG.USERNAME.toLowerCase()) {
      isWhite = true;
    }
    
    const myColor = isWhite ? 'white' : 'black';
    
    // Get rating changes
    let myRatingChange = isWhite ? gameData.ratingChangeWhite : gameData.ratingChangeBlack;
    let oppRatingChange = isWhite ? gameData.ratingChangeBlack : gameData.ratingChangeWhite;
    
    // Get player objects
    const myPlayer = isWhite ? whitePlayer : blackPlayer;
    const oppPlayer = isWhite ? blackPlayer : whitePlayer;
    
    // Get current ratings
    const myRating = myPlayer.rating || null;
    const oppRating = oppPlayer.rating || null;
    
    // Calculate ratings before
    let myRatingBefore = null;
    let oppRatingBefore = null;
    
    if (myRating !== null && myRatingChange !== null && myRatingChange !== undefined) {
      myRatingBefore = myRating - myRatingChange;
    }
    if (oppRating !== null && oppRatingChange !== null && oppRatingChange !== undefined) {
      oppRatingBefore = oppRating - oppRatingChange;
    }
    
    Logger.log(`Callback data fetched successfully for game ${gameId}`);
    Logger.log(`  My rating: ${myRatingBefore} → ${myRating} (${myRatingChange > 0 ? '+' : ''}${myRatingChange})`);
    Logger.log(`  Opp rating: ${oppRatingBefore} → ${oppRating} (${oppRatingChange > 0 ? '+' : ''}${oppRatingChange})`);
    
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
    Logger.log(`Stack: ${error.stack}`);
    return null;
  }
}

// ===== STORE CALLBACK DATA TO SHEET =====
function storeCallbackData(sheet, callbackData) {
  if (!callbackData) return;
  
  const now = new Date();
  
  // Convert epoch timestamps to dates if needed
  const endTime = callbackData.endTime ? new Date(callbackData.endTime * 1000) : null;
  const myMemberSince = callbackData.myMemberSince ? new Date(callbackData.myMemberSince * 1000) : null;
  const oppMemberSince = callbackData.oppMemberSince ? new Date(callbackData.oppMemberSince * 1000) : null;
  
  const row = [
    callbackData.gameId,
    callbackData.gameUrl,
    callbackData.callbackUrl,
    endTime,
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
    myMemberSince,
    callbackData.myDefaultTab,
    callbackData.myPostMoveAction,
    callbackData.myLocation,
    callbackData.oppUsername,
    callbackData.oppCountry,
    callbackData.oppMembership,
    oppMemberSince,
    callbackData.oppDefaultTab,
    callbackData.oppPostMoveAction,
    callbackData.oppLocation,
    now
  ];
  
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
  
  Logger.log(`Stored callback data for game ${callbackData.gameId} in row ${lastRow + 1}`);
}

// ===== ENRICH GAMES WITH CALLBACKS =====
function enrichNewGamesWithCallbacks(count = 20) {
  const ss = getMainSpreadsheet();
  const gamesSheet = getGamesSheet();
  const callbackSheet = setupCallbackSheet();
  
  const lastRow = gamesSheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('No games to enrich', 'ℹ️', 3);
    return;
  }
  
  // Find games that need enrichment
  const statusCol = GAMES_COLS.CALLBACK_STATUS;
  const allData = gamesSheet.getRange(2, 1, lastRow - 1, gamesSheet.getLastColumn()).getValues();
  
  // Filter to pending games
  const pendingGames = [];
  for (let i = 0; i < allData.length; i++) {
    const status = allData[i][statusCol - 1];
    if (!status || status === '' || status === 'pending' || String(status).startsWith('error')) {
      pendingGames.push({
        rowIndex: i + 2,
        gameId: allData[i][GAMES_COLS.GAME_ID - 1],
        gameUrl: allData[i][GAMES_COLS.GAME_URL - 1],
        timeClass: allData[i][GAMES_COLS.TIME_CLASS - 1],
        format: allData[i][GAMES_COLS.FORMAT - 1]
      });
    }
  }
  
  if (pendingGames.length === 0) {
    ss.toast('No pending games found', 'ℹ️', 3);
    return;
  }
  
  const gamesToEnrich = pendingGames.slice(0, count);
  ss.toast(`Enriching ${gamesToEnrich.length} games...`, '⏳', -1);
  
  let successCount = 0;
  let noRatingCount = 0;
  let errorCount = 0;
  const errorDetails = [];
  
  for (const game of gamesToEnrich) {
    try {
      // Mark as processing
      gamesSheet.getRange(game.rowIndex, statusCol).setValue('processing');
      SpreadsheetApp.flush();
      
      // Fetch callback data
      Logger.log(`\n=== Processing game ${game.gameId} ===`);
      const callbackData = fetchCallbackData(game);
      
      if (!callbackData) {
        gamesSheet.getRange(game.rowIndex, statusCol).setValue('error: no data');
        errorCount++;
        errorDetails.push(`${game.gameId}: No callback data returned`);
        continue;
      }
      
      // Store in callback sheet
      storeCallbackData(callbackSheet, callbackData);
      
      // Update ratings in Games sheet if available
      if (callbackData.myRatingBefore !== null && callbackData.myRatingChange !== null) {
        gamesSheet.getRange(game.rowIndex, GAMES_COLS.RATING_BEFORE).setValue(callbackData.myRatingBefore);
        gamesSheet.getRange(game.rowIndex, GAMES_COLS.RATING_DELTA).setValue(callbackData.myRatingChange);
        gamesSheet.getRange(game.rowIndex, statusCol).setValue('fetched');
        gamesSheet.getRange(game.rowIndex, GAMES_COLS.CALLBACK_DATE).setValue(new Date());
        successCount++;
      } else {
        gamesSheet.getRange(game.rowIndex, statusCol).setValue('no_rating');
        gamesSheet.getRange(game.rowIndex, GAMES_COLS.CALLBACK_DATE).setValue(new Date());
        noRatingCount++;
      }
      
      gamesSheet.getRange(game.rowIndex, GAMES_COLS.LAST_UPDATED).setValue(new Date());
      
      // Rate limiting
      Utilities.sleep(500);
      
    } catch (error) {
      Logger.log(`Error enriching game ${game.gameId}: ${error.message}`);
      Logger.log(`Stack: ${error.stack}`);
      gamesSheet.getRange(game.rowIndex, statusCol).setValue(`error: ${error.message}`);
      errorCount++;
      errorDetails.push(`${game.gameId}: ${error.message}`);
    }
  }
  
  // Show detailed status
  let statusMsg = `✅ Enrichment complete!\n\n` +
    `Success: ${successCount}\n` +
    `No Rating Data: ${noRatingCount}\n` +
    `Errors: ${errorCount}`;
  
  if (errorDetails.length > 0 && errorDetails.length <= 5) {
    statusMsg += `\n\nError details:\n${errorDetails.join('\n')}`;
  }
  
  Logger.log('\n' + statusMsg);
  if (errorDetails.length > 0) {
    Logger.log('All errors: ' + JSON.stringify(errorDetails, null, 2));
  }
  
  ss.toast(statusMsg, errorCount > 0 ? '⚠️' : '✅', 8);
}

// ===== ENRICH ALL PENDING =====
function enrichAllPendingCallbacks() {
  const gamesSheet = getGamesSheet();
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No games found');
    return;
  }
  
  // Count pending games
  const statuses = gamesSheet.getRange(2, GAMES_COLS.CALLBACK_STATUS, lastRow - 1, 1).getValues();
  const pendingCount = statuses.filter(row => 
    !row[0] || row[0] === '' || row[0] === 'pending' || String(row[0]).startsWith('error')
  ).length;
  
  if (pendingCount === 0) {
    SpreadsheetApp.getUi().alert('No pending games to enrich');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Enrich All Pending Games?',
    `This will enrich ${pendingCount} pending games.\n\n` +
    'This may take several minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    enrichNewGamesWithCallbacks(pendingCount);
  }
}

// ===== TEST CALLBACK FETCH =====
function testCallbackFetch() {
  const gamesSheet = getGamesSheet();
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No games found');
    return;
  }
  
  // Get the most recent game
  const gameData = gamesSheet.getRange(lastRow, 1, 1, gamesSheet.getLastColumn()).getValues()[0];
  
  const game = {
    gameId: gameData[GAMES_COLS.GAME_ID - 1],
    gameUrl: gameData[GAMES_COLS.GAME_URL - 1],
    timeClass: gameData[GAMES_COLS.TIME_CLASS - 1]
  };
  
  Logger.log('Testing callback fetch for most recent game...');
  Logger.log('Game: ' + JSON.stringify(game));
  
  const callbackData = fetchCallbackData(game);
  
  if (callbackData) {
    Logger.log('\n=== SUCCESS! ===');
    Logger.log(JSON.stringify(callbackData, null, 2));
    
    SpreadsheetApp.getUi().alert(
      'Callback Test Success!',
      `Successfully fetched callback data!\n\n` +
      `Game: ${callbackData.gameId}\n` +
      `My Rating: ${callbackData.myRatingBefore} → ${callbackData.myRating} (${callbackData.myRatingChange > 0 ? '+' : ''}${callbackData.myRatingChange})\n` +
      `Opp Rating: ${callbackData.oppRatingBefore} → ${callbackData.oppRating} (${callbackData.oppRatingChange > 0 ? '+' : ''}${callbackData.oppRatingChange})\n\n` +
      'Check View > Logs for full details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    Logger.log('\n=== FAILED ===');
    Logger.log('No callback data returned');
    
    SpreadsheetApp.getUi().alert(
      'Callback Test Failed',
      'Could not fetch callback data.\n\n' +
      'Check View > Logs for error details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== HELPER FUNCTIONS =====

class ArchiveBasedGameIndex {
  constructor() {
    this.index = null;
    this.archiveAnchors = {};
  }
  
  build() {
    if (this.index) return this.index;
    
    const gamesSheet = getGamesSheet();
    const lastRow = gamesSheet.getLastRow();
    
    this.index = new Set();
    
    if (lastRow > 1) {
      const data = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, lastRow - 1, 1).getValues();
      data.forEach(row => this.index.add(String(row[0])));
      
      const fullData = gamesSheet.getRange(2, 1, lastRow - 1, GAMES_COLS.ARCHIVE).getValues();
      
      for (let i = fullData.length - 1; i >= 0; i--) {
        const gameId = String(fullData[i][GAMES_COLS.GAME_ID - 1]);
        const archive = fullData[i][GAMES_COLS.ARCHIVE - 1];
        
        if (archive && !this.archiveAnchors[archive]) {
          this.archiveAnchors[archive] = gameId;
        }
      }
      
      Logger.log(`Built index with ${this.index.size} games`);
    }
    
    return this.index;
  }
  
  has(gameId) {
    if (!this.index) this.build();
    return this.index.has(String(gameId));
  }
}

function filterNewGames(games) {
  const index = new ArchiveBasedGameIndex();
  index.build();
  
  return games.filter(game => {
    const gameId = game.url.split('/').pop();
    return !index.has(gameId);
  });
}
