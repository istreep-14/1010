// ===== OPTIMIZED GAME PROCESSING - BASED ON OLD.GS =====
// Streamlined, fast processing using proven old.gs logic
// Integrates with current project structure

// ===== CONFIGURATION =====
const CONFIG = {
  USERNAME: 'ians141',
  MONTHS_TO_FETCH: 2 // 0 = all history
};

const SHEETS = {
  GAMES: 'Games'
};

// ===== MENU =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('♟️ Chess.com')
    .addItem('⚙️ Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('🔄 Update Games', 'fetchChesscomGames')
    .addItem('📥 Fetch All History', 'fetchAllGames')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⭐ Callbacks')
      .addItem('Test Callback Fetch', 'testCallbackFetch')
      .addItem('View Callback Logs', 'viewCallbackLogs')
      .addItem('Clear Callback Logs', 'clearCallbackLogs')
      .addSeparator()
      .addItem('Update Pending Callbacks', 'updatePendingCallbacks')
      .addItem('Enrich Recent Games (20)', 'enrichRecentGamesImmediate')
      .addItem('Enrich All Games', 'enrichAllPendingCallbacks'))
    .addSeparator()
    .addItem('📊 Update Summary Stats', 'updateSummaryStats')
    .addToUi();
}

// ===== FETCH FUNCTIONS =====
function fetchAllGames() {
  fetchChesscomGames(true);
}

function fetchChesscomGames(fetchAll = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Run "Setup Sheets" first!');
    return;
  }
  
  try {
    // Get archives to fetch
    const archives = fetchAll || CONFIG.MONTHS_TO_FETCH === 0 
      ? getAllArchives(CONFIG.USERNAME)
      : getRecentArchives(CONFIG.USERNAME, CONFIG.MONTHS_TO_FETCH);
    
    if (!archives.length) {
      ss.toast('No archives found', 'ℹ️', 3);
      return;
    }
    
    // Fetch all games from archives
    ss.toast(`Fetching ${archives.length} archive(s)...`, '⏳', -1);
    const allGames = fetchGamesFromArchives(archives);
    
    if (!allGames.length) {
      ss.toast('No games found', 'ℹ️', 3);
      return;
    }
    
    // Filter to new games only
    const newGames = filterNewGames(allGames, sheet);
    
    if (!newGames.length) {
      ss.toast('No new games', 'ℹ️', 3);
      return;
    }
    
    // Get current ratings ledger
    const ledger = getLastLedger(sheet);
    Logger.log('Starting ledger loaded: ' + JSON.stringify(ledger));
    
    // Process and write new games
    ss.toast(`Processing ${newGames.length} new games...`, '⏳', -1);
    const rows = processGames(newGames, CONFIG.USERNAME, ledger);
    writeGamesToSheet(sheet, rows);
    
    ss.toast(`✅ ${newGames.length} new games!`, '✅', 5);
    
    // Automatically enrich with callbacks if we have new games
    if (newGames.length > 0) {
      ss.toast('Starting callback enrichment...', '⏳', 3);
      enrichNewGamesWithCallbacks(newGames.length);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
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
  
  for (const url of archiveUrls) {
    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());
      if (data.games) allGames.push(...data.games);
      Utilities.sleep(300);
    } catch (e) {
      Logger.log(`Failed to fetch ${url}: ${e.message}`);
    }
  }
  
  return allGames.sort((a, b) => a.end_time - b.end_time);
}

// ===== NEW GAME DETECTION =====
function filterNewGames(games, sheet) {
  const existingGames = new Set();
  
  if (sheet.getLastRow() > 1) {
    const gameIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
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
  if (sheet.getLastRow() <= 1) return {};
  
  try {
    // Column AW is now column 49 (after removing Link column)
    const lastLedgerCell = sheet.getRange(sheet.getLastRow(), 49).getValue();
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
  
  // Note: Callback enrichment should be run manually via menu
  // or automatically after the main fetch process completes
}

// ===== GAME PROCESSING (OPTIMIZED FROM OLD.GS) =====
function processGames(games, username, ratingsLedger = {}) {
  const rows = [];
  let currentLedger = JSON.parse(JSON.stringify(ratingsLedger));
  
  for (const game of games) {
    try {
      if (!game || !game.url || !game.end_time) continue;
      
      // ===== BASIC INFO =====
      const gameId = game.url.split('/').pop();
      const gameType = (game.time_class || '').toLowerCase() === 'daily' ? 'daily' : 'live';
      const gameUrl = game.url;
      
      // ===== DATES & TIMES =====
      const endDate = new Date(game.end_time * 1000);
      const startDate = extractStartFromPGN(game.pgn);
      const duration = extractDurationFromPGN(game.pgn) || 0;
      
      const startDateTimeFormatted = startDate ? formatDateTime(startDate) : null;
      const endDateTimeFormatted = formatDateTime(endDate);
      
      const startDateFormatted = startDate ? formatDate(startDate) : null;
      const startTimeFormatted = startDate ? formatTime(startDate) : null;
      const startEpoch = startDate ? Math.floor(startDate.getTime() / 1000) : null;
      
      const endDateFormatted = formatDate(endDate);
      const endTimeFormatted = formatTime(endDate);
      const endEpoch = Math.floor(endDate.getTime() / 1000);
      
      const endSerial = dateToSerial(endDate);
      const archive = `${endDate.getFullYear()}-${String(endDate.getMonth() + 1).padStart(2, '0')}`;
      
      // ===== GAME DETAILS =====
      const rules = (game.rules || 'chess').toLowerCase();
      const isLive = gameType === 'live';      
      let timeClass = (game.time_class || '').toLowerCase();
      if (!timeClass || timeClass === 'unknown') {
        timeClass = calculateTimeClass(game.time_control);
      }
      const format = getGameFormat(game).toLowerCase();
      const rated = game.rated || false;
      
      // ===== TIME CONTROL =====
      const tcParsed = parseTimeControl(game.time_control, game.time_class);
      const baseTime = tcParsed.baseTime;
      const increment = tcParsed.increment;
      const corrTime = tcParsed.correspondenceTime;
      const timeControlLabel = formatTimeControlLabel(baseTime, increment, corrTime);
      
      const durationFormatted = formatDuration(duration);
      const durationSeconds = duration;
      
      // ===== PLAYER INFO =====
      const isWhite = game.white?.username.toLowerCase() === username.toLowerCase();
      const color = isWhite ? 'white' : 'black';
      const opponent = (isWhite ? game.black?.username : game.white?.username || '').toLowerCase();
      const myRating = isWhite ? game.white?.rating : game.black?.rating;
      const oppRating = isWhite ? game.black?.rating : game.white?.rating;
      
      // ===== RATING CALCULATIONS =====
      const ratingBefore = currentLedger[format] || null;
      const ratingAfter = myRating || null;
      const ratingDelta = (ratingBefore !== null && ratingAfter !== null) ? (ratingAfter - ratingBefore) : null;
      
      // Update ledger for next game
      if (ratingAfter !== null) {
        currentLedger[format] = ratingAfter;
      }
      
      // ===== GAME RESULT =====
      const outcome = getGameOutcome(game, username).toLowerCase();
      const termination = getGameTermination(game, username).toLowerCase();
      
      // ===== OPENING INFO =====
      const ecoCode = extractECOCodeFromPGN(game.pgn) || '';
      const ecoUrl = extractECOFromPGN(game.pgn) || '';
      const openingData = getOpeningDataForGame(ecoUrl);
      
      // ===== MOVE DATA =====
      const moveData = extractMovesWithClocks(game.pgn, baseTime, increment);
      const movesCount = moveData.plyCount > 0 ? Math.ceil(moveData.plyCount / 2) : 0;
      const tcn = game.tcn || '';
      const clocks = encodeClocksBase36(moveData.clocks);
      
      // ===== LEDGER =====
      const ledgerString = JSON.stringify(currentLedger);
      
      // ===== BUILD ROW =====
      rows.push([
        gameId,                    // A: Game ID
        gameType,                  // B: Type
        gameUrl,                   // C: Game URL
        startDateTimeFormatted,    // D: Start Date/Time
        startDateFormatted,        // E: Start Date
        startTimeFormatted,        // F: Start Time
        startEpoch,                // G: Start (s)
        endDateTimeFormatted,      // H: End Date/Time
        endDateFormatted,          // I: End Date
        endTimeFormatted,          // J: End Time
        endEpoch,                  // K: End (s)
        endSerial,                 // L: End Serial
        archive,                   // M: Archive
        rules,                     // N: Rules
        isLive,                    // O: Live
        timeClass,                 // P: Time Class
        format,                    // Q: Format
        rated,                     // R: Rated
        timeControlLabel,          // S: Time Control
        baseTime,                  // T: Base
        increment,                 // U: Inc
        corrTime,                  // V: Corr
        durationFormatted,         // W: Duration
        durationSeconds,           // X: Duration (s)
        color,                     // Y: Color
        opponent,                  // Z: Opponent
        myRating,                  // AA: My Rating
        oppRating,                 // AB: Opp Rating
        ratingBefore,              // AC: Rating Before
        ratingDelta,               // AD: Rating Δ
        'pending',                 // AE: Callback Status
        outcome,                   // AF: Outcome
        termination,               // AG: Termination
        ecoCode,                   // AH: ECO
        ecoUrl,                    // AI: ECO URL
        openingData[0],            // AJ: Opening Name
        openingData[1],            // AK: Opening Slug
        openingData[2],            // AL: Opening Family
        openingData[3],            // AM: Opening Base
        openingData[4],            // AN: Variation 1
        openingData[5],            // AO: Variation 2
        openingData[6],            // AP: Variation 3
        openingData[7],            // AQ: Variation 4
        openingData[8],            // AR: Variation 5
        openingData[9],            // AS: Variation 6
        openingData[10],           // AT: Extra Moves
        movesCount,                // AU: Moves
        tcn,                       // AV: TCN
        clocks,                    // AW: Clocks
        ledgerString               // AX: Ratings Ledger
      ]);
      
    } catch (error) {
      Logger.log(`Error processing game ${game?.url}: ${error.message}`);
      continue;
    }
  }
  
  return rows;
}

// ===== SETUP =====
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES) || ss.insertSheet(SHEETS.GAMES);
  
  // DON'T clear existing content - only clear header row
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  }
  
  // Set headers
  const headers = [
    'Game ID', 'Type', 'Game URL',
    'Start Date/Time', 'Start Date', 'Start Time', 'Start (s)',
    'End Date/Time', 'End Date', 'End Time', 'End (s)', 'End Serial', 'Archive',
    'Rules', 'Live', 'Time Class', 'Format', 'Rated', 'Time Control', 'Base', 'Inc', 'Corr', 'Duration', 'Duration (s)',
    'Color', 'Opponent', 'My Rating', 'Opp Rating', 'Rating Before', 'Rating Δ', 'Callback Status',
    'Outcome', 'Termination',
    'ECO', 'ECO URL',
    'Opening Name', 'Opening Slug', 'Opening Family', 'Opening Base',
    'Variation 1', 'Variation 2', 'Variation 3', 'Variation 4', 'Variation 5', 'Variation 6',
    'Extra Moves',
    'Moves', 'TCN', 'Clocks', 'Ratings Ledger'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  // Format combined datetime columns (E, I)
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('H:H').setNumberFormat('@');
  
  // Format date columns (F, J)
  sheet.getRange('E:E').setNumberFormat('M/D/YY');
  sheet.getRange('I:I').setNumberFormat('M/D/YY');
  
  // Format time columns (G, K)
  sheet.getRange('F:F').setNumberFormat('h:mm AM/PM');
  sheet.getRange('J:J').setNumberFormat('h:mm AM/PM');
  
  // Format duration column (X)
  sheet.getRange('W:W').setNumberFormat('[h]:mm:ss');
  
  // Format duration seconds column (Y)
  sheet.getRange('X:X').setNumberFormat('0');
  
  // Format rating columns as numbers (AB:AE)
  sheet.getRange('AA:AD').setNumberFormat('0');
  
  // Format moves column (AU)
  sheet.getRange('AT:AT').setNumberFormat('0');
  
  // Archive column (N)
  sheet.getRange('M:M').setNumberFormat('@');
  
  // ECO code column (AH)
  sheet.getRange('AG:AG').setNumberFormat('@');
  
  // ECO URL column (AI)
  sheet.getRange('AH:AH').setNumberFormat('@');
  
  // Extra Moves column (AS)
  sheet.getRange('AR:AR').setNumberFormat('@');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Column widths
  sheet.setColumnWidth(1, 90);   // A: Game ID
  sheet.setColumnWidth(2, 60);   // B: Type
  sheet.setColumnWidth(3, 250);  // C: Game URL (will be hidden)
  sheet.setColumnWidth(4, 180);  // E: Start Date/Time
  sheet.setColumnWidths(5, 2, 90); // F-G: Start Date/Time (separate)
  sheet.setColumnWidth(8, 180);  // I: End Date/Time
  sheet.setColumnWidths(9, 2, 90); // J-K: End Date/Time (separate)
  sheet.setColumnWidth(13, 90);  // N: Archive
  sheet.setColumnWidth(14, 100); // O: Rules
  sheet.setColumnWidth(17, 80);  // R: Format
  sheet.setColumnWidth(18, 60);  // S: Rated
  sheet.setColumnWidth(19, 100); // T: Time Control
  sheet.setColumnWidth(23, 90);  // X: Duration
  sheet.setColumnWidth(24, 90);  // Y: Duration (s)
  sheet.setColumnWidth(32, 125); // AH: ECO (code)
  sheet.setColumnWidth(33, 65);  // AI: ECO URL
  sheet.setColumnWidth(34, 90); // AJ: Opening Name
  sheet.setColumnWidth(35, 150); // AK: Opening Slug
  sheet.setColumnWidth(36, 150); // AL: Opening Family
  sheet.setColumnWidth(37, 200); // AM: Opening Base
  sheet.setColumnWidth(38, 120); // AN: Variation 1
  sheet.setColumnWidth(39, 120); // AO: Variation 2
  sheet.setColumnWidth(40, 120); // AP: Variation 3
  sheet.setColumnWidth(41, 120); // AQ: Variation 4
  sheet.setColumnWidth(42, 120); // AR: Variation 5
  sheet.setColumnWidth(43, 120); // AS: Variation 6
  sheet.setColumnWidth(44, 200); // AT: Extra Moves

  // === FONT AND ALIGNMENT ===
  const maxRows = sheet.getMaxRows();
  const maxCols = headers.length;
  sheet.getRange(1, 1, maxRows, maxCols).setFontFamily('Montserrat');
  sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment('center');

  // === REMOVE GRIDLINES ===
  sheet.setHiddenGridlines(true);

  // === ALTERNATING ROW COLORS (BANDING) ===
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  const banding = dataRange.getBandings()[0];
  if (banding) {
    banding.remove();
  }
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // === BORDERS ===
  sheet.getRange(1, 1, 1, headers.length).setBorder(null, null, true, null, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Remove all existing column groups
  try {
    const lastCol = sheet.getMaxColumns();
    if (lastCol > 0) {
      sheet.getRange(1, 1, 1, lastCol).shiftColumnGroupDepth(-1);
    }
  } catch (e) {
    // No groups exist, continue
  }

  // Create fresh groups
  sheet.getRange('A1:B1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('A:B'));

  sheet.getRange('D1:H1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('D:H'));

  sheet.getRange('K1:P1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('K:P'));

  sheet.getRange('T1:W1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('T:W'));

  sheet.getRange('AL1:AS1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('AL:AS'));

  sheet.getRange('AI1:AJ1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('AI:AJ'));

  sheet.getRange('AU1:AV1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('AU:AV'));

  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  // === NAMED RANGES ===
  ss.setNamedRange('GamesData', sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length));
  ss.setNamedRange('GameIDs', sheet.getRange('A2:A'));
  ss.setNamedRange('Outcomes', sheet.getRange('AE2:AE'));
  ss.setNamedRange('MyRatings', sheet.getRange('AA2:AA'));
  ss.setNamedRange('Opponents', sheet.getRange('Z2:Z'));
  ss.setNamedRange('OpeningNames', sheet.getRange('AI2:AI'));

  // === CONDITIONAL FORMATTING ===
  sheet.clearConditionalFormatRules();
  const newRules = [];

  // Outcome: Win (green)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('win')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  // Outcome: Loss (red)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('loss')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  // Outcome: Draw (yellow)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('draw')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  // Rating Delta: Positive (green text)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  // Rating Delta: Negative (red text)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  // Callback Status: Override (blue background)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('callback_override')
    .setBackground('#cfe2f3')
    .setBold(true)
    .setRanges([sheet.getRange('AF2:AF')])
    .build());

  // Callback Status: Fetched (green background)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('fetched')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('AF2:AF')])
    .build());

  // Callback Status: Fetched Zero (yellow background)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('fetched_zero')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('AF2:AF')])
    .build());

  // Callback Status: No Data (red background)
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('no_data')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('AF2:AF')])
    .build());

  // Time Class colors
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('bullet')
    .setBackground('#c9daf8')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('blitz')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('rapid')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('daily')
    .setBackground('#d9d2e9')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  sheet.setConditionalFormatRules(newRules);

  SpreadsheetApp.getUi().alert('✅ Sheet setup complete with enhanced formatting!');
}

// ===== GAME OUTCOME HELPERS =====
const RESULT_MAP = {
  'win': 'Win',
  'checkmated': 'Loss',
  'agreed': 'Draw',
  'repetition': 'Draw',
  'timeout': 'Loss',
  'resigned': 'Loss',
  'stalemate': 'Draw',
  'lose': 'Loss',
  'insufficient': 'Draw',
  '50move': 'Draw',
  'abandoned': 'Loss',
  'kingofthehill': 'Loss',
  'threecheck': 'Loss',
  'timevsinsufficient': 'Draw',
  'bughousepartnerlose': 'Loss'
};

function getGameOutcome(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white.username?.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  
  if (!myResult) return 'Unknown';
  
  return RESULT_MAP[myResult] || 'Unknown';
}

function getGameTermination(game, username) {
  if (!game || !game.white || !game.black) return 'Unknown';
  
  const isWhite = game.white.username?.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  const opponentResult = isWhite ? game.black.result : game.white.result;
  
  if (!myResult) return 'Unknown';
  
  // If I won, use opponent's result for termination
  if (myResult === 'win') {
    return opponentResult;
  }
  
  // Otherwise use my result
  return myResult;
}

// ===== MOVE EXTRACTION =====
function extractMovesWithClocks(pgn, baseTime, increment) {
  if (!pgn) return { moves: [], clocks: [], times: [] };
  
  const moveSection = pgn.split(/\n\n/)[1] || pgn;
  const moves = [];
  const clocks = [];
  const times = [];
  
  // Regex to match move and its clock: "e4 {[%clk 0:02:59.9]}"
  const movePattern = /([NBRQK]?[a-h]?[1-8]?x?[a-h][1-8](?:=[NBRQK])?|O-O(?:-O)?)\s*\{?\[%clk\s+(\d+):(\d+):(\d+)(?:\.(\d+))?\]?\}?/g;
  
  let match;
  let prevClock = [baseTime || 0, baseTime || 0]; // [white, black] previous clocks
  let moveIndex = 0;
  
  while ((match = movePattern.exec(moveSection)) !== null) {
    const move = match[1];
    const hours = parseInt(match[2]) || 0;
    const minutes = parseInt(match[3]) || 0;
    const seconds = parseInt(match[4]) || 0;
    const deciseconds = parseInt(match[5]) || 0;
    
    // Convert clock to total seconds
    const clockSeconds = hours * 3600 + minutes * 60 + seconds + deciseconds / 10;
    
    moves.push(move);
    clocks.push(clockSeconds);
    
    // Calculate time spent on this move
    const playerIndex = moveIndex % 2; // 0 = white, 1 = black
    const prevPlayerClock = prevClock[playerIndex];
    
    // Time spent = previous clock - current clock + increment
    let timeSpent = prevPlayerClock - clockSeconds + (increment || 0);
    // Allow 0.0 seconds moves (e.g., premove)
    if (timeSpent < 0) timeSpent = 0;
    
    times.push(Math.round(timeSpent * 10) / 10); // Round to 1 decimal
    
    // Update previous clock for this player
    prevClock[playerIndex] = clockSeconds;
    
    moveIndex++;
  }
  
  return { 
    moveList: moves.join(', '), 
    clocks: clocks.join(', '), 
    times: times.join(', '),
    plyCount: moves.length
  };
}

// ===== TIME CONTROL PARSING =====
function parseTimeControl(timeControl, timeClass) {
  const result = {
    type: timeClass === 'daily' ? 'Daily' : 'Live',
    baseTime: null,
    increment: null,
    correspondenceTime: null
  };
  
  if (!timeControl) return result;
  
  const tcStr = String(timeControl);
  
  // Check if correspondence/daily format (1/value)
  if (tcStr.includes('/')) {
    const parts = tcStr.split('/');
    if (parts.length === 2) {
      result.correspondenceTime = parseInt(parts[1]) || null;
    }
  }
  // Check if live format with increment (value+value)
  else if (tcStr.includes('+')) {
    const parts = tcStr.split('+');
    if (parts.length === 2) {
      result.baseTime = parseInt(parts[0]) || null;
      result.increment = parseInt(parts[1]) || null;
    }
  }
  // Simple live format (just value)
  else {
    result.baseTime = parseInt(tcStr) || null;
    result.increment = 0;
  }
  
  return result;
}

function getGameFormat(game) {
  const rules = game.rules || 'chess';
  let timeClass = game.time_class || '';
  
  if (rules === 'chess') {
    // Use time class for standard chess (Bullet, Blitz, Rapid, Daily)
    return timeClass.toLowerCase();
  } else if (rules === 'chess960') {
    return timeClass === 'daily' ? 'daily960' : 'live960';
  } else {
    // For other variants, return the rules name
    return rules;
  }
}

function calculateTimeClass(timeControl) {
  if (!timeControl) return 'unknown';
  
  const match = timeControl.match(/(\d+)\+(\d+)/);
  if (!match) return 'unknown';
  
  const base = parseInt(match[1]);
  const inc = parseInt(match[2]);
  const estimated = base + 40 * inc;
  
  if (estimated < 180) return 'bullet';
  if (estimated < 600) return 'blitz';
  return 'rapid';
}

// ===== DURATION EXTRACTION =====
function extractDurationFromPGN(pgn) {
  if (!pgn) return null;
  
  const dateMatch = pgn.match(/\[UTCDate "([^"]+)"\]/);
  const timeMatch = pgn.match(/\[UTCTime "([^"]+)"\]/);
  const endDateMatch = pgn.match(/\[EndDate "([^"]+)"\]/);
  const endTimeMatch = pgn.match(/\[EndTime "([^"]+)"\]/);
  
  if (!dateMatch || !timeMatch || !endDateMatch || !endTimeMatch) {
    return null;
  }
  
  try {
    const startDateParts = dateMatch[1].split('.');
    const startTimeParts = timeMatch[1].split(':');
    const startDate = new Date(Date.UTC(
      parseInt(startDateParts[0]),
      parseInt(startDateParts[1]) - 1,
      parseInt(startDateParts[2]),
      parseInt(startTimeParts[0]),
      parseInt(startTimeParts[1]),
      parseInt(startTimeParts[2])
    ));
    
    const endDateParts = endDateMatch[1].split('.');
    const endTimeParts = endTimeMatch[1].split(':');
    const endDate = new Date(Date.UTC(
      parseInt(endDateParts[0]),
      parseInt(endDateParts[1]) - 1,
      parseInt(endDateParts[2]),
      parseInt(endTimeParts[0]),
      parseInt(endTimeParts[1]),
      parseInt(endTimeParts[2])
    ));
    
    const durationMs = endDate.getTime() - startDate.getTime();
    return Math.round(durationMs / 1000);
  } catch (error) {
    Logger.log(`Error parsing duration: ${error.message}`);
    return null;
  }
}

function extractStartFromPGN(pgn) {
  if (!pgn) return null;
  
  const dateMatch = pgn.match(/\[UTCDate "([^"]+)"\]/);
  const timeMatch = pgn.match(/\[UTCTime "([^"]+)"\]/);
  
  if (!dateMatch || !timeMatch) return null;
  
  try {
    // Parse "2009.10.19" and "14:52:57"
    const d = dateMatch[1].split('.');
    const t = timeMatch[1].split(':');
    
    return new Date(Date.UTC(
      parseInt(d[0]),      // year
      parseInt(d[1]) - 1,  // month (0-indexed)
      parseInt(d[2]),      // day
      parseInt(t[0]),      // hour
      parseInt(t[1]),      // minute
      parseInt(t[2])       // second
    ));
  } catch (e) {
    Logger.log(`Error parsing PGN date/time: ${e.message}`);
    return null;
  }
}

// ===== ECO AND OPENING EXTRACTION =====
function extractECOCodeFromPGN(pgn) {
  if (!pgn) return '';
  
  // Look for [ECO "B08"] pattern
  const ecoMatch = pgn.match(/\[ECO\s+"([A-E]\d{2})"\]/i);
  if (ecoMatch && ecoMatch[1]) {
    return ecoMatch[1].toUpperCase();
  }
  
  return '';
}

function extractECOFromPGN(pgn) {
  if (!pgn) return '';
  
  // Chess.com includes [ECOUrl "..."] in their PGNs
  const ecoUrlMatch = pgn.match(/\[ECOUrl\s+"([^"]+)"\]/i);
  if (ecoUrlMatch && ecoUrlMatch[1]) {
    return ecoUrlMatch[1];
  }
  
  // Fallback: try to find [Link "...openings/..."]
  const linkMatch = pgn.match(/\[Link\s+"([^"]*openings\/[^"]+)"\]/i);
  if (linkMatch && linkMatch[1]) {
    return linkMatch[1];
  }
  
  return '';
}

// ===== ENHANCED OPENING DATA WITH FAMILY EXTRACTION =====
function getOpeningDataForGame(ecoUrl) {
  if (!ecoUrl) return ['', '', '', '', '', '', '', '', '', '', ''];
  
  // Extract opening name from URL
  const match = ecoUrl.match(/\/openings\/([^"]+)$/);
  if (!match) return ['', '', '', '', '', '', '', '', '', '', ''];
  
  const slug = match[1];
  const openingName = slug
    .split('-')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
  
  // Extract family from opening name (first major part)
  let openingFamily = '';
  if (openingName) {
    const familyParts = openingName.split(' ');
    if (familyParts.length >= 2) {
      // Take first 2-3 words as family (e.g., "Sicilian Defense", "King's Indian")
      openingFamily = familyParts.slice(0, Math.min(3, familyParts.length)).join(' ');
    } else {
      openingFamily = openingName;
    }
  }
  
  return [
    openingName,     // Opening Name
    slug,            // Opening Slug
    openingFamily,   // Opening Family
    openingName,     // Opening Base (same as name for now)
    '',              // Variation 1
    '',              // Variation 2
    '',              // Variation 3
    '',              // Variation 4
    '',              // Variation 5
    '',              // Variation 6
    ''               // Extra Moves
  ];
}

// ===== CLOCK ENCODING =====
function encodeClocksBase36(clocksCsv) {
  if (!clocksCsv) return '';
  const parts = String(clocksCsv).split(',').map(s => s.trim()).filter(Boolean);
  if (parts.length === 0) return '';
  return parts.map(p => {
    const ds = Math.round(parseFloat(p) * 10);
    const val = isFinite(ds) && ds >= 0 ? ds : 0;
    return val.toString(36);
  }).join('.');
}

// ===== TIME CONTROL LABEL =====
function formatTimeControlLabel(baseTime, increment, corrTime) {
  // Daily/correspondence games
  if (corrTime != null) {
    const days = Math.floor(corrTime / 86400);
    return days === 1 ? '1 day' : `${days} days`;
  }
  
  // Live games
  if (baseTime == null) return 'unknown';
  
  const hasIncrement = increment != null && increment > 0;
  
  // Check if base time is evenly divisible by 60 (whole minutes)
  const isWholeMinutes = baseTime % 60 === 0;
  const minutes = baseTime / 60;
  
  if (isWholeMinutes && !hasIncrement) {
    // Format as "X min" (e.g., "1 min", "3 min", "10 min", "60 min")
    return `${minutes} min`;
  } else if (isWholeMinutes && hasIncrement) {
    // Format as "X | inc" without "min" (e.g., "3 | 2", "10 | 5")
    return `${minutes} | ${increment}`;
  } else if (!isWholeMinutes && !hasIncrement) {
    // Format as "X sec" (e.g., "20 sec", "30 sec")
    return `${baseTime} sec`;
  } else {
    // Format as "X sec | inc" (e.g., "20 sec | 1", "45 sec | 2")
    return `${baseTime} sec | ${increment}`;
  }
}

// ===== FORMATTING HELPERS =====
function formatDateTime(date) {
  const datePart = `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
  
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12 || 12;
  
  return `${datePart} ${hours}:${minutes}:${seconds} ${ampm}`;
}

function formatDate(date) {
  return `${date.getMonth() + 1}/${date.getDate()}/${String(date.getFullYear()).slice(-2)}`;
}

function formatTime(date) {
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12 || 12;
  return `${hours}:${minutes}:${seconds} ${ampm}`;
}

function formatDuration(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;
  return `${hours}:${String(minutes).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

function dateToSerial(date) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const epoch = new Date(1899, 11, 30);
  const localDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.floor((localDate.getTime() - epoch.getTime()) / msPerDay);
}

// ===== SUMMARY STATS (FROM OLD.GS) =====
function updateSummaryStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName('Games');
  
  if (!gamesSheet || gamesSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert('❌ No games data found!');
    return;
  }
  
  ss.toast('Calculating summary statistics...', '⏳', -1);
  
  // Define which formats to include (in this specific order)
  const includedFormats = ['bullet', 'blitz', 'rapid'];
  
  // Get or create Summary sheet
  let summarySheet = ss.getSheetByName('Summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('Summary');
    setupSummarySheet(summarySheet);
  }
  
  // Get all game data
  const data = gamesSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find column indices
  const colIndices = {
    endDate: headers.indexOf('End Date'),
    format: headers.indexOf('Format'),
    outcome: headers.indexOf('Outcome'),
    duration: headers.indexOf('Duration (s)'),
    ratingDelta: headers.indexOf('Rating Δ'),
    ledger: headers.indexOf('Ratings Ledger')
  };
  
  // Build summary data structure
  const summaryMap = new Map(); // key: "date_format"
  const allDates = new Set();
  
  // FIXED: Collect ALL dates from ALL games (not just included formats)
  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][colIndices.endDate];
    if (dateVal) {
      const date = new Date(dateVal);
      const dateKey = formatDateKey(date);
      allDates.add(dateKey);
    }
  }
  
  // Second pass: process only included formats for stats
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[colIndices.endDate];
    const format = row[colIndices.format];
    const outcome = row[colIndices.outcome];
    const duration = row[colIndices.duration] || 0;
    const ratingDelta = row[colIndices.ratingDelta] || 0;
    const ledgerStr = row[colIndices.ledger];
    
    if (!dateVal || !format) continue;
    
    // Only process included formats for stats
    if (!includedFormats.includes(format)) continue;
    
    // Convert date to simple date string
    const date = new Date(dateVal);
    const dateKey = formatDateKey(date);
    
    const key = `${dateKey}_${format}`;
    
    if (!summaryMap.has(key)) {
      summaryMap.set(key, {
        date: dateKey,
        format: format,
        wins: 0,
        losses: 0,
        draws: 0,
        totalGames: 0,
        durationSum: 0,
        ratingDeltaSum: 0,
        lastRating: null
      });
    }
    
    const entry = summaryMap.get(key);
    entry.totalGames++;
    entry.durationSum += duration;
    entry.ratingDeltaSum += ratingDelta;
    
    if (outcome === 'win') entry.wins++;
    else if (outcome === 'loss') entry.losses++;
    else if (outcome === 'draw') entry.draws++;
    
    // Parse ledger for last rating of this format
    try {
      if (ledgerStr) {
        const ledger = JSON.parse(ledgerStr);
        if (ledger[format]) {
          entry.lastRating = ledger[format];
        }
      }
    } catch (e) {
      // Skip bad ledger
    }
  }
  
  // Generate complete date range from ALL games
  const sortedDates = Array.from(allDates).sort((a, b) => {
    return parseDateKey(a).getTime() - parseDateKey(b).getTime();
  });
  
  if (sortedDates.length === 0) {
    ss.toast('No valid dates found', '❌', 3);
    return;
  }
  
  const minDate = parseDateKey(sortedDates[0]);
  const maxDate = parseDateKey(sortedDates[sortedDates.length - 1]);
  const allDatesInRange = generateDateRange(minDate, maxDate);
  
  // Build rating tracker per format (forward fill) - only for included formats
  const ratingTracker = {};
  for (const format of includedFormats) {
    ratingTracker[format] = null;
  }
  
  // Build complete grid with all date-format combinations
  const rows = [];
  for (const dateKey of allDatesInRange) {
    // Track totals for this date
    let dateTotalGames = 0;
    let dateTotalWins = 0;
    let dateTotalLosses = 0;
    let dateTotalDraws = 0;
    let dateTotalDuration = 0;
    let dateTotalRatingDelta = 0;
    let dateTotalRating = 0;
    let dateRatingCount = 0;
    
    // Add rows for each format (in specified order)
    for (const format of includedFormats) {
      const key = `${dateKey}_${format}`;
      const entry = summaryMap.get(key);
      
      if (entry) {
        // Update rating tracker if we have a new rating
        if (entry.lastRating !== null) {
          ratingTracker[format] = entry.lastRating;
        }
        
        rows.push([
          entry.date,
          entry.format,
          entry.totalGames,
          entry.wins,
          entry.losses,
          entry.draws,
          ratingTracker[format],  // Use tracked rating (forward filled)
          entry.durationSum,
          entry.ratingDeltaSum
        ]);
        
        dateTotalGames += entry.totalGames;
        dateTotalWins += entry.wins;
        dateTotalLosses += entry.losses;
        dateTotalDraws += entry.draws;
        dateTotalDuration += entry.durationSum;
        dateTotalRatingDelta += entry.ratingDeltaSum;
        
        // Sum ratings for total
        if (ratingTracker[format] !== null) {
          dateTotalRating += ratingTracker[format];
          dateRatingCount++;
        }
      } else {
        // Empty row for date-format with no games - use forward filled rating
        rows.push([
          dateKey,
          format,
          0,      // totalGames
          0,      // wins
          0,      // losses
          0,      // draws
          ratingTracker[format],  // Forward filled rating
          0,      // durationSum
          0       // ratingDeltaSum
        ]);
        
        // Sum ratings for total (even if no games on this date)
        if (ratingTracker[format] !== null) {
          dateTotalRating += ratingTracker[format];
          dateRatingCount++;
        }
      }
    }
    
    // Add TOTAL row for this date
    rows.push([
      dateKey,
      'TOTAL',
      dateTotalGames,
      dateTotalWins,
      dateTotalLosses,
      dateTotalDraws,
      dateRatingCount > 0 ? dateTotalRating : null,  // Sum of all format ratings
      dateTotalDuration,
      dateTotalRatingDelta
    ]);
  }
  
  // Write to sheet
  if (rows.length > 0) {
    const dataRange = summarySheet.getRange(2, 1, rows.length, 9);
    dataRange.setValues(rows);
    
    // Clear any extra rows below
    const lastRow = summarySheet.getLastRow();
    if (lastRow > rows.length + 1) {
      summarySheet.getRange(rows.length + 2, 1, lastRow - rows.length - 1, 9).clearContent();
    }
  }
  
  ss.toast(`✅ Summary updated: ${rows.length} rows`, '✅', 3);
}

// ===== SETUP SUMMARY SHEET =====
function setupSummarySheet(sheet) {
  const headers = [
    'Date',
    'Format',
    'Games',
    'Wins',
    'Losses',
    'Draws',
    'Rating',
    'Duration (s)',
    'Rating Δ'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  // Format columns
  sheet.getRange('A:A').setNumberFormat('M/D/YY');  // Date
  sheet.getRange('B:B').setNumberFormat('@');        // Format
  sheet.getRange('C:I').setNumberFormat('0');        // Numbers
  
  // Column widths
  sheet.setColumnWidth(1, 90);   // Date
  sheet.setColumnWidth(2, 100);  // Format
  sheet.setColumnWidth(3, 70);   // Games
  sheet.setColumnWidth(4, 70);   // Wins
  sheet.setColumnWidth(5, 70);   // Losses
  sheet.setColumnWidth(6, 70);   // Draws
  sheet.setColumnWidth(7, 80);   // Rating
  sheet.setColumnWidth(8, 100);  // Duration
  sheet.setColumnWidth(9, 90);   // Rating Δ
  
  // Styling
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setFontFamily('Montserrat');
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setHorizontalAlignment('center');
  
  // Border under header
  sheet.getRange(1, 1, 1, headers.length).setBorder(null, null, true, null, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Alternating rows
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  // Add filter
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).createFilter();
  
  // Conditional formatting for wins/losses
  sheet.clearConditionalFormatRules();
  const rules = [];
  
  // Wins > 0 (green)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('D2:D')])
    .build());
  
  // Losses > 0 (red)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('E2:E')])
    .build());
  
  // Draws > 0 (yellow)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('F2:F')])
    .build());
  
  // Rating Δ positive (green text)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('I2:I')])
    .build());
  
  // Rating Δ negative (red text)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('I2:I')])
    .build());
  
  // TOTAL rows - bold and light blue background
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TOTAL')
    .setBackground('#cfe2f3')
    .setBold(true)
    .setRanges([sheet.getRange('B2:B')])
    .build());
  
  sheet.setConditionalFormatRules(rules);
}

// ===== HELPER FUNCTIONS =====
function formatDateKey(date) {
  // Returns M/D/YY format for consistency
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const year = date.getFullYear() % 100;
  return `${month}/${day}/${String(year).padStart(2, '0')}`;
}

function parseDateKey(dateKey) {
  // Parse "M/D/YY" back to Date object
  const parts = dateKey.split('/');
  const month = parseInt(parts[0]) - 1;
  const day = parseInt(parts[1]);
  const year = 2000 + parseInt(parts[2]);
  return new Date(year, month, day);
}

function generateDateRange(startDate, endDate) {
  const dates = [];
  const current = new Date(startDate);
  
  while (current <= endDate) {
    dates.push(formatDateKey(current));
    current.setDate(current.getDate() + 1);
  }
  
  return dates;
}

// ===== CALLBACK INTEGRATION SYSTEM =====
// Fetches callback data and overrides ratings when non-zero and different

function enrichNewGamesWithCallbacks(count = 20) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName('Games');
  
  const lastRow = gamesSheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('No games to enrich', 'ℹ️', 3);
    return;
  }
  
  // Get the most recent games
  const startRow = Math.max(2, lastRow - count + 1);
  const numRows = lastRow - startRow + 1;
  
  const gameData = gamesSheet.getRange(startRow, 1, numRows, 53).getValues();
  
  ss.toast(`Enriching ${numRows} recent games with callbacks...`, '⏳', -1);
  
  let successCount = 0;
  let overrideCount = 0;
  let errorCount = 0;
  
  for (let i = 0; i < gameData.length; i++) {
    const row = gameData[i];
    const gameId = row[0]; // Game ID
    const gameUrl = row[2]; // Game URL
    const timeClass = row[15]; // Time Class
    const currentRatingBefore = row[28]; // Rating Before
    const currentRatingDelta = row[29]; // Rating Delta
    
    try {
      // Fetch callback data
      const callbackData = fetchCallbackData({
        gameId: gameId,
        gameUrl: gameUrl,
        timeClass: timeClass
      });
      
      if (callbackData && callbackData.myRatingBefore !== null && callbackData.myRatingChange !== null) {
        const callbackRatingBefore = callbackData.myRatingBefore;
        const callbackRatingChange = callbackData.myRatingChange;
        
        // Store comprehensive callback data
        storeCallbackData(gameId, callbackData);
        
        // Check if callback data is non-zero and different from current data
        const isDifferent = (callbackRatingBefore !== currentRatingBefore) || (callbackRatingChange !== currentRatingDelta);
        const isNonZero = callbackRatingChange !== 0;
        
        if (isDifferent && isNonZero) {
          // Override with callback data
          const actualRow = startRow + i;
          gamesSheet.getRange(actualRow, 29).setValue(callbackRatingBefore); // Rating Before
          gamesSheet.getRange(actualRow, 30).setValue(callbackRatingChange); // Rating Delta
          gamesSheet.getRange(actualRow, 31).setValue('callback_override'); // Mark as overridden
          
          overrideCount++;
          Logger.log(`Overrode ratings for game ${gameId}: ${currentRatingBefore}→${callbackRatingBefore}, ${currentRatingDelta}→${callbackRatingChange}`);
        }
        
        successCount++;
      }
      
      // Rate limiting
      Utilities.sleep(500);
      
    } catch (error) {
      const errorLog = {
        timestamp: new Date().toISOString(),
        gameId: gameId,
        error: {
          message: error.message,
          stack: error.stack
        },
        context: {
          gameUrl: gameUrl,
          timeClass: timeClass
        }
      };
      
      Logger.log(`\n=== CALLBACK ERROR FOR GAME ${gameId} ===`);
      Logger.log(JSON.stringify(errorLog, null, 2));
      Logger.log(`=== END CALLBACK ERROR ===\n`);
      
      errorCount++;
    }
  }
  
  const statusMsg = `✅ Callback enrichment complete!\n\n` +
    `Success: ${successCount}\n` +
    `Rating Overrides: ${overrideCount}\n` +
    `Errors: ${errorCount}`;
  
  ss.toast(statusMsg, errorCount > 0 ? '⚠️' : '✅', 8);
  Logger.log(statusMsg);
}

// ===== CALLBACK LOGGING =====
function logCallbackData(gameId, callbackData, context) {
  const logEntry = {
    timestamp: new Date().toISOString(),
    gameId: gameId,
    context: {
      gameUrl: context.gameUrl,
      timeClass: context.timeClass,
      currentRatingBefore: context.currentRatingBefore,
      currentRatingDelta: context.currentRatingDelta
    },
    callbackData: {
      myRating: callbackData.myRating,
      oppRating: callbackData.oppRating,
      myRatingChange: callbackData.myRatingChange,
      oppRatingChange: callbackData.oppRatingChange,
      myRatingBefore: callbackData.myRatingBefore,
      oppRatingBefore: callbackData.oppRatingBefore
    },
    analysis: {
      isDifferent: (callbackData.myRatingBefore !== context.currentRatingBefore) || 
                   (callbackData.myRatingChange !== context.currentRatingDelta),
      isNonZero: callbackData.myRatingChange !== 0,
      willOverride: (callbackData.myRatingBefore !== context.currentRatingBefore) || 
                    (callbackData.myRatingChange !== context.currentRatingDelta) && 
                    callbackData.myRatingChange !== 0
    }
  };
  
  // Log as formatted JSON for easy reading
  Logger.log(`\n=== CALLBACK DATA FOR GAME ${gameId} ===`);
  Logger.log(JSON.stringify(logEntry, null, 2));
  Logger.log(`=== END CALLBACK DATA ===\n`);
  
  // Also log a summary line
  const summary = `Game ${gameId}: My ${callbackData.myRatingBefore}→${callbackData.myRating} (${callbackData.myRatingChange > 0 ? '+' : ''}${callbackData.myRatingChange}), ` +
                  `Opp ${callbackData.oppRatingBefore}→${callbackData.oppRating} (${callbackData.oppRatingChange > 0 ? '+' : ''}${callbackData.oppRatingChange})`;
  Logger.log(summary);
}

// ===== CALLBACK DATA FETCHING =====
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
    
    return {
      gameId: gameId,
      myRating: myRating,
      oppRating: oppRating,
      myRatingChange: myRatingChange,
      oppRatingChange: oppRatingChange,
      myRatingBefore: myRatingBefore,
      oppRatingBefore: oppRatingBefore
    };
    
  } catch (error) {
    Logger.log(`Error fetching callback data for game ${gameId}: ${error.message}`);
    return null;
  }
}

// ===== MANUAL CALLBACK ENRICHMENT =====
function enrichAllPendingCallbacks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName('Games');
  
  const lastRow = gamesSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No games found');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Enrich All Games with Callbacks?',
    `This will fetch callback data for all games and override ratings where different.\n\n` +
    'This may take several minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    enrichNewGamesWithCallbacks(lastRow - 1);
  }
}

// ===== IMMEDIATE CALLBACK ENRICHMENT =====
function enrichRecentGamesImmediate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName('Games');
  
  const lastRow = gamesSheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('No games found', '⚠️', 3);
    return;
  }
  
  // Enrich last 20 games immediately
  enrichNewGamesWithCallbacks(20);
}

// ===== UPDATE PENDING CALLBACKS =====
function updatePendingCallbacks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName('Games');
  
  const lastRow = gamesSheet.getLastRow();
  if (lastRow <= 1) {
    ss.toast('No games found', '⚠️', 3);
    return;
  }
  
  // Find all rows with "pending" status
  const statusColumn = 31; // AF column
  const statusRange = gamesSheet.getRange(2, statusColumn, lastRow - 1, 1);
  const statusValues = statusRange.getValues();
  
  let pendingCount = 0;
  const pendingRows = [];
  
  for (let i = 0; i < statusValues.length; i++) {
    if (statusValues[i][0] === 'pending') {
      pendingCount++;
      pendingRows.push(i + 2); // +2 because we start from row 2
    }
  }
  
  if (pendingCount === 0) {
    ss.toast('No pending callbacks found', 'ℹ️', 3);
    return;
  }
  
  ss.toast(`Found ${pendingCount} pending callbacks. Processing...`, '⏳', 3);
  
  // Process each pending row
  let successCount = 0;
  let overrideCount = 0;
  let errorCount = 0;
  
  for (const rowNum of pendingRows) {
    try {
      const gameData = gamesSheet.getRange(rowNum, 1, 1, 53).getValues()[0];
      const gameId = gameData[0];
      const gameUrl = gameData[2];
      const timeClass = gameData[15];
      const currentRatingBefore = gameData[28];
      const currentRatingDelta = gameData[29];
      
      const callbackData = fetchCallbackData({
        gameId: gameId,
        gameUrl: gameUrl,
        timeClass: timeClass
      });
      
      if (callbackData && callbackData.myRatingBefore !== null && callbackData.myRatingChange !== null) {
        const callbackRatingBefore = callbackData.myRatingBefore;
        const callbackRatingChange = callbackData.myRatingChange;
        
        // Log comprehensive callback data
        logCallbackData(gameId, callbackData, {
          currentRatingBefore,
          currentRatingDelta,
          gameUrl,
          timeClass
        });
        
        // Check if callback data is non-zero and different from current data
        const isDifferent = (callbackRatingBefore !== currentRatingBefore) || (callbackRatingChange !== currentRatingDelta);
        const isNonZero = callbackRatingChange !== 0;
        
        let status = 'fetched';
        
        if (isDifferent && isNonZero) {
          // Override with callback data
          gamesSheet.getRange(rowNum, 29).setValue(callbackRatingBefore); // Rating Before
          gamesSheet.getRange(rowNum, 30).setValue(callbackRatingChange); // Rating Delta
          status = 'callback_override';
          overrideCount++;
          
          Logger.log(`✅ OVERRIDE: Game ${gameId} - Ratings changed from ${currentRatingBefore}→${callbackRatingBefore}, ${currentRatingDelta}→${callbackRatingChange}`);
        } else if (isDifferent && !isNonZero) {
          status = 'fetched_zero';
          Logger.log(`ℹ️ SAME ZERO: Game ${gameId} - Callback data same as sheet (both zero rating change)`);
        } else {
          Logger.log(`ℹ️ SAME DATA: Game ${gameId} - Callback data matches sheet data`);
        }
        
        // Update status
        gamesSheet.getRange(rowNum, 31).setValue(status);
        successCount++;
      } else {
        // No callback data available
        gamesSheet.getRange(rowNum, 31).setValue('no_data');
        Logger.log(`❌ NO DATA: Game ${gameId} - No callback data available`);
      }
      
      // Rate limiting
      Utilities.sleep(500);
      
    } catch (error) {
      Logger.log(`Error processing pending game at row ${rowNum}: ${error.message}`);
      errorCount++;
    }
  }
  
  const statusMsg = `✅ Pending callbacks processed!\n\n` +
    `Success: ${successCount}\n` +
    `Rating Overrides: ${overrideCount}\n` +
    `Errors: ${errorCount}`;
  
  ss.toast(statusMsg, errorCount > 0 ? '⚠️' : '✅', 8);
  Logger.log(statusMsg);
}

// ===== VIEW CALLBACK LOGS =====
function viewCallbackLogs() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Callback Logs',
    'Callback data is logged to the Apps Script console.\n\n' +
    'To view logs:\n' +
    '1. Go to Extensions > Apps Script\n' +
    '2. Click "View" > "Logs"\n' +
    '3. Look for "CALLBACK DATA FOR GAME" entries\n\n' +
    'Each log entry contains:\n' +
    '• Game ID and context\n' +
    '• Current vs Callback ratings\n' +
    '• Analysis of differences\n' +
    '• Override decisions\n\n' +
    'Note: Logs are automatically cleared when you run new operations.',
    ui.ButtonSet.OK
  );
}

// ===== CLEAR CALLBACK LOGS =====
function clearCallbackLogs() {
  // Clear the console logs
  console.clear();
  Logger.log('Callback logs cleared at ' + new Date().toISOString());
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Callback logs cleared', 'ℹ️', 3);
}

// ===== TEST CALLBACK FETCH =====
function testCallbackFetch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName('Games');
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No games found');
    return;
  }
  
  // Get the most recent game
  const gameData = gamesSheet.getRange(lastRow, 1, 1, 53).getValues()[0];
  
  const game = {
    gameId: gameData[0],
    gameUrl: gameData[2],
    timeClass: gameData[15]
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
      `My Rating: ${callbackData.myRatingBefore} → ${callbackData.myRating} (${callbackData.myRatingChange > 0 ? '+' : ''}${callbackData.myRatingChange})\n\n` +
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