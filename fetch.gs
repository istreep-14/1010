// ===== EXPANDED GAME FETCHING WITH ALL DATA =====

function fetchChesscomGames(options = {}) {
  const { monthsToFetch = CONFIG.MONTHS_TO_FETCH, specificArchive = null } = options;
  const ss = getMainSpreadsheet();
  
  try {
    ss.toast('Getting archives list...', '⏳', -1);
    
    // Get archives to fetch (auto-syncs archive list)
    const archives = getArchivesToFetch({ monthsToFetch, specificArchive });
    
    if (!archives.length) {
      ss.toast('No archives to fetch', 'ℹ️', 3);
      return;
    }
    
    ss.toast(`Fetching ${archives.length} archive(s)...`, '⏳', -1);
    
    // Fetch games from archives
    const allGames = fetchGamesFromArchives(archives);
    
    if (!allGames.length) {
      ss.toast('No games found in archives', 'ℹ️', 3);
      return;
    }
    
    // Filter to new games only (using archive anchors when possible)
    const newGames = filterNewGames(allGames);
    
    if (!newGames.length) {
      ss.toast('No new games found', 'ℹ️', 3);
      updateArchiveStatuses(archives, allGames);
      return;
    }
    
    ss.toast(`Processing ${newGames.length} new games...`, '⏳', -1);
    
    // Initialize ratings tracker
    const ratingsTracker = new ImprovedRatingsTracker();
    ratingsTracker.loadFromGamesSheet();
    
    // Process and write games
    const gamesRows = processExpandedGames(newGames, ratingsTracker);
    writeGamesToSheet(gamesRows);
    
    // Update archives with last game IDs
    updateArchiveStatuses(archives, allGames);
    
    // Update property
    setProperty(PROP_KEYS.TOTAL_GAMES, 
      (parseInt(getProperty(PROP_KEYS.TOTAL_GAMES, '0')) + newGames.length).toString()
    );
    
    ss.toast(`✅ Added ${newGames.length} new games!`, '✅', 5);
    
    // Ask about callback enrichment
    if (newGames.length > 0) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Fetch Callbacks?',
        `Would you like to fetch callback data for these ${newGames.length} new games?\n\n` +
        'This will get accurate pre-game ratings and rating changes.',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        enrichNewGamesWithCallbacks(newGames.length);
      }
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
    Logger.log(error);
  }
}

// ===== FETCH GAMES FROM ARCHIVES =====
function fetchGamesFromArchives(archiveUrls) {
  const allGamesByArchive = {};
  
  for (const url of archiveUrls) {
    try {
      updateArchiveStatus(url, { status: 'fetching' });
      
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());
      
      if (data.games) {
        allGamesByArchive[url] = data.games;
        
        const etag = response.getHeaders()['ETag'] || '';
        updateArchiveStatus(url, {
          etag: etag,
          lastChecked: new Date(),
          lastFetched: new Date(),
          gameCount: data.games.length
        });
      }
      
      Utilities.sleep(300);
      
    } catch (e) {
      Logger.log(`Failed to fetch ${url}: ${e.message}`);
      updateArchiveStatus(url, { status: 'error' });
    }
  }
  
  // Flatten and sort by end time
  const allGames = [];
  for (const url in allGamesByArchive) {
    allGamesByArchive[url].forEach(game => {
      game._archiveUrl = url;  // Tag with archive URL
      allGames.push(game);
    });
  }
  
  return allGames.sort((a, b) => a.end_time - b.end_time);
}

// ===== PROCESS EXPANDED GAMES =====
function processExpandedGames(games, ratingsTracker) {
  const rows = [];
  const now = new Date();
  
  for (const game of games) {
    try {
      if (!game || !game.url || !game.end_time) continue;
      
      const gameId = game.url.split('/').pop();
      const gameType = (game.time_class || '').toLowerCase() === 'daily' ? 'daily' : 'live';
      const endDate = new Date(game.end_time * 1000);
      const archive = `${endDate.getFullYear()}-${String(endDate.getMonth() + 1).padStart(2, '0')}`;
      
      // Extract start time and duration from PGN
      const startDate = extractStartFromPGN(game.pgn);
      const duration = extractDurationFromPGN(game.pgn);
      
      // Determine player info
      const isWhite = game.white?.username.toLowerCase() === CONFIG.USERNAME.toLowerCase();
      const color = isWhite ? 'white' : 'black';
      const opponent = (isWhite ? game.black?.username : game.white?.username || '').toLowerCase();
      const myRating = isWhite ? game.white?.rating : game.black?.rating;
      const oppRating = isWhite ? game.black?.rating : game.white?.rating;
      
      // Game details
      const rules = (game.rules || 'chess').toLowerCase();
      const isLive = gameType === 'live';
      const timeClass = (game.time_class || '').toLowerCase();
      const format = getGameFormat(game).toLowerCase();
      const rated = game.rated || false;
      
      // Time control
      const tcParsed = parseTimeControl(game.time_control, game.time_class);
      
      // Calculate ratings using tracker
      const ratings = ratingsTracker.calculateRating(format, myRating);
      
      // Get outcome
      const outcome = getGameOutcome(game, CONFIG.USERNAME).toLowerCase();
      const termination = getGameTermination(game, CONFIG.USERNAME).toLowerCase();
      
      // Extract opening details
      const ecoCode = extractECOCodeFromPGN(game.pgn) || '';
      const ecoUrl = extractECOFromPGN(game.pgn) || '';
      const openingData = getOpeningDataForGame(ecoUrl);
      
      // Move data
      const moveData = extractMovesWithClocks(game.pgn, tcParsed.baseTime, tcParsed.increment);
      const movesCount = moveData.plyCount > 0 ? Math.ceil(moveData.plyCount / 2) : 0;
      
      // Build row with ALL data
      rows.push([
        gameId,
        gameType,
        game.url,
        game.pgn || '',
        
        // Dates & Times
        startDate ? formatDateTime(startDate) : null,
        startDate ? formatDate(startDate) : null,
        startDate ? formatTime(startDate) : null,
        startDate ? Math.floor(startDate.getTime() / 1000) : null,
        formatDateTime(endDate),
        formatDate(endDate),
        formatTime(endDate),
        game.end_time,
        archive,
        
        // Game Details
        rules,
        isLive,
        timeClass,
        format,
        rated,
        game.time_control || '',
        tcParsed.baseTime,
        tcParsed.increment,
        duration ? formatDuration(duration) : null,
        duration || 0,
        
        // Players
        color,
        opponent,
        
        // Ratings
        myRating,
        oppRating,
        ratings.before,
        ratings.delta,
        
        // Result
        outcome,
        termination,
        
        // Opening (11 columns)
        ecoCode,
        ecoUrl,
        ...openingData,
        
        // Move Data
        movesCount,
        game.tcn || '',
        
        // Enrichment Status
        null,  // Callback status
        null,  // Callback date
        null,  // Lichess status
        null,  // Lichess URL
        
        // Metadata
        now,   // Fetch date
        now    // Last updated
      ]);
      
      // Add to registry
      addToRegistry(gameId, archive, format);
      
    } catch (error) {
      Logger.log(`Error processing game ${game?.url}: ${error.message}`);
      continue;
    }
  }
  
  return rows;
}

// ===== WRITE GAMES TO SHEET =====
function writeGamesToSheet(rows) {
  if (!rows.length) return;
  
  const gamesSheet = getGamesSheet();
  const startRow = gamesSheet.getLastRow() + 1;
  
  gamesSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  
  Logger.log(`Wrote ${rows.length} games to sheet`);
}

// ===== UPDATE ARCHIVE STATUSES WITH LAST GAME IDS =====
function updateArchiveStatuses(archives, allGames) {
  const now = new Date();
  
  // Group games by archive
  const gamesByArchive = {};
  for (const game of allGames) {
    const archiveUrl = game._archiveUrl;
    if (archiveUrl) {
      if (!gamesByArchive[archiveUrl]) {
        gamesByArchive[archiveUrl] = [];
      }
      gamesByArchive[archiveUrl].push(game);
    }
  }
  
  for (const archiveUrl of archives) {
    const parts = archiveUrl.split('/');
    const year = parseInt(parts[parts.length - 2]);
    const month = parseInt(parts[parts.length - 1]);
    
    // Archive is complete if it's not current month
    const archiveEndDate = new Date(year, month, 0);
    const isComplete = now > archiveEndDate;
    
    // Get last game ID for this archive
    const archiveGames = gamesByArchive[archiveUrl] || [];
    const lastGameId = archiveGames.length > 0 
      ? archiveGames[archiveGames.length - 1].url.split('/').pop()
      : null;
    
    updateArchiveStatus(archiveUrl, {
      status: isComplete ? 'complete' : 'pending',
      lastChecked: now,
      lastGameId: lastGameId
    });
  }
}

// ===== HELPER: FORMAT DATE/TIME =====
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d/yy');
}

function formatTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'h:mm a');
}

function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d/yy h:mm a');
}

function formatDuration(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;
  return `${hours}:${String(minutes).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

// ===== HELPER: PARSE TIME CONTROL =====
function parseTimeControl(timeControl, timeClass) {
  if (!timeControl) {
    return { baseTime: null, increment: null, correspondenceTime: null };
  }
  
  // Daily/correspondence
  if (timeClass === 'daily') {
    const match = timeControl.match(/(\d+)/);
    return {
      baseTime: null,
      increment: null,
      correspondenceTime: match ? parseInt(match[1]) : null
    };
  }
  
  // Live games (base+increment)
  const match = timeControl.match(/(\d+)\+(\d+)/);
  if (match) {
    return {
      baseTime: parseInt(match[1]),
      increment: parseInt(match[2]),
      correspondenceTime: null
    };
  }
  
  return { baseTime: null, increment: null, correspondenceTime: null };
}
