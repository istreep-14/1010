// ===== GAME PROCESSING =====
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
        outcome,                   // AE: Outcome
        termination,               // AF: Termination
        ecoCode,                   // AG: ECO
        ecoUrl,                    // AH: ECO URL
        openingData[0],            // AI: Opening Name
        openingData[1],            // AJ: Opening Slug
        openingData[2],            // AK: Opening Family
        openingData[3],            // AL: Opening Base
        openingData[4],            // AM: Variation 1
        openingData[5],            // AN: Variation 2
        openingData[6],            // AO: Variation 3
        openingData[7],            // AP: Variation 4
        openingData[8],            // AQ: Variation 5
        openingData[9],            // AR: Variation 6
        openingData[10],           // AS: Extra Moves
        movesCount,                // AT: Moves
        tcn,                       // AU: TCN
        clocks,                    // AV: Clocks
        ledgerString               // AW: Ratings Ledger
      ]);
      
    } catch (error) {
      Logger.log(`Error processing game ${game?.url}: ${error.message}`);
      continue;
    }
  }
  
  return rows;
}
