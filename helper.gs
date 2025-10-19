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

// Extract moves with clock times from PGN
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

// Helpers for base-36 sequences in deciseconds
function decodeBase36Seq(s) { return String(s).split('.').filter(Boolean).map(t => { const v = parseInt(t, 36); return isFinite(v) && v >= 0 ? v : 0; }); }
function encodeBase36Seq(arr) { return (arr || []).map(v => (v >= 0 ? v : 0).toString(36)).join('.'); }
function reconstructTimesFromClocksDeci(baseDeci, incDeci, clocksDeci) {
  const times = [];
  let prev = [baseDeci || 0, baseDeci || 0];
  for (let i = 0; i < clocksDeci.length; i++) {
    const p = i % 2;
    const t = (prev[p] - (clocksDeci[i] || 0) + (incDeci || 0));
    times.push(t >= 0 ? t : 0);
    prev[p] = clocksDeci[i] || 0;
  }
  return times;
}

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

function formatRatingDelta(delta) {
  if (delta > 0) return `+${delta}`;
  if (delta < 0) return `${delta}`;
  return '0';
}

function dateToSerial(date) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const epoch = new Date(1899, 11, 30);
  const localDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.floor((localDate.getTime() - epoch.getTime()) / msPerDay);
}

function extractECOCodeFromPGN(pgn) {
  if (!pgn) return '';
  
  // Look for [ECO "B08"] pattern
  const ecoMatch = pgn.match(/\[ECO\s+"([A-E]\d{2})"\]/i);
  if (ecoMatch && ecoMatch[1]) {
    return ecoMatch[1].toUpperCase();
  }
  
  return '';
}

function testECOCodeExtraction() {
  const username = CONFIG.USERNAME;
  const archiveUrl = `https://api.chess.com/pub/player/${username}/games/2025/10`;
  
  try {
    const response = UrlFetchApp.fetch(archiveUrl);
    const data = JSON.parse(response.getContentText());
    
    if (data.games && data.games.length > 0) {
      const game = data.games[0];
      
      const ecoCode = extractECOCodeFromPGN(game.pgn);
      const ecoUrl = extractECOFromPGN(game.pgn);
      
      Logger.log('=== ECO Extraction Test ===');
      Logger.log('Game: ' + game.url);
      Logger.log('ECO Code: ' + ecoCode);
      Logger.log('ECO URL: ' + ecoUrl);
      
      SpreadsheetApp.getUi().alert(
        'ECO Extraction Test\n\n' +
        'Game: ' + game.url + '\n\n' +
        'ECO Code: ' + ecoCode + '\n' +
        'ECO URL: ' + ecoUrl + '\n\n' +
        'Check logs for full PGN'
      );
      
      Logger.log('\n=== Full PGN ===');
      Logger.log(game.pgn);
    } else {
      SpreadsheetApp.getUi().alert('No games found');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    Logger.log(error.stack);
  }
}
function extractECOSlug(ecoUrl) {
  if (!ecoUrl || !ecoUrl.includes('chess.com/openings/')) return '';
  
  // Extract the slug part after '/openings/'
  const slug = ecoUrl.split('/openings/')[1] || '';
  if (!slug) return '';
  
  // Strategy: Find the first move sequence that's NOT part of a "with" pattern and trim from there
  
  // Pattern 1: with-NUMBER-MOVE-and-NUMBER-MOVE (keep this entire pattern)
  // Pattern 2: with-NUMBER-MOVE (keep this entire pattern)
  // Pattern 3: -NUMBER (where NUMBER is followed by . or ... indicating moves) - REMOVE from here onward
  
  // First, protect "with" patterns by replacing them temporarily
  let protected = slug;
  const withPatterns = [];
  
  // Match: with-NUMBER-MOVE-and-NUMBER-MOVE
  // Move can be: standard notation (Nf3, e4, etc.) or castling (O-O, O-O-O)
  const withAndPattern = /with-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)-and-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)/g;
  protected = protected.replace(withAndPattern, (match) => {
    const placeholder = `__WITH_AND_${withPatterns.length}__`;
    withPatterns.push(match);
    return placeholder;
  });
  
  // Match: with-NUMBER-MOVE (but not followed by -and-)
  // Move can be: standard notation (Nf3, e4, etc.) or castling (O-O, O-O-O)
  const withPattern = /with-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)(?!-and-)/g;
  protected = protected.replace(withPattern, (match) => {
    const placeholder = `__WITH_${withPatterns.length}__`;
    withPatterns.push(match);
    return placeholder;
  });
  
  // Now find the first move sequence indicator
  // Look for patterns like: -3...Nf6 or -4.g3 or -7...g6 or ...8.Nf3 or ...5.cxd4 or ...e6
  // These indicate the start of move notation
  // Pattern matches: 
  //   -NUMBER. or -NUMBER... (dash followed by move number)
  //   ...NUMBER. (three dots followed by move number)
  //   ...[a-zA-Z] (three dots followed by move notation without number)
  const movePattern = /(-\d+\.{0,3}[a-zA-Z]|\.{3}\d+\.|\.{3}[a-zA-Z])/;
  const moveMatch = protected.match(movePattern);
  
  if (moveMatch) {
    // Trim from the first move sequence
    protected = protected.substring(0, moveMatch.index);
  }
  
  // Restore "with" patterns
  for (let i = 0; i < withPatterns.length; i++) {
    protected = protected.replace(`__WITH_AND_${i}__`, withPatterns[i]);
    protected = protected.replace(`__WITH_${i}__`, withPatterns[i]);
  }
  
  return protected;
}

function loadOpeningsDbCache() {
  if (OPENINGS_DB_CACHE) return OPENINGS_DB_CACHE;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEETS.OPENINGS_DB);
  const cache = new Map();
  if (!dbSheet) {
    OPENINGS_DB_CACHE = cache;
    return cache;
  }
  const values = dbSheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    OPENINGS_DB_CACHE = cache;
    return cache;
  }
  const header = values[0];
  const slugIdx = header.indexOf('Trim Slug');
  const familyIdx = header.indexOf('Family');
  // The TSV has two 'Name' headers. We'll treat the first as Full Name and the second as Base Name.
  // Positions per OPENINGS_DB_HEADERS:
  // 0: Name (Full), 1: Trim Slug, 2: Family, 3: Name (Base), 4..9: Variation 1..6
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const trimSlug = String(row[1] || '').trim();
    if (!trimSlug) continue;
    const fullName = String(row[0] || '');
    const baseName = String(row[3] || '');
    const family = String(row[2] || '');
    const v1 = String(row[4] || '');
    const v2 = String(row[5] || '');
    const v3 = String(row[6] || '');
    const v4 = String(row[7] || '');
    const v5 = String(row[8] || '');
    const v6 = String(row[9] || '');
    cache.set(trimSlug, [fullName, family, baseName, v1, v2, v3, v4, v5, v6]);
  }
  OPENINGS_DB_CACHE = cache;
  return cache;
}

function normalizeSlugForDb(ecoSlug) {
  if (!ecoSlug) return '';
  // The DB uses Title-Case with hyphens; ecoSlug looks similar but may include lowercase and numbers like with-3-Nc3
  // We'll convert to Title-Case tokens separated by '-' and ensure castling tokens are normalized.
  const tokens = ecoSlug
    .replace(/_/g, '-')
    .split('-')
    .filter(Boolean)
    .map(tok => {
      if (/^with$/i.test(tok) || /^and$/i.test(tok)) return tok.charAt(0).toUpperCase() + tok.slice(1).toLowerCase();
      if (/^o$/i.test(tok)) return 'O';
      if (/^o\so$/i.test(tok)) return 'O-O';
      // Preserve chess move tokens/case like Nf3, e4, O-O, but capitalize words
      if (/^[a-z][a-z]+$/i.test(tok)) {
        return tok.charAt(0).toUpperCase() + tok.slice(1);
      }
      return tok;
    });
  return tokens.join('-');
}

function getDbMappingValues(ecoSlug) {
  // Returns array matching DERIVED_DB_HEADERS order
  const empty = ['', '', '', '', '', '', '', '', ''];
  if (!ecoSlug) return empty;
  const db = loadOpeningsDbCache();
  // Try direct match first
  if (db.has(ecoSlug)) return db.get(ecoSlug);
  // Try normalized form
  const normalized = normalizeSlugForDb(ecoSlug);
  if (db.has(normalized)) return db.get(normalized);
  // Try loosening: drop trailing move qualifiers like 'with-3-Nc3' if not found
  const withoutWith = ecoSlug.split('-with-')[0];
  if (withoutWith && db.has(withoutWith)) return db.get(withoutWith);
  const normalizedWithoutWith = normalized.split('-with-')[0];
  if (normalizedWithoutWith && db.has(normalizedWithoutWith)) return db.get(normalizedWithoutWith);
  return empty;
}

function getOpeningOutputs(ecoSlug) {
  const db = loadOpeningsDbCache();
  if (!ecoSlug) return ['', ''];
  const keys = [ecoSlug, normalizeSlugForDb(ecoSlug), ecoSlug.split('-with-')[0] || '', normalizeSlugForDb(ecoSlug).split('-with-')[0] || ''];
  for (const k of keys) {
    if (k && db.has(k)) {
      const row = db.get(k);
      return [row[0] || '', row[2] || '']; // Full Name, Family
    }
  }
  return ['', ''];
}

// ===== PGN PARSING HELPER FUNCTIONS =====
// Add these functions to your main code (document 1)

/**
 * Extract Chess.com opening URL from PGN
 * Returns the full URL like: https://www.chess.com/openings/Sicilian-Defense-Najdorf-Variation
 */
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

/**
 * Extract start time from PGN
 */
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

/**
 * Extract game duration from PGN (in seconds)
 */
function extractDurationFromPGN(pgn) {
  if (!pgn) return 0;
  
  const endMatch = pgn.match(/\[EndTime "([^"]+)"\]/);
  const startMatch = pgn.match(/\[StartTime "([^"]+)"\]/);
  
  if (!endMatch || !startMatch) return 0;
  
  try {
    // Parse time strings like "01:44:22" into seconds
    const startParts = startMatch[1].split(':').map(Number);
    const endParts = endMatch[1].split(':').map(Number);
    
    const startSeconds = startParts[0] * 3600 + startParts[1] * 60 + startParts[2];
    const endSeconds = endParts[0] * 3600 + endParts[1] * 60 + endParts[2];
    
    const duration = endSeconds - startSeconds;
    
    // Handle midnight crossover (negative duration)
    return duration >= 0 ? duration : duration + 86400;
  } catch (e) {
    Logger.log(`Error calculating duration: ${e.message}`);
    return 0;
  }
}

/**
 * Debug function - test opening extraction on a single game
 */
function testOpeningExtraction() {
  const username = CONFIG.USERNAME;
  const archiveUrl = `https://api.chess.com/pub/player/${username}/games/2025/10`;
  
  try {
    const response = UrlFetchApp.fetch(archiveUrl);
    const data = JSON.parse(response.getContentText());
    
    if (data.games && data.games.length > 0) {
      const game = data.games[0];
      const ecoUrl = extractECOFromPGN(game.pgn);
      const openingData = getOpeningDataForGame(ecoUrl);
      
      Logger.log('=== Test Results ===');
      Logger.log('Game URL: ' + game.url);
      Logger.log('ECO URL: ' + ecoUrl);
      Logger.log('Opening Data: ' + JSON.stringify(openingData));
      
      SpreadsheetApp.getUi().alert(
        'Test Results:\n\n' +
        'ECO URL: ' + ecoUrl + '\n\n' +
        'Opening Name: ' + openingData[0] + '\n' +
        'Opening Slug: ' + openingData[1] + '\n' +
        'Opening Family: ' + openingData[2] + '\n' +
        'See logs for full details'
      );
    } else {
      SpreadsheetApp.getUi().alert('No games found in October 2025');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    Logger.log(error);
  }
}
