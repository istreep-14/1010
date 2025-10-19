// ================================
// OPENINGS DATABASE - EXTERNAL LOOKUP
// ================================

const OPENINGS_DB_CONFIG = {
  SPREADSHEET_ID: '1PWyey0pm7IkI8T_y6BLt9eisvOb51jdpgEWe91tyq44', // Replace with your database spreadsheet ID
  SHEET_NAME: 'Openings',
  cache: null,
  lastCacheTime: null,
  CACHE_DURATION_MS: 5 * 60 * 1000 // 5 minutes
};

// What we store in the games sheet
const DERIVED_OPENING_HEADERS = [
  'Opening Name', 'Opening Slug', 'Opening Family', 'Opening Base',
  'Variation 1', 'Variation 2', 'Variation 3', 'Variation 4', 'Variation 5', 'Variation 6',
  'Extra Moves'
];

// ================================
// MAIN LOOKUP FUNCTION
// ================================

/**
 * Main function: Takes ECO URL, returns all opening data
 * This is the ONLY function you need to call from processGames()
 */
function getOpeningDataForGame(ecoUrl) {
  const empty = ['', '', '', '', '', '', '', '', '', '', ''];
  if (!ecoUrl) return empty;
  
  // Step 1: Split URL into base slug + extra moves
  const { baseSlug, extraMoves } = splitEcoUrl(ecoUrl);
  if (!baseSlug) return empty;
  
  // Step 2: Load database (cached)
  const db = loadOpeningsDb();
  
  // Step 3: Lookup in database
  const dbRow = lookupInDb(db, baseSlug);
  
  // Step 4: Format extra moves from slug to PGN notation
  const formattedExtraMoves = formatExtraMovesV2(extraMoves);
  
  // Step 5: Return all fields + formatted extra moves
  return [...dbRow, formattedExtraMoves];
}


// ================================
// HELPER FUNCTIONS (Internal)
// ================================

/**
 * Split ECO URL into base slug and extra moves
 * Example: "...openings/Sicilian-Defense-5.Nc3" 
 *   → { baseSlug: "sicilian-defense", extraMoves: "5.Nc3" }
 */
function splitEcoUrl(ecoUrl) {
  if (!ecoUrl || !ecoUrl.includes('chess.com/openings/')) {
    return { baseSlug: '', extraMoves: '' };
  }
  
  const fullSlug = ecoUrl.split('/openings/')[1] || '';
  if (!fullSlug) return { baseSlug: '', extraMoves: '' };
  
  let slug = fullSlug;
  const withPatterns = [];
  
  // Protect "with-NUMBER-MOVE" patterns (these are part of opening names)
  slug = slug.replace(/with-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+)(?:-and-(\d+)-(O-O(?:-O)?|[a-zA-Z0-9]+))?/g, (match) => {
    const placeholder = `__WITH_${withPatterns.length}__`;
    withPatterns.push(match);
    return placeholder;
  });
  
  // Find first move sequence: -3...Nf6 or -4.g3 or ...8.Nf3
  const movePattern = /(-\d+\.{0,3}[a-zA-Z]|\.{3}\d+\.|\.{3}[a-zA-Z])/;
  const moveMatch = slug.match(movePattern);
  
  let baseSlug, extraMoves;
  if (moveMatch) {
    baseSlug = slug.substring(0, moveMatch.index);
    extraMoves = slug.substring(moveMatch.index);
  } else {
    baseSlug = slug;
    extraMoves = '';
  }
  
  // Restore "with" patterns
  withPatterns.forEach((pattern, i) => {
    baseSlug = baseSlug.replace(`__WITH_${i}__`, pattern);
  });
  
  return { 
    baseSlug: baseSlug.toLowerCase(), 
    extraMoves 
  };
}

/**
 * Load database from external spreadsheet (with caching)
 */
function loadOpeningsDb() {
  const now = Date.now();
  
  // Return cache if valid
  if (OPENINGS_DB_CONFIG.cache && 
      OPENINGS_DB_CONFIG.lastCacheTime &&
      (now - OPENINGS_DB_CONFIG.lastCacheTime) < OPENINGS_DB_CONFIG.CACHE_DURATION_MS) {
    return OPENINGS_DB_CONFIG.cache;
  }
  
  const cache = new Map();
  
  try {
    const dbSpreadsheet = SpreadsheetApp.openById(OPENINGS_DB_CONFIG.SPREADSHEET_ID);
    const dbSheet = dbSpreadsheet.getSheetByName(OPENINGS_DB_CONFIG.SHEET_NAME);
    
    if (!dbSheet) {
      Logger.log('Openings DB sheet not found');
      OPENINGS_DB_CONFIG.cache = cache;
      OPENINGS_DB_CONFIG.lastCacheTime = now;
      return cache;
    }
    
    const values = dbSheet.getDataRange().getValues();
    if (values.length < 2) {
      OPENINGS_DB_CONFIG.cache = cache;
      OPENINGS_DB_CONFIG.lastCacheTime = now;
      return cache;
    }
    
    // Load rows: [Name, Trim Slug, Family, Base Name, Var1-6]
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const trimSlug = String(row[1] || '').trim().toLowerCase();
      if (!trimSlug) continue;
      
      cache.set(trimSlug, [
        String(row[0] || ''),  // Full Name
        trimSlug,              // Slug
        String(row[2] || ''),  // Family
        String(row[3] || ''),  // Base Name
        String(row[4] || ''),  // Variation 1
        String(row[5] || ''),  // Variation 2
        String(row[6] || ''),  // Variation 3
        String(row[7] || ''),  // Variation 4
        String(row[8] || ''),  // Variation 5
        String(row[9] || '')   // Variation 6
      ]);
    }
    
    OPENINGS_DB_CONFIG.cache = cache;
    OPENINGS_DB_CONFIG.lastCacheTime = now;
    Logger.log(`Loaded ${cache.size} openings`);
    
  } catch (error) {
    Logger.log(`Error loading openings: ${error.message}`);
  }
  
  return cache;
}

/**
 * Lookup slug in database with fallback logic
 */
function lookupInDb(db, slug) {
  const empty = ['', '', '', '', '', '', '', '', '', ''];
  
  // Try direct match
  if (db.has(slug)) {
    return db.get(slug);
  }
  
  // Try without "with-" suffix
  const withoutWith = slug.split('-with-')[0];
  if (withoutWith && withoutWith !== slug && db.has(withoutWith)) {
    return db.get(withoutWith);
  }
  
  // Not found - return empty with just the slug
  return ['', slug, '', '', '', '', '', '', '', ''];
}

// ================================
// UTILITY FUNCTIONS
// ================================

/**
 * Refresh all games with latest database data
 */
function refreshOpeningDataFromExternalDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Games sheet not found!');
    return;
  }
  
  // Force cache refresh
  OPENINGS_DB_CONFIG.cache = null;
  loadOpeningsDb();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('ℹ️ No games to update');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const ecoIdx = headers.indexOf('ECO');
  const openingStartIdx = headers.indexOf('Opening Name');
  
  if (ecoIdx === -1 || openingStartIdx === -1) {
    SpreadsheetApp.getUi().alert('❌ Required columns not found!');
    return;
  }
  
  ss.toast('Refreshing opening data...', '⏳', -1);
  
  const ecoUrls = sheet.getRange(2, ecoIdx + 1, lastRow - 1, 1).getValues();
  const updates = ecoUrls.map(row => getOpeningDataForGame(String(row[0] || '')));
  
  sheet.getRange(2, openingStartIdx + 1, updates.length, DERIVED_OPENING_HEADERS.length)
    .setValues(updates);
  
  ss.toast(`✅ Updated ${updates.length} games!`, '✅', 5);
}

/**
 * Test database connection
 */
function testOpeningsDbConnection() {
  OPENINGS_DB_CONFIG.cache = null;
  const db = loadOpeningsDb();
  const size = db.size;
  
  if (size > 0) {
    const samples = [];
    let count = 0;
    for (const [slug, data] of db.entries()) {
      samples.push(`${slug} → ${data[0]}`);
      if (++count >= 3) break;
    }
    
    SpreadsheetApp.getUi().alert(
      `✅ Connected! ${size} openings loaded\n\n` +
      'Samples:\n' + samples.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert('⚠️ Database empty or not found');
  }
}

function formatExtraMovesV2(extraMovesSlug) {
  if (!extraMovesSlug || extraMovesSlug.trim() === '') {
    return '';
  }
  
  let slug = extraMovesSlug.trim();
  slug = slug.replace(/^[-\.]+/, '');
  if (!slug) return '';
  
  const tokens = slug.split('-').filter(Boolean);
  if (tokens.length === 0) return '';
  
  const moves = [];
  let i = 0;
  
  while (i < tokens.length) {
    const token = tokens[i];
    const moveNumMatch = token.match(/^(\d+)(\.{0,3})$/);
    
    if (moveNumMatch) {
      const num = moveNumMatch[1];
      const dots = moveNumMatch[2];
      
      if (dots === '...') {
        moves.push(`${num}...`);
        i++;
        if (i < tokens.length && !tokens[i].match(/^\d+\.{0,3}$/)) {
          moves.push(tokens[i]);
          i++;
        }
      } else {
        moves.push(`${num}.`);
        i++;
        if (i < tokens.length && !tokens[i].match(/^\d+\.{0,3}$/)) {
          moves.push(tokens[i]);
          i++;
          if (i < tokens.length && !tokens[i].match(/^\d+\.{0,3}$/)) {
            moves.push(tokens[i]);
            i++;
          }
        }
      }
    } else {
      moves.push(token);
      i++;
    }
  }
  
  return moves.join(' ');
}
