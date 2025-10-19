// ===== EXPANDED GAMES SHEET SETUP =====
// Includes all original data: duration, opening details, moves, etc.

const GAMES_COLS = {
  GAME_ID: 1,
  TYPE: 2,
  GAME_URL: 3,
  PGN: 4,
  
  // Dates & Times
  START_DATETIME: 5,
  START_DATE: 6,
  START_TIME: 7,
  START_EPOCH: 8,
  END_DATETIME: 9,
  END_DATE: 10,
  END_TIME: 11,
  END_EPOCH: 12,
  ARCHIVE: 13,
  
  // Game Details
  RULES: 14,
  LIVE: 15,
  TIME_CLASS: 16,
  FORMAT: 17,
  RATED: 18,
  TIME_CONTROL: 19,
  BASE: 20,
  INC: 21,
  DURATION: 22,
  DURATION_S: 23,
  
  // Players
  COLOR: 24,
  OPPONENT: 25,
  
  // Ratings
  MY_RATING: 26,
  OPP_RATING: 27,
  RATING_BEFORE: 28,
  RATING_DELTA: 29,
  
  // Result
  OUTCOME: 30,
  TERMINATION: 31,
  
  // Opening
  ECO: 32,
  ECO_URL: 33,
  OPENING_NAME: 34,
  OPENING_SLUG: 35,
  OPENING_FAMILY: 36,
  OPENING_BASE: 37,
  VAR_1: 38,
  VAR_2: 39,
  VAR_3: 40,
  VAR_4: 41,
  VAR_5: 42,
  VAR_6: 43,
  EXTRA_MOVES: 44,
  
  // Move Data
  MOVES: 45,
  TCN: 46,
  
  // Enrichment Status
  CALLBACK_STATUS: 47,
  CALLBACK_DATE: 48,
  LICHESS_STATUS: 49,
  LICHESS_URL: 50,
  
  // Metadata
  FETCH_DATE: 51,
  LAST_UPDATED: 52
};

function setupExpandedGamesSheet() {
  const ss = getMainSpreadsheet();
  const sheet = ss.getSheetByName('Games') || ss.insertSheet('Games');
  
  // Clear if exists
  if (sheet.getLastRow() > 0) {
    sheet.clear();
  }
  
  const headers = [
    'Game ID', 'Type', 'Game URL', 'PGN',
    'Start Date/Time', 'Start Date', 'Start Time', 'Start (s)',
    'End Date/Time', 'End Date', 'End Time', 'End (s)', 'Archive',
    'Rules', 'Live', 'Time Class', 'Format', 'Rated', 'Time Control', 'Base', 'Inc', 'Duration', 'Duration (s)',
    'Color', 'Opponent',
    'My Rating', 'Opp Rating', 'Rating Before', 'Rating Δ',
    'Outcome', 'Termination',
    'ECO', 'ECO URL', 'Opening Name', 'Opening Slug', 'Opening Family', 'Opening Base',
    'Variation 1', 'Variation 2', 'Variation 3', 'Variation 4', 'Variation 5', 'Variation 6', 'Extra Moves',
    'Moves', 'TCN',
    'Callback Status', 'Callback Date',
    'Lichess Status', 'Lichess URL',
    'Fetch Date', 'Last Updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // Formatting
  sheet.setFrozenRows(1);
  sheet.getRange('F:F').setNumberFormat('M/D/YY');
  sheet.getRange('G:G').setNumberFormat('h:mm AM/PM');
  sheet.getRange('J:J').setNumberFormat('M/D/YY');
  sheet.getRange('K:K').setNumberFormat('h:mm AM/PM');
  sheet.getRange('V:V').setNumberFormat('[h]:mm:ss');
  sheet.getRange('W:W').setNumberFormat('0');
  sheet.getRange('Z:AC').setNumberFormat('0');
  sheet.getRange('AS:AS').setNumberFormat('0');
  sheet.getRange('AV:AV').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('AZ:AZ').setNumberFormat('M/D/YY h:mm');
  
  // Column widths
  sheet.setColumnWidth(1, 90);   // Game ID
  sheet.setColumnWidth(2, 60);   // Type
  sheet.setColumnWidth(3, 250);  // URL
  sheet.setColumnWidth(4, 400);  // PGN
  sheet.setColumnWidth(25, 120); // Opponent
  sheet.setColumnWidth(34, 200); // Opening Name
  
  // Conditional formatting
  const rules = [];
  
  // Outcomes
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('win')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('loss')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());
  
  // Rating changes
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('AC2:AC')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('AC2:AC')])
    .build());
  
  // Callback status
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('fetched')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('AU2:AU')])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  
  // Hide columns
  sheet.hideColumns(1, 2);  // ID and Type
  sheet.hideColumns(4);     // PGN
  sheet.hideColumns(5);     // Start Date/Time
  sheet.hideColumns(8);     // Start Epoch
  sheet.hideColumns(9);     // End Date/Time
  sheet.hideColumns(12);    // End Epoch
  sheet.hideColumns(20, 2); // Base and Inc
  sheet.hideColumns(23);    // Duration (s)
  sheet.hideColumns(33);    // ECO URL
  sheet.hideColumns(35, 10); // Opening details
  sheet.hideColumns(46);    // TCN
  
  SpreadsheetApp.getUi().alert('✅ Expanded Games sheet setup complete!');
}

// ===== IMPROVED FORMAT DETECTION =====
function getGameFormat(game) {
  const rules = (game.rules || 'chess').toLowerCase();
  let timeClass = (game.time_class || '').toLowerCase();
  
  // Handle variants
  if (rules === 'chess960') {
    return timeClass === 'daily' ? 'daily960' : 'live960';
  } else if (rules !== 'chess') {
    // Return variant name: bughouse, crazyhouse, kingofthehill, threecheck, oddschess
    return rules;
  }
  
  // Standard chess - use time class
  if (['bullet', 'blitz', 'rapid', 'daily'].includes(timeClass)) {
    return timeClass;
  }
  
  // Fallback: parse time control
  const tc = game.time_control || '';
  const match = tc.match(/(\d+)\+(\d+)/);
  
  if (!match) return timeClass || 'unknown';
  
  const base = parseInt(match[1]);
  const inc = parseInt(match[2]);
  const estimated = base + 40 * inc;
  
  if (estimated < 180) return 'bullet';
  if (estimated < 600) return 'blitz';
  return 'rapid';
}

// ===== GET OPENING DATA FOR GAME =====
function getOpeningDataForGame(ecoUrl) {
  if (!ecoUrl) return ['', '', '', '', '', '', '', '', '', '', ''];
  
  // Extract slug from ECO URL
  const match = ecoUrl.match(/\/openings\/([^"]+)$/);
  if (!match) return ['', '', '', '', '', '', '', '', '', '', ''];
  
  const slug = match[1];
  const parts = slug.split('-');
  
  // Parse opening structure
  const openingName = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
  const openingSlug = slug;
  
  // Determine family, base, and variations
  let family = '';
  let base = '';
  const variations = [];
  
  if (parts.length > 0) {
    family = parts[0];
  }
  if (parts.length > 1) {
    base = parts.slice(0, 2).join(' ');
  }
  if (parts.length > 2) {
    for (let i = 2; i < Math.min(parts.length, 8); i++) {
      variations.push(parts[i]);
    }
  }
  
  const extraMoves = parts.length > 8 ? parts.slice(8).join(' ') : '';
  
  // Pad variations array to 6 elements
  while (variations.length < 6) {
    variations.push('');
  }
  
  return [
    openingName,
    openingSlug,
    family,
    base,
    ...variations.slice(0, 6),
    extraMoves
  ];
}
