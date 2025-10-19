// ===== COMPREHENSIVE HELPER FUNCTIONS =====
// PGN parsing, opening extraction, game processing utilities

// ===== EXTRACT START TIME FROM PGN =====
function extractStartFromPGN(pgn) {
  if (!pgn) return null;
  
  try {
    // Look for StartTime tag: [StartTime "HH:MM:SS"]
    const startTimeMatch = pgn.match(/\[StartTime\s+"([^"]+)"\]/);
    
    // Look for Date and UTCDate tags
    const dateMatch = pgn.match(/\[Date\s+"([^"]+)"\]/);
    const utcDateMatch = pgn.match(/\[UTCDate\s+"([^"]+)"\]/);
    
    if (startTimeMatch && (dateMatch || utcDateMatch)) {
      const time = startTimeMatch[1];
      const date = utcDateMatch ? utcDateMatch[1] : dateMatch[1];
      
      // Parse date (format: YYYY.MM.DD)
      const dateParts = date.split('.');
      if (dateParts.length === 3) {
        const [year, month, day] = dateParts.map(p => parseInt(p));
        
        // Parse time (format: HH:MM:SS)
        const timeParts = time.split(':');
        if (timeParts.length === 3) {
          const [hour, minute, second] = timeParts.map(p => parseInt(p));
          
          return new Date(year, month - 1, day, hour, minute, second);
        }
      }
    }
  } catch (error) {
    Logger.log(`Error parsing start time from PGN: ${error.message}`);
  }
  
  return null;
}

// ===== EXTRACT DURATION FROM PGN =====
function extractDurationFromPGN(pgn) {
  if (!pgn) return null;
  
  try {
    const startDate = extractStartFromPGN(pgn);
    
    // Also look for EndTime or EndDate
    const endTimeMatch = pgn.match(/\[EndTime\s+"([^"]+)"\]/);
    const endDateMatch = pgn.match(/\[EndDate\s+"([^"]+)"\]/);
    const utcTimeMatch = pgn.match(/\[UTCTime\s+"([^"]+)"\]/);
    
    if (startDate && endTimeMatch && endDateMatch) {
      const time = endTimeMatch[1];
      const date = endDateMatch[1];
      
      const dateParts = date.split('.');
      if (dateParts.length === 3) {
        const [year, month, day] = dateParts.map(p => parseInt(p));
        
        const timeParts = time.split(':');
        if (timeParts.length === 3) {
          const [hour, minute, second] = timeParts.map(p => parseInt(p));
          
          const endDate = new Date(year, month - 1, day, hour, minute, second);
          const durationMs = endDate - startDate;
          
          return Math.floor(durationMs / 1000);  // Return seconds
        }
      }
    }
  } catch (error) {
    Logger.log(`Error parsing duration from PGN: ${error.message}`);
  }
  
  return null;
}

// ===== EXTRACT ECO CODE FROM PGN =====
function extractECOCodeFromPGN(pgn) {
  if (!pgn) return '';
  
  const match = pgn.match(/\[ECO\s+"([^"]+)"\]/);
  return match ? match[1] : '';
}

// ===== EXTRACT ECO URL FROM PGN =====
function extractECOFromPGN(pgn) {
  if (!pgn) return '';
  
  const match = pgn.match(/\[ECOUrl\s+"([^"]+)"\]/);
  return match ? match[1] : '';
}

// ===== EXTRACT OPENING NAME FROM PGN =====
function extractOpeningNameFromPGN(pgn) {
  if (!pgn) return '';
  
  // Try ECOUrl first
  const ecoUrl = extractECOFromPGN(pgn);
  if (ecoUrl) {
    const match = ecoUrl.match(/\/openings\/([^"]+)$/);
    if (match) {
      const slug = match[1];
      return slug
        .split('-')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
    }
  }
  
  // Fallback to Opening tag
  const match = pgn.match(/\[Opening\s+"([^"]+)"\]/);
  return match ? match[1] : '';
}

// ===== EXTRACT MOVES WITH CLOCKS =====
function extractMovesWithClocks(pgn, baseTime, increment) {
  const result = {
    plyCount: 0,
    clocks: []
  };
  
  if (!pgn) return result;
  
  try {
    // Extract move section (after headers, before result)
    const moveSection = pgn.split(/\n\n/)[1] || '';
    
    // Count plies (individual moves)
    const moves = moveSection.match(/\d+\.\s+\S+(\s+\S+)?/g) || [];
    result.plyCount = moves.reduce((count, move) => {
      // Each "1. e4 e5" is 2 plies
      const plies = move.match(/\S+/g).length - 1;  // -1 for move number
      return count + plies;
    }, 0);
    
    // Extract clock times
    const clockMatches = moveSection.match(/\{\[%clk\s+([^\]]+)\]\}/g) || [];
    result.clocks = clockMatches.map(match => {
      const timeStr = match.match(/\[%clk\s+([^\]]+)\]/)[1];
      return parseClockTime(timeStr);
    });
    
  } catch (error) {
    Logger.log(`Error extracting moves: ${error.message}`);
  }
  
  return result;
}

// ===== PARSE CLOCK TIME =====
function parseClockTime(timeStr) {
  // Parse format like "0:05:32" or "1:23:45.6"
  const parts = timeStr.split(':');
  
  if (parts.length === 3) {
    const hours = parseInt(parts[0]);
    const minutes = parseInt(parts[1]);
    const seconds = parseFloat(parts[2]);
    
    return Math.floor(hours * 3600 + minutes * 60 + seconds);
  }
  
  return 0;
}

// ===== GET GAME OUTCOME =====
function getGameOutcome(game, username) {
  if (!game.white || !game.black) return 'unknown';
  
  const isWhite = game.white.username.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  
  // Win conditions
  if (myResult === 'win') return 'win';
  if (['checkmated', 'resigned', 'timeout', 'abandoned'].includes(myResult)) return 'loss';
  
  // Draw conditions
  if (['agreed', 'stalemate', 'repetition', 'insufficient', 'timevsinsufficient', '50move'].includes(myResult)) {
    return 'draw';
  }
  
  return 'unknown';
}

// ===== GET GAME TERMINATION =====
function getGameTermination(game, username) {
  if (!game.white || !game.black) return 'unknown';
  
  const isWhite = game.white.username.toLowerCase() === username.toLowerCase();
  const myResult = isWhite ? game.white.result : game.black.result;
  
  return myResult || 'unknown';
}

// ===== IMPROVED FORMAT DETECTION =====
function getGameFormat(game) {
  const rules = (game.rules || 'chess').toLowerCase();
  let timeClass = (game.time_class || '').toLowerCase();
  
  // Handle Chess960
  if (rules === 'chess960') {
    return timeClass === 'daily' ? 'daily960' : 'live960';
  }
  
  // Handle other variants
  if (rules !== 'chess') {
    // Return variant name: bughouse, crazyhouse, kingofthehill, threecheck, oddschess
    return rules;
  }
  
  // Standard chess - use time class if valid
  if (['bullet', 'blitz', 'rapid', 'daily'].includes(timeClass)) {
    return timeClass;
  }
  
  // Fallback: calculate from time control
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
  
  // Build opening name
  const openingName = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
  const openingSlug = slug;
  
  // Determine structure
  let family = parts[0] || '';
  let base = parts.length > 1 ? parts.slice(0, 2).join('-') : '';
  
  // Extract variations (positions 2-7)
  const variations = [];
  for (let i = 2; i < Math.min(parts.length, 8); i++) {
    variations.push(parts[i]);
  }
  
  // Extra moves (position 8+)
  const extraMoves = parts.length > 8 ? parts.slice(8).join('-') : '';
  
  // Pad variations to 6 elements
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

// ===== PARSE TIME CONTROL =====
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

// ===== FORMAT TIME CONTROL LABEL =====
function formatTimeControlLabel(base, inc, corr) {
  if (corr !== null) {
    return `${corr} days`;
  }
  if (base !== null && inc !== null) {
    return `${base}+${inc}`;
  }
  return '';
}

// ===== DATE UTILITIES =====
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

function dateToSerial(date) {
  // Excel serial date
  const epoch = new Date(1899, 11, 30);
  const diff = date - epoch;
  return diff / (1000 * 60 * 60 * 24);
}

// ===== CALCULATE TIME CLASS FROM TIME CONTROL =====
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

// ===== ENCODE CLOCKS (SIMPLIFIED) =====
function encodeClocksBase36(clocks) {
  if (!clocks || !clocks.length) return '';
  
  return clocks.map(c => c.toString(36)).join('.');
}

// ===== DECODE CLOCKS (SIMPLIFIED) =====
function decodeClocksBase36(encoded) {
  if (!encoded) return [];
  
  return encoded.split('.').map(s => parseInt(s, 36) || 0);
}
