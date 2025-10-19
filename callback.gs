// ===== IMPROVED RATINGS TRACKER =====
// Uses global format-specific ledger, not just last 100 games

class ImprovedRatingsTracker {
  constructor() {
    this.ledger = {};  // Format -> current rating
    this.loaded = false;
  }
  
  // Build index from ALL existing games (more reliable for rare formats)
  loadFromGamesSheet() {
    if (this.loaded) return;
    
    const gamesSheet = getGamesSheet();
    const lastRow = gamesSheet.getLastRow();
    
    if (lastRow <= 1) {
      this.loaded = true;
      return;
    }
    
    // Read ALL games to get most recent rating by format
    const data = gamesSheet.getRange(2, GAMES_COLS.FORMAT, lastRow - 1, 2).getValues();
    // [format, myRating]
    
    // Build ledger: most recent (last row) rating for each format
    const formatLastRating = {};
    
    for (let i = 0; i < data.length; i++) {
      const format = data[i][0];
      const rating = data[i][1];
      
      if (format && rating) {
        // Keep updating - last one wins (most recent)
        formatLastRating[format] = rating;
      }
    }
    
    this.ledger = formatLastRating;
    this.loaded = true;
    
    Logger.log('Loaded ratings ledger: ' + JSON.stringify(this.ledger));
  }
  
  // Calculate rating before and delta for a new game
  calculateRating(format, currentRating) {
    const ratingBefore = this.ledger[format] || null;
    const ratingDelta = (ratingBefore !== null && currentRating !== null) 
      ? (currentRating - ratingBefore) 
      : null;
    
    // Update ledger for next game
    if (currentRating !== null) {
      this.ledger[format] = currentRating;
    }
    
    return {
      before: ratingBefore,
      delta: ratingDelta
    };
  }
}

// ===== SIMPLIFIED REGISTRY (OPTIONAL) =====
// Registry is now optional since Games sheet has all status info
// Keep it only for quick lookups if needed

function setupRegistrySheet() {
  const ss = getControlSpreadsheet();
  const sheet = ss.getSheetByName('Registry') || ss.insertSheet('Registry');
  
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Game ID',
      'Archive',
      'Format',
      'Date Added'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    sheet.setFrozenRows(1);
    sheet.getRange('D:D').setNumberFormat('M/D/YY h:mm');
  }
  
  return sheet;
}

// ===== ADD GAME TO REGISTRY (MINIMAL) =====
function addToRegistry(gameId, archive, format) {
  const registrySheet = setupRegistrySheet();
  const now = new Date();
  
  const row = [
    gameId,
    archive,
    format,
    now
  ];
  
  const lastRow = registrySheet.getLastRow();
  registrySheet.getRange(lastRow + 1, 1, 1, 4).setValues([row]);
}

// ===== BUILD GAME ID INDEX (USING ARCHIVE ANCHORS) =====
// More efficient duplicate detection using archive anchors

class ArchiveBasedGameIndex {
  constructor() {
    this.index = null;
    this.archiveAnchors = {};  // Archive -> last game ID
  }
  
  build() {
    if (this.index) return this.index;
    
    const gamesSheet = getGamesSheet();
    const lastRow = gamesSheet.getLastRow();
    
    this.index = new Set();
    
    if (lastRow > 1) {
      // Read game IDs and archives
      const data = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, lastRow - 1, 2).getValues();
      // [gameId, archive or next column if needed]
      
      data.forEach(row => {
        const gameId = String(row[0]);
        this.index.add(gameId);
      });
      
      // Build archive anchors (last game per archive)
      const archiveData = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, lastRow - 1, 2).getValues();
      // [gameId, need archive col]
      const fullData = gamesSheet.getRange(2, GAMES_COLS.GAME_ID, lastRow - 1, GAMES_COLS.ARCHIVE).getValues();
      
      for (let i = fullData.length - 1; i >= 0; i--) {
        const gameId = String(fullData[i][GAMES_COLS.GAME_ID - 1]);
        const archive = fullData[i][GAMES_COLS.ARCHIVE - 1];
        
        if (archive && !this.archiveAnchors[archive]) {
          this.archiveAnchors[archive] = gameId;
        }
      }
      
      Logger.log(`Built index with ${this.index.size} games and ${Object.keys(this.archiveAnchors).length} archive anchors`);
    }
    
    return this.index;
  }
  
  has(gameId) {
    if (!this.index) this.build();
    return this.index.has(String(gameId));
  }
  
  add(gameId) {
    if (!this.index) this.build();
    this.index.add(String(gameId));
  }
  
  // Get anchor (last imported game ID) for an archive
  getAnchor(archive) {
    if (!this.index) this.build();
    return this.archiveAnchors[archive] || null;
  }
  
  // Set new anchor for an archive
  setAnchor(archive, gameId) {
    if (!this.index) this.build();
    this.archiveAnchors[archive] = String(gameId);
  }
}

// ===== FILTER NEW GAMES (USING ANCHORS) =====
function filterNewGamesWithAnchors(games, archive) {
  const index = new ArchiveBasedGameIndex();
  index.build();
  
  const anchor = getLastGameIdForArchive(archive);
  
  if (!anchor) {
    // No anchor - check all games against index
    return games.filter(game => {
      const gameId = game.url.split('/').pop();
      return !index.has(gameId);
    });
  }
  
  // Have anchor - only check games after anchor
  const anchorIndex = games.findIndex(game => game.url.split('/').pop() === anchor);
  
  if (anchorIndex === -1) {
    // Anchor not found in current batch - something changed, check all
    Logger.log(`Warning: Anchor ${anchor} not found in archive ${archive}`);
    return games.filter(game => {
      const gameId = game.url.split('/').pop();
      return !index.has(gameId);
    });
  }
  
  // Return only games after anchor
  const newGames = games.slice(anchorIndex + 1);
  Logger.log(`Using anchor: ${anchor}, found ${newGames.length} new games in archive`);
  
  return newGames;
}

// ===== STANDARD FILTER (FALLBACK) =====
function filterNewGames(games) {
  const index = new ArchiveBasedGameIndex();
  index.build();
  
  return games.filter(game => {
    const gameId = game.url.split('/').pop();
    return !index.has(gameId);
  });
}
