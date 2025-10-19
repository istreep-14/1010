// ===== CUSTOM MENU =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('♟️ Chess Tracker')
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('Setup All Sheets', 'setupAll')
      .addItem('Setup Games Sheet', 'setupExpandedGamesSheet')
      .addItem('Setup Summary Sheet', 'setupSummarySheet')
      .addSeparator()
      .addItem('Show Configuration', 'showConfiguration')
      .addItem('Reset All Connections', 'resetAllConnections'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('📥 Fetch Games')
      .addItem('Fetch New Games', 'fetchNewGames')
      .addItem('Fetch All History', 'fetchAllHistory')
      .addSeparator()
      .addItem('Fetch Specific Month...', 'fetchSpecificMonthPrompt'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('⭐ Enrichment')
      .addItem('Enrich Recent Games (20)', 'enrichRecentGames')
      .addItem('Enrich All Pending', 'enrichAllPendingCallbacks')
      .addSeparator()
      .addItem('Show Enrichment Status', 'showEnrichmentStatus'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('♟️ Lichess')
      .addItem('Export Games to Lichess', 'exportGamesToLichess')
      .addItem('Mark as Imported', 'markGamesAsLichessImported')
      .addSeparator()
      .addItem('Lichess Guide', 'showLichessMenu'))
    
    .addSeparator()
    
    .addItem('📊 View Summary', 'goToSummary')
    
    .addToUi();
}

// ===== SETUP ALL =====
function setupAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Setup All Sheets?',
    'This will set up:\n' +
    '• Games sheet (main spreadsheet)\n' +
    '• Summary sheet (main spreadsheet)\n' +
    '• Control sheets (auto-created)\n' +
    '• Enrichment sheets (auto-created)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Setting up sheets...', '⏳', -1);
    
    // Main sheets
    setupExpandedGamesSheet();
    setupSummarySheet();
    
    // Control sheets
    setupArchivesSheet();
    setupRegistrySheet();
    
    // Enrichment sheets
    setupCallbackSheet();
    setupLichessSheet();
    
    ui.alert(
      '✅ Setup Complete!',
      'All sheets have been created.\n\n' +
      'Next steps:\n' +
      '1. Run "Fetch New Games" to start importing\n' +
      '2. Run "Enrich Recent Games" to get accurate ratings\n\n' +
      'Use "Show Configuration" to see your spreadsheet IDs.'
    );
    
  } catch (error) {
    ui.alert('❌ Error: ' + error.message);
    Logger.log(error);
  }
}

// ===== ENRICHMENT HELPERS =====
function enrichRecentGames() {
  enrichNewGamesWithCallbacks(20);
}

function showEnrichmentStatus() {
  const gamesSheet = getGamesSheet();
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No games found');
    return;
  }
  
  const statuses = gamesSheet.getRange(2, GAMES_COLS.CALLBACK_STATUS, lastRow - 1, 1).getValues();
  
  let fetched = 0;
  let noRating = 0;
  let pending = 0;
  let errors = 0;
  
  for (const [status] of statuses) {
    if (status === 'fetched') {
      fetched++;
    } else if (status === 'no_rating') {
      noRating++;
    } else if (status === 'error') {
      errors++;
    } else {
      pending++;
    }
  }
  
  const totalGames = lastRow - 1;
  const completePercent = totalGames > 0 ? ((fetched / totalGames) * 100).toFixed(1) : 0;
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { color: #4285f4; }
      .stat { margin: 10px 0; font-size: 14px; }
      .complete { color: #0f9d58; font-weight: bold; }
      .pending { color: #f4b400; }
      .norating { color: #9e9e9e; }
      .error { color: #db4437; }
      .progress-bar {
        width: 100%;
        height: 30px;
        background: #e0e0e0;
        border-radius: 4px;
        overflow: hidden;
        margin: 10px 0;
      }
      .progress-fill {
        height: 100%;
        background: #0f9d58;
        text-align: center;
        line-height: 30px;
        color: white;
        font-weight: bold;
      }
    </style>
    <h2>📊 Enrichment Progress</h2>
    <div class="stat"><strong>Total Games:</strong> ${totalGames}</div>
    
    <div class="progress-bar">
      <div class="progress-fill" style="width: ${completePercent}%">
        ${completePercent}%
      </div>
    </div>
    
    <div class="stat complete">✅ Fetched (Valid Ratings): ${fetched}</div>
    <div class="stat norating">📊 Fetched (No Rating Change): ${noRating}</div>
    <div class="stat pending">⏳ Pending: ${pending}</div>
    <div class="stat error">⚠️ Errors: ${errors}</div>
    
    <hr>
    <p>Run "Enrich Recent Games" or "Enrich All Pending" from the menu to continue.</p>
  `)
    .setWidth(450)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enrichment Status');
}

// ===== NAVIGATION =====
function goToSummary() {
  const ss = getMainSpreadsheet();
  const summarySheet = ss.getSheetByName('Summary');
  
  if (summarySheet) {
    ss.setActiveSheet(summarySheet);
  } else {
    SpreadsheetApp.getUi().alert('Summary sheet not found. Run "Setup Summary Sheet" first.');
  }
}

// ===== FETCH SPECIFIC MONTH =====
function fetchSpecificMonthPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Fetch Specific Month',
    'Enter archive in format YYYY-MM (e.g., 2024-10):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const archiveMonth = response.getResponseText().trim();
  
  if (!/^\d{4}-\d{2}$/.test(archiveMonth)) {
    ui.alert('Invalid format. Please use YYYY-MM (e.g., 2024-10)');
    return;
  }
  
  const [year, month] = archiveMonth.split('-');
  const archiveUrl = `https://api.chess.com/pub/player/${CONFIG.USERNAME}/games/${year}/${month}`;
  
  fetchChesscomGames({ specificArchive: archiveUrl });
}

// ===== HELPER: EXTRACT OPENING NAME FROM PGN =====
function extractOpeningNameFromPGN(pgn) {
  if (!pgn) return '';
  
  const match = pgn.match(/\[ECOUrl\s+"https:\/\/www\.chess\.com\/openings\/([^"]+)"\]/);
  if (!match) return '';
  
  const slug = match[1];
  
  // Convert slug to readable name
  const name = slug
    .split('-')
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
  
  return name;
}

// ===== HELPER: EXTRACT ECO CODE FROM PGN =====
function extractECOCodeFromPGN(pgn) {
  if (!pgn) return '';
  
  const match = pgn.match(/\[ECO\s+"([^"]+)"\]/);
  return match ? match[1] : '';
}

// ===== HELPER: GET GAME OUTCOME =====
function getGameOutcome(game, username) {
  if (!game.white || !game.black) return 'unknown';
  
  const isWhite = game.white.username.toLowerCase() === username.toLowerCase();
  const result = game.white.result;
  
  if (result === 'win' && isWhite) return 'win';
  if (result === 'win' && !isWhite) return 'loss';
  if (result === 'checkmated' && isWhite) return 'loss';
  if (result === 'checkmated' && !isWhite) return 'win';
  if (result === 'resigned' && isWhite) return 'loss';
  if (result === 'resigned' && !isWhite) return 'win';
  if (result === 'timeout' && isWhite) return 'loss';
  if (result === 'timeout' && !isWhite) return 'win';
  if (result === 'abandoned' && isWhite) return 'loss';
  if (result === 'abandoned' && !isWhite) return 'win';
  
  // Draws
  if (['agreed', 'stalemate', 'repetition', 'insufficient', 'timevsinsufficient', '50move'].includes(result)) {
    return 'draw';
  }
  
  return 'unknown';
}

// ===== HELPER: GET GAME TERMINATION =====
function getGameTermination(game, username) {
  if (!game.white || !game.black) return 'unknown';
  
  const isWhite = game.white.username.toLowerCase() === username.toLowerCase();
  const result = isWhite ? game.white.result : game.black.result;
  
  return result || 'unknown';
}
