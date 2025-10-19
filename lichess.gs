// ===== LICHESS INTEGRATION =====

function setupLichessSheet() {
  const ss = getEnrichmentSpreadsheet();
  const sheet = ss.getSheetByName('Lichess Analysis') || ss.insertSheet('Lichess Analysis');
  
  if (sheet.getLastRow() === 0) {
    const headers = [
      'Game ID',
      'Lichess URL',
      'Analysis Status',
      'Average Centipawn Loss',
      'Accuracy %',
      'Mistakes',
      'Blunders',
      'Inaccuracies',
      'Date Imported',
      'Date Analyzed'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

// ===== EXPORT GAMES TO LICHESS =====
function exportGamesToLichess() {
  const ui = SpreadsheetApp.getUi();
  
  // Ask for number of games
  const response = ui.prompt(
    'Export to Lichess',
    'How many recent games would you like to export?\n\n' +
    '(This will open Lichess import page with PGN)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const count = parseInt(response.getResponseText());
  if (isNaN(count) || count <= 0) {
    ui.alert('Invalid number');
    return;
  }
  
  const gamesSheet = getGamesSheet();
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) {
    ui.alert('No games found');
    return;
  }
  
  // Get recent games with PGN
  const startRow = Math.max(2, lastRow - count + 1);
  const numRows = lastRow - startRow + 1;
  
  const data = gamesSheet.getRange(startRow, GAMES_COLS.GAME_ID, numRows, GAMES_COLS.PGN).getValues();
  
  const pgns = [];
  for (const row of data) {
    const pgn = row[GAMES_COLS.PGN - 1];
    if (pgn && pgn.trim()) {
      pgns.push(pgn.trim());
    }
  }
  
  if (pgns.length === 0) {
    ui.alert('No PGN data found for selected games');
    return;
  }
  
  // Combine PGNs
  const combinedPGN = pgns.join('\n\n');
  
  // Create temporary HTML to copy PGN
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      textarea { width: 100%; height: 300px; font-family: monospace; font-size: 12px; }
      button { 
        background: #4285f4; 
        color: white; 
        border: none; 
        padding: 10px 20px; 
        font-size: 14px; 
        cursor: pointer; 
        border-radius: 4px;
        margin-top: 10px;
      }
      button:hover { background: #357ae8; }
      .success { color: #0f9d58; margin-top: 10px; display: none; }
    </style>
    <h2>📋 Export to Lichess</h2>
    <p>${pgns.length} games ready to export</p>
    <textarea id="pgnText" readonly>${combinedPGN}</textarea>
    <br>
    <button onclick="copyToClipboard()">Copy PGN to Clipboard</button>
    <button onclick="openLichess()">Open Lichess Import</button>
    <div class="success" id="success">✅ Copied to clipboard!</div>
    
    <script>
      function copyToClipboard() {
        const textarea = document.getElementById('pgnText');
        textarea.select();
        document.execCommand('copy');
        document.getElementById('success').style.display = 'block';
        setTimeout(() => {
          document.getElementById('success').style.display = 'none';
        }, 3000);
      }
      
      function openLichess() {
        window.open('https://lichess.org/paste', '_blank');
      }
    </script>
  `)
    .setWidth(600)
    .setHeight(500);
  
  ui.showModalDialog(htmlOutput, 'Export to Lichess');
}

// ===== IMPORT LICHESS ANALYSIS =====
function importLichessAnalysis() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Import Lichess Analysis',
    'Enter the Lichess study URL or game URL:\n\n' +
    'Example: https://lichess.org/study/abc123',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const url = response.getResponseText().trim();
  
  if (!url.includes('lichess.org')) {
    ui.alert('Invalid Lichess URL');
    return;
  }
  
  ui.alert(
    'Lichess Analysis Import',
    'Automatic analysis import from Lichess API is not yet implemented.\n\n' +
    'For now, you can:\n' +
    '1. Export your games to Lichess\n' +
    '2. Request computer analysis on Lichess\n' +
    '3. Manually record statistics in the Lichess Analysis sheet',
    ui.ButtonSet.OK
  );
}

// ===== MARK GAMES AS IMPORTED TO LICHESS =====
function markGamesAsLichessImported() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Mark as Imported',
    'How many recent games were imported to Lichess?',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const count = parseInt(response.getResponseText());
  if (isNaN(count) || count <= 0) {
    ui.alert('Invalid number');
    return;
  }
  
  const gamesSheet = getGamesSheet();
  const lastRow = gamesSheet.getLastRow();
  
  if (lastRow <= 1) return;
  
  const startRow = Math.max(2, lastRow - count + 1);
  const numRows = lastRow - startRow + 1;
  const now = new Date();
  
  // Update Lichess status
  const updates = [];
  for (let i = 0; i < numRows; i++) {
    updates.push(['imported', now]);
  }
  
  gamesSheet.getRange(startRow, GAMES_COLS.LICHESS_STATUS, numRows, 2).setValues(
    updates.map(u => [u[0], ''])  // Status only, URL stays empty for now
  );
  
  gamesSheet.getRange(startRow, GAMES_COLS.LAST_UPDATED, numRows, 1).setValue(now);
  
  ui.alert(`✅ Marked ${numRows} games as imported to Lichess`);
}

// ===== BATCH OPERATIONS MENU =====
function showLichessMenu() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { color: #4285f4; }
      .option { 
        margin: 15px 0; 
        padding: 15px; 
        background: #f5f5f5; 
        border-radius: 4px;
      }
      .option h3 { margin-top: 0; }
    </style>
    <h2>♟️ Lichess Integration</h2>
    
    <div class="option">
      <h3>1. Export Games to Lichess</h3>
      <p>Select recent games and get their PGN for import to Lichess.</p>
      <p><strong>Menu:</strong> Lichess → Export Games to Lichess</p>
    </div>
    
    <div class="option">
      <h3>2. Mark as Imported</h3>
      <p>After importing to Lichess, mark games as imported in your tracker.</p>
      <p><strong>Menu:</strong> Lichess → Mark as Imported</p>
    </div>
    
    <div class="option">
      <h3>3. View Analysis Sheet</h3>
      <p>The Lichess Analysis sheet stores analysis data for your games.</p>
      <p><strong>Location:</strong> Enrichment Spreadsheet → Lichess Analysis tab</p>
    </div>
    
    <hr>
    <p><em>Note: Automatic analysis import from Lichess API is planned for a future update.</em></p>
  `)
    .setWidth(500)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Lichess Integration Guide');
}
