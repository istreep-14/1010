// ===== POPULATE ARCHIVES =====
function populateArchives() {
  try {
    const archivesSheet = getArchivesSheet();
    
    // Show progress
    SpreadsheetApp.getActiveSpreadsheet().toast('Fetching archives list from Chess.com...', '⏳', -1);
    
    // Fetch all archives from Chess.com
    const url = `https://api.chess.com/pub/player/${CONFIG.USERNAME}/games/archives`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    const archives = data.archives;
    
    if (!archives || !archives.length) {
      SpreadsheetApp.getUi().alert('No archives found for user: ' + CONFIG.USERNAME);
      return;
    }
    
    // Prepare rows for all archives
    const rows = [];
    const now = new Date();
    
    for (const archiveUrl of archives) {
      // Extract year and month from URL
      // URL format: https://api.chess.com/pub/player/username/games/YYYY/MM
      const parts = archiveUrl.split('/');
      const year = parts[parts.length - 2];
      const month = parts[parts.length - 1];
      
      rows.push([
        archiveUrl,           // Archive
        year,                 // Year
        month,                // Month
        'pending',            // Status
        0,                    // Game Count
        now,                  // Date Created
        null,                 // Last Checked
        null,                 // Last Fetched
        '',                   // ETag
        'Games',              // Data Location
        ''                    // Notes
      ]);
    }
    
    // Write all archives to sheet
    const startRow = archivesSheet.getLastRow() + 1;
    archivesSheet.getRange(startRow, 1, rows.length, ARCHIVE_COLS.NOTES).setValues(rows);
    
    // Update Config
    setConfig('Total Archives', archives.length);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `✅ Added ${archives.length} archives!`, 
      '✅', 
      5
    );
    
    Logger.log(`Populated ${archives.length} archives`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error populating archives: ' + error.message);
    Logger.log(error);
  }
}
