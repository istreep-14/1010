// ===== SETUP GAMES SHEET =====
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES) || ss.insertSheet(SHEETS.GAMES);
  
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  }
  
  const headers = [
    'Game ID', 'Type', 'Game URL',
    'Start Date/Time', 'Start Date', 'Start Time', 'Start (s)',
    'End Date/Time', 'End Date', 'End Time', 'End (s)', 'End Serial', 'Archive',
    'Rules', 'Live', 'Time Class', 'Format', 'Rated', 'Time Control', 'Base', 'Inc', 'Corr', 'Duration', 'Duration (s)',
    'Color', 'Opponent', 'My Rating', 'Opp Rating', 'Rating Before', 'Rating Δ',
    'Outcome', 'Termination',
    'ECO', 'ECO URL',
    'Opening Name', 'Opening Slug', 'Opening Family', 'Opening Base',
    'Variation 1', 'Variation 2', 'Variation 3', 'Variation 4', 'Variation 5', 'Variation 6',
    'Extra Moves',
    'Moves', 'TCN', 'Clocks', 'Ratings Ledger'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  sheet.getRange('D:D').setNumberFormat('@');
  sheet.getRange('H:H').setNumberFormat('@');
  sheet.getRange('E:E').setNumberFormat('M/D/YY');
  sheet.getRange('I:I').setNumberFormat('M/D/YY');
  sheet.getRange('F:F').setNumberFormat('h:mm AM/PM');
  sheet.getRange('J:J').setNumberFormat('h:mm AM/PM');
  sheet.getRange('W:W').setNumberFormat('[h]:mm:ss');
  sheet.getRange('X:X').setNumberFormat('0');
  sheet.getRange('AA:AD').setNumberFormat('0');
  sheet.getRange('AT:AT').setNumberFormat('0');
  sheet.getRange('M:M').setNumberFormat('@');
  sheet.getRange('AG:AG').setNumberFormat('@');
  sheet.getRange('AH:AH').setNumberFormat('@');
  sheet.getRange('AR:AR').setNumberFormat('@');
  
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidths(5, 2, 90);
  sheet.setColumnWidth(8, 180);
  sheet.setColumnWidths(9, 2, 90);
  sheet.setColumnWidth(13, 90);
  sheet.setColumnWidth(14, 100);
  sheet.setColumnWidth(17, 80);
  sheet.setColumnWidth(18, 60);
  sheet.setColumnWidth(19, 100);
  sheet.setColumnWidth(23, 90);
  sheet.setColumnWidth(24, 90);
  sheet.setColumnWidth(32, 125);
  sheet.setColumnWidth(33, 65);
  sheet.setColumnWidth(34, 90);
  sheet.setColumnWidth(35, 150);
  sheet.setColumnWidth(36, 150);
  sheet.setColumnWidth(37, 200);
  sheet.setColumnWidth(38, 120);
  sheet.setColumnWidth(39, 120);
  sheet.setColumnWidth(40, 120);
  sheet.setColumnWidth(41, 120);
  sheet.setColumnWidth(42, 120);
  sheet.setColumnWidth(43, 120);
  sheet.setColumnWidth(44, 200);

  const maxRows = sheet.getMaxRows();
  const maxCols = headers.length;
  sheet.getRange(1, 1, maxRows, maxCols).setFontFamily('Montserrat');
  sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment('center');

  sheet.setHiddenGridlines(true);

  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  const banding = dataRange.getBandings()[0];
  if (banding) {
    banding.remove();
  }
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  sheet.getRange(1, 1, 1, headers.length).setBorder(null, null, true, null, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  try {
    const lastCol = sheet.getMaxColumns();
    if (lastCol > 0) {
      sheet.getRange(1, 1, 1, lastCol).shiftColumnGroupDepth(-1);
    }
  } catch (e) {
    // No groups exist
  }

  sheet.getRange('A1:B1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('A:B'));

  sheet.getRange('D1:H1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('D:H'));

  sheet.getRange('K1:P1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('K:P'));

  sheet.getRange('T1:W1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('T:W'));

  sheet.getRange('AL1:AS1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('AL:AS'));

  sheet.getRange('AI1:AJ1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('AI:AJ'));

  sheet.getRange('AU1:AV1').shiftColumnGroupDepth(1);
  sheet.hideColumn(sheet.getRange('AU:AV'));

  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  ss.setNamedRange('GamesData', sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length));
  ss.setNamedRange('GameIDs', sheet.getRange('A2:A'));
  ss.setNamedRange('Outcomes', sheet.getRange('AE2:AE'));
  ss.setNamedRange('MyRatings', sheet.getRange('AA2:AA'));
  ss.setNamedRange('Opponents', sheet.getRange('Z2:Z'));
  ss.setNamedRange('OpeningNames', sheet.getRange('AI2:AI'));

  sheet.clearConditionalFormatRules();
  const newRules = [];

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('win')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('loss')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('draw')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('AE2:AE')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('AD2:AD')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('AD2:AD')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('bullet')
    .setBackground('#c9daf8')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('blitz')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('rapid')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('daily')
    .setBackground('#d9d2e9')
    .setRanges([sheet.getRange('Q2:Q')])
    .build());

  sheet.setConditionalFormatRules(newRules);

  SpreadsheetApp.getUi().alert('✅ Games sheet setup complete!');
}

// ===== SETUP CONTROL SPREADSHEET =====
function setupControlSpreadsheet() {
  try {
    const controlSS = SpreadsheetApp.create('Chess Control');
    const controlId = controlSS.getId();
    
    CONFIG.CONTROL_SPREADSHEET_ID = controlId;
    
    setupArchivesTab();
    setupGameRegistryTab();
    setupRatingsTab();
    setupConfigTab();
    
    const defaultSheet = controlSS.getSheetByName('Sheet1');
    if (defaultSheet) {
      controlSS.deleteSheet(defaultSheet);
    }
    
    SpreadsheetApp.getUi().alert(
      '✅ Control Spreadsheet Created!\n\n' +
      'Spreadsheet ID: ' + controlId + '\n' +
      'URL: ' + controlSS.getUrl() + '\n\n' +
      'Please add this ID to CONFIG.CONTROL_SPREADSHEET_ID in your script.\n\n' +
      'Next step: Run populateArchives() to fetch all available archives from Chess.com.'
    );
    
    Logger.log('✅ Control spreadsheet created: ' + controlId);
    
    return controlId;
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error creating Control spreadsheet: ' + error.message);
    Logger.log(error);
  }
}

// ===== SETUP ARCHIVES TAB =====
function setupArchivesTab() {
  const controlSS = getControlSpreadsheet();
  const sheet = controlSS.getSheetByName(SHEETS.CONTROL.ARCHIVES) || controlSS.insertSheet(SHEETS.CONTROL.ARCHIVES);
  
  if (sheet.getLastRow() > 0) {
    sheet.clearContents().clearFormats();
  }
  
  const headers = [
    'Archive', 'Year', 'Month', 'Status', 'Game Count',
    'Date Created', 'Last Checked', 'Last Fetched', 'ETag',
    'Data Location', 'Notes'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  sheet.getRange('F:F').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('G:G').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('H:H').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('E:E').setNumberFormat('0');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('I:I').setNumberFormat('@');
  sheet.getRange('J:J').setNumberFormat('@');
  sheet.getRange('K:K').setNumberFormat('@');
  
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 250);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(11, 200);
  
  const maxRows = sheet.getMaxRows();
  const maxCols = headers.length;
  sheet.getRange(1, 1, maxRows, maxCols).setFontFamily('Montserrat');
  sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment('center');
  
  sheet.setHiddenGridlines(true);
  
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  const banding = dataRange.getBandings()[0];
  if (banding) banding.remove();
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  sheet.getRange(1, 1, 1, headers.length).setBorder(
    null, null, true, null, null, null,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  sheet.clearConditionalFormatRules();
  const rules = [];
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('complete')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('D2:D')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('pending')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('D2:D')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('error')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('D2:D')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('fetching')
    .setBackground('#c9daf8')
    .setRanges([sheet.getRange('D2:D')])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('✅ Archives tab setup complete');
}

// ===== SETUP GAME REGISTRY TAB =====
function setupGameRegistryTab() {
  const controlSS = getControlSpreadsheet();
  const sheet = controlSS.getSheetByName(SHEETS.CONTROL.REGISTRY) || controlSS.insertSheet(SHEETS.CONTROL.REGISTRY);
  
  if (sheet.getLastRow() > 0) {
    sheet.clearContents().clearFormats();
  }
  
  const headers = [
    'Game ID', 'Archive', 'Date', 'Format', 'Date Added',
    'Callback Status', 'Callback Date',
    'Future Enrich 1 Status', 'Future Enrich 1 Date',
    'Data Location'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  sheet.getRange('C:C').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('E:E').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('G:G').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('I:I').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('J:J').setNumberFormat('@');
  
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 150);
  sheet.setColumnWidth(10, 120);
  
  const maxRows = sheet.getMaxRows();
  const maxCols = headers.length;
  sheet.getRange(1, 1, maxRows, maxCols).setFontFamily('Montserrat');
  sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment('center');
  
  sheet.setHiddenGridlines(true);
  
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  const banding = dataRange.getBandings()[0];
  if (banding) banding.remove();
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  sheet.getRange(1, 1, 1, headers.length).setBorder(
    null, null, true, null, null, null,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  sheet.clearConditionalFormatRules();
  const rules = [];
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('fetched')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('F2:F')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('invalid')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('F2:F')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('error')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('F2:F')])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('✅ Game Registry tab setup complete');
}

// ===== SETUP RATINGS TAB =====
function setupRatingsTab() {
  const controlSS = getControlSpreadsheet();
  const sheet = controlSS.getSheetByName(SHEETS.CONTROL.RATINGS) || controlSS.insertSheet(SHEETS.CONTROL.RATINGS);
  
  if (sheet.getLastRow() > 0) {
    sheet.clearContents().clearFormats();
  }
  
  const headers = [
    'Game ID', 'Game URL', 'Archive', 'Date', 'Format',
    'My Rating', 'My Rating Last', 'My Rating Δ Last',
    'Opp Rating', 'Opp Rating Δ Last', 'Opp Rating Last',
    'Callback Status', 'Callback Date',
    'My Rating Pregame Callback', 'My Rating Δ Callback',
    'Opp Rating Pregame Callback', 'Opp Rating Δ Callback',
    'My Rating Pregame Effective', 'My Rating Δ Effective',
    'Opp Rating Pregame Effective', 'Opp Rating Δ Effective'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  sheet.getRange('D:D').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('M:M').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('@');
  
  sheet.getRange('F:F').setNumberFormat('0');
  sheet.getRange('G:G').setNumberFormat('0');
  sheet.getRange('H:H').setNumberFormat('0');
  sheet.getRange('I:I').setNumberFormat('0');
  sheet.getRange('J:J').setNumberFormat('0');
  sheet.getRange('K:K').setNumberFormat('0');
  sheet.getRange('N:N').setNumberFormat('0');
  sheet.getRange('O:O').setNumberFormat('0');
  sheet.getRange('P:P').setNumberFormat('0');
  sheet.getRange('Q:Q').setNumberFormat('0');
  sheet.getRange('R:R').setNumberFormat('0');
  sheet.getRange('S:S').setNumberFormat('0');
  sheet.getRange('T:T').setNumberFormat('0');
  sheet.getRange('U:U').setNumberFormat('0');
  
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 90);
  sheet.setColumnWidth(10, 140);
  sheet.setColumnWidth(11, 130);
  sheet.setColumnWidth(12, 120);
  sheet.setColumnWidth(13, 130);
  sheet.setColumnWidth(14, 180);
  sheet.setColumnWidth(15, 160);
  sheet.setColumnWidth(16, 190);
  sheet.setColumnWidth(17, 170);
  sheet.setColumnWidth(18, 180);
  sheet.setColumnWidth(19, 160);
  sheet.setColumnWidth(20, 190);
  sheet.setColumnWidth(21, 170);
  
  const maxRows = sheet.getMaxRows();
  const maxCols = headers.length;
  sheet.getRange(1, 1, maxRows, maxCols).setFontFamily('Montserrat');
  sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment('center');
  
  sheet.setHiddenGridlines(true);
  
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  const banding = dataRange.getBandings()[0];
  if (banding) banding.remove();
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  sheet.getRange(1, 1, 1, headers.length).setBorder(
    null, null, true, null, null, null,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  try {
    sheet.getRange('B1:B1').shiftColumnGroupDepth(1);
    sheet.hideColumn(sheet.getRange('B:B'));
  } catch (e) {
    Logger.log('Could not create column group: ' + e);
  }
  
  sheet.clearConditionalFormatRules();
  const rules = [];
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('fetched')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('L2:L')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('invalid')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('L2:L')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('error')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('L2:L')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('H2:H')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('H2:H')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('O2:O')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('O2:O')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor('#38761d')
    .setBold(true)
    .setRanges([sheet.getRange('S2:S')])
    .build());
  
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#cc0000')
    .setBold(true)
    .setRanges([sheet.getRange('S2:S')])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('✅ Ratings tab setup complete');
}

// ===== SETUP CONFIG TAB =====
function setupConfigTab() {
  const controlSS = getControlSpreadsheet();
  const sheet = controlSS.getSheetByName(SHEETS.CONTROL.CONFIG) || controlSS.insertSheet(SHEETS.CONTROL.CONFIG);
  
  if (sheet.getLastRow() > 0) {
    sheet.clearContents().clearFormats();
  }
  
  const headers = ['Setting', 'Value'];
  sheet.getRange(1, 1, 1, 2).setValues([headers]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  sheet.getRange(1, 1, 1, 2).setBackground('#4285f4').setFontColor('#ffffff');
  
  const settings = [
    ['Username', CONFIG.USERNAME],
    ['Months to Fetch', CONFIG.MONTHS_TO_FETCH],
    ['Callback Batch Size', 50],
    ['Duplicate Check Rows', 200],
    ['Game Index Cache Size', 500],
    [''],
    ['=== PROGRESS ===', ''],
    ['Last Registry Row Processed', 1],
    ['Total Games', 0],
    ['Callbacks Complete', 0],
    ['Callbacks Pending', 0],
    ['Callbacks Invalid', 0],
    ['Callbacks Error', 0],
    [''],
    ['=== TIMESTAMPS ===', ''],
    ['Last Full Fetch Date', ''],
    ['Last Callback Run Date', ''],
    ['Last Validation Run Date', ''],
    [''],
    ['=== SPREADSHEET IDS ===', ''],
    ['Control Spreadsheet ID', controlSS.getId()],
    ['Games Spreadsheet ID', SpreadsheetApp.getActiveSpreadsheet().getId()],
    ['Callback Data Spreadsheet ID', ''],
    [''],
    ['=== ERROR LOG ===', ''],
    ['Recent Errors', '']
  ];
  
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  
  sheet.getRange('B16:B18').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('@');
  
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);
  
  sheet.getRange('A:A').setFontFamily('Montserrat').setHorizontalAlignment('left');
  sheet.getRange('B:B').setFontFamily('Montserrat').setHorizontalAlignment('left');
  sheet.getRange('A1:B1').setHorizontalAlignment('center');
  
  sheet.getRange('A7').setFontWeight('bold').setBackground('#e0e0e0');
  sheet.getRange('A14').setFontWeight('bold').setBackground('#e0e0e0');
  sheet.getRange('A19').setFontWeight('bold').setBackground('#e0e0e0');
  sheet.getRange('A24').setFontWeight('bold').setBackground('#e0e0e0');
  
  sheet.setHiddenGridlines(true);
  
  sheet.getRange(1, 1, 1, 2).setBorder(
    null, null, true, null, null, null,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  Logger.log('✅ Config tab setup complete');
}

// ===== SETUP CALLBACK DATA SPREADSHEET =====
function setupCallbackDataSpreadsheet() {
  const ss = SpreadsheetApp.create('Chess Callback Data');
  const sheet = ss.getSheets()[0];
  sheet.setName('Callback Data');
  
  const headers = [
    'Game ID', 'Game URL', 'Callback URL', 'End Time', 'My Color', 'Time Class',
    'My Rating', 'Opp Rating', 'My Rating Change', 'Opp Rating Change',
    'My Rating Before', 'Opp Rating Before', 'Base Time', 'Time Increment',
    'Move Timestamps', 'My Username', 'My Country', 'My Membership',
    'My Member Since', 'My Default Tab', 'My Post Move Action', 'My Location',
    'Opp Username', 'Opp Country', 'Opp Membership', 'Opp Member Since',
    'Opp Default Tab', 'Opp Post Move Action', 'Opp Location', 'Date Fetched'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4').setFontColor('#ffffff');
  
  sheet.getRange('D:D').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('S:S').setNumberFormat('M/D/YY');
  sheet.getRange('Z:Z').setNumberFormat('M/D/YY');
  sheet.getRange('AD:AD').setNumberFormat('M/D/YY h:mm');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('B:B').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('@');
  sheet.getRange('O:O').setNumberFormat('@');
  sheet.getRange('G:L').setNumberFormat('0');
  sheet.getRange('M:M').setNumberFormat('0');
  sheet.getRange('N:N').setNumberFormat('0');
  
  sheet.setFrozenRows(1);
  
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(9, 130);
  sheet.setColumnWidth(10, 140);
  sheet.setColumnWidth(11, 130);
  sheet.setColumnWidth(12, 140);
  sheet.setColumnWidth(13, 90);
  sheet.setColumnWidth(14, 110);
  sheet.setColumnWidth(15, 150);
  sheet.setColumnWidth(16, 120);
  sheet.setColumnWidth(17, 100);
  sheet.setColumnWidth(18, 120);
  sheet.setColumnWidth(19, 120);
  sheet.setColumnWidth(20, 120);
  sheet.setColumnWidth(21, 150);
  sheet.setColumnWidth(22, 120);
  sheet.setColumnWidth(23, 120);
  sheet.setColumnWidth(24, 100);
  sheet.setColumnWidth(25, 120);
  sheet.setColumnWidth(26, 120);
  sheet.setColumnWidth(27, 120);
  sheet.setColumnWidth(28, 150);
  sheet.setColumnWidth(29, 120);
  sheet.setColumnWidth(30, 130);
  
  const maxRows = sheet.getMaxRows();
  const maxCols = headers.length;
  sheet.getRange(1, 1, maxRows, maxCols).setFontFamily('Montserrat');
  sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment('center');
  
  sheet.setHiddenGridlines(true);
  
  const dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  const banding = dataRange.getBandings()[0];
  if (banding) banding.remove();
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  
  sheet.getRange(1, 1, 1, headers.length).setBorder(
    null, null, true, null, null, null,
    '#cccccc', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  try {
    sheet.getRange('B1:C1').shiftColumnGroupDepth(1);
    sheet.hideColumn(sheet.getRange('B:C'));
  } catch (e) {
    Logger.log('Could not create column group: ' + e);
  }
  
  SpreadsheetApp.getUi().alert(
    '✅ Callback Data Spreadsheet Created!\n\n' +
    'Spreadsheet ID: ' + ss.getId() + '\n' +
    'URL: ' + ss.getUrl() + '\n\n' +
    'Please save this ID to CONFIG.CALLBACK_DATA_SPREADSHEET_ID in your main script.'
  );
  
  Logger.log('✅ Callback Data spreadsheet created: ' + ss.getId());
  
  return ss.getId();
}
