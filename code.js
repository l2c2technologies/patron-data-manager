/**
 * @OnlyCurrentDoc
 *
 * L2C2 Patron Data Manager for cleanup, validation and transformation of patron data for Koha ILS.
 *
 * @author      Indranil Das Gupta <indradg@l2c2.co.in>
 * @copyright   (c) 2023 - 2025 L2C2 Technologies
 * @version     3.3
 * @license     AGPL v3+
 */

/**
 * Creates the main menu in the Google Sheet UI when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('L2C2 Patron Data Manager')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Cleaning Tools')
          .addItem('Remove Line Breaks in Column', 'removeLineBreaksAndExtraSpaces')
          .addItem('Advanced Cleanup in Range', 'advancedTextCleanup'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Structural Tools')
          .addItem('Add Column with Preset Value', 'addNewColumnWithPresetValue')
          .addItem('Replicate Column', 'replicateColumn')
          .addItem('Rename Column Header', 'renameColumnHeader')
          .addItem('Delete Column', 'deleteColumn'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Transformation Tools')
          .addItem('Conditional Population', 'conditionalPopulation'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Validation Tools')
          .addItem('Find & Handle Duplicates', 'findAndHandleDuplicates')
          .addItem('Validate & Clean Mobile Numbers', 'validateMobileNumbers')
          .addItem('Validate & Clean Emails (Syntax/Domain)', 'validateEmails')
          .addItem('Validate & Clean Aadhaar Numbers', 'validateAadhaarNumbers')
          .addItem('Validate & Format Dates', 'validateAndFormatDates'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Documentation Tools')
          .addItem('Generate Koha Field Map', 'mapKohaFields'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Export Tools')
          .addItem('Export Filtered Data as CSV', 'exportFilteredDataAsCsv'))
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Help')
          .addItem('About this Tool', 'showHelp'))
      .addToUi();
}

// --- HELP FUNCTION ---

/**
 * Displays a dialog box with descriptions of all available tools.
 */
function showHelp() {
  var ui = SpreadsheetApp.getUi();
  var message = 'L2C2 Patron Data Manager - Function Guide:\n\n' +
    'This tool helps clean, manage, and prepare your patron data for import into the Koha ILS. All actions that modify the sheet are automatically recorded in a sheet named "Action Log" for full traceability.\n\n' +
    '--- CLEANING TOOLS ---\n' +
    '• Remove Line Breaks in Column: Cleans an entire column by removing line breaks and extra spaces.\n' +
    '• Advanced Cleanup in Range: Removes extra spaces, handles spacing around punctuation, and removes line breaks within a selected cell range. Protects email addresses from being altered.\n\n' +
    '--- STRUCTURAL TOOLS ---\n' +
    '• Add Column with Preset Value: Adds a new column with a default value, asking you where to place it.\n' +
    '• Replicate Column: Duplicates a column and lets you choose where to place the new copy.\n' +
    '• Rename Column Header: Changes the title in the first row of a specified column.\n' +
    '• Delete Column: Permanently removes a specified column after confirmation.\n\n' +
    '--- TRANSFORMATION TOOLS ---\n' +
    '• Conditional Population: Fills a target column with a specific value based on a condition in another column.\n\n' +
    '--- VALIDATION TOOLS ---\n' +
    '• Find & Handle Duplicates: Scans for duplicates and lets you choose how to handle them (interactively, remove row, or clear cell).\n' +
    '• Validate & Clean Mobile Numbers: Formats valid 10-digit Indian mobile numbers and removes any that are invalid.\n' +
    '• Validate & Clean Emails: Checks email syntax and if the domain can receive mail. Removes invalid emails.\n' +
    '• Validate & Clean Aadhaar Numbers: Validates 12-digit Aadhaar numbers using the Verhoeff algorithm. Removes/logs invalid numbers.\n' +
    '• Validate & Format Dates: Formats various date formats into YYYY-MM-DD. It will ask for help if a date has an ambiguous two-digit year.\n\n' +
    '--- DOCUMENTATION TOOLS ---\n' +
    '• Generate Koha Field Map: Interactively maps your sheet\'s columns to the standard Koha patron import fields and notes any custom re-purposing.\n\n' +
    '--- EXPORT TOOLS ---\n' +
    '• Export Filtered Data as CSV: Creates a CSV file in your Google Drive containing only the rows that match your filter criteria.';
  
  ui.alert('About this Tool', message, ui.ButtonSet.OK);
}


// --- UNIVERSAL LOGGING FUNCTION ---

/**
 * Logs any action to a dedicated "Action Log" sheet with a dedicated cell reference column.
 * @param {string} actionType The category of action.
 * @param {string} target A general description of the target.
 * @param {string} details A description of the change.
 * @param {string} [cellReference] Optional. The specific cell/range/column A1 notation.
 */
function logAction(actionType, target, details, cellReference) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Action Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Action Log", 0);
    logSheet.appendRow(["Timestamp", "Action Type", "Target", "Cell Reference", "Details"]);
    logSheet.getRange("A1:E1").setFontWeight("bold");
    logSheet.setFrozenRows(1);
  } else {
    // Retroactively add the new column if the log sheet already exists from an older version
    if (logSheet.getRange("D1").getValue() !== "Cell Reference") {
      logSheet.insertColumnAfter(3);
      logSheet.getRange("D1").setValue("Cell Reference").setFontWeight("bold");
    }
  }
  var sheetName = SpreadsheetApp.getActiveSheet().getName();
  logSheet.appendRow([new Date(), actionType, target, sheetName + "!" + (cellReference || ''), details]);
}


// --- DOCUMENTATION FUNCTIONS ---

/**
 * Interactively creates a sheet that maps current column headers to standard Koha patron fields.
 */
function mapKohaFields() {
  const kohaFields = [
    'cardnumber','surname','firstname','middle_name','title','othernames','initials','pronouns','streetnumber','streettype',
    'address','address2','city','state','zipcode','country','email','phone','mobile','fax','emailpro','phonepro','B_streetnumber',
    'B_streettype','B_address','B_address2','B_city','B_state','B_zipcode','B_country','B_email','B_phone','dateofbirth',
    'branchcode','categorycode','dateenrolled','dateexpiry','password_expiration_date','date_renewed','gonenoaddress','lost',
    'debarred','debarredcomment','contactname','contactfirstname','contacttitle','borrowernotes','relationship','sex','password',
    'secret','auth_method','flags','userid','opacnote','contactnote','sort1','sort2','altcontactfirstname','altcontactsurname',
    'altcontactaddress1','altcontactaddress2','altcontactaddress3','altcontactstate','altcontactzipcode','altcontactcountry',
    'altcontactphone','smsalertnumber','sms_provider_id','privacy','privacy_guarantor_fines','privacy_guarantor_checkouts',
    'checkprevcheckout','updated_on','lastseen','lang','login_attempts','overdrive_auth_token','anonymized',
    'autorenew_checkouts','primary_contact_method','protected','patron_attributes','guarantor_relationship','guarantor_id'
  ];

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var mappingSheet = ss.getSheetByName("Field Mapping");
  if (mappingSheet) {
    var confirm = ui.alert('Sheet "Field Mapping" already exists.', 'Do you want to overwrite it?', ui.ButtonSet.YES_NO);
    if (confirm === ui.Button.YES) {
      mappingSheet.clear();
    } else {
      ui.alert('Operation cancelled.');
      return;
    }
  } else {
    mappingSheet = ss.insertSheet("Field Mapping", 0);
  }

  mappingSheet.appendRow(["Source Column Header", "Mapped Koha Field", "Notes / Purpose"]);
  mappingSheet.getRange("A1:C1").setFontWeight("bold");
  mappingSheet.setFrozenRows(1);

  for (var i = 0; i < headers.length; i++) {
    var sourceHeader = headers[i];
    if (!sourceHeader) continue;

    var mappedField = '';
    while (true) {
      var result = ui.prompt('Map Source Field: "' + sourceHeader + '"', 
                              'Enter the corresponding Koha field name below. Type "SKIP" to ignore this field or "CANCEL" to stop.', 
                              ui.ButtonSet.OK_CANCEL);

      if (result.getSelectedButton() !== ui.Button.OK) {
        ui.alert('Mapping cancelled.');
        return;
      }
      
      mappedField = result.getResponseText().trim();
      
      if (mappedField.toLowerCase() === 'cancel') {
        ui.alert('Mapping cancelled.');
        return;
      }
      if (mappedField.toLowerCase() === 'skip' || kohaFields.indexOf(mappedField.toLowerCase()) !== -1) {
        break;
      } else {
        ui.alert('Invalid Koha Field', '"' + mappedField + '" is not a valid Koha patron field. Please try again.', ui.ButtonSet.OK);
      }
    }
    
    if (mappedField.toLowerCase() === 'skip') {
      mappingSheet.appendRow([sourceHeader, "--- SKIPPED ---", ""]);
      continue;
    }

    var notes = "";
    var repurposeResult = ui.alert('Is the Koha field "' + mappedField + '" being re-purposed?',
                                   '(For example, using the "country" field to store an Aadhaar number).',
                                   ui.ButtonSet.YES_NO);
    
    if (repurposeResult === ui.Button.YES) {
      var notesResult = ui.prompt('Describe the purpose of this field:', ui.ButtonSet.OK_CANCEL);
      if (notesResult.getSelectedButton() === ui.Button.OK) {
        notes = notesResult.getResponseText();
      }
    }
    
    mappingSheet.appendRow([sourceHeader, mappedField, notes]);
  }
  
  logAction("Documentation", "Koha Field Mapping", "Generated or updated the Koha field mapping documentation.", "Field Mapping");
  ui.alert('Field Mapping Complete', 'The "Field Mapping" sheet has been created/updated successfully.', ui.ButtonSet.OK);
}


// --- DATA CLEANING FUNCTIONS ---

function removeLineBreaksAndExtraSpaces() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter the column letter to clean up:', ui.ButtonSet.OK_CANCEL);
  var columnLetter = result.getResponseText().toUpperCase().trim();
  if (result.getSelectedButton() == ui.Button.OK && columnLetter) {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      var range = sheet.getRange(columnLetter + ':' + columnLetter);
      var values = range.getValues();
      var changesMade = 0;
      for (var i = 1; i < values.length; i++) { // Start from row 2
        if (typeof values[i][0] == 'string') {
          var originalValue = values[i][0];
          var cleanedText = originalValue.replace(/(\r\n|\n|\r)/gm, " ").replace(/\s{2,}/g, " ");
          if (originalValue !== cleanedText) {
            values[i][0] = cleanedText;
            changesMade++;
          }
        }
      }
      if (changesMade > 0) {
        range.setValues(values);
        logAction("Data Cleanup", "Column Cleanup", "Removed line breaks/spaces from " + changesMade + " cells.", columnLetter);
        ui.alert('Cleanup Complete', 'Cleaned ' + changesMade + ' cells in column ' + columnLetter + '.', ui.ButtonSet.OK);
      } else {
        ui.alert('No changes were needed in column ' + columnLetter + '.');
      }
    } catch (e) {
      ui.alert('Error: Please enter a valid column letter.');
    }
  }
}

function advancedTextCleanup() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter the cell range to clean up (e.g., A2:C50):', ui.ButtonSet.OK_CANCEL);
  var rangeString = result.getResponseText().trim();
  if (result.getSelectedButton() == ui.Button.OK && rangeString) {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      var range = sheet.getRange(rangeString);
      var values = range.getValues();
      var changesMade = 0;
      for (var i = 0; i < values.length; i++) {
        for (var j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] == 'string') {
            var originalValue = values[i][j];
            var cleanedText = originalValue;
            
            if (cleanedText.indexOf('@') === -1) {
              var puncChars = ",.?!:;)\\}\\]\\/"; 
              cleanedText = cleanedText.replace(new RegExp("\\s+([" + puncChars + "])", "g"), '$1');
              cleanedText = cleanedText.replace(new RegExp("([" + puncChars + "])(?!\\s|[" + puncChars + "]|$)", "g"), '$1 ');
            }
            
            cleanedText = cleanedText.replace(/(\r\n|\n|\r)/gm, " ");
            cleanedText = cleanedText.replace(/\s{2,}/g, ' ');
            cleanedText = cleanedText.trim();
            
            if (originalValue !== cleanedText) {
              values[i][j] = cleanedText;
              changesMade++;
            }
          }
        }
      }
      if (changesMade > 0) {
        range.setValues(values);
        logAction("Data Cleanup", "Range Cleanup", "Performed advanced cleanup on " + changesMade + " cells.", rangeString);
        ui.alert('Cleanup Complete', 'Cleaned ' + changesMade + ' cells in the range ' + rangeString + '.', ui.ButtonSet.OK);
      } else {
        ui.alert('No changes were needed in the specified range.');
      }
    } catch (e) {
      ui.alert('Error: Please enter a valid range.');
    }
  }
}

// --- STRUCTURAL MANIPULATION FUNCTIONS ---

function addNewColumnWithPresetValue() {
  var ui = SpreadsheetApp.getUi();
  var headerResult = ui.prompt('Step 1 of 4: Header', 'Enter the header for the new column:', ui.ButtonSet.OK_CANCEL);
  var newHeaderText = headerResult.getResponseText().trim();
  if (headerResult.getSelectedButton() !== ui.Button.OK || !newHeaderText) return;

  var valueResult = ui.prompt('Step 2 of 4: Preset Value', 'Enter the value to fill this new column with:', ui.ButtonSet.OK_CANCEL);
  var presetValue = valueResult.getResponseText();
  if (valueResult.getSelectedButton() !== ui.Button.OK) return;

  var refColResult = ui.prompt('Step 3 of 4: Reference Column', 'Enter the column letter to place the new column next to (e.g., C):', ui.ButtonSet.OK_CANCEL);
  var refColLetter = refColResult.getResponseText().toUpperCase().trim();
  if (refColResult.getSelectedButton() !== ui.Button.OK || !refColLetter) return;

  var html = HtmlService.createHtmlOutput(
      '<p>Place new column relative to column ' + refColLetter + '?</p>' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).finishAddingColumn(true);">Before</button> ' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).finishAddingColumn(false);">After</button>')
      .setWidth(300).setHeight(100);

  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('newHeaderText', newHeaderText);
  scriptProperties.setProperty('presetValue', presetValue);
  scriptProperties.setProperty('refColLetter', refColLetter);
  
  ui.showModalDialog(html, 'Step 4 of 4: Position');
}

function finishAddingColumn(insertBefore) {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties();
  var newHeaderText = scriptProperties.getProperty('newHeaderText');
  var presetValue = scriptProperties.getProperty('presetValue');
  var refColLetter = scriptProperties.getProperty('refColLetter');

  scriptProperties.deleteProperty('newHeaderText');
  scriptProperties.deleteProperty('presetValue');
  scriptProperties.deleteProperty('refColLetter');
  
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    var refColIndex = sheet.getRange(refColLetter + "1").getColumn();
    
    var newColIndex;
    if (insertBefore) {
      sheet.insertColumnBefore(refColIndex);
      newColIndex = refColIndex;
    } else {
      sheet.insertColumnAfter(refColIndex);
      newColIndex = refColIndex + 1;
    }

    sheet.getRange(1, newColIndex).setValue(newHeaderText);
    if (lastRow > 1) {
      var values = Array(lastRow - 1).fill([presetValue]);
      sheet.getRange(2, newColIndex, values.length, 1).setValues(values);
    }
    
    var newColLetter = sheet.getRange(1, newColIndex).getA1Notation().slice(0, -1);
    logAction("Structural Change", "Add Column", "Added new column with header '" + newHeaderText + "' and populated " + (lastRow > 1 ? lastRow - 1 : 0) + " rows with value '" + presetValue + "'.", newColLetter);
    ui.alert('Success!', 'The new column "' + newHeaderText + '" has been added.', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('An error occurred. Please ensure the reference column letter is valid.');
  }
}

function replicateColumn() {
  var ui = SpreadsheetApp.getUi();
  var sourceResult = ui.prompt('Step 1 of 4: Source Column', 'Enter the letter of the column to replicate:', ui.ButtonSet.OK_CANCEL);
  var sourceColumnLetter = sourceResult.getResponseText().toUpperCase().trim();
  if (sourceResult.getSelectedButton() !== ui.Button.OK || !sourceColumnLetter) return;

  var headerResult = ui.prompt('Step 2 of 4: New Header', 'Enter the header for the new (replicated) column:', ui.ButtonSet.OK_CANCEL);
  var newHeaderText = headerResult.getResponseText();
  if (headerResult.getSelectedButton() !== ui.Button.OK) return;

  var refColResult = ui.prompt('Step 3 of 4: Reference Column', 'Enter the column letter to place the new column next to (e.g., C):', ui.ButtonSet.OK_CANCEL);
  var refColLetter = refColResult.getResponseText().toUpperCase().trim();
  if (refColResult.getSelectedButton() !== ui.Button.OK || !refColLetter) return;

  var html = HtmlService.createHtmlOutput(
      '<p>Place new column relative to column ' + refColLetter + '?</p>' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).finishReplicatingColumn(true);">Before</button> ' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).finishReplicatingColumn(false);">After</button>')
      .setWidth(300).setHeight(100);
      
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('sourceColumnLetter', sourceColumnLetter);
  scriptProperties.setProperty('newHeaderText', newHeaderText);
  scriptProperties.setProperty('refColLetter', refColLetter);
  
  ui.showModalDialog(html, 'Step 4 of 4: Position');
}

function finishReplicatingColumn(insertBefore) {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties();
  var sourceColumnLetter = scriptProperties.getProperty('sourceColumnLetter');
  var newHeaderText = scriptProperties.getProperty('newHeaderText');
  var refColLetter = scriptProperties.getProperty('refColLetter');
  
  scriptProperties.deleteProperty('sourceColumnLetter');
  scriptProperties.deleteProperty('newHeaderText');
  scriptProperties.deleteProperty('refColLetter');
  
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var sourceRange = sheet.getRange(sourceColumnLetter + ":" + sourceColumnLetter);
    var refColIndex = sheet.getRange(refColLetter + "1").getColumn();
    
    var newColIndex;
    if (insertBefore) {
      sheet.insertColumnBefore(refColIndex);
      newColIndex = refColIndex;
    } else {
      sheet.insertColumnAfter(refColIndex);
      newColIndex = refColIndex + 1;
    }
    
    var newColumnRange = sheet.getRange(1, newColIndex, sheet.getMaxRows(), 1);
    sourceRange.copyTo(newColumnRange);
    sheet.getRange(1, newColIndex).setValue(newHeaderText);
    
    var newColLetter = newColumnRange.getA1Notation().slice(0,-1);
    logAction("Structural Change", "Replicate Column", "Replicated data from column " + sourceColumnLetter + " to " + newColLetter + " with new header '" + newHeaderText + "'.", newColLetter);
    ui.alert('Successfully replicated column ' + sourceColumnLetter + '.');
  } catch (e) {
    ui.alert('An error occurred. Please ensure all column letters are valid.');
  }
}

function renameColumnHeader() {
  var ui = SpreadsheetApp.getUi();
  var colResult = ui.prompt('Enter the letter of the column to rename:', ui.ButtonSet.OK_CANCEL);
  var colLetter = colResult.getResponseText().toUpperCase().trim();
  if (colResult.getSelectedButton() !== ui.Button.OK || !colLetter) return;

  var headerResult = ui.prompt('Enter the new header text:', ui.ButtonSet.OK_CANCEL);
  var newHeaderText = headerResult.getResponseText();
  if (headerResult.getSelectedButton() !== ui.Button.OK) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  var headerCell = sheet.getRange(colLetter + '1');
  var oldHeaderText = headerCell.getValue();
  headerCell.setValue(newHeaderText);
  
  logAction("Structural Change", "Rename Header", "Renamed header from '" + oldHeaderText + "' to '" + newHeaderText + "'.", colLetter + '1');
  ui.alert('Column ' + colLetter + ' has been renamed to "' + newHeaderText + '".');
}

function deleteColumn() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter the letter of the column to permanently delete:', ui.ButtonSet.OK_CANCEL);
  var columnLetter = result.getResponseText().toUpperCase().trim();
  if (result.getSelectedButton() === ui.Button.OK && columnLetter) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(columnLetter + ":" + columnLetter);
    var header = sheet.getRange(columnLetter + "1").getValue();
    var confirm = ui.alert('Are you sure?', 'You are about to permanently delete column ' + columnLetter + ' ("' + header + '"). This cannot be undone.', ui.ButtonSet.YES_NO);
    if (confirm == ui.Button.YES) {
      sheet.deleteColumn(range.getColumn());
      logAction("Structural Change", "Delete Column", "Permanently deleted column (header was '" + header + "').", columnLetter);
      ui.alert('Column ' + columnLetter + ' has been deleted.');
    } else {
      ui.alert('Deletion cancelled.');
    }
  }
}

// --- DATA TRANSFORMATION FUNCTIONS ---

function conditionalPopulation() {
  var ui = SpreadsheetApp.getUi();
  var sourceResult = ui.prompt('Step 1: Source Column', 'Enter letter of column to check:', ui.ButtonSet.OK_CANCEL);
  var sourceCol = sourceResult.getResponseText().toUpperCase().trim();
  if (sourceResult.getSelectedButton() !== ui.Button.OK || !sourceCol) return;

  var conditionResult = ui.prompt('Step 2: Condition', 'Enter the value to look for in Column ' + sourceCol + ':', ui.ButtonSet.OK_CANCEL);
  var conditionVal = conditionResult.getResponseText();
  if (conditionResult.getSelectedButton() !== ui.Button.OK) return;

  var targetResult = ui.prompt('Step 3: Target Column', 'Enter letter of column to populate:', ui.ButtonSet.OK_CANCEL);
  var targetCol = targetResult.getResponseText().toUpperCase().trim();
  if (targetResult.getSelectedButton() !== ui.Button.OK || !targetCol) return;
  
  var valueResult = ui.prompt('Step 4: Value to Set', 'Enter the value to write in Column ' + targetCol + ':', ui.ButtonSet.OK_CANCEL);
  var valueToSet = valueResult.getResponseText();
  if (valueResult.getSelectedButton() !== ui.Button.OK) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var targetHeaderCell = sheet.getRange(targetCol + '1');
  if (targetHeaderCell.getValue() === '') {
    var headerPrompt = ui.prompt('New Column Detected', 'Enter a header for the new column ' + targetCol + ':', ui.ButtonSet.OK_CANCEL);
    if (headerPrompt.getSelectedButton() === ui.Button.OK) {
      targetHeaderCell.setValue(headerPrompt.getResponseText());
    }
  }
  
  var sourceValues = sheet.getRange(sourceCol + '1:' + sourceCol + lastRow).getValues();
  var targetRange = sheet.getRange(targetCol + '1:' + targetCol + lastRow);
  var targetValues = targetRange.getValues();
  var populatedCount = 0;
  
  for (var i = 1; i < sourceValues.length; i++) { // Start from row 2
    if (sourceValues[i][0] == conditionVal) {
      targetValues[i][0] = valueToSet;
      populatedCount++;
    }
  }
  
  if(populatedCount > 0){
    targetRange.setValues(targetValues);
    logAction("Data Transformation", "Conditional Population", "Populated " + populatedCount + " rows with value '" + valueToSet + "' where column " + sourceCol + " was '" + conditionVal + "'.", targetCol);
    ui.alert('Operation Complete', populatedCount + ' rows in Column ' + targetCol + ' were populated.', ui.ButtonSet.OK);
  } else {
    ui.alert('No rows met the specified condition.');
  }
}

// --- DATA VALIDATION FUNCTIONS (WITH LOGGING & REMOVAL) ---

function findAndHandleDuplicates() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Find & Handle Duplicates', 'Enter column letters to check (e.g., A, C):', ui.ButtonSet.OK_CANCEL);
  var inputText = result.getResponseText().trim();
  if (result.getSelectedButton() !== ui.Button.OK || !inputText) return;

  var html = HtmlService.createHtmlOutput(
      '<p>How should duplicates be handled?</p>' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processDuplicates(\'INTERACTIVE\');">Decide for Each Duplicate</button> ' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processDuplicates(\'REMOVE_ROW\');">Remove All Duplicate Rows</button> ' +
      '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processDuplicates(\'CLEAR_CELL\');">Clear All Duplicate Cells</button> ' +
      '<button onclick="google.script.host.close()">Cancel</button>')
      .setWidth(400).setHeight(150);
  
  PropertiesService.getScriptProperties().setProperty('duplicateFinderInput', inputText);
  ui.showModalDialog(html, 'Choose Action for Duplicates');
}

function processDuplicates(action) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var inputText = PropertiesService.getScriptProperties().getProperty('duplicateFinderInput');
  PropertiesService.getScriptProperties().deleteProperty('duplicateFinderInput');

  var cellsToClear = [];
  var rowsToDelete = [];
  var handledCount = 0;

  try {
    var columnLetters = inputText.toUpperCase().split(',').map(function(letter) { return letter.trim(); });
    for (var c = 0; c < columnLetters.length; c++) {
      var letter = columnLetters[c];
      var range = sheet.getRange(letter + ":" + letter);
      var values = range.getValues();
      var seen = {};

      for (var i = 1; i < values.length; i++) { // Start from row 2
        var cellValue = values[i][0];
        if (cellValue === "" || typeof cellValue === 'boolean') continue;
        
        var valueKey = typeof cellValue + '_' + cellValue;
        if (!seen[valueKey]) {
          seen[valueKey] = { firstOccurrence: letter + (i + 1) };
        } else {
          var rowIndex = i + 1;
          var cellAddress = letter + rowIndex;
          var userChoice = action;

          if (action === 'INTERACTIVE') {
            var promptMessage = 'Duplicate Found!\n\nValue: "' + cellValue + '" at ' + cellAddress +
                                '\nFirst seen at ' + seen[valueKey].firstOccurrence +
                                '\n\nYES = Remove Row\nNO = Clear Cell\nCANCEL = Skip';
            var interactiveResult = ui.alert(promptMessage, ui.ButtonSet.YES_NO_CANCEL);
            if (interactiveResult === ui.Button.YES) userChoice = 'REMOVE_ROW';
            else if (interactiveResult === ui.Button.NO) userChoice = 'CLEAR_CELL';
            else userChoice = 'SKIP';
          }

          if (userChoice === 'REMOVE_ROW') {
            if (rowsToDelete.indexOf(rowIndex) === -1) rowsToDelete.push(rowIndex);
            logAction("Validation", "Duplicate Removal", "Removed duplicate row. Value was '" + cellValue + "', first seen at " + seen[valueKey].firstOccurrence + ".", cellAddress);
            handledCount++;
          } else if (userChoice === 'CLEAR_CELL') {
            cellsToClear.push(cellAddress);
            logAction("Validation", "Duplicate Removal", "Cleared duplicate cell. Value was '" + cellValue + "', first seen at " + seen[valueKey].firstOccurrence + ".", cellAddress);
            handledCount++;
          }
        }
      }
    }

    if (cellsToClear.length > 0) sheet.getRangeList(cellsToClear).clearContent();
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort(function(a, b) { return b - a; }).forEach(function(rowIndex) {
        sheet.deleteRow(rowIndex);
      });
    }

    if (handledCount > 0) ui.alert('Operation Complete', 'Handled ' + handledCount + ' duplicate entries.', ui.ButtonSet.OK);
    else ui.alert('No duplicate values found.');
  } catch (e) {
    ui.alert('An error occurred: ' + e.message);
  }
}

function validateMobileNumbers() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter column letter with mobile numbers:', ui.ButtonSet.OK_CANCEL);
  var columnLetter = result.getResponseText().toUpperCase().trim();
  if (result.getSelectedButton() !== ui.Button.OK || !columnLetter) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(columnLetter + ":" + columnLetter);
  var values = range.getValues();
  var changesMade = 0;

  for (var i = 1; i < values.length; i++) { // Start from row 2
    var originalValue = values[i][0];
    var cellAddress = columnLetter + (i + 1);
    if (originalValue === null || originalValue === '') continue;

    var cleanedNumber = String(originalValue).replace(/\D/g, '');
    if (cleanedNumber.startsWith('91') && cleanedNumber.length === 12) cleanedNumber = cleanedNumber.substring(2);
    else if (cleanedNumber.startsWith('0') && cleanedNumber.length === 11) cleanedNumber = cleanedNumber.substring(1);
    
    if (cleanedNumber.length === 10 && /^[6789]/.test(cleanedNumber)) {
      if (values[i][0] != cleanedNumber) {
        logAction("Validation", "Mobile Validation", "Formatted '" + originalValue + "' to '" + cleanedNumber + "'.", cellAddress);
        values[i][0] = cleanedNumber;
        changesMade++;
      }
    } else {
      if (originalValue !== '') {
         logAction("Validation", "Mobile Validation", "Removed invalid number: '" + originalValue + "'.", cellAddress);
         values[i][0] = '';
         changesMade++;
      }
    }
  }

  if (changesMade > 0) {
    range.setValues(values);
    ui.alert('Validation Complete', 'Processed ' + changesMade + ' entries. See "Action Log" for details.', ui.ButtonSet.OK);
  } else {
    ui.alert('Validation Complete', 'No invalid or unformatted numbers found.');
  }
}

function validateEmails() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter the column letter with email addresses:', ui.ButtonSet.OK_CANCEL);
  var columnLetter = result.getResponseText().toUpperCase().trim();
  if (result.getSelectedButton() !== ui.Button.OK || !columnLetter) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(columnLetter + ":" + columnLetter);
  var values = range.getValues();
  var changesMade = 0;
  var domainCache = {};
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i;

  for (var i = 1; i < values.length; i++) { // Start from row 2
    var email = values[i][0];
    var cellAddress = columnLetter + (i + 1);
    if (typeof email !== 'string' || email.trim() === '') continue;
    
    email = email.trim();

    if (!emailRegex.test(email)) {
      logAction("Validation", "Email Validation", "Removed email with invalid syntax: '" + email + "'.", cellAddress);
      values[i][0] = '';
      changesMade++;
      continue;
    }

    var domain = email.substring(email.lastIndexOf("@") + 1).toLowerCase();
    if (domain === 'gmail.com') domainCache[domain] = true;

    if (domainCache[domain] === undefined) {
      try {
        var response = UrlFetchApp.fetch('https://dns.google/resolve?name=' + domain + '&type=MX', { 'muteHttpExceptions': true });
        var jsonResponse = JSON.parse(response.getContentText());
        domainCache[domain] = (jsonResponse.Status === 0 && jsonResponse.hasOwnProperty('Answer'));
      } catch (e) {
        console.error("DNS API call for " + domain + " failed. Assuming valid. Error: " + e.message);
        domainCache[domain] = true;
      }
    }
    
    if (domainCache[domain] === false) {
      logAction("Validation", "Email Validation", "Removed email with invalid domain (no MX record): '" + email + "'.", cellAddress);
      values[i][0] = '';
      changesMade++;
    }
  }

  if (changesMade > 0) {
    range.setValues(values);
    ui.alert('Validation Complete', 'Removed ' + changesMade + ' invalid emails. See "Action Log" for details.', ui.ButtonSet.OK);
  } else {
    ui.alert('Validation Complete', 'No invalid emails found.');
  }
}

/**
 * Validates Aadhaar numbers in a column using the Verhoeff algorithm.
 */
function validateAadhaarNumbers() {
  // Verhoeff algorithm tables
  const d = [[0, 1, 2, 3, 4, 5, 6, 7, 8, 9], [1, 2, 3, 4, 0, 6, 7, 8, 9, 5], [2, 3, 4, 0, 1, 7, 8, 9, 5, 6], [3, 4, 0, 1, 2, 8, 9, 5, 6, 7], [4, 0, 1, 2, 3, 9, 5, 6, 7, 8], [5, 9, 8, 7, 6, 0, 4, 3, 2, 1], [6, 5, 9, 8, 7, 1, 0, 4, 3, 2], [7, 6, 5, 9, 8, 2, 1, 0, 4, 3], [8, 7, 6, 5, 9, 3, 2, 1, 0, 4], [9, 8, 7, 6, 5, 4, 3, 2, 1, 0]];
  const p = [[0, 1, 2, 3, 4, 5, 6, 7, 8, 9], [1, 5, 7, 6, 2, 8, 3, 0, 9, 4], [5, 8, 0, 3, 7, 9, 6, 1, 4, 2], [8, 9, 1, 6, 0, 4, 3, 5, 2, 7], [9, 4, 5, 3, 1, 2, 6, 8, 7, 0], [4, 2, 8, 6, 5, 7, 3, 9, 0, 1], [2, 7, 9, 3, 8, 0, 6, 4, 1, 5], [7, 0, 4, 6, 9, 1, 3, 2, 5, 8]];
  
  function verhoeffValidate(num) {
    var c = 0;
    var num_array = String(num).split('').reverse();
    for (var i = 0; i < num_array.length; i++) {
      c = d[c][p[i % 8][parseInt(num_array[i], 10)]];
    }
    return (c === 0);
  }

  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter column letter with Aadhaar numbers:', ui.ButtonSet.OK_CANCEL);
  var columnLetter = result.getResponseText().toUpperCase().trim();
  if (result.getSelectedButton() !== ui.Button.OK || !columnLetter) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(columnLetter + ":" + columnLetter);
  var values = range.getValues();
  var changesMade = 0;

  for (var i = 1; i < values.length; i++) { // Start from row 2
    var originalValue = values[i][0];
    var cellAddress = columnLetter + (i + 1);
    if (originalValue === null || originalValue === '') continue;

    var cleanedNumber = String(originalValue).replace(/\D/g, '');
    
    if (cleanedNumber.length === 12) {
      if (verhoeffValidate(cleanedNumber)) {
        if (values[i][0] != cleanedNumber) {
          logAction("Validation", "Aadhaar Validation", "Formatted '" + originalValue + "' to '" + cleanedNumber + "'.", cellAddress);
          values[i][0] = cleanedNumber;
          changesMade++;
        }
      } else {
        logAction("Validation", "Aadhaar Validation", "Removed invalid Aadhaar (failed checksum): '" + originalValue + "'.", cellAddress);
        values[i][0] = '';
        changesMade++;
      }
    } else if (cleanedNumber.length > 0) {
      logAction("Validation", "Aadhaar Validation", "Removed invalid Aadhaar (not 12 digits): '" + originalValue + "'.", cellAddress);
      values[i][0] = '';
      changesMade++;
    }
  }

  if (changesMade > 0) {
    range.setValues(values);
    ui.alert('Validation Complete', 'Processed ' + changesMade + ' Aadhaar number entries. See "Action Log" for details.', ui.ButtonSet.OK);
  } else {
    ui.alert('Validation Complete', 'No invalid or unformatted Aadhaar numbers found.');
  }
}

/**
 * Validates and formats dates, using a DD/MM/YYYY assumption for ambiguous dates and prompting for ambiguous 2-digit years.
 */
function validateAndFormatDates() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter column letter with dates to validate (e.g., G):', ui.ButtonSet.OK_CANCEL);
  var columnLetter = result.getResponseText().toUpperCase().trim();
  if (result.getSelectedButton() !== ui.Button.OK || !columnLetter) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  try {
    var range = sheet.getRange(columnLetter + ":" + columnLetter);
  } catch (e) {
    ui.alert("Error: Invalid column letter provided.");
    return;
  }
  
  var values = range.getValues();
  var changesMade = 0;
  
  var twoDigitYearRegex = /^(\d{1,2}[-\/]\d{1,2}[-\/])(\d{2})$/;
  var dayMonthYearRegex = /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/;

  for (var i = 1; i < values.length; i++) { // Start from row 2
    var originalValue = values[i][0];
    var cellAddress = columnLetter + (i + 1);
    
    if (originalValue === null || originalValue === '' || originalValue instanceof Date) continue;

    var dateStr = String(originalValue).trim();
    var twoDigitMatch = dateStr.match(twoDigitYearRegex);
    var dayMonthMatch = dateStr.match(dayMonthYearRegex);
    var date;

    // --- STAGE 1: Handle Ambiguous 2-Digit Years (INTERACTIVE) ---
    if (twoDigitMatch) {
      var datePart = twoDigitMatch[1];
      var yearPart = twoDigitMatch[2];
      
      var year19xx = "19" + yearPart;
      var year20xx = "20" + yearPart;

      var html = HtmlService.createHtmlOutput(
        '<p>Ambiguous Year: <b>' + dateStr + '</b> at cell ' + cellAddress + '</p>' +
        '<p>Please choose the correct century:</p>' +
        '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processYearChoice({row: ' + (i+1) + ', value: \''+ datePart + year19xx +'\'});">Choose ' + year19xx + '</button> ' +
        '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processYearChoice({row: ' + (i+1) + ', value: \''+ datePart + year20xx +'\'});">Choose ' + year20xx + '</button> ' +
        '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processYearChoice({row: ' + (i+1) + ', value: \'INVALID\'});">Mark as Invalid</button>')
        .setWidth(400).setHeight(150);
      
      PropertiesService.getScriptProperties().setProperty('dateColumnLetter', columnLetter);
      ui.showModalDialog(html, 'Clarification Needed: Ambiguous Year');
      continue;
    }
    
    // --- STAGE 2: Handle DD/MM/YYYY vs MM/DD/YYYY (AUTOMATIC) ---
    else if (dayMonthMatch) {
      var day = parseInt(dayMonthMatch[1], 10);
      var month = parseInt(dayMonthMatch[2], 10);
      var year = parseInt(dayMonthMatch[3], 10);
      date = new Date(year, month - 1, day); 
    }
    
    // --- STAGE 3: Standard Parsing for All Other Formats ---
    else {
      date = new Date(dateStr);
    }
    
    // --- Final Validation and Formatting ---
    if (date && !isNaN(date.getTime())) {
      var year = date.getFullYear();
      var month = ('0' + (date.getMonth() + 1)).slice(-2);
      var day = ('0' + date.getDate()).slice(-2);
      var formattedDate = year + '-' + month + '-' + day;
      
      if (originalValue != formattedDate) {
        logAction("Validation", "Date Validation", "Formatted '" + originalValue + "' to '" + formattedDate + "'.", cellAddress);
        values[i][0] = formattedDate;
        changesMade++;
      }
    } else {
      logAction("Validation", "Date Validation", "Removed invalid date value: '" + originalValue + "'.", cellAddress);
      values[i][0] = '';
      changesMade++;
    }
  }

  if (changesMade > 0) {
    range.setValues(values);
    ui.alert('Validation Complete', 'Processed ' + changesMade + ' non-interactive entries. See "Action Log" for details.', ui.ButtonSet.OK);
  } else {
    ui.alert('Validation Complete', 'No unambiguous invalid or unformatted dates were found in the initial scan.', ui.ButtonSet.OK);
  }
}

/**
 * Handler function for the ambiguous year dialog.
 */
function processYearChoice(choice) {
  var columnLetter = PropertiesService.getScriptProperties().getProperty('dateColumnLetter');
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getRange(columnLetter + choice.row);
  var originalValue = cell.getValue();

  if (choice.value === 'INVALID') {
    cell.clearContent();
    logAction("Validation", "Date Validation (Manual)", "User marked ambiguous year date '" + originalValue + "' as invalid.", columnLetter + choice.row);
    return;
  }
  
  var dayMonthYearRegex = /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/;
  var dayMonthMatch = choice.value.match(dayMonthYearRegex);
  var date;
  
  if (dayMonthMatch) {
    var day = parseInt(dayMonthMatch[1], 10);
    var month = parseInt(dayMonthMatch[2], 10);
    var year = parseInt(dayMonthMatch[3], 10);
    date = new Date(year, month - 1, day);
  } else {
    date = new Date(choice.value);
  }

  if (date && !isNaN(date.getTime())) {
      var year = date.getFullYear();
      var month = ('0' + (date.getMonth() + 1)).slice(-2);
      var day = ('0' + date.getDate()).slice(-2);
      var formattedDate = year + '-' + month + '-' + day;
      cell.setValue(formattedDate);
      logAction("Validation", "Date Validation (Manual)", "User clarified ambiguous year '" + originalValue + "' as '" + formattedDate + "'.", columnLetter + choice.row);
  } else {
      cell.clearContent();
      logAction("Validation", "Date Validation (Manual)", "User-clarified date '" + choice.value + "' was still invalid and was removed.", columnLetter + choice.row);
  }
}


// --- DATA EXPORT FUNCTIONS ---

function exportFilteredDataAsCsv() {
  var ui = SpreadsheetApp.getUi();
  var colResult = ui.prompt('Step 1: Filter Column', 'Enter column letter to filter by:', ui.ButtonSet.OK_CANCEL);
  var filterColLetter = colResult.getResponseText().toUpperCase().trim();
  if (colResult.getSelectedButton() !== ui.Button.OK || !filterColLetter) return;

  var valueResult = ui.prompt('Step 2: Filter Value', 'Enter the value to export:', ui.ButtonSet.OK_CANCEL);
  var filterValue = valueResult.getResponseText();
  if (valueResult.getSelectedButton() !== ui.Button.OK) return;

  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var allData = sheet.getDataRange().getValues();
    var colIndex = sheet.getRange(filterColLetter + "1").getColumn() - 1;

    if (colIndex < 0 || colIndex >= allData[0].length) {
      throw new Error("Invalid column letter provided.");
    }
    
    var filteredData = allData.filter(function(row, index) {
      return index === 0 || (String(row[colIndex]).trim() === filterValue.trim());
    });

    if (filteredData.length <= 1) {
      ui.alert('No data found matching your criteria.');
      return;
    }

    var csvContent = filteredData.map(function(row) {
      return row.map(formatCsvField).join(',');
    }).join('\n');

    var fileName = 'Export_' + filterValue + '_from_' + sheet.getName() + '.csv';
    var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);

    logAction("Export", "CSV Export", "Exported " + (filteredData.length - 1) + " rows where column " + filterColLetter + " was '" + filterValue + "'.", "File: " + fileName);
    
    var html = 'Success! <a href="' + file.getUrl() + '" target="_blank">Click here to open the CSV file.</a>';
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(300).setHeight(80), 'Export Complete');
  } catch (e) {
    ui.alert('An error occurred: ' + e.message);
  }
}

/**
 * Helper function to format a single cell's value according to the specified CSV rules.
 * @param {any} value The cell value.
 * @return {string} The formatted CSV field.
 */
function formatCsvField(value) {
  if (value === null || value === undefined) {
    return '""';
  }
  var stringValue = String(value);
  var escapedValue = stringValue.replace(/"/g, '\\"');
  return '"' + escapedValue + '"';
}
