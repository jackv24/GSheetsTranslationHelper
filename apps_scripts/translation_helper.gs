/** @NotOnlyCurrentDoc */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Translation Helper')
      .addItem('Mark as Up-to-date', 'setActiveGreen')
      .addItem('Mark as Needs Updating', 'setActiveAmber')
      .addItem('Validate Range', 'validateRange')
      .addItem('Copy Protected Ranges from First to Current Sheet', 'copyProtectedRanges')
      .addToUi();
}

function setActiveGreen() {
  setActiveState('setTrue');
}

function setActiveAmber() {
  setActiveState('setFalse');
}

function validateRange() {
  setActiveState('setValid');
}

function setActiveState(checkType) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const values = range.getValues();
  const notes = range.getNotes();
  const bgColors = range.getBackgrounds();

  for (let j = 1; j <= numRows; j++) {
    for (let i = 1; i <= numCols; i++) {
      const cell = range.getCell(j, i);
      const column = cell.getColumn();
      const row = cell.getRow();

      // Ignore key column and title row
      if (column <= 1 || row <= 1) continue;

      const titleCell = sheet.getRange(1, column);
      const titleVal = titleCell.getValue();

      // Only operate on other language cells
      if (!titleVal || titleVal == "EN" || titleVal == "Notes" || titleVal == "Type") continue;

      if (checkType == 'setValid') {
        const cellVal = values[j-1][i-1];
        const cellEN = sheet.getRange(row, 2);
        const cellENVal = cellEN.getValue();

        if (!cellVal) {
          if (cellENVal) {
            // Needs updating if cell is empty but EN cell isn't
            bgColors[j-1][i-1] = "#ffd5b8";
          } else {
            // Up to date if both EN and self are empty
            bgColors[j-1][i-1] = "#d6ffb8";
          }

          // Note no longer relevant when cell is empty
          notes[j-1][i-1] = "";
        }
      } else if (checkType == 'setTrue') {
        const cellEN = sheet.getRange(row, 2);
        const cellENVal = cellEN.getValue();

        bgColors[j-1][i-1] = "#d6ffb8";
        notes[j-1][i-1] = cellENVal;
      } else if (checkType == 'setFalse') {
        bgColors[j-1][i-1] = "#ffd5b8";
        notes[j-1][i-1] = "";
      }
    }
  }

  range.setBackgrounds(bgColors);
  range.setNotes(notes);
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEditInstallable(e) {
  // Open sheet seperately so no undo operations are recorded for user
  var id = e.source.getId();
  var sheetName = e.source.getActiveSheet().getSheetName();
  var sheet = SpreadsheetApp.openById(id).getSheetByName(sheetName);

  // Get range in seperately opened sheet to match current sheet
  const range = sheet.getRange(e.range.getRow(), e.range.getColumn(), e.range.getNumRows(), e.range.getNumColumns());

  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const values = range.getValues();
  const notes = range.getNotes();
  const bgColors = range.getBackgrounds();
  const startColumn = range.getColumn();
  const startRow = range.getRow();

  for (let i = 0; i < numCols; i++) {
    const column = startColumn + i;

    // Ignore changes in key column
    if (column <= 1) continue;

    const titleCell = sheet.getRange(1, column);
    const titleVal = titleCell.getValue();

    // Only need translation tracking on language columns
    if (!titleVal || titleVal == "Notes" || titleVal == "Type") continue;

    for (let j = 0; j < numRows; j++) {
      const row = startRow + j;

      // Ignore changes in title row
      if (row <= 1) continue;

      if (titleVal == "EN") {
        // If English changed, update other languages states to say needs update

        // Get large enough row range to cover all languages
        const rowRange = sheet.getRange(row, 1, 1, 20);

        updateOtherLanguages(values[j][i], rowRange, sheet);

      } else {
        // If other language changed, update state to say translation up to date
        const cellEN = sheet.getRange(row, 2);
        const cellENVal = cellEN.getValue();

        const cellVal = values[j][i];

        if (!cellVal && cellENVal) {
          // Cell is empty but EN cell isn't, it must need updating
          notes[j][i] = "";
          bgColors[j][i] = "#ffd5b8";
        } else {
          // Set note to EN value so we know what value this was translated from
          notes[j][i] = cellENVal;
          
          // Mark up to date
          bgColors[j][i] = "#d6ffb8";
        }
      }
    }

    range.setBackgrounds(bgColors);
    range.setNotes(notes);
  }
}

function updateOtherLanguages(cellENVal, range, sheet) {
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const values = range.getValues();
  const notes = range.getNotes();
  const bgColors = range.getBackgrounds();
  const startColumn = range.getColumn();

  for (let i = 0; i < numCols; i++) {
    const column = startColumn + i;

    // Ignore key column
    if (column <= 1) continue;

    const titleCell = sheet.getRange(1, column);
    const titleVal = titleCell.getValue();

    // Only operate on other language cells
    if (!titleVal || titleVal == "EN" || titleVal == "Notes" || titleVal == "Type") continue;

    for (let j = 0; j < numRows; j++) {
      const cellVal = values[j][i];
      const note = notes[j][i];

      if (!cellVal && cellENVal) {
        // Cell is empty but EN cell isn't, it must need updating
        notes[j][i] = "";
        bgColors[j][i] = "#ffd5b8";
      } else if (note == cellENVal) {
        // EN value is the same as it was when this cell was updated, must be up to date
        bgColors[j][i] = "#d6ffb8";
      } else {
        // EN value has changed, this cell needs updating
        bgColors[j][i] = "#ffd5b8";
      }
    }
  }

  range.setBackgrounds(bgColors);
  range.setNotes(notes);
}

function copyProtectedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const firstSheet = sheets[0];

  const targetSheet = ss.getActiveSheet();

  if (targetSheet.getName() == firstSheet.getName()) {
    SpreadsheetApp.getUi().alert("Current sheet is source sheet!");
    return;
  }
  
  // Get protected ranges from the first sheet
  const protections = firstSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const columnHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].map(s => s.trim());
  
  // Clear existing protections in the target sheet (optional)
  const targetProtections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  targetProtections.forEach(protection => protection.remove());
  
  protections.forEach(protection => {
    const range = protection.getRange();
    const sourceColumnIndex = range.getColumn();
    const headerValue = firstSheet.getRange(1, sourceColumnIndex).getValue().trim();
    const targetColumnIndex = columnHeaders.indexOf(headerValue) + 1; // Match header and get column index

    if (targetColumnIndex > 0) {
      // Select entire column
      const targetRange = targetSheet.getRange(`${targetSheet.getRange(1, targetColumnIndex).getA1Notation()[0]}:${targetSheet.getRange(1, targetColumnIndex).getA1Notation()[0]}`); 
      const newProtection = targetRange.protect();
      
      // Copy permissions from source protection
      if (protection.canDomainEdit()) {
        newProtection.setDomainEdit(true); // Allow domain editing if enabled
      } else {
        const editors = protection.getEditors();
        newProtection.removeEditors(newProtection.getEditors()); // Clear default editors
        newProtection.addEditors(editors); // Add source editors
        
        if (protection.canEdit()) {
          newProtection.setDomainEdit(false); // Ensure only added editors can edit
        }
      }
      
      // Copy description (optional, useful for debugging or documentation)
      if (protection.getDescription()) {
        newProtection.setDescription(protection.getDescription());
      }
    }
  });
}

