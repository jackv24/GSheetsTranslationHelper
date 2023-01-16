/** @NotOnlyCurrentDoc */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Translation Helper')
      .addItem('Mark as Up-to-date', 'setActiveGreen')
      .addItem('Mark as Needs Updating', 'setActiveAmber')
      //.addSeparator()
      .addToUi();
}

function setActiveGreen() {
  setActiveState(true);
}

function setActiveAmber() {
  setActiveState(false);
}

function setActiveState(isUpToDate) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

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

      // Only operate on cells with values
      const cellVal = cell.getValue();
      if (!cellVal) continue;

      if (isUpToDate) {
        const cellEN = sheet.getRange(row, 2);
        const cellENVal = cellEN.getValue();

        // Don't do expensive write if not needed
        if (cell.getNote() == cellENVal) continue;

        cell.setBackground("#d6ffb8");
        cell.setNote(cellENVal);
      } else {
        // Don't do expensive write if not needed
        if (!cell.getNote()) continue;

        cell.setBackground("#ffd5b8");
        cell.setNote("");
      }
    }
  }
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */
function onEditInstallable(e) {
  //return;

  // Open sheet seperately so no undo operations are recorded for user
  var id = e.source.getId();
  var sheetName = e.source.getActiveSheet().getSheetName();
  var sheet = SpreadsheetApp.openById(id).getSheetByName(sheetName);

  // Get range in seperately opened sheet to match current sheet
  const range = sheet.getRange(e.range.getRow(), e.range.getColumn(), e.range.getNumRows(), e.range.getNumColumns());

  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  for (let i = 1; i <= numCols; i++) {
    for (let j = 1; j <= numRows; j++) {
      const cell = range.getCell(j, i);
      const column = cell.getColumn();
      const row = cell.getRow();

      // Ignore changes in key column or title row
      if (column <= 1 || row <= 1) continue;

      const titleCell = sheet.getRange(1, column);
      const titleVal = titleCell.getValue();

      // Only need translation tracking on language columns
      if (!titleVal || titleVal == "Notes" || titleVal == "Type") continue;

      if (titleVal == "EN") {
        // If English changed, update other languages states to say needs update

        // Get large enough row range to cover all languages
        const rowRange = sheet.getRange(row, 1, 1, 20);

        updateOtherLanguages(cell, rowRange, sheet);

      } else {
        // If other language changed, update state to say translation up to date
        const cellEN = sheet.getRange(row, 2);
        const cellENVal = cellEN.getValue();

        // Set note to EN value so we know what value this was translated from
        cell.setNote(cellENVal);
        
        // Mark up to date
        cell.setBackground("#d6ffb8");
      }
    }
  }
}

function updateOtherLanguages(cellEN, rowRange, sheet) {
  const cellENVal = cellEN.getValue();

  const numRows = rowRange.getNumRows();
  const numCols = rowRange.getNumColumns();

  for (let i = 1; i <= numCols; i++) {
    for (let j = 1; j <= numRows; j++) {
      const cell = rowRange.getCell(j, i);
      const column = cell.getColumn();

      // Ignore key column
      if (column <= 1) continue;

      const titleCell = sheet.getRange(1, column);
      const titleVal = titleCell.getValue();

      // Only operate on other language cells
      if (!titleVal || titleVal == "EN" || titleVal == "Notes" || titleVal == "Type") continue;

      // Only operate on cells with values
      const cellVal = cell.getValue();
      if (!cellVal) continue;

      const note = cell.getNote();
      if (note == cellENVal) {
        // EN value is the same as it was when this cell was updated, must be up to date
        cell.setBackground("#d6ffb8");
      } else {
        // EN value has changed, this cell needs updating
        cell.setBackground("#ffd5b8");
      }
    }
  }
}
