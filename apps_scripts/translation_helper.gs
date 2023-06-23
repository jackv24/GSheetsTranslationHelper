/** @NotOnlyCurrentDoc */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Translation Helper')
      .addItem('Mark as Up-to-date', 'setActiveGreen')
      .addItem('Mark as Needs Updating', 'setActiveAmber')
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

      if (isUpToDate) {
        const cellEN = sheet.getRange(row, 2);
        const cellENVal = cellEN.getValue();

        bgColors[j-1][i-1] = "#d6ffb8";
        notes[j-1][i-1] = cellENVal;
      } else {
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

        // Set note to EN value so we know what value this was translated from
        notes[j][i] = cellENVal;
        
        // Mark up to date
        bgColors[j][i] = "#d6ffb8";
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
      const note = notes[j][i];
      if (note == cellENVal) {
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
