function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Custom Tools')
    .addItem('Generate Month Template', 'generateMonthTemplate')
    .addItem('Generate Schedule', 'generateSchedule')
    .addSeparator()
    .addItem('Help', 'showHelpAlert')
    .addToUi();
}

/**
 * Optional: A simple helper function for the 'Help' menu item
 */
function showHelpAlert() {
  SpreadsheetApp.getUi().alert("Enter Month in A1, Year in B1, and Names starting at A4.");
}