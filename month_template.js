/**
 * Generates a monthly schedule template with large, roomy boxes.
 * Includes a sidebar with Name (Col A) and an Input column (Col B).
 */
function generateMonthTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  
  // 1. Get Month and Year
  const monthInput = sourceSheet.getRange("A1").getValue();
  const year = sourceSheet.getRange("B1").getValue();
  
  if (!monthInput || !year) {
    SpreadsheetApp.getUi().alert("Please enter a Month (1-12) in A1 and Year in B1");
    return;
  }

  // 2. Grab names AND their background colors from source Column A
  const lastRowNames = sourceSheet.getLastRow();
  let names = [];
  let nameBackgrounds = []; 
  
  if (lastRowNames >= 4) {
    const nameRange = sourceSheet.getRange(4, 1, lastRowNames - 3, 1);
    names = nameRange.getValues();
    nameBackgrounds = nameRange.getBackgrounds();
  }
  
  const monthIndex = monthInput - 1; 

  // 3. Create or get the target month sheet
  const monthName = new Intl.DateTimeFormat('en-US', { month: 'long' }).format(new Date(year, monthIndex));
  let targetSheet = ss.getSheetByName(monthName);
  
  if (!targetSheet) {
    targetSheet = ss.insertSheet(monthName);
  } else {
    targetSheet.clear();
  }

  // 4. Calculate days
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const dayNames = [];
  const dayNumbers = [];

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, monthIndex, d);
    const dayName = new Intl.DateTimeFormat('en-US', { weekday: 'short' }).format(date);
    dayNames.push(dayName);
    dayNumbers.push(d);
  }

  // 5. Set the headers (Schedule now starts at Column 3 / Column C)
  const scheduleStartCol = 3;
  targetSheet.getRange(3, scheduleStartCol, 1, daysInMonth).setValues([dayNames]);
  targetSheet.getRange(4, scheduleStartCol, 1, daysInMonth).setValues([dayNumbers]);

  // 6. Paste names and apply row background colors
  if (names.length > 0) {
    const dataRowStart = 5;
    // Paste names into Column A
    targetSheet.getRange(dataRowStart, 1, names.length, 1).setValues(names);
    
    // Loop through each name and apply its color to the entire row
    for (let i = 0; i < nameBackgrounds.length; i++) {
      const color = nameBackgrounds[i][0];
      if (color !== "#ffffff") {
        targetSheet.getRange(dataRowStart + i, 1, 1, daysInMonth + 2).setBackground(color);
      }
    }
  }

  // 7. BIG BOX STYLING
  const totalRows = names.length > 0 ? names.length + 4 : 15;
  const totalCols = daysInMonth + 2; // Col A (Name) + Col B (Input) + Days
  const fullRange = targetSheet.getRange(3, 1, totalRows - 2, totalCols);

  fullRange.setFontSize(14)
           .setVerticalAlignment("middle")
           .setHorizontalAlignment("center");

  targetSheet.setColumnWidth(1, 180); // Name Column
  targetSheet.setColumnWidth(2, 100); // New Input Column
  
  // Style the sidebar labels (Col A and B)
  targetSheet.getRange(5, 1, names.length, 2).setFontWeight("bold").setHorizontalAlignment("left");

  targetSheet.setColumnWidths(scheduleStartCol, daysInMonth, 70);
  targetSheet.setRowHeights(3, totalRows - 2, 45);

  // Header Colors
  const headerRange = targetSheet.getRange(3, 1, 2, totalCols);
  headerRange.setBackground("#e2efda").setFontWeight("bold");
  
  // Apply borders
  fullRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Freeze Name and Input columns
  targetSheet.setFrozenColumns(2);
}