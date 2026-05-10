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

  // 2. Grab names, values, AND their background colors from source Columns A, B & C
  const lastRow = sourceSheet.getLastRow();
  let sidebarData = [];
  let nameBackgrounds = []; 
  let holidayData = [];
  let weeklyOffData = [];
  let locationData = [];
  
  if (lastRow >= 4) {
    // Fetch names and values (A, B and C)
    const sidebarRange = sourceSheet.getRange(4, 1, lastRow - 3, 3);
    sidebarData = sidebarRange.getValues();
    nameBackgrounds = sidebarRange.getBackgrounds();

    // Fetch holidays (E) independently (Shifted from D)
    holidayData = sourceSheet.getRange(4, 5, lastRow - 3, 1).getValues()
      .flat()
      .filter(h => h !== "" && !isNaN(h));

    // Fetch locations (F) independently (Shifted from E)
    locationData = sourceSheet.getRange(4, 6, lastRow - 3, 1).getValues().flat();

    // Fetch weekly days off (G-M) independently (Shifted from F-L)
    weeklyOffData = sourceSheet.getRange(4, 7, lastRow - 3, 7).getValues();
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

  // 5. Set the headers (Schedule now starts at Column 4 / Column D)
  const scheduleStartCol = 4;
  targetSheet.getRange(3, scheduleStartCol, 1, daysInMonth).setValues([dayNames]);
  targetSheet.getRange(4, scheduleStartCol, 1, daysInMonth).setValues([dayNumbers]);

  // Determine which days of the week have an "X" in columns F-L
  // Columns F-L map to: 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat, 6=Sun
  const closedDaysOfWeek = new Map();
  if (weeklyOffData.length > 0) {
    weeklyOffData.forEach((row, rowIndex) => {
      const locationName = locationData[rowIndex];
      row.forEach((cell, index) => {
        if (cell && cell.toString().toUpperCase() === "X") {
          // JS Date getDay(): 0=Sun, 1=Mon, 2=Tue, 3=Wed, 4=Thu, 5=Fri, 6=Sat
          const jsDay = index === 6 ? 0 : index + 1;
          if (!closedDaysOfWeek.has(jsDay)) {
            closedDaysOfWeek.set(jsDay, []);
          }
          if (locationName && !closedDaysOfWeek.get(jsDay).includes(locationName)) {
            closedDaysOfWeek.get(jsDay).push(locationName);
          }
        }
      });
    });
  }

  // Mark Weekly Off Days on Row 2
  for (let d = 1; d <= daysInMonth; d++) {
    // Skip labeling for the 1st and last day of the month
    if (d === 1 || d === daysInMonth) continue;

    const date = new Date(year, monthIndex, d);
    const jsDay = date.getDay();
    if (closedDaysOfWeek.has(jsDay)) {
      const label = closedDaysOfWeek.get(jsDay).join(" & ");
      targetSheet.getRange(2, scheduleStartCol + d - 1)
        .setValue(label)
        .setBackground("yellow")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    }
  }

  // Mark Holidays on Row 2
  if (holidayData.length > 0) {
    holidayData.forEach(h => {
      const dayNum = parseInt(h);
      if (dayNum >= 1 && dayNum <= daysInMonth) {
        targetSheet.getRange(2, scheduleStartCol + dayNum - 1)
          .setValue("HOLIDAY")
          .setBackground("yellow")
          .setFontWeight("bold")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      }
    });
  }
  
  // Optional: Set row height for the holiday row
  targetSheet.setRowHeight(2, 30);

  // 6. Paste names/values and apply row background colors
  if (sidebarData.length > 0) {
    const dataRowStart = 5;
    targetSheet.getRange(dataRowStart, 1, sidebarData.length, 3).setValues(sidebarData);
    
    for (let i = 0; i < sidebarData.length; i++) {
      const color = nameBackgrounds[i][0];
      if (color !== "#ffffff") {
        targetSheet.getRange(dataRowStart + i, 1, 1, daysInMonth + 3).setBackground(color);
      }
    }
  }

  // 7. BIG BOX STYLING
  const totalRows = sidebarData.length > 0 ? sidebarData.length + 4 : 15;
  const totalCols = daysInMonth + 3; 
  const fullRange = targetSheet.getRange(3, 1, totalRows - 2, totalCols);

  fullRange.setFontSize(14)
           .setVerticalAlignment("middle")
           .setHorizontalAlignment("center");

  targetSheet.setColumnWidth(1, 180); // Name Column
  targetSheet.setColumnWidth(2, 5);   // Value Column (Almost collapsed)
  targetSheet.setColumnWidth(3, 5);   // Consecutive Column (Almost collapsed)
  
  // Style the sidebar labels (Col A, B and C)
  targetSheet.getRange(5, 1, sidebarData.length, 3).setFontWeight("bold").setHorizontalAlignment("left");

  targetSheet.setColumnWidths(scheduleStartCol, daysInMonth, 70);
  targetSheet.setRowHeights(3, totalRows - 2, 45);

  // Header Colors
  const headerRange = targetSheet.getRange(3, 1, 2, totalCols);
  headerRange.setBackground("#e2efda").setFontWeight("bold");
  
  // Apply borders
  fullRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  // Freeze Name and Input columns
  targetSheet.setFrozenColumns(3);
}