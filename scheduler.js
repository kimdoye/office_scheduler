/**
 * Auto-assigns locations based on specific business rules and admin availability.
 * Adjusted for a horizontal layout (Dates in columns, Admins in rows).
 */
function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Configurations mapped directly to your screenshot
  const dayRow = 3;       // The row containing "Sun", "Mon", "Tue"...
  const adminStartRow = 5; // Jeff is on Row 5
  const startCol = 3;      // Column C is index 3 (Calendar starts here)
  const initialValueCol = 2; // Column B contains the initial "days worked" value
  
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const numAdmins = lastRow - adminStartRow + 1;

  if (lastCol < startCol || numAdmins <= 0) return; // Exit if sheet is empty

  // 1. Fetch data in bulk
  const daysData = sheet.getRange(dayRow, startCol, 1, lastCol - startCol + 1).getValues()[0];
  const initialValues = sheet.getRange(adminStartRow, initialValueCol, numAdmins, 1).getValues();
  const scheduleRange = sheet.getRange(adminStartRow, startCol, numAdmins, lastCol - startCol + 1);
  const scheduleData = scheduleRange.getValues();
  
  // Track weekly work counts and "off yesterday" status dynamically
  // Initialize from Column B values
  let weeklyWorkCount = initialValues.map(row => Number(row[0]) || 0);
  let wasOffYesterday = new Array(numAdmins).fill(false);

  // Map text days to numbers
  const dayMap = {
    "Sun": 0, "Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5, "Sat": 6
  };

  // Iterate through each column (day)
  for (let c = 0; c < daysData.length; c++) {
    let dayText = daysData[c];
    if (!dayText) continue; 
    
    let dayOfWeek = dayMap[dayText];

    // Reset weekly counter on Monday (1)
    if (dayOfWeek === 1) {
      weeklyWorkCount = new Array(numAdmins).fill(0);
    }

    let adminStatuses = []; 

    // 2. Analyze Admin Availability for the day
    for (let r = 0; r < numAdmins; r++) {
      let cellValue = scheduleData[r][c]; 
      let isRequestedOff = (cellValue === "NE");
      let hitMaxDays = (weeklyWorkCount[r] >= 5);
      
      adminStatuses.push({
        index: r,
        canWork: !isRequestedOff && !hitMaxDays,
        requestedOff: isRequestedOff,
        prefersOff: wasOffYesterday[r] 
      });
    }

    // 3. Define Needs based on Location Logic
    let needsPalmovka = 0;
    let needsStrizkov = 0;

    const isFirstDay = (c === 0);
    const isLastDay = (c === daysData.length - 1);

    if (isFirstDay || isLastDay) {
      needsPalmovka = 1;
      needsStrizkov = 2;
    } else if (dayOfWeek === 3 || dayOfWeek === 6) { // Wed or Sat
      needsStrizkov = 2; 
    } else if (dayOfWeek === 5) { // Fri
      needsPalmovka = 1; 
    } else {
      needsPalmovka = 1;
      needsStrizkov = 2;
    }

    // 4. Sorting logic for assignment
    let availableAdmins = adminStatuses
      .filter(a => a.canWork)
      .sort((a, b) => (a.prefersOff === b.prefersOff) ? 0 : a.prefersOff ? 1 : -1);

    let results = new Array(numAdmins).fill("");

    // 5. Assigning Roles (Priority: 1st Střížkov -> 1st Palmovka -> 2nd Střížkov)
    
    // Assign 1st person to Střížkov
    if (needsStrizkov > 0 && availableAdmins.length > 0) {
      let admin = availableAdmins.shift();
      results[admin.index] = "Střížkov";
      weeklyWorkCount[admin.index]++;
    }

    // Assign 1st person to Palmovka
    if (needsPalmovka > 0 && availableAdmins.length > 0) {
      let admin = availableAdmins.shift();
      results[admin.index] = "Palmovka";
      weeklyWorkCount[admin.index]++;
    }

    // Assign 2nd person to Střížkov (the "Preferred" staff)
    if (needsStrizkov > 1 && availableAdmins.length > 0) {
      let admin = availableAdmins.shift();
      results[admin.index] = "Střížkov";
      weeklyWorkCount[admin.index]++;
    }

    // 6. Update statuses & memory array for the next day
    for (let r = 0; r < numAdmins; r++) {
      wasOffYesterday[r] = (results[r] === "" && scheduleData[r][c] !== "NE");
      if (scheduleData[r][c] !== "NE") {
        scheduleData[r][c] = results[r];
      }
    }
  }

  // 7. Write everything back
  scheduleRange.setValues(scheduleData);
}