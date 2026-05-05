/**
 * Auto-assigns locations based on specific business rules and admin availability.
 * Adjusted for a horizontal layout (Dates in columns, Admins in rows).
 */
function generateSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // Configurations mapped directly to your screenshot
  const closureRow = 2;   // The row containing closures/holidays
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
  const closureData = sheet.getRange(closureRow, startCol, 1, lastCol - startCol + 1).getValues()[0];
  const initialValues = sheet.getRange(adminStartRow, initialValueCol, numAdmins, 1).getValues();
  const scheduleRange = sheet.getRange(adminStartRow, startCol, numAdmins, lastCol - startCol + 1);
  const scheduleData = scheduleRange.getValues();
  
  // Track remaining work days and "off yesterday" status dynamically
  // Initialize remaining days from Column B's "days worked" value.
  let weeklyWorkLeft = initialValues.map(row => Math.max(5 - Number(row[0]), 0) || 0);
  let wasOffYesterday = new Array(numAdmins).fill(false);

  // Map text days to numbers
  const dayMap = {
    "Sun": 0, "Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5, "Sat": 6
  };

  // Iterate through each column (day)
  for (let c = 0; c < daysData.length; c++) {
    let dayText = daysData[c] ? daysData[c].toString().trim() : "";
    if (!dayText) continue; 
    
    let dayOfWeek = dayMap[dayText];

    // Reset weekly counter on Monday (1)
    if (dayOfWeek === 1) {
      weeklyWorkLeft = new Array(numAdmins).fill(5);
    }

    let adminStatuses = []; 

    // 2. Analyze Admin Availability & Calculate Priority Score for the day
    for (let r = 0; r < numAdmins; r++) {
      let cellValue = scheduleData[r][c]; 
      let isRequestedOff = (cellValue && cellValue.toString().toUpperCase().trim() === "NE");
      let hitMaxDays = (weeklyWorkLeft[r] == 0);
      
      let score = 0;
      
      // Points for scheduling scarcity (Max 50)
      // The fewer available days they have left compared to the shifts they need, the more points.
      let endOfWeek = c;
      while (endOfWeek < daysData.length - 1) {
        if (dayMap[daysData[endOfWeek + 1]] === 1) break; // Next day is Monday
        endOfWeek++;
      }
      
      let availableDaysLeft = 0;
      for (let checkCol = c; checkCol <= endOfWeek; checkCol++) {
        let isNEFuture = scheduleData[r][checkCol] && scheduleData[r][checkCol].toString().toUpperCase().trim() === "NE";
        if (!isNEFuture) {
          availableDaysLeft++;
        }
      }

      let shiftsNeeded = weeklyWorkLeft[r];
      if (shiftsNeeded > 0) {
        let buffer = Math.max(0, availableDaysLeft - shiftsNeeded);
        // buffer 0 = 50 pts, buffer 1 = 40 pts, buffer 2 = 30 pts, etc.
        score += Math.max(0, 50 - (buffer * 10));
      }

      // Points for working yesterday (Max 20)
      if (!wasOffYesterday[r]) {
        score += 20;
      }

      // Points for more weekly capacity (Max 30)
      let capacity = weeklyWorkLeft[r];
      score += (capacity / 5) * 30;

      adminStatuses.push({
        index: r,
        canWork: !isRequestedOff && !hitMaxDays,
        requestedOff: isRequestedOff,
        prefersOff: wasOffYesterday[r],
        score: score
      });
    }

    // 3. Define Needs based on Location Logic
    const { needsPalmovka, needsStrizkov } = getNeedsForDay(dayOfWeek, closureData[c]);

    // 3b. Special Holiday Logic: Everyone uses one weekly day, nobody is scheduled
    const isHoliday = closureData[c] && closureData[c].toString().toUpperCase().trim() === "HOLIDAY";
    if (isHoliday) {
      for (let r = 0; r < numAdmins; r++) {
        weeklyWorkLeft[r] = Math.max(0, weeklyWorkLeft[r] - 1);
      }
    }

    // 4. Sorting logic for assignment
    let availableAdmins = (isHoliday) ? [] : adminStatuses
      .filter(a => a.canWork)
      .sort((a, b) => b.score - a.score);

    let results = new Array(numAdmins).fill("");

    // 5. Assigning Roles
    // Priority: 1st Střížkov -> 1st Palmovka -> 2nd Střížkov -> 2nd Palmovka
    
    // Assign 1st person to Střížkov
    if (needsStrizkov > 0 && availableAdmins.length > 0) {
      let admin = availableAdmins.shift();
      results[admin.index] = "Střížkov";
      weeklyWorkLeft[admin.index] = Math.max(0, weeklyWorkLeft[admin.index] - 1);
    }

    // Assign 1st person to Palmovka
    if (needsPalmovka > 0 && availableAdmins.length > 0) {
      let admin = availableAdmins.shift();
      results[admin.index] = "Palmovka";
      weeklyWorkLeft[admin.index] = Math.max(0, weeklyWorkLeft[admin.index] - 1);
    }

    // Assign 2nd person to Střížkov (the "Preferred" staff)
    // Only assign if we have enough capacity for mandatory shifts for the rest of the week
    if (needsStrizkov > 1 && availableAdmins.length > 0) {
      if (hasEnoughCapacityForRestOfWeek(c, weeklyWorkLeft, scheduleData, daysData, dayMap, numAdmins, closureData)) {
        let admin = availableAdmins.shift();
        results[admin.index] = "Střížkov";
        weeklyWorkLeft[admin.index] = Math.max(0, weeklyWorkLeft[admin.index] - 1);
      }
    }

    // Assign 2nd person to Palmovka as the last-priority optional shift
    // Only assign if we still have enough capacity for mandatory shifts for the rest of the week
    if (needsPalmovka > 1 && availableAdmins.length > 0) {
      if (hasEnoughCapacityForRestOfWeek(c, weeklyWorkLeft, scheduleData, daysData, dayMap, numAdmins, closureData)) {
        let admin = availableAdmins.shift();
        results[admin.index] = "Palmovka";
        weeklyWorkLeft[admin.index] = Math.max(0, weeklyWorkLeft[admin.index] - 1);
      }
    }

    // 6. Update statuses & memory array for the next day
    for (let r = 0; r < numAdmins; r++) {
      const isNE = (scheduleData[r][c] && scheduleData[r][c].toString().toUpperCase().trim() === "NE");
      // "Off yesterday" is true if they didn't work. This resets their "consecutive days" bonus.
      wasOffYesterday[r] = (results[r] === "");
      
      if (isNE) {
        scheduleData[r][c] = "NE";
      } else {
        scheduleData[r][c] = results[r];
      }
    }
  }

  // 7. Write everything back
  const baseBackgrounds = sheet.getRange(adminStartRow, 1, numAdmins, 1).getBackgrounds();
  const newBackgrounds = scheduleData.map((row, r) => 
    row.map(cell => (cell === "NE" ? "yellow" : baseBackgrounds[r][0]))
  );

  scheduleRange.setValues(scheduleData);
  scheduleRange.setBackgrounds(newBackgrounds);
}

/**
 * Helper to determine staffing needs for a given day.
 */
function getNeedsForDay(dayOfWeek, closureLabel = "") {
  let needsPalmovka = 2;
  let needsStrizkov = 2;


  // Handle closures based on row 2 labels
  if (closureLabel) {
    const labelUpper = closureLabel.toString().toUpperCase().trim();
    if (labelUpper === "HOLIDAY") {
      needsPalmovka = 0;
      needsStrizkov = 0;
    } else {
      if (labelUpper.includes("PALMOVKA")) {
        needsPalmovka = 0;
      }
      if (labelUpper.includes("STRIZKOV") || labelUpper.includes("STŘÍŽKOV")) {
        needsStrizkov = 0;
      }
    }
  }

  return { needsPalmovka, needsStrizkov };
}

/**
 * Checks if assigning an extra optional shift today would leave enough capacity
 * for mandatory shifts (1st Strizkov and 1st Palmovka) for the rest of the week.
 */
function hasEnoughCapacityForRestOfWeek(currentCol, weeklyWorkLeft, scheduleData, daysData, dayMap, numAdmins, closureData) {
  // Identify the end of the current scheduling week
  let endOfWeek = currentCol + 1;
  while (endOfWeek < daysData.length) {
    if (dayMap[daysData[endOfWeek]] === 1) break; // Monday starts a new week
    endOfWeek++;
  }

  // Copy remaining work days to simulate future mandatory shifts
  let capacities = [...weeklyWorkLeft];

  // Simulate assigning future mandatory shifts (1st Strizkov and 1st Palmovka)
  for (let c = currentCol + 1; c < endOfWeek; c++) {
    const isHolidayFuture = closureData[c] && closureData[c].toString().toUpperCase().trim() === "HOLIDAY";
    if (isHolidayFuture) {
      for (let r = 0; r < numAdmins; r++) {
        capacities[r] = Math.max(0, capacities[r] - 1);
      }
      continue; // No mandatory shifts on holidays
    }

    const { needsPalmovka, needsStrizkov } = getNeedsForDay(dayMap[daysData[c]], closureData ? closureData[c] : "");
    let mandatoryToday = (needsPalmovka > 0 ? 1 : 0) + (needsStrizkov > 0 ? 1 : 0);

    // Get admins available today (not NE), sorted by remaining work days descending
    let availableToday = [];
    for (let r = 0; r < numAdmins; r++) {
      const isNE = (scheduleData[r][c] && scheduleData[r][c].toString().toUpperCase() === "NE");
      if (!isNE && capacities[r] > 0) {
        availableToday.push(r);
      }
    }
    availableToday.sort((a, b) => capacities[b] - capacities[a]);

    // If we can't find enough distinct people for mandatory shifts today, return false
    if (availableToday.length < mandatoryToday) {
      return false;
    }

    // "Consume" capacity for the top candidates
    for (let i = 0; i < mandatoryToday; i++) {
      capacities[availableToday[i]]--;
    }
  }

  // After fulfilling all future mandatory shifts, we must have at least 1 spare capacity 
  // left across all admins to afford assigning today's optional shift.
  let remainingCapacity = capacities.reduce((sum, val) => sum + val, 0);
  return remainingCapacity > 0;
}
