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
  const adminStartRow = 5; // First admin is on Row 5
  const startCol = 4;      // Column D is index 4 (Calendar starts here)
  const initialValueCol = 2; // Column B contains the initial "days worked" value
  const consecutiveCol = 3;  // Column C contains the initial "consecutive days" value
  
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const numAdmins = lastRow - adminStartRow + 1;

  if (lastCol < startCol || numAdmins <= 0) return; // Exit if sheet is empty

  // 1. Fetch data in bulk
  const daysData = sheet.getRange(dayRow, startCol, 1, lastCol - startCol + 1).getValues()[0];
  const closureData = sheet.getRange(closureRow, startCol, 1, lastCol - startCol + 1).getValues()[0];
  const initialValues = sheet.getRange(adminStartRow, initialValueCol, numAdmins, 2).getValues(); // Fetch B and C
  const scheduleRange = sheet.getRange(adminStartRow, startCol, numAdmins, lastCol - startCol + 1);
  const scheduleData = scheduleRange.getValues();
  
  // Track remaining work days dynamically
  // weeklyWorkLeft: Math.max(5 - Column B, 0)
  // current_streak: Column C
  let weeklyWorkLeft = initialValues.map(row => Math.max(5 - Number(row[0]), 0) || 0);
  let current_streak = initialValues.map(row => Number(row[1]) || 0);

  // Iterate through each column (day)
  for (let c = 0; c < daysData.length; c++) {
    let dayText = normalizeCell(daysData[c]);
    if (!dayText) continue; 
    
    let dayOfWeek = DAY_MAP[dayText];

    // Reset weekly counter on Monday (1)
    if (dayOfWeek === 1) {
      weeklyWorkLeft = new Array(numAdmins).fill(5);
    }

    const adminStatuses = getAdminStatusesForDay(c, numAdmins, scheduleData, daysData, weeklyWorkLeft, current_streak, closureData);

    // 3. Define Needs based on Location Logic
    const { needsPalmovka, needsStrizkov } = getNeedsForDay(closureData[c]);

    // 3b. Special Holiday Logic: Everyone uses one weekly day, nobody is scheduled
    const holiday = isHoliday(closureData[c]);
    if (holiday) {
      reduceWeeklyWorkForAll(weeklyWorkLeft);
    }

    // 4. Sorting logic for assignment
    let availableAdmins = (holiday) ? [] : adminStatuses
      .filter(a => a.canWork)
      .sort((a, b) => b.score - a.score);

    let results = new Array(numAdmins).fill("");
    const isLastDay = (c === daysData.length - 1);
    let mandatoryAssigned = 0;

    // 5. Assigning Roles
    // Priority: 1st Střížkov -> 1st Palmovka -> 2nd Střížkov -> 2nd Palmovka
    
    // Assign 1st person to Střížkov
    if (needsStrizkov > 0 && availableAdmins.length > 0) {
      assignShift(results, availableAdmins, weeklyWorkLeft, "Střížkov");
      mandatoryAssigned++;
    }

    // Assign 1st person to Palmovka
    if (needsPalmovka > 0 && availableAdmins.length > 0) {
      assignShift(results, availableAdmins, weeklyWorkLeft, "Palmovka");
      mandatoryAssigned++;
    }

    // Assign 2nd person to Střížkov (the "Preferred" staff)
    // Only assign if this specific admin can take the optional shift safely.
    if (needsStrizkov > 1 && availableAdmins.length > 0) {
      if (isLastDay && mandatoryAssigned < 3) {
        assignShift(results, availableAdmins, weeklyWorkLeft, "Střížkov");
        mandatoryAssigned++;
      } else {
        assignOptionalShiftIfSafe(results, availableAdmins, weeklyWorkLeft, "Střížkov", c, scheduleData, daysData, numAdmins, closureData);
      }
    }

    // Assign 2nd person to Palmovka as the last-priority optional shift
    // Only assign if this specific admin can take the optional shift safely.
    if (needsPalmovka > 1 && availableAdmins.length > 0) {
      if (isLastDay && mandatoryAssigned < 3) {
        assignShift(results, availableAdmins, weeklyWorkLeft, "Palmovka");
        mandatoryAssigned++;
      } else {
        assignOptionalShiftIfSafe(results, availableAdmins, weeklyWorkLeft, "Palmovka", c, scheduleData, daysData, numAdmins, closureData);
      }
    }

    applyDayResults(c, numAdmins, scheduleData, results, current_streak, closureData);
  }

  // 7. Write everything back
  const baseBackgrounds = sheet.getRange(adminStartRow, 1, numAdmins, 1).getBackgrounds();
  const newBackgrounds = buildScheduleBackgrounds(scheduleData, baseBackgrounds);

  scheduleRange.setValues(scheduleData);
  scheduleRange.setBackgrounds(newBackgrounds);
}

const DAY_MAP = {
  "Sun": 0,
  "Mon": 1,
  "Tue": 2,
  "Wed": 3,
  "Thu": 4,
  "Fri": 5,
  "Sat": 6
};

function normalizeCell(value) {
  return value ? value.toString().trim() : "";
}

function normalizeCellUpper(value) {
  return normalizeCell(value).toUpperCase();
}

function isRequestedOff(value) {
  return normalizeCellUpper(value) === "NE";
}

function isHoliday(value) {
  return normalizeCellUpper(value) === "HOLIDAY";
}

function findEndOfWeek(currentCol, daysData) {
  let endOfWeek = currentCol;
  while (endOfWeek < daysData.length - 1) {
    if (DAY_MAP[normalizeCell(daysData[endOfWeek + 1])] === 1) break; // Next day is Monday
    endOfWeek++;
  }

  return endOfWeek;
}

function getAdminStatusesForDay(currentCol, numAdmins, scheduleData, daysData, weeklyWorkLeft, currentStreak, closureData) {
  const statuses = [];

  for (let r = 0; r < numAdmins; r++) {
    const streak = currentStreak[r];
    const isHolidayToday = isHoliday(closureData[currentCol]);
    
    // Check if taking today off would make it impossible to reach 5 shifts in some future week
    const forcedToWork = !isHolidayToday && !isRequestedOff(scheduleData[r][currentCol]) && 
                         isForcedToWork(r, currentCol, scheduleData, daysData, weeklyWorkLeft, currentStreak, closureData);

    statuses.push({
      index: r,
      canWork: !isRequestedOff(scheduleData[r][currentCol]) && weeklyWorkLeft[r] !== 0 && streak < 6,
      score: calculateAdminScore(r, currentCol, scheduleData, daysData, weeklyWorkLeft, streak, forcedToWork)
    });
  }

  return statuses;
}

function calculateAdminScore(adminIndex, currentCol, scheduleData, daysData, weeklyWorkLeft, streak, forcedToWork) {
  let score = 0;

  // PRIORITY 1: Forced to work (Lookahead failure)
  if (forcedToWork) {
    score += 1000;
  }

  // Points for scheduling scarcity (Max 50)
  // The fewer available days they have left compared to the shifts they need, the more points.
  const availableDaysLeft = countAvailableDaysLeft(adminIndex, currentCol, scheduleData, daysData);
  const shiftsNeeded = weeklyWorkLeft[adminIndex];

  if (shiftsNeeded > 0) {
    const buffer = Math.max(0, availableDaysLeft - shiftsNeeded);
    // buffer 0 = 50 pts, buffer 1 = 40 pts, buffer 2 = 30 pts, etc.
    score += Math.max(0, 50 - (buffer * 10));
  }

  // Points for consecutive work days
  if (streak === 5) {
    score -= 100; // Sharp dropoff for the 6th day
  } else if (streak < 5) {
    score += streak * 10; // Points go up to 5 days
  }

  // Points for more weekly capacity (Max 30)
  score += (shiftsNeeded / 5) * 30;

  return score;
}

function countAvailableDaysLeft(adminIndex, currentCol, scheduleData, daysData) {
  const endOfWeek = findEndOfWeek(currentCol, daysData);
  let availableDaysLeft = 0;

  for (let checkCol = currentCol; checkCol <= endOfWeek; checkCol++) {
    if (!isRequestedOff(scheduleData[adminIndex][checkCol])) {
      availableDaysLeft++;
    }
  }

  return availableDaysLeft;
}

function reduceWeeklyWorkForAll(weeklyWorkLeft) {
  for (let r = 0; r < weeklyWorkLeft.length; r++) {
    weeklyWorkLeft[r] = Math.max(0, weeklyWorkLeft[r] - 1);
  }
}

function assignShift(results, availableAdmins, weeklyWorkLeft, location) {
  const admin = availableAdmins.shift();
  results[admin.index] = location;
  weeklyWorkLeft[admin.index] = Math.max(0, weeklyWorkLeft[admin.index] - 1);
}

function assignOptionalShiftIfSafe(results, availableAdmins, weeklyWorkLeft, location, currentCol, scheduleData, daysData, numAdmins, closureData) {
  for (let i = 0; i < availableAdmins.length; i++) {
    const admin = availableAdmins[i];
    const capacities = [...weeklyWorkLeft];
    const capacityAfterOptionalShift = Math.max(0, capacities[admin.index] - 1);
    capacities[admin.index] = capacityAfterOptionalShift;

    if (canCoverFutureMandatoryShifts(currentCol, capacities, scheduleData, daysData, numAdmins, closureData)) {
      availableAdmins.splice(i, 1);
      results[admin.index] = location;
      weeklyWorkLeft[admin.index] = capacityAfterOptionalShift;
      return true;
    }
  }

  return false;
}

function applyDayResults(currentCol, numAdmins, scheduleData, results, currentStreak, closureData) {
  for (let r = 0; r < numAdmins; r++) {
    const holiday = isHoliday(closureData[currentCol]);

    if (holiday) {
      currentStreak[r] = 0;
    } else if (results[r] !== "") {
      currentStreak[r]++;
    } else {
      currentStreak[r] = 0;
    }

    if (holiday) {
      scheduleData[r][currentCol] = "";
    } else if (isRequestedOff(scheduleData[r][currentCol])) {
      scheduleData[r][currentCol] = "NE";
    } else if (!holiday) {
      scheduleData[r][currentCol] = results[r];
    }
  }
}

function buildScheduleBackgrounds(scheduleData, baseBackgrounds) {
  return scheduleData.map((row, r) =>
    row.map(cell => (cell === "NE" ? "yellow" : baseBackgrounds[r][0]))
  );
}

/**
 * Helper to determine staffing needs for a given day.
 */
function getNeedsForDay(closureLabel = "") {
  let needsPalmovka = 2;
  let needsStrizkov = 2;


  // Handle closures based on row 2 labels
  if (closureLabel) {
    const labelUpper = normalizeCellUpper(closureLabel);
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
 * Checks whether the remaining capacities can cover mandatory future shifts
 * (1st Střížkov and 1st Palmovka) for the rest of the current week.
 */
function canCoverFutureMandatoryShifts(currentCol, capacities, scheduleData, daysData, numAdmins, closureData) {
  // Identify the end of the current scheduling week
  const endOfWeek = findEndOfWeek(currentCol, daysData);
  const lastDayIndex = daysData.length - 1;

  // Simulate assigning future mandatory shifts (1st Strizkov and 1st Palmovka)
  for (let c = currentCol + 1; c <= endOfWeek; c++) {
    if (isHoliday(closureData[c])) {
      reduceWeeklyWorkForAll(capacities);
      continue; // No mandatory shifts on holidays
    }

    const { needsPalmovka, needsStrizkov } = getNeedsForDay(closureData ? closureData[c] : "");
    let mandatoryToday = (needsPalmovka > 0 ? 1 : 0) + (needsStrizkov > 0 ? 1 : 0);

    // Requirement: On the last day of the month, we need 3 people available.
    if (c === lastDayIndex) {
      mandatoryToday = Math.min(3, needsPalmovka + needsStrizkov);
    }

    // Get admins available today (not NE), sorted by remaining work days descending
    let availableToday = [];
    for (let r = 0; r < numAdmins; r++) {
      if (!isRequestedOff(scheduleData[r][c]) && capacities[r] > 0) {
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

  return true;
}

/**
 * Determines if an admin MUST work today because taking it off would
 * lead to a failed schedule (under 5 shifts) in a future week due to streak limits.
 */
function isForcedToWork(adminIndex, currentCol, scheduleData, daysData, weeklyWorkLeft, currentStreak, closureData) {
  // Scenario A: Work Today
  const canIfWork = canSatisfyVisibleWeeklyTargets(adminIndex, currentCol, true, scheduleData, daysData, weeklyWorkLeft, currentStreak, closureData);
  
  // Scenario B: Take Today Off
  const canIfOff = canSatisfyVisibleWeeklyTargets(adminIndex, currentCol, false, scheduleData, daysData, weeklyWorkLeft, currentStreak, closureData);

  // Forced only when working today fixes a failure that taking today off would cause.
  return !canIfOff && canIfWork;
}

/**
 * Simulates whether an admin can still satisfy all full visible weekly targets.
 * A trailing partial week is ignored unless the visible data ends on Sunday.
 */
function canSatisfyVisibleWeeklyTargets(adminIndex, currentCol, workToday, scheduleData, daysData, weeklyWorkLeft, currentStreak, closureData) {
  let streak = currentStreak[adminIndex];
  let remainingNeed = weeklyWorkLeft[adminIndex];

  // Apply today's choice
  if (isHoliday(closureData[currentCol])) {
    remainingNeed = Math.max(0, remainingNeed - 1);
    streak = 0;
  } else if (workToday) {
    streak++;
    remainingNeed = Math.max(0, remainingNeed - 1);
  } else {
    streak = 0;
  }

  for (let c = currentCol + 1; c < daysData.length; c++) {
    const dayOfWeek = DAY_MAP[normalizeCell(daysData[c])];
    
    // A Monday starts a new week, so the previous visible week must be complete.
    if (dayOfWeek === 1) {
      if (remainingNeed > 0) return false;
      remainingNeed = 5;
    }

    const holiday = isHoliday(closureData[c]);
    const requestedOff = isRequestedOff(scheduleData[adminIndex][c]);

    if (holiday) {
      remainingNeed = Math.max(0, remainingNeed - 1);
      streak = 0;
    } else if (!requestedOff && streak < 6 && remainingNeed > 0) {
      remainingNeed--;
      streak++;
    } else {
      streak = 0;
    }
  }

  const lastDay = DAY_MAP[normalizeCell(daysData[daysData.length - 1])];
  return lastDay !== 0 || remainingNeed === 0;
}
