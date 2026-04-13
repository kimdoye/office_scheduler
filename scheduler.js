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
  const startCol = 2;      // Column B is index 2
  const numAdmins = 3;     // Jeff, Tim, Ariana
  
  const lastCol = sheet.getLastColumn();
  if (lastCol < startCol) return; // Exit if sheet is empty

  // 1. Fetch data in bulk (much faster than cell-by-cell)
  const daysData = sheet.getRange(dayRow, startCol, 1, lastCol - startCol + 1).getValues()[0];
  const scheduleRange = sheet.getRange(adminStartRow, startCol, numAdmins, lastCol - startCol + 1);
  const scheduleData = scheduleRange.getValues();
  
  // Track weekly work counts (Index 0: Jeff, 1: Tim, 2: Ariana)
  let weeklyWorkCount = [0, 0, 0];
  let wasOffYesterday = [false, false, false];

  // Map text days to numbers to match original logic
  const dayMap = {
    "Sun": 0, "Mon": 1, "Tue": 2, "Wed": 3, "Thu": 4, "Fri": 5, "Sat": 6
  };

  // Iterate through each column (day)
  for (let c = 0; c < daysData.length; c++) {
    let dayText = daysData[c];
    if (!dayText) continue; // Skip if header is blank
    
    let dayOfWeek = dayMap[dayText];

    // Reset weekly counter on Sunday
    if (dayOfWeek === 0) {
      weeklyWorkCount = [0, 0, 0];
    }

    let adminStatuses = []; 

    // 2. Analyze Admin Availability for the day
    for (let r = 0; r < numAdmins; r++) {
      let cellValue = scheduleData[r][c]; // Get cell for specific admin on this day
      let isRequestedOff = (cellValue === "NE");
      let hitMaxDays = (weeklyWorkCount[r] >= 5);
      
      adminStatuses.push({
        index: r,
        canWork: !isRequestedOff && !hitMaxDays,
        requestedOff: isRequestedOff,
        prefersOff: wasOffYesterday[r] // Two-in-a-row logic
      });
    }

    // 3. Define Needs based on Location Logic
    let needsPalmovka = 0;
    let needsStrizkov = 0;

    if (dayOfWeek === 3 || dayOfWeek === 6) { // Wed or Sat
      needsStrizkov = 2; // Palmovka closed
    } else if (dayOfWeek === 5) { // Fri
      needsPalmovka = 1; // Střížkov closed
    } else {
      needsPalmovka = 1;
      needsStrizkov = 2;
    }

    // 4. Sorting logic for assignment
    // Prioritize admins who can work. Deprioritize those who "prefer off" (yesterday was off).
    let availableAdmins = adminStatuses
      .filter(a => a.canWork)
      .sort((a, b) => (a.prefersOff === b.prefersOff) ? 0 : a.prefersOff ? 1 : -1);

    let results = ["", "", ""];

    // 5. Assigning Roles
    // Assign Střížkov first (usually needs more people)
    for (let i = 0; i < needsStrizkov; i++) {
      if (availableAdmins.length > 0) {
        let admin = availableAdmins.shift();
        results[admin.index] = "Střížkov";
        weeklyWorkCount[admin.index]++;
      }
    }

    // Assign Palmovka
    for (let i = 0; i < needsPalmovka; i++) {
      if (availableAdmins.length > 0) {
        let admin = availableAdmins.shift();
        results[admin.index] = "Palmovka";
        weeklyWorkCount[admin.index]++;
      }
    }

    // 6. Update statuses & memory array for the next day
    for (let r = 0; r < numAdmins; r++) {
      wasOffYesterday[r] = (results[r] === "" && scheduleData[r][c] !== "NE");
      
      // If they didn't specifically request "NE", update their schedule block
      if (scheduleData[r][c] !== "NE") {
        scheduleData[r][c] = results[r];
      }
    }
  }

  // 7. Write everything back to the sheet in one swift motion
  scheduleRange.setValues(scheduleData);
}