# Office Scheduler for Google Sheets

An automated scheduling solution built with Google Apps Script to manage staff assignments across multiple locations. This tool generates monthly calendar templates and auto-assigns staff based on availability, workload limits, consecutive work day constraints, and configurable office closure rules.

## 🚀 Features

- **Custom Spreadsheet Menu**: Adds a "Custom Tools" menu directly to the Google Sheets toolbar for easy access.
- **Dynamic Template Generation**:
  - Creates a full monthly calendar based on month and year inputs.
  - Automatically pulls staff names and inherits their background colors for better visual organization.
  - Applies professional formatting (borders, column widths, and font styles).
  - Automatically maps configurable weekly location closures and holidays into the generated month calendar.
- **Intelligent Scheduling Logic**:
  - **Availability Aware**: Respects "NE" (Requested Off) markers on the schedule.
  - **Workload Balancing**: Limits staff to a maximum of 5 days per week, tracking initial state for accurate mid-week transitions.
  - **Consecutive Days Logic**: Encourages contiguous work days while enforcing a maximum limit (penalizes working a 6th consecutive day).
  - **Dynamic Location Needs**: Assigns staff to "Střížkov" and "Palmovka", automatically prioritizing mandatory coverage before optional secondary shifts.
  - **Holiday & Closure Aware**: Adjusts capacity requirements intelligently when locations are marked as closed or on a global holiday.

## 🛠️ Setup & Installation

Since this is a Google Apps Script "Bound Script," follow these steps:

1. **Create a Google Sheet**: Open a new or existing Google Spreadsheet.
2. **Open Script Editor**: Go to `Extensions` > `Apps Script`.
3. **Add Files**: Create three files in the editor named `menu.gs`, `month_template.gs`, and `scheduler.gs`.
4. **Copy Code**: Paste the contents of the corresponding `.js` files from this repository into the `.gs` files in the editor.
5. **Save**: Click the disk icon or press `Cmd/Ctrl + S`.
6. **Refresh Sheet**: Refresh your Google Sheet. You should see a new **Custom Tools** menu appear.

## 📖 How to Use

### 1. Prepare the Input Data (Source Sheet)
On your active sheet (e.g., "Config"), define the parameters starting from row 4:
- **A1**: Month number (e.g., `4` for April).
- **B1**: Year (e.g., `2026`).
- **Column A (A4 downwards)**: Staff Names. You can set background colors for these names to color-code their rows in the generated template.
- **Column B**: Initial "Days Worked" count for the current week (useful when generating partial weeks or mid-month).
- **Column C**: Initial "Consecutive Days" worked prior to the 1st of the month.
- **Column E**: Holiday dates (day numbers, e.g., `15`).
- **Column F**: Location names (e.g., "Palmovka", "Střížkov").
- **Columns G-M**: Mark "X" to define weekly closures (Monday to Sunday) for the corresponding location in Column F.

### 2. Generate Template
Select `Custom Tools > Generate Month Template`. This creates a new tab named after the month (e.g., "April") with the calendar layout, dates, closed days labeled at the top, and all staff loaded.

### 3. Mark Availability
In the newly generated month sheet, mark specific days where staff are unavailable by entering **"NE"** in their respective cells.

### 4. Run Scheduler
Select `Custom Tools > Generate Schedule`. The script will calculate the optimal schedule based on availability, remaining weekly capacity, streak continuity, and dynamic staffing needs. It will fill the cells with the appropriate location names (e.g., "Střížkov", "Palmovka").

## 📂 Project Structure

- `menu.js`: UI configuration for the Google Sheets menu.
- `month_template.js`: Logic for calculating dates, layout generation, parsing input closures/holidays, and setting up the sheet formatting.
- `scheduler.js`: The "brain" of the project containing the assignment algorithm, capacity tracking, lookahead validation, and scoring constraints.

---

*Developed as part of the IFSA Internship project.*
