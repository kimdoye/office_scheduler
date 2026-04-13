# Office Scheduler for Google Sheets

An automated scheduling solution built with Google Apps Script to manage staff assignments across multiple locations for bro-coli. This tool generates monthly calendar templates and auto-assigns staff based on availability, workload limits, and specific office closure rules.

## 🚀 Features

- **Custom Spreadsheet Menu**: Adds a "Custom Tools" menu directly to the Google Sheets toolbar.
- **Dynamic Template Generation**:
  - Creates a full monthly calendar based on month/year inputs.
  - Automatically pulls staff names and inherits their background colors for better visual organization.
  - Applies professional formatting (borders, column widths, and font styles).
- **Intelligent Scheduling Logic**:
  - **Availability Aware**: Respects "NE" (Requested Off) markers.
  - **Workload Balancing**: Limits staff to a maximum of 5 days per week.
  - **Location-Specific Rules**:
    - **Wed/Sat**: Palmovka closed, assigns 2 people to Střížkov.
    - **Fri**: Střížkov closed, assigns 1 person to Palmovka.
    - **Other days**: Assigns 1 to Palmovka and 2 to Střížkov.
  - **Rest Logic**: Prioritizes staff who had the previous day off.

## 🛠️ Setup & Installation

Since this is a Google Apps Script "Bound Script," follow these steps:

1. **Create a Google Sheet**: Open a new or existing Google Spreadsheet.
2. **Open Script Editor**: Go to `Extensions` > `Apps Script`.
3. **Add Files**: Create three files in the editor named `menu.gs`, `month_template.gs`, and `scheduler.gs`.
4. **Copy Code**: Paste the contents of the corresponding `.js` files from this repository into the `.gs` files in the editor.
5. **Save**: Click the disk icon or press `Cmd/Ctrl + S`.
6. **Refresh Sheet**: Refresh your Google Sheet. You should see a new **Custom Tools** menu appear.

## 📖 How to Use

### 1. Prepare the Input
On your active sheet:
- Set **A1** to the Month number (e.g., `4` for April).
- Set **B1** to the Year (e.g., `2026`).
- List staff names in column **A** starting from **A4**. You can set background colors for these names to color-code their rows in the generated template.

### 2. Generate Template
Select `Custom Tools > Generate Month Template`. This creates a new tab named after the month (e.g., "April").

### 3. Mark Availability
In the new month sheet, mark days where staff are unavailable by entering **"NE"** in their respective cells.

### 4. Run Scheduler
Select `Custom Tools > Generate Schedule`. The script will fill in the remaining cells with "Střížkov" or "Palmovka" based on the business rules.

## 📂 Project Structure

- `menu.js`: UI configuration for the Google Sheets menu.
- `month_template.js`: Logic for calculating dates, formatting cells, and setting up the sheet layout.
- `scheduler.js`: The "brain" of the project containing the assignment algorithm and business constraints.

---

*Developed as part of the IFSA Internship project.*
