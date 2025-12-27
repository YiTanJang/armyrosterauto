# Army Roster Automation (부대 근무표 작성)

This project is a VBA-powered Excel tool designed to automate and manage military duty rosters. It aims to distribute duties fairly among soldiers while adhering to strict military regulations and constraints.

## Features

### 📅 Automated Scheduling
- **Smart Assignment:** Automatically assigns soldiers to shifts based on a fairness algorithm.
- **Fairness Score:** Calculates a priority score for each soldier (`Total Duty Score` / `Service Days`) to ensure equal distribution of workload.
- **Conflict Resolution:** Prevents consecutive duties (double shifts) and respects exclusion dates (leave/sick).

### 🛡️ Constraints & Rules
- **Rank Requirements:** Distinguishes between "Shift Leader" (requiring a minimum rank) and "Helper".
- **Holiday Handling:** Detects weekends and holidays, adjusting duty scores or rules accordingly.
- **Exemptions:** Skips soldiers who are on leave or have valid excuses.

### 🛠️ Management Tools
- **Dashboard UI:** A user-friendly "Main" sheet with buttons for all major actions.
- **Manual Adjustments:**
  - **Swap Duties:** Select two cells and press a button to swap personnel.
  - **Undo:** Revert the last roster generation or change.
- **Calendar View:** Generates a visual calendar sheet from the list-based roster for easy printing and viewing.
- **Archiving:** automated archiving of past duty records to keep the main dataset clean.

## Getting Started

1. **Download:** Clone this repository or download the `.xlsm` file.
2. **Open:** Open `부대 근무표 작성.xlsm` in Microsoft Excel.
3. **Enable Macros:** You must enable macros for the automation to work.
   - Click "Enable Content" on the yellow security bar if it appears.
4. **Setup:**
   - Go to the **Personnel (인원현황)** sheet to add soldiers and their details (Rank, Name, Enlistment Date).
   - Go to the **Settings (설정)** sheet to define shifts, difficulty scores, and holidays.
5. **Run:** Go to the **Main (메인)** sheet and click **"Generate Roster" (근무표 이어쓰기_생성)**.

## File Structure

- `부대 근무표 작성.xlsm`: The main Excel application file.
- `부대 근무표 작성.bas`: The source code module exported from the VBA editor.

## Requirements

- Microsoft Excel (Windows version recommended for full VBA compatibility).

## License

This project is open-source. Feel free to modify it to fit your unit's specific needs.
