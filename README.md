# Seminar Form Scheduler (Google Apps Script)

A lightweight **mini ETL pipeline in JavaScript** for seminar attendance and feedback collection using **Google Sheets + Apps Script + Google Forms**.

The workflow is designed to support a **single QR code** (same Google Form URL all year) while enabling a team of organisers to collaboratively maintain the seminar calendar and add new seminars as they get confirmed.

---

## What it does

- Reads seminar rows from a worksheet (for example: `Seminar_Calendar`)
- Uses `Start_AEST` and `End_AEST` to determine when the form should be available
- Publishes + opens the form during an active seminar window
- Unpublishes + closes the form outside seminar windows
- Uses one Google Form and one QR code across the seminar year
- Streams Google Form responses back into same workbook for easy analysis and reporting

---

## Why this is useful (data/operations perspective)

This project is designed as a **low-maintenance operational data pipeline**:

- **Shared source of truth:** organisers collaborate in one Google Sheet
- **Controlled release:** the attendance/feedback form is only available during valid windows
- **Continuous data capture:** form responses land in the linked spreadsheet automatically
- **Actionable insights:** attendance and feedback can be analysed directly in Sheets or exported

This allows the seminar organising team to keep adding new seminars as they are confirmed throughout the year without changing the QR code or rebuilding forms while ensuring that all data is collected and stored in a centralised way.

---

## Core logic (hard requirement)

### `TITLE_HEADER` is mandatory

The script uses a **hard-coded row gate** based on `TITLE_HEADER` (for example, `Title`).

A row is **eligible to activate the form only if all of the following are true**:

1. `TITLE_HEADER` column value is **non-empty**
2. `Start_AEST` is present and valid
3. `End_AEST` is present and valid
4. Current time is between `Start_AEST` and `End_AEST`

### Why this matters

This rule prevents form publication for weeks when there is no seminars and allows the organising team to add more seminars as they get confirmed. Supports implement-and-forget.

In practice, it means the organising team can:
- add draft seminar rows early
- fill in dates/times progressively
- leave title blank until confirmed
- only trigger form publication once the row is complete enough to be considered real

**Important:** even if a row has values in `Start_AEST` and `End_AEST`, the script will **ignore it** if `TITLE_HEADER` is empty.

---

## Example workflow for a shared organising team

1. Organisers maintain a shared `Seminar_Organisation` sheet
2. New seminar rows are added as events are planned
3. Dates/times may be added early (e.g every Tuesday from 11:00-13:00)
4. The seminar title is filled in once confirmed
5. The script checks the sheet on a time-driven schedule
6. When the current time enters a valid row window **and** the title is non-empty, the form is published/opened
7. Attendees submit via the same QR code
8. Responses are written back to the linked response tab in the same workbook
9. Team analyses attendance + feedback in Google Sheets

---

## Requirements

- A Google Sheet workbook with a worksheet (e.g. `Seminar_Organisation`)
- Required columns:
  - `TITLE_HEADER` target column (e.g. `Title`)
  - `Start_AEST`
  - `End_AEST`
- A single Google Form (reused all year)
- Google Form responses linked to a sheet in the same workbook
- An Apps Script project bound to the spreadsheet
- A time-driven trigger

---

## Configuration constants (important)

In the script config:

- `TITLE_HEADER` = the exact header text used as the **confirmation gate**
  - Example: `'Title'`
  - If your sheet uses a different column name, update this constant
- `START_HEADER` = `'Start_AEST'`
- `END_HEADER` = `'End_AEST'`

### Header requirements

- Header names must match **exactly**
- The `TITLE_HEADER`, `Start_AEST`, and `End_AEST` headers should each be **unique**
- If `TITLE_HEADER` points to the wrong column (or a duplicate header exists), the script may fail or ignore intended rows

---

## Setup (high level)

1. Open the shared seminar workbook in Google Sheets
2. Open **Extensions → Apps Script**
3. Paste the script into `Code.gs`
4. Set the config constants:
   - Spreadsheet ID
   - Sheet name
   - Form ID
   - `TITLE_HEADER`
5. Run the setup function once
6. Confirm the time-driven trigger is installed
7. Link the Google Form responses to a tab in the same workbook
8. Keep using the same QR code for the Google Form URL

---

## Data model view (mini ETL framing)

### Source table (control table)
`Seminar_Organisation` worksheet

Key control fields:
- `TITLE_HEADER` (confirmation gate)
- `Start_AEST`
- `End_AEST`

### Operational state
- Form published/unpublished
- Form accepting/not accepting responses

### Output dataset
Google Form response tab in the same workbook (automatically appended rows), ready for:
- attendance summaries
- feedback review
- downstream analysis
- export to other systems if needed

---

## Notes

- This is intentionally a **minimal** script for reliability and maintainability
- It uses one form and one QR code instead of creating a new form per seminar
- The spreadsheet and Apps Script project should use the correct timezone (e.g. `Australia/Melbourne`)
- The form response destination can be the same workbook for simpler analysis and shared visibility

---

## License

MIT
