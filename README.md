# Seminar Form Scheduler (Google Apps Script)

A lightweight Google Apps Script that automatically publishes/unpublishes a single Google Form based on seminar time windows stored in a Google Sheet.

## What it does

- Reads `Start_AEST` and `End_AEST` from a worksheet (e.g. `Seminar_Organisation`)
- Publishes + opens the form during an active time window
- Unpublishes + closes the form outside the window
- Supports a single QR code workflow (same form URL all year)

## Use case

Designed for seminar attendance/feedback collection where:
- one Google Form is reused
- one QR code is displayed throughout the seminar year
- access is only available during seminar days/times

## Requirements

- Google Sheet with columns:
  - `Start_AEST`
  - `End_AEST`
- Google Form (single form reused)
- Apps Script project bound to the spreadsheet
- Time-driven trigger(s)

## Setup (high level)

1. Open the Google Sheet
2. Go to **Extensions → Apps Script**
3. Paste the script into `Code.gs`
4. Update:
   - Spreadsheet ID
   - Sheet name
   - Form ID
5. Run the setup function once (authorise)
6. Confirm the time-driven trigger is created
7. Keep using the same QR code pointing to the Google Form URL

## Notes

- The script relies on the spreadsheet/script timezone being set correctly (e.g. `Australia/Melbourne`)
- The form can be linked directly to a response destination spreadsheet
- No per-seminar form creation is required

## License

MIT
