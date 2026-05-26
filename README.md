# CBS Food Signup Builder

This repository now includes a Google Apps Script that **creates the Google Sheet + Google Form for you**.

## File
- `google_apps_script_setup.gs`

## What it builds
- Spreadsheet tabs:
  - `Events`
  - `Slots`
  - `Signups`
  - `Helper_OpenSlots`
  - `Public_View`
- Form fields:
  - Name (required)
  - Email (required)
  - Event Date (dropdown)
  - Food Slot (dropdown)
  - Notes (optional)
- Form settings:
  - Collect email
  - Allow response edits after submit
  - Confirmation message

## How to run
1. Go to [script.google.com](https://script.google.com) and create a new project.
2. Replace the default code with `google_apps_script_setup.gs` from this repo.
3. Run `buildCbsFoodSignupSystem()`.
4. Authorize requested Google permissions.
5. Check execution logs for:
   - Spreadsheet URL
   - Live Form URL
   - Form Edit URL

## Ongoing maintenance
- To refresh Event Date and Food Slot choices based on latest sheet data:
  - Open the linked form script project and run `refreshFormChoicesFromSheets_()`.
- Recommended: add a time-driven trigger for periodic refresh.

## Notes
- `current_signup_count`, `open_spots`, and `status` are formula-driven in the `Slots` tab.
- Form choices include `slot_id` so records map cleanly to relational IDs for future migration.


## Better UX mode (recommended)
After running `buildCbsFoodSignupSystem()`, you can switch the form to **event-specific sections** (users choose event first, then only see slots for that event):

1. Open the form-linked Apps Script project.
2. Run `buildSectionBasedFormFromSheets_()`.
3. (Optional) Re-run this function weekly after changing events/slots.

This reduces mismatched Event/Slot selections and is usually easier for non-technical users.
