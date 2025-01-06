# google-calendar-time-tracker
Sync google calendar events to a google sheet

This Google Apps Script automatically:

    Syncs your chosen Google Calendar events into a Google Sheet by day.
    Creates consecutive date columns for the past X days (no gaps).
    Logs event durations (excluding any events that haven't started yet).
    Skips clearing cells that contain formulas (e.g., summary formulas).
    Copies formulas from Column B into the newly updated date column.
    Colors Calendar events by the first letter of the event title (W, L, H, etc.).
    Prevents overlapping or rapid-fire runs via a locking system, so multiple triggers won't collide.
    Auto-expires the lock after 5 minutes if the script fails or is canceled.

How It Works

    The script uses a time-based or installable trigger to run at an interval you choose (e.g., hourly).
    For each day in the last DAYS_TO_SYNC days (plus today), it finds or creates a date column in your year-labeled sheet (e.g., “2025”, “2026”) and accumulates how many hours were spent on each category (parsed from event titles).
    Categories are recognized by short codes in the event title, e.g. "WP" for “Work Project,” "H" for “Health,” etc.
    If any row has a formula in Column B (like =SUM(...)), the script replicates it to the new date column, adjusting references from B to the new column’s letter.
    The script also checks whether the event is already colored correctly. If not, it sets the event color based on the first letter of the event title (W → red, L → blue, H → green, else → purple).

Setup Instructions

    Create or Open a Google Sheet for your time tracking.
    Create one sheet per year: For example, name a sheet exactly "2025". (If the script references a year that doesn’t exist, it will create a new sheet automatically.)
    Copy the Script:
        Open the Sheet → go to Extensions → Apps Script, or visit script.google.com directly.
        Create a new project or paste the script code into the editor.
    Update the IDs:
        In the code, find and replace YOUR_CALENDAR_ID@group.calendar.google.com with your actual Google Calendar ID.
        Replace YOUR_SPREADSHEET_ID with your Google Sheet ID (from the sheet’s URL).
    Set a Trigger:
        In the Apps Script Editor, click Triggers (alarm clock icon) or go to Project Settings → Triggers.
        Add Trigger for the function runScriptSafely() (not syncCalendarEvents()), choose Time-driven, and pick an interval (e.g., once per hour).
    Authorize the script when prompted, so it can access your Calendar and Spreadsheet.
