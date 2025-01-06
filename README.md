## About

I wanted a simple, automated way to track how my time is spent, so I made this. This script automatically logs your Google Calendar events into a Google Sheet. It adds columns for each day, sums up event durations by user-defined categories, preserves any existing spreadsheet formulas, and color-codes events on the calendar. Essentially, it turns Google Calendar into a powerful time-tracking system with minimal effort.



## Setup
### Create a Google Sheet

1. Create or Open a Google Sheet
2. Create one sheet (click the + icon on the bottom left) per year: For example, name a sheet exactly "2025". (If the script references a year that doesn’t exist, it will create a new sheet automatically)
3. Format row 1 to use mm/dd/yyyy
4. Add category names in Column 1 (skip row 1)
5. Add category abbreviations and formulas in column 2

   * This script works by checking if any google calendar event names match abbreviations in column 2, then distributing the time in the correct date column. The script also creates new columns when needed, and will copy any formulas used in column 2. Example:
         ![TimeTrackerExample](https://github.com/user-attachments/assets/4310b892-7276-49da-9b68-f5880c894150)


### Copy the code from GitHub into AppsScript

1. From this GitHub page, Open the "SyncCalendarToSheet.gs" file in a new tab, and copy all the code
3. In your Google Sheet, go to Extensions → Apps Script
4. Paste into the AppsScript editor
5. Find and replace YOUR_GMAIL@gmail.com (line 25) with your Gmail. Make sure it is the same Gmail used with your google calendar
6. Find and replace Replace YOUR_SPREADSHEET_ID (line 26) with your Google Sheet ID

    * Find your Google Sheet ID in the sheet URL. Open your spreadsheet in a browser, and copy the string of letters and numbers that come after the "/d/" portion of the URL. Example:
         ![image](https://github.com/user-attachments/assets/def719d9-3dba-4cdb-bda4-f4de3a67af55)


### Add a Trigger

1. In the Apps Script Editor, click Triggers (alarm clock icon) or go to Project Settings → Triggers
2. Add Trigger for the function runScriptSafely(), choose Time-driven, and pick an interval (e.g., once per hour)
3. Authorize the script when prompted, so it can access your Calendar and Spreadsheet



## Usage

* To log an event, open the google calendar app, create an event with the desired time, and name it with one of your abbreviations
* To add/remove categories, add/romove rows to the google sheet with the same format shown above
* Event titles can have more than one category. The full event time will be distributed to each category in the event title.



## Customization

* Change event colors: In the Apps Script Editor, modify lines 269-273
* Change number of days that sync: In the Apps Script Editor, modify line 29
* Trigger when adding an event: In the Apps Script Editor, click Triggers, add a trigger for "runScriptSafely", select event source "from spreadsheet", select event type "on edit"



## How It Works

* The script uses a time-based or installable trigger to run at an interval you choose (e.g., hourly).
* For each day in the last DAYS_TO_SYNC days (plus today), it finds or creates a date column in your year-labeled sheet (e.g., “2025”, “2026”) and accumulates how many hours were spent on each category (parsed from event titles).
* Categories are recognized by short codes in the event title, e.g. "WP" for “Work Project,” "H" for “Health,” etc.
* If any row has a formula in Column B (like =SUM(...)), the script replicates it to the new date column, adjusting references from B to the new column’s letter.
* The script also checks whether the event is already colored correctly. If not, it sets the event color based on the first letter of the event title (W → red, L → blue, H → green, else → purple).



## Future Plans

* Add an example google sheet for copying format
* Add a dashboard for viewing data
* Add survey question support foe tracking goals (sleep scores, weekly goals, etc)
* Create a single app that handles everything