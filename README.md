## About The Project

I wanted a simple, automated way to track how my time is spent, so I made this. This script automatically logs your Google Calendar events into a Google Sheet. It adds columns for each day, sums up event durations by user-defined categories, preserves any existing spreadsheet formulas, and color-codes events on the calendar. Essentially, it turns Google Calendar into a powerful time-tracking system with minimal effort.



## Setup Instructions

### Create a Google Sheet

1. Create or Open a Google Sheet
2. Create one sheet (click the + icon on the bottom left) per year: For example, name a sheet exactly "2025". (If the script references a year that doesn’t exist, it will create a new sheet automatically)
3. Format your sheets:
    * Row 1 contains dates using format mm/dd/yyy
    * Column 1 contains category names
    * Column 2 contains category abbreviations (your google calendar event titles)
    * Column 2 can also contain formulas, which will be automatically copied to new columns created by the script. See example:


    ![TimeTrackerExample](https://github.com/user-attachments/assets/4310b892-7276-49da-9b68-f5880c894150)


### Copy the code from GitHub into AppsScript

1. From this GitHub page, Open the "SyncCalendarToSheet.gs" file in a new tab, and copy all the code
3. In your Google Sheet, go to Extensions → Apps Script
4. Paste into the AppsScript editor
5. Find and replace YOUR_GMAIL@gmail.com (line 25) with your Gmail. Make sure it is the same Gmail used with your google calendar
6. Find and replace Replace YOUR_SPREADSHEET_ID (line 26) with your Google Sheet ID
    * Get your Google Sheet ID from the sheet URL. Open your spreadsheet in a browser, and copy the string of letters and numbers that come after the "/d/" portion of the URL. Example:


       ![image](https://github.com/user-attachments/assets/def719d9-3dba-4cdb-bda4-f4de3a67af55)


### Add a Trigger

1. In the Apps Script Editor, click Triggers (alarm clock icon) or go to Project Settings → Triggers
2. Add Trigger for the function runScriptSafely(), choose Time-driven, and pick an interval (e.g., once per hour)
3. Authorize the script when prompted, so it can access your Calendar and Spreadsheet



## Usage
*To log an event, open the google calendar app, create an event with the desired time, and name it with one of your abbreviations. For example, after a 30 minute breakfast, I would open the app and create a 30 minute event with the title HD (HD for Health: Diet). 
*To add/remove categories, open the google sheet and create a new row with the abbreviation in column 2 (same format as shown in the example image above). The script will automatically sort calendar events and their times according to the matching abbreviations in column 2. Add as many categories as you want, at any time.
*Adding multiple abbreviations to the same event will distribute the time to all of the events. For example, if you have a category Vacation (v), and you eat a 30 minute breakfast on vacation, you could create a 30 minute event titled "HD v" and both rows will receive 30 minutes. This allows you to track total vacation time, travel time, or other states, while still logging everything else. Make sure to separate abbreviations with a space when adding them to the calendar (HDv will not work, HD v will work).



## Customization
*Change event colors: In the Apps Script Editor, modify lines 269-273
*Change number of days that sync: In the Apps Script Editor, modify line 29
*Trigger when adding an event: In the Apps Script Editor, click Triggers, add a trigger for "runScriptSafely", select event source "from spreadsheet", select event type "on edit"



## How It Works

*The script uses a time-based or installable trigger to run at an interval you choose (e.g., hourly).
*For each day in the last DAYS_TO_SYNC days (plus today), it finds or creates a date column in your year-labeled sheet (e.g., “2025”, “2026”) and accumulates how many hours were spent on each category (parsed from event titles).
*Categories are recognized by short codes in the event title, e.g. "WP" for “Work Project,” "H" for “Health,” etc.
*If any row has a formula in Column B (like =SUM(...)), the script replicates it to the new date column, adjusting references from B to the new column’s letter.
*The script also checks whether the event is already colored correctly. If not, it sets the event color based on the first letter of the event title (W → red, L → blue, H → green, else → purple).
