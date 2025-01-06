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
5. Name your AppsScript project anything you like

### Update the IDs

1.  In the pasted code, find and replace YOUR_GMAIL@gmail.com (line 25) with your gmail. Make sure it is the same gmail used with your google calendar
2.  In the pasted code, find and replace Replace YOUR_SPREADSHEET_ID (line 26) with your Google Sheet ID
  * Get your Google Sheet ID from the URL. Open your spreadsheet in a browser, and copy the string of letters and numbers that come after the "/d/" portion of the URL. Example:


    ![image](https://github.com/user-attachments/assets/def719d9-3dba-4cdb-bda4-f4de3a67af55)


### Add a Trigger

1. In the Apps Script Editor, click Triggers (alarm clock icon) or go to Project Settings → Triggers
2. Add Trigger for the function runScriptSafely(), choose Time-driven, and pick an interval (e.g., once per hour)
3. Authorize the script when prompted, so it can access your Calendar and Spreadsheet
