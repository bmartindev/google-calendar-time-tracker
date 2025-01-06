# google-calendar-time-tracker
Sync google calendar events to a google sheet



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

![TimeTrackerExample](https://github.com/user-attachments/assets/a09af55b-fe44-4f56-bdef-8370f8b7d6ef)

### Copy the Script

1. In your Google Sheet, go to Extensions → Apps Script
2. Name your project
3. Create a new project or paste the script code into the editor. AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA COPY PASTE

### Update the IDs

1.  In the code, find and replace YOUR_CALENDAR_ID@group.calendar.google.com with your actual Google Calendar ID
2.  Replace YOUR_SPREADSHEET_ID with your Google Sheet ID (from the sheet’s URL)

### Set a Trigger

1. In the Apps Script Editor, click Triggers (alarm clock icon) or go to Project Settings → Triggers
2. Add Trigger for the function runScriptSafely(), choose Time-driven, and pick an interval (e.g., once per hour)
3. Authorize the script when prompted, so it can access your Calendar and Spreadsheet
