/*******************************************************
 * 1) SAFE ENTRY POINT (Locking via Script Properties)
 *******************************************************/
function runScriptSafely() {
  if (isScriptLocked()) {
    Logger.log("Script locked (and not timed out), skipping.");
    return;
  }
  lockScript(); // Mark locked with timestamp

  try {
    syncCalendarEvents(); 
  } catch (err) {
    Logger.log("Error: " + err);
    throw err;
  } finally {
    unlockScript(); // Always unlock so we don't remain stuck
  }
}

/*******************************************************
 * 2) MAIN LOGIC: Sync Calendar Events
 *******************************************************/
function syncCalendarEvents() {
  const calendarId = "YOUR_GMAIL@gmail.com";  //REPLACE WITH YOUR GOOGLE CALENDAR GMAIL
  const spreadsheetId = "YOUR_SPREADSHEET_ID";  // REPLACE WITH YOUR SPREADSHEET ID

  // How many days back to include (plus today)
  const DAYS_TO_SYNC = 7;

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error("Calendar not found; check your calendarId.");
  }

  // Process from oldest => newest day
  for (let dayOffset = DAYS_TO_SYNC; dayOffset >= 0; dayOffset--) {
    const dayDate = getDateNDaysAgo(dayOffset);
    syncOneDay(spreadsheet, calendar, dayDate);
  }
}

/*******************************************************
 * SYNC ONE DAY
 *  1) Find/create date column (no gaps)
 *  2) Clear old data in that column (non-formula cells only) from row 2 downward
 *  3) Gather & write event durations (excluding future events)
 *  4) AFTER writing data, replicate any formula from Column B 
 *     to the same row in this date column
 *  5) Only color events if:
 *        a) the event has started, AND
 *        b) the event title is listed in Column B, AND
 *        c) the current color differs from desired color
 *******************************************************/
function syncOneDay(spreadsheet, calendar, dayDate) {
  const year = dayDate.getFullYear();
  const sheet = getOrCreateYearSheet(spreadsheet, year);

  // 1) Ensure there's a column for dayDate
  const { colIndex: dateCol } = findOrCreateDateColumn(sheet, dayDate);

  // 2) Clear old data (non-formula cells) in this column (row 2..lastRow)
  clearColumnValues(sheet, dateCol);

  // 3) Gather event data (skip future events) => write hours
  const dayStart = new Date(year, dayDate.getMonth(), dayDate.getDate(), 0, 0, 0);
  const dayEnd   = new Date(year, dayDate.getMonth(), dayDate.getDate() + 1, 0, 0, 0);
  const events   = calendar.getEvents(dayStart, dayEnd);
  const now      = new Date();

  const categoryHours = {};
  events.forEach(event => {
    if (event.getStartTime() > now) return; // skip future
    const duration = getDurationOfEventOnDate(event, dayDate);
    if (duration <= 0) return;

    // Parse categories from event title (if you like multiple categories)
    const categories = getCategoriesFromTitle(event.getTitle());
    categories.forEach(cat => {
      categoryHours[cat] = (categoryHours[cat] || 0) + duration;
    });
  });

  // Write the hours in one pass
  if (Object.keys(categoryHours).length > 0) {
    writeCategoryHours(sheet, dateCol, categoryHours);
  }

  // 4) Copy any formula from Column B to the same row in this column
  replicateFormulasFromColB(sheet, dateCol);

  // 5) Color events only if they're listed in Column B and have started
  const validTitles = getTitlesInColumnB(sheet); // read all titles in col B
  events.forEach(event => {
    if (event.getStartTime() > now) return; 
    // Only color if the event title is in col B
    const title = event.getTitle().trim();
    if (validTitles.includes(title)) {
      setEventColorIfNeeded(event);
    }
  });
}

/*******************************************************
 * GET or CREATE YEAR SHEET
 *******************************************************/
function getOrCreateYearSheet(spreadsheet, year) {
  const name = String(year);
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  return sheet;
}

/*******************************************************
 * FIND or CREATE DATE COLUMN
 * If missing, create a new rightmost column
 *******************************************************/
function findOrCreateDateColumn(sheet, dayDate) {
  let lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    // If sheet is fully blank
    sheet.getRange(1, 1).setValue("Header");
    lastCol = 1;
  }

  // Check row 1 for matching dayDate
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let col = 1; col <= lastCol; col++) {
    if (isSameDay(headers[col - 1], dayDate)) {
      return { colIndex: col };
    }
  }

  // If not found, create a new column
  const newCol = lastCol + 1;
  const headerCell = sheet.getRange(1, newCol);
  headerCell.setValue(dayDate);
  headerCell.setNumberFormat("M/d"); // display mm/dd
  return { colIndex: newCol };
}

/*******************************************************
 * CLEAR COLUMN VALUES (non-formula only, row 2..lastRow)
 *******************************************************/
function clearColumnValues(sheet, colIndex) {
  const lastRow = sheet.getMaxRows();
  if (lastRow < 2) return;

  // We'll clear from row 2..N
  const rangeToClear = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const formulas = rangeToClear.getFormulas();
  const values   = rangeToClear.getValues();

  for (let r = 0; r < values.length; r++) {
    // If no formula => clear it
    if (!formulas[r][0]) {
      values[r][0] = "";
    }
  }
  rangeToClear.setValues(values);
}

/*******************************************************
 * COPY FORMULAS from COLUMN B to the same row in dateCol
 * For any row that has a formula in col B, replicate it
 * adjusting references from B -> {dateColLetter}
 *******************************************************/
function replicateFormulasFromColB(sheet, dateCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const bFormulas  = sheet.getRange(1, 2, lastRow, 1).getFormulas(); 
  const newColLetter = columnIndexToLetter(dateCol);

  // For each row 1..lastRow, if there's a formula in col B, replicate it
  for (let i = 0; i < lastRow; i++) {
    const formulaB = bFormulas[i][0];
    if (!formulaB) continue; // no formula in col B => skip

    // e.g. "=SUM(B3:B7)" => "=SUM(K3:K7)"
    const adjusted = formulaB.replace(
      /(\$?)B(\$?\d+)/gi,
      `$1${newColLetter}$2`
    );
    // Write to same row, dateCol
    sheet.getRange(i + 1, dateCol).setFormula(adjusted);
  }
}

/*******************************************************
 * WRITE CATEGORY HOURS in one pass
 * We skip row 1 so we don't overwrite the date header
 *******************************************************/
function writeCategoryHours(sheet, dateCol, categoryHours) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Read col B for category names (row 1..N)
  const catVals = sheet.getRange(1, 2, lastRow, 1).getValues();

  // Also read existing formulas in dateCol
  const formVals = sheet.getRange(1, dateCol, lastRow, 1).getFormulas();

  // We'll build a 2D array for rows 1..lastRow
  const writeData = [];
  for (let i = 0; i < lastRow; i++) {
    writeData.push([""]); 
  }

  // For each row, if category matches => set hours (skipping formula cells)
  for (let i = 0; i < catVals.length; i++) {
    const category = catVals[i][0];
    if (!category) continue; 
    const hours = categoryHours[category];
    if (!hours) continue;

    // skip if there's a formula in that cell
    if (!formVals[i][0]) {
      writeData[i][0] = hours;
    }
  }

  // Write back but skipping row 1
  const numRowsToWrite = lastRow - 1; // excludes row 1
  const sliceOfData = writeData.slice(1); // skip index 0
  if (numRowsToWrite > 0 && sliceOfData.length >= numRowsToWrite) {
    sheet
      .getRange(2, dateCol, numRowsToWrite, 1)
      .setValues(sliceOfData.slice(0, numRowsToWrite));
  }
}

/*******************************************************
 * ONLY COLOR IF EVENT TITLE APPEARS IN COLUMN B
 *******************************************************/
function setEventColorIfNeeded(event) {
  const desiredColor = determineColorByTitle(event.getTitle());
  if (event.getColor() !== desiredColor) {
    event.setColor(desiredColor);
  }
}

/*******************************************************
 * GET ALL TITLES IN COLUMN B
 * (We'll call this in syncOneDay)
 *******************************************************/
function getTitlesInColumnB(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  const rawValues = sheet.getRange(1, 2, lastRow, 1).getValues();
  // Build a simple array of trimmed strings
  const titles = rawValues
    .map(row => row[0])
    .filter(val => val) // non-empty
    .map(val => String(val).trim());
  return titles;
}

/*******************************************************
 * DETERMINE COLOR BY TITLE
 *******************************************************/
function determineColorByTitle(title) {
  if (!title) return CalendarApp.EventColor.GRAY;
  const c = title.trim().charAt(0).toUpperCase();
  switch (c) {
    case "W": return CalendarApp.EventColor.RED; //Events beginning with W will be RED
    case "L": return CalendarApp.EventColor.CYAN; //Events beginning with L will be BLUE (CYAN)
    case "H": return CalendarApp.EventColor.PALE_GREEN; //Events beginning with H will be GREEN (PALE_GREEN)
    case "S": return CalendarApp.EventColor.PALE_BLUE //Events beginning with S will be PURPLE (PALE_BLUE)
    default:  return CalendarApp.EventColor.GRAY; //Events remaining will be GRAY
  }
}

/*******************************************************
 * GET CATEGORIES FROM TITLE
 *******************************************************/
function getCategoriesFromTitle(title) {
  // If you prefer one category per event, adjust accordingly
  // For multiple categories e.g. "WP HD", we split by spaces
  return title.trim().split(/\s+/).filter(Boolean);
}

/*******************************************************
 * GET DURATION OF EVENT ON A PARTICULAR DAY (hours)
 *******************************************************/
function getDurationOfEventOnDate(event, dayDate) {
  const dayStart = new Date(dayDate.getFullYear(), dayDate.getMonth(), dayDate.getDate(), 0, 0, 0);
  const dayEnd   = new Date(dayDate.getFullYear(), dayDate.getMonth(), dayDate.getDate() + 1, 0, 0, 0);

  const start = event.getStartTime();
  const end   = event.getEndTime();
  const effStart = new Date(Math.max(start.getTime(), dayStart.getTime()));
  const effEnd   = new Date(Math.min(end.getTime(), dayEnd.getTime()));
  const duration = effEnd - effStart;
  if (duration <= 0) return 0;
  return duration / (1000 * 60 * 60);
}

/*******************************************************
 * CHECK IF TWO DATES ARE THE SAME (Y/M/D)
 *******************************************************/
function isSameDay(cellValue, dayDate) {
  if (!(cellValue instanceof Date)) {
    const parsed = new Date(cellValue);
    if (isNaN(parsed.valueOf())) return false;
    cellValue = parsed;
  }
  return (
    cellValue.getFullYear() === dayDate.getFullYear() &&
    cellValue.getMonth() === dayDate.getMonth() &&
    cellValue.getDate() === dayDate.getDate()
  );
}

/*******************************************************
 * GET DATE N DAYS AGO (no time)
 *******************************************************/
function getDateNDaysAgo(n) {
  const now = new Date();
  const d   = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  d.setDate(d.getDate() - n);
  return d;
}

/*******************************************************
 * HELPER: COLUMN INDEX -> LETTER
 *******************************************************/
function columnIndexToLetter(index) {
  let temp = "";
  let letter = "";
  while (index > 0) {
    temp   = (index - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    index  = (index - temp - 1) / 26;
  }
  return letter;
}

/*******************************************************
 * LOCKING UTILITIES (Script Properties)
 *******************************************************/
function isScriptLocked() {
  const props = PropertiesService.getScriptProperties();
  const locked = props.getProperty("LOCKED");
  if (locked !== "true") {
    return false; // not locked
  }

  const stamp = parseInt(props.getProperty("LOCK_TIMESTAMP") || "0", 10);
  if (!stamp) {
    // Missing stamp => auto-unlock
    unlockScript();
    return false;
  }

  const now = Date.now();
  const elapsed = now - stamp;
  const FIVE_MINUTES = 5 * 60 * 1000;
  if (elapsed > FIVE_MINUTES) {
    // Stale lock => override
    unlockScript();
    return false;
  }

  // Otherwise still locked + not timed out
  return true;
}

function lockScript() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("LOCKED", "true");
  props.setProperty("LOCK_TIMESTAMP", Date.now().toString());
}

function unlockScript() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("LOCKED");
  props.deleteProperty("LOCK_TIMESTAMP");
}
