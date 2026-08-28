/**
 * Time Tracking sync: Google Calendar -> Google Sheet
 *
 * SETUP (one time):
 * 1. Create a Google Sheet (any name). Add two tabs: "Time Log" and "Buckets".
 * 2. In "Buckets", list your allowed bucket names, one per row, column A:
 *      Networking
 *      Applications
 *      Interview
 *      InterviewPrep
 *      BusinessCase
 *      Research
 *      Learning
 *      Workout
 *      Break
 *      Other
 *    (Edit this list any time — it's just used to catch typos, not to limit what you type.
 *    Adding a new bucket later needs no code changes, just a new row here.)
 * 3. In the Sheet, go to Extensions > Apps Script. Delete the placeholder code
 *    and paste this whole file in. Save (the little floppy disk icon).
 * 4. In the function dropdown at the top, choose "createNightlyTrigger" and
 *    click Run. Google will ask you to authorize access to your Calendar —
 *    approve it (this is a one-time permission you're granting yourself,
 *    inside your own Google account).
 * 5. Also run "syncTimeTracking" once manually to pull in your first batch of data.
 *
 * From then on, it runs automatically every night at 11pm and refreshes
 * the last 180 days of events into "Time Log". You can also refresh any
 * time you like via the "Time Tracking > Sync Now" menu in the Sheet.
 *
 * TAGGING CONVENTION (in your calendar event titles):
 *   One hashtag per event, e.g. "Coffee with Sam #Networking"
 *   Keep bucket names as single words (no spaces) — use #InterviewPrep,
 *   not #Interview Prep.
 *   Rule of thumb: talking to a person -> #Networking. Solo desk work
 *   (market/company/role research) -> #Research. Actually being
 *   interviewed -> #Interview. Prepping for it beforehand -> #InterviewPrep
 *   (or #BusinessCase if it's specifically case-prep).
 */

const CALENDAR_NAME = "Time Tracking";
const LOG_SHEET_NAME = "Time Log";
const BUCKETS_SHEET_NAME = "Buckets";
const DAYS_TO_SYNC = 180;

function syncTimeTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  const bucketsSheet = ss.getSheetByName(BUCKETS_SHEET_NAME);

  const validBuckets = bucketsSheet
    ? bucketsSheet
        .getRange(1, 1, Math.max(bucketsSheet.getLastRow(), 1), 1)
        .getValues()
        .flat()
        .filter(String)
        .map(String)
    : [];

  const calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
  if (!calendar) {
    throw new Error(`No calendar named "${CALENDAR_NAME}" found. Check the name matches exactly.`);
  }

  const now = new Date();
  const start = new Date(now.getTime() - DAYS_TO_SYNC * 24 * 60 * 60 * 1000);
  const events = calendar.getEvents(start, now);

  const rows = events
    .filter((e) => !e.isAllDayEvent())
    .map((e) => {
      const title = e.getTitle();
      const match = title.match(/#(\w+)/);
      const bucket = match ? match[1] : "Untagged";

      const durationHrs = (e.getEndTime() - e.getStartTime()) / (1000 * 60 * 60);
      const isKnownBucket = bucket === "Untagged" || validBuckets.length === 0 || validBuckets.includes(bucket);

      return [
        e.getStartTime(),
        title,
        e.getStartTime(),
        e.getEndTime(),
        durationHrs,
        bucket,
        isKnownBucket ? "" : "check spelling",
      ];
    });

  logSheet.clearContents();
  logSheet.appendRow(["Date", "Event Title", "Start", "End", "Duration (hrs)", "Bucket", "Flag"]);
  if (rows.length) {
    logSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function createNightlyTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "syncTimeTracking") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("syncTimeTracking").timeBased().everyDays(1).atHour(23).create();
}

/**
 * Adds a "Time Tracking > Sync Now" menu to the Sheet so you can refresh
 * on demand, any time, without opening the Apps Script editor.
 * Runs automatically whenever you open the Sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Time Tracking").addItem("Sync Now", "syncTimeTracking").addToUi();
}
