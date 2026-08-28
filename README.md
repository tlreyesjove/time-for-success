# time-for-success (Google Calendar → Sheets → Looker Studio)

A lightweight, personal time-tracking system for job-search "hustle time." No third-party apps —
just Google Calendar, a Google Sheet, and Looker Studio.

## How it works

1. Time is blocked/logged on a dedicated **"Time Tracking"** Google Calendar (kept separate from
   the main calendar so it can be toggled off and doesn't clash with shared/booking events).
2. Each event title ends with one hashtag naming what the time was spent on, e.g.:
   `Coffee with Sam #Networking`
3. A Google Apps Script (`time_tracking_sync.gs`), bound to a Google Sheet, runs nightly. It reads
   the calendar, pulls the hashtag out of each event, and writes one row per event into a
   **"Time Log"** tab: date, title, start, end, duration in hours, and bucket.
4. A second **"Buckets"** tab in the same Sheet is the official, editable list of valid bucket
   names. If an event's hashtag doesn't match anything on that list (a typo like `#breek` instead
   of `#break`), the script flags that row in a `Flag` column so it's easy to spot and fix.
5. Looker Studio connects to the "Time Log" tab for the actual dashboard (hours by bucket, weekly
   trend, bucket share over time).

## Buckets (current list)

- `#Networking` — talking to a person (informational interviews, scoping calls)
- `#Applications` — job applications
- `#Interview` — actually being interviewed
- `#InterviewPrep` — general interview prep
- `#BusinessCase` — case-specific prep
- `#Research` — solo desk research (market/company/role)
- `#Learning` — professional development / upskilling
- `#Consulting` — billable hours work
- `#Workout`
- `#Break`
- `#Other`

Add or remove buckets any time by editing the "Buckets" tab in the Sheet — no code changes needed.

## Setup

1. Create a Google Calendar named **"Time Tracking"**.
2. Create a Google Sheet with two tabs: **"Time Log"** and **"Buckets"**. Put the bucket list above
   into "Buckets", one per row, column A.
3. In the Sheet: **Extensions → Apps Script**, paste in `time_tracking_sync.gs`, save.
4. Run `createNightlyTrigger` once from the Apps Script editor (this prompts a one-time Google
   authorization). Also run `syncTimeTracking` once manually to pull in the first batch of data.
5. In Looker Studio, add a data source pointing at the "Time Log" tab and build charts from there.

Refreshes automatically every night, or any time on demand via the **Time Tracking > Sync Now**
menu that appears in the Sheet.

## Screenshots

_To add once the Sheet and dashboard are live._
