# Sitterwise · Care.com Invoice Builder

Drop the monthly exports, answer the judgment calls, download the finished
invoice — official template layout, live formulas, full formatting.

Everything runs in the browser. The spreadsheets never leave your computer.

## Deploy to Netlify (about 60 seconds)

1. Go to https://app.netlify.com/drop (log in if asked)
2. Drag this whole folder onto the page
3. Netlify gives you a URL — that's the app. Bookmark it.

To update later: drag the folder onto the same site's "Deploys" page.

## Pulling the exports

> **The Care.com export must include BOTH cancelled and active/completed jobs.**
> Don't filter it to one status when you pull it from the Care portal.

The cancelled rows are not optional detail — they are what the invoice is built
from:

| What the cancelled rows give you | Without them |
| --- | --- |
| Care.com's own 7-digit job ID for each cancellation | you send Care.com a Sitterwise booking number and they bounce it |
| A cancellation **timestamp**, so >24 hr vs <24 hr is measured | every cancellation defaults to the $30 flat fee and you guess the rest |
| Proof of reassignment "shadows" — the same job ID coming back both Cancelled and Done | the app falls back to a client+date+time guess |
| Proof that a multi-day booking was cancelled all at once (one fee, not one per day) | consecutive days are only flagged, never resolved |

A completed-only export still builds an invoice — the app says so in orange at
the top of step 2 — but every cancellation falls back to the bookings file with
no job ID and no notice time.

The bookings export needs no special handling: pull the whole month.

## Using it each month

1. **Upload** the Care.com export (`CareAtWorkBackupCareJobs…xlsx`, cancelled +
   active/completed) and the bookings export (`bookings-Month-Year.xlsx`)
2. **Answer the questions** — the app applies every rule automatically
   (CA daily overtime, 4-hour minimum, weekly-OT check, reassignment-shadow
   and couldn't-fill exclusion, multi-day cancellation grouping) and only
   asks about the true judgment calls:
   - rate / mileage rate / invoice number (pre-filled, editable)
   - any client names it couldn't match
   - which mileage reimbursements to include, and the billable miles
   - how to bill each cancellation — now **pre-set from the measured notice**
     (>24hr flat · <24hr hours · no charge), so you're confirming rather than
     guessing. Each row shows where it came from: `✓ Care` means Care.com listed
     the cancellation and the ID is theirs; `⚠ not in Care export` means only
     Sitterwise has it, so look their job ID up before sending.
3. **Download** the .xlsx and send it to Care.com yourself.

The receipt tape on the right shows the running total as you answer.
The spreadsheet's own formulas do the final math.

**Mileage:** upload the caregiver Mileage Request form export (Cognito) as an
optional third file — it is the source of truth for billable miles. The app
matches each submission to its invoice job, pre-fills the over-40 miles,
auto-skips jobs that also carry a bonus (Care.com won't pay both unless
pre-approved), and explains every submission it leaves off. Without the form,
mileage falls back to reimbursement-based estimates.

If the bookings export has an **Admin Notes** column, each note appears right
next to the judgment call it explains (cancellations, mileage, unmatched
clients), plus a full list under "All admin notes this month." Notes are
internal guidance only — they are never printed on the invoice.

## Rules baked in

- $42/hr regular · 1.5× CA overtime over 8 hrs/day (rate editable)
- 4-hour minimum on short bookings
- Mileage over 40-mile round trip at $0.76/mi (rate editable)
- Cancellations: $30 flat for >24 hr notice; booked hours (capped at 8) at the
  hourly rate for <24 hr; one fee for multi-day or double-staffed bookings
- Bonuses from the Care.com export, never duplicated onto overtime lines
- Tips are never billed to Care.com (flagged for OnPay instead)
- Lifesaver bonuses are a Sitterwise caregiver incentive, never a Care.com
  charge — they're listed under "what was excluded" so you can see them
- State comes from the Care export's city/state/zip, so out-of-state jobs
  (e.g. a Miami Beach booking) aren't stamped `CA`

## Export formats

Both exports gained columns in August 2026. Columns are matched **by header
name**, with aliases and fallbacks, so older files still load:

| Care.com export | |
| --- | --- |
| `Status` | new — `Done` / `Cancelled`. Splits worked jobs from the cancellation block. |
| `Cancellation Date` | now populated with a timestamp, which sets the >24/<24 fee. |

| Bookings export | |
| --- | --- |
| `Total Hours` | now `Hours Worked` + `Hours Billed`. Billed is what gets billed; any of the three names works. |
| `Care.com Job Number` | new — an exact join to the Care export for client names, replacing the old caregiver+date+time guess. One cell may list every day of a multi-day booking. |
| `Mileage Approval Status`, `Payable Miles`, `Mileage Approved Miles`, `Round Trip Miles`, `Mileage Amount` | new — used as the mileage source when populated, ahead of the Cognito form and the reimbursement estimate. |
| `Lifesaver Bonus` | new — surfaced as excluded, never billed. |
| `Minimum Applied` | new — read alongside the app's own 4-hour-minimum check. |

If a future export renames something again, the app fails with the missing
column's name rather than silently billing the wrong number.
