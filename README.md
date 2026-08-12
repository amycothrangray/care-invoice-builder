# Sitterwise · Care.com Invoice Builder

Drop the monthly exports, answer the judgment calls, download the finished
invoice — official template layout, live formulas, full formatting.

Everything runs in the browser. The spreadsheets never leave your computer.

## Deploy to Netlify (about 60 seconds)

1. Go to https://app.netlify.com/drop (log in if asked)
2. Drag this whole folder onto the page
3. Netlify gives you a URL — that's the app. Bookmark it.

To update later: drag the folder onto the same site's "Deploys" page.

## Using it each month

1. **Upload** the Care.com export (`CareAtWorkBackupCareJobs…xlsx`) and the
   bookings export (`bookings-Month-Year.xlsx`)
2. **Answer the questions** — the app applies every rule automatically
   (CA daily overtime, 4-hour minimum, weekly-OT check, reassignment-shadow
   and couldn't-fill exclusion, multi-day cancellation grouping) and only
   asks about the true judgment calls:
   - rate / mileage rate / invoice number (pre-filled, editable)
   - any client names it couldn't match
   - which mileage reimbursements to include, and the billable miles
   - how to bill each cancellation (>24hr flat · <24hr hours · no charge)
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
