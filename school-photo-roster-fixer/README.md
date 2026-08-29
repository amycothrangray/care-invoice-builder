# School Photo Roster Fixer

Drop in whatever spreadsheet the school sent. Get back one clean row per
student, ready to import into GotPhoto.

Everything runs in the browser. The roster never leaves your computer, and the
file you upload is never modified.

**The one rule:** formatting problems get fixed automatically, identity problems
get handed back to you. The tool will never quietly turn two children into one
row, and it will never throw away a parent's contact details.

---

## Deploy it (about 60 seconds)

1. Go to https://app.netlify.com/drop
2. Drag **this folder** onto the page
3. Netlify gives you a URL. Bookmark it.

To update later, drag the folder onto the same site's "Deploys" page.

It is four static files with no build step, so any static host works — and
double-clicking `index.html` works too, as long as you are online (SheetJS and
ExcelJS load from a CDN).

---

## Using it

**Upload → check the columns → fix → review a handful → download.**

### 1. Upload

`.xlsx`, `.xls` or `.csv`. Title rows, logos, blank rows, repeated headers,
footer totals, merged cells, hidden rows, formulas and multiple tabs are all
expected.

### 2. Check the columns

The app shows every tab it found, which one looks like a roster, which row it
believes is the header, and what it thinks each column means:

| School column | Exported as |
| --- | --- |
| StuFirst | Student First Name |
| GradeLvl | Grade |
| HR Teacher | Teacher |
| PermID | Student ID |
| MomEmail | Parent Email 1 |
| DadCell | Parent Cell 2 |

Change anything that is wrong before it processes — school information systems
name things wildly differently, and this is the screen that absorbs that.

A tab named `Grade 3` with no grade column will offer to fill Grade = 3. A tab
named `Sheet1` will not offer anything.

### 3. Fix & review

A dashboard of what was fixed automatically, what needs you, and which columns
are being left out. Then the questions — each one a card with the two records
side by side and buttons like **Merge as same student** / **Keep as two
students**.

### 4. Download

`[Original File Name] - CLEANED.xlsx`

| Sheet | What is on it |
| --- | --- |
| **GOTPHOTO READY** | the cleaned roster, one row per student, every cell text |
| **NEEDS REVIEW** | every issue with the student, source sheet, source row and an explanation |
| **CLEANUP REPORT** | counts of everything that changed, plus the columns left out |
| **SOURCE TRACE** | optional: every student, the rows behind them, and what changed |

---

## When two rows are the same child

This is the whole point of the tool. Schools export one row per **parent**, so a
child in two households arrives twice:

| Student | Parent | Email |
| --- | --- | --- |
| Emma Smith | Jane Smith | jane@email.com |
| Emma Smith | John Smith | john@email.com |

That is one student with two buyers, not two students.

**Merged automatically:**

- identical rows (reported, not silently dropped)
- rows sharing a valid **Student ID** — the primary identity key
- rows with no ID that match on **first + last + grade + teacher**, where every
  one of those the file actually provides is filled in

**Never merged automatically — you decide:**

- same name and grade, **different teacher**
- same name and grade, teacher blank on one of them when the file has teachers
- similar-but-not-identical names (`Sofia` / `Sophia`, nicknames, typos)
- the same name in two different grades
- the same name with two different Student IDs

Same last name is never treated as evidence — siblings stay separate.

One Student ID carrying two spellings (`Isabella` / `Isabela`) stays a single
student, because the ID says so, but the name discrepancy is flagged and you
choose which spelling to keep — or split them back apart.

Parent contacts are matched on the normalised email and phone, so
`(619) 555-1234` and `619-555-1234` are one number, and `MOM@GMAIL.COM` and
`mom@gmail.com` are one person.

---

## What it fixes without asking

| | |
| --- | --- |
| Header row | finds the real one under the school name, report title and blank rows |
| Junk rows | blank rows, `Total Students: 427`, page numbers, repeated headers, footers |
| Names | `Smith, Emma` → Emma / Smith; SHOUTING → Shouting; keeps `DeShawn`, `McRae`, `van der Berg`, `O'Connor`, hyphens, accents, `Jr.` |
| Grade | `Kindergarten` `Kinder` `KG` → `K`; `01` `1st` `Grade 1` `First Grade` → `1`; PS, PK, TK, K, 1–12 |
| Teacher | spacing and case, `Mrs` → `Mrs.`, `SMITH, JENNIFER` → `Jennifer Smith` |
| Student ID | always text — `0018472` stays `0018472`, never `18472` and never `1.8472E+04` |
| Email | trims, folds case, repairs `amy @ gmail.com`, drops `none`/`N/A` |
| Phone | `+1 (619) 555-1212 x204` → `(619) 555-1212 x204`, extensions kept |
| Two in one cell | `amy@e.com; john@e.com` becomes Parent 1 and Parent 2 |
| Spreadsheet mess | merged cells filled down, formulas flattened, non-breaking spaces, smart quotes, line breaks, zero-width characters |

Unknown grades (`2026`, `Room 12 Enrichment`) are kept exactly as written and
flagged. Nothing is ever guessed into a real grade number.

## What it asks about

Possible duplicate students · one ID with two names · ambiguous name splits
(`Mary Kate Johnson`) · invalid emails and phones · missing names, grades,
teachers and IDs (when the rest of the roster has them) · rows with no parent
contact at all.

**Parent phone with no email** gets its own warning: GotPhoto's names-list
import needs an email address for a buyer contact, so a phone-only parent will
not import. No dummy addresses are ever invented.

Invalid emails and phones are held out of the export until they are right —
they are on the Needs Review sheet and editable right in the review screen, so
nothing is lost.

## What it leaves out

Date of birth, home address, medical, lunch status, IEP/504, gender,
race/ethnicity, discipline, logins and internal notes are listed under "not
exported" rather than shipped to GotPhoto. Photo/media release columns are
called out separately in case you want one. Any of them can be switched on from
the mapping screen.

Emergency contacts are **not** treated as buyers unless you say they are.

---

## The files

| | |
| --- | --- |
| `index.html` | the page and all of its styling |
| `lib.js` | the whole cleanup engine — no DOM, no globals, testable on its own |
| `app.js` | reads the file, draws the screen, feeds your choices back to the engine |
| `test/` | the rules that are not allowed to break |

The engine is one pure function from (parsed workbook + your choices) to
everything on screen and in the download. Every checkbox, mapping change and
merge decision re-runs it from scratch, so nothing has to be undone and the
download always matches what you are looking at.

## Tests

```
npm install     # xlsx + exceljs, only needed for the tests
npm test
```

The suite builds deliberately awful workbooks — title rows, camel-case headers,
split households, an exact duplicate, a repeated header row, a totals footer,
leading-zero IDs, two emails in one cell, non-breaking spaces, an emergency
contact, a `2026` grade — and asserts what comes out the other end, including
the finished .xlsx.
