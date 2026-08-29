/* Sitterwise → Care.com Invoice Builder
   All logic runs client-side. Rules mirror the monthly workflow:
   $42/hr (editable) · CA daily OT 1.5× over 8h/day · 4-hr minimum ·
   mileage over 40mi RT at $0.76/mi (editable) · cancellation reconciliation
   with reassignment-shadow + couldn't-fill exclusion and multi-day grouping.

   ── PULLING THE CARE.COM EXPORT ────────────────────────────────────────────
   The Care.com export MUST include BOTH cancelled AND active/completed jobs.
   Do not filter it to one status when you pull it from the Care portal.
   Cancelled rows are what give us Care.com's own job IDs and — since the
   "Cancellation Date" column carries a timestamp — the >24hr / <24hr notice
   window, which used to be a manual guess every single month. A Care export
   pulled with only completed jobs still builds, but every cancellation falls
   back to the bookings file: no Care job ID, no notice window, no automatic
   reassignment-shadow proof. See README.md → "Pulling the exports".

   Both exports gained columns (Aug 2026). Columns are resolved by HEADER NAME
   with aliases, never by fixed position, so older exports still load:
     Care export ....... + Status, + Cancellation Date
     bookings export ... Total Hours → Hours Worked / Hours Billed,
                         + Care.com Job Number (exact client-name join),
                         + Minimum Applied, + Round Trip Miles,
                         + Mileage Approved Miles, + Mileage Approval Status,
                         + Payable Miles, + Mileage Amount, + Lifesaver Bonus */

"use strict";

const S = {
  careFile: null, bookFile: null, milFile: null, xeroFile: null,
  jobs: [],            // completed/active Care.com jobs (Status != Cancelled)
  cancBillable: [],    // billable cancellations (with UI decisions)
  cancExcluded: [],    // {why, label}
  unmatched: [],       // jobs needing client names
  mileageCand: [],     // mileage candidates on the invoice
  month: null, year: null,
  weeklyOTWarn: [],
  adminNotes: [],      // internal guidance only — never written to the invoice
  dblBill: [],         // completed jobs that may duplicate a billed cancellation
  milSkipped: [],      // mileage form entries that don't belong on this invoice
  milCol: "", milBlank: 0,
  milSource: "none",   // "book" (bookings mileage cols) · "form" (Cognito) · "reimb" (estimate)
  lifesavers: [],      // Sitterwise lifesaver bonuses — never billed to Care.com
  careHasStatus: false,// did the Care export carry a Status column?
  oddStatus: [],       // Care rows with a status that is neither done/completed nor cancelled
  excludeJobs: new Set(), // completed Care jobs excluded from the invoice (not actually worked)
  built: null,         // {blob, filename, total}
  clientInvoices: [],  // step 4 — one draft per non-Care.com Xero contact
  clientSkipped: [],   // non-Care rows with no matching Xero client
};

const $ = id => document.getElementById(id);
const esc = s => String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const money = n => "$" + n.toLocaleString("en-US", {minimumFractionDigits: 2, maximumFractionDigits: 2});
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];

/* ---------------- file handling ---------------- */
function wireDrop(dropId, inputId, nameId, key) {
  const drop = $(dropId), input = $(inputId);
  input.addEventListener("change", () => {
    if (input.files[0]) {
      S[key] = input.files[0];
      $(nameId).textContent = input.files[0].name;
      drop.classList.add("filled");
      $("btn-analyze").disabled = !(S.careFile && S.bookFile);
    }
  });
  ["dragover","dragenter"].forEach(ev => drop.addEventListener(ev, e => e.preventDefault()));
  drop.addEventListener("drop", e => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) { input.files = e.dataTransfer.files; input.dispatchEvent(new Event("change")); }
  });
}
wireDrop("drop-care","file-care","fname-care","careFile");
wireDrop("drop-book","file-book","fname-book","bookFile");
wireDrop("drop-mil","file-mil","fname-mil","milFile");
wireDrop("drop-xero","file-xero","fname-xero","xeroFile");

function readText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result);
    r.onerror = () => rej(new Error("Could not read " + file.name));
    r.readAsText(file);
  });
}
function readWB(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => { try { res(XLSX.read(new Uint8Array(e.target.result), {type:"array", cellDates:true})); } catch(err){ rej(err); } };
    r.onerror = () => rej(new Error("Could not read " + file.name));
    r.readAsArrayBuffer(file);
  });
}

/* ---------------- normalizers ---------------- */
function isoDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    return v.getFullYear() + "-" + String(v.getMonth()+1).padStart(2,"0") + "-" + String(v.getDate()).padStart(2,"0");
  }
  const s = String(v ?? "").trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[1]+"-"+m[2]+"-"+m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) { let y = m[3].length===2 ? "20"+m[3] : m[3]; return y+"-"+String(m[1]).padStart(2,"0")+"-"+String(m[2]).padStart(2,"0"); }
  return null;
}
function hhmm(v) { // -> "HH:MM" 24h
  if (v instanceof Date && !isNaN(v)) return String(v.getHours()).padStart(2,"0")+":"+String(v.getMinutes()).padStart(2,"0");
  if (typeof v === "number") { const t = Math.round(v*1440); return String(Math.floor(t/60)%24).padStart(2,"0")+":"+String(t%60).padStart(2,"0"); }
  const s = String(v ?? "").trim();
  let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm|AM|PM)?/);
  if (!m) return null;
  let h = parseInt(m[1]), ap = (m[3]||"").toLowerCase();
  if (ap==="pm" && h!==12) h+=12; if (ap==="am" && h===12) h=0;
  return String(h).padStart(2,"0")+":"+m[2];
}
function money2num(v) {
  if (typeof v === "number") return v;
  const n = parseFloat(String(v ?? "").replace(/[$,\s]/g,""));
  return isNaN(n) ? 0 : n;
}
const firstName = s => (String(s||"").trim().toLowerCase().split(/\s+/)[0] || "");
const lastName = s => { const p = String(s||"").trim().split(/\s+/); return p[p.length-1] || ""; };
const cleanName = s => String(s||"").replace(/\s+/g," ").replace(/\s*\(NULL\)\s*/i,"").trim();
const CG_FIX = {"Jacque Sniff":"Jacquelyn Sniff"};

/* Care export "Cancellation Date" cell: "07/30/2026 03:57pm" -> Date (local) */
function parseStamp(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v ?? "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:[ ,T]+(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm)?)?/i);
  if (m) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    let h = m[4] ? +m[4] : 0; const ap = (m[6]||"").toLowerCase();
    if (ap === "pm" && h !== 12) h += 12; if (ap === "am" && h === 12) h = 0;
    return new Date(y, +m[1]-1, +m[2], h, m[5] ? +m[5] : 0);
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (iso) return new Date(+iso[1], +iso[2]-1, +iso[3], iso[4]?+iso[4]:0, iso[5]?+iso[5]:0);
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
/* local Date for an ISO date + "HH:MM" */
const atTime = (iso, hm) => { const [Y,M,D]=iso.split("-").map(Number); const [h,m]=(hm||"00:00").split(":").map(Number); return new Date(Y,M-1,D,h,m); };
/* "Chula Vista, CA 91911" -> "CA" */
const stateOf = v => { const m = String(v||"").match(/,\s*([A-Za-z]{2})\s+\d{5}/); return m ? m[1].toUpperCase() : "CA"; };
/* A Care.com Job Number cell can hold several ids, comma/tab separated, and the
   occasional stray date ("8/17, 5586324, 8/18, 5586328"). Keep the 6-9 digit ids. */
const jobNums = v => String(v ?? "").split(/[^0-9]+/).filter(t => /^\d{6,9}$/.test(t));
/* header lookup with aliases: first header that matches any of the patterns */
function colFinder(headerRow) {
  const hdr = (headerRow||[]).map(h => String(h ?? "").replace(/\s+/g," ").trim().toLowerCase());
  return (...pats) => {
    for (const p of pats) { const i = hdr.findIndex(h => h === p); if (i > -1) return i; }
    for (const p of pats) { const i = hdr.findIndex(h => h.includes(p)); if (i > -1) return i; }
    return -1;
  };
}

/* Care export time cell: "07:30am - 05:30pm (10)" */
function parseCareTimes(s) {
  const m = String(s||"").match(/(\d{2}):(\d{2})(am|pm)\s*-\s*(\d{2}):(\d{2})(am|pm)\s*\(([\d.]+)\)/i);
  if (!m) return null;
  const to24 = (h,mn,ap) => { h=parseInt(h); ap=ap.toLowerCase(); if(ap==="pm"&&h!==12)h+=12; if(ap==="am"&&h===12)h=0; return String(h).padStart(2,"0")+":"+mn; };
  return { start: to24(m[1],m[2],m[3]), end: to24(m[4],m[5],m[6]), hrs: parseFloat(m[7]) };
}
const addMin = (t, mins) => { const [h,m]=t.split(":").map(Number); const x=(h*60+m+mins)%1440; return String(Math.floor(x/60)).padStart(2,"0")+":"+String(x%60).padStart(2,"0"); };
const fmt12 = t => { let [h,m]=t.split(":").map(Number); const ap=h<12?"AM":"PM"; h=h%12||12; return h+":"+String(m).padStart(2,"0")+" "+ap; };
const timeFrac = t => { const [h,m]=t.split(":").map(Number); return (h*60+m)/1440; };
function isoWeek(dstr){ const d=new Date(dstr+"T12:00:00"); const t=new Date(d); t.setDate(d.getDate()+3-((d.getDay()+6)%7)); const w1=new Date(t.getFullYear(),0,4); return t.getFullYear()+"-W"+(1+Math.round(((t-w1)/86400000-3+((w1.getDay()+6)%7))/7)); }

/* ---------------- analysis ---------------- */
$("btn-analyze").addEventListener("click", async () => {
  const errBox = $("err-upload"); errBox.style.display="none";
  try {
    const [careWB, bookWB, milWB, xeroTxt] = await Promise.all([
      readWB(S.careFile), readWB(S.bookFile),
      S.milFile ? readWB(S.milFile) : Promise.resolve(null),
      S.xeroFile ? readText(S.xeroFile) : Promise.resolve(null)]);
    analyze(careWB, bookWB, milWB, xeroTxt);
  } catch(e) {
    errBox.textContent = "Couldn't read the files: " + e.message + " — make sure these are the two monthly .xlsx exports.";
    errBox.style.display = "block";
  }
});

function sheetRows(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
}

/* The Cognito mileage-request export. Column names on this form have changed
   before, so match them generously — and when the miles column genuinely can't
   be found, say so with the headers we DID see. Silently defaulting the miles to
   0 is how a whole month of mileage reached the invoice blank: the job numbers
   still matched, so every row looked like a healthy "form submission". */
function parseMileageForm(wb) {
  const rows = sheetRows(wb);
  if (!rows.length) throw new Error("the mileage file has no rows");
  // the header row isn't always row 1 (Cognito can emit a title/filter row first)
  let hi = rows.findIndex(r => (r||[]).some(c => /job/i.test(String(c ?? ""))));
  if (hi === -1) hi = 0;
  const raw = (rows[hi]||[]).map(h => String(h ?? "").trim());
  // Cognito exports field NAMES, not labels: "TotalMilesAbove40", "CarecomJob",
  // "Name_First". Split camelCase and underscores so word-boundary patterns can
  // match at all — matching the raw string is what made every miles column miss.
  const hdr = raw.map(h => h.replace(/_/g," ")
                            .replace(/(?<=[a-z0-9])(?=[A-Z])/g," ")
                            .replace(/(?<=[A-Za-z])(?=\d)/g," ")
                            .replace(/\s+/g," ").trim().toLowerCase());
  const find = (...pats) => { for (const re of pats) { const i = hdr.findIndex(h => re.test(h)); if (i > -1) return i; } return -1; };

  const jc = find(/care\.?com job/, /job\s*#/, /job number/, /job/);
  if (jc === -1) throw new Error(`mileage file: no job-number column. Headers found: ${raw.filter(Boolean).join(" | ") || "(none)"}`);
  const nc  = find(/caregiver name/, /your name/, /^name$/, /name first/, /first name/, /name/);
  const nc2 = find(/name last/, /last name/);
  const ac = find(/amount/, /reimburse/, /\$/);
  // miles, best first: an explicit over-40 figure, then a billable/payable figure,
  // then a round-trip total we can subtract the 40-mile threshold from.
  const oc = find(/(?:over|above)\s*40/, /miles (?:over|above)/, /over/, /above/, /billable/, /payable/, /excess/);
  const tc = find(/total number/, /number of miles/, /total miles/, /round.?trip/, /^miles$/, /miles driven/, /mileage/, /miles/);
  if (oc === -1 && tc === -1)
    throw new Error(`mileage file: found the job column but no miles column, so every job would import as 0 miles. Headers found: ${raw.filter(Boolean).join(" | ")}`);

  const used = oc > -1 ? raw[oc] : raw[tc] + " (minus the 40-mile threshold)";
  const out = []; let blank = 0;
  for (let i=hi+1;i<rows.length;i++){
    const r = rows[i]; if (!r || r[jc]==null || String(r[jc]).trim()==="") continue;
    const total = tc>-1 ? money2num(r[tc]) : 0;
    let over = oc>-1 ? money2num(r[oc]) : 0;
    if (!over && total) over = Math.max(Math.round(total-40), 0);
    if (!over) blank++;
    const nm = [nc>-1?r[nc]:"", nc2>-1?r[nc2]:""].filter(Boolean).join(" ");
    out.push({ job:String(r[jc]).trim(), name:cleanName(nm), over, total, amt: ac>-1?money2num(r[ac]):0 });
  }
  if (out.length && blank === out.length)
    throw new Error(`mileage file: read ${out.length} submission(s) from “${used}” but every one came out at 0 billable miles — that column doesn't hold the mileage. Headers found: ${raw.filter(Boolean).join(" | ")}`);
  out.milesCol = used; out.blank = blank;
  return out;
}

function analyze(careWB, bookWB, milWB, xeroTxt) {
  const care = sheetRows(careWB);
  const book = sheetRows(bookWB);

  // ---- bookings columns by header name (aliases: the export gained columns Aug 2026) ----
  const bc = colFinder(book[0]);
  const IDX = {
    bk:     bc("booking id"),
    jobno:  bc("care.com job number","care.com job #","care job number"),
    client: bc("client name"),
    email:  bc("client email"),
    svc:    bc("service type"),
    date:   bc("start date"),
    start:  bc("start time"),
    end:    bc("end time"),
    // "Total Hours" was split into Hours Worked / Hours Billed. Billed is what we bill.
    hrsB:   bc("hours billed"),
    hrsW:   bc("hours worked","total hours"),
    minApp: bc("minimum applied"),
    cg:     bc("caregiver name"),
    status: bc("status"),
    reimb:  bc("reimbursement"),
    reimbD: bc("reimbursement description"),
    rtMi:   bc("round trip miles"),
    apprMi: bc("mileage approved miles"),
    milSt:  bc("mileage approval status"),
    payMi:  bc("payable miles"),
    milAmt: bc("mileage amount"),
    tip:    bc("tip"),
    bonus:  bc("bonus"),
    life:   bc("lifesaver bonus"),
    note:   bc("admin notes"),
  };
  for (const [k,label] of [["bk","Booking ID"],["client","Client Name"],["svc","Service Type"],["date","Start Date"],["start","Start Time"],["status","Status"]])
    if (IDX[k] === -1) throw new Error(`bookings file is missing the "${label}" column`);
  if (IDX.hrsB === -1 && IDX.hrsW === -1) throw new Error('bookings file has no hours column ("Hours Billed", "Hours Worked" or "Total Hours")');

  const at = (r, i) => i > -1 ? r[i] : null;
  const bRows = [];
  for (let i=1;i<book.length;i++){
    const r = book[i]; if (!r || r.every(c=>c===null||c==="")) continue;
    const hrsB = money2num(at(r,IDX.hrsB)), hrsW = money2num(at(r,IDX.hrsW));
    bRows.push({
      bk: at(r,IDX.bk), jobs: jobNums(at(r,IDX.jobno)), jobRaw: String(at(r,IDX.jobno) ?? "").trim(),
      client: cleanName(at(r,IDX.client)), email: String(at(r,IDX.email) ?? "").trim(),
      svc: String(at(r,IDX.svc)||"").trim(),
      date: isoDate(at(r,IDX.date)), start: hhmm(at(r,IDX.start)), end: hhmm(at(r,IDX.end)),
      hrs: hrsB || hrsW, hrsWorked: hrsW || hrsB,
      minApplied: /^(true|yes|1)$/i.test(String(at(r,IDX.minApp) ?? "").trim()),
      cg: cleanName(at(r,IDX.cg)), status: String(at(r,IDX.status)||"").trim().toLowerCase(),
      reimb: money2num(at(r,IDX.reimb)), reimbDesc: String(at(r,IDX.reimbD) ?? "").replace(/\s+/g," ").trim(),
      rtMiles: money2num(at(r,IDX.rtMi)), apprMiles: money2num(at(r,IDX.apprMi)),
      milStatus: String(at(r,IDX.milSt) ?? "").trim().toLowerCase(),
      payMiles: money2num(at(r,IDX.payMi)), milAmt: money2num(at(r,IDX.milAmt)),
      tip: money2num(at(r,IDX.tip)), bonus: money2num(at(r,IDX.bonus)), lifesaver: money2num(at(r,IDX.life)),
      note: String(at(r,IDX.note) ?? "").replace(/\s+/g," ").trim(),
    });
  }
  const corp = bRows.filter(r => r.svc === "Corporate (Invoiced)");

  // client lookup. Best key by far is Care.com's own job number, which the bookings
  // export now carries (one cell can list every day of a multi-day booking).
  const byJob = new Map();
  for (const r of bRows) for (const j of r.jobs) if (!byJob.has(j)) byJob.set(j, r);
  // fallbacks for older exports / rows with no job number: (date|start), then (cgLast|date)
  const byDT = new Map(), byCgDate = new Map();
  for (const r of bRows) {
    if (!r.date || !r.start) continue;
    const k = r.date+"|"+r.start;
    if (!byDT.has(k)) byDT.set(k, []);
    byDT.get(k).push(r);
    if (r.cg) {
      const k2 = lastName(r.cg).toLowerCase()+"|"+r.date;
      if (!byCgDate.has(k2)) byCgDate.set(k2, []);
      byCgDate.get(k2).push(r);
    }
  }

  // ---- Care export rows (headers, with the old fixed positions as a fallback) ----
  // The export now carries BOTH cancelled and active/completed jobs — the Status
  // column decides which pile a row lands in. Cancelled rows are billed from the
  // CANCELLATIONS block, never as worked hours.
  const cc = colFinder(care[0]);
  const CIDX = {
    jid:  cc("job id","job #") , date: cc("date"), times: cc("hours (local time)","hours"),
    multi: cc("multi day","multi-day"), loc: cc("city, state zip","city"),
    bonus: cc("bonus"), cg: cc("caregiver"), status: cc("status"), canc: cc("cancellation date"),
  };
  const CFALLBACK = {jid:0,date:1,times:2,multi:3,loc:4,bonus:5,cg:6,status:7,canc:8};
  for (const k in CIDX) if (CIDX[k] === -1) CIDX[k] = CFALLBACK[k];
  S.careHasStatus = cc("status") !== -1;

  const jobs = []; const unmatched = []; const careCanc = [];
  const doneIds = new Set(), cancIds = new Set();
  S.oddStatus = [];
  for (let i=1;i<care.length;i++){
    const r = care[i]; if (!r || r[CIDX.jid]===null || String(r[CIDX.jid]).trim()==="") continue;
    const jid = String(r[CIDX.jid]).trim();
    const date = isoDate(r[CIDX.date]); const t = parseCareTimes(r[CIDX.times]);
    if (!date || !t) continue;
    let cg = cleanName(r[CIDX.cg]); cg = CG_FIX[cg] || cg;
    const bonus = money2num(r[CIDX.bonus]);
    const status = String(r[CIDX.status] ?? "").trim();
    const cancelled = /cancel/i.test(status);
    // hours from the actual time span (Care's parenthetical rounds up, e.g. 4.75 -> "4.8")
    const spanMin = (()=>{ const [h1,m1]=t.start.split(":").map(Number), [h2,m2]=t.end.split(":").map(Number);
      return ((h2*60+m2)-(h1*60+m1)+1440)%1440; })();
    t.hrs = spanMin/60;

    // client match — exact on Care.com's job number first, then the old heuristics
    let client = null, matchedBy = "";
    const jr = byJob.get(jid);
    if (jr && jr.client) { client = jr.client; matchedBy = "job#"; }
    if (!client) {
      const cand = byDT.get(date+"|"+t.start) || [];
      const cf = firstName(cg);
      for (const c of cand) { const bf = firstName(c.cg); if (bf===cf || (cf.length>=4 && bf.slice(0,4)===cf.slice(0,4))) { client=c.client; matchedBy="date+time+caregiver"; break; } }
      if (!client && cand.length===1) { client = cand[0].client; matchedBy="date+time"; }
    }
    if (!client) {
      const cd = (byCgDate.get(lastName(cg).toLowerCase()+"|"+date) || []);
      const corpCd = cd.filter(x=>x.svc==="Corporate (Invoiced)" && x.status!=="cancelled");
      if (corpCd.length===1) { client = corpCd[0].client; matchedBy="caregiver+date"; }
      else if (cd.length===1) { client = cd[0].client; matchedBy="caregiver+date"; }
    }

    const row = { jid, date, cg, client: client || "Care.com Family", start: t.start, end: t.end, hrs: t.hrs,
                  bonus, miles: 0, state: stateOf(r[CIDX.loc]), multi: String(r[CIDX.multi] ?? "").trim(),
                  status, matchedBy, bkRow: jr || null, oddId: !/^\d{7}$/.test(jid) };

    if (cancelled) {
      cancIds.add(jid);
      row.stamp = parseStamp(r[CIDX.canc]);
      careCanc.push(row);
    } else {
      if (status && !/^(done|completed|complete|active)$/i.test(status)) S.oddStatus.push(`${jid} — "${status}"`);
      doneIds.add(jid);
      if (!client) unmatched.push(row);
      jobs.push(row);
    }
  }
  if (!jobs.length) throw new Error(
    careCanc.length
      ? `the Care.com export has ${careCanc.length} cancelled job(s) but no completed ones — re-pull it with BOTH cancelled and active/completed jobs included`
      : "no completed jobs found in the Care.com export — is this the right file?");

  // month/year from most common job date
  const mc = {}; jobs.forEach(j=>{ const k=j.date.slice(0,7); mc[k]=(mc[k]||0)+1; });
  const ym = Object.entries(mc).sort((a,b)=>b[1]-a[1])[0][0];
  S.year = parseInt(ym.slice(0,4)); S.month = parseInt(ym.slice(5,7));

  // mileage: the caregiver mileage-request form export is the source of truth.
  const mileageCand = [];
  const mileageOffInvoice = [];
  S.milSkipped = [];
  const milEntries = milWB ? parseMileageForm(milWB) : null;
  if (milEntries) {
    S.milSource = "form";
    S.milCol = milEntries.milesCol || ""; S.milBlank = milEntries.blank || 0;
    const byJid = new Map(jobs.map(j=>[j.jid, j]));
    const bkMap = new Map(bRows.map(r=>[String(r.bk), r]));
    const seenJid = new Set();
    for (const e of milEntries) {
      let j = byJid.get(e.job), how = "";
      if (!j) {
        const bkr = bkMap.get(e.job);
        if (bkr) {
          if (bkr.svc !== "Corporate (Invoiced)") { S.milSkipped.push(`${e.name} — #${e.job} is a ${bkr.svc} booking (Sitterwise-side reimbursement, not a Care.com invoice item)`); continue; }
          j = jobs.find(x => x.date===bkr.date && (firstName(x.cg)===firstName(bkr.cg) || firstName(x.cg).slice(0,4)===firstName(bkr.cg).slice(0,4)));
          how = ` (matched from Sitterwise booking #${e.job})`;
          if (!j) { S.milSkipped.push(`${e.name} — booking #${e.job} (${bkr.client||""} ${bkr.date||""}) has no completed Care.com job this month — cancelled or not on this invoice`); continue; }
        } else { S.milSkipped.push(`${e.name} — job ${e.job} isn't on this month's Care.com export (likely next month's invoice, or already billed)`); continue; }
      }
      if (seenJid.has(j.jid)) continue; // duplicate submission for the same job
      seenJid.add(j.jid);
      // Who filed the claim isn't always who worked the job. A caregiver cancelled off
      // a job can still file mileage against its number — and then file again against
      // the job she did work, so the same drive arrives twice under two job numbers.
      const sameCg = e.name && j.cg && firstName(e.name)===firstName(j.cg) && lastName(e.name)===lastName(j.cg);
      const mismatch = !!(e.name && !sameCg);
      mileageCand.push({ jid: j.jid, cg: j.cg, submitter: e.name || "", mismatch,
        client: j.client, date: j.date, reimb: e.amt, miles: e.over, subMiles: e.over, subAmt: e.amt,
        include: !(j.bonus > 0) && !mismatch, conflict: j.bonus > 0, src: "form"+how,
        note: mismatch ? `Filed by ${e.name}, but the Care.com export has ${j.cg} on job ${j.jid}. Confirm who actually drove before billing this.` : "" });
    }
  } else {
    // The bookings export now carries mileage approval itself. Use it when it is
    // populated; otherwise fall back to estimating from the Reimbursement column.
    const jobFor = r => (r.jobs.map(j=>jobs.find(x=>x.jid===j)).find(Boolean))
      || jobs.find(j => j.date===r.date && r.cg && (firstName(j.cg)===firstName(r.cg) || firstName(j.cg).slice(0,4)===firstName(r.cg).slice(0,4)));
    const approved = corp.filter(r => /appro/.test(r.milStatus) && (r.payMiles>0 || r.apprMiles>0 || r.rtMiles>40));
    if (approved.length) {
      S.milSource = "book";
      for (const r of approved) {
        const j = jobFor(r);
        // prefer explicitly payable/approved miles; else derive the over-40 portion
        const miles = Math.round(r.payMiles || r.apprMiles || Math.max(r.rtMiles-40, 0));
        if (!j) { mileageOffInvoice.push(r); continue; }
        mileageCand.push({ jid: j.jid, cg: j.cg, client: r.client||j.client, date: r.date,
          reimb: r.milAmt || r.reimb, miles, subMiles: miles, subAmt: r.milAmt || r.reimb,
          include: !(j.bonus > 0), conflict: j.bonus > 0,
          src: "approved in the bookings export", note: r.note || "" });
      }
    } else {
      S.milSource = "reimb";
      // Not every reimbursement is mileage — "$5 X 8hrs for an extra child" is not.
      const looksMileage = r => /mile|mileage|\bmi\b|gas|drive|driving/i.test(r.reimbDesc + " " + r.note) || !r.reimbDesc;
      for (const r of corp) {
        if (!(r.reimb > 0)) continue;
        if (!looksMileage(r)) { S.milSkipped.push(`${r.client} ${r.date ? r.date.slice(5) : ""} — $${r.reimb.toFixed(2)} reimbursement is not mileage (“${r.reimbDesc}”), left off the invoice`); continue; }
        const match = jobFor(r);
        const desc = r.reimbDesc ? `reimbursement “${r.reimbDesc}” ($${r.reimb.toFixed(2)})` : `$${r.reimb.toFixed(2)} reimbursement`;
        const mixed = /\+|\/|extra|additional|kid|child/i.test(r.reimbDesc) && /mile/i.test(r.reimbDesc);
        const est = Math.round(r.reimb/0.76);
        if (match) mileageCand.push({ jid: match.jid, cg: match.cg, client: r.client, date: r.date, reimb: r.reimb,
          miles: est, subMiles: est, subAmt: r.reimb, include: !(match.bonus > 0), conflict: match.bonus > 0,
          src: "estimated from reimbursement — upload the mileage form export for exact figures",
          mixed, note: [mixed ? `⚠ ${desc} covers more than mileage — enter only the mileage portion` : desc, r.note].filter(Boolean).join(" · ") });
        else mileageOffInvoice.push(r);
      }
    }
  }

  // tips on corporate jobs → warn (OnPay, never the invoice)
  const tips = corp.filter(r => r.tip > 0);
  // Sitterwise lifesaver bonuses are a caregiver incentive we pay, not a Care.com
  // billable. Care.com bonuses come from the Care export's own Bonus column.
  S.lifesavers = corp.filter(r => r.lifesaver > 0).map(r => ({client:r.client, date:r.date, cg:r.cg, amt:r.lifesaver}));

  // admin notes across all corporate rows (guidance only — never written to the invoice)
  S.adminNotes = corp.filter(r => r.note).map(r => ({ client:r.client, date:r.date, cg:r.cg, status:r.status,
    note:r.note, jobs:r.jobs, bk:String(r.bk), hasJobNo:!!r.jobRaw }));

  // ---- cancellations ----
  // Two sources now. The Care export's own cancelled rows are authoritative: they
  // carry Care.com's job ID and a cancellation TIMESTAMP, so the >24hr/<24hr call
  // and the reassignment-shadow test stop being guesses. The bookings file still
  // contributes anything Care.com didn't list, flagged for confirmation.
  const excluded = []; let billable = [];
  const noticeHrs = c => (c.stamp && c.date) ? (atTime(c.date, c.start) - c.stamp) / 3600000 : null;

  // (a) from the Care export
  for (const c of careCanc) {
    if (doneIds.has(c.jid)) {  // same Care job ID also came back Done → reassigned and worked
      excluded.push({why:"shadow", label:`${c.client} ${c.date.slice(5)} — job ${c.jid} was reassigned to another caregiver and worked (already billed as completed)`});
      continue;
    }
    const bk = c.bkRow;
    billable.push({ src:"care", bk:c.jid, careId:c.jid, client:c.client, date:c.date, start:c.start,
      hrs:c.hrs||0, cg:c.cg, stamp:c.stamp, multi:c.multi, state:c.state,
      note:(bk && bk.note) || "" });
  }

  // (b) from the bookings file — only what the Care export didn't already cover
  const paidKeys = new Set(corp.filter(r=>["paid","completed","confirmed"].includes(r.status)).map(r=>r.client+"|"+r.date+"|"+r.start));
  for (const r of corp.filter(r=>r.status==="cancelled")) {
    if (r.jobs.some(j=>cancIds.has(j))) continue;                        // already on the Care list
    if (r.jobs.some(j=>doneIds.has(j))) {
      excluded.push({why:"shadow", label:`${r.client} ${r.date ? r.date.slice(5) : ""} — job ${r.jobs.find(j=>doneIds.has(j))} was reassigned and worked (already billed as completed)`});
      continue;
    }
    if (paidKeys.has(r.client+"|"+r.date+"|"+r.start)) { excluded.push({why:"shadow", label:`${r.client} ${r.date.slice(5)} — job was reassigned and worked (already billed as completed)`}); continue; }
    if (!r.cg) { excluded.push({why:"nofill", label:`${r.client} ${r.date ? r.date.slice(5) : ""} — no caregiver assigned (agency couldn't fill; no fee)`}); continue; }
    // Cancelled in Sitterwise but absent from the Care.com export. The only ID we have
    // for it is a Sitterwise booking number, which will never match a Care.com job —
    // Care.com can't pay a line it can't identify, so it does not go on the invoice.
    excluded.push({why:"notincare", label:`${r.client} ${r.date ? r.date.slice(5) : ""} (${r.cg}) — cancelled in Sitterwise but not in the Care.com export${r.jobRaw ? `; booking #${r.bk} lists job ${r.jobRaw}` : `; booking #${r.bk} has no Care.com job number`} — not billable`});
  }

  // double-staffed: same client+date+start with 2+ billable rows → keep one, note it
  const seen = new Map();
  billable = billable.filter(r => {
    const k = r.client+"|"+r.date+"|"+r.start;
    if (seen.has(k)) { excluded.push({why:"double", label:`${r.client} ${r.date.slice(5)} — second caregiver on the same cancelled job (one fee only)`}); return false; }
    seen.set(k, true); return true;
  });
  billable.sort((a,b)=> lastName(a.client).toLowerCase().localeCompare(lastName(b.client).toLowerCase()) || a.date.localeCompare(b.date));

  // multi-day groups: same client, consecutive dates. Care.com policy 3 — if every
  // day was cancelled at once (within 15 minutes), it is ONE fee. With cancellation
  // timestamps we can now prove that, so the extra days default to "no charge".
  let gid = 0;
  for (let i=1;i<billable.length;i++){
    const cur = billable[i], prev = billable[i-1];
    if (prev.client !== cur.client) continue;
    if ((new Date(cur.date) - new Date(prev.date))/86400000 !== 1) continue;
    cur._g = prev._g ?? (prev._g = ++gid);
    if (prev.stamp && cur.stamp && Math.abs(cur.stamp - prev.stamp) <= 15*60000) cur._simul = true;
  }
  // a day is auto-zeroed only if it is simultaneous with the day before it AND the
  // whole chain back to the group's first day is simultaneous
  for (let i=1;i<billable.length;i++){
    const cur = billable[i], prev = billable[i-1];
    if (cur._g && cur._g === prev._g) cur._chain = cur._simul && (prev._chain !== false);
    else if (cur._g) cur._chain = cur._simul;
  }

  S.cancBillable = billable.map(r => {
    const nh = noticeHrs(r);
    const grouped = !!r._g && r._chain === true;   // a follow-on day of a simultaneous multi-day cancel
    return { bk:String(r.bk), careId:r.careId, src:r.src, client:r.client, date:r.date, start:r.start,
             hrs:r.hrs||0, cg:r.cg, group:r._g||null, grouped, notice:nh, state:r.state||"CA",
             mode: grouped ? "none" : (nh !== null && nh < 24 ? "<24" : ">24"),
             autoMode: true, note:r.note||"" };
  });

  // double-bill check: a billable cancellation whose caregiver ALSO has a completed
  // Care.com job the same day may be the same shift on the invoice twice. Cancellations
  // that came from the Care export carry their own distinct job ID, so they are already
  // proven to be a different job — only the bookings-only ones need this warning.
  S.dblBill = [];
  for (const c of S.cancBillable) {
    if (c.src === "care") continue;
    const same = (a,b) => firstName(a)===firstName(b) && lastName(a)===lastName(b);
    const hit = jobs.find(j => j.date===c.date && c.cg && same(j.cg, c.cg));
    if (hit) { c.dbl = hit.jid; if (!S.dblBill.some(d=>d.jid===hit.jid)) S.dblBill.push({jid:hit.jid, cg:hit.cg, client:c.client, date:c.date}); }
  }
  S.cancExcluded = excluded;
  S.jobs = jobs; S.unmatched = unmatched; S.mileageCand = mileageCand;

  // weekly OT check (warn only — historically never triggers)
  const wk = {};
  jobs.forEach(j => { const k=j.cg+"|"+isoWeek(j.date); wk[k]=(wk[k]||0)+Math.min(j.hrs,8); });
  S.weeklyOTWarn = Object.entries(wk).filter(([,v])=>v>40).map(([k,v])=>k.split("|")[0]+" ("+v.toFixed(1)+"h)");

  // ---- step 4: every invoiced job that isn't Care.com ----
  S.clientInvoices = []; S.clientSkipped = [];
  if (xeroTxt) {
    const isCare = r => r.jobs.length > 0 || /@care\.com$/i.test(r.email);
    const nonCare = bRows.filter(r => (r.svc === "Corporate (Invoiced)" || r.svc === "Group Childcare (Invoiced)")
                                      && !isCare(r) && r.date && r.date.slice(0,7) === ym);
    const learned = learnClients(parseCSV(xeroTxt));
    const built = buildClientInvoices(nonCare, learned, S.year, S.month); // month is 1-based -> the 1st of the NEXT month
    S.clientInvoices = built.invoices; S.clientSkipped = built.skipped;
  }

  renderQuestions({tips, mileageOffInvoice});
  renderClients();
}

/* ---------------- questions UI ---------------- */
function renderQuestions(extra) {
  setStep(2);
  $("card-questions").classList.remove("locked");
  const mn = MONTHS[S.month-1];
  $("tape-month").textContent = mn + " " + S.year;
  $("analysis-sub").textContent = `Read both files for ${mn} ${S.year}. Here's what the rules found — confirm the calls below.`;

  const otJobs = S.jobs.filter(j=>j.hrs>8).length;
  const minJobs = S.jobs.filter(j=>j.hrs<4).length;
  const bonusN = S.jobs.filter(j=>j.bonus>0).length;
  const lastMin = S.cancBillable.filter(c=>c.mode==="<24").length;
  const autoGrouped = S.cancBillable.filter(c=>c.grouped).length;
  const bookOnly = S.cancBillable.filter(c=>c.src==="book").length;
  const byJobN = S.jobs.filter(j=>j.matchedBy==="job#").length;
  const chips = [
    `<span class="chip"><b>${S.jobs.length}</b> completed jobs</span>`,
    `<span class="chip"><b>${otJobs}</b> CA overtime splits</span>`,
    minJobs ? `<span class="chip"><b>${minJobs}</b> under 4-hr minimum</span>` : "",
    `<span class="chip"><b>${bonusN}</b> bonuses</span>`,
    `<span class="chip"><b>${S.cancBillable.length}</b> billable cancellations</span>`,
    `<span class="chip good"><b>${S.cancExcluded.length}</b> excluded (shadows / unfilled)</span>`,
    S.weeklyOTWarn.length ? `<span class="chip warn">⚠ weekly OT: ${S.weeklyOTWarn.join(", ")}</span>` : `<span class="chip good">weekly OT: none</span>`,
    extra.tips.length ? `<span class="chip warn">⚠ ${extra.tips.length} tip(s) on corporate jobs — left OFF the invoice (route via OnPay)</span>` : "",
    extra.mileageOffInvoice.length ? `<span class="chip warn">${extra.mileageOffInvoice.length} mileage reimbursement(s) on jobs NOT in the Care export — excluded</span>` : "",
    S.jobs.some(j=>j.oddId) ? `<span class="chip warn">⚠ non-standard job ID: ${S.jobs.filter(j=>j.oddId).map(j=>j.jid).join(", ")} — confirm it's legit</span>` : "",
    S.milSource==="form" ? `<span class="chip good">🚗 mileage form: ${S.mileageCand.length} matched, ${S.milSkipped.length} skipped — miles read from “${esc(S.milCol)}”</span>` : `<span class="chip warn">no mileage form uploaded \u2014 mileage below is estimated from reimbursements</span>`,
    (S.milSource==="form" && S.milBlank) ? `<span class="chip warn">⚠ ${S.milBlank} mileage submission(s) had no billable miles in “${esc(S.milCol)}” — they will bill $0 unless you type the miles in below</span>` : "",
    S.mileageCand.some(m=>m.include && !m.miles) ? `<span class="chip warn">⚠ ${S.mileageCand.filter(m=>m.include && !m.miles).length} mileage row(s) are ticked but set to 0 miles — they add nothing to the invoice</span>` : "",
    S.mileageCand.some(m=>m.mismatch) ? `<span class="chip warn">⚠ ${S.mileageCand.filter(m=>m.mismatch).length} mileage claim(s) filed by someone other than the caregiver on the job — unticked, confirm who drove</span>` : "",
    S.dblBill.length ? `<span class="chip warn">⚠ ${S.dblBill.length} possible double-bill(s): a billed cancellation's caregiver also has a completed job that day — see below</span>` : "",
    S.adminNotes.length ? `<span class="chip note">📝 ${S.adminNotes.length} admin note(s) in the bookings file — shown next to the calls below</span>` : "",
    S.careHasStatus
      ? `<span class="chip good">✓ Care export includes cancelled + completed jobs</span>`
      : `<span class="chip warn">⚠ Care export has no Status column — it looks like completed jobs only. Re-pull it with cancelled jobs included so cancellations get Care.com job IDs and notice times.</span>`,
    lastMin ? `<span class="chip warn">⏱ ${lastMin} cancellation(s) under 24 hr notice — set from the Care export's cancellation timestamps</span>` : "",
    autoGrouped ? `<span class="chip good">${autoGrouped} multi-day cancellation day(s) auto-zeroed (cancelled together → one fee)</span>` : "",
    bookOnly ? `<span class="chip warn">⚠ ${bookOnly} cancellation(s) are in the bookings file but NOT in the Care export — confirm the job ID before sending</span>` : "",
    `<span class="chip good">${byJobN}/${S.jobs.length} clients matched on Care.com job number</span>`,
    S.oddStatus.length ? `<span class="chip warn">⚠ unexpected Care status: ${S.oddStatus.join(", ")} — billed as worked; confirm</span>` : "",
    S.lifesavers.length ? `<span class="chip warn">⚠ ${S.lifesavers.length} lifesaver bonus(es) (${money(S.lifesavers.reduce((a,b)=>a+b.amt,0))}) — Sitterwise-side incentive, left OFF the invoice</span>` : "",
  ].filter(Boolean).join("");
  $("chips").innerHTML = chips;

  $("q-settings").style.display = "block";
  $("set-invnum").value = `SW-CARE-${S.year}-${String(S.month).padStart(2,"0")}`;

  // unmatched clients
  if (S.unmatched.length) {
    $("q-clients").style.display = "block";
    $("tbl-clients").innerHTML = `<tr><th>Job</th><th>Date</th><th>Caregiver</th><th>Client name</th></tr>` +
      S.unmatched.map((j,i)=>{
        const n = S.adminNotes.find(a=>a.date===j.date && a.cg && j.cg && a.cg.split(" ")[0].toLowerCase()===j.cg.split(" ")[0].toLowerCase());
        return `<tr><td class="mono">${j.jid}</td><td class="mono">${j.date.slice(5)}</td><td>${j.cg}</td>
        <td><input type="text" data-uidx="${i}" class="inp-client" value="" placeholder="Care.com Family" style="width:200px"></td></tr>${n?`<tr><td colspan="4" class="adminnote">📝 ${esc(n.note)}</td></tr>`:""}`;
      }).join("");
  } else $("q-clients").style.display = "none";

  // mileage
  if (S.mileageCand.length || S.milSkipped.length) {
    $("q-mileage").style.display = "block";
    // Saudia approves a DOLLAR AMOUNT, not a mileage count, so the amount is a
    // first-class editable field alongside the miles the invoice actually bills.
    $("tbl-mileage").innerHTML = `<tr><th>Approve</th><th>Job</th><th>Caregiver</th><th>Date</th><th>Submitted</th><th>Miles over 40 RT</th><th>Amount to bill</th><th>Status</th></tr>` +
      S.mileageCand.map((m,i)=>{
        const st = m.mismatch
          ? `<span class="pill less" title="The mileage form was filed by ${esc(m.submitter)}, but the Care.com export shows ${esc(m.cg)} worked job ${esc(m.jid)}. Unticked until you confirm who drove.">\u26a0 filed by ${esc(m.submitter)}, not ${esc(m.cg)}</span>`
          : m.conflict
          ? `<span class="pill less" title="Care.com does not pay mileage AND a bonus on the same job unless both are pre-approved.">\u26a0 $50 bonus on this job \u2014 pick one</span>`
          : (S.milSource==="form" ? `<span class="pill multi">\u2713 form submission</span>` : `<span class="pill skip">estimate \u2014 verify</span>`);
        const sub = m.subMiles ? `${+(+m.subMiles).toFixed(2)} mi · ${money(m.subAmt||0)}` : "\u2014";
        return `<tr class="salrow" data-mrow="${i}">
        <td><span class="seg" role="group">
             <button type="button" class="sg ok${m.include?" on":""}" data-app="${i}" data-v="1">Approved</button>
             <button type="button" class="sg no${m.include?"":" on"}" data-app="${i}" data-v="0">Skip</button>
           </span></td>
        <td class="mono">${m.jid}</td><td>${m.cg}</td><td class="mono">${m.date.slice(5)}</td>
        <td class="mono submitted">${sub}</td>
        <td><input type="number" data-miles="${i}" class="inp-miles" value="${+(+m.miles).toFixed(2)}" min="0" step="0.01" style="width:84px"></td>
        <td><span class="amtwrap">$<input type="number" data-amt="${i}" class="inp-amt" value="${(m.miles*0.76).toFixed(2)}" min="0" step="0.01" style="width:88px"></span></td>
        <td>${st} <span class="adj" data-adj="${i}"></span></td></tr>${m.note?`<tr class="salrow"><td></td><td colspan="7" class="adminnote">📝 ${esc(m.note)}</td></tr>`:""}`;
      }).join("");
    if (S.milSkipped.length) {
      $("mil-skipped").style.display = "block";
      $("mil-skipped-list").innerHTML = S.milSkipped.map(s=>`<li>${esc(s)}</li>`).join("");
    } else $("mil-skipped").style.display = "none";
  } else { $("q-mileage").style.display = "none"; }

  // cancellations
  if (S.cancBillable.length) {
    $("q-canc").style.display = "block";
    $("tbl-canc").innerHTML = `<tr><th>Care.com Job ID</th><th>Client</th><th>Date</th><th>Caregiver</th><th>Booked hrs</th><th>Notice</th><th>How to bill</th></tr>` +
      S.cancBillable.map((c,i)=>{
        const grp = c.group ? `<span class="pill multi" title="Consecutive-day booking for the same client${c.grouped?" — cancelled at the same time as the day before, so Care.com pays one fee for the whole booking":""}">multi-day #${c.group}</span> ` : "";
        const dbl = c.dbl ? `<span class="pill less" title="This caregiver also has completed Care.com job ${c.dbl} on this date — either the cancellation or the completed job should come off">⚠ dbl ${c.dbl}</span> ` : "";
        const srcPill = c.src === "care"
          ? `<span class="pill multi" title="Came from the Care.com export — this is Care.com's own job ID">✓ Care</span> `
          : `<span class="pill less" title="Cancelled in Sitterwise but NOT listed in the Care.com export. The ID below is a Sitterwise booking number — look up Care.com's 7-digit job ID in the Care portal before sending.">⚠ not in Care export</span> `;
        const idWarn = /^\d{7}$/.test(c.bk) ? "" : ` title="This looks like a Sitterwise booking number — Care.com needs THEIR 7-digit job ID (find it in the Care portal)" style="border-color:var(--amber)"`;
        const notice = c.notice === null || c.notice === undefined
          ? `<span class="pill skip" title="No cancellation timestamp — pull the Care export with cancelled jobs included to get one">unknown</span>`
          : `<span class="pill ${c.notice < 24 ? "less" : "multi"}" title="Cancelled ${c.notice.toFixed(1)} hours before the shift started">${c.notice < 24 ? c.notice.toFixed(1)+" hr" : Math.round(c.notice/24)+" d"}</span>`;
        const sel = v => c.mode===v ? " selected" : "";
        return `<tr>
          <td><input type="text" data-cidx="${i}" class="inp-cid mono" value="${esc(c.bk)}" style="width:90px"${idWarn}></td>
          <td>${srcPill}${dbl}${grp}${c.client}</td><td class="mono">${c.date.slice(5)}</td><td>${c.cg}</td>
          <td class="mono">${c.hrs ? c.hrs.toFixed(2) : "—"}</td>
          <td>${notice}</td>
          <td><select data-cidx="${i}" class="inp-cmode">
            <option value=">24"${sel(">24")}>&gt;24 hr — flat fee</option>
            <option value="<24"${sel("<24")}>&lt;24 hr — booked hrs @ rate (cap 8)</option>
            <option value="none"${sel("none")}>No charge</option>
          </select></td></tr>${c.grouped?`<tr><td colspan="7" class="adminnote">↳ cancelled within 15 minutes of the previous day — Care.com policy 3 allows one fee for the whole multi-day booking, so this day is set to no charge.</td></tr>`:""}${c.note?`<tr><td colspan="7" class="adminnote">📝 ${esc(c.note)}</td></tr>`:""}`;
      }).join("");
  } else $("q-canc").style.display = "none";

  const extraExcl = S.lifesavers.map(l=>`${esc(l.client)} ${l.date?l.date.slice(5):""} (${esc(l.cg||"—")}) — ${money(l.amt)} Sitterwise lifesaver bonus, not a Care.com charge`);
  $("excl-list").innerHTML = (S.cancExcluded.map(e=>`<li>${e.label}</li>`).concat(extraExcl.map(t=>`<li>${t}</li>`))).join("") || "<li>Nothing was excluded.</li>";
  $("excl-details").style.display = (S.cancExcluded.length || S.lifesavers.length) ? "block" : "none";
  if (S.adminNotes.length) {
    // Say plainly whether each noted booking actually reached the invoice — the notes
    // are where "invoice this to someone else" gets written down, so "is it on there?"
    // needs to be answerable without cross-referencing anything.
    const onJobs = new Set(S.jobs.filter(j=>!S.excludeJobs.has(j.jid)).map(j=>j.jid));
    const onCanc = new Set(S.cancBillable.filter(c=>c.mode!=="none").map(c=>String(c.bk)));
    const verdict = n => {
      const hit = n.jobs.find(j=>onJobs.has(j));
      if (hit) return {cls:"multi", txt:`on the invoice · job ${hit}`};
      if (n.jobs.some(j=>onCanc.has(j)) || onCanc.has(n.bk)) return {cls:"multi", txt:"on the invoice · cancellation"};
      if (!n.hasJobNo) return {cls:"skip", txt:"not on the invoice · no Care.com job number"};
      return {cls:"skip", txt:"not on the invoice · no matching job in the Care export"};
    };
    $("notes-list").innerHTML = S.adminNotes.map(n=>{
      const v = verdict(n);
      return `<li><span class="pill ${v.cls}">${v.txt}</span> <b>${esc(n.client)}</b> ${n.date.slice(5)} (${esc(n.cg||"—")}, ${esc(n.status)}): ${esc(n.note)}</li>`;
    }).join("");
    $("notes-details").style.display = "block";
  } else $("notes-details").style.display = "none";

  // possible double-bills: offer to exclude the completed job (if it wasn't actually worked)
  if (S.dblBill.length) {
    $("q-exclude").style.display = "block";
    $("tbl-exclude").innerHTML = `<tr><th>Exclude</th><th>Care Job</th><th>Caregiver</th><th>Client</th><th>Date</th></tr>` +
      S.dblBill.map((d,i)=>`<tr class="salrow"><td><input type="checkbox" data-xidx="${i}" class="inp-excl"></td>
        <td class="mono">${d.jid}</td><td>${d.cg}</td><td>${d.client}</td><td class="mono">${d.date.slice(5)}</td></tr>`).join("");
  } else $("q-exclude").style.display = "none";

  $("q-actions").style.display = "flex";
  // Approved / Skip owns `include`; the two number fields are two views of the
  // same figure (the invoice bills miles in column N, O = N x rate), so editing
  // either keeps the other honest.
  document.querySelectorAll(".sg[data-app]").forEach(b=>b.addEventListener("click", ()=>{
    const i = +b.dataset.app, on = b.dataset.v === "1";
    S.mileageCand[i].include = on;
    document.querySelectorAll(`.sg[data-app="${i}"]`).forEach(x=>x.classList.toggle("on", (x.dataset.v==="1")===on));
    const row = document.querySelector(`[data-mrow="${i}"]`); if (row) row.classList.toggle("skipped", !on);
    updateTape();
  }));
  const mrate = () => parseFloat($("set-mile").value) || 0.76;
  const syncAdj = i => {
    const m = S.mileageCand[i], el = document.querySelector(`[data-adj="${i}"]`);
    if (!el) return;
    const d = Math.round((m.miles - (m.subMiles||0))*100)/100;
    // a caregiver can add up to 10 miles at checkout via reimbursements
    el.innerHTML = !d ? "" :
      `<span class="pill ${(d>10||d<0)?"less":"multi"}">${d>0?"+":""}${d} mi vs submitted${d>10?" — over the 10-mile limit":""}</span>`;
  };
  document.querySelectorAll(".inp-miles").forEach(inp=>inp.addEventListener("input", ()=>{
    const i = +inp.dataset.miles, v = Math.max(0, parseFloat(inp.value)||0);
    S.mileageCand[i].miles = v;
    const a = document.querySelector(`[data-amt="${i}"]`); if (a) a.value = (v*mrate()).toFixed(2);
    syncAdj(i); updateTape();
  }));
  document.querySelectorAll(".inp-amt").forEach(inp=>inp.addEventListener("input", ()=>{
    const i = +inp.dataset.amt, amt = Math.max(0, parseFloat(inp.value)||0);
    const v = Math.round((amt/mrate())*100)/100;
    S.mileageCand[i].miles = v;
    const mi = document.querySelector(`[data-miles="${i}"]`); if (mi) mi.value = v;
    syncAdj(i); updateTape();
  }));
  $("set-mile").addEventListener("input", ()=>{
    S.mileageCand.forEach((m,i)=>{ const a=document.querySelector(`[data-amt="${i}"]`); if (a) a.value=(m.miles*mrate()).toFixed(2); });
  });
  S.mileageCand.forEach((m,i)=>{ syncAdj(i); const r=document.querySelector(`[data-mrow="${i}"]`); if (r) r.classList.toggle("skipped", !m.include); });
  document.querySelectorAll("#card-questions input, #card-questions select").forEach(el=>el.addEventListener("input", updateTape));
  updateTape();
  $("card-questions").scrollIntoView({behavior:"smooth", block:"start"});
}

function collectAnswers() {
  const rate = parseFloat($("set-rate").value)||42;
  const mileRate = parseFloat($("set-mile").value)||0.76;
  const cancFee = parseFloat($("set-cancfee").value)||30;
  document.querySelectorAll(".inp-client").forEach(inp=>{
    const j = S.unmatched[+inp.dataset.uidx];
    j.client = cleanName(inp.value) || "Care.com Family";
  });
  document.querySelectorAll(".inp-miles").forEach(inp=>{ S.mileageCand[+inp.dataset.miles].miles = Math.max(0, parseFloat(inp.value)||0); });
  S.jobs.forEach(j=>j.miles=0);
  S.mileageCand.forEach(m=>{ if (m.include) { const j=S.jobs.find(x=>x.jid===m.jid); if (j) j.miles = m.miles; } });
  document.querySelectorAll(".inp-cmode").forEach(sel=>{ S.cancBillable[+sel.dataset.cidx].mode = sel.value; });
  document.querySelectorAll(".inp-cid").forEach(inp=>{ const v=inp.value.trim(); if (v) S.cancBillable[+inp.dataset.cidx].bk = v; });
  S.excludeJobs = new Set();
  document.querySelectorAll(".inp-excl").forEach(cb=>{ if (cb.checked) S.excludeJobs.add(S.dblBill[+cb.dataset.xidx].jid); });
  return {rate, otRate: Math.round(rate*1.5*100)/100, mileRate, cancFee, invNum: $("set-invnum").value.trim()};
}

/* ---------------- live tape ---------------- */
function computeTotals(a) {
  let regH=0, otH=0, bonus=0, miles=0;
  for (const j of S.jobs) {
    if (S.excludeJobs.has(j.jid)) continue;
    const h = Math.max(j.hrs, 4); // 4-hr minimum
    regH += Math.min(h,8); otH += Math.max(h-8,0);
    bonus += j.bonus; miles += j.miles;
  }
  let canc = 0;
  for (const c of S.cancBillable) {
    if (c.mode===">24") canc += a.cancFee;
    else if (c.mode==="<24") canc += Math.min(c.hrs||8, 8) * a.rate;
  }
  const mileCost = Math.round(miles*a.mileRate*100)/100;
  return {regH, otH, bonus, mileCost, canc,
    total: Math.round((regH*a.rate + otH*a.otRate + bonus + mileCost + canc)*100)/100};
}
function updateTape() {
  if (!S.jobs.length) return;
  const a = collectAnswers(); const t = computeTotals(a);
  const set=(id,val,on)=>{ const el=$(id); el.querySelector(".amt").textContent=val; el.classList.toggle("subtle",!on); };
  set("tr-reg", `${t.regH.toFixed(2)}h · ${money(t.regH*a.rate)}`, t.regH>0);
  set("tr-ot", `${t.otH.toFixed(2)}h · ${money(t.otH*a.otRate)}`, t.otH>0);
  set("tr-bonus", money(t.bonus), t.bonus>0);
  set("tr-miles", money(t.mileCost), t.mileCost>0);
  set("tr-canc", money(t.canc), t.canc>0);
  $("tape-total").textContent = money(t.total);
}

/* ---------------- build the workbook ---------------- */
$("btn-build").addEventListener("click", async () => {
  const err=$("err-build"); err.style.display="none";
  try {
    const a = collectAnswers();
    const {blob, total} = await buildWorkbook(a);
    const fname = `Sitterwise_CareAtWork_Invoice_${MONTHS[S.month-1]}${S.year}.xlsx`;
    S.built = {blob, fname, total};
    setStep(3);
    $("card-download").classList.remove("locked");
    $("dl-zone").style.display="block";
    $("dl-total").textContent = money(total);
    $("dl-fname").textContent = fname;
    $("card-download").scrollIntoView({behavior:"smooth"});
  } catch(e) {
    console.error(e);
    err.textContent = "Something went wrong building the invoice: " + e.message;
    err.style.display="block";
  }
});
$("btn-download").addEventListener("click", ()=> {
  if (!S.built) return;
  const url = URL.createObjectURL(S.built.blob);
  const aEl = document.createElement("a"); aEl.href=url; aEl.download=S.built.fname; aEl.click();
  URL.revokeObjectURL(url);
});

async function buildWorkbook(a) {
  const wb = new ExcelJS.Workbook();
  wb.calcProperties.fullCalcOnLoad = true;
  const ws = wb.addWorksheet("Sheet1");

  const SALMON = "FFFFCCCC";
  const thin = {style:"thin", color:{argb:"FF000000"}};
  const BORD = {top:thin,left:thin,bottom:thin,right:thin};
  const HFONT = {name:"Calibri", size:11, bold:true};
  const DFONT = {name:"Aptos Narrow", size:12};
  const MONEYFMT = '"$"#,##0.00';

  // column widths (template)
  const widths = [11.1,13.8,13.1,4.6,9.6,11.6,10.6,7.6,11.3,11.0,7.6,8.3,7.1,11.8,7.8,8.6,42];
  widths.forEach((w,i)=> ws.getColumn(i+1).width = w);

  // rows 1-2 meta + red notes in I1:I3
  ws.getCell("A1").value="Agency Name"; ws.getCell("A1").font={...HFONT};
  ws.getCell("B1").value="Invoice #";  ws.getCell("B1").font={...HFONT};
  ws.getCell("A2").value="Sitterwise, Inc."; ws.getCell("A2").font=DFONT;
  ws.getCell("B2").value=a.invNum; ws.getCell("B2").font=DFONT;
  const notes=[
    "* Please provide a note in the notes column for any time changes that are greater than one hour.",
    "* Mileage must be pre-approved by your Relationship Manager and CANNOT be charged in addition to a bonus. ",
    `* Mileage applies to mileage over 40 miles round trip at $${a.mileRate}/mile.`];
  notes.forEach((t,i)=>{ const c=ws.getCell(1+i,9); c.value=t; c.font={name:"Calibri",size:9,color:{argb:"FFCC0000"}}; });

  // header row 5
  const HEADERS=["Booking ID","Caregiver Name","Client Name","State","Start Date","Actual Start Time","Actual End Time","Total Hours","Requested Start Time","Requested End Time","Total Hours","Hourly Rate","Bonus","# of miles over 40 round trip","Mileage Cost","Total Cost","Notes"];
  HEADERS.forEach((h,i)=>{ const c=ws.getCell(5,i+1); c.value=h; c.font={...HFONT}; c.alignment={wrapText:true, vertical:"bottom"}; c.border=BORD; });
  ws.getRow(5).height = 33;

  // completed rows
  const SALMON_COLS=[6,7,9,10,13,14,15];
  const jobsSorted=[...S.jobs].filter(j=>!S.excludeJobs.has(j.jid)).sort((x,y)=> lastName(x.cg).toLowerCase().localeCompare(lastName(y.cg).toLowerCase()) || x.cg.localeCompare(y.cg) || x.date.localeCompare(y.date) || x.start.localeCompare(y.start));
  let r=6;
  const styleDataRow = row => { for (let c=1;c<=17;c++){ const cell=ws.getCell(row,c); cell.font=DFONT; cell.border=BORD; if (SALMON_COLS.includes(c)) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:SALMON}}; } };
  const dateSerial = iso => { const [Y,M,D]=iso.split("-").map(Number); return Math.round((Date.UTC(Y,M-1,D)-Date.UTC(1899,11,30))/86400000); };
  const setCommon=(row,j)=>{ ws.getCell(row,1).value=/^\d+$/.test(j.jid)?Number(j.jid):j.jid; ws.getCell(row,2).value=j.cg; ws.getCell(row,3).value=j.client; ws.getCell(row,4).value=j.state||"CA";
    const dc=ws.getCell(row,5); dc.value=dateSerial(j.date); dc.numFmt="mm-dd-yy"; };
  const setTimes=(row,s,e)=>{ [[6,s],[7,e],[9,s],[10,e]].forEach(([c,t])=>{ const cell=ws.getCell(row,c); cell.value=timeFrac(t); cell.numFmt="h:mm AM/PM"; }); };
  const setFormulas=(row)=>{ ws.getCell(row,8).value={formula:`(MOD(G${row}-F${row},1)*24)`};
    ws.getCell(row,11).value={formula:`(MOD(J${row}-I${row},1)*24)`};
    ws.getCell(row,15).value={formula:`SUM(N${row}*${a.mileRate})`};
    ws.getCell(row,16).value={formula:`SUM(K${row}*L${row})+M${row}+O${row}`};
    ws.getCell(row,12).numFmt=MONEYFMT; ws.getCell(row,15).numFmt=MONEYFMT; ws.getCell(row,16).numFmt=MONEYFMT; };

  for (const j of jobsSorted) {
    let bh=j.hrs, rs=j.start, re=j.end, note="";
    if (bh<4){ re=addMin(rs,240); note=`4-hr minimum applied (booked ${bh} hrs).`; bh=4; }
    if (j.miles>0) note=(note?note+" ":"")+`Mileage ${j.miles} mi round trip over 40-mi threshold @ $${a.mileRate}.`+(j.bonus>0?" Bonus + mileage both pre-approved.":"");
    if (bh<=8.000001) {
      styleDataRow(r); setCommon(r,j); setTimes(r,rs,re);
      ws.getCell(r,12).value=a.rate; if (j.bonus) ws.getCell(r,13).value=j.bonus; if (j.miles) ws.getCell(r,14).value=j.miles;
      if (note) ws.getCell(r,17).value=note;
      setFormulas(r); r++;
    } else {
      // Care.com OT format: row 1 shows the FULL shift in the Actual columns and the
      // first 8 hrs in the Requested columns; row 2 has no repeated identity, just
      // "CA OT" and the overtime span in the Requested columns.
      const regEnd=addMin(rs,480);
      styleDataRow(r); setCommon(r,j);
      [[6,rs],[7,re],[9,rs],[10,regEnd]].forEach(([c,t])=>{ const cell=ws.getCell(r,c); cell.value=timeFrac(t); cell.numFmt="h:mm AM/PM"; });
      ws.getCell(r,12).value=a.rate; if (j.bonus) ws.getCell(r,13).value=j.bonus; if (j.miles) ws.getCell(r,14).value=j.miles;
      if (note) ws.getCell(r,17).value=note;
      setFormulas(r); r++;
      styleDataRow(r); // identity columns intentionally left blank on the OT line
      ws.getCell(r,8).value="CA OT";
      [[9,regEnd],[10,re]].forEach(([c,t])=>{ const cell=ws.getCell(r,c); cell.value=timeFrac(t); cell.numFmt="h:mm AM/PM"; });
      ws.getCell(r,11).value={formula:`(MOD(J${r}-I${r},1)*24)`};
      ws.getCell(r,12).value=a.otRate; ws.getCell(r,12).numFmt=MONEYFMT;
      ws.getCell(r,15).value={formula:`SUM(N${r}*${a.mileRate})`}; ws.getCell(r,15).numFmt=MONEYFMT;
      ws.getCell(r,16).value={formula:`SUM(K${r}*L${r})+M${r}+O${r}`}; ws.getCell(r,16).numFmt=MONEYFMT;
      r++;
    }
  }
  const lastCompleted=r-1;

  // cancellation section at template positions 295/296/297
  const CANC_TITLE=lastCompleted+2, CANC_HDR=CANC_TITLE+1, CANC_START=CANC_TITLE+2; // right after the jobs, no blank gap
  const tcell=ws.getCell(CANC_TITLE,1); tcell.value="CANCELLATIONS"; tcell.font={...HFONT};
  tcell.fill={type:"pattern",pattern:"solid",fgColor:{argb:SALMON}};
  const CH=["Booking ID","More than or less than 24 hrs","Client Name","State","Date","","","","","","Total Hrs","Rate","","","","",""];
  CH.forEach((h,i)=>{ const c=ws.getCell(CANC_HDR,i+1); if(h)c.value=h; c.font={...HFONT}; c.alignment={wrapText:true}; c.border=BORD;
    if ([1,2,3,4,5,11,12].includes(i+1)) c.fill={type:"pattern",pattern:"solid",fgColor:{argb:SALMON}}; });

  // billable rows (mode 'none' dropped); multi-day note if group had a no-charge sibling
  const billed=S.cancBillable.filter(c=>c.mode!=="none");
  const dropped=S.cancBillable.filter(c=>c.mode==="none");
  let cr=CANC_START;
  for (const c of billed) {
    for (let col=1;col<=16;col++){ const cell=ws.getCell(cr,col); cell.font=DFONT; cell.border=BORD; }
    ws.getCell(cr,1).value=/^\d+$/.test(c.bk)?Number(c.bk):c.bk;
    ws.getCell(cr,2).value = c.mode===">24" ? "More than 24 hrs" : "Less than 24 hrs";
    ws.getCell(cr,3).value=c.client; ws.getCell(cr,4).value=c.state||"CA";
    const dc=ws.getCell(cr,5); dc.value=dateSerial(c.date); dc.numFmt="mm-dd-yy";
    if (c.mode===">24") { ws.getCell(cr,11).value=1.00; ws.getCell(cr,12).value=a.cancFee; }
    else { const h=Math.min(c.hrs||8,8); ws.getCell(cr,11).value=h; ws.getCell(cr,12).value=a.rate;
      const when = (c.notice !== null && c.notice !== undefined) ? ` Cancelled ${c.notice.toFixed(1)} hrs before the shift.` : "";
      ws.getCell(cr,17).value=`Cancelled <24 hrs; billed ${h} hrs @ $${a.rate} (8-hr cap). Caregiver ${c.cg}.${when}`; ws.getCell(cr,17).font=DFONT; }
    ws.getCell(cr,11).numFmt="0.00"; ws.getCell(cr,12).numFmt=MONEYFMT;
    ws.getCell(cr,16).value={formula:`SUM(K${cr}*L${cr})+M${cr}+O${cr}`}; ws.getCell(cr,16).numFmt=MONEYFMT;
    if (c.group && dropped.some(d=>d.group===c.group)) {
      const cur=ws.getCell(cr,17).value;
      ws.getCell(cr,17).value=(cur?cur+" ":"")+"Part of a multi-day booking; other day(s) not charged per policy."; ws.getCell(cr,17).font=DFONT;
    }
    cr++;
  }

  // total at 360 like the template
  const TOTAL_ROW = cr;
  ws.getCell(TOTAL_ROW,16).value={formula:`SUM(P6:P${cr-1})`}; ws.getCell(TOTAL_ROW,16).numFmt=MONEYFMT; ws.getCell(TOTAL_ROW,16).font={...HFONT};
  ws.getCell(TOTAL_ROW,17).value="Invoice Total"; ws.getCell(TOTAL_ROW,17).font={...HFONT};

  // policy text 361-369
  const POLICY=["Cancellation Policy ",
   "1. If a job is cancelled with more than 24 hours\u2019 notice, agencies are permitted to charge a $30 fee and should list 1.00 under the total hrs column.",
   "2. If a job is cancelled with less than 24 hours notice, agencies are permitted to charge up to 8 hrs at their normal hourly rate.",
   "3. If all days of a multi-day job are cancelled simultaneously, or within 15 minutes of each other, agencies are permitted to    ",
   " charge us either one less than 24 hour cancellation fee or one more than 24 hours cancellation fee, depending on when the cancellation occurred.",
   "4. Bonuses are not paid on cancelled jobs, unless a caregiver arrives at a client's home and is then told that the job has been cancelled. ",
   "5. If a caregiver arrives at a client's home and finds out that a job has been cancelled, agencies are permitted to charge the full shift.",
   "6. If a client cancels a job, due to a caregiver not making their introductory phone call, we would not pay that cancellation fee.",
   "7. Cancellation fees will not be paid if a client cancels due to a caregiver not arriving on time."];
  POLICY.forEach((t,i)=>{ const c=ws.getCell(TOTAL_ROW+2+i,1); c.value=t; c.font= i===0? {...HFONT} : {name:"Calibri",size:10}; });

  const t = computeTotals(a);
  const buf = await wb.xlsx.writeBuffer();
  return { blob: new Blob([buf], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}), total: t.total };
}

/* ================================================================
   STEP 4 — every invoiced client that isn't Care.com
   The Xero sales-invoice export is the config: each client's billing
   contact, rate, payment terms and line wording are read back out of
   their own invoice history. The bookings export is NOT trusted for
   price — its "Charge to Client" carries $42 (the Care.com rate) for
   anything typed "Corporate (Invoiced)", including clients Xero has
   always billed at $36.
   ================================================================ */

const GENERIC_DOMAINS = new Set(["email.com","gmail.com","yahoo.com","hotmail.com","outlook.com","icloud.com","care.com","aol.com","me.com"]);
const NAME_STOP = new Set(["the","at","of","and","a","an","inc","llc","group","groups","campus","church","services","service","youth","family","families","community","center","centre","school","co"]);
const domainOf = e => { const m = String(e||"").toLowerCase().match(/@([^@\s]+)$/); return m ? m[1] : ""; };
const ARTICLES = new Set(["the","at","of","and","a","an","&"]);
const rawTokens = n => String(n||"").toLowerCase().replace(/[^a-z0-9 ]+/g," ").split(/\s+/).filter(w=>w.length>1 && !ARTICLES.has(w));
const nameTokens = n => rawTokens(n).filter(w=>!NAME_STOP.has(w));

/* minimal RFC4180 CSV parse (quoted fields, embedded commas/newlines) */
function parseCSV(text) {
  const rows=[]; let row=[], cur="", q=false;
  text = String(text).replace(/\r\n?/g,"\n");
  for (let i=0;i<text.length;i++){
    const c=text[i];
    if (q) { if (c==='"'){ if (text[i+1]==='"'){cur+='"';i++;} else q=false; } else cur+=c; }
    else if (c==='"') q=true;
    else if (c===",") { row.push(cur); cur=""; }
    else if (c==="\n") { row.push(cur); rows.push(row); row=[]; cur=""; }
    else cur+=c;
  }
  if (cur!=="" || row.length) { row.push(cur); rows.push(row); }
  if (!rows.length) return [];
  const hdr = rows[0].map(h=>h.trim());
  return rows.slice(1).filter(r=>r.some(c=>String(c).trim()!==""))
             .map(r => Object.fromEntries(hdr.map((h,i)=>[h, r[i] ?? ""])));
}

const mdy = s => { const m=String(s||"").match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if(!m) return null;
  const y=m[3].length===2?2000+ +m[3]:+m[3]; return new Date(y, +m[1]-1, +m[2]); };
const fmtMDY = d => `${d.getMonth()+1}/${d.getDate()}/${String(d.getFullYear()).slice(2)}`;
const usd = n => (Math.round(n*100)/100).toFixed(4);

/* Learn one profile per Xero contact from their invoice history. */
function learnClients(xrows) {
  const byInv = new Map();
  for (const r of xrows) {
    if (!r.InvoiceNumber) continue;
    if (!byInv.has(r.InvoiceNumber)) byInv.set(r.InvoiceNumber, []);
    byInv.get(r.InvoiceNumber).push(r);
  }
  const profiles = new Map();
  let maxNum = 0;
  for (const [num, lines] of byInv) {
    const nm = (num.match(/(\d+)\s*$/)||[])[1];
    if (nm) maxNum = Math.max(maxNum, +nm);
    const f = lines[0], contact = (f.ContactName||"").trim();
    if (!contact) continue;
    const date = mdy(f.InvoiceDate), due = mdy(f.DueDate);
    let p = profiles.get(contact);
    if (!p) { p = { contact, email:f.EmailAddress||"", addr:{...f}, invoices:0, lastDate:null,
                    rate:0, terms:14, style:"date", quote:false, acct:"460", tax:f.TaxType||"Tax on Sales",
                    cur:f.Currency||"USD", billed:[], rateVotes:new Map() };
             profiles.set(contact, p); }
    p.invoices++;
    // service rates: positive per-unit amounts on quantity-priced lines
    for (const l of lines) {
      const u = parseFloat(l.UnitAmount)||0, q = parseFloat(l.Quantity)||0;
      if (u > 0 && q > 1) p.rateVotes.set(u, (p.rateVotes.get(u)||0)+1);
    }
    // every M/D mentioned on this invoice — used to spot work already billed
    const marks = new Set();
    for (const l of lines) for (const m of String(l.Description||"").matchAll(/(\d{1,2})\/(\d{1,2})(?!\d)/g)) marks.add(+m[1]+"/"+ +m[2]);
    if (date) p.billed.push({date, marks});
    if (!p.lastDate || (date && date > p.lastDate)) {
      p.lastDate = date; p.email = f.EmailAddress || p.email; p.addr = {...f}; p.acct = f.AccountCode || p.acct;
      if (date && due) p.terms = Math.max(Math.round((due-date)/86400000), 0);
      p.quote = /^QU-/i.test((f.Reference||"").trim());
    }
    const svc = lines.filter(l => (parseFloat(l.Quantity)||0) > 1);
    if (svc.length && (!p.styleDate || (date && date > p.styleDate))) {
      p.styleDate = date;
      const descs = svc.map(l=>String(l.Description||""));
      p.style = descs.some(d=>/^childcare services\s+\d{1,2}\//i.test(d)) ? "date"
              : descs.some(d=>/^childcare services\s*[-:]/i.test(d))      ? "named"
              : descs.some(d=>/^[A-Za-z'.-]+ [A-Z]/.test(d) && !/\d/.test(d)) ? "caregiver" : "date";
      p.sample = descs[0] || "";
    }
  }
  for (const p of profiles.values()) {
    p.rate = [...p.rateVotes.entries()].sort((a,b)=>b[1]-a[1])[0]?.[0] || 36;
  }
  return { profiles, nextNum: maxNum + 1 };
}

/* Sitterwise client name/email -> Xero contact. */
function matchContact(name, email, profiles) {
  const e = String(email||"").trim().toLowerCase();
  for (const p of profiles.values()) if (p.email && p.email.toLowerCase() === e) return {p, how:"email"};
  const dom = domainOf(e);
  if (dom && !GENERIC_DOMAINS.has(dom))
    for (const p of profiles.values()) if (domainOf(p.email) === dom) return {p, how:"domain " + dom};
  const t = nameTokens(name);
  if (t.length) {
    let best = null;
    for (const p of profiles.values()) {
      const u = nameTokens(p.contact);
      if (!u.length) continue;
      const shared = t.filter(w=>u.includes(w));
      if (!shared.length) continue;
      const score = shared.length / Math.min(t.length, u.length);
      if (!best || score > best.score) best = {p, score, shared};
    }
    // 0.67 keeps every real match (they all score 1.0 — the contact's distinctive
    // tokens are fully contained in the client name) and rejects a lone shared
    // first name, e.g. "David Russo" vs "St. David's Episcopal Church" at 0.5.
    if (best && best.score >= 0.67) return {p:best.p, how:"name “" + best.shared.join(" ") + "”"};
  }
  return null;
}

/* the part of the Sitterwise client name the Xero contact doesn't already say
   ("Ymca East County Group" under "YMCA Youth & Family Services" -> "East County Group") */
function subLabel(clientName, contactName) {
  const u = new Set(rawTokens(contactName));
  return String(clientName||"").split(/\s+/)
    .filter(w => { const k = w.toLowerCase().replace(/[^a-z0-9]/g,""); return k && k.length>1 && !u.has(k) && !ARTICLES.has(k); })
    .join(" ").trim();
}

/* Group the non-Care invoiced bookings into one draft invoice per Xero contact. */
function buildClientInvoices(corpRows, learned, year, month) {
  const {profiles, nextNum} = learned;
  const skipped = [], invoices = [];
  const groups = new Map();       // xero contact -> rows
  for (const r of corpRows) {
    const hit = matchContact(r.client, r.email, profiles);
    if (!hit) { skipped.push({cls:"skip", txt:`${r.client} ${r.date ? r.date.slice(5) : ""} — no matching client in the Xero history, so there is no contact, rate or line wording to copy. Invoice this one by hand (or re-export Xero over a longer period).`}); continue; }
    const key = hit.p.contact;
    if (!groups.has(key)) groups.set(key, {p:hit.p, how:hit.how, rows:[]});
    groups.get(key).rows.push(r);
  }
  let num = nextNum;
  // billed on the 1st of the month after the work, matching the existing pattern
  const invDate = new Date(year, month, 1);
  for (const [contact, g] of [...groups.entries()].sort((a,b)=>a[0].localeCompare(b[0]))) {
    const p = g.p;
    const done = g.rows.filter(r=>r.status==="completed");
    const canc = g.rows.filter(r=>r.status==="cancelled");
    const other = g.rows.filter(r=>r.status!=="completed" && r.status!=="cancelled");
    const lines = [];
    if (p.style === "caregiver") {
      const byCg = new Map();
      for (const r of done) {
        const k = r.cg || "Caregiver";
        if (!byCg.has(k)) byCg.set(k, {hrs:0, dates:[]});
        const e = byCg.get(k); e.hrs += r.hrs; e.dates.push(r.date);
      }
      for (const [cg,e] of [...byCg.entries()].sort((a,b)=>a[0].localeCompare(b[0])))
        lines.push({ desc:cg, qty:e.hrs, rate:p.rate, dates:[e.dates.sort()[0]] });
    } else if (p.style === "named") {
      for (const r of done.sort((a,b)=>a.date.localeCompare(b.date)||a.start.localeCompare(b.start))) {
        const d = atTime(r.date, "12:00");
        lines.push({ desc:`Childcare Services - ${r.client} ${fmtMDY(d)} ${fmt12(r.start)}-${fmt12(r.end)}`,
                     qty:r.hrs, rate:p.rate, dates:[r.date] });
      }
    } else {
      const byKey = new Map();
      for (const r of done) {
        const label = subLabel(r.client, contact);
        const k = r.date + "|" + label;
        if (!byKey.has(k)) byKey.set(k, {date:r.date, label, hrs:0, cgs:new Set()});
        const e = byKey.get(k); e.hrs += r.hrs; if (r.cg) e.cgs.add(r.cg);
      }
      for (const e of [...byKey.values()].sort((a,b)=>a.date.localeCompare(b.date)||a.label.localeCompare(b.label))) {
        const n = Math.max(e.cgs.size, 1);
        const d = atTime(e.date, "12:00");
        const each = e.hrs / n;
        lines.push({ desc:`Childcare Services ${d.getMonth()+1}/${d.getDate()}${e.label ? " "+e.label : ""} (${n} caregiver${n>1?"s":""} x ${+each.toFixed(2)} hours)`,
                     qty:e.hrs, rate:p.rate, dates:[e.date] });
      }
    }
    // anything whose service date already appears on one of this client's invoices
    for (const l of lines) {
      const d = atTime(l.dates[0], "12:00"), mark = (d.getMonth()+1)+"/"+d.getDate();
      const prior = p.billed.find(b => b.date >= d && b.marks.has(mark));
      if (prior) { l.already = true; l.priorOn = fmtMDY(prior.date); }
    }
    const total = lines.filter(l=>!l.already).reduce((a,l)=>a+l.qty*l.rate, 0);
    invoices.push({
      contact, p, how:g.how, style:p.style, quote:p.quote,
      num: "INV-" + (num++), invDate, dueDate: new Date(invDate.getTime() + p.terms*86400000),
      lines, canc, other, total,
      include: !p.quote && lines.some(l=>!l.already),
      extra: {desc:"", amt:""},
    });
  }
  return {invoices, skipped};
}

/* Xero sales-invoice import CSV — the same 43 columns their export uses. */
const XCOLS = ["ContactName","EmailAddress","POAddressLine1","POAddressLine2","POAddressLine3","POAddressLine4","POCity","PORegion","POPostalCode","POCountry","SAAddressLine1","SAAddressLine2","SAAddressLine3","SAAddressLine4","SACity","SARegion","SAPostalCode","SACountry","InvoiceNumber","Reference","InvoiceDate","DueDate","PlannedDate","Total","TaxTotal","InvoiceAmountPaid","InvoiceAmountDue","InventoryItemCode","Description","Quantity","UnitAmount","Discount","LineAmount","AccountCode","TaxType","TaxAmount","TrackingName1","TrackingOption1","TrackingName2","TrackingOption2","Currency","Type","Sent","Status"];
const csvCell = v => { const s = String(v ?? ""); return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; };

function buildXeroCSV() {
  const out = [XCOLS.join(",")];
  for (const inv of S.clientInvoices) {
    if (!inv.include) continue;
    const rows = inv.lines.filter(l=>!l.already).map(l=>({desc:l.desc, qty:l.qty, rate:l.rate}));
    const ex = inv.extra;
    if (ex.desc.trim() && parseFloat(ex.amt)) rows.push({desc:ex.desc.trim(), qty:1, rate:parseFloat(ex.amt)});
    if (!rows.length) continue;
    const total = rows.reduce((a,r)=>a + r.qty*r.rate, 0);
    for (const r of rows) {
      const rec = Object.fromEntries(XCOLS.map(c=>[c,""]));
      const a = inv.p.addr || {};
      rec.ContactName = inv.contact; rec.EmailAddress = inv.p.email;
      for (const k of ["POAddressLine1","POCity","PORegion","POPostalCode","POCountry","SAAddressLine1","SACity","SARegion","SAPostalCode","SACountry"]) rec[k] = a[k] || "";
      rec.InvoiceNumber = inv.num;
      rec.InvoiceDate = fmtMDY(inv.invDate); rec.DueDate = fmtMDY(inv.dueDate);
      rec.Total = usd(total); rec.TaxTotal = usd(0); rec.InvoiceAmountPaid = usd(0); rec.InvoiceAmountDue = usd(total);
      rec.Description = r.desc; rec.Quantity = usd(r.qty); rec.UnitAmount = usd(r.rate); rec.LineAmount = usd(r.qty*r.rate);
      rec.AccountCode = inv.p.acct || "460"; rec.TaxType = inv.p.tax || "Tax on Sales"; rec.TaxAmount = usd(0);
      rec.Currency = inv.p.cur || "USD"; rec.Type = "Sales invoice"; rec.Sent = "Unsent"; rec.Status = "Draft";
      out.push(XCOLS.map(c=>csvCell(rec[c])).join(","));
    }
  }
  return out.join("\n") + "\n";
}

/* ---------------- step 4 UI ---------------- */
function renderClients() {
  const card = $("card-clients");
  if (!S.clientInvoices.length && !S.clientSkipped.length) {
    card.classList.add("locked");
    $("clients-list").innerHTML = `<p class="hint">Upload your Xero sales-invoice export in step 1 and every invoiced client that isn't Care.com shows up here.</p>`;
    $("cl-actions").style.display = "none"; $("cl-skipped").style.display = "none"; $("client-chips").innerHTML = "";
    return;
  }
  card.classList.remove("locked");

  const live = S.clientInvoices.filter(i=>i.include);
  const already = S.clientInvoices.reduce((a,i)=>a + i.lines.filter(l=>l.already).length, 0);
  const cancN = S.clientInvoices.reduce((a,i)=>a + i.canc.length, 0);
  $("client-chips").innerHTML = [
    `<span class="chip"><b>${S.clientInvoices.length}</b> client${S.clientInvoices.length===1?"":"s"}</span>`,
    `<span class="chip"><b>${live.length}</b> invoice${live.length===1?"":"s"} to send</span>`,
    `<span class="chip good">rates read from your Xero history</span>`,
    already ? `<span class="chip good">${already} line(s) already invoiced — left off</span>` : "",
    S.clientInvoices.some(i=>i.quote) ? `<span class="chip warn">⚠ ${S.clientInvoices.filter(i=>i.quote).length} quote-billed client(s) — not generated, see below</span>` : "",
    cancN ? `<span class="chip warn">${cancN} cancelled booking(s) — add a fee line only if you agreed one</span>` : "",
    S.clientSkipped.length ? `<span class="chip warn">⚠ ${S.clientSkipped.length} booking(s) with no Xero client</span>` : "",
  ].filter(Boolean).join("");

  $("clients-list").innerHTML = S.clientInvoices.map((inv,i)=>{
    const rows = inv.lines.map(l=>`<tr class="${l.already?"done":""}">
        <td>${esc(l.desc)}${l.already?` <span class="pill skip" title="A line for this service date is already on an invoice dated ${l.priorOn}">already invoiced ${esc(l.priorOn)}</span>`:""}</td>
        <td class="q">${(+l.qty.toFixed(2))} × ${money(l.rate)}</td>
        <td class="a">${l.already?"—":money(l.qty*l.rate)}</td></tr>`).join("");
    const meta = [
      `<span class="pill multi" title="How this booking's client was matched to a Xero contact">matched by ${esc(inv.how)}</span>`,
      `<span class="pill skip">${money(inv.p.rate)}/hr from ${inv.p.invoices} past invoice${inv.p.invoices===1?"":"s"}</span>`,
      `<span class="pill skip">net ${inv.p.terms}</span>`,
      `<span class="pill skip">${esc(inv.num)} · ${fmtMDY(inv.invDate)}</span>`,
    ].join(" ");
    const quoteNote = inv.quote
      ? `<div class="cl-note">Billed from a quote, not from hours — the last invoice referenced ${esc((inv.p.addr.Reference||"a quote").trim())} and its line quantities didn't come from the bookings. Left out of the CSV; invoice this one the way you always do.</div>` : "";
    const cancNote = inv.canc.length
      ? `<div class="cl-note">${inv.canc.length} cancelled booking(s) this month: ${esc(inv.canc.map(c=>`${c.date.slice(5)} ${c.cg||"unassigned"}`).join(", "))}. Nothing is billed for these unless you add a line below.</div>` : "";
    const otherNote = inv.other.length
      ? `<div class="cl-note">${inv.other.length} booking(s) still ${esc([...new Set(inv.other.map(o=>o.status))].join("/"))} — not invoiced.</div>` : "";
    return `<div class="cl${inv.include?"":" off"}">
      <div class="cl-h">
        <label style="display:flex;gap:9px;align-items:center;cursor:pointer">
          <input type="checkbox" class="inp-clinc" data-ci="${i}"${inv.include?" checked":""}${inv.quote?" disabled":""}>
          <span class="nm">${esc(inv.contact)}</span>
        </label>
        <span class="to">${esc(inv.p.email||"no email on file")}</span>
        <span class="amt" data-cltot="${i}">${money(inv.total)}</span>
      </div>
      <div class="cl-meta">${meta}</div>
      ${quoteNote}${cancNote}${otherNote}
      <div class="cl-b"><table>${rows || `<tr><td colspan="3" class="hint">Nothing new to invoice this month.</td></tr>`}</table>
        <div class="cl-add">
          <span>Extra line</span>
          <input type="text" class="inp-clx" data-ci="${i}" placeholder="e.g. Cancellation Fee: 8/14 (1 caregiver x 6 hours) — or Resort Fee">
          <input type="number" class="inp-clxa" data-ci="${i}" step="0.01" placeholder="amount" style="width:110px">
        </div>
      </div>
    </div>`;
  }).join("");

  if (S.clientSkipped.length) {
    $("cl-skipped").style.display = "block";
    $("cl-skipped-list").innerHTML = S.clientSkipped.map(x=>`<li><span class="pill ${x.cls}">no Xero client</span> ${esc(x.txt)}</li>`).join("");
  } else $("cl-skipped").style.display = "none";

  $("cl-actions").style.display = S.clientInvoices.some(i=>i.include) ? "flex" : "none";
  document.querySelectorAll(".inp-clinc,.inp-clx,.inp-clxa").forEach(el=>el.addEventListener("input", syncClients));
}

function syncClients() {
  document.querySelectorAll(".inp-clinc").forEach(cb=>{ S.clientInvoices[+cb.dataset.ci].include = cb.checked; });
  document.querySelectorAll(".inp-clx").forEach(inp=>{ S.clientInvoices[+inp.dataset.ci].extra.desc = inp.value; });
  document.querySelectorAll(".inp-clxa").forEach(inp=>{ S.clientInvoices[+inp.dataset.ci].extra.amt = inp.value; });
  S.clientInvoices.forEach((inv,i)=>{
    const ex = (inv.extra.desc.trim() && parseFloat(inv.extra.amt)) ? parseFloat(inv.extra.amt) : 0;
    const base = inv.lines.filter(l=>!l.already).reduce((a,l)=>a+l.qty*l.rate, 0);
    inv.total = Math.round((base+ex)*100)/100;
    const el = document.querySelector(`[data-cltot="${i}"]`); if (el) el.textContent = money(inv.total);
    const box = document.querySelectorAll(".cl")[i]; if (box) box.classList.toggle("off", !inv.include);
  });
  $("cl-actions").style.display = S.clientInvoices.some(i=>i.include) ? "flex" : "none";
}

$("btn-xero").addEventListener("click", ()=>{
  const err = $("err-clients"); err.style.display = "none";
  try {
    syncClients();
    const csv = buildXeroCSV();
    if (csv.trim().split("\n").length < 2) throw new Error("nothing selected to invoice");
    const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `Sitterwise_XeroImport_${MONTHS[S.month-1]}${S.year}.csv`; a.click();
    URL.revokeObjectURL(url);
    setStep(4);
  } catch(e) { err.textContent = "Couldn't build the CSV: " + e.message; err.style.display = "block"; }
});

/* ---------------- stepper ---------------- */
function setStep(n) {
  document.querySelectorAll(".step").forEach(s=>{
    const k=+s.dataset.step;
    s.classList.toggle("active", k===n);
    s.classList.toggle("done", k<n);
  });
}
