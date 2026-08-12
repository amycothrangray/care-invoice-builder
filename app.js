/* Sitterwise → Care.com Invoice Builder
   All logic runs client-side. Rules mirror the monthly workflow:
   $42/hr (editable) · CA daily OT 1.5× over 8h/day · 4-hr minimum ·
   mileage over 40mi RT at $0.76/mi (editable) · cancellation reconciliation
   with reassignment-shadow + couldn't-fill exclusion and multi-day grouping. */

"use strict";

const S = {
  careFile: null, bookFile: null, milFile: null,
  jobs: [],            // completed Care.com jobs
  cancBillable: [],    // billable cancellations (with UI decisions)
  cancExcluded: [],    // {why, label}
  unmatched: [],       // jobs needing client names
  mileageCand: [],     // mileage candidates on the invoice
  month: null, year: null,
  weeklyOTWarn: [],
  adminNotes: [],      // internal guidance only — never written to the invoice
  dblBill: [],         // completed jobs that may duplicate a billed cancellation
  milSkipped: [],      // mileage form entries that don't belong on this invoice
  milSource: "none",   // "form" (Cognito export) or "reimb" (estimated fallback)
  excludeJobs: new Set(), // completed Care jobs excluded from the invoice (not actually worked)
  built: null,         // {blob, filename, total}
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
    const [careWB, bookWB, milWB] = await Promise.all([readWB(S.careFile), readWB(S.bookFile), S.milFile ? readWB(S.milFile) : Promise.resolve(null)]);
    analyze(careWB, bookWB, milWB);
  } catch(e) {
    errBox.textContent = "Couldn't read the files: " + e.message + " — make sure these are the two monthly .xlsx exports.";
    errBox.style.display = "block";
  }
});

function sheetRows(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
}

function parseMileageForm(wb) {
  const rows = sheetRows(wb);
  if (!rows.length) return [];
  const hdr = rows[0].map(h => String(h||"").toLowerCase());
  const find = re => hdr.findIndex(h => re.test(h));
  const jc = find(/job/), nc = find(/name/), oc = find(/over/), tc = find(/total number|number of miles/), ac = find(/amount/);
  if (jc === -1) throw new Error('mileage file: no "Care.com Job #" column found — is this the mileage request export?');
  const out = [];
  for (let i=1;i<rows.length;i++){
    const r = rows[i]; if (!r || r[jc]==null || String(r[jc]).trim()==="") continue;
    const total = tc>-1 ? money2num(r[tc]) : 0;
    let over = oc>-1 ? money2num(r[oc]) : 0;
    if (!over && total) over = Math.max(Math.round(total-40), 0);
    out.push({ job:String(r[jc]).trim(), name: nc>-1?cleanName(r[nc]):"", over, total, amt: ac>-1?money2num(r[ac]):0 });
  }
  return out;
}

function analyze(careWB, bookWB, milWB) {
  const care = sheetRows(careWB);
  const book = sheetRows(bookWB);

  // ---- bookings columns by header name ----
  const bh = book[0].map(h => String(h||"").trim());
  const bc = name => bh.indexOf(name);
  const need = ["Booking ID","Client Name","Service Type","Start Date","Start Time","End Time","Total Hours","Caregiver Name","Status","Reimbursement","Tip","Bonus"];
  for (const n of need) if (bc(n) === -1) throw new Error(`bookings file is missing the "${n}" column`);

  const bRows = [];
  for (let i=1;i<book.length;i++){
    const r = book[i]; if (!r || r.every(c=>c===null||c==="")) continue;
    bRows.push({
      bk: r[bc("Booking ID")], client: cleanName(r[bc("Client Name")]),
      svc: String(r[bc("Service Type")]||"").trim(),
      date: isoDate(r[bc("Start Date")]), start: hhmm(r[bc("Start Time")]), end: hhmm(r[bc("End Time")]),
      hrs: money2num(r[bc("Total Hours")]),
      cg: cleanName(r[bc("Caregiver Name")]), status: String(r[bc("Status")]||"").trim().toLowerCase(),
      reimb: money2num(r[bc("Reimbursement")]), tip: money2num(r[bc("Tip")]),
      note: bc("Admin Notes") !== -1 ? String(r[bc("Admin Notes")] ?? "").trim() : "",
    });
  }
  const corp = bRows.filter(r => r.svc === "Corporate (Invoiced)");

  // client lookup: (date|start) -> [{cgFirst, client}], plus (cgLast|date) fallback
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

  // ---- Care export rows (fixed positions: A jobID, B date, C times, D multiday, E loc, F bonus, G caregiver, H status) ----
  const jobs = []; const unmatched = [];
  for (let i=1;i<care.length;i++){
    const r = care[i]; if (!r || r[0]===null || String(r[0]).trim()==="") continue;
    const jid = String(r[0]).trim();
    const date = isoDate(r[1]); const t = parseCareTimes(r[2]);
    if (!date || !t) continue;
    let cg = cleanName(r[6]); cg = CG_FIX[cg] || cg;
    const bonus = money2num(r[5]);
    // hours from the actual time span (Care's parenthetical rounds up, e.g. 4.75 -> "4.8")
    const spanMin = (()=>{ const [h1,m1]=t.start.split(":").map(Number), [h2,m2]=t.end.split(":").map(Number);
      return ((h2*60+m2)-(h1*60+m1)+1440)%1440; })();
    t.hrs = spanMin/60;
    // client match
    let client = null;
    const cand = byDT.get(date+"|"+t.start) || [];
    const cf = firstName(cg);
    for (const c of cand) { const bf = firstName(c.cg); if (bf===cf || (cf.length>=4 && bf.slice(0,4)===cf.slice(0,4))) { client=c.client; break; } }
    if (!client && cand.length===1) client = cand[0].client;
    if (!client) {
      const cd = (byCgDate.get(lastName(cg).toLowerCase()+"|"+date) || []);
      const corpCd = cd.filter(x=>x.svc==="Corporate (Invoiced)" && x.status!=="cancelled");
      if (corpCd.length===1) client = corpCd[0].client;
      else if (cd.length===1) client = cd[0].client;
    }
    const job = { jid, date, cg, client: client || "Care.com Family", start: t.start, end: t.end, hrs: t.hrs, bonus, miles: 0, oddId: !/^\d{7}$/.test(jid) };
    if (!client) unmatched.push(job);
    jobs.push(job);
  }
  if (!jobs.length) throw new Error("no completed jobs found in the Care.com export — is this the right file?");

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
      mileageCand.push({ jid: j.jid, cg: j.cg, client: j.client, date: j.date, reimb: e.amt, miles: e.over,
        include: !(j.bonus > 0), conflict: j.bonus > 0, src: "form"+how, note: "" });
    }
  } else {
    // fallback (no form uploaded): estimate from the bookings Reimbursement column
    S.milSource = "reimb";
    for (const r of corp) {
      if (r.reimb > 0) {
        const match = jobs.find(j => j.date===r.date && (firstName(j.cg)===firstName(r.cg) || firstName(j.cg).slice(0,4)===firstName(r.cg).slice(0,4)));
        if (match) mileageCand.push({ jid: match.jid, cg: match.cg, client: r.client, date: r.date, reimb: r.reimb, miles: Math.round(r.reimb/0.76), include: !(match.bonus > 0), conflict: match.bonus > 0, src: "estimated from reimbursement — upload the mileage form export for exact figures", note: r.note || "" });
        else mileageOffInvoice.push(r);
      }
    }
  }

  // tips on corporate jobs → warn (OnPay, never the invoice)
  const tips = corp.filter(r => r.tip > 0);

  // admin notes across all corporate rows (guidance only — never written to the invoice)
  S.adminNotes = corp.filter(r => r.note).map(r => ({ client:r.client, date:r.date, cg:r.cg, status:r.status, note:r.note }));

  // ---- cancellations ----
  const paidKeys = new Set(corp.filter(r=>["paid","completed","confirmed"].includes(r.status)).map(r=>r.client+"|"+r.date+"|"+r.start));
  const cancRaw = corp.filter(r=>r.status==="cancelled");
  const excluded = []; let billable = [];
  for (const r of cancRaw) {
    const key = r.client+"|"+r.date+"|"+r.start;
    if (paidKeys.has(key)) excluded.push({why:"shadow", label:`${r.client} ${r.date.slice(5)} — job was reassigned and worked (already billed as completed)`});
    else if (!r.cg) excluded.push({why:"nofill", label:`${r.client} ${r.date.slice(5)} — no caregiver assigned (agency couldn't fill; no fee)`});
    else billable.push(r);
  }
  // double-staffed: same client+date+start with 2+ billable rows → keep one, note it
  const seen = new Map();
  billable = billable.filter(r => {
    const k = r.client+"|"+r.date+"|"+r.start;
    if (seen.has(k)) { excluded.push({why:"double", label:`${r.client} ${r.date.slice(5)} — second caregiver on the same cancelled job (one fee only)`}); return false; }
    seen.set(k, true); return true;
  });
  billable.sort((a,b)=> lastName(a.client).toLowerCase().localeCompare(lastName(b.client).toLowerCase()) || a.date.localeCompare(b.date));
  // multi-day groups: same client, consecutive dates
  const groups = {}; let gid = 0;
  for (let i=0;i<billable.length;i++){
    const cur = billable[i];
    const prev = billable[i-1];
    if (prev && prev.client===cur.client) {
      const d1=new Date(prev.date), d2=new Date(cur.date);
      if ((d2-d1)/86400000 === 1) { cur._g = prev._g ?? (prev._g = ++gid); continue; }
    }
  }
  S.cancBillable = billable.map(r => ({ bk:String(r.bk), client:r.client, date:r.date, hrs:r.hrs||0, cg:r.cg, group:r._g||null, mode:">24", note:r.note||"" }));
  // double-bill check: a billable cancellation whose caregiver ALSO has a completed
  // Care.com job the same day usually means the same shift is on the invoice twice.
  S.dblBill = [];
  for (const c of S.cancBillable) {
    const hit = jobs.find(j => j.date===c.date && firstName(j.cg)===firstName(c.cg));
    if (hit) { c.dbl = hit.jid; if (!S.dblBill.some(d=>d.jid===hit.jid)) S.dblBill.push({jid:hit.jid, cg:hit.cg, client:c.client, date:c.date}); }
  }
  S.cancExcluded = excluded;
  S.jobs = jobs; S.unmatched = unmatched; S.mileageCand = mileageCand;

  // weekly OT check (warn only — historically never triggers)
  const wk = {};
  jobs.forEach(j => { const k=j.cg+"|"+isoWeek(j.date); wk[k]=(wk[k]||0)+Math.min(j.hrs,8); });
  S.weeklyOTWarn = Object.entries(wk).filter(([,v])=>v>40).map(([k,v])=>k.split("|")[0]+" ("+v.toFixed(1)+"h)");

  renderQuestions({tips, mileageOffInvoice});
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
    S.milSource==="form" ? `<span class="chip good">🚗 mileage form loaded: ${S.mileageCand.length} matched, ${S.milSkipped.length} skipped</span>` : `<span class="chip warn">no mileage form uploaded \u2014 mileage below is estimated from reimbursements</span>`,
    S.dblBill.length ? `<span class="chip warn">⚠ ${S.dblBill.length} possible double-bill(s): a billed cancellation's caregiver also has a completed job that day — see below</span>` : "",
    S.adminNotes.length ? `<span class="chip note">📝 ${S.adminNotes.length} admin note(s) in the bookings file — shown next to the calls below</span>` : "",
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
    $("tbl-mileage").innerHTML = `<tr><th>Include</th><th>Job</th><th>Caregiver</th><th>Date</th><th>Miles over 40 RT</th><th>Status</th></tr>` +
      S.mileageCand.map((m,i)=>{
        const st = m.conflict
          ? `<span class="pill less" title="Care.com does not pay mileage AND a bonus on the same job unless both are pre-approved.">\u26a0 $50 bonus on this job \u2014 pick one</span>`
          : (S.milSource==="form" ? `<span class="pill multi">\u2713 form submission</span>` : `<span class="pill skip">estimate \u2014 verify</span>`);
        return `<tr class="salrow">
        <td><input type="checkbox" data-midx="${i}" class="inp-minc"${m.include?" checked":""}></td>
        <td class="mono">${m.jid}</td><td>${m.cg}</td><td class="mono">${m.date.slice(5)}</td>
        <td><input type="number" data-midx="${i}" class="inp-miles" value="${m.miles}" min="0" style="width:80px"></td>
        <td>${st}</td></tr>${m.note?`<tr class="salrow"><td></td><td colspan="5" class="adminnote">📝 ${esc(m.note)}</td></tr>`:""}`;
      }).join("");
    if (S.milSkipped.length) {
      $("mil-skipped").style.display = "block";
      $("mil-skipped-list").innerHTML = S.milSkipped.map(s=>`<li>${esc(s)}</li>`).join("");
    } else $("mil-skipped").style.display = "none";
  } else { $("q-mileage").style.display = "none"; }

  // cancellations
  if (S.cancBillable.length) {
    $("q-canc").style.display = "block";
    $("tbl-canc").innerHTML = `<tr><th>Care.com Job ID</th><th>Client</th><th>Date</th><th>Caregiver</th><th>Booked hrs</th><th>How to bill</th></tr>` +
      S.cancBillable.map((c,i)=>{
        const grp = c.group ? `<span class="pill multi" title="Consecutive-day booking for the same client — if cancelled together, only one day should be charged">multi-day #${c.group}</span> ` : "";
        const dbl = c.dbl ? `<span class="pill less" title="This caregiver also has completed Care.com job ${c.dbl} on this date — either the cancellation or the completed job should come off">⚠ dbl ${c.dbl}</span> ` : "";
        const idWarn = /^\d{7}$/.test(c.bk) ? "" : ` title="This looks like a Sitterwise booking number — Care.com needs THEIR 7-digit job ID (find it in the Care portal)" style="border-color:var(--amber)"`;
        return `<tr>
          <td><input type="text" data-cidx="${i}" class="inp-cid mono" value="${esc(c.bk)}" style="width:90px"${idWarn}></td>
          <td>${dbl}${grp}${c.client}</td><td class="mono">${c.date.slice(5)}</td><td>${c.cg}</td>
          <td class="mono">${c.hrs ? c.hrs.toFixed(2) : "—"}</td>
          <td><select data-cidx="${i}" class="inp-cmode">
            <option value=">24" selected>&gt;24 hr — flat fee</option>
            <option value="<24">&lt;24 hr — booked hrs @ rate (cap 8)</option>
            <option value="none">No charge</option>
          </select></td></tr>${c.note?`<tr><td colspan="6" class="adminnote">📝 ${esc(c.note)}</td></tr>`:""}`;
      }).join("");
  } else $("q-canc").style.display = "none";

  $("excl-list").innerHTML = S.cancExcluded.map(e=>`<li>${e.label}</li>`).join("") || "<li>Nothing was excluded.</li>";
  $("excl-details").style.display = S.cancExcluded.length ? "block" : "none";
  if (S.adminNotes.length) {
    $("notes-list").innerHTML = S.adminNotes.map(n=>`<li><b>${esc(n.client)}</b> ${n.date.slice(5)} (${esc(n.cg||"—")}, ${esc(n.status)}): ${esc(n.note)}</li>`).join("");
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
  document.querySelectorAll(".inp-minc").forEach(cb=>{ S.mileageCand[+cb.dataset.midx].include = cb.checked; });
  document.querySelectorAll(".inp-miles").forEach(inp=>{ S.mileageCand[+inp.dataset.midx].miles = Math.max(0, parseInt(inp.value)||0); });
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
  const setCommon=(row,j)=>{ ws.getCell(row,1).value=/^\d+$/.test(j.jid)?Number(j.jid):j.jid; ws.getCell(row,2).value=j.cg; ws.getCell(row,3).value=j.client; ws.getCell(row,4).value="CA";
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
    ws.getCell(cr,3).value=c.client; ws.getCell(cr,4).value="CA";
    const dc=ws.getCell(cr,5); dc.value=dateSerial(c.date); dc.numFmt="mm-dd-yy";
    if (c.mode===">24") { ws.getCell(cr,11).value=1.00; ws.getCell(cr,12).value=a.cancFee; }
    else { const h=Math.min(c.hrs||8,8); ws.getCell(cr,11).value=h; ws.getCell(cr,12).value=a.rate;
      ws.getCell(cr,17).value=`Cancelled <24 hrs; billed ${h} hrs @ $${a.rate} (8-hr cap). Caregiver ${c.cg}.`; ws.getCell(cr,17).font=DFONT; }
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

/* ---------------- stepper ---------------- */
function setStep(n) {
  document.querySelectorAll(".step").forEach(s=>{
    const k=+s.dataset.step;
    s.classList.toggle("active", k===n);
    s.classList.toggle("done", k<n);
  });
}
