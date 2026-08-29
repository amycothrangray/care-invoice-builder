/* School Photo Roster Fixer -- the screen.

   All the thinking lives in lib.js. This file reads the dropped file, draws
   what the engine found, and hands every choice the user makes straight back
   to it: the engine re-runs from scratch each time, so nothing has to be undone
   and the download always matches what is on screen. */

"use strict";

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? "" : s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const plural = (n, one, many) => n + " " + (n === 1 ? one : (many || one + "s"));

const S = {
  file: null, book: null, result: null,
  opts: { sheets: {}, mapping: {}, implied: {}, decisions: {}, fixes: {}, names: {} },
  touched: {}, showReviews: 30,
  step: 1
};

/* ------------------------------------------------------------- the file */

function wireDrop() {
  const drop = $("drop"), input = $("file");
  drop.addEventListener("click", () => input.click());
  input.addEventListener("change", () => { if (input.files[0]) load(input.files[0]); });
  ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.add("over");
  }));
  ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => {
    e.preventDefault(); drop.classList.remove("over");
  }));
  drop.addEventListener("drop", e => {
    if (e.dataTransfer.files[0]) load(e.dataTransfer.files[0]);
  });
}

function readFile(file, asText) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error("Could not read that file."));
    if (asText) r.readAsText(file); else r.readAsArrayBuffer(file);
  });
}

async function load(file) {
  $("err1").innerHTML = "";
  S.file = file;
  S.opts = { sheets: {}, mapping: {}, implied: {}, decisions: {}, fixes: {}, names: {} };
  S.touched = {};
  S.showReviews = 30;
  $("fname").textContent = file.name;
  $("drop").classList.add("filled");
  try {
    const isText = /\.(csv|txt|tsv)$/i.test(file.name);
    const data = await readFile(file, isText);
    S.book = RF.readWorkbook(XLSX, isText ? data : new Uint8Array(data), file.name);
    if (!S.book.sheets.length) throw new Error("That workbook has no sheets in it.");
    run();
    setStep(2);
    $("card2").scrollIntoView({ block: "start" });
  } catch (e) {
    $("err1").innerHTML = '<div class="err"><b>Could not open that file.</b> ' + esc(e.message) +
      " If it opens in Excel, try File &rarr; Save As &rarr; .xlsx and drop it in again.</div>";
    $("drop").classList.remove("filled");
    S.book = null;
  }
}

/* ------------------------------------------------------- run + draw all */

function run() {
  if (!S.book) return;
  S.result = RF.analyze(S.book, S.opts);
  renderSheets();
  renderMapping();
  renderDash();
  renderReviews();
  renderWarnings();
  renderDownload();
  renderSidebar();
}

/* ---------------------------------------------------------- 2. the tabs */

function renderSheets() {
  const r = S.result;
  const multi = r.sheets.length > 1;
  const html = [];
  if (multi) {
    html.push('<h3>Tabs in this workbook</h3>');
    html.push('<div class="sub" style="margin-bottom:10px">Ticked tabs get combined into one roster.</div>');
  }
  r.sheets.forEach(s => {
    if (!multi && s.included) return;
    const rows = s.headerRow >= 0 ? plural(s.dataRows, "row") : "";
    html.push('<div class="sheet' + (s.included ? " on" : "") + '">' +
      '<div class="top">' +
        '<label class="ck"><input type="checkbox" data-sheet="' + esc(s.name) + '"' +
          (s.included ? " checked" : "") + '> <span class="nm">' + esc(s.name) + "</span></label>" +
        '<span class="meta">' + esc(rows) +
          (s.headerRow >= 0 ? " &middot; header on row " + (s.headerRow + 1) : "") + "</span>" +
        (s.looksLikeRoster ? '<span class="tag sure">looks like a roster</span>'
          : '<span class="tag off">' + esc(s.why || "not a roster") + "</span>") +
      "</div>" +
      (s.headerRow >= 0 ? '<div class="hdr" title="' + esc(s.headerText) + '">' + esc(s.headerText) + "</div>" : "") +
      (s.included && s.offerGrade ? offerLine("g:" + s.name, s.useGrade,
        'This tab has no grade column. Use the tab name and set Grade = <b>' + esc(s.implies.grade) +
        "</b> where it is blank?") : "") +
      (s.included && s.offerTeacher ? offerLine("t:" + s.name, s.useTeacher,
        'This tab has no teacher column. Use the tab name and set Teacher = <b>' +
        esc(s.implies.teacher) + "</b> where it is blank?") : "") +
      "</div>");
  });
  $("sheets").innerHTML = html.join("");
  $("sheets").querySelectorAll("input[data-sheet]").forEach(cb => {
    cb.addEventListener("change", () => { S.opts.sheets[cb.dataset.sheet] = cb.checked; run(); });
  });
  $("sheets").querySelectorAll("input[data-implied]").forEach(cb => {
    cb.addEventListener("change", () => { S.opts.implied[cb.dataset.implied] = cb.checked; run(); });
  });
}

const offerLine = (k, on, text) => '<div class="offer"><label class="ck"><input type="checkbox" ' +
  'data-implied="' + esc(k) + '"' + (on ? " checked" : "") + "> <span>" + text + "</span></label></div>";

/* ------------------------------------------------------ 2. the mapping */

function fieldOptions(slots) {
  const F = RF.F, L = RF.FIELD_LABEL;
  const out = [["ignore", "Not exported"]];
  [F.FIRST, F.LAST, F.MIDDLE, F.FULL, F.SUFFIX, F.GRADE, F.TEACHER, F.ID]
    .forEach(f => out.push([f, L[f]]));
  for (let s = 1; s <= slots; s++) {
    [F.P_FIRST, F.P_LAST, F.P_FULL, F.P_EMAIL, F.P_CELL].forEach(f => out.push([f + ":" + s, L[f] + " " + s]));
  }
  return out;
}

function renderMapping() {
  const r = S.result;
  const on = r.sheets.filter(s => s.included && s.headerRow >= 0);
  if (!on.length) {
    $("mapping").innerHTML = '<div class="err">No tab with student data is selected. Tick one above, or ' +
      "try a different export from the school.</div>";
    $("btn-fix").disabled = true;
    $("mapbadge").style.display = "none";
    return;
  }
  $("btn-fix").disabled = false;
  const maxSlot = on.reduce((n, s) => s.cols.reduce((m, c) => Math.max(m, c.slot || 0), n), 0);
  const slots = Math.max(3, Math.min(maxSlot + 1, 8));
  const opts = fieldOptions(slots);
  const seen = new Map();
  const html = [];

  on.forEach(info => {
    if (seen.has(info.sig)) {
      html.push('<h3>' + esc(info.name) + '</h3><div class="sub">Same columns as <b>' +
        esc(seen.get(info.sig)) + "</b> &mdash; one mapping covers both.</div>");
      return;
    }
    seen.set(info.sig, info.name);
    if (on.length > 1) html.push("<h3>" + esc(info.name) + "</h3>");
    html.push('<div class="scroll"><table class="t"><thead><tr>' +
      "<th>School column</th><th>Example</th><th>Exported as</th><th></th></tr></thead><tbody>");
    info.cols.forEach(c => {
      const cur = RF.PARENT_FIELDS.has(c.field) ? c.field + ":" + (c.slot || 1) : c.field;
      const tag = c.overridden ? '<span class="tag you">you set this</span>'
        : c.field === "ignore" ? (c.sensitive ? '<span class="tag off">' + esc(c.sensitive) + "</span>" : "")
        : c.conf === "check" ? '<span class="tag check">check this</span>'
        : '<span class="tag sure">sure</span>';
      html.push("<tr><td><b>" + esc(c.header) + "</b>" +
        (c.note ? '<span class="why">' + esc(c.note) + "</span>" : "") + "</td>" +
        '<td><span class="sample" title="' + esc(c.sample) + '">' + esc(c.sample || "-") + "</span></td>" +
        '<td><select data-map="' + esc(c.mapKey) + '">' +
        opts.map(o => '<option value="' + o[0] + '"' + (o[0] === cur ? " selected" : "") + ">" +
          esc(o[1]) + "</option>").join("") + "</select></td>" +
        "<td>" + tag + "</td></tr>");
    });
    html.push("</tbody></table></div>");
  });

  $("mapping").innerHTML = html.join("");
  $("mapping").querySelectorAll("select[data-map]").forEach(sel => {
    sel.addEventListener("change", () => { S.opts.mapping[sel.dataset.map] = sel.value; run(); });
  });

  const checks = on.reduce((n, s) => n + s.cols.filter(c => c.conf === "check" && c.field !== "ignore").length, 0);
  $("mapbadge").textContent = checks ? plural(checks, "column") + " worth a look" : "";
  $("mapbadge").style.display = checks ? "" : "none";
  $("maphint").innerHTML = plural(r.stats.sourceRows, "source row") + " on " +
    plural(on.length, "tab") + " &middot; " + plural(r.unused.length, "column") + " will not be exported";
}

/* ------------------------------------------------------- 3. the dashboard */

function renderDash() {
  const r = S.result, st = r.stats, c = st.counts;
  const fixed = [
    ["Duplicate household rows merged", st.rowsMerged],
    ["Exact duplicate rows removed", st.exactDupes],
    ["Duplicate parent contacts collapsed", st.dupContacts],
    ["Names standardized", st.namesStandardized],
    ["Grade values normalized", st.gradesFixed],
    ["Teacher names tidied", st.teachersFixed],
    ["Phone numbers normalized", st.phonesFixed],
    ["Email addresses cleaned", st.emailsFixed],
    ["Cells holding two contacts split", st.splitCells],
    ["Grades filled from the tab name", st.impliedGrade],
    ["Teachers filled from the tab name", st.impliedTeacher],
    ["Blank rows ignored", st.blankRows],
    ["Title / total / footer rows ignored", st.junkRows],
    ["Repeated header rows ignored", st.repeatHeaders],
    ["Merged cells filled down", st.merges],
    ["Formulas flattened to values", st.formulas],
    ["Odd characters cleaned up", st.chars.odd]
  ].filter(x => x[1] > 0);

  const review = [
    ["Possible duplicate students", c.duplicates, "reviews"],
    ["One Student ID, two names", c.idConflicts, "reviews"],
    ["Missing student name", c["no-name"], "w-no-name"],
    ["Missing last name", c["missing-last"], "w-missing-last"],
    ["Missing first name", c["missing-first"], "w-missing-first"],
    ["Names needing a split check", c["name-parse"], "reviews"],
    ["Invalid emails", c["bad-email"], "reviews"],
    ["Invalid phones", c["bad-phone"], "reviews"],
    ["Missing grade", c["missing-grade"], "w-missing-grade"],
    ["Unusual grade values", c["odd-grade"], "w-odd-grade"],
    ["Missing teacher", c["missing-teacher"], "w-missing-teacher"],
    ["Missing Student ID", c["missing-id"], "w-missing-id"],
    ["No parent contact at all", c["no-contact"], "w-no-contact"],
    ["Parent phone without email", c["phone-no-email"], "w-phone-no-email"],
    ["Same student spelled two ways", c["name-variant"], "w-name-variant"],
    ["Conflicting grade after merge", c["grade-conflict"], "w-grade-conflict"],
    ["Conflicting teacher after merge", c["teacher-conflict"], "w-teacher-conflict"],
    ["Conflicting Student ID after merge", c["id-conflict-merged"], "w-id-conflict-merged"],
    ["Teacher spelled more than one way", c.teacherVariants, "w-teacher-variants"]
  ].filter(x => x[1] > 0);

  const cols = r.unused.slice(0, 14).map(u =>
    '<div class="stat"><span>' + esc(u.header) + (r.stats.sheetsUsed > 1 ? ' <span class="pill">' +
      esc(u.sheet) + "</span>" : "") + "</span><span class=\"hint\">" + esc(u.why) + "</span></div>").join("");

  $("dash").innerHTML =
    '<div class="dash">' +
      '<div class="panel"><h4>Students</h4>' +
        '<div class="count-big">' + st.students + "</div>" +
        '<div class="hint" style="margin-top:4px">from ' + plural(st.sourceRows, "source row") +
        " on " + plural(st.sheetsUsed, "tab") + "</div>" +
        '<div class="stat" style="margin-top:8px"><span>With a parent email</span><b>' + st.withEmail + "</b></div>" +
        '<div class="stat"><span>With a parent phone</span><b>' + st.withPhone + "</b></div>" +
        '<div class="stat"><span>Parent columns each</span><b>' + st.parentSlots + "</b></div>" +
      "</div>" +
      '<div class="panel good"><h4>Fixed automatically</h4>' +
        (fixed.length ? fixed.map(x => '<div class="stat"><span>' + x[0] + "</span><b>" + x[1] + "</b></div>").join("")
          : '<div class="none">Nothing needed fixing.</div>') +
      "</div>" +
      '<div class="panel review"><h4>Needs review</h4>' +
        (review.length ? review.map(x => '<div class="stat"><span><a href="#" data-jump="' + x[2] + '">' +
          x[0] + "</a></span><b>" + x[1] + "</b></div>").join("")
          : '<div class="none">Nothing to decide. Every row came through clean.</div>') +
      "</div>" +
      '<div class="panel cols"><h4>Not exported (' + r.unused.length + ")</h4>" +
        (cols || '<div class="none">Every column is being used.</div>') +
        (r.unused.length > 14 ? '<div class="hint">+ ' + (r.unused.length - 14) + " more</div>" : "") +
        (r.unused.some(u => u.sensitive) ? '<div class="hint" style="margin-top:6px">Personal columns are ' +
          "left out on purpose. Change any of them on the mapping screen if you need one.</div>" : "") +
      "</div>" +
    "</div>";

  $("dash").querySelectorAll("a[data-jump]").forEach(a => a.addEventListener("click", e => {
    e.preventDefault();
    const t = $(a.dataset.jump);
    if (t) { t.scrollIntoView({ block: "center" }); if (t.tagName === "DETAILS") t.open = true; }
  }));
}

/* ---------------------------------------------------- 3. the review cards */

function cardHTML(c) {
  return '<div class="pc"><div class="pn">' + esc(c.name) + "</div>" +
    '<div class="pl">Grade ' + esc(c.grade) + " &middot; " + esc(c.teacher) + " &middot; ID: " + esc(c.id) + "</div>" +
    (c.contacts.length ? c.contacts.map(x => '<div class="pe">' + esc(x) + "</div>").join("")
      : '<div class="pe" style="color:#6B7A87">no parent contact</div>') +
    '<div class="pl">' + esc(c.where.map(w => "row " + w.split(":")[1]).join(", ")) + "</div></div>";
}

function renderReviews() {
  const r = S.result;
  const html = [];
  const open = r.reviews.filter(x => x.type === "id-conflict" || x.choice !== "separate" || true);

  if (r.reviews.length) {
    html.push("<h3>Is this one child or two?</h3>");
    html.push('<div class="sub">Nothing here has been merged. Photos get mixed up when two ' +
      "children become one row, so these wait for you.</div>");
    const shown = r.reviews.slice(0, S.showReviews);
    shown.forEach(rev => {
      const settled = !!S.touched[rev.id];
      html.push('<div class="rev' + (settled ? " settled" : "") + '">' +
        '<div class="rt">' + esc(rev.title) + "</div>" +
        '<div class="rd">' + esc(rev.detail) + "</div>" +
        '<div class="pair">' + rev.cards.map(cardHTML).join("") + "</div>" +
        '<div class="opts">' + rev.options.map(o =>
          '<button class="opt' + (o.value === rev.choice ? " on" : "") + '" data-rev="' + esc(rev.id) +
          '" data-val="' + esc(o.value) + '">' + esc(o.label) +
          (o.desc ? "<small>" + esc(o.desc) + "</small>" : "") + "</button>").join("") +
        "</div></div>");
    });
    if (r.reviews.length > shown.length) {
      html.push('<div class="callout"><b>' + (r.reviews.length - shown.length) + " more like this.</b> " +
        "Every one of them is on the Needs Review sheet of the download too, and all of them stay as " +
        'separate students unless you say otherwise. <button class="btn btn-ghost btn-sm" id="more-rev" ' +
        'style="margin-left:6px">Show 30 more</button></div>');
    }
    if (r.stats.moreDuplicates) {
      html.push('<div class="hint">' + r.stats.moreDuplicates + " further look-alike pairs were found " +
        "beyond the first 400 and left as separate students. They are counted in the cleanup report.</div>");
    }
    if (r.stats.crowdedNames) {
      html.push('<div class="hint">' + r.stats.crowdedNames + " students sit in name blocks too large to " +
        "compare pair by pair, so only exact-name matches were checked there.</div>");
    }
  }

  const names = r.issues.filter(i => i.type === "name-parse");
  if (names.length) {
    html.push("<h3>Which part is the last name?</h3>");
    html.push('<div class="sub">Multi-word names can go either way. The last name is right either ' +
      "way &mdash; this is only about how much of it is the first name.</div>");
    names.forEach(i => {
      const s = i.student;
      const a = { first: s.first, last: s.last }, b = { first: i.alt, last: s.last };
      const chosen = S.opts.names[i.nameKey];
      html.push('<div class="rev' + (chosen ? " settled" : "") + '">' +
        '<div class="rt">' + esc(i.raw) + "</div>" +
        '<div class="rd">' + esc(s.grade ? "Grade " + s.grade : "") +
          (s.teacher ? " &middot; " + esc(s.teacher) : "") + "</div>" +
        '<div class="opts" style="margin-top:9px">' +
        [a, b].filter(x => x.first).map(x =>
          '<button class="opt' + (chosen ? (chosenName(i.nameKey, x) ? " on" : "")
            : (x.first === s.first ? " on" : "")) + '" data-name="' + esc(i.nameKey) +
          '" data-first="' + esc(x.first) + '" data-last="' + esc(x.last) + '">' +
          esc(x.first) + " / " + esc(x.last) + "<small>first / last</small></button>").join("") +
        "</div></div>");
    });
  }

  const bad = r.issues.filter(i => i.type === "bad-email" || i.type === "bad-phone");
  if (bad.length) {
    html.push("<h3>" + plural(bad.length, "contact") + " the import will choke on</h3>");
    html.push('<div class="sub">These are held out of the export until they are right &mdash; nothing is ' +
      "lost, they are all on the Needs Review sheet too. Fix one here and it goes straight in.</div>");
    bad.forEach(i => {
      const s = i.student;
      html.push('<div class="rev' + (S.opts.fixes[i.fixKey] !== undefined ? " settled" : "") + '">' +
        '<div class="rt">' + esc([s.first, s.last].filter(Boolean).join(" ") || "(no name)") +
        ' <span class="pill">' + esc(i.kind) + "</span></div>" +
        '<div class="rd">' + esc(i.who) + ": " + esc(i.value || "(blank)") + " &mdash; " +
        esc(i.text.split(" - ").pop()) + "</div>" +
        '<div class="fixrow"><input type="text" data-fix="' + esc(i.fixKey) + '" value="' +
        esc(S.opts.fixes[i.fixKey] !== undefined ? S.opts.fixes[i.fixKey] : i.value) +
        '" placeholder="' + (i.kind === "email" ? "name@example.com" : "(619) 555-1212") + '">' +
        '<button class="btn btn-ghost btn-sm" data-apply="' + esc(i.fixKey) + '">Use this</button>' +
        '<button class="btn btn-ghost btn-sm" data-drop="' + esc(i.fixKey) + '">Leave it out</button>' +
        "</div></div>");
    });
  }

  $("reviews").innerHTML = html.join("");

  const more = $("more-rev");
  if (more) more.addEventListener("click", () => { S.showReviews += 30; renderReviews(); });
  $("reviews").querySelectorAll("button[data-rev]").forEach(b => b.addEventListener("click", () => {
    S.opts.decisions[b.dataset.rev] = b.dataset.val;
    S.touched[b.dataset.rev] = true;
    run();
  }));
  $("reviews").querySelectorAll("button[data-name]").forEach(b => b.addEventListener("click", () => {
    S.opts.names[b.dataset.name] = { first: b.dataset.first, last: b.dataset.last };
    S.touched[b.dataset.name] = true;
    run();
  }));
  $("reviews").querySelectorAll("button[data-apply]").forEach(b => b.addEventListener("click", () => {
    const input = $("reviews").querySelector('input[data-fix="' + b.dataset.apply.replace(/"/g, '\\"') + '"]');
    S.opts.fixes[b.dataset.apply] = input ? input.value : "";
    run();
  }));
  $("reviews").querySelectorAll("button[data-drop]").forEach(b => b.addEventListener("click", () => {
    S.opts.fixes[b.dataset.drop] = "";
    run();
  }));
  $("reviews").querySelectorAll("input[data-fix]").forEach(inp => inp.addEventListener("keydown", e => {
    if (e.key === "Enter") { S.opts.fixes[inp.dataset.fix] = inp.value; run(); }
  }));
}

function chosenName(k, x) {
  const c = S.opts.names[k];
  return c && c.first === x.first && c.last === x.last;
}

/* ------------------------------------------------------- 3. the warnings */

const WARN_TITLE = {
  "no-name": "Rows with no student name", "missing-last": "Missing last name",
  "missing-first": "Missing first name", "missing-grade": "Missing grade",
  "odd-grade": "Unusual grade values", "missing-teacher": "Missing teacher",
  "missing-id": "Missing Student ID", "no-contact": "No parent contact at all",
  "phone-no-email": "Parent phone with no email", "name-variant": "Spelled two ways",
  "grade-conflict": "Merged rows disagree on grade", "teacher-conflict": "Merged rows disagree on teacher",
  "id-conflict-merged": "Merged rows carry different Student IDs"
};
const WARN_NOTE = {
  "phone-no-email": "GotPhoto's names-list import needs an email for a buyer contact. A phone-only " +
    "parent will not come through. No fake addresses were invented.",
  "missing-teacher": "Most of this roster has a teacher, so these stand out.",
  "missing-id": "Most of this roster has a Student ID, so these stand out.",
  "odd-grade": "Left exactly as the school wrote it - never turned into a made-up grade number."
};

function renderWarnings() {
  const r = S.result;
  const order = ["no-name", "missing-last", "missing-first", "no-contact", "id-conflict-merged",
    "name-variant", "grade-conflict", "teacher-conflict", "odd-grade", "missing-grade",
    "missing-teacher", "missing-id", "phone-no-email"];
  const html = [];
  order.forEach(type => {
    const list = r.issues.filter(i => i.type === type);
    if (!list.length) return;
    const critical = list[0].severity === "critical";
    html.push('<details id="w-' + type + '"' + (critical ? " open" : "") + "><summary>" +
      esc(WARN_TITLE[type] || type) + '<span class="c' + (critical ? "" : " q") + '">' + list.length + "</span>" +
      "</summary>" + (WARN_NOTE[type] ? '<div class="hint" style="margin-top:8px">' + WARN_NOTE[type] + "</div>" : "") +
      '<div class="list">' + list.slice(0, 300).map(i => {
        const s = i.student;
        return '<div class="li"><span class="who">' +
          esc([s.first, s.last].filter(Boolean).join(" ") || "(no name)") + "</span>" +
          "<span>" + esc(i.text) + "</span>" +
          '<span class="where">' + esc((s.sources[0] || {}).sheet || "") + " row " +
          esc((s.sources[0] || {}).row || "?") + "</span></div>";
      }).join("") +
      (list.length > 300 ? '<div class="hint">+ ' + (list.length - 300) + " more, all on the Needs Review sheet</div>" : "") +
      "</div></details>");
  });

  if (r.teacherVariants.length) {
    html.push('<details id="w-teacher-variants"><summary>Teacher spelled more than one way' +
      '<span class="c q">' + r.teacherVariants.length + "</span></summary>" +
      '<div class="hint" style="margin-top:8px">These may be one classroom written two ways. Nothing was ' +
      "merged - change them in the source file if they should match.</div><div class=\"list\">" +
      r.teacherVariants.map(t => '<div class="li"><span class="who">' + esc(t.key) + "</span><span>" +
        esc(t.names.map(n => n.name + " (" + plural(n.n, "student") + ")").join("  vs  ")) +
        "</span></div>").join("") + "</div></details>");
  }
  if (r.junk.length) {
    html.push('<details id="w-junk"><summary>Rows skipped as titles, totals or footers' +
      '<span class="c q">' + r.junk.length + "</span></summary><div class=\"list\">" +
      r.junk.slice(0, 100).map(j => '<div class="li"><span class="who">' + esc(j.sheet) + " row " + j.row +
        '</span><span class="mono">' + esc(j.text) + "</span></div>").join("") + "</div></details>");
  }
  $("warnings").innerHTML = html.join("");
}

/* -------------------------------------------------------- 4. download */

function renderDownload() {
  const r = S.result, st = r.stats;
  const left = st.counts.duplicates + st.counts.idConflicts + st.counts["bad-email"] + st.counts["bad-phone"];
  $("dlinfo").innerHTML =
    '<div class="callout teal"><b>' + esc(RF.outputFileName(S.file ? S.file.name : "roster")) + "</b><br>" +
    "Sheet 1 <b>GOTPHOTO READY</b> &mdash; " + plural(st.students, "student") + ", " +
    r.columns.length + " columns, every value as text so IDs keep their leading zeros.<br>" +
    "Sheet 2 <b>NEEDS REVIEW</b> &mdash; everything above that a human should look at.<br>" +
    "Sheet 3 <b>CLEANUP REPORT</b> &mdash; what changed, and which columns were left out.</div>" +
    (left ? '<div class="callout"><b>' + plural(left, "thing") + " still waiting on you</b> in step 3. " +
      "You can download anyway &mdash; unresolved rows stay separate and are listed on the Needs Review " +
      "sheet.</div>" : "");
  renderAudit();
}

function renderAudit() {
  const r = S.result;
  const rows = r.students.slice(0, 400).map(s =>
    "<tr><td><b>" + esc([s.first, s.last].filter(Boolean).join(" ")) + "</b></td>" +
    "<td>" + esc(s.grade) + "</td><td>" + esc(s.teacher) + "</td>" +
    '<td class="mono">' + esc(s.id) + "</td>" +
    '<td class="mono">' + esc(s.sources.map(x => x.sheet + ":" + x.row +
      (x.exact ? " (exact dup)" : x.merged ? " (merged)" : "")).join(", ")) + "</td>" +
    '<td class="hint">' + esc(s.changes.join("; ")) +
      (s.dropped.length ? " &mdash; not exported: " + esc(s.dropped.map(d => d.value).join(", ")) : "") +
    "</td></tr>").join("");
  $("audit").innerHTML = '<details id="auditbox"><summary>Where every student came from' +
    '<span class="c q">' + r.students.length + "</span></summary>" +
    '<div class="hint" style="margin:8px 0">Every final row, the source rows behind it, and what was ' +
    "changed on the way. Tick the source-trace box above to get this as a fourth sheet in the download.</div>" +
    '<div class="scroll" style="max-height:420px;overflow:auto"><table class="t"><thead><tr>' +
    "<th>Student</th><th>Grade</th><th>Teacher</th><th>ID</th><th>Came from</th><th>What changed</th>" +
    "</tr></thead><tbody>" + rows + "</tbody></table></div>" +
    (r.students.length > 400 ? '<div class="hint">Showing the first 400.</div>' : "") + "</details>";
}

async function download() {
  const btn = $("btn-dl");
  btn.disabled = true;
  const was = btn.textContent;
  btn.textContent = "Building...";
  try {
    const buf = await RF.buildWorkbook(ExcelJS, S.result, { trace: $("trace").checked });
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = RF.outputFileName(S.file ? S.file.name : "roster");
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    $("dlnote").innerHTML = '<div class="hint" style="margin-top:10px">Saved. Your original file was ' +
      "not touched.</div>";
  } catch (e) {
    $("dlnote").innerHTML = '<div class="err">Could not build the file: ' + esc(e.message) + "</div>";
  }
  btn.textContent = was;
  btn.disabled = false;
}

/* ---------------------------------------------------------- sidebar */

function renderSidebar() {
  const st = S.result.stats, c = st.counts;
  const needs = c.duplicates + c.idConflicts + c["bad-email"] + c["bad-phone"] + c["no-name"] +
    c["missing-last"] + c["missing-first"];
  $("sb-count").textContent = st.students;
  $("sb-strip").innerHTML =
    row("Source rows", st.sourceRows) +
    row("Rows merged", st.rowsMerged, st.rowsMerged ? "good" : "") +
    row("Exact dupes gone", st.exactDupes, st.exactDupes ? "good" : "") +
    row("Needs you", needs, needs ? "warn" : "good") +
    row("Not exported", S.result.unused.length);
  $("sb-note").innerHTML = needs
    ? "Step 3 has " + plural(needs, "decision") + " waiting. Nothing was merged without asking."
    : "Everything resolved. The download is ready whenever you are.";
}
const row = (l, v, cls) => '<div class="srow ' + (cls || "") + '"><span class="l">' + l + "</span><b>" + v + "</b></div>";

/* ------------------------------------------------------------- steps */

function setStep(n) {
  S.step = n;
  for (let i = 1; i <= 4; i++) {
    const card = $("card" + i);
    if (i <= n) card.classList.remove("locked"); else card.classList.add("locked");
  }
  document.querySelectorAll(".step").forEach(el => {
    const s = parseInt(el.dataset.step, 10);
    el.classList.toggle("active", s === n);
    el.classList.toggle("done", s < n);
  });
}

$("btn-fix").addEventListener("click", () => {
  setStep(4);
  $("card3").scrollIntoView({ block: "start" });
});
$("btn-dl").addEventListener("click", download);
document.querySelectorAll(".step").forEach(el => el.addEventListener("click", () => {
  const s = parseInt(el.dataset.step, 10);
  if (S.book && s <= S.step) $("card" + s).scrollIntoView({ block: "start" });
}));
wireDrop();
