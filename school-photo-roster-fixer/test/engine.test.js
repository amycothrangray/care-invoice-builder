/* Run with:  npm test   (needs xlsx + exceljs installed)
   These are the rules the tool is not allowed to get wrong. */
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const RF = require("../lib.js");
const { messyWorkbook, householdCSV } = require("./fixtures.js");

let pass = 0, fail = 0;
const ok = (name, cond, extra) => {
  if (cond) { pass++; console.log("  ok   " + name); }
  else { fail++; console.log("  FAIL " + name + (extra !== undefined ? "  -> " + JSON.stringify(extra) : "")); }
};
const eq = (name, got, want) => ok(name + "  (" + JSON.stringify(got) + ")", JSON.stringify(got) === JSON.stringify(want), got);
const group = t => console.log("\n" + t);

/* ------------------------------------------------ pieces */
group("names");
eq("Smith, Emma", pick(RF.splitName("Smith, Emma")), ["Emma", "Smith"]);
eq("Emma Smith", pick(RF.splitName("Emma Smith")), ["Emma", "Smith"]);
eq("Mary Kate Johnson keeps Johnson", pick(RF.splitName("Mary Kate Johnson")), ["Mary", "Johnson"]);
ok("Mary Kate Johnson is flagged", RF.splitName("Mary Kate Johnson").confidence === "check");
eq("Emma R. Smith", pick(RF.splitName("Emma R. Smith")), ["Emma", "Smith"]);
ok("Emma R. Smith not flagged", RF.splitName("Emma R. Smith").confidence === "likely");
eq("Lucas van der Berg", pick(RF.splitName("Lucas van der Berg")), ["Lucas", "van der Berg"]);
eq("Jack Miller Jr.", pick(RF.splitName("Jack Miller Jr.")), ["Jack", "Miller"]);
eq("suffix kept", RF.splitName("Jack Miller Jr.").suffix, "Jr.");
eq("hyphens survive", pick(RF.splitName("Ana Lopez-Reyes")), ["Ana", "Lopez-Reyes"]);
eq("apostrophes survive", pick(RF.splitName("liam o'connor")), ["Liam", "O'Connor"]);
eq("SHOUTING is fixed", pick(RF.splitName("EMMA MCRAE")), ["Emma", "McRae"]);
eq("mixed case is left alone", pick(RF.splitName("DeShawn McRae")), ["DeShawn", "McRae"]);
eq("accents survive", RF.splitName("Jose Munoz̃").last.length > 0, true);
ok("one word flags", RF.splitName("Madonna").confidence === "check");

group("grade");
["Kindergarten", "Kinder", "K", "KG"].forEach(g => eq(g, RF.normGrade(g).value, "K"));
["1", "01", "1st", "Grade 1", "First Grade", "Gr 1"].forEach(g => eq(g, RF.normGrade(g).value, "1"));
eq("Pre-K", RF.normGrade("Pre-K").value, "PK");
eq("TK", RF.normGrade("Transitional Kindergarten").value, "TK");
eq("12", RF.normGrade("12th Grade").value, "12");
eq("2026 stays 2026", RF.normGrade("2026").value, "2026");
ok("2026 is flagged", RF.normGrade("2026").ok === false);
eq("unknown program kept verbatim", RF.normGrade("Room 12 Enrichment").value, "Room 12 Enrichment");
ok("unknown program flagged", RF.normGrade("Room 12 Enrichment").ok === false);

group("phone + email");
["(619) 555-1212", "619-555-1212", "619 555 1212", "6195551212", "+1 619 555 1212"]
  .forEach(p => eq(p, RF.cleanPhone(p).value, "(619) 555-1212"));
eq("extension kept", RF.cleanPhone("619-555-1212 x204").value, "(619) 555-1212 x204");
ok("7 digits flagged", RF.cleanPhone("555-1212").ok === false);
ok("text flagged", RF.cleanPhone("call mom").ok === false);
ok("fake flagged", RF.cleanPhone("111-111-1111").ok === false);
eq("case folded", RF.cleanEmail("MOM@GMAIL.COM").value, "mom@gmail.com");
eq("spaces round @ repaired", RF.cleanEmail("amy @ gmail.com").value, "amy@gmail.com");
ok("amy@gmail flagged", RF.cleanEmail("amy@gmail").ok === false);
ok("double dot flagged", RF.cleanEmail("amy@gmail..com").ok === false);
eq("two emails split", RF.splitEmails("amy@e.com; john@e.com").length, 2);
eq("comma split", RF.splitEmails("amy@e.com, john@e.com").length, 2);
eq("name with comma not split", RF.splitNames("Smith, John").length, 1);
eq("two phones split", RF.splitPhones("619-555-1212 / 619-555-3434").length, 2);

group("column mapping");
const maps = { "StuFirst": "student_first", "First Name": "student_first", "Given Name": "student_first",
  "StudentLastName": "student_last", "Surname": "student_last", "Family Name": "student_last",
  "Student Name": "student_full", "GradeLvl": "grade", "HR Teacher": "teacher", "Homeroom": "teacher",
  "Advisor": "teacher", "PermID": "student_id", "Student #": "student_id", "SIS ID": "student_id",
  "MomEmail": "parent_email", "Guardian 2 Cell": "parent_cell", "Mother Phone": "parent_cell",
  "Parent 1 Email": "parent_email", "Date of Birth": "ignore", "Lunch Status": "ignore" };
Object.keys(maps).forEach(h => eq(h, RF.detectColumn(h).field, maps[h]));
ok("DOB marked sensitive", RF.detectColumn("Date of Birth").sensitive === "date of birth");
ok("emergency contact is not a buyer by default", RF.detectColumn("Emergency Contact Email").questionable === true);

/* ------------------------------------------------ the whole pipeline */
group("household merge (the csv case)");
const csvBook = RF.readWorkbook(XLSX, householdCSV, "smallschool.csv");
const csv = RF.analyze(csvBook, {});
eq("4 rows become 3 students", csv.stats.students, 3);
const emma = csv.students.find(s => s.first === "Emma");
eq("Emma has two parent emails", emma.contacts.filter(c => c.email).length, 2);
eq("Emma keeps both emails", emma.contacts.map(c => c.email).sort(), ["jane@email.com", "john@email.com"]);
ok("Liam stays his own student", !!csv.students.find(s => s.first === "Liam"));
eq("shared parent is not duplicated onto Liam", csv.students.find(s => s.first === "Liam").contacts.length, 1);
eq("uppercase email folded", csv.students.find(s => s.first === "Ana").contacts[0].email, "rosa@email.com");

group("the messy workbook");
const wbBook = RF.readWorkbook(XLSX, messyWorkbook(), "Oakhaven Roster.xlsx");
const r = RF.analyze(wbBook, {});
eq("three tabs seen", r.sheets.length, 3);
eq("two look like rosters", r.sheets.filter(s => s.looksLikeRoster).length, 2);
ok("summary tab excluded", !r.sheets.find(s => s.name === "Summary").included);
eq("header row found on Grade 3", r.sheets[0].headerRow, 4);
eq("K tab implies a grade", r.sheets[1].implies.grade, "K");
ok("K tab offers the grade", r.sheets[1].offerGrade === true);

const by = n => r.students.find(s => (s.first + " " + s.last) === n);
eq("Emma Smith merged to one row", r.students.filter(s => s.first === "Emma" && s.last === "Smith").length, 1);
eq("Emma has mom and dad", by("Emma Smith").contacts.filter(c => c.email).length, 2);
eq("Emma's duplicate phone collapsed", by("Emma Smith").contacts.length, 2);
eq("leading zeros kept", by("Emma Smith").id, "0018472");
eq("exact duplicate removed", r.stats.exactDupes, 1);
eq("Noah appears once", r.students.filter(s => s.first === "Noah").length, 1);
eq("repeated header ignored", r.stats.repeatHeaders, 1);
ok("totals footer ignored", r.stats.junkRows >= 2);
eq("Sofia and Sophia stay separate", r.students.filter(s => s.last === "Martinez").length, 2);
ok("Sofia/Sophia flagged", r.reviews.some(x => x.type === "possible-duplicate" && x.cards[0].name.indexOf("Martinez") > 0));
eq("Isabella merged on her ID", r.students.filter(s => s.last === "Rodriguez").length, 1);
ok("Isabella's spelling flagged", r.reviews.some(x => x.type === "id-conflict"));
eq("Grade 3 normalized", by("Liam O'Connor").grade, "3");
eq("teacher flipped from LAST, FIRST", by("Liam O'Connor").teacher, "Susan Brown");
ok("teacher variants flagged", r.teacherVariants.length >= 1);
eq("2026 grade kept and flagged", by("Zoe Adams").grade, "2026");
ok("odd grade issue raised", r.issues.some(i => i.type === "odd-grade"));
ok("ava@gmail flagged", r.issues.some(i => i.type === "bad-email" && i.value === "ava@gmail"));
ok("555-1212 flagged", r.issues.some(i => i.type === "bad-phone"));
ok("missing teacher flagged", r.issues.some(i => i.type === "missing-teacher"));
ok("phone without email flagged", r.issues.some(i => i.type === "phone-no-email"));
eq("Oliver split from Last, First", by("Oliver Smith") ? "yes" : "no", "yes");
eq("two emails in one cell became two parents", by("Oliver Smith").contacts.filter(c => c.email).length, 2);
eq("van der Berg kept whole", by("Lucas van der Berg") ? by("Lucas van der Berg").last : "", "van der Berg");
eq("extension survived the pipeline", by("Lucas van der Berg").contacts[0].cell, "(619) 555-1818 x22");
eq("nbsp cleaned", by("Amara Okafor") ? "yes" : "no", "yes");
ok("emergency contact not exported", !r.students.some(s => s.contacts.some(c => (c.first + c.last).indexOf("Peg") >= 0)));
ok("DOB and lunch listed as unused", r.unused.filter(u => u.sensitive).length >= 2);
eq("K students got their grade from the tab", by("Oliver Smith").grade, "K");
ok("merged cells noted", r.stats.merges >= 1);
ok("hidden row still read", r.stats.hiddenRows >= 1);

group("decisions");
const merged = RF.analyze(wbBook, { decisions: (function () {
  const d = {};
  const dup = r.reviews.find(x => x.type === "possible-duplicate");
  d[dup.id] = "merge";
  return d;
})() });
eq("merging on request removes a student", merged.stats.students, r.stats.students - 1);
const split = RF.analyze(wbBook, { decisions: (function () {
  const d = {}; d[r.reviews.find(x => x.type === "id-conflict").id] = "split"; return d;
})() });
eq("splitting an ID conflict adds a student", split.stats.students, r.stats.students + 1);
const fixed = RF.analyze(wbBook, { fixes: (function () {
  const f = {}; f[r.issues.find(i => i.type === "bad-email").fixKey] = "ava@gmail.com"; return f;
})() });
ok("a hand-fixed email lands in the export", fixed.students.some(s => s.contacts.some(c => c.email === "ava@gmail.com")));
ok("and stops being an issue", fixed.issues.filter(i => i.type === "bad-email").length === r.issues.filter(i => i.type === "bad-email").length - 1);

group("remapping a column");
const idCol = r.sheets[0].cols.find(c => c.header === "DOB");
const remap = RF.analyze(wbBook, { mapping: (function () { const m = {}; m[idCol.mapKey] = "ignore"; return m; })() });
ok("mapping override accepted", remap.stats.students === r.stats.students);

group("the workbook that comes out");
RF.buildWorkbook(ExcelJS, r, { trace: true }).then(async buf => {
  const back = new ExcelJS.Workbook();
  await back.xlsx.load(buf);
  const names = back.worksheets.map(w => w.name);
  eq("sheet names", names, ["GOTPHOTO READY", "NEEDS REVIEW", "CLEANUP REPORT", "SOURCE TRACE"]);
  const ready = back.getWorksheet("GOTPHOTO READY");
  eq("header row", ready.getRow(1).values.slice(1, 6),
    ["Student First Name", "Student Last Name", "Grade", "Teacher", "Student ID"]);
  eq("one row per student", ready.rowCount - 1, r.stats.students);
  const ids = [];
  ready.eachRow((row, n) => { if (n > 1) ids.push(row.getCell(5).value); });
  ok("IDs are still text with leading zeros", ids.indexOf("0018472") >= 0, ids.slice(0, 3));
  ok("IDs are strings, not numbers", ids.every(v => v === null || typeof v === "string"));
  eq("output file name", RF.outputFileName("Oakhaven Roster.xlsx"), "Oakhaven Roster - CLEANED.xlsx");
  ok("review sheet has rows", back.getWorksheet("NEEDS REVIEW").rowCount > 1);
  ok("report sheet has rows", back.getWorksheet("CLEANUP REPORT").rowCount > 10);

  console.log("\n" + pass + " passed, " + fail + " failed");
  process.exit(fail ? 1 : 0);
});

function pick(n) { return [n.first, n.last]; }

/* appended: edge cases that used to invent duplicates out of blank names */
function edgeCases() {
  group("blank names");
  const csv = ["Student ID,Student Name,Grade,Parent Email,Parent Phone",
    "0001,Emma Smith,3,jane@e.com,",
    "0001,,3,john@e.com,",            // contacts-only second household row
    "0002,,3,rita@e.com,",            // a row with no name at all
    "0003,,3,carl@e.com,",            // ...and another
    "0004,Ana Lee,3,ana@e.com,"].join("\n");
  const res = RF.analyze(RF.readWorkbook(XLSX, csv, "blank.csv"), {});
  const emma = res.students.find(s => s.first === "Emma");
  eq("nameless row merges into its ID", emma.contacts.filter(c => c.email).length, 2);
  eq("no phantom ID conflict", res.reviews.filter(r => r.type === "id-conflict").length, 0);
  eq("two nameless students are not duplicates of each other",
    res.reviews.filter(r => r.type === "possible-duplicate").length, 0);
  eq("they are flagged instead", res.issues.filter(i => i.type === "no-name").length, 2);
}
edgeCases();
