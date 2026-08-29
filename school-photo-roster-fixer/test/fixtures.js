/* Builds the deliberately awful workbooks the tests run against.
   Every ugly thing in here is something a school has actually sent. */
const XLSX = require("xlsx");
const NBSP = String.fromCharCode(0xa0), ZWSP = String.fromCharCode(0x200b);
const RSQUO = String.fromCharCode(0x2019);

function sheetFromRows(rows) {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  return ws;
}

/* A four-tab elementary export: title junk, camel-case headers, split
   households, an exact duplicate, IDs with leading zeros, a totals footer. */
function messyWorkbook() {
  const wb = XLSX.utils.book_new();

  const g3 = [
    ["Oakhaven Elementary School"],
    ["Student Roster - Picture Day"],
    ["Generated 09/12/2026"],
    [],
    ["StuFirst", "StuLast", "GradeLvl", "HR Teacher", "PermID", "MomEmail", "MomCell", "DadEmail", "DadCell", "DOB", "Lunch Status"],
    ["Emma", "Smith", "3", "Mrs Brown", "0018472", "jane@email.com", "(619) 555-1234", "", "", "2017-04-02", "Free"],
    ["Emma", "Smith", "3", "Mrs. Brown", "0018472", "", "619-555-1234", "john@email.com", "6195559999", "2017-04-02", "Free"],
    ["Liam", "O" + RSQUO + "CONNOR", "03", "BROWN, SUSAN", "0018473", "  MOM@GMAIL.COM ", "619.555.0001", "", "", "2017-01-09", "Paid"],
    ["Sofia", "Martinez", "3", "Mrs. Brown", "0018474", "sofia.m@email.com", "", "", "", "2017-06-11", "Paid"],
    ["Sophia", "Martinez", "3", "Mrs. Brown", "0018999", "sophia.m@email.com", "", "", "", "2017-06-11", "Paid"],
    [],
    ["Noah", "Kim", "3", "Mr. Jones", "0018475", "kim.family@email.com", "619 555 7777", "", "", "2017-02-02", "Free"],
    ["Noah", "Kim", "3", "Mr. Jones", "0018475", "kim.family@email.com", "619 555 7777", "", "", "2017-02-02", "Free"],
    ["StuFirst", "StuLast", "GradeLvl", "HR Teacher", "PermID", "MomEmail", "MomCell", "DadEmail", "DadCell", "DOB", "Lunch Status"],
    ["Ava" + ZWSP, "Nguyen ", "3", "Mr. Jones", "0018476", "ava@gmail", "555-1212", "", "", "2017-08-30", "Paid"],
    ["Mateo", "Garcia", "3", "", "0018477", "", "(619) 555-4444", "", "", "2017-03-14", "Paid"],
    ["Isabella", "Rodriguez", "3", "Mr. Jones", "0018478", "iz@email.com", "", "", "", "2017-05-05", "Paid"],
    ["Isabela", "Rodriguez", "3", "Mr. Jones", "0018478", "", "619-555-8888", "", "", "2017-05-05", "Paid"],
    ["Zoe", "Adams", "2026", "Mr. Jones", "0018479", "zoe.mom@email.com", "6195552222", "", "", "2017-07-07", "Paid"],
    ["Total Students: 12"],
    ["Page 1 of 1"]
  ];
  const wsG3 = sheetFromRows(g3);
  /* Student IDs must survive as text, so force the cell type. */
  for (let r = 5; r <= 19; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 4 });
    if (wsG3[addr] && wsG3[addr].v) { wsG3[addr].t = "s"; wsG3[addr].v = String(wsG3[addr].v); }
  }
  wsG3["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  wsG3["!rows"] = []; wsG3["!rows"][12] = { hidden: true };
  XLSX.utils.book_append_sheet(wb, wsG3, "Grade 3");

  /* Kindergarten tab: one full-name column, Last-First order, two emails in one
     cell, an emergency contact, and no grade column at all. */
  const k = [
    ["Kindergarten - Mrs. Alvarez"],
    [],
    ["Student Name", "Teacher", "Parent Name", "Parent Email", "Parent Phone", "Emergency Contact", "Emergency Phone"],
    ["Smith, Oliver", "Mrs. Alvarez", "Dana Smith", "dana@email.com; greg@email.com", "619-555-3131", "Aunt Peg", "619-555-9090"],
    ["Mary Kate Johnson", "Mrs. Alvarez", "Rob Johnson", "rob@email.com", "(619)555-1717", "", ""],
    ["Lucas van der Berg", "Mrs. Alvarez", "Ana van der Berg", "ana@email.com", "619 555 1818 x22", "", ""],
    ["Chen, Wei Ming", "Mrs. Alvarez", "Li Chen", "li@email..com", "not provided", "", ""],
    ["Amara" + NBSP + "Okafor", "Mrs. Alvarez", "", "", "", "", ""],
    ["Jack Miller Jr.", "Mrs. Alvarez", "Sara Miller", "sara@email.com", "619-555-2323", "", ""]
  ];
  XLSX.utils.book_append_sheet(wb, sheetFromRows(k), "K");

  /* A report tab with no roster on it at all. */
  XLSX.utils.book_append_sheet(wb, sheetFromRows([
    ["Enrollment Summary"], [], ["Grade", "Count"], ["K", 6], ["3", 12], ["Total", 18]
  ]), "Summary");

  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

/* The spec's headline case: one row per parent, nothing else to go on. */
const householdCSV = [
  "Student,Parent,Email,Phone",
  "Emma Smith,Jane Smith,jane@email.com,(619) 555-1234",
  "Emma Smith,John Smith,john@email.com,619-555-9999",
  "Liam Smith,Jane Smith,jane@email.com,(619) 555-1234",
  "Ana Torres,Rosa Torres,ROSA@EMAIL.COM,6195550000"
].join("\n");

module.exports = { messyWorkbook, householdCSV };
