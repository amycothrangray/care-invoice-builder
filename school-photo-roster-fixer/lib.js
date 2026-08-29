/* School Photo Roster Fixer -- cleanup engine.

   Pure logic: no DOM, no globals, no file reading. The browser app (app.js)
   feeds it parsed sheets plus the user's choices and re-runs it after every
   change; the Node test harness feeds it the same shapes. SheetJS / ExcelJS are
   handed in as arguments so this file never depends on a script tag.

   The rule that outranks everything else here: never silently combine two rows
   that might be two different children. Formatting problems get fixed
   automatically; identity problems get flagged for a human. */

(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  else root.RF = api;
})(typeof self !== "undefined" ? self : this, function () {
"use strict";

/* ============================================================ text scrubbing

   Everything a school spreadsheet smuggles in -- non-breaking spaces, smart
   quotes, zero-width joiners, alt+enter line breaks -- is handled in one pass
   over the characters, so the table below is the whole list of what we touch. */

const DROP_CHARS = new Set([0x200b, 0x200c, 0x200d, 0x200e, 0x200f, 0x2060, 0x2061,
  0x2062, 0x2063, 0x202a, 0x202b, 0x202c, 0x202d, 0x202e, 0xfeff, 0x00ad, 0x180e]);
const SPACE_CHARS = new Set([0x00a0, 0x1680, 0x2000, 0x2001, 0x2002, 0x2003, 0x2004,
  0x2005, 0x2006, 0x2007, 0x2008, 0x2009, 0x200a, 0x202f, 0x205f, 0x3000]);
const SWAP_CHARS = {
  0x2018: "'", 0x2019: "'", 0x201a: "'", 0x201b: "'", 0x02bc: "'", 0x02b9: "'", 0x2032: "'",
  0x201c: '"', 0x201d: '"', 0x201e: '"', 0x2033: '"',
  0x2010: "-", 0x2011: "-", 0x2012: "-", 0x2013: "-", 0x2014: "-", 0x2015: "-", 0x2212: "-",
  0x2026: "...", 0xfb01: "fi", 0xfb02: "fl"
};
const COMBINING = new RegExp("[\\u0300-\\u036f]", "g");
const LETTERS = new RegExp("[A-Za-z\\u00c0-\\u024f]+", "g");

/* Flatten one cell to clean text. `keepBreaks` leaves line breaks in place so a
   cell holding two contacts stacked with alt+enter can still be split apart.
   Returns the string; `scrubStats` reports what had to be repaired. */
function scrubWith(v, keepBreaks, stats) {
  if (v == null) return "";
  const s = String(v);
  let out = "";
  for (const ch of s) {
    const c = ch.codePointAt(0);
    if (DROP_CHARS.has(c)) { if (stats) stats.odd++; continue; }
    if (SPACE_CHARS.has(c)) { if (stats) stats.odd++; out += " "; continue; }
    if (SWAP_CHARS[c] !== undefined) { if (stats) stats.odd++; out += SWAP_CHARS[c]; continue; }
    if (c === 10 || c === 13) { if (stats) stats.breaks++; out += keepBreaks ? "\n" : " "; continue; }
    if (c === 9) { out += " "; continue; }
    if (c < 32 || c === 127) { if (stats) stats.odd++; out += " "; continue; }
    out += ch;
  }
  out = keepBreaks
    ? out.replace(/[ \t]+/g, " ").replace(/ ?\n ?/g, "\n").replace(/\n{2,}/g, "\n")
    : out.replace(/\s+/g, " ");
  return out.trim();
}
const scrub = (v, stats) => scrubWith(v, false, stats);
const scrubMulti = (v, stats) => scrubWith(v, true, stats);

const deaccent = s => String(s == null ? "" : s).normalize("NFD").replace(COMBINING, "");
/* Comparison keys: accent-free, punctuation-free, lower case. Never exported. */
const key = s => deaccent(scrub(s)).toLowerCase().replace(/[^a-z0-9]+/g, "");
const looseKey = s => deaccent(scrub(s)).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

const titleWord = w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();

/* Case-fix a person's name -- but only the pieces that are shouting or
   whispering. Mixed-case input is left exactly as the school typed it, so
   DeShawn, McRae and van Dijk all survive untouched. */
function fixCase(raw) {
  const s = scrub(raw);
  if (!s) return "";
  return s.split(" ").map(word => {
    const bare = deaccent(word);
    const isUpper = word === word.toUpperCase() && /[A-Z]/.test(bare);
    const isLower = word === word.toLowerCase() && /[a-z]/.test(bare);
    if (!isUpper && !isLower) return word;                     // already mixed -- hands off
    if (/^[A-Za-z]\.?$/.test(bare)) return word.toUpperCase(); // single initial
    let out = word.replace(LETTERS, m => titleWord(m));
    if (isUpper) {
      out = out.replace(/^(Mc)([a-z])/, (m, a, b) => a + b.toUpperCase());
      out = out.replace(/^(Mac)([a-z]{2,})/, (m, a, b) => a + b.charAt(0).toUpperCase() + b.slice(1));
      out = out.replace(/^(O')([a-z])/, (m, a, b) => a + b.toUpperCase());
      out = out.replace(/(-)([a-z])/g, (m, a, b) => a + b.toUpperCase());
    }
    return out;
  }).join(" ");
}

/* =========================================================== student names */

const SUFFIXES = { jr: "Jr.", sr: "Sr.", ii: "II", iii: "III", iv: "IV", "2nd": "II", "3rd": "III" };
const PARTICLES = new Set(["van", "von", "de", "del", "dela", "della", "der", "den", "da", "das",
  "dos", "di", "du", "la", "le", "les", "st", "ste", "bin", "ibn", "al", "el", "ter", "ten",
  "op", "vander", "vanden", "abu"]);
const NOT_A_NAME = /^(n\/?a|none|unknown|tbd|test|no name|\.|-|--)$/i;
const suffixOf = w => SUFFIXES[String(w == null ? "" : w).toLowerCase().replace(/\./g, "")] || null;

/* Split one full-name string into parts.
   confidence: "sure" (nothing to guess) / "likely" (a defensible rule fired) /
   "check" (a human should glance at it). Never invents or drops characters. */
function splitName(raw) {
  const s = scrub(raw).replace(/^[,;\s]+|[,;\s]+$/g, "");
  const out = { first: "", last: "", middle: "", suffix: "", confidence: "sure", note: "", raw: scrub(raw) };
  if (!s || NOT_A_NAME.test(s)) { out.confidence = "check"; out.note = s ? "not a name" : "blank"; return out; }

  /* "Smith, Emma" / "Smith, Emma Rose" / "Smith, Emma, Jr." */
  if (s.indexOf(",") >= 0) {
    const parts = s.split(",").map(p => p.trim()).filter(Boolean);
    if (parts.length >= 2) {
      let suffix = "";
      if (parts.length >= 3 && suffixOf(parts[2])) suffix = suffixOf(parts[2]);
      const rest = parts[1].split(" ").filter(Boolean);
      if (rest.length > 1 && suffixOf(rest[rest.length - 1])) suffix = suffixOf(rest.pop());
      out.first = fixCase(rest.shift() || "");
      out.middle = fixCase(rest.join(" "));
      out.last = fixCase(parts[0]);
      out.suffix = suffix;
      out.confidence = "likely";
      out.note = "read as Last, First";
      if (!out.first) { out.confidence = "check"; out.note = "only one name either side of the comma"; }
      return out;
    }
  }

  const tok = s.split(" ").filter(Boolean);
  if (tok.length > 1 && suffixOf(tok[tok.length - 1])) out.suffix = suffixOf(tok.pop());

  if (tok.length === 0) { out.confidence = "check"; out.note = "blank"; return out; }
  if (tok.length === 1) {
    out.first = fixCase(tok[0]);
    out.confidence = "check";
    out.note = "only one word - no last name";
    return out;
  }
  if (tok.length === 2) { out.first = fixCase(tok[0]); out.last = fixCase(tok[1]); return out; }

  /* Surname particles glue the tail together: "Emma van der Berg" */
  const pIdx = tok.findIndex((t, i) =>
    i > 0 && i < tok.length - 1 && PARTICLES.has(deaccent(t).toLowerCase().replace(/\./g, "")));
  if (pIdx > 0) {
    out.first = fixCase(tok[0]);
    out.middle = fixCase(tok.slice(1, pIdx).join(" "));
    /* "van der Berg" is how that family writes it -- leave their particles alone. */
    out.last = tok.slice(pIdx).map(t =>
      (PARTICLES.has(t.toLowerCase()) && t === t.toLowerCase()) ? t : fixCase(t)).join(" ");
    out.confidence = "likely";
    out.note = "multi-word last name";
    return out;
  }

  out.first = fixCase(tok[0]);
  out.middle = fixCase(tok.slice(1, -1).join(" "));
  out.last = fixCase(tok[tok.length - 1]);
  if (tok.slice(1, -1).every(m => /^[A-Za-z]\.?$/.test(deaccent(m)))) {
    out.confidence = "likely";
    out.note = "middle initial";
  } else {
    /* "Mary Kate Johnson" -- Johnson is the surname either way, but the first
       name could be "Mary" or "Mary Kate". Offer both rather than guessing. */
    out.confidence = "check";
    out.altFirst = fixCase(tok.slice(0, 2).join(" "));
    out.altMiddle = fixCase(tok.slice(2, -1).join(" "));
    out.note = 'first name might be "' + out.altFirst + '"';
  }
  return out;
}

/* ================================================================== grade */

const WORD_NUM = { one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8,
  nine: 9, ten: 10, eleven: 11, twelve: 12, first: 1, second: 2, third: 3, fourth: 4, fifth: 5,
  sixth: 6, seventh: 7, eighth: 8, ninth: 9, tenth: 10, eleventh: 11, twelfth: 12 };

/* Unknown programs keep their original text and get flagged -- they never
   become an invented grade number. */
function normGrade(raw) {
  const s = scrub(raw);
  if (!s) return { value: "", changed: false, ok: true, note: "" };
  const k = looseKey(s).replace(/\b(grade|gr|grd|level|lvl|year|yr)\b/g, " ").replace(/\s+/g, " ").trim();
  const bare = k.replace(/\s/g, "");
  const same = v => ({ value: v, changed: s !== v, ok: true, note: "" });

  if (/^(prek|pk|prekindergarten|prekinder|pre)$/.test(bare)) return same("PK");
  if (/^(preschool|ps|nursery)$/.test(bare)) return same("PS");
  if (/^(tk|transitionalkindergarten|transk|transitionalk)$/.test(bare)) return same("TK");
  if (/^(k|kg|kn|kd|kinder|kindergarten|kindergarden|k5|fulldayk|halfdayk)$/.test(bare)) return same("K");

  const m = bare.match(/^0*(\d{1,2})(st|nd|rd|th)?$/);
  if (m) {
    const n = parseInt(m[1], 10);
    if (n === 0) return { value: "K", changed: true, ok: true, note: "grade 0 read as Kindergarten" };
    if (n >= 1 && n <= 12) return same(String(n));
    return { value: s, changed: false, ok: false, note: "not a grade level" };
  }
  if (WORD_NUM[bare]) return { value: String(WORD_NUM[bare]), changed: true, ok: true, note: "" };
  return { value: s, changed: false, ok: false, note: "unrecognized grade" };
}

const GRADE_ORDER = ["PS", "PK", "TK", "K", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
const ELEMENTARY = new Set(["PS", "PK", "TK", "K", "1", "2", "3", "4", "5", "6"]);
const gradeRank = g => { const i = GRADE_ORDER.indexOf(g); return i < 0 ? 99 : i; };

/* ================================================================ teacher */

const TITLES = { mr: "Mr.", mrs: "Mrs.", ms: "Ms.", miss: "Miss", mx: "Mx.", dr: "Dr.",
  madame: "Madame", mme: "Mme.", sra: "Sra.", srta: "Srta.", profe: "Profe." };

/* Tidy spacing and case, and put the period back on Mrs -- nothing more. A
   teacher's actual name is never rewritten. */
function normTeacher(raw) {
  const s = scrub(raw);
  if (!s) return { value: "", changed: false, note: "" };
  if (NOT_A_NAME.test(s)) return { value: "", changed: true, note: "not a teacher name" };
  let flipped = false, text = s;

  /* "SMITH, JENNIFER" names the same classroom as "Jennifer Smith". Flip it so
     the two collapse instead of splitting one class in half. */
  const c = text.split(",").map(p => p.trim()).filter(Boolean);
  if (c.length === 2 && !TITLES[looseKey(c[0])] && c[0].split(" ").length <= 3 && c[1].split(" ").length <= 2) {
    text = c[1] + " " + c[0];
    flipped = true;
  }

  const value = text.split(" ").filter(Boolean).map((w, i) => {
    const t = TITLES[deaccent(w).toLowerCase().replace(/\./g, "")];
    return (t && i === 0) ? t : fixCase(w);
  }).join(" ");
  return { value, changed: value !== s, note: flipped ? "read as Last, First" : "" };
}

/* Surname-only key, for spotting "Mrs Smith" / "Mrs. Smith" / "Jennifer Smith"
   variants of one teacher. Used to flag them, never to merge them. */
function teacherKey(t) {
  const words = looseKey(t).split(" ").filter(w => w && TITLES[w] === undefined);
  return words.length ? words[words.length - 1] : "";
}

/* ============================================================= student ID */

/* IDs are identifiers, not numbers. `numeric` is the raw cell value when Excel
   stored it as a number -- that is where leading zeros get lost upstream, so we
   say so rather than pretending we can put them back. */
function normId(text, numeric) {
  const s = scrub(text);
  if (numeric != null && Number.isFinite(numeric)) {
    const asInt = Number.isInteger(numeric) ? numeric.toFixed(0) : String(numeric);
    const scientific = /e[+-]/i.test(s);
    return { value: (scientific || !s) ? asInt : s, wasNumeric: true,
      note: scientific ? "recovered from scientific notation" : "" };
  }
  return { value: s.replace(/\.0+$/, ""), wasNumeric: false, note: "" };
}

/* ================================================================= emails */

const EMAIL_OK = /^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+$/;
const EMAIL_DEAD = /^(none|n\/?a|no ?e-?mail|unknown|not provided|nomail|-|\.)$/i;

/* Repairs only what is unambiguous (stray spaces, mailto:, trailing
   punctuation) and never guesses at a domain. */
function cleanEmail(raw) {
  const before = scrub(raw);
  if (!before) return { value: "", ok: true, changed: false, reason: "" };
  if (EMAIL_DEAD.test(before))
    return { value: "", ok: true, changed: true, reason: "placeholder", dropped: before };
  let s = before.replace(/^mailto:/i, "").replace(/^[<("']+|[>)"'.,;:]+$/g, "");
  s = s.replace(/\s*@\s*/g, "@").replace(/\s*\.\s*/g, ".").replace(/\s+/g, "");
  const value = s.toLowerCase();
  if (!EMAIL_OK.test(value)) {
    let reason = "does not look like an email address";
    if (value.indexOf("@") < 0) reason = "no @ sign";
    else if (value.split("@").length > 2) reason = "more than one @ sign";
    else if (value.split("@")[1].indexOf(".") < 0) reason = "no domain ending (.com, .org...)";
    else if (/\.\./.test(value)) reason = "double dot";
    else if (/@\./.test(value) || /\.$/.test(value)) reason = "misplaced dot";
    return { value, ok: false, changed: value !== before, reason };
  }
  return { value, ok: true, changed: value !== before, reason: "" };
}

/* ================================================================= phones */

const FAKE_PHONE = /^(\d)\1{9}$|^0123456789$|^1234567890$/;

/* US/Canada normalisation for comparison and display. Extensions survive.
   Missing area codes are never invented. */
function cleanPhone(raw) {
  const before = scrub(raw);
  if (!before) return { value: "", key: "", ok: true, changed: false, reason: "" };
  if (/^(none|n\/?a|no ?phone|unknown|not provided|-|\.)$/i.test(before))
    return { value: "", key: "", ok: true, changed: true, reason: "placeholder", dropped: before };

  let s = before, ext = "";
  const em = s.match(/(?:x|ext|extn|extension)\.?\s*[:#]?\s*(\d{1,6})\s*$/i);
  if (em) { ext = em[1]; s = s.slice(0, em.index); }

  const letters = s.replace(/[^A-Za-z]/g, "");
  let digits = s.replace(/\D/g, "");
  if (digits.length === 11 && digits.charAt(0) === "1") digits = digits.slice(1);
  const bad = reason => ({ value: "", key: "", ok: false, changed: true, reason, dropped: before });

  if (!digits) return bad("no digits");
  if (letters.length > 2 && digits.length < 10) return bad("text, not a phone number");
  if (digits.length === 7) return bad("7 digits - no area code");
  if (digits.length < 10) return bad("only " + digits.length + " digits");
  if (digits.length > 10) return bad(digits.length + " digits - too long");
  if (FAKE_PHONE.test(digits) || /^[01]/.test(digits) || /^\d{3}[01]/.test(digits))
    return bad("not a valid US number");

  const value = "(" + digits.slice(0, 3) + ") " + digits.slice(3, 6) + "-" + digits.slice(6) +
    (ext ? " x" + ext : "");
  return { value, key: digits + (ext ? "x" + ext : ""), ok: true, changed: value !== before, reason: "" };
}

/* ======================================================== many in one cell */

const PHONE_SCAN = /(?:\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}(?:\s*(?:x|ext|extension)\.?\s*\d{1,6})?/gi;

/* "amy@a.com; john@b.com" gives two emails. A comma is only trusted when what
   follows it looks like an email -- a comma inside a name must not split it. */
function splitEmails(raw) {
  const s = scrubMulti(raw);
  if (!s) return [];
  if ((s.match(/@/g) || []).length <= 1) return [s];
  const parts = s.split(/[;\n]|\s+and\s+|\s{2,}|,\s*(?=[^\s@,]+@)|\/(?=[^\s@/]+@)/i)
    .map(p => scrub(p)).filter(Boolean);
  const good = parts.filter(p => p.indexOf("@") >= 0);
  return good.length >= 2 ? good : [s];
}

/* Same idea for phones: only split when each piece parses as a real number. */
function splitPhones(raw) {
  const s = scrubMulti(raw);
  if (!s) return [];
  const found = s.match(PHONE_SCAN);
  if (!found || found.length < 2) return [s];
  const rest = s.replace(PHONE_SCAN, " ").replace(/[\s,;/&|()\-.]+/g, " ").trim();
  /* Short labels between numbers ("mom", "cell") are fine; a long word stuck to
     a number could mean something else, so leave that cell alone. */
  if (rest.split(" ").some(w => w.length > 12)) return [s];
  return found.map(f => scrub(f));
}

/* Names never split on a comma ("Smith, John" is one person). */
function splitNames(raw) {
  const s = scrubMulti(raw);
  if (!s) return [];
  const parts = s.split(/[;\n]|\s+&\s+|\s+and\s+/i).map(p => scrub(p)).filter(Boolean);
  return parts.length >= 2 ? parts : [s];
}

/* ====================================================== reading the headers

   School exports never use our words. Everything below turns whatever the
   office typed into one of our fields, with a confidence we can show the user
   on the mapping screen so they can overrule it. */

const ABBR = {
  stu: "student", stud: "student", studnt: "student", st: "student", chld: "child",
  fname: "first name", firstnm: "first name", lname: "last name", lastnm: "last name",
  mname: "middle name", nm: "name", nme: "name",
  grd: "grade", gr: "grade", gradelvl: "grade level", lvl: "level", lev: "level",
  tchr: "teacher", teach: "teacher", tch: "teacher", hr: "homeroom", hmrm: "homeroom",
  rm: "room", sect: "section", sec: "section",
  prnt: "parent", par: "parent", pnt: "parent", grdn: "guardian", guard: "guardian",
  gaurdian: "guardian", mthr: "mother", fthr: "father",
  eml: "email", emal: "email", mail: "email", ph: "phone", phn: "phone", tel: "phone",
  telephone: "phone", cel: "cell", cll: "cell", mob: "mobile", mbl: "mobile",
  num: "number", nbr: "number", no: "number", numb: "number",
  addr: "address", dob: "date of birth", bday: "date of birth", perm: "permanent",
  sid: "student id", ident: "id", identifier: "id"
};

/* "StuFirstName" and "stu_first_nm" both land on "student first name". */
function headerNorm(raw) {
  let s = scrub(raw);
  s = s.replace(/([a-z0-9])([A-Z])/g, "$1 $2").replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2");
  s = s.replace(/([A-Za-z])(\d)/g, "$1 $2").replace(/#/g, " number ");
  s = deaccent(s).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  if (!s) return "";
  return s.split(" ").map(w => ABBR[w] || w).join(" ").replace(/\s+/g, " ").trim();
}

const F = {
  IGNORE: "ignore", FIRST: "student_first", LAST: "student_last", MIDDLE: "student_middle",
  FULL: "student_full", SUFFIX: "student_suffix", GRADE: "grade", TEACHER: "teacher",
  ID: "student_id", P_FIRST: "parent_first", P_LAST: "parent_last", P_FULL: "parent_full",
  P_EMAIL: "parent_email", P_CELL: "parent_cell"
};
const FIELD_LABEL = {
  ignore: "Not exported", student_first: "Student First Name", student_last: "Student Last Name",
  student_middle: "Student Middle Name", student_full: "Student Full Name (split it)",
  student_suffix: "Student Suffix", grade: "Grade", teacher: "Teacher", student_id: "Student ID",
  parent_first: "Parent First Name", parent_last: "Parent Last Name",
  parent_full: "Parent Full Name (split it)", parent_email: "Parent Email", parent_cell: "Parent Cell"
};
const PARENT_FIELDS = new Set([F.P_FIRST, F.P_LAST, F.P_FULL, F.P_EMAIL, F.P_CELL]);
const fieldLabel = (field, slot) =>
  (FIELD_LABEL[field] || field) + (PARENT_FIELDS.has(field) && slot ? " " + slot : "");

/* Contact roles. `prefer` is the parent slot the role usually deserves, which is
   why MomEmail lands in Parent Email 1 and DadEmail in Parent Email 2. */
const ROLES = [
  { re: /\b(step ?mother|step ?mom|step ?father|step ?dad|step ?parent)\b/, kind: "step" },
  { re: /\b(grand ?(mother|father|parent|ma|pa)|grandma|grandpa)\b/, kind: "grandparent" },
  { re: /\b(mother|mom|mommy|maternal)\b/, kind: "mother", prefer: 1 },
  { re: /\b(father|dad|daddy|paternal)\b/, kind: "father", prefer: 2 },
  { re: /\bprimary\b/, kind: "primary", prefer: 1 },
  { re: /\bsecondary\b/, kind: "secondary", prefer: 2 },
  { re: /\bemergency\b/, kind: "emergency", questionable: true },
  { re: /\b(parent|guardian|caregiver|custodian|carer)\b/, kind: "parent" },
  { re: /\b(household|buyer|contact|adult|family)\b/, kind: "contact" }
];
const SENSITIVE = [
  { re: /\b(date of birth|birth ?date|birthday|born)\b/, what: "date of birth" },
  { re: /\b(address|street|apt|apartment|city|state|zip|postal)\b/, what: "home address" },
  { re: /\b(medical|medication|allerg|health|nurse|immuni|vaccin)\b/, what: "medical" },
  { re: /\b(lunch|free reduced|frl|meal)\b/, what: "lunch status" },
  { re: /\b(iep|504|sped|special ed|ell|esl|gifted|gate)\b/, what: "program / IEP" },
  { re: /\b(gender|sex|pronoun)\b/, what: "gender" },
  { re: /\b(race|ethnic|hispanic|nationality)\b/, what: "race / ethnicity" },
  { re: /\b(discipline|suspension|referral|behavior)\b/, what: "discipline" },
  { re: /\b(password|login|username|pin|ssn|social security)\b/, what: "login / password" },
  { re: /\b(note|notes|comment|remark|internal)\b/, what: "internal notes" },
  { re: /\b(bus|transport|route)\b/, what: "transportation" }
];
const PHOTO_REL = /\b(photo|picture|media|publicity|directory|yearbook)\b.*\b(release|permission|consent|opt|allow|flag|ok)\b|\b(do not|no)\b.*\b(photo|picture)\b|\bopt ?out\b/;

/* One source column to one of our fields. Returns the field, which contact
   group it belongs to, a confidence, and anything worth telling the user. */
function detectColumn(raw) {
  const h = headerNorm(raw);
  const out = { field: F.IGNORE, group: "", roleKind: "", prefer: null, conf: "sure",
    note: "", questionable: false, sensitive: "", phoneKind: "" };
  if (!h) { out.note = "unnamed column"; return out; }

  if (PHOTO_REL.test(h)) { out.sensitive = "photo / media release"; out.note = "photo release - not a GotPhoto field, but you may want it"; return out; }
  for (const s of SENSITIVE) if (s.re.test(h)) { out.sensitive = s.what; return out; }

  /* "Family Name" is a surname, not a family contact. */
  if (/\b(family|sur) name\b/.test(h) && !/\b(parent|guardian|contact)\b/.test(h)) {
    out.field = F.LAST; return out;
  }
  let role = null;
  for (const r of ROLES) if (r.re.test(h)) { role = r; break; }
  const numMatch = h.match(/\b([1-9])\b/);
  const num = numMatch ? parseInt(numMatch[1], 10) : null;
  const isStudent = /\b(student|child|pupil|scholar|kid)\b/.test(h);

  const wantsEmail = /\be-? ?mail\b/.test(h);
  const wantsPhone = /\b(phone|cell|mobile|fone)\b/.test(h);
  const wantsFirst = /\b(first|given|fore)\b/.test(h);
  const wantsLast = /\b(last|surname|family name|sur)\b/.test(h);
  const wantsMiddle = /\b(middle|mi)\b/.test(h);
  const wantsName = /\b(name|names)\b/.test(h);

  if (role) {
    out.roleKind = role.kind;
    out.group = role.kind + (num || "");
    out.prefer = num || role.prefer || null;
    out.questionable = !!role.questionable;
    if (role.questionable) out.note = "emergency contact - not exported unless you say it is a parent";
    if (wantsEmail) out.field = F.P_EMAIL;
    else if (wantsPhone) {
      out.field = F.P_CELL;
      out.phoneKind = /\b(cell|mobile)\b/.test(h) ? "cell" : /\bhome\b/.test(h) ? "home"
        : /\b(work|business|office)\b/.test(h) ? "work" : "other";
      if (out.phoneKind === "home" || out.phoneKind === "work")
        out.note = out.phoneKind + " phone - used only if there is no cell";
    }
    else if (wantsFirst) out.field = F.P_FIRST;
    else if (wantsLast) out.field = F.P_LAST;
    else if (wantsName) out.field = F.P_FULL;
    else { out.field = F.IGNORE; out.note = out.note || "no contact value in this column"; }
    if (role.kind === "grandparent" || role.kind === "step" || role.kind === "contact") {
      out.conf = "check";
      if (!out.note) out.note = role.kind + " contact - check before exporting";
    }
    if (out.field !== F.IGNORE && role.questionable) out.conf = "check";
    return out;
  }

  /* No role word. Student fields, or a bare Email/Phone that is almost always
     the guardian's -- flagged so the user confirms it on the mapping screen. */
  if (wantsEmail) {
    if (isStudent) { out.note = "student's own email - not a buyer"; return out; }
    out.field = F.P_EMAIL; out.group = "contact"; out.conf = "check";
    out.note = "no parent/guardian word in the header - confirm this is the buyer";
    return out;
  }
  if (wantsPhone) {
    if (isStudent) { out.note = "student's own phone - not a buyer"; return out; }
    out.field = F.P_CELL; out.group = "contact"; out.conf = "check";
    out.phoneKind = /\b(cell|mobile)\b/.test(h) ? "cell" : /\bhome\b/.test(h) ? "home" : "other";
    out.note = "no parent/guardian word in the header - confirm this is the buyer";
    return out;
  }

  if (/\bclass of\b/.test(h) || /\bgraduat/.test(h)) { out.note = "graduation year"; return out; }
  if (/\b(grade|level)\b/.test(h) && !/\b(reading|math|score|gpa|average|mark)\b/.test(h)) {
    out.field = F.GRADE; return out;
  }
  if (/\b(teacher|homeroom|advisor|advisory|instructor)\b/.test(h)) { out.field = F.TEACHER; return out; }
  if (/\b(classroom|class|section|room)\b/.test(h)) {
    out.field = F.TEACHER; out.conf = "check";
    out.note = "reading this as the teacher/classroom";
    return out;
  }
  if (/\bid\b/.test(h) || /\b(sis|powerschool|aeries|skyward|infinite campus|permanent|local|state)\b.*\b(id|number)\b/.test(h)
      || (isStudent && /\bnumber\b/.test(h))) {
    out.field = F.ID; return out;
  }
  if (wantsFirst) { out.field = F.FIRST; return out; }
  if (wantsLast) { out.field = F.LAST; return out; }
  if (wantsMiddle) { out.field = F.MIDDLE; return out; }
  if (/\bsuffix\b/.test(h)) { out.field = F.SUFFIX; return out; }
  if (wantsName || isStudent) {
    out.field = F.FULL;
    if (!isStudent && !/\b(full|student|legal|display)\b/.test(h)) {
      out.conf = "check";
      out.note = 'a bare "' + scrub(raw) + '" column - reading it as the student name';
    }
    return out;
  }
  if (/\bnumber\b/.test(h)) { out.field = F.ID; out.conf = "check"; out.note = "reading this as the student ID"; return out; }
  return out;
}

/* Parent slots. Explicit numbers win, then the usual roles (mother 1, father 2),
   then everything else fills the first free slot in column order. */
function assignSlots(columns) {
  const groups = [];
  const byName = new Map();
  columns.forEach(col => {
    if (!PARENT_FIELDS.has(col.field) || !col.group) return;
    if (!byName.has(col.group)) {
      const g = { name: col.group, prefer: col.prefer, slot: null, questionable: col.questionable };
      byName.set(col.group, g);
      groups.push(g);
    } else if (col.prefer && !byName.get(col.group).prefer) byName.get(col.group).prefer = col.prefer;
  });
  const taken = new Set();
  const explicit = groups.filter(g => /\d$/.test(g.name) && g.prefer);
  explicit.forEach(g => { if (!taken.has(g.prefer)) { g.slot = g.prefer; taken.add(g.slot); } });
  groups.filter(g => !g.slot && g.prefer).forEach(g => {
    if (!taken.has(g.prefer)) { g.slot = g.prefer; taken.add(g.slot); }
  });
  groups.filter(g => !g.slot).forEach(g => {
    let n = 1;
    while (taken.has(n)) n++;
    g.slot = n; taken.add(n);
  });
  const slotOf = new Map(groups.map(g => [g.name, g.slot]));
  columns.forEach(col => { if (PARENT_FIELDS.has(col.field)) col.slot = slotOf.get(col.group) || 1; });
  return groups;
}

/* ================================================== finding the header row

   Schools put the school name, a report title, a print date and three blank
   rows above the actual table. Score every row near the top and take the one
   that reads like field names with data underneath it. */

const JUNK_ROW = /^(total|totals|grand total|page \d|page \d+ of|end of report|report generated|generated on|printed|continued|confidential)\b/i;
const TOTAL_LINE = /\b(total|count)\b.*[:\s]\d+|^\d+\s+(students?|records?|rows?)\b/i;

function headerScore(cells, rowsBelow) {
  const texts = cells.map(c => scrub(c));
  const filled = texts.filter(Boolean);
  if (filled.length < 2) return { score: -99, fields: new Set(), mapped: 0 };
  let score = 0, mapped = 0;
  const fields = new Set();
  texts.forEach(t => {
    if (!t) return;
    if (t.length > 45) { score -= 3; return; }
    if (/@/.test(t) || /^\(?\d{3}\)?[\s.-]?\d{3}/.test(t)) { score -= 4; return; } // that is data
    if (/^\d+([.,]\d+)?$/.test(t)) { score -= 2; return; }
    const d = detectColumn(t);
    if (d.field !== F.IGNORE) {
      mapped++;
      fields.add(d.field);
      score += d.conf === "sure" ? 3 : d.conf === "likely" ? 2 : 1;
    } else if (d.sensitive) { score += 1; fields.add("sensitive"); }
  });
  const hasName = fields.has(F.FIRST) || fields.has(F.LAST) || fields.has(F.FULL);
  if (!hasName) score -= 6;
  if (fields.size >= 3) score += 2;
  /* A header with nothing under it is a title, not a header. */
  const below = rowsBelow.filter(r => r.some(c => scrub(c)));
  if (!below.length) score -= 10;
  else {
    const width = filled.length;
    const avg = below.reduce((n, r) => n + r.filter(c => scrub(c)).length, 0) / below.length;
    if (avg >= width * 0.5) score += 3;
    else if (avg < width * 0.25) score -= 2;
  }
  return { score, fields, mapped, hasName };
}

/* Returns {row, score, fields, mapped} or null when nothing looks like a table. */
function findHeaderRow(grid) {
  const limit = Math.min(grid.length, 40);
  let best = null;
  for (let r = 0; r < limit; r++) {
    const s = headerScore(grid[r] || [], grid.slice(r + 1, r + 6));
    if (s.score >= 4 && (!best || s.score > best.score)) best = { row: r, ...s };
  }
  return best;
}

/* ============================================================ reading files

   Nothing here mutates the uploaded file -- we only ever read it. Values are
   flattened to text at the door: formulas become their cached result, merged
   cells are filled across their range, dates become the string Excel was
   showing, and anything Excel stored as a number keeps its raw value alongside
   the text so Student IDs can be rescued from scientific notation. */

function parseCSV(text) {
  let t = String(text || "");
  if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
  const head = t.slice(0, 4000);
  const counts = [[",", 0], [";", 0], ["\t", 0], ["|", 0]].map(([d]) => [d, (head.split(d).length - 1)]);
  counts.sort((a, b) => b[1] - a[1]);
  const delim = counts[0][1] > 0 ? counts[0][0] : ",";

  const rows = [];
  let row = [], field = "", quoted = false;
  for (let i = 0; i < t.length; i++) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else field += ch;
      continue;
    }
    if (ch === '"') { quoted = true; continue; }
    if (ch === delim) { row.push(field); field = ""; continue; }
    if (ch === "\r") continue;
    if (ch === "\n") { row.push(field); rows.push(row); row = []; field = ""; continue; }
    field += ch;
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows;
}

const isCSVName = name => /\.(csv|txt|tsv)$/i.test(String(name || ""));

/* data: an ArrayBuffer for xlsx/xls, or a string for csv. */
function readWorkbook(XLSX, data, fileName) {
  const book = { fileName: String(fileName || "roster"), sheets: [] };

  if (isCSVName(fileName) || typeof data === "string") {
    const rows = parseCSV(typeof data === "string" ? data : new TextDecoder().decode(data));
    book.sheets.push({
      name: String(fileName || "CSV").replace(/\.[^.]+$/, ""),
      grid: rows, nums: rows.map(r => r.map(() => null)),
      meta: { hiddenRows: 0, hiddenCols: 0, merges: 0, formulas: 0, dates: 0, filtered: false, csv: true }
    });
    return book;
  }

  const wb = XLSX.read(data, { type: "array", cellDates: false, cellNF: true, cellText: true, cellStyles: true });
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const meta = { hiddenRows: 0, hiddenCols: 0, merges: 0, formulas: 0, dates: 0, filtered: !!ws["!autofilter"], csv: false };
    if (!ws || !ws["!ref"]) { book.sheets.push({ name, grid: [], nums: [], meta }); return; }
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const rowInfo = ws["!rows"] || [], colInfo = ws["!cols"] || [];
    const grid = [], nums = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const line = [], nline = [];
      if (rowInfo[r] && rowInfo[r].hidden) meta.hiddenRows++;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell) { line.push(""); nline.push(null); continue; }
        if (cell.f) meta.formulas++;
        const isDate = cell.t === "d" || (cell.z && /[ymdhs]/.test(String(cell.z)) && cell.t === "n");
        if (isDate) meta.dates++;
        let text = cell.w != null ? cell.w : (cell.v != null ? String(cell.v) : "");
        if (cell.t === "d" && cell.v instanceof Date && !cell.w) text = cell.v.toISOString().slice(0, 10);
        line.push(text);
        nline.push(cell.t === "n" && !isDate && typeof cell.v === "number" ? cell.v : null);
      }
      grid.push(line); nums.push(nline);
    }
    for (let c = range.s.c; c <= range.e.c; c++) if (colInfo[c] && colInfo[c].hidden) meta.hiddenCols++;

    /* Un-merge: a teacher name merged down a class block belongs on every row. */
    (ws["!merges"] || []).forEach(m => {
      meta.merges++;
      const sr = m.s.r - range.s.r, sc = m.s.c - range.s.c;
      const val = grid[sr] && grid[sr][sc] !== undefined ? grid[sr][sc] : "";
      if (!scrub(val)) return;
      for (let r = m.s.r - range.s.r; r <= m.e.r - range.s.r; r++) {
        for (let c = m.s.c - range.s.c; c <= m.e.c - range.s.c; c++) {
          if (!grid[r]) continue;
          if (!scrub(grid[r][c])) grid[r][c] = val;
        }
      }
    });

    book.sheets.push({ name, grid, nums, meta, offsetRow: range.s.r, offsetCol: range.s.c });
  });
  return book;
}

/* ============================================ what a sheet name can tell us */

const SHEET_NOISE = /^(sheet\d*|data|roster|list|report|export|students?|master|all|main|table|page \d+)$/i;

/* "Grade 3" or "3rd" as a tab name means grade 3 -- but only when the tab name
   is unambiguous, and only to fill blanks. */
function sheetImplies(name) {
  const s = scrub(name);
  const out = { grade: "", teacher: "" };
  if (!s || SHEET_NOISE.test(s)) return out;
  const g = normGrade(s);
  if (g.ok && g.value && GRADE_ORDER.indexOf(g.value) >= 0) { out.grade = g.value; return out; }
  const m = s.match(/^(.*?)\s*[-_]\s*(.*)$/);
  if (m) {
    const left = normGrade(m[1]), right = normGrade(m[2]);
    if (left.ok && GRADE_ORDER.indexOf(left.value) >= 0) {
      out.grade = left.value;
      if (/[A-Za-z]/.test(m[2])) out.teacher = normTeacher(m[2]).value;
      return out;
    }
    if (right.ok && GRADE_ORDER.indexOf(right.value) >= 0) {
      out.grade = right.value;
      if (/[A-Za-z]/.test(m[1])) out.teacher = normTeacher(m[1]).value;
      return out;
    }
  }
  /* A tab named "Mrs. Alvarez" is a classroom. A tab named "Alvarez" might be
     anything, so only trust it when there is a title in front. */
  if (/^(mr|mrs|ms|miss|mx|dr)\.?\s+\S+/i.test(s)) out.teacher = normTeacher(s).value;
  return out;
}

/* ==================================================== source rows to records */

const sheetSignature = headerCells => headerCells.map(h => headerNorm(h)).join("|");

/* One column definition per source column, with the user's override applied. */
function buildColumns(sheet, headerRow, overrides) {
  const cells = (sheet.grid[headerRow] || []);
  const sig = sheetSignature(cells);
  const cols = cells.map((raw, index) => {
    const d = detectColumn(raw);
    const col = {
      index, header: scrub(raw) || "(unnamed column " + colLetter(index) + ")",
      mapKey: sig + "#" + index,
      field: d.field, group: d.group, prefer: d.prefer, conf: d.conf, note: d.note,
      sensitive: d.sensitive, questionable: d.questionable, phoneKind: d.phoneKind,
      auto: d.field, autoSlot: null, slot: null, overridden: false, sample: ""
    };
    const ov = overrides[col.mapKey];
    if (ov !== undefined) {
      const [f, s] = String(ov).split(":");
      col.field = f || F.IGNORE;
      col.prefer = s ? parseInt(s, 10) : col.prefer;
      col.group = PARENT_FIELDS.has(col.field) ? "user" + (s || "1") : "";
      col.overridden = true;
      col.conf = "sure";
      col.questionable = false;
    }
    /* An emergency contact is never a buyer until the user says so. */
    if (!col.overridden && col.questionable) col.field = F.IGNORE;
    return col;
  });
  assignSlots(cols);
  cols.forEach(c => {
    if (c.overridden && PARENT_FIELDS.has(c.field) && c.prefer) c.slot = c.prefer;
    for (let r = headerRow + 1; r < Math.min(sheet.grid.length, headerRow + 40) && !c.sample; r++) {
      const v = scrub((sheet.grid[r] || [])[c.index]);
      if (v) c.sample = v;
    }
  });
  return { cols, sig };
}

function colLetter(i) {
  let s = "", n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

const PHONE_RANK = { cell: 0, other: 1, "": 1, home: 2, work: 3 };

/* Group the values collected for one parent slot back into the columns they
   came from -- "src" is sheet:row:column:index. */
function byColumn(items) {
  const map = new Map();
  items.forEach(it => {
    const col = it.src.split(":").slice(0, 3).join(":");
    if (!map.has(col)) map.set(col, []);
    map.get(col).push(it);
  });
  return Array.from(map.values());
}
const colMax = cols => cols.reduce((n, c) => Math.max(n, c.length), 0);

/* Best value at position i across the columns, plus whatever it beat. */
function pickAt(cols, i) {
  const here = cols.map(c => c[i]).filter(Boolean);
  const good = here.filter(v => !v.bad);
  const chosen = good[0] || here[0] || {};
  return Object.assign({}, chosen, { extras: here.filter(v => v !== chosen) });
}

/* Turn one source row into a record, or return null when the row is not a
   student (blank, a repeated header, a subtotal, a page footer). */
function readRow(sheet, sheetIdx, r, cols, ctx, headerKeys) {
  const cells = sheet.grid[r] || [], nums = sheet.nums[r] || [];
  const anyText = cells.some(c => scrub(c));
  if (!anyText) { ctx.stats.blankRows++; return null; }

  const filled = cells.map(c => scrub(c)).filter(Boolean);
  const asKeys = cells.map(c => key(c)).filter(Boolean).join("|");
  if (headerKeys && asKeys && asKeys === headerKeys) { ctx.stats.repeatHeaders++; return null; }
  if (filled.length <= 2 && filled.some(t => JUNK_ROW.test(t) || TOTAL_LINE.test(t))) {
    ctx.stats.junkRows++; ctx.junk.push({ sheet: sheet.name, row: r + 1, text: filled.join(" ") });
    return null;
  }
  if (filled.length === 1 && (/^\d{1,4}$/.test(filled[0]) || filled[0].length > 60)) {
    ctx.stats.junkRows++; ctx.junk.push({ sheet: sheet.name, row: r + 1, text: filled[0] });
    return null;
  }

  const rec = {
    sheet: sheet.name, sheetIdx, row: r + 1, first: "", last: "", middle: "", suffix: "",
    grade: "", teacher: "", id: "", idNumeric: false, contacts: [], changes: [], notes: [],
    nameConf: "sure", nameNote: "", altFirst: "", rawName: "", rawGrade: "", rawTeacher: "",
    gradeOK: true, dropped: []
  };
  const slots = new Map();
  const getSlot = n => {
    if (!slots.has(n)) slots.set(n, { names: [], firsts: [], lasts: [], emails: [], phones: [] });
    return slots.get(n);
  };
  let fullName = "", firstCell = "", lastCell = "";

  cols.forEach(col => {
    const raw = cells[col.index];
    const text = scrub(raw, ctx.stats.chars);
    if (!text && col.field !== F.ID) return;
    const src = sheetIdx + ":" + (r + 1) + ":" + col.index;
    switch (col.field) {
      case F.FIRST: firstCell = firstCell || text; break;
      case F.LAST: lastCell = lastCell || text; break;
      case F.MIDDLE: rec.middle = rec.middle || fixCase(text); break;
      case F.SUFFIX: rec.suffix = rec.suffix || scrub(text); break;
      case F.FULL: fullName = fullName || text; break;
      case F.GRADE: if (!rec.rawGrade) rec.rawGrade = text; break;
      case F.TEACHER: if (!rec.rawTeacher) rec.rawTeacher = text; break;
      case F.ID: {
        if (!rec.id) {
          const id = normId(raw, nums[col.index]);
          rec.id = id.value;
          rec.idNumeric = id.wasNumeric;
          if (id.note) rec.changes.push("Student ID " + id.note);
        }
        break;
      }
      case F.P_FIRST: getSlot(col.slot).firsts.push(text); break;
      case F.P_LAST: getSlot(col.slot).lasts.push(text); break;
      case F.P_FULL: splitNames(raw).forEach(n => getSlot(col.slot).names.push(n)); break;
      case F.P_EMAIL: {
        const parts = splitEmails(raw);
        if (parts.length > 1) ctx.stats.splitCells++;
        parts.forEach((p, i) => getSlot(col.slot).emails.push({ raw: p, src: src + ":" + i }));
        break;
      }
      case F.P_CELL: {
        const parts = splitPhones(raw);
        if (parts.length > 1) ctx.stats.splitCells++;
        parts.forEach((p, i) => getSlot(col.slot).phones.push({ raw: p, kind: col.phoneKind || "other", src: src + ":" + i }));
        break;
      }
      default: break;
    }
  });

  /* ---- student name ---- */
  rec.rawName = fullName || [firstCell, lastCell].filter(Boolean).join(" ");
  if (firstCell || lastCell) {
    const f = fixCase(firstCell), l = fixCase(lastCell);
    if (f !== scrub(firstCell) || l !== scrub(lastCell)) ctx.stats.namesFixed++;
    rec.first = f; rec.last = l;
    /* A single "Last, First" value dropped into a First Name column still parses. */
    if (f && !l && f.indexOf(",") >= 0) {
      const p = splitName(firstCell);
      rec.first = p.first; rec.last = p.last; rec.nameConf = p.confidence; rec.nameNote = p.note;
      rec.changes.push("Name read as Last, First");
    }
    if (!rec.first && !rec.last && fullName) {
      const p = splitName(fullName);
      Object.assign(rec, { first: p.first, last: p.last, middle: p.middle || rec.middle,
        suffix: p.suffix || rec.suffix, nameConf: p.confidence, nameNote: p.note, altFirst: p.altFirst || "" });
    }
  } else if (fullName) {
    const p = splitName(fullName);
    rec.first = p.first; rec.last = p.last;
    rec.middle = p.middle || rec.middle;
    rec.suffix = p.suffix || rec.suffix;
    rec.nameConf = p.confidence; rec.nameNote = p.note;
    rec.altFirst = p.altFirst || ""; rec.altMiddle = p.altMiddle || "";
    if (p.first || p.last) {
      ctx.stats.namesSplit++;
      rec.changes.push('Split "' + scrub(fullName) + '" into ' + (p.first || "?") + " / " + (p.last || "?"));
    }
  }

  /* ---- grade ---- */
  if (rec.rawGrade) {
    const g = normGrade(rec.rawGrade);
    rec.grade = g.value; rec.gradeOK = g.ok;
    if (g.changed) { ctx.stats.gradesFixed++; rec.changes.push('Grade "' + rec.rawGrade + '" to ' + g.value); }
    if (g.note) rec.notes.push(g.note);
  }

  /* ---- teacher ---- */
  if (rec.rawTeacher) {
    const t = normTeacher(rec.rawTeacher);
    rec.teacher = t.value;
    if (t.changed && t.value) { ctx.stats.teachersFixed++; rec.changes.push('Teacher "' + rec.rawTeacher + '" to ' + t.value); }
  }

  /* ---- contacts ----
     Values split out of ONE cell ("mom@x.com; dad@y.com") are two people, and
     stack into consecutive parent slots. Values in DIFFERENT columns of the same
     group (a cell phone and a home phone for Parent 1) are two numbers for one
     person, so the best one is exported and the other is kept in the audit. */
  Array.from(slots.keys()).sort((a, b) => a - b).forEach(slot => {
    const s = slots.get(slot);
    const names = [];
    s.names.forEach(n => names.push(splitName(n)));
    const maxNamed = Math.max(s.firsts.length, s.lasts.length);
    for (let i = 0; i < maxNamed; i++) {
      const f = s.firsts[i] || "", l = s.lasts[i] || "";
      if (f || l) names.push({ first: fixCase(f), last: fixCase(l) });
    }

    const emailCols = byColumn(s.emails).map(list => list.map(e => {
      const fixed = ctx.fixes[e.src];
      const c = cleanEmail(fixed !== undefined ? fixed : e.raw);
      if (c.dropped) rec.dropped.push({ what: "email", value: c.dropped, why: c.reason });
      else if (c.changed && c.value) ctx.stats.emailsFixed++;
      return { value: c.ok ? c.value : "", bad: !c.ok, raw: c.ok ? c.value : (fixed !== undefined ? fixed : e.raw),
        why: c.reason, src: e.src, fixed: fixed !== undefined };
    }).filter(e => e.value || e.bad));

    const phoneCols = byColumn(s.phones).map(list => list.map(p => {
      const fixed = ctx.fixes[p.src];
      const c = cleanPhone(fixed !== undefined ? fixed : p.raw);
      if (c.dropped && c.ok) rec.dropped.push({ what: "phone", value: c.dropped, why: c.reason });
      else if (c.changed && c.value) ctx.stats.phonesFixed++;
      return { value: c.value, cellKey: c.key, bad: !c.ok, raw: c.ok ? c.value : (fixed !== undefined ? fixed : p.raw),
        why: c.reason, src: p.src, kind: p.kind, fixed: fixed !== undefined };
    }).filter(p => p.value || p.bad))
      .sort((a, b) => (PHONE_RANK[a[0].kind] || 1) - (PHONE_RANK[b[0].kind] || 1));

    const width = Math.max(names.length, colMax(emailCols), colMax(phoneCols));
    for (let i = 0; i < width; i++) {
      const nm = names[i] || {};
      const em = pickAt(emailCols, i), ph = pickAt(phoneCols, i);
      if (!nm.first && !nm.last && !em.value && !em.bad && !ph.value && !ph.bad) continue;
      em.extras.forEach(x => { ctx.stats.extraContacts++; rec.dropped.push({ what: "email", value: x.value || x.raw, why: "second email for the same parent" }); });
      ph.extras.forEach(x => { ctx.stats.extraContacts++; rec.dropped.push({ what: "phone", value: x.value || x.raw, why: "second number (" + (x.kind || "other") + ") for the same parent" }); });
      rec.contacts.push({
        slot, first: nm.first || "", last: nm.last || "",
        email: em.value || "", emailBad: !!em.bad, emailRaw: em.raw || "", emailWhy: em.why || "",
        emailSrc: em.src || "", emailFixed: !!em.fixed,
        cell: ph.value || "", cellKey: ph.cellKey || "", cellBad: !!ph.bad, cellRaw: ph.raw || "",
        cellWhy: ph.why || "", cellSrc: ph.src || "", cellKind: ph.kind || "", cellFixed: !!ph.fixed
      });
    }
  });

  /* The review screen can overrule an ambiguous name split. */
  const nameFix = ctx.names[sheetIdx + ":" + (r + 1)];
  if (nameFix) {
    rec.first = nameFix.first; rec.last = nameFix.last;
    rec.nameConf = "sure"; rec.nameNote = "";
    rec.changes.push("Name split set by hand: " + nameFix.first + " / " + nameFix.last);
  }

  const hasStudent = !!(rec.first || rec.last);
  const hasAnything = hasStudent || rec.id || rec.contacts.length;
  if (!hasAnything) { ctx.stats.blankRows++; return null; }
  if (!hasStudent && !rec.contacts.length) { ctx.stats.junkRows++; return null; }
  return rec;
}

/* =================================================== is this the same child?

   Everything below decides identity, so it is deliberately conservative. Rows
   are only combined without asking when the evidence is strong: the same
   Student ID, or the same first name, last name, grade AND teacher. Anything
   short of that becomes a question for the user instead of a silent merge. */

const NICK_SETS = [
  ["robert", "bob", "rob", "bobby", "robbie"], ["william", "will", "bill", "billy", "liam"],
  ["elizabeth", "liz", "beth", "lizzie", "eliza", "betsy", "libby"],
  ["katherine", "catherine", "kate", "katie", "kathy", "cathy", "kat"],
  ["michael", "mike", "mikey", "mick"], ["matthew", "matt"], ["christopher", "chris"],
  ["nicholas", "nick", "nicky"], ["alexander", "alex", "xander", "sasha"],
  ["alexandra", "alex", "allie", "sasha"], ["daniel", "dan", "danny"], ["joseph", "joe", "joey"],
  ["james", "jim", "jimmy", "jamie"], ["john", "johnny", "jack", "jon"], ["thomas", "tom", "tommy"],
  ["anthony", "tony"], ["benjamin", "ben", "benny"], ["samuel", "sam", "sammy"],
  ["samantha", "sam", "sammy"], ["jennifer", "jen", "jenny"], ["jessica", "jess", "jessie"],
  ["stephanie", "steph"], ["patricia", "pat", "patty", "trish"], ["margaret", "maggie", "meg", "peggy"],
  ["deborah", "deb", "debbie"], ["rebecca", "becca", "becky"], ["victoria", "vicky", "tori"],
  ["gabriel", "gabe"], ["gabriella", "gabriela", "gabby", "gabi"], ["isabella", "isabela", "bella", "izzy"],
  ["olivia", "liv", "livy"], ["sophia", "sofia", "sophie"], ["madison", "maddie", "madi"],
  ["andrew", "andy", "drew"], ["edward", "ed", "eddie", "ted"], ["richard", "rick", "ricky", "richie"],
  ["charles", "charlie", "chuck"], ["timothy", "tim"], ["theodore", "teddy", "theo"],
  ["zachary", "zach"], ["jonathan", "jon", "jonny"], ["nathaniel", "nathan", "nate"],
  ["vincent", "vinny"], ["dominic", "dom"], ["frederick", "fred"], ["gregory", "greg"],
  ["lawrence", "larry"], ["kenneth", "ken", "kenny"], ["ronald", "ron", "ronnie"],
  ["donald", "don", "donnie"], ["steven", "stephen", "steve"], ["peter", "pete"],
  ["philip", "phillip", "phil"], ["raymond", "ray"], ["russell", "russ"], ["eugene", "gene"],
  ["francisco", "frank", "paco"], ["jose", "pepe"], ["guadalupe", "lupe"], ["angela", "angie"],
  ["amanda", "mandy"], ["cassandra", "cassie"], ["natalie", "nat"], ["penelope", "penny"],
  ["josephine", "josie"], ["mackenzie", "kenzie"], ["alejandro", "alejandra", "alex"],
  ["antonio", "tony"], ["emmanuel", "manny"], ["salvador", "sal"], ["ezekiel", "zeke"]
];
const NICK_OF = (() => {
  const m = new Map();
  NICK_SETS.forEach((set, i) => set.forEach(n => {
    if (!m.has(n)) m.set(n, []);
    m.get(n).push(i);
  }));
  return m;
})();
const shareNickname = (a, b) => {
  const ga = NICK_OF.get(a), gb = NICK_OF.get(b);
  return !!(ga && gb && ga.some(x => gb.indexOf(x) >= 0));
};

function editDistance(a, b, max) {
  if (a === b) return 0;
  if (Math.abs(a.length - b.length) > max) return max + 1;
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    let best = i;
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      best = Math.min(best, cur[j]);
    }
    if (best > max) return max + 1;
    prev = cur;
  }
  return prev[b.length];
}

/* How alike are two student names? Never used to merge on its own -- only to
   decide whether to ask the user about a pair. */
function nameLikeness(a, b) {
  const af = key(a.first), al = key(a.last), bf = key(b.first), bl = key(b.last);
  if (af === bf && al === bl) return { kind: "exact", why: "" };
  if (al !== bl) {
    if (af === bf && al && bl && editDistance(al, bl, 1) <= 1)
      return { kind: "typo", why: "last name differs by one letter" };
    return { kind: "none", why: "" };
  }
  if (!af || !bf) return { kind: "none", why: "" };
  if (shareNickname(af, bf)) return { kind: "nickname", why: "one first name is a nickname of the other" };
  const max = Math.min(af.length, bf.length) >= 7 ? 2 : 1;
  if (editDistance(af, bf, max) <= max) return { kind: "typo", why: "first names differ by a letter or two" };
  if (af.length >= 3 && bf.length >= 3 && (af.indexOf(bf) === 0 || bf.indexOf(af) === 0))
    return { kind: "short", why: "one first name is a shortened form of the other" };
  return { kind: "none", why: "" };
}

/* ------------------------------------------------------------- contacts */

const emailKeyOf = c => key(c.email);
const phoneKeyOf = c => c.cellKey || key(c.cell);
const contactNameKey = c => key(c.first) + key(c.last);

function sameContact(a, b) {
  const ae = emailKeyOf(a), be = emailKeyOf(b);
  if (ae && be) return ae === be;
  const ap = phoneKeyOf(a), bp = phoneKeyOf(b);
  const an = contactNameKey(a), bn = contactNameKey(b);
  if (ap && bp && ap === bp) return !an || !bn || an === bn ||
    nameLikeness({ first: a.first, last: a.last }, { first: b.first, last: b.last }).kind !== "none";
  if (an && bn && an === bn) return !(ae && be && ae !== be);
  return false;
}

/* Keep the most complete, best-looking version of each field. */
function absorbContact(into, from) {
  if (!into.first && from.first) into.first = from.first;
  if (!into.last && from.last) into.last = from.last;
  if (!into.email && from.email) { into.email = from.email; into.emailSrc = from.emailSrc; }
  if (!into.cell && from.cell) {
    into.cell = from.cell; into.cellKey = from.cellKey; into.cellSrc = from.cellSrc; into.cellKind = from.cellKind;
  }
  if (into.emailBad && from.email && !from.emailBad) { into.emailBad = false; into.email = from.email; }
  if (into.cellBad && from.cell && !from.cellBad) { into.cellBad = false; into.cell = from.cell; }
  if (from.emailBad && !into.email && !into.emailBad) {
    into.emailBad = true; into.emailRaw = from.emailRaw; into.emailWhy = from.emailWhy; into.emailSrc = from.emailSrc;
  }
  if (from.cellBad && !into.cell && !into.cellBad) {
    into.cellBad = true; into.cellRaw = from.cellRaw; into.cellWhy = from.cellWhy; into.cellSrc = from.cellSrc;
  }
}

function mergeContacts(list, stats) {
  const out = [];
  list.forEach(c => {
    const hit = out.find(o => sameContact(o, c));
    if (hit) { absorbContact(hit, c); if (stats) stats.dupContacts++; }
    else out.push(Object.assign({}, c));
  });
  return out;
}

/* ================================================== grouping rows into kids */

const rowSignature = r => [key(r.first), key(r.last), key(r.grade), key(r.teacher), key(r.id),
  r.contacts.map(c => contactNameKey(c) + key(c.email) + phoneKeyOf(c)).sort().join("~")].join("|");

const refOf = r => r.sheetIdx + ":" + r.row;
const commonest = vals => {
  const counts = new Map();
  vals.filter(Boolean).forEach(v => counts.set(v, (counts.get(v) || 0) + 1));
  let best = "", n = 0;
  counts.forEach((c, v) => { if (c > n) { n = c; best = v; } });
  return best;
};

function clusterRecords(records, ctx) {
  const reviews = [];
  const uf = records.map((_, i) => i);
  const find = i => { while (uf[i] !== i) { uf[i] = uf[uf[i]]; i = uf[i]; } return i; };
  const join = (a, b) => { const ra = find(a), rb = find(b); if (ra === rb) return false;
    if (ra < rb) uf[rb] = ra; else uf[ra] = rb; return true; };
  const push = (map, k, v) => { if (!map.has(k)) map.set(k, []); map.get(k).push(v); };

  /* --- byte-identical rows: gone, counted, and remembered in the audit --- */
  const seen = new Map();
  const active = [];
  records.forEach((r, i) => {
    const sig = rowSignature(r);
    if (seen.has(sig)) {
      const keep = records[seen.get(sig)];
      keep.exactDupRows = (keep.exactDupRows || []).concat([{ sheet: r.sheet, row: r.row }]);
      ctx.stats.exactDupes++;
      r.droppedAsExactDupe = true;
      return;
    }
    seen.set(sig, i);
    active.push(i);
  });

  /* --- same Student ID: the same child, per the spec's primary key --- */
  const byId = new Map();
  active.forEach(i => { if (records[i].id) push(byId, key(records[i].id), i); });
  byId.forEach((idxs, idk) => {
    if (idxs.length < 2) return;
    const variants = new Map();
    idxs.forEach(i => {
      const k = key(records[i].first) + "|" + key(records[i].last);
      if (k === "|") return;            // a nameless row is not a rival spelling
      if (!variants.has(k)) variants.set(k, { k, idxs: [], rec: records[i] });
      variants.get(k).idxs.push(i);
    });
    const list = Array.from(variants.values()).sort((a, b) => b.idxs.length - a.idxs.length || a.idxs[0] - b.idxs[0]);
    if (list.length <= 1) {
      for (let i = 1; i < idxs.length; i++) if (join(idxs[0], idxs[i])) ctx.stats.mergedById++;
      return;
    }
    /* Two spellings on one ID. The spec is explicit: same ID means the same
       student pending review, so they stay together and the name gets flagged. */
    const rid = "idc:" + idk;
    const choice = ctx.decisions[rid] || list[0].k;
    if (choice === "split") {
      list.forEach(v => v.idxs.forEach(i => { if (i !== v.idxs[0] && join(v.idxs[0], i)) ctx.stats.mergedById++; }));
    } else {
      for (let i = 1; i < idxs.length; i++) if (join(idxs[0], idxs[i])) ctx.stats.mergedById++;
      const chosen = variants.get(choice) || list[0];
      idxs.forEach(i => { records[i].forceName = { first: chosen.rec.first, last: chosen.rec.last }; });
    }
    reviews.push({
      id: rid, type: "id-conflict", severity: "critical",
      title: "One Student ID, two names",
      detail: "Student ID " + records[idxs[0]].id + " appears on " + idxs.length + " rows with " +
        list.length + " different names.", rank: 0,
      choice, default: list[0].k,
      options: list.map(v => ({ value: v.k, label: "Use " + [v.rec.first, v.rec.last].filter(Boolean).join(" "),
        desc: v.idxs.length + " row" + (v.idxs.length > 1 ? "s" : "") }))
        .concat([{ value: "split", label: "Keep both / review manually", desc: "treat them as different students" }]),
      cards: list.map(v => cardOf(records[v.idxs[0]], v.idxs.map(i => refOf(records[i]))))
    });
  });

  /* --- no ID: first + last + grade + teacher, and only when all of those the
         file actually provides are filled in on both rows --- */
  const groups = new Map();
  active.forEach(i => {
    const r = records[i];
    if (r.id) return;
    if (!key(r.first) || !key(r.last)) return;
    push(groups, key(r.first) + "|" + key(r.last) + "|" +
      (ctx.hasGrade ? key(r.grade) : "") + "|" + (ctx.hasTeacher ? key(r.teacher) : ""), i);
  });
  groups.forEach(idxs => {
    if (idxs.length < 2) return;
    const r0 = records[idxs[0]];
    const complete = (!ctx.hasGrade || r0.grade) && (!ctx.hasTeacher || r0.teacher);
    if (!complete) return;   // handled as a possible-duplicate question below
    for (let i = 1; i < idxs.length; i++) if (join(idxs[0], idxs[i])) ctx.stats.mergedByName++;
  });

  /* --- everything still separate that looks alike: ask, never assume --- */
  const rootsOf = () => {
    const m = new Map();
    active.forEach(i => push(m, find(i), i));
    return m;
  };
  const summaries = Array.from(rootsOf().entries()).map(([root, idxs]) => {
    const r = records[idxs[0]];
    return { root, idxs, first: r.first, last: r.last, grade: commonest(idxs.map(i => records[i].grade)),
      teacher: commonest(idxs.map(i => records[i].teacher)), id: commonest(idxs.map(i => records[i].id)), rec: r };
  });
  /* Two cheap passes instead of comparing everyone with everyone: same last
     name (catches nicknames, typos and re-spellings of the first name), and
     same first name (catches a typo in the last name). */
  const byLast = new Map(), byFirst = new Map();
  summaries.forEach(s => {
    if (!key(s.first) || !key(s.last)) return;   // half a name proves nothing
    push(byLast, key(s.last), s);
    push(byFirst, key(s.first) + "|" + key(s.last).length, s);
  });
  const buckets = [];
  let crowded = 0;
  [byLast, byFirst].forEach(map => map.forEach(list => {
    if (list.length < 2) return;
    if (list.length > 300) { crowded += list.length; return; }
    buckets.push(list);
  }));
  ctx.stats.crowdedNames = crowded;
  const pairs = [];
  const paired = new Set();
  buckets.forEach(list => {
    for (let a = 0; a < list.length; a++) for (let b = a + 1; b < list.length; b++) {
      const A = list[a], B = list[b];
      if (find(A.root) === find(B.root)) continue;
      const pid = Math.min(A.root, B.root) + "-" + Math.max(A.root, B.root);
      if (paired.has(pid)) continue;
      paired.add(pid);
      const like = nameLikeness(A, B);
      if (like.kind === "none") continue;
      const sameGrade = key(A.grade) === key(B.grade);
      if (!sameGrade && A.grade && B.grade && like.kind !== "exact") continue;

      let title = "Possible duplicate students", why = like.why, severity = "warn", rank = 3;
      if (like.kind === "exact") {
        rank = 1;
        if (A.id && B.id && key(A.id) !== key(B.id)) {
          title = "Same name, two different Student IDs"; severity = "critical"; rank = 0;
          why = "same name and grade but the IDs do not match - one of them may be a typo";
        } else if (A.grade && B.grade && !sameGrade) {
          title = "Same name in two grades"; rank = 4;
          why = "identical names in " + A.grade + " and " + B.grade + " - siblings share a last name, not a first name";
        } else if (A.teacher && B.teacher && key(A.teacher) !== key(B.teacher)) {
          why = "same name and grade, different teacher - not merged automatically";
        } else if (!A.teacher || !B.teacher) {
          why = "same name and grade, but the teacher is blank on one of them";
        } else if ((A.id && !B.id) || (B.id && !A.id)) {
          why = "same student details, but only one row has a Student ID";
        } else why = "identical name, grade and teacher";
      }
      pairs.push({ A, B, title, why, severity, rank, id: "dup:" + refOf(A.rec) + "|" + refOf(B.rec) });
    }
  });

  /* Rank first, then keep a workable number of questions: a roster where every
     name repeats could produce thousands, and burying the real conflicts under
     them would be worse than saying how many were left. */
  pairs.sort((a, b) => a.rank - b.rank || (a.id < b.id ? -1 : 1));
  pairs.forEach((p, i) => {
    const choice = ctx.decisions[p.id] || "separate";
    if (choice === "merge") { join(p.A.root, p.B.root); ctx.stats.mergedByHand++; }
    else if (i >= 400) { ctx.stats.moreDuplicates++; return; }
    reviews.push({
      id: p.id, type: "possible-duplicate", severity: p.severity, title: p.title, detail: p.why,
      rank: p.rank, choice, default: "separate",
      options: [{ value: "merge", label: "Merge as same student", desc: "one row, contacts combined" },
        { value: "separate", label: "Keep as two students", desc: "leave both rows in the roster" }],
      cards: [cardOf(p.A.rec, p.A.idxs.map(i2 => refOf(records[i2]))),
        cardOf(p.B.rec, p.B.idxs.map(i2 => refOf(records[i2])))]
    });
  });

  reviews.sort((a, b) => (a.rank || 0) - (b.rank || 0) || (a.id < b.id ? -1 : 1));
  const finalRoots = rootsOf();
  const students = Array.from(finalRoots.values())
    .sort((x, y) => x[0] - y[0])
    .map(idxs => materialize(idxs.map(i => records[i]), ctx));
  return { students, reviews };
}

function cardOf(rec, refs) {
  return {
    name: [rec.first, rec.last].filter(Boolean).join(" ") || "(no name)",
    grade: rec.grade || "-", teacher: rec.teacher || "-", id: rec.id || "blank",
    contacts: rec.contacts.map(c => [c.first, c.last].filter(Boolean).join(" ") +
      (c.email ? " " + c.email : "") + (c.cell ? " " + c.cell : "")).filter(Boolean),
    where: refs
  };
}

/* One student record out of every row that turned out to be them. */
function materialize(recs, ctx) {
  const first = recs[0];
  const forced = recs.find(r => r.forceName);
  const s = {
    first: forced ? forced.forceName.first : commonest(recs.map(r => r.first)) || first.first,
    last: forced ? forced.forceName.last : commonest(recs.map(r => r.last)) || first.last,
    grade: commonest(recs.map(r => r.grade)),
    teacher: commonest(recs.map(r => r.teacher)),
    id: commonest(recs.map(r => r.id)),
    contacts: mergeContacts([].concat.apply([], recs.map(r => r.contacts)), ctx.stats),
    sources: [], changes: [], notes: [], dropped: [], issues: [],
    mergedRows: recs.length, nameConf: first.nameConf, nameNote: first.nameNote,
    altFirst: first.altFirst || "", altMiddle: first.altMiddle || "",
    nameRef: first.sheetIdx + ":" + first.row, rawName: first.rawName, gradeOK: recs.every(r => r.gradeOK),
    rawGrade: first.rawGrade, idNumeric: recs.some(r => r.idNumeric)
  };
  recs.forEach(r => {
    s.sources.push({ sheet: r.sheet, row: r.row, merged: r !== first });
    (r.exactDupRows || []).forEach(d => s.sources.push({ sheet: d.sheet, row: d.row, exact: true }));
    r.changes.forEach(c => { if (s.changes.indexOf(c) < 0) s.changes.push(c); });
    r.notes.forEach(c => { if (s.notes.indexOf(c) < 0) s.notes.push(c); });
    r.dropped.forEach(d => s.dropped.push(d));
  });
  /* Values that disagree between merged rows are worth saying out loud. */
  const grades = Array.from(new Set(recs.map(r => r.grade).filter(Boolean)));
  const teachers = Array.from(new Set(recs.map(r => r.teacher).filter(Boolean)));
  const ids = Array.from(new Set(recs.map(r => r.id).filter(Boolean)));
  if (grades.length > 1) s.gradeConflict = grades;
  if (teachers.length > 1) s.teacherConflict = teachers;
  if (ids.length > 1) s.idConflict = ids;
  const names = Array.from(new Set(recs.map(r => [r.first, r.last].filter(Boolean).join(" ")).filter(Boolean)));
  if (names.length > 1) s.nameVariants = names;
  if (recs.length > 1) s.changes.push("Merged " + recs.length + " rows into one student");
  return s;
}

/* ================================================================ analyze

   One pure function from (parsed workbook + the user's choices) to everything
   the screen and the download need. It re-runs from scratch after every change
   the user makes, so nothing has to be undone. */

function analyze(book, opts) {
  opts = opts || {};
  const ctx = {
    decisions: opts.decisions || {}, fixes: opts.fixes || {}, names: opts.names || {},
    junk: [], hasGrade: false, hasTeacher: false,
    stats: { sourceRows: 0, blankRows: 0, junkRows: 0, repeatHeaders: 0, exactDupes: 0,
      mergedById: 0, mergedByName: 0, mergedByHand: 0, dupContacts: 0, namesFixed: 0, namesSplit: 0,
      gradesFixed: 0, teachersFixed: 0, emailsFixed: 0, phonesFixed: 0, splitCells: 0,
      extraContacts: 0, impliedGrade: 0, impliedTeacher: 0, moreDuplicates: 0, crowdedNames: 0,
      merges: 0, formulas: 0,
      hiddenRows: 0, hiddenCols: 0, dateCells: 0, chars: { odd: 0, breaks: 0 } }
  };
  const mapping = opts.mapping || {};
  const sheetChoice = opts.sheets || {};
  const useImplied = opts.implied || {};

  /* ---- 1. what is on each tab ---- */
  const sheets = book.sheets.map((sh, i) => {
    const head = findHeaderRow(sh.grid);
    const built = head ? buildColumns(sh, head.row, mapping) : { cols: [], sig: "" };
    const dataRows = head ? sh.grid.slice(head.row + 1).filter(r => r.some(c => scrub(c))).length : 0;
    const implies = sheetImplies(sh.name);
    const hasGradeCol = built.cols.some(c => c.field === F.GRADE);
    const hasTeacherCol = built.cols.some(c => c.field === F.TEACHER);
    const info = {
      index: i, name: sh.name, headerRow: head ? head.row : -1,
      headerText: head ? (sh.grid[head.row] || []).map(c => scrub(c)).filter(Boolean).join("  |  ") : "",
      cols: built.cols, sig: built.sig, dataRows, meta: sh.meta,
      looksLikeRoster: !!(head && head.hasName && dataRows > 0),
      why: !head ? "no table of student data found on this tab"
        : !head.hasName ? "no student name column on this tab"
        : dataRows === 0 ? "no rows under the header" : "",
      implies, hasGradeCol, hasTeacherCol,
      offerGrade: !!implies.grade && !hasGradeCol,
      offerTeacher: !!implies.teacher && !hasTeacherCol
    };
    info.included = sheetChoice[sh.name] === undefined ? info.looksLikeRoster : !!sheetChoice[sh.name];
    info.useGrade = useImplied["g:" + sh.name] === undefined ? info.offerGrade : !!useImplied["g:" + sh.name];
    info.useTeacher = useImplied["t:" + sh.name] === undefined ? info.offerTeacher : !!useImplied["t:" + sh.name];
    return info;
  });

  const on = sheets.filter(s => s.included && s.headerRow >= 0);
  ctx.hasGrade = on.some(s => s.hasGradeCol || (s.useGrade && s.implies.grade));
  ctx.hasTeacher = on.some(s => s.hasTeacherCol || (s.useTeacher && s.implies.teacher));

  /* ---- 2. rows to records ---- */
  const records = [];
  on.forEach(info => {
    const sh = book.sheets[info.index];
    ctx.stats.merges += sh.meta.merges; ctx.stats.formulas += sh.meta.formulas;
    ctx.stats.hiddenRows += sh.meta.hiddenRows; ctx.stats.hiddenCols += sh.meta.hiddenCols;
    ctx.stats.dateCells += sh.meta.dates;
    const headerKeys = (sh.grid[info.headerRow] || []).map(c => key(c)).filter(Boolean).join("|");
    for (let r = info.headerRow + 1; r < sh.grid.length; r++) {
      ctx.stats.sourceRows++;
      const rec = readRow(sh, info.index, r, info.cols, ctx, headerKeys);
      if (!rec) continue;
      if (!rec.grade && info.useGrade && info.implies.grade) {
        rec.grade = info.implies.grade; ctx.stats.impliedGrade++;
        rec.changes.push('Grade ' + rec.grade + ' taken from the tab name "' + info.name + '"');
      }
      if (!rec.teacher && info.useTeacher && info.implies.teacher) {
        rec.teacher = info.implies.teacher; ctx.stats.impliedTeacher++;
        rec.changes.push('Teacher taken from the tab name "' + info.name + '"');
      }
      records.push(rec);
    }
  });

  /* ---- 3. rows to students ---- */
  const { students, reviews } = clusterRecords(records, ctx);

  /* ---- 4. what still needs a human ---- */
  const withTeacher = students.filter(s => s.teacher).length;
  const withId = students.filter(s => s.id).length;
  const teacherExpected = ctx.hasTeacher && students.length > 0 && withTeacher / students.length >= 0.6;
  const idExpected = students.length > 0 && withId / students.length >= 0.8 && withId > 1;
  const issues = [];
  const add = (student, type, severity, text, extra) => {
    const item = Object.assign({ type, severity, text, student, id: type + ":" + refOf2(student) }, extra || {});
    issues.push(item);
    student.issues.push(item);
    return item;
  };

  students.forEach(s => {
    if (!s.first && !s.last) add(s, "no-name", "critical", "No student name on this row");
    else if (!s.last) add(s, "missing-last", "critical", "Missing last name" + (s.rawName ? ' (source said "' + s.rawName + '")' : ""));
    else if (!s.first) add(s, "missing-first", "critical", "Missing first name" + (s.rawName ? ' (source said "' + s.rawName + '")' : ""));
    if (!s.grade) add(s, "missing-grade", "warn", "No grade");
    else if (!s.gradeOK) add(s, "odd-grade", "warn", 'Unusual grade value "' + s.grade + '" - left exactly as the school wrote it');
    if (teacherExpected && !s.teacher) add(s, "missing-teacher", "warn", "No teacher, but most of this roster has one");
    if (idExpected && !s.id) add(s, "missing-id", "warn", "No Student ID, but most of this roster has one");
    if (s.gradeConflict) add(s, "grade-conflict", "warn", "Merged rows disagree on grade: " + s.gradeConflict.join(" / "));
    if (s.teacherConflict) add(s, "teacher-conflict", "warn", "Merged rows disagree on teacher: " + s.teacherConflict.join(" / "));
    if (s.idConflict) add(s, "id-conflict-merged", "critical", "Merged rows carry different Student IDs: " + s.idConflict.join(" / "));
    if (s.nameVariants && s.nameVariants.length > 1)
      add(s, "name-variant", "warn", "Same student spelled " + s.nameVariants.length + " ways: " + s.nameVariants.join(" / "));
    if (s.nameConf === "check" && s.rawName && (s.first || s.last))
      add(s, "name-parse", "review", "Name split needs a look: " + (s.nameNote || "ambiguous"),
        { alt: s.altFirst, altMiddle: s.altMiddle, raw: s.rawName, nameKey: s.nameRef });

    const usable = s.contacts.filter(c => c.email || c.cell);
    if (!usable.length) add(s, "no-contact", "warn", "No parent email or phone at all");
    s.contacts.forEach((c, i) => {
      const who = [c.first, c.last].filter(Boolean).join(" ") || "Parent " + (i + 1);
      if (c.emailBad) add(s, "bad-email", "review", who + ': email "' + c.emailRaw + '" - ' + c.emailWhy,
        { fixKey: c.emailSrc, value: c.emailRaw, kind: "email", who });
      if (c.cellBad) add(s, "bad-phone", "review", who + ': phone "' + c.cellRaw + '" - ' + c.cellWhy,
        { fixKey: c.cellSrc, value: c.cellRaw, kind: "phone", who });
      if (c.cell && !c.email && !c.emailBad)
        add(s, "phone-no-email", "warn", who + " has a phone but no email - GotPhoto may not import this buyer contact");
    });
  });

  /* Teacher spellings that look like one classroom written two ways. */
  const tGroups = new Map();
  students.forEach(s => {
    if (!s.teacher) return;
    const k = teacherKey(s.teacher);
    if (!k) return;
    if (!tGroups.has(k)) tGroups.set(k, new Map());
    const m = tGroups.get(k);
    m.set(s.teacher, (m.get(s.teacher) || 0) + 1);
  });
  const teacherVariants = [];
  tGroups.forEach((m, k) => {
    if (m.size < 2) return;
    teacherVariants.push({ key: k, names: Array.from(m.entries()).map(([name, n]) => ({ name, n })) });
  });

  /* ---- 5. columns we are not exporting ---- */
  const unused = [];
  on.forEach(info => info.cols.forEach(c => {
    if (c.field !== F.IGNORE) return;
    if (!c.header && !c.sample) return;
    unused.push({ sheet: info.name, header: c.header, sample: c.sample,
      why: c.sensitive ? c.sensitive : (c.note || "not needed for GotPhoto"),
      sensitive: !!c.sensitive, mapKey: c.mapKey });
  }));

  const maxParents = students.reduce((n, s) => Math.max(n, s.contacts.filter(c => c.email || c.cell || c.first || c.last).length), 0);
  const parentSlots = Math.max(2, Math.min(maxParents, 8));
  const counted = t => issues.filter(i => i.type === t).length;
  const stats = Object.assign({}, ctx.stats, {
    students: students.length,
    sheetsUsed: on.length, sheetsSkipped: sheets.length - on.length,
    rowsMerged: ctx.stats.mergedById + ctx.stats.mergedByName + ctx.stats.mergedByHand,
    namesStandardized: ctx.stats.namesFixed + ctx.stats.namesSplit,
    parentSlots,
    withEmail: students.filter(s => s.contacts.some(c => c.email)).length,
    withPhone: students.filter(s => s.contacts.some(c => c.cell)).length,
    counts: {
      "no-name": counted("no-name"), "missing-last": counted("missing-last"),
      "missing-first": counted("missing-first"), "missing-grade": counted("missing-grade"),
      "odd-grade": counted("odd-grade"), "missing-teacher": counted("missing-teacher"),
      "missing-id": counted("missing-id"), "no-contact": counted("no-contact"),
      "bad-email": counted("bad-email"), "bad-phone": counted("bad-phone"),
      "phone-no-email": counted("phone-no-email"), "name-parse": counted("name-parse"),
      "name-variant": counted("name-variant"), "grade-conflict": counted("grade-conflict"),
      "teacher-conflict": counted("teacher-conflict"), "id-conflict-merged": counted("id-conflict-merged"),
      duplicates: reviews.filter(r => r.type === "possible-duplicate" && r.choice === "separate").length,
      idConflicts: reviews.filter(r => r.type === "id-conflict").length,
      teacherVariants: teacherVariants.length
    }
  });

  return { book, sheets, students, reviews, issues, unused, stats, teacherVariants,
    junk: ctx.junk, rows: outputRows(students, parentSlots), columns: outputColumns(parentSlots) };
}

const refOf2 = s => (s.sources[0] ? s.sources[0].sheet + ":" + s.sources[0].row : "?");

/* ============================================================ the output */

function outputColumns(parents) {
  const cols = ["Student First Name", "Student Last Name", "Grade", "Teacher", "Student ID"];
  for (let i = 1; i <= parents; i++) {
    cols.push("Parent First Name " + i, "Parent Last Name " + i, "Parent Email " + i, "Parent Cell " + i);
  }
  return cols;
}

function outputRows(students, parents) {
  return students.map(s => {
    const row = [s.first, s.last, s.grade, s.teacher, s.id];
    const live = s.contacts.filter(c => c.email || c.cell || c.first || c.last);
    for (let i = 0; i < parents; i++) {
      const c = live[i] || {};
      row.push(c.first || "", c.last || "", c.email || "", c.cell || "");
    }
    return row;
  });
}

/* ====================================================== the cleaned workbook

   Three sheets: the roster GotPhoto imports, everything that still needs a
   human, and a report of what happened. Every cell is written as text so a
   Student ID of 0018472 stays 0018472. */

const ISSUE_LABEL = {
  "no-name": "No student name", "missing-last": "Missing last name",
  "missing-first": "Missing first name", "missing-grade": "Missing grade",
  "odd-grade": "Unusual grade value", "missing-teacher": "Missing teacher",
  "missing-id": "Missing Student ID", "no-contact": "No parent contact",
  "bad-email": "Invalid email", "bad-phone": "Invalid phone",
  "phone-no-email": "Phone without email", "name-parse": "Name split needs review",
  "name-variant": "Name spelled two ways", "grade-conflict": "Conflicting grade",
  "teacher-conflict": "Conflicting teacher", "id-conflict-merged": "Conflicting Student ID",
  "possible-duplicate": "Possible duplicate students", "id-conflict": "One ID, two names"
};
const SEVERITY_LABEL = { critical: "Fix before import", warn: "Check", review: "Decide" };
const NAVY = "FF16324A", CREAM = "FFFDF6EC", TEAL = "FF177E7E", CORAL = "FFE4643B";

function sheetName(file) { return String(file || "roster").replace(/\.[^.]+$/, ""); }
const outputFileName = file => sheetName(file) + " - CLEANED.xlsx";

function styleHeader(ws, row, color) {
  row.eachCell(cell => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color || NAVY } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    cell.border = { bottom: { style: "thin", color: { argb: "FFB9C6D1" } } };
  });
  row.height = 20;
}
function autoWidth(ws, min, max) {
  ws.columns.forEach(col => {
    let w = min || 10;
    col.eachCell({ includeEmpty: false }, c => {
      const len = String(c.value == null ? "" : c.value).length + 2;
      if (len > w) w = len;
    });
    col.width = Math.min(w, max || 42);
  });
}

function buildWorkbook(ExcelJS, result, opts) {
  opts = opts || {};
  const wb = new ExcelJS.Workbook();
  wb.creator = "School Photo Roster Fixer";
  wb.created = new Date();

  /* ---------------- 1. GOTPHOTO READY ---------------- */
  const ready = wb.addWorksheet("GOTPHOTO READY", { views: [{ state: "frozen", ySplit: 1 }] });
  styleHeader(ready, ready.addRow(result.columns));
  result.rows.forEach(r => ready.addRow(r.map(v => String(v == null ? "" : v))));
  ready.columns.forEach(c => { c.numFmt = "@"; });
  autoWidth(ready, 12, 34);
  if (result.rows.length) {
    ready.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: result.columns.length } };
  }

  /* ---------------- 2. NEEDS REVIEW ---------------- */
  const rev = wb.addWorksheet("NEEDS REVIEW", { views: [{ state: "frozen", ySplit: 1 }] });
  styleHeader(rev, rev.addRow(["Issue", "What to do", "Student", "Grade", "Teacher", "Student ID",
    "Source sheet", "Source row", "Details"]), CORAL);
  const where = s => (s.sources || []).map(x => x.row).join(", ");
  const sheetOf = s => Array.from(new Set((s.sources || []).map(x => x.sheet))).join(", ");
  result.issues.filter(i => i.severity !== "info").forEach(i => {
    const s = i.student;
    rev.addRow([ISSUE_LABEL[i.type] || i.type, SEVERITY_LABEL[i.severity] || i.severity,
      [s.first, s.last].filter(Boolean).join(" ") || "(no name)", s.grade || "", s.teacher || "",
      s.id || "", sheetOf(s), where(s), i.text]);
  });
  result.reviews.filter(r => (r.type === "possible-duplicate" && r.choice === "separate") || r.type === "id-conflict")
    .forEach(r => {
      const names = r.cards.map(c => c.name + " (" + c.where.join(", ") + ")").join("  vs  ");
      rev.addRow([ISSUE_LABEL[r.type] || r.type, "Decide", r.cards.map(c => c.name).join(" / "),
        r.cards.map(c => c.grade).join(" / "), r.cards.map(c => c.teacher).join(" / "),
        r.cards.map(c => c.id).join(" / "),
        Array.from(new Set(r.cards.map(c => c.where.map(w => w.split(":")[0]).join(",")))).join(" / "),
        r.cards.map(c => c.where.join(",")).join(" / "), r.title + " - " + r.detail + ". " + names]);
    });
  result.teacherVariants.forEach(t => {
    rev.addRow(["Teacher spelled two ways", "Check", "", "", t.names.map(n => n.name).join(" / "), "", "", "",
      "These may be one classroom: " + t.names.map(n => n.name + " (" + n.n + ")").join(", ") +
      ". Nothing was merged."]);
  });
  if (rev.rowCount === 1) rev.addRow(["Nothing", "", "", "", "", "", "", "", "Every row came through clean."]);
  autoWidth(rev, 10, 70);

  /* ---------------- 3. CLEANUP REPORT ---------------- */
  const rep = wb.addWorksheet("CLEANUP REPORT");
  const st = result.stats;
  const section = title => {
    const r = rep.addRow([title, ""]);
    r.getCell(1).font = { bold: true, size: 12, color: { argb: NAVY } };
    r.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };
    r.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: CREAM } };
  };
  const line = (label, value) => { if (value !== 0 || opts.allLines) rep.addRow([label, value]); };
  styleHeader(rep, rep.addRow(["School Photo Roster Fixer", sheetName(result.book.fileName)]), TEAL);
  rep.addRow(["Cleaned", new Date().toLocaleString()]);
  rep.addRow([]);
  section("Where it came from");
  line("Source rows read", st.sourceRows);
  line("Sheets used", st.sheetsUsed);
  line("Sheets skipped", st.sheetsSkipped);
  line("Final students", st.students);
  rep.addRow([]);
  section("Fixed automatically");
  line("Duplicate household rows merged", st.rowsMerged);
  line("  ...matched on Student ID", st.mergedById);
  line("  ...matched on name + grade + teacher", st.mergedByName);
  line("  ...merged because you said so", st.mergedByHand);
  line("Exact duplicate rows removed", st.exactDupes);
  line("Duplicate parent contacts collapsed", st.dupContacts);
  line("Names standardized", st.namesStandardized);
  line("Full names split into first/last", st.namesSplit);
  line("Grade values normalized", st.gradesFixed);
  line("Teacher names tidied", st.teachersFixed);
  line("Phone numbers normalized", st.phonesFixed);
  line("Email addresses cleaned", st.emailsFixed);
  line("Cells holding two contacts split", st.splitCells);
  line("Grades taken from the tab name", st.impliedGrade);
  line("Teachers taken from the tab name", st.impliedTeacher);
  line("Blank rows ignored", st.blankRows);
  line("Title / total / footer rows ignored", st.junkRows);
  line("Repeated header rows ignored", st.repeatHeaders);
  line("Merged cells filled down", st.merges);
  line("Formulas flattened to values", st.formulas);
  line("Hidden rows included", st.hiddenRows);
  line("Hidden columns read", st.hiddenCols);
  line("Odd characters cleaned up", st.chars.odd);
  rep.addRow([]);
  section("Needs review");
  const c = st.counts;
  line("Possible duplicate students", c.duplicates);
  line("One Student ID, two names", c.idConflicts);
  line("Missing student name", c["no-name"]);
  line("Missing last name", c["missing-last"]);
  line("Missing first name", c["missing-first"]);
  line("Missing grade", c["missing-grade"]);
  line("Unusual grade values", c["odd-grade"]);
  line("Missing teacher", c["missing-teacher"]);
  line("Missing Student ID", c["missing-id"]);
  line("No parent contact at all", c["no-contact"]);
  line("Invalid emails", c["bad-email"]);
  line("Invalid phones", c["bad-phone"]);
  line("Parent phone without email", c["phone-no-email"]);
  line("Names needing a split check", c["name-parse"]);
  line("Teacher spelled more than one way", c.teacherVariants);
  line("Look-alike pairs found but not listed", st.moreDuplicates);
  line("Students in name blocks too big to compare pair by pair", st.crowdedNames);
  rep.addRow([]);
  section("Columns not exported");
  if (!result.unused.length) rep.addRow(["(none)", ""]);
  result.unused.forEach(u => rep.addRow([u.sheet + " - " + u.header, u.why + (u.sample ? '  e.g. "' + u.sample + '"' : "")]));
  rep.addRow([]);
  rep.addRow(["Contacts per student exported", st.parentSlots]);
  rep.addRow(["Students with an email", st.withEmail]);
  rep.addRow(["Students with a phone", st.withPhone]);
  rep.getColumn(1).width = 44;
  rep.getColumn(2).width = 54;

  /* ---------------- 4. SOURCE TRACE (optional) ---------------- */
  if (opts.trace) {
    const tr = wb.addWorksheet("SOURCE TRACE", { views: [{ state: "frozen", ySplit: 1 }] });
    styleHeader(tr, tr.addRow(["Student", "Grade", "Teacher", "Student ID", "Came from", "Rows merged",
      "What changed", "Not exported"]), TEAL);
    result.students.forEach(s => {
      tr.addRow([[s.first, s.last].filter(Boolean).join(" "), s.grade, s.teacher, s.id,
        s.sources.map(x => x.sheet + " row " + x.row + (x.exact ? " (exact duplicate)" : x.merged ? " (merged)" : "")).join("; "),
        s.sources.length, s.changes.join("; "),
        s.dropped.map(d => d.what + ' "' + d.value + '" - ' + d.why).join("; ")]);
    });
    tr.columns.forEach(col => { col.numFmt = "@"; });
    autoWidth(tr, 12, 60);
  }

  return wb.xlsx.writeBuffer();
}

/* ================================================================= exports */

return {
  scrub, scrubMulti, key, looseKey, fixCase, splitName, normGrade, normTeacher, teacherKey, normId,
  cleanEmail, cleanPhone, splitEmails, splitPhones, splitNames, headerNorm, detectColumn, assignSlots,
  findHeaderRow, headerScore, parseCSV, readWorkbook, sheetImplies, buildColumns, analyze,
  buildWorkbook, outputColumns, outputFileName, nameLikeness, editDistance, mergeContacts,
  colLetter, fieldLabel, FIELD_LABEL, F, PARENT_FIELDS, GRADE_ORDER, ISSUE_LABEL, SEVERITY_LABEL
};
});
