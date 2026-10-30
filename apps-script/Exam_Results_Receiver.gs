/**
 * Master Guide Examinations 2026 — results receiver
 * Instructor: Ambuso Dismas
 *
 * Every time a candidate submits a paper on https://ambuso.github.io/PsdaMasterguide/
 * it is added as a new row to the sheet "Master Guide Exam Results 2026".
 * When the candidate has finished BOTH papers, they get ONE email with both results.
 *
 * TO UPDATE: paste ALL of this file over the old code, press Ctrl + S, then
 * Deploy → Manage deployments → pencil → Version: New version → Deploy.
 */

const SHEET_ID = "1RBaEx5wziGCB-0XMMqk8kf7Fwc909fyYGytdF3qaP5I";
const INSTRUCTOR = "Ambuso Dismas";
const SEND_EMAILS = true;   // set to false to stop emailing candidates
const FOOTNOTE = "How can I lose when I came here with nothing?";   // shown at the bottom of every results email

// ── Settings for "Send all results" (sendAllResults) ──
const EXAM_NAMES = ["Church Heritage Examination 2026", "Bible Truths Examination 2026"];
const RESULTS_NOTE = "I went through the short answers (Section C) again and added marks where the answer was right but the spelling or wording was different. These are your final scores.";   // shown in the "send all" emails
const SKIP_EMAILS = ["dismasmik3@gmail.com", "dismas@firesidefoundation.org"];   // addresses that should NOT get an email, e.g. ["dismasmik3@gmail.com"] for your own test attempts

const HEADERS = ["Submitted at", "Examination", "Full name", "Local church", "Email",
  "Score /60", "A True/False /15", "B Multiple choice /15", "C Short answers /15", "D Matching /15",
  "Minutes used", "Time ran out", "Email sent",
  "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15",
  "All answers", "Attempt ID"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    // Keep the header row up to date.
    const first = sheet.getLastRow() ? sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0] : [];
    if (first.join("|") !== HEADERS.join("|")) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Ignore a repeat of an attempt that is already recorded (e.g. a retry after a poor connection).
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, HEADERS.length, last - 1, 1).getValues().flat().map(String);
      if (ids.indexOf(String(r.attemptId)) !== -1) return reply({ ok: true, duplicate: true });
    }

    const clean = (v, max) => String(v == null ? "" : v).slice(0, max || 500).replace(/^[=+\-@]/, "'$&");
    const parts = toParts(r.parts);
    const total = Number(r.total) || 0;
    const email = String(r.email || r.candidateEmail || "").trim();
    const shortAns = (r.shortAnswers || []).slice(0, 15).map(v => clean(v));
    while (shortAns.length < 15) shortAns.push("");

    // Write the row first; the email decision is made from the sheet itself.
    sheet.appendRow([
      new Date(), clean(r.exam), clean(r.name), clean(r.church), clean(email),
      total, parts[0], parts[1], parts[2], parts[3],
      Number(r.minutes) || 0, r.timedOut ? "Yes" : "No", ""
    ].concat(shortAns, [clean(r.allAnswers, 2000), String(r.attemptId)]));
    const thisRow = sheet.getLastRow();
    sheet.getRange(thisRow, 13).setValue(decideEmail(sheet, email, String(r.name || "Candidate")));

    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * One email per candidate: only when BOTH papers are in the sheet under the same email address,
 * and only if that candidate has not already been sent their combined results.
 * This works whatever version of the exam page the candidate's phone is using.
 */
function decideEmail(sheet, email, name) {
  if (!SEND_EMAILS) return "Emails turned off";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return "No email given";
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const mine = rows.filter(row => String(row[4]).trim().toLowerCase() === email.toLowerCase());
  if (mine.some(row => String(row[12]).indexOf("Yes (both") === 0)) return "Already emailed";
  const latest = {};
  mine.forEach(row => { latest[String(row[1])] = row; });   // latest row for each paper
  const papers = Object.keys(latest).map(exam => {
    const row = latest[exam];
    return { exam: exam, total: row[5], parts: [row[6], row[7], row[8], row[9]], minutes: row[10], timedOut: row[11] === "Yes" };
  });
  if (papers.length < 2) return "Waits for second paper";
  papers.sort((a, b) => a.exam < b.exam ? 1 : -1);   // Church Heritage first, then Bible Truths
  try {
    sendResultEmail(email, name, papers);
    return "Yes (both papers)";
  } catch (mailErr) {
    return "Failed: " + String(mailErr).slice(0, 120);
  }
}

function toParts(p) {
  const parts = (p || []).slice(0, 4).map(n => Number(n) || 0);
  while (parts.length < 4) parts.push(0);
  return parts;
}

/** Sends one email listing every paper in `papers` (each: exam, total, parts, minutes, timedOut). */
function sendResultEmail(email, name, papers, opts) {
  opts = opts || {};
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const labels = ["A · True or False", "B · Multiple choice", "C · Short answers", "D · Matching"];
  const list = papers.map(p => ({
    exam: String(p.exam || "Examination"), total: Number(p.total) || 0, parts: toParts(p.parts),
    minutes: Number(p.minutes) || 0, timedOut: !!p.timedOut
  }));
  const grand = list.reduce((a, p) => a + p.total, 0);
  const max = 60 * list.length;
  const pct = Math.round(grand / max * 100);
  const both = list.length > 1;
  const subject = opts.subject || ((both || opts.note) ? "Your results: Master Guide Examinations 2026" : "Your results: " + list[0].exam);
  const intro = opts.subject ? "Here are your updated results for the Master Guide Examinations 2026."
              : both ? "Here are your results for the Master Guide Examinations 2026."
                     : "Here are your results for the " + list[0].exam + ".";
  const note = opts.note || "Short answers in Section C were marked automatically from key words. Your instructor will check them, so your final mark may change slightly.";
  const missing = (opts.missing || []).length
    ? "You have not yet sat the " + opts.missing.join(" or the ") + ". Please speak to your instructor." : "";
  const named = both || !!missing;

  const text =
    "Dear " + name + ",\n\n" + intro + "\n\n" +
    (both ? "Total: " + grand + " out of " + max + " (" + pct + "%)\n\n" : "") +
    list.map(p =>
      p.exam + ": " + p.total + "/60 (" + Math.round(p.total / 60 * 100) + "%)\n" +
      labels.map((l, i) => "  " + l + ": " + p.parts[i] + "/15").join("\n") + "\n" +
      "  Time used: " + p.minutes + " minutes" + (p.timedOut ? " (time ran out)" : "")
    ).join("\n\n") + "\n\n" + (missing ? missing + "\n\n" : "") + note + "\n\n" +
    "God bless you,\n" + INSTRUCTOR + "\nInstructor, Master Guide class\n\n——\n" + FOOTNOTE;

  const cell = 'padding:6px 8px;border-bottom:1px solid #d9e0e7';
  const html =
    '<div style="font-family:Arial,sans-serif;color:#18212b;max-width:520px">' +
    "<p>Dear " + esc(name) + ",</p>" +
    "<p>" + esc(intro) + "</p>" +
    (both ? '<p style="font-size:28px;margin:12px 0"><b>' + grand + "</b> / " + max +
      ' <span style="color:#5b6774;font-size:16px">(' + pct + "%)</span></p>" : "") +
    list.map(p =>
      '<p style="margin:20px 0 6px;font-size:' + (named ? "16px" : "28px") + '"><b>' + (named ? esc(p.exam) + " · " : "") + p.total + "</b> / 60" +
      ' <span style="color:#5b6774;font-size:14px">(' + Math.round(p.total / 60 * 100) + "%)</span></p>" +
      '<table style="border-collapse:collapse;width:100%">' +
      labels.map((l, i) => '<tr><td style="' + cell + '">' + l + '</td><td style="' + cell + ';text-align:right"><b>' +
        p.parts[i] + "</b> / 15</td></tr>").join("") +
      "</table>" +
      '<p style="color:#5b6774;margin:6px 0 0">Time used: ' + p.minutes + " minutes" + (p.timedOut ? " (time ran out)" : "") + "</p>"
    ).join("") +
    (missing ? '<p style="margin-top:18px"><b>' + esc(missing) + "</b></p>" : "") +
    '<p style="color:#5b6774;margin-top:18px">' + esc(note) + "</p>" +
    "<p>God bless you,<br>" + esc(INSTRUCTOR) + "<br>Instructor, Master Guide class</p>" +
    '<p style="margin-top:24px;padding-top:12px;border-top:1px solid #d9e0e7;color:#5b6774;font-size:13px;font-style:italic">' + esc(FOOTNOTE) + "</p></div>";

  MailApp.sendEmail({ to: email, subject: subject, body: text, htmlBody: html, name: INSTRUCTOR + " – Master Guide class" });
}

/** Run this once from the editor (select it, click Run) to send yourself a test email with both papers. */
function testEmail() {
  sendResultEmail(Session.getActiveUser().getEmail(), "Test Candidate", [
    { exam: "Church Heritage Examination 2026", total: 50, parts: [13, 12, 11, 14], minutes: 38, timedOut: false },
    { exam: "Bible Truths Examination 2026", total: 46, parts: [12, 13, 9, 12], minutes: 41, timedOut: false }
  ]);
}

/* ════════════════════════════════════════════════════════════════════
 * SEND ALL RESULTS
 *  1. Select previewAllResults → Run. Check the Execution log and the new "Summary" tab.
 *  2. Select sendAllResults → Run. Each person gets ONE email with their latest result for each paper.
 *     People already sent their results are skipped, so it is safe to run again later
 *     (for example after more people finish, or if the daily email limit was reached).
 *  3. To correct a mark first, change the section score (columns G–J) in that person's latest row.
 *     The total is worked out from the four sections.
 *  To send everyone a fresh copy after changing marks, run resetSentList, then sendAllResults.
 * ════════════════════════════════════════════════════════════════════ */
function previewAllResults() {
  const props = PropertiesService.getScriptProperties();
  const people = collectCandidates_();
  writeSummary_(people, props);
  people.forEach(c => {
    const done = EXAM_NAMES.filter(n => c.papers[n]);
    Logger.log((props.getProperty("final:" + c.email) ? "ALREADY SENT  " : "WILL SEND     ") +
      c.name + " <" + c.email + "> · " + done.map(n => n.split(" Examination")[0] + " " + c.papers[n].total + "/60").join(", ") +
      (done.length < EXAM_NAMES.length ? " · missing " + EXAM_NAMES.filter(n => !c.papers[n]).join(", ") : ""));
  });
  Logger.log(people.length + " people. Emails left today: " + MailApp.getRemainingDailyQuota());
}

function sendAllResults() {
  const props = PropertiesService.getScriptProperties();
  const people = collectCandidates_();
  let sent = 0, skipped = 0, stopped = false;
  for (const c of people) {
    if (props.getProperty("final:" + c.email)) { skipped++; continue; }
    if (MailApp.getRemainingDailyQuota() < 1) { stopped = true; break; }
    const papers = EXAM_NAMES.filter(n => c.papers[n]).map(n => c.papers[n]);
    const missing = EXAM_NAMES.filter(n => !c.papers[n]);
    try {
      sendResultEmail(c.email, c.name, papers, { note: RESULTS_NOTE, missing: missing,
        subject: "Your updated results: Master Guide Examinations 2026" });
      props.setProperty("final:" + c.email, new Date().toISOString());
      sent++;
      Logger.log("Sent to " + c.name + " <" + c.email + ">");
    } catch (err) {
      Logger.log("FAILED for " + c.email + ": " + err);
    }
  }
  writeSummary_(people, props);
  Logger.log("Done. Sent: " + sent + ". Already sent before: " + skipped + "." +
    (stopped ? " Stopped because today's email limit was reached. Run sendAllResults again tomorrow to finish." : ""));
}

/** Clears the record of who has been sent their results, so sendAllResults sends to everyone again. */
function resetSentList() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(k => { if (k.indexOf("final:") === 0) props.deleteProperty(k); });
  Logger.log("Cleared. The next sendAllResults will email everyone.");
}

/** One entry per email address, using each person's LATEST attempt at each paper. */
function collectCandidates_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const skip = SKIP_EMAILS.map(e => String(e).trim().toLowerCase());
  const nice = n => (n === n.toUpperCase() ? n.toLowerCase().replace(/\b\w/g, ch => ch.toUpperCase()) : n);
  const byEmail = {};
  sheet.getRange(2, 1, last - 1, 13).getValues().forEach(row => {
    const email = String(row[4]).trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || skip.indexOf(email) !== -1) return;
    const exam = String(row[1]).trim();
    if (EXAM_NAMES.indexOf(exam) === -1) return;
    const c = byEmail[email] || (byEmail[email] = { email: email, name: "", church: "", papers: {} });
    if (String(row[2]).trim()) c.name = nice(String(row[2]).trim());
    if (String(row[3]).trim()) c.church = String(row[3]).trim();
    const parts = [row[6], row[7], row[8], row[9]].map(n => Number(n) || 0);
    c.papers[exam] = { exam: exam, parts: parts, total: parts.reduce((a, b) => a + b, 0),
                       minutes: Number(row[10]) || 0, timedOut: row[11] === "Yes" };
  });
  return Object.keys(byEmail).map(k => byEmail[k]).sort((a, b) => a.name.localeCompare(b.name));
}

/** Writes a "Summary" tab: one row per person with both scores and the total. */
function writeSummary_(people, props) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tab = ss.getSheetByName("Summary") || ss.insertSheet("Summary");
  tab.clear();
  const head = ["Full name", "Email", "Local church"].concat(EXAM_NAMES.map(n => n.replace(" Examination 2026", "") + " /60"),
    ["Total /" + 60 * EXAM_NAMES.length, "Papers done", "Results emailed"]);
  const rows = people.map(c => {
    const scores = EXAM_NAMES.map(n => c.papers[n] ? c.papers[n].total : "");
    const sentAt = props.getProperty("final:" + c.email);
    return [c.name, c.email, c.church].concat(scores, [
      scores.reduce((a, b) => a + (Number(b) || 0), 0),
      EXAM_NAMES.filter(n => c.papers[n]).length + " of " + EXAM_NAMES.length,
      sentAt ? new Date(sentAt) : ""]);
  });
  tab.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  if (rows.length) tab.getRange(2, 1, rows.length, head.length).setValues(rows);
  tab.setFrozenRows(1);
  tab.autoResizeColumns(1, head.length);
}

/* ════════════════════════════════════════════════════════════════════
 * SECTION C RE-MARKING (checked by hand)
 * Select applyMarkCorrections → Run. It adds the marks below to Section C (column I),
 * updates the score (column F) and leaves a note on each changed cell explaining why.
 * Running it twice does nothing extra. Run it BEFORE sendAllResults.
 * ════════════════════════════════════════════════════════════════════ */
// Your call: in Church Heritage C15 (Revelation 14:12) these four wrote "testimony of Jesus"
// (which is Revelation 12:17) instead of "faith of Jesus". Set to true to give them the mark.
const ACCEPT_TESTIMONY_FOR_CH_C15 = false;

const MARK_CORRECTIONS = [
  { id: "ch-1793294221805-43442", add: 1, why: "Jackson Mwamba, CH C2: 'Warldeness' = Waldenses (spelling)" },
  { id: "ch-1793294277791-6580",  add: 2, why: "Abigael Chelangat, CH C5: 'sign of times' = Signs of the Times; C11: 'Swizerland' = Switzerland (spelling)" },
  { id: "bt-1793296240194-66424", add: 2, why: "Alice Nyanduko, BT C4: 'Spirit of prophesy' (spelling); C15: commandments + spirit of prophecy (= testimony of Jesus)" },
  { id: "bt-1793299206596-69041", add: 1, why: "Gladys Kebaya, BT C5: 'the sin offering and the scapegoat' (the Lord's goat was the sin offering)" },
  { id: "bt-1793301180504-35466", add: 1, why: "Palmer Osuri, BT C5: 'Goat and scapegoat'" },
  { id: "bt-1793294368237-25176", add: 1, why: "Dickson Onyango, BT C8: 'Mass/ population' = peoples" },
  { id: "bt-1793295243109-39154", add: 1, why: "Cecilia Volenzo, BT C9: 'Biblios' = biblos (spelling)" }
];
const TESTIMONY_CORRECTIONS = [
  { id: "ch-1793294319447-70583", add: 1, why: "Lucy Ndiao, CH C15: 'testimony of Jesus' accepted" },
  { id: "ch-1793294280811-81220", add: 1, why: "Cecilia Volenzo, CH C15: 'testimony of Jesus' accepted" },
  { id: "ch-1793294221805-43442", add: 1, why: "Jackson Mwamba, CH C15: 'testimony of Jesus' accepted", key: "ch-1793294221805-43442#c15" },
  { id: "ch-1793295047371-12830", add: 1, why: "Osoro Nyambane, CH C15: 'testimony of Jesus' accepted" }
];

function applyMarkCorrections() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const props = PropertiesService.getScriptProperties();
  const last = sheet.getLastRow();
  const ids = sheet.getRange(2, HEADERS.length, last - 1, 1).getValues().map(r => String(r[0]));
  const list = MARK_CORRECTIONS.concat(ACCEPT_TESTIMONY_FOR_CH_C15 ? TESTIMONY_CORRECTIONS : []);
  let done = 0;
  list.forEach(c => {
    const key = "corr:" + (c.key || c.id);
    if (props.getProperty(key)) { Logger.log("Already applied: " + c.why); return; }
    const i = ids.indexOf(c.id);
    if (i === -1) { Logger.log("Row not found: " + c.why); return; }
    const row = i + 2;
    const secs = sheet.getRange(row, 7, 1, 4).getValues()[0].map(n => Number(n) || 0);   // G..J
    secs[2] = Math.min(15, secs[2] + c.add);                                               // I = Section C
    sheet.getRange(row, 9).setValue(secs[2]);
    sheet.getRange(row, 6).setValue(secs.reduce((a, b) => a + b, 0));                      // F = score /60
    const cell = sheet.getRange(row, 9);
    cell.setNote((cell.getNote() ? cell.getNote() + "\n" : "") + "+" + c.add + " " + c.why);
    props.setProperty(key, new Date().toISOString());
    done++;
    Logger.log("Applied +" + c.add + ": " + c.why);
  });
  Logger.log(done + " correction(s) applied. Now run previewAllResults, then sendAllResults.");
}

function doGet() {
  return reply({ ok: true, message: "Master Guide results receiver is running." });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
