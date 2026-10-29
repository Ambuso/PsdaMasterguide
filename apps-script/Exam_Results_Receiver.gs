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
function sendResultEmail(email, name, papers) {
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
  const subject = both ? "Your results: Master Guide Examinations 2026" : "Your results: " + list[0].exam;
  const intro = both ? "Here are your results for the Master Guide Examinations 2026."
                     : "Here are your results for the " + list[0].exam + ".";
  const note = "Short answers in Section C were marked automatically from key words. Your instructor will check them, so your final mark may change slightly.";

  const text =
    "Dear " + name + ",\n\n" + intro + "\n\n" +
    (both ? "Total: " + grand + " out of " + max + " (" + pct + "%)\n\n" : "") +
    list.map(p =>
      p.exam + ": " + p.total + "/60 (" + Math.round(p.total / 60 * 100) + "%)\n" +
      labels.map((l, i) => "  " + l + ": " + p.parts[i] + "/15").join("\n") + "\n" +
      "  Time used: " + p.minutes + " minutes" + (p.timedOut ? " (time ran out)" : "")
    ).join("\n\n") + "\n\n" + note + "\n\n" +
    "God bless you,\n" + INSTRUCTOR + "\nInstructor, Master Guide class\n\n——\n" + FOOTNOTE;

  const cell = 'padding:6px 8px;border-bottom:1px solid #d9e0e7';
  const html =
    '<div style="font-family:Arial,sans-serif;color:#18212b;max-width:520px">' +
    "<p>Dear " + esc(name) + ",</p>" +
    "<p>" + esc(intro) + "</p>" +
    (both ? '<p style="font-size:28px;margin:12px 0"><b>' + grand + "</b> / " + max +
      ' <span style="color:#5b6774;font-size:16px">(' + pct + "%)</span></p>" : "") +
    list.map(p =>
      '<p style="margin:20px 0 6px;font-size:' + (both ? "16px" : "28px") + '"><b>' + (both ? esc(p.exam) + " · " : "") + p.total + "</b> / 60" +
      ' <span style="color:#5b6774;font-size:14px">(' + Math.round(p.total / 60 * 100) + "%)</span></p>" +
      '<table style="border-collapse:collapse;width:100%">' +
      labels.map((l, i) => '<tr><td style="' + cell + '">' + l + '</td><td style="' + cell + ';text-align:right"><b>' +
        p.parts[i] + "</b> / 15</td></tr>").join("") +
      "</table>" +
      '<p style="color:#5b6774;margin:6px 0 0">Time used: ' + p.minutes + " minutes" + (p.timedOut ? " (time ran out)" : "") + "</p>"
    ).join("") +
    '<p style="color:#5b6774;margin-top:18px">' + note + "</p>" +
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

function doGet() {
  return reply({ ok: true, message: "Master Guide results receiver is running." });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
