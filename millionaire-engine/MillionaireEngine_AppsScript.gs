/**
 * ========================================================================
 *  MILLIONAIRE ENGINE — Apps Script (Form pipeline + Calendar + Email + Webhook)
 * ========================================================================
 *
 *  This script powers:
 *   1. Google Form → MASTER_LOG pipeline (onFormSubmit)
 *   2. Auto-sort MASTER_LOG by date (sortMasterLog)
 *   3. Calendar sync of MONTHLY_PAYABLES to Google Calendar (syncToCalendar)
 *   4. Daily HTML email digest to personal Gmail (sendDailyEmailDigest)
 *   5. Webhook for Claude/Cowork direct edits (doPost — see WEBHOOK_SETUP.md)
 *
 *  SETUP CHECKLIST:
 *   1. Paste this entire file into Apps Script editor (Code.gs).
 *   2. Confirm SPREADSHEET_ID matches your Sheet URL.
 *   3. Confirm WEBHOOK_SECRET is your private password.
 *   4. (Optional) Set PERSONAL_EMAIL if using email digest.
 *   5. Save (Ctrl+S).
 *   6. Deploy → New deployment → Web app → Execute as Me → Anyone.
 *   7. Test webhookPing from the Run dropdown before sharing URL.
 *
 *  MASTER_LOG column layout:
 *    A: Month formula =IF(B{r}="","",TEXT(B{r},"mmmm yyyy"))
 *    B: Date     C: Item     D: Category     E: Amount
 *    F: Payment Method (manual)     G: Status     H: Notes (manual)
 * ========================================================================
 */

const MASTER_LOG_SHEET_NAME = "MASTER_LOG";
const MONTHLY_PAYABLES_SHEET = "MONTHLY_PAYABLES";
const PAYABLES_FIRST_ROW = 24;
const PAYABLES_LAST_ROW = 60;

const FORM_QUESTION_DATE = "Date of Log Entry";
const FORM_QUESTION_ITEM = "Item";
const FORM_QUESTION_CATEGORY = "Category";
const FORM_QUESTION_AMOUNT = "Amount";
const FORM_QUESTION_STATUS = "Status";

const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";          // ← the token between /d/ and /edit in your Sheet URL
const WEBHOOK_SECRET = "SET_A_STRONG_UNIQUE_SECRET";   // ← generate a fresh 32+ char random string; never reuse across instances
const PERSONAL_EMAIL = "REPLACE_WITH_YOUR_PERSONAL_EMAIL@gmail.com"; // ← edit for email digest

const CAL_NAME = "Millionaire Engine Bills";
const CAL_TAG_PREFIX = "[engine:";
const CAL_TAG_SUFFIX = "]";
const COLOR_PENDING = "5";   // banana yellow
const COLOR_UNPAID  = "11";  // tomato red
const COLOR_PAID    = "2";   // sage green

/* ============================================================
 * Spreadsheet access — works bound OR standalone
 * ============================================================ */
function getSS() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* ignore */ }
  if (ss) return ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "REPLACE_ME_WITH_SPREADSHEET_ID") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  throw new Error("Cannot open spreadsheet: SPREADSHEET_ID not set.");
}

/* ============================================================
 * Custom menu — appears in the spreadsheet UI
 * ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️  Millionaire Engine")
    .addItem("Sort MASTER_LOG by date", "sortMasterLog")
    .addItem("Backfill from Form responses 1", "backfillFromFormResponses")
    .addSeparator()
    .addItem("📅  Sync bills to Google Calendar (now)", "syncToCalendar")
    .addItem("📅  Install daily calendar sync (6 AM)", "installCalendarTrigger")
    .addItem("📅  Remove all Engine-managed calendar events", "purgeCalendarEvents")
    .addSeparator()
    .addItem("📧  Send bill digest email NOW", "sendDailyEmailDigest")
    .addItem("📧  Install daily email digest (6 AM)", "installEmailDigestTrigger")
    .addSeparator()
    .addItem("Reinstall form trigger", "installTrigger")
    .addToUi();
}

/* ============================================================
 * FORM PIPELINE
 * ============================================================ */
function onFormSubmit(e) {
  try {
    // Defensive guard: this function is meant to be triggered by an actual
    // form submission. If you click Run from the editor with no event, exit
    // cleanly instead of crashing with "namedValues of undefined".
    if (!e || !e.namedValues) {
      console.log("onFormSubmit called without event object — exiting cleanly. " +
                  "This function is only meaningful when triggered by a real form submission.");
      return;
    }
    const ss = getSS();
    const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
    if (!sheet) { console.error("MASTER_LOG sheet not found"); return; }

    const nv = e.namedValues || {};
    const date = nv[FORM_QUESTION_DATE] ? new Date(nv[FORM_QUESTION_DATE][0]) : new Date();
    const item = nv[FORM_QUESTION_ITEM] ? nv[FORM_QUESTION_ITEM][0] : "";
    const category = nv[FORM_QUESTION_CATEGORY] ? nv[FORM_QUESTION_CATEGORY][0] : "";
    const amountRaw = nv[FORM_QUESTION_AMOUNT] ? nv[FORM_QUESTION_AMOUNT][0] : "0";
    const amount = parseFloat(String(amountRaw).replace(/[^0-9.\-]/g, "")) || 0;
    const status = nv[FORM_QUESTION_STATUS] ? nv[FORM_QUESTION_STATUS][0] : "Pending";

    const lastRow = findNextEmptyRow(sheet);
    sheet.getRange(lastRow, 1).setFormula(`=IF(B${lastRow}="","",TEXT(B${lastRow},"mmmm yyyy"))`);
    sheet.getRange(lastRow, 2).setValue(date);
    sheet.getRange(lastRow, 2).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(lastRow, 3).setValue(item);
    sheet.getRange(lastRow, 4).setValue(category);
    sheet.getRange(lastRow, 5).setValue(amount);
    sheet.getRange(lastRow, 5).setNumberFormat('"₱"#,##0.00');
    sheet.getRange(lastRow, 7).setValue(status);

    markSynced(ss);
    sortMasterLog();
  } catch (err) {
    console.error("onFormSubmit error: " + err.toString());
    console.error(err.stack);
  }
}

function findNextEmptyRow(sheet) {
  const colB = sheet.getRange("B:B").getValues();
  for (let i = 1; i < colB.length; i++) {
    if (!colB[i][0]) return i + 1;
  }
  return sheet.getLastRow() + 1;
}

function markSynced(ss) {
  const fr = ss.getSheetByName("Form responses 1");
  if (!fr) return;
  const lastRow = fr.getLastRow();
  if (lastRow < 2) return;
  const headers = fr.getRange(1, 1, 1, fr.getLastColumn()).getValues()[0];
  let syncCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().includes("synced")) { syncCol = i + 1; break; }
  }
  if (syncCol === -1) return;
  const now = new Date();
  const stamp = "✓ " +
    String(now.getMonth() + 1).padStart(2, "0") + "/" +
    String(now.getDate()).padStart(2, "0") + " " +
    String(now.getHours()).padStart(2, "0") + ":" +
    String(now.getMinutes()).padStart(2, "0");
  fr.getRange(lastRow, syncCol).setValue(stamp);
}

function installTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onFormSubmit") ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(getSS())
    .onFormSubmit()
    .create();
  SpreadsheetApp.getUi().alert("✓ Form trigger installed.");
}

function backfillFromFormResponses() {
  const ss = getSS();
  const fr = ss.getSheetByName("Form responses 1");
  const ml = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!fr || !ml) return;
  const data = fr.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0];
  const colIdx = {
    date:   headers.findIndex(h => /date/i.test(h)),
    item:   headers.findIndex(h => /^item/i.test(h)),
    cat:    headers.findIndex(h => /category/i.test(h)),
    amt:    headers.findIndex(h => /amount/i.test(h)),
    status: headers.findIndex(h => /status/i.test(h)),
    sync:   headers.findIndex(h => /synced/i.test(h))
  };
  let mirrored = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[colIdx.date]) continue;
    if (colIdx.sync >= 0 && String(row[colIdx.sync]).includes("✓")) continue;
    const targetRow = findNextEmptyRow(ml);
    ml.getRange(targetRow, 1).setFormula(`=IF(B${targetRow}="","",TEXT(B${targetRow},"mmmm yyyy"))`);
    ml.getRange(targetRow, 2).setValue(new Date(row[colIdx.date]));
    ml.getRange(targetRow, 2).setNumberFormat("yyyy-mm-dd");
    ml.getRange(targetRow, 3).setValue(row[colIdx.item]);
    ml.getRange(targetRow, 4).setValue(row[colIdx.cat]);
    ml.getRange(targetRow, 5).setValue(parseFloat(String(row[colIdx.amt]).replace(/[^0-9.\-]/g, "")) || 0);
    ml.getRange(targetRow, 5).setNumberFormat('"₱"#,##0.00');
    ml.getRange(targetRow, 7).setValue(row[colIdx.status] || "Pending");
    if (colIdx.sync >= 0) {
      const now = new Date();
      fr.getRange(r + 1, colIdx.sync + 1).setValue(
        "✓ " + String(now.getMonth() + 1).padStart(2, "0") + "/" +
        String(now.getDate()).padStart(2, "0") + " " +
        String(now.getHours()).padStart(2, "0") + ":" +
        String(now.getMinutes()).padStart(2, "0"));
    }
    mirrored++;
  }
  SpreadsheetApp.getUi().alert(`Backfill complete: ${mirrored} row(s) mirrored.`);
}

/* ============================================================
 * SORT MASTER_LOG by date — sweep-and-rewrite pattern
 * ============================================================ */
function sortMasterLog() {
  const ss = getSS();
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const numCols = Math.max(sheet.getLastColumn() - 1, 7);
  const values = sheet.getRange(2, 2, lastRow - 1, numCols).getValues();

  const realRows = [];
  values.forEach((row, idx) => {
    const hasAnyData = row.slice(0, 7).some(v => v !== "" && v !== null && v !== undefined);
    if (hasAnyData) {
      realRows.push({ origIdx: idx, values: row, dateVal: row[0], hasDate: row[0] instanceof Date });
    }
  });

  realRows.sort((a, b) => {
    if (a.hasDate && b.hasDate) {
      const cmp = a.dateVal - b.dateVal;
      return cmp !== 0 ? cmp : a.origIdx - b.origIdx;
    }
    if (a.hasDate && !b.hasDate) return -1;
    if (!a.hasDate && b.hasDate) return 1;
    return a.origIdx - b.origIdx;
  });

  sheet.getRange(2, 1, lastRow - 1, numCols + 1).clearContent();
  if (realRows.length > 0) {
    sheet.getRange(2, 2, realRows.length, numCols).setValues(realRows.map(r => r.values));
  }
  const lastDataRow = 1 + realRows.length;
  const aFormulas = [];
  for (let r = 2; r <= lastDataRow + 1; r++) {
    aFormulas.push([`=IF(B${r}="","",TEXT(B${r},"mmmm yyyy"))`]);
  }
  if (aFormulas.length > 0) sheet.getRange(2, 1, aFormulas.length, 1).setFormulas(aFormulas);

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, lastDataRow, 8).createFilter();
  return realRows.length;
}

/* ============================================================
 * CALENDAR SYNC
 * ============================================================ */
function syncToCalendar() {
  const ss = getSS();
  const sheet = ss.getSheetByName(MONTHLY_PAYABLES_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert("MONTHLY_PAYABLES sheet not found"); return; }
  const cal = getOrCreateEngineCalendar();
  const numRows = PAYABLES_LAST_ROW - PAYABLES_FIRST_ROW + 1;
  const rows = sheet.getRange(PAYABLES_FIRST_ROW, 2, numRows, 9).getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let created = 0, updated = 0, skipped = 0;

  rows.forEach(row => {
    const [item, provider, dueDay, amount, status, lastPaid, category, autoDebit, notes] = row;
    if (!item || !dueDay || !amount) { skipped++; return; }
    const targetDate = computeNextDueDate(dueDay, today);
    const tag = CAL_TAG_PREFIX + slugify(String(item)) + CAL_TAG_SUFFIX;
    const title = `💰 ${item} — ₱${Number(amount).toLocaleString("en-PH", {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
    const description = buildDescription(item, provider, amount, status, category, autoDebit, notes, tag);

    const searchStart = new Date(targetDate); searchStart.setDate(searchStart.getDate() - 35);
    const searchEnd = new Date(targetDate); searchEnd.setDate(searchEnd.getDate() + 35);
    const existing = cal.getEvents(searchStart, searchEnd).filter(ev => (ev.getDescription() || "").includes(tag));

    if (existing.length > 0) {
      const ev = existing[0];
      ev.setTitle(title);
      ev.setDescription(description);
      if (ev.getAllDayStartDate().getTime() !== targetDate.getTime()) ev.setAllDayDate(targetDate);
      applyEventColor(ev, status);
      for (let k = 1; k < existing.length; k++) existing[k].deleteEvent();
      updated++;
    } else {
      const ev = cal.createAllDayEvent(title, targetDate, { description: description });
      applyEventColor(ev, status);
      ev.removeAllReminders();
      ev.addPopupReminder(1440);
      ev.addPopupReminder(60);
      created++;
    }
  });
  SpreadsheetApp.getUi().alert(`✓ Calendar sync: ${created} created, ${updated} updated, ${skipped} skipped.`);
}

function getOrCreateEngineCalendar() {
  const cals = CalendarApp.getCalendarsByName(CAL_NAME);
  if (cals.length > 0) return cals[0];
  const cal = CalendarApp.createCalendar(CAL_NAME, {
    summary: "Recurring bills tracked by Millionaire Engine.",
    color: CalendarApp.Color.BLUE
  });
  cal.setTimeZone(Session.getScriptTimeZone());
  return cal;
}

function computeNextDueDate(dueDay, today) {
  const day = Math.max(1, Math.min(31, Number(dueDay)));
  const candidate = new Date(today.getFullYear(), today.getMonth(), day);
  candidate.setHours(0, 0, 0, 0);
  if (candidate < today) {
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const lastDayNext = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();
    candidate.setFullYear(nextMonth.getFullYear());
    candidate.setMonth(nextMonth.getMonth());
    candidate.setDate(Math.min(day, lastDayNext));
  }
  return candidate;
}

function buildDescription(item, provider, amount, status, category, autoDebit, notes, tag) {
  return [
    `Status: ${status || "—"}`,
    `Amount: ₱${Number(amount).toLocaleString("en-PH", {minimumFractionDigits: 2, maximumFractionDigits: 2})}`,
    `Provider: ${provider || "—"}`,
    `Category: ${category || "—"}`,
    `Auto-debit: ${autoDebit || "No"}`,
    "",
    notes ? `Notes: ${notes}` : "",
    "",
    "—",
    "Managed by Millionaire Engine — do not edit description (tag will desync).",
    tag
  ].filter(s => s !== undefined && s !== null).join("\n");
}

function applyEventColor(ev, status) {
  const s = String(status || "").toLowerCase();
  let colorId = COLOR_PENDING;
  if (s.includes("paid")) colorId = COLOR_PAID;
  else if (s.includes("unpaid") || s.includes("overdue") || s.includes("priority")) colorId = COLOR_UNPAID;
  try { ev.setColor(colorId); } catch (e) { console.warn("setColor failed: " + e); }
}

function slugify(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").substring(0, 50);
}

function installCalendarTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "syncToCalendar") ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger("syncToCalendar").timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert("✓ Daily calendar sync installed (6 AM).");
}

function purgeCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("Purge Engine-managed calendar events?",
    "Delete ALL [engine:*] tagged events in '" + CAL_NAME + "'. Continue?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const cals = CalendarApp.getCalendarsByName(CAL_NAME);
  if (cals.length === 0) { ui.alert("Calendar '" + CAL_NAME + "' does not exist."); return; }
  const cal = cals[0];
  const start = new Date(); start.setFullYear(start.getFullYear() - 2);
  const end = new Date(); end.setFullYear(end.getFullYear() + 2);
  const events = cal.getEvents(start, end).filter(ev => (ev.getDescription() || "").includes(CAL_TAG_PREFIX));
  let deleted = 0;
  events.forEach(ev => { ev.deleteEvent(); deleted++; });
  ui.alert("Purged " + deleted + " events.");
}

/* ============================================================
 * EMAIL DIGEST
 * ============================================================ */
function sendDailyEmailDigest() {
  const ss = getSS();
  const sheet = ss.getSheetByName(MONTHLY_PAYABLES_SHEET);
  if (!sheet) return;
  if (PERSONAL_EMAIL.startsWith("REPLACE_")) {
    try { SpreadsheetApp.getUi().alert("⚠ PERSONAL_EMAIL not set. Edit the constant near the top of the script."); } catch (e) {}
    return;
  }
  const numRows = PAYABLES_LAST_ROW - PAYABLES_FIRST_ROW + 1;
  const rows = sheet.getRange(PAYABLES_FIRST_ROW, 2, numRows, 9).getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const todayDay = today.getDate();
  const overdue = [], dueToday = [], dueSoon = [], upcoming = [];
  rows.forEach(row => {
    const [item, provider, dueDay, amount, status, lastPaid, category, autoDebit, notes] = row;
    if (!item || !dueDay || !amount) return;
    const s = String(status || "").toLowerCase();
    if (s.includes("paid")) return;
    const d = Number(dueDay);
    const entry = { item, provider, dueDay: d, amount, status, category, notes };
    if (d < todayDay) overdue.push(entry);
    else if (d === todayDay) dueToday.push(entry);
    else if (d <= todayDay + 7) dueSoon.push(entry);
    else if (d <= todayDay + 14) upcoming.push(entry);
  });
  overdue.sort((a, b) => a.dueDay - b.dueDay);
  dueToday.sort((a, b) => a.dueDay - b.dueDay);
  dueSoon.sort((a, b) => a.dueDay - b.dueDay);
  upcoming.sort((a, b) => a.dueDay - b.dueDay);

  const monthYear = today.toLocaleString("en-US", { month: "long", year: "numeric" });
  const dateStr = today.toLocaleString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" });
  const totalOverdue = overdue.reduce((s, e) => s + Number(e.amount), 0);
  const totalDueToday = dueToday.reduce((s, e) => s + Number(e.amount), 0);
  const totalDueSoon = dueSoon.reduce((s, e) => s + Number(e.amount), 0);

  const subject = overdue.length > 0
    ? `🔴 ${overdue.length} OVERDUE · ${dueToday.length} due today — Millionaire Engine ${monthYear}`
    : (dueToday.length > 0
        ? `🟡 ${dueToday.length} bills due today — Millionaire Engine ${monthYear}`
        : `📅 ${dueSoon.length} bills due this week — Millionaire Engine ${monthYear}`);

  const html = buildDigestHtml({ dateStr, monthYear, overdue, dueToday, dueSoon, upcoming, totalOverdue, totalDueToday, totalDueSoon });
  GmailApp.sendEmail(PERSONAL_EMAIL, subject, "Open in HTML-enabled client.", { htmlBody: html, name: "Millionaire Engine" });
  try {
    SpreadsheetApp.getUi().alert(`✓ Digest sent to ${PERSONAL_EMAIL}\nOverdue: ${overdue.length}, Due today: ${dueToday.length}, Next 7d: ${dueSoon.length}`);
  } catch (e) { /* not run from UI */ }
}

function buildDigestHtml(d) {
  const { dateStr, monthYear, overdue, dueToday, dueSoon, upcoming, totalOverdue, totalDueToday, totalDueSoon } = d;
  const peso = n => "₱" + Number(n).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  function bucketHtml(title, color, items, total) {
    if (items.length === 0) return "";
    const rowsHtml = items.map(e => `
      <tr>
        <td style="padding:8px 12px;border-bottom:1px solid #eceff1;font-size:13px;">Day ${e.dueDay}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eceff1;font-size:13px;font-weight:600;">${e.item}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eceff1;font-size:13px;color:#546e7a;">${e.category || ""}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eceff1;font-size:13px;text-align:right;font-weight:600;">${peso(e.amount)}</td>
      </tr>`).join("");
    return `
      <div style="margin:24px 0;">
        <div style="background:${color};color:#fff;padding:10px 14px;font-weight:700;font-size:14px;border-radius:6px 6px 0 0;">
          ${title} — ${items.length} item${items.length===1?"":"s"} · ${peso(total)}
        </div>
        <table style="width:100%;border-collapse:collapse;background:#fff;border:1px solid #eceff1;border-top:none;border-radius:0 0 6px 6px;">
          <thead><tr style="background:#f5f5f5;">
            <th style="padding:8px 12px;text-align:left;font-size:11px;color:#546e7a;font-weight:600;text-transform:uppercase;">Due</th>
            <th style="padding:8px 12px;text-align:left;font-size:11px;color:#546e7a;font-weight:600;text-transform:uppercase;">Item</th>
            <th style="padding:8px 12px;text-align:left;font-size:11px;color:#546e7a;font-weight:600;text-transform:uppercase;">Category</th>
            <th style="padding:8px 12px;text-align:right;font-size:11px;color:#546e7a;font-weight:600;text-transform:uppercase;">Amount</th>
          </tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>`;
  }
  return `
    <div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:680px;margin:0 auto;padding:24px;background:#fafafa;">
      <div style="background:#1a237e;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0;">
        <div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;opacity:0.85;">Millionaire Engine · Daily Bill Digest</div>
        <div style="font-size:20px;font-weight:700;margin-top:4px;">${dateStr}</div>
        <div style="font-size:13px;opacity:0.9;margin-top:4px;">${monthYear} obligations</div>
      </div>
      <div style="background:#fff;padding:8px 22px 22px 22px;border-radius:0 0 8px 8px;border:1px solid #eceff1;border-top:none;">
        ${overdue.length === 0 && dueToday.length === 0 && dueSoon.length === 0
          ? `<div style="padding:32px;text-align:center;color:#2e7d32;font-size:16px;font-weight:600;">✓ Nothing overdue, due today, or due this week.</div>` : ""}
        ${bucketHtml("🔴 OVERDUE", "#c62828", overdue, totalOverdue)}
        ${bucketHtml("🟡 DUE TODAY", "#ef6c00", dueToday, totalDueToday)}
        ${bucketHtml("⏰ DUE NEXT 7 DAYS", "#1565c0", dueSoon, totalDueSoon)}
        ${bucketHtml("📋 UPCOMING (8–14 days)", "#37474f", upcoming, upcoming.reduce((s,e)=>s+Number(e.amount),0))}
        <div style="margin-top:24px;padding:14px 16px;background:#fff8e1;border-radius:6px;font-size:12px;color:#5d4037;">
          Open MONTHLY_PAYABLES → flip Status to "Paid" when each bill clears.
        </div>
      </div>
    </div>`;
}

function installEmailDigestTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "sendDailyEmailDigest") ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger("sendDailyEmailDigest").timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getUi().alert(`✓ Daily email digest installed (6 AM → ${PERSONAL_EMAIL}).`);
}

/* ============================================================
 * WEBHOOK — Claude/Cowork direct write to the Sheet
 * Deploy this script as a Web App, then POST JSON to its URL.
 * ============================================================ */
function doGet() {
  return jsonResponse({
    success: true,
    info: "Millionaire Engine webhook is live. POST {secret, action, data} to use it."
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: "no payload" });
    }
    const payload = JSON.parse(e.postData.contents);
    if (payload.secret !== WEBHOOK_SECRET) {
      return jsonResponse({ success: false, error: "invalid_secret" });
    }
    const action = payload.action;
    const data = payload.data || {};
    let result;
    switch (action) {
      case "ping":                        result = webhookPing(); break;
      case "add_master_log_row":          result = webhookAddMasterLogRow(data); break;
      case "update_master_log_status":    result = webhookUpdateMasterLogStatus(data); break;
      case "add_payable":                 result = webhookAddPayable(data); break;
      case "update_payable_status":       result = webhookUpdatePayableStatus(data); break;
      case "query":                       result = webhookQuery(data); break;
      case "month_summary":               result = webhookMonthSummary(); break;
      case "list_upcoming":               result = webhookListUpcoming(data); break;
      case "monthly_history":             result = webhookMonthlyHistory(data); break;
      case "category_breakdown":          result = webhookCategoryBreakdown(data); break;
      default: return jsonResponse({ success: false, error: "unknown_action: " + action });
    }
    return jsonResponse({ success: true, action: action, result: result });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString(), stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function webhookPing() {
  const ss = getSS();
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  return {
    pong: true,
    time: new Date().toISOString(),
    spreadsheet: ss.getName(),
    spreadsheet_id: ss.getId(),
    sheets: ss.getSheets().map(s => s.getName()),
    master_log_last_row: sheet ? sheet.getLastRow() : null,
    timezone: Session.getScriptTimeZone(),
    deployed_as: Session.getEffectiveUser().getEmail()
  };
}

function webhookAddMasterLogRow(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!sheet) throw new Error("MASTER_LOG sheet not found");
  const date = data.date ? new Date(data.date) : new Date();
  const item = data.item || "";
  const category = data.category || "";
  const amount = parseFloat(data.amount) || 0;
  const paymentMethod = data.payment_method || "";
  const status = data.status || "Pending";
  const notes = data.notes || "";
  if (!item) throw new Error("item is required");
  if (!amount) throw new Error("amount is required");
  const row = findNextEmptyRow(sheet);
  sheet.getRange(row, 1).setFormula(`=IF(B${row}="","",TEXT(B${row},"mmmm yyyy"))`);
  sheet.getRange(row, 2).setValue(date);
  sheet.getRange(row, 2).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(row, 3).setValue(item);
  sheet.getRange(row, 4).setValue(category);
  sheet.getRange(row, 5).setValue(amount);
  sheet.getRange(row, 5).setNumberFormat('"₱"#,##0.00');
  sheet.getRange(row, 6).setValue(paymentMethod);
  sheet.getRange(row, 7).setValue(status);
  sheet.getRange(row, 8).setValue(notes);
  sortMasterLog();
  return { row_inserted: row, item: item, amount: amount, sorted: true };
}

function webhookUpdateMasterLogStatus(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!sheet) throw new Error("MASTER_LOG sheet not found");
  const search = String(data.search || "").toLowerCase();
  const searchDate = data.date ? new Date(data.date) : null;
  const newStatus = data.status || "Paid";
  if (!search) throw new Error("data.search is required");
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 2, lastRow - 1, 7).getValues();
  let matched = -1;
  let oldStatus = "";
  for (let i = 0; i < rows.length; i++) {
    const item = String(rows[i][1] || "").toLowerCase();
    const dateVal = rows[i][0];
    if (item.includes(search)) {
      if (searchDate) {
        if (dateVal instanceof Date &&
            dateVal.getFullYear() === searchDate.getFullYear() &&
            dateVal.getMonth() === searchDate.getMonth() &&
            dateVal.getDate() === searchDate.getDate()) {
          matched = i + 2; oldStatus = rows[i][5]; break;
        }
      } else {
        matched = i + 2; oldStatus = rows[i][5]; break;
      }
    }
  }
  if (matched < 0) throw new Error("no matching row for: " + search);
  sheet.getRange(matched, 7).setValue(newStatus);
  return { matched_row: matched, old_status: oldStatus, new_status: newStatus };
}

function webhookAddPayable(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName(MONTHLY_PAYABLES_SHEET);
  if (!sheet) throw new Error("MONTHLY_PAYABLES sheet not found");
  let row = PAYABLES_FIRST_ROW;
  while (row <= PAYABLES_LAST_ROW && sheet.getRange(row, 2).getValue()) row++;
  if (row > PAYABLES_LAST_ROW) throw new Error("MONTHLY_PAYABLES data range full");
  sheet.getRange(row, 2).setValue(data.item || "");
  sheet.getRange(row, 3).setValue(data.provider || "");
  sheet.getRange(row, 4).setValue(parseInt(data.due_day) || 1);
  sheet.getRange(row, 5).setValue(parseFloat(data.amount) || 0);
  sheet.getRange(row, 5).setNumberFormat('"₱"#,##0.00');
  sheet.getRange(row, 6).setValue(data.status || "Pending");
  if (data.last_paid) {
    sheet.getRange(row, 7).setValue(new Date(data.last_paid));
    sheet.getRange(row, 7).setNumberFormat("yyyy-mm-dd");
  }
  sheet.getRange(row, 8).setValue(data.category || "");
  sheet.getRange(row, 9).setValue(data.auto_debit || "No");
  sheet.getRange(row, 10).setValue(data.notes || "");
  return { row: row };
}

function webhookUpdatePayableStatus(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName(MONTHLY_PAYABLES_SHEET);
  if (!sheet) throw new Error("MONTHLY_PAYABLES sheet not found");
  const search = String(data.item || "").toLowerCase();
  if (!search) throw new Error("data.item is required");
  for (let r = PAYABLES_FIRST_ROW; r <= PAYABLES_LAST_ROW; r++) {
    const item = String(sheet.getRange(r, 2).getValue() || "").toLowerCase();
    if (item.includes(search)) {
      const old = sheet.getRange(r, 6).getValue();
      sheet.getRange(r, 6).setValue(data.status || "Paid");
      if (data.last_paid) {
        sheet.getRange(r, 7).setValue(new Date(data.last_paid));
        sheet.getRange(r, 7).setNumberFormat("yyyy-mm-dd");
      }
      return { matched_row: r, old_status: old, new_status: data.status || "Paid" };
    }
  }
  throw new Error("no matching payable for: " + search);
}

function webhookQuery(data) {
  const ss = getSS();
  const sheet = ss.getSheetByName(data.sheet);
  if (!sheet) throw new Error("sheet not found: " + data.sheet);
  const range = data.range || "A1:H20";
  return { sheet: data.sheet, range: range, values: sheet.getRange(range).getValues() };
}

function webhookMonthSummary() {
  const ss = getSS();
  const ml = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!ml) throw new Error("MASTER_LOG not found");
  const now = new Date();
  const monthYear = Utilities.formatDate(now, Session.getScriptTimeZone(), "MMMM yyyy");
  const lastRow = ml.getLastRow();
  const data = ml.getRange(2, 1, lastRow - 1, 8).getValues();
  let incomeReceived = 0, expensesPaid = 0, expensesPending = 0, expensesOverdue = 0;
  let overdueItems = [];
  data.forEach(row => {
    const month = row[0];
    const date = row[1];
    const item = row[2];
    const category = row[3];
    const amount = parseFloat(row[4]) || 0;
    const status = row[6];
    if (month !== monthYear) return;
    if (status === "Paid" && category === "Money ++") incomeReceived += amount;
    else if (status === "Paid") expensesPaid += amount;
    else if (status === "Pending") expensesPending += amount;
    else if (status === "Unpaid" || status === "Priority") {
      expensesOverdue += amount;
      overdueItems.push({ item: item, amount: amount, status: status, date: date });
    }
  });
  return {
    month: monthYear,
    income_received: incomeReceived,
    expenses_paid: expensesPaid,
    expenses_pending: expensesPending,
    expenses_overdue: expensesOverdue,
    net_to_date: incomeReceived - expensesPaid,
    projected_net: incomeReceived - expensesPaid - expensesPending,
    overdue_items: overdueItems
  };
}

function webhookMonthlyHistory(data) {
  const ss = getSS();
  const ml = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!ml) throw new Error("MASTER_LOG not found");
  const months = parseInt(data.months) || 6;
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  // Build target month list: oldest first
  const targets = [];
  for (let i = months - 1; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const label = Utilities.formatDate(d, tz, "MMMM yyyy");
    const short = Utilities.formatDate(d, tz, "MMM");
    targets.push({ label: label, short: short, income: 0, expense_paid: 0, expense_pending: 0 });
  }
  const lookup = {};
  targets.forEach(t => { lookup[t.label] = t; });

  const lastRow = ml.getLastRow();
  if (lastRow < 2) return { months: targets };
  const rows = ml.getRange(2, 1, lastRow - 1, 8).getValues();
  rows.forEach(row => {
    const month = row[0];
    const category = row[3];
    const amount = parseFloat(row[4]) || 0;
    const status = row[6];
    const t = lookup[month];
    if (!t) return;
    if (status === "Paid" && category === "Money ++") t.income += amount;
    else if (status === "Paid") t.expense_paid += amount;
    else if (status === "Pending" || status === "Unpaid" || status === "Priority") t.expense_pending += amount;
  });
  return { months: targets };
}

function webhookCategoryBreakdown(data) {
  const ss = getSS();
  const ml = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!ml) throw new Error("MASTER_LOG not found");
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  let targetMonth = data.month;
  if (!targetMonth) targetMonth = Utilities.formatDate(now, tz, "MMMM yyyy");
  const statusFilter = data.status_filter || "Paid";  // "Paid", "Pending", or "All"

  const lastRow = ml.getLastRow();
  if (lastRow < 2) return { month: targetMonth, categories: [] };
  const rows = ml.getRange(2, 1, lastRow - 1, 8).getValues();
  const totals = {};
  rows.forEach(row => {
    const month = row[0];
    const category = row[3] || "(Uncategorized)";
    const amount = parseFloat(row[4]) || 0;
    const status = row[6];
    if (month !== targetMonth) return;
    if (category === "Money ++") return;  // exclude income
    if (statusFilter !== "All" && status !== statusFilter) return;
    totals[category] = (totals[category] || 0) + amount;
  });
  const categories = Object.entries(totals)
    .map(([name, amount]) => ({ name, amount }))
    .sort((a, b) => b.amount - a.amount);
  return { month: targetMonth, status_filter: statusFilter, categories: categories };
}

function webhookListUpcoming(data) {
  const ss = getSS();
  const ml = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!ml) throw new Error("MASTER_LOG not found");
  const days = parseInt(data.days) || 7;
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const cutoff = new Date(now); cutoff.setDate(cutoff.getDate() + days);
  const lastRow = ml.getLastRow();
  const rows = ml.getRange(2, 2, lastRow - 1, 7).getValues();
  const upcoming = [];
  rows.forEach((row, i) => {
    const d = row[0];
    const item = row[1];
    const category = row[2];
    const amount = parseFloat(row[3]) || 0;
    const status = row[5];
    if (!(d instanceof Date)) return;
    if (d < now || d > cutoff) return;
    if (status === "Paid") return;
    upcoming.push({
      row: i + 2,
      date: Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      item: item, category: category, amount: amount, status: status
    });
  });
  upcoming.sort((a, b) => a.date.localeCompare(b.date) || a.amount - b.amount);
  return { window_days: days, count: upcoming.length, items: upcoming };
}
