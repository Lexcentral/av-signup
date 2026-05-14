// AV Signup — Google Apps Script backend
//
// SETUP (one time):
//   1. Run createEmailTrigger() from the editor to schedule weekly emails.
//   2. Deploy as web app: Execute as Me, Who has access: Anyone (even anonymous).
//   3. Copy the deployed URL into index.html as GAS_URL.

const SIGNUP_SHEET_ID  = '1itsYgObIh34c8Lp1XgWBvKSASCm_hZ2iib9Q-H4gwko';
const WORSHIP_SHEET_ID = '19CUEXrS4gP7O7dSqInmbkUzWpJV0W_s9HLlQtDyqC_c';
const SITE_URL         = 'https://lexcentral.github.io/av-signup/';
const WEEKS_AHEAD      = 4;

// ── Web app ────────────────────────────────────────────────────────────────

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  let result;
  try {
    if (action === 'schedule') {
      result = getScheduleData();
    } else {
      result = { error: 'Unknown action. Use ?action=schedule' };
    }
  } catch (err) {
    result = { error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'signup') {
      result = addSignup(data);
    } else {
      result = { ok: false, error: 'Unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Schedule data ──────────────────────────────────────────────────────────

function getScheduleData() {
  const ss     = SpreadsheetApp.openById(SIGNUP_SHEET_ID);
  const signups = sheetRows(ss, 'Signup');
  const jobs    = sheetRows(ss, 'Jobs').map(r => String(r.job)).filter(Boolean);
  const dates   = upcomingServiceDates();

  const slots = [];
  dates.forEach(({ iso, label }) => {
    jobs.forEach(job => {
      const id       = `${iso}_${job}`;
      const existing = signups.find(s => String(s.id) === id);
      slots.push({ id, iso, label, job, volunteer: existing ? String(existing.volunteer) : '' });
    });
  });

  return { slots };
}

function upcomingServiceDates() {
  const ws      = SpreadsheetApp.openById(WORSHIP_SHEET_ID);
  const sheet   = ws.getSheetByName('planning');
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  const dateIdx = headers.indexOf('date');
  if (dateIdx < 0) throw new Error('No "date" column in planning sheet');

  const tz     = Session.getScriptTimeZone();
  const today  = new Date(); today.setHours(0, 0, 0, 0);
  const cutoff = new Date(today.getTime() + WEEKS_AHEAD * 7 * 86400000);

  const seen  = new Set();
  const dates = [];

  for (let i = 1; i < data.length; i++) {
    const raw = data[i][dateIdx];
    if (!raw) continue;
    const d = raw instanceof Date ? new Date(raw) : new Date(String(raw));
    if (isNaN(d.getTime())) continue;
    d.setHours(0, 0, 0, 0);
    if (d < today || d > cutoff) continue;
    const iso = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    if (seen.has(iso)) continue;
    seen.add(iso);
    dates.push({ iso, label: Utilities.formatDate(d, tz, 'EEE, MMM d, yyyy') });
  }

  dates.sort((a, b) => a.iso < b.iso ? -1 : 1);
  return dates;
}

// ── Signup write ───────────────────────────────────────────────────────────

function addSignup({ id, date, job, name }) {
  if (!id || !date || !job || !name || !String(name).trim()) {
    return { ok: false, error: 'Missing required fields.' };
  }
  const ss      = SpreadsheetApp.openById(SIGNUP_SHEET_ID);
  const sheet   = ss.getSheetByName('Signup');
  const signups = sheetRows(ss, 'Signup');
  const existing = signups.find(s => String(s.id) === String(id));
  if (existing && existing.volunteer) {
    return { ok: false, error: `${existing.volunteer} is already signed up for this slot.` };
  }
  sheet.appendRow([id, date, job, String(name).trim()]);
  sheet.sort(1);
  return { ok: true };
}

// ── Email ──────────────────────────────────────────────────────────────────
// Run createEmailTrigger() once from the editor to schedule Tuesday + Friday sends.

// Run from the editor to preview the email as yourself.
function sendTestEmail() {
  const { slots } = getScheduleData();
  MailApp.sendEmail({
    to: 'aaron@aaronaustin.com',
    subject: '[TEST] AV Signups',
    htmlBody: buildEmailHtml('Aaron Austin', slots),
  });
}

function createEmailTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendSignupEmails')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendSignupEmails').timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(9).create();
  ScriptApp.newTrigger('sendSignupEmails').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(9).create();
}

function sendSignupEmails() {
  const ss     = SpreadsheetApp.openById(SIGNUP_SHEET_ID);
  const people = sheetRows(ss, 'People');
  const { slots } = getScheduleData();
  people.forEach(({ name, email }) => {
    if (!name || !email) return;
    MailApp.sendEmail({
      to: String(email).trim(),
      subject: 'AV Signups',
      htmlBody: buildEmailHtml(String(name).trim(), slots),
    });
  });
}

function buildEmailHtml(recipientName, slots) {
  const byDate = {};
  slots.forEach(slot => {
    if (!byDate[slot.iso]) byDate[slot.iso] = { label: slot.label, slots: [] };
    byDate[slot.iso].slots.push(slot);
  });

  const blocks = Object.values(byDate).map(({ label, slots: dateSlots }) => {
    const rows = dateSlots.map(slot => {
      const filled  = !!slot.volunteer;
      const pageUrl = `${SITE_URL}?name=${encodeURIComponent(recipientName)}&id=${encodeURIComponent(slot.id)}`;
      const status  = filled
        ? `<span style="color:#555">${slot.volunteer}</span>`
        : `<span style="color:#0fa481;font-weight:bold">Open</span>`;
      const btn = filled
        ? `<span style="display:inline-block;padding:6px 16px;background:#ccc;color:#fff;border-radius:4px;font-size:14px">Filled</span>`
        : `<a href="${pageUrl}" style="display:inline-block;padding:6px 16px;background:#068ccd;color:#fff;border-radius:4px;text-decoration:none;font-size:14px">Sign Up</a>`;
      return `<tr>
        <td style="padding:8px 12px;font-size:15px;border-bottom:1px solid #eee"><strong>${slot.job.toUpperCase()}</strong> — ${status}</td>
        <td style="padding:8px 12px;text-align:right;border-bottom:1px solid #eee">${btn}</td>
      </tr>`;
    }).join('');

    return `<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;border-collapse:collapse;font-family:Arial,sans-serif;border:1px solid #ddd;border-radius:6px">
      <tr><td colspan="2" style="background:#f0f4f8;padding:10px 12px;font-size:16px;font-weight:bold">${label}</td></tr>
      ${rows}
    </table>`;
  }).join('');

  return `<!DOCTYPE html>
<html><body style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;padding:20px;background:#fff">
<h2 style="color:#068ccd;margin-bottom:4px">AV Signups</h2>
<p style="color:#555;margin-top:0">Hi ${recipientName}, here are the upcoming services that need AV volunteers.
Click <strong>Sign Up</strong> to claim an open slot.</p>
${blocks}
<p style="margin-top:24px;font-size:13px;color:#999">
  <a href="${SITE_URL}?name=${encodeURIComponent(recipientName)}" style="color:#068ccd">View full schedule online</a>
</p>
</body></html>`;
}

// ── Utilities ──────────────────────────────────────────────────────────────

function sheetRows(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[String(h)] = row[i]; });
    return obj;
  });
}
