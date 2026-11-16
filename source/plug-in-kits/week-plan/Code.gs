/**
 * Lance Week Plan: Google Sheet + Apps Script
 * Built by Basav Sharma for the Lance Chief of Staff take-home assignment.
 *
 * The Sheet is the plan. This script keeps everything else in step with it:
 *   1. Blocks tab -> Google Calendar events (re-running updates them, never duplicates).
 *   2. Emails marked Ready -> Gmail drafts. A person presses Send.
 *   3. Slack messages marked Ready -> sent when their block starts (checked every 15 minutes).
 *   4. 5:15 PM on weekdays -> end-of-day note to the CEO, built from the Inbox tab.
 *   5. The dashboard (HTML file "Dashboard") is served from here, private to the company's Google domain.
 * Nothing goes out unless a person has set the row to Ready.
 *
 * Setup, about 15 minutes:
 *   - Upload Lance_Week_Plan_Template.xlsx to Google Drive and save it as a Google Sheet.
 *   - Extensions > Apps Script: paste this file as Code.gs, add an HTML file named Dashboard (the dashboard page).
 *   - Project Settings > Script properties: SLACK_TOKEN (Slack app token with chat:write; a user token posts as you)
 *     and CEO_SLACK_ID (the CEO's Slack member ID, for the end-of-day note).
 *   - Reload the Sheet, then Week plan > Turn on automation (approve Google's prompt once).
 *   - Deploy > New deployment > Web app (execute as me, access: anyone at the company) for the dashboard link.
 */

const TABS = ['Days', 'Inbox', 'Blocks', 'Messages', 'Background', 'Deadlines'];
const COLORS = { 'Inbox fire': 'ORANGE', 'Onboarding project': 'GREEN', 'Investor update': 'YELLOW', 'CEO + closing loops': 'GRAY', 'Team': 'MAUVE' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Week plan')
    .addItem('Put my blocks on Google Calendar', 'pushBlocksToCalendar')
    .addItem('Make Gmail drafts (emails marked Ready)', 'createEmailDrafts')
    .addItem('Send Slack messages that are due (marked Ready)', 'sendDueSlackMessages')
    .addItem('Send end-of-day note to CEO', 'sendEodNote')
    .addSeparator()
    .addItem('Turn on automation', 'installTriggers')
    .addToUi();
}

// ---------- 1. Calendar ----------
function pushBlocksToCalendar() {
  const cal = CalendarApp.getDefaultCalendar();
  let added = 0, updated = 0;
  readTab_('Blocks').forEach(b => {
    const start = at_(b.date, b.start), end = at_(b.date, b.end);
    const desc = String(b.steps || '') + (b.thinking ? '\n\nWhy: ' + b.thinking : '');
    let ev = b.calendar_event_id ? cal.getEventById(b.calendar_event_id) : null;
    if (ev) {
      ev.setTime(start, end); ev.setTitle(b.title); ev.setDescription(desc); updated++;
    } else {
      ev = cal.createEvent(b.title, start, end, { description: desc });
      setCell_('Blocks', b._row, 'calendar_event_id', ev.getId()); added++;
    }
    if (COLORS[b.workstream]) ev.setColor(CalendarApp.EventColor[COLORS[b.workstream]]);
  });
  toast_(`${added} blocks added and ${updated} updated on your calendar.`);
  return { added: added, updated: updated };
}

// ---------- 2. Gmail drafts ----------
function createEmailDrafts() {
  let n = 0;
  readTab_('Messages').forEach(m => {
    if (m.channel !== 'Email' || m.status !== 'Ready') return;
    if (!/@/.test(m.send_to)) { mark_(m, 'Needs email address'); return; }
    GmailApp.createDraft(String(m.send_to), m.subject || '(no subject)', String(m.text));
    mark_(m, 'Drafted');
    n++;
  });
  toast_(n + ' Gmail drafts ready to review and send.');
  return n;
}

// ---------- 3. Slack, when the block starts ----------
function sendDueSlackMessages() {
  const token = prop_('SLACK_TOKEN');
  if (!token) throw new Error('Add SLACK_TOKEN in Project Settings > Script properties.');
  const blocks = byKey_(readTab_('Blocks'), 'block_id'), now = new Date();
  let n = 0;
  readTab_('Messages').forEach(m => {
    if (!/^Slack/.test(m.channel) || m.status !== 'Ready') return;
    const b = blocks[m.block_id];
    if (!b || at_(b.date, b.start) > now) return;
    if (!m.send_to) { mark_(m, 'Needs Slack ID'); return; }
    const errors = String(m.send_to).split(/[,\s]+/).filter(Boolean)
      .map(id => slack_(token, id, String(m.text))).filter(r => !r.ok).map(r => r.error);
    mark_(m, errors.length ? 'Error: ' + errors.join(', ') : 'Sent');
    if (!errors.length) n++;
  });
  return n;
}

// ---------- 4. End-of-day note to the CEO ----------
function sendEodNote() {
  const now = new Date(), today = startOfDay_(now);
  if (now.getDay() === 0 || now.getDay() === 6) return '';
  const rows = readTab_('Inbox');
  const line = r => '• ' + r.item + (r.resolution ? ': ' + r.resolution : '');
  const sections = [
    ['Done today', rows.filter(r => r.status === 'Done' && r.closed_on && startOfDay_(asDate_(r.closed_on)).getTime() === today.getTime())],
    ['At risk', rows.filter(r => r.status === 'At risk')],
    ['Need from you', rows.filter(r => r.status === 'Need CEO')],
  ];
  const parts = [`*End of day, ${fmt_(now, 'EEE MMM d')}*`];
  sections.forEach(([label, list]) => { if (list.length) parts.push(`*${label}*\n` + list.map(line).join('\n')); });
  parts.push(`Still open: ${rows.filter(r => r.status === 'Open').length}. Full plan: ${SpreadsheetApp.getActive().getUrl()}`);
  const text = parts.join('\n\n');
  const token = prop_('SLACK_TOKEN'), ceo = prop_('CEO_SLACK_ID');
  if (!token || !ceo) { Logger.log(text); toast_('Add SLACK_TOKEN and CEO_SLACK_ID to send. Note written to the log.'); return text; }
  const r = slack_(token, ceo, text);
  if (!r.ok) throw new Error('Slack said: ' + r.error);
  return text;
}

// ---------- Automation ----------
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['sendDueSlackMessages', 'sendEodNote'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDueSlackMessages').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('sendEodNote').timeBased().everyDays(1).atHour(17).nearMinute(15).create();
  toast_('Automation is on: Ready Slack messages every 15 min, CEO note at 5:15 PM.');
}

// ---------- Dashboard (web app) ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Week plan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called by the dashboard: every tab as plain rows (dates become text). */
function getData() {
  const out = {};
  TABS.forEach(n => { out[n] = readTab_(n).map(plain_); });
  return out;
}

/** Called by the dashboard's Approve / Pull back buttons. Sending stays with the script. */
function setMessageStatus(row, status) {
  if (['Draft', 'Ready'].indexOf(status) < 0) throw new Error('The dashboard can only set Draft or Ready.');
  const m = readTab_('Messages').filter(r => r._row === Number(row))[0];
  if (!m) throw new Error('No message on row ' + row);
  if (['Sent', 'Drafted'].indexOf(m.status) >= 0) throw new Error('Already ' + m.status + '.');
  setCell_('Messages', m._row, 'status', status);
  return true;
}

// ---------- Helpers ----------
function slack_(token, channel, text) {
  const res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ channel: channel, text: text }),
    muteHttpExceptions: true,
  });
  return JSON.parse(res.getContentText());
}
function mark_(m, status) {
  setCell_('Messages', m._row, 'status', status);
  if (status === 'Sent' || status === 'Drafted') setCell_('Messages', m._row, 'sent_at', new Date());
}
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing tab: ' + name);
  return sh;
}
/** A tab as objects keyed by the header row; _row is the Sheet row number. */
function readTab_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const head = values.shift().map(h => String(h).trim());
  return values.map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => { if (h) o[h] = r[j]; });
    return o;
  }).filter(o => head.some(h => h && o[h] !== ''));
}
function setCell_(name, row, col, value) {
  const sh = sheet_(name);
  const j = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(col);
  if (j < 0) throw new Error(`Tab ${name} has no column ${col}`);
  sh.getRange(row, j + 1).setValue(value);
}
function byKey_(rows, key) { const o = {}; rows.forEach(r => { o[r[key]] = r; }); return o; }
function prop_(key) { return PropertiesService.getScriptProperties().getProperty(key); }
function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg); } catch (e) { Logger.log(msg); } }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function fmt_(d, pattern) { return Utilities.formatDate(d, tz_(), pattern); }
function asDate_(v) {
  if (v instanceof Date) return new Date(v.getTime());
  const m = String(v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m) return new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  const d = new Date(v);
  if (isNaN(d.getTime())) throw new Error('Not a date: ' + v);
  return d;
}
function startOfDay_(d) { const x = new Date(d.getTime()); x.setHours(0, 0, 0, 0); return x; }
/** A date cell plus "9:15" (or a Sheets time cell) = one moment. */
function at_(date, t) {
  const s = t instanceof Date ? fmt_(t, 'H:mm') : String(t);
  const m = s.match(/(\d{1,2}):(\d{2})\s*(am|pm)?/i);
  if (!m) throw new Error('Not a time: ' + t);
  let h = +m[1];
  if (m[3] && /pm/i.test(m[3]) && h < 12) h += 12;
  if (m[3] && /am/i.test(m[3]) && h === 12) h = 0;
  const x = startOfDay_(asDate_(date));
  x.setHours(h, +m[2], 0, 0);
  return x;
}
/** Dates can't travel to the browser, so turn them into text first. */
function plain_(row) {
  const o = {};
  Object.keys(row).forEach(k => {
    const v = row[k];
    if (!(v instanceof Date)) { o[k] = v; return; }
    if (v.getFullYear() < 1900) o[k] = fmt_(v, 'H:mm');
    else o[k] = fmt_(v, 'HH:mm') === '00:00' ? fmt_(v, 'yyyy-MM-dd') : fmt_(v, 'yyyy-MM-dd HH:mm');
  });
  return o;
}
