/**
 * Lance Onboarding Hub: Google Sheet + Apps Script
 * Built by Basav Sharma for the Lance Chief of Staff take-home assignment.
 *
 * The Sheet is the database. This script does the repeat work:
 *   1. Every new hire row becomes a full checklist (Tasks tab), due dates set from the start date.
 *   2. 8 AM on weekdays: one Slack message to #onboarding with what is overdue or due by tomorrow, by owner.
 *   3. 3 days before a start date: a welcome email waits in Gmail drafts. A person presses Send.
 *   4. The week before a start date: Week 1 meetings go on the calendar.
 *   5. The dashboard (HTML file "Dashboard") is served from here, private to the company's Google domain.
 *
 * Setup, about 15 minutes:
 *   - Upload Lance_Onboarding_Template.xlsx to Google Drive and save it as a Google Sheet.
 *   - Extensions > Apps Script: paste this file as Code.gs, add an HTML file named Dashboard (the dashboard page).
 *   - Project Settings > Script properties: SLACK_WEBHOOK_URL = incoming webhook for #onboarding.
 *   - Reload the Sheet, then Onboarding > Turn on daily automation (approve Google's prompt once).
 *   - Deploy > New deployment > Web app (execute as me, access: anyone at the company) for the dashboard link.
 */

const TABS = ['Settings', 'Hires', 'Checklist', 'Tasks', 'Owners', 'Access', 'Week1'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Onboarding')
    .addItem('Create tasks for new hires', 'createTasks')
    .addItem('Send Slack digest now', 'sendSlackDigest')
    .addItem('Draft welcome emails', 'draftWelcomeEmails')
    .addItem('Add Week 1 to calendars', 'addWeekOneEvents')
    .addSeparator()
    .addItem('Turn on daily automation', 'installTriggers')
    .addToUi();
}

/** Runs at 8 AM every day once automation is on; does nothing on weekends. */
function dailyRun() {
  const day = asOf_().getDay();
  if (day === 0 || day === 6) return;
  createTasks();
  sendSlackDigest();
  draftWelcomeEmails();
  addWeekOneEvents();
}

// ---------- 1. Hires x Checklist = Tasks ----------
/** Adds missing tasks for every hire, and moves due dates if a start date changed. Safe to run any time. */
function createTasks() {
  const tasks = readTab_('Tasks'), list = readTab_('Checklist');
  const have = {};
  tasks.forEach(t => { have[t.hire_id + '|' + t.task_id] = t; });
  const rows = [];
  let moved = 0;
  readTab_('Hires').forEach(h => {
    if (!h.hire_id || !h.start_date) return;
    list.filter(c => appliesTo_(c, h)).forEach(c => {
      const due = dueDate_(h.start_date, c.days_from_start), t = have[h.hire_id + '|' + c.task_id];
      if (!t) { rows.push([h.hire_id, c.task_id, due, 'Open', '', '']); return; }
      if (!t.due_date || startOfDay_(asDate_(t.due_date)).getTime() !== due.getTime()) { setCell_('Tasks', t._row, 'due_date', due); moved++; }
    });
  });
  appendRows_('Tasks', rows);
  toast_(rows.length + ' tasks created, ' + moved + ' due dates moved.');
  return rows.length;
}

// ---------- 2. Daily Slack digest ----------
function sendSlackDigest() {
  const now = asOf_(), today = startOfDay_(now), tomorrow = addDays_(today, 1);
  const hires = byKey_(readTab_('Hires'), 'hire_id');
  const list = byKey_(readTab_('Checklist'), 'task_id');
  const owners = byKey_(readTab_('Owners'), 'owner');
  const tasks = readTab_('Tasks');
  const groups = {};
  tasks.forEach(t => {
    const h = hires[t.hire_id], c = list[t.task_id];
    if (!h || !c || t.status === 'Done') return;
    const due = dueDate_(h.start_date, c.days_from_start); // always from today's start date
    if (due > tomorrow) return;
    const who = ownerName_(c.owner, h);
    const late = due < today ? ':red_circle: overdue · ' : '';
    (groups[who] = groups[who] || []).push(`• ${late}${c.task} for ${h.name} (${fmt_(due, 'EEE MMM d')})`);
  });
  const lines = [`*Onboarding, ${fmt_(now, 'EEE MMM d')}*`];
  Object.keys(groups).sort().forEach(who => {
    const o = owners[who];
    lines.push((o && o.slack_id ? `<@${o.slack_id}>` : `*${who}*`) + '\n' + groups[who].join('\n'));
  });
  if (lines.length === 1) lines.push('Nothing overdue or due by tomorrow. :white_check_mark:');
  // who starts in the next 2 days, and are they ready?
  Object.keys(hires).forEach(id => {
    const h = hires[id], days = Math.round((startOfDay_(asDate_(h.start_date)) - today) / 864e5);
    if (days < 0 || days > 2) return;
    const open = tasks.filter(t => t.hire_id === id && t.status !== 'Done' && Number((list[t.task_id] || {}).days_from_start) < 0).length;
    lines.push(`${h.name} starts ${fmt_(asDate_(h.start_date), 'EEE')}: ` + (open ? `:warning: ${open} prep tasks open` : 'ready for Day 1 :white_check_mark:'));
  });
  const text = lines.join('\n\n');
  const url = prop_('SLACK_WEBHOOK_URL');
  if (!url) { Logger.log(text); toast_('No SLACK_WEBHOOK_URL yet. Digest written to the log.'); return text; }
  UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: text }) });
  return text;
}

// ---------- 3. Welcome emails (drafts only) ----------
function draftWelcomeEmails() {
  const s = settings_(), before = Number(s.welcome_days_before || 3), today = startOfDay_(asOf_());
  const welcome = readTab_('Checklist').filter(c => /welcome email/i.test(c.task))[0];
  const tasks = readTab_('Tasks');
  let n = 0;
  readTab_('Hires').forEach(h => {
    const start = startOfDay_(asDate_(h.start_date)), days = Math.round((start - today) / 864e5);
    if (!h.email || days < 0 || days > before) return;
    const t = welcome && tasks.filter(r => r.hire_id === h.hire_id && r.task_id === welcome.task_id)[0];
    if (t && (t.status === 'Done' || /draft/i.test(t.note))) return;
    const first = String(h.name).split(' ')[0];
    const body = [
      `Hi ${first},`, '',
      `We're excited you start on ${fmt_(start, 'EEEE, MMMM d')}! A few things before Day 1:`,
      `1. Arrive at 9:30 AM at ${s.office_address || '[office address]'} and ask for ${h.buddy}, your onboarding buddy.`,
      '2. Your laptop will be ready, with your accounts set up. Sign in to Google first.',
      '3. Your first week is on your calendar, starting with lunch with the team.', '',
      'If anything comes up before then, just reply here.', '',
      'See you soon,', s.sender_name || '',
    ].join('\n');
    GmailApp.createDraft(h.email, 'Welcome to Lance: your first day', body);
    if (t) setCell_('Tasks', t._row, 'note', 'Draft in Gmail ' + fmt_(new Date(), 'MMM d'));
    n++;
  });
  toast_(n + ' welcome emails waiting in Gmail drafts.');
  return n;
}

// ---------- 4. Week 1 on the calendar ----------
function addWeekOneEvents() {
  const s = settings_(), cal = CalendarApp.getDefaultCalendar(), today = startOfDay_(asOf_());
  const invites = readTab_('Checklist').filter(c => /invites/i.test(c.task))[0];
  const tasks = readTab_('Tasks'), week1 = readTab_('Week1');
  let n = 0;
  readTab_('Hires').forEach(h => {
    const start = startOfDay_(asDate_(h.start_date)), days = Math.round((start - today) / 864e5);
    if (!h.email || days < 0 || days > 7) return;
    const t = invites && tasks.filter(r => r.hire_id === h.hire_id && r.task_id === invites.task_id)[0];
    if (t && /invites created/i.test(t.note)) return;
    week1.forEach(w => {
      const st = atTime_(addWorkdays_(start, Number(w.day) - 1), w.time);
      const en = new Date(st.getTime() + Number(w.minutes) * 60000);
      cal.createEvent(`${w.title} · ${h.name}`, st, en, { guests: h.email, sendInvites: s.send_invites === 'Yes' });
      n++;
    });
    if (t) setCell_('Tasks', t._row, 'note', 'Invites created ' + fmt_(new Date(), 'MMM d'));
  });
  toast_(n + ' Week 1 events added.');
  return n;
}

// ---------- Automation ----------
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'dailyRun') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyRun').timeBased().everyDays(1).atHour(8).create();
  toast_('Daily automation is on: 8 AM on weekdays.');
}

/** Stamps done_on when someone sets a task to Done in the Sheet. */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Tasks' || e.range.getRow() === 1) return;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (head[e.range.getColumn() - 1] !== 'status') return;
  sh.getRange(e.range.getRow(), head.indexOf('done_on') + 1).setValue(e.value === 'Done' ? new Date() : '');
}

// ---------- Dashboard (web app) ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Onboarding Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Called by the dashboard: every tab as plain rows (dates become text). */
function getData() {
  const out = {};
  TABS.forEach(n => { out[n] = readTab_(n).map(plain_); });
  return out;
}

/** Called by the dashboard when someone ticks a task. */
function setTaskStatus(hireId, taskId, done) {
  const t = readTab_('Tasks').filter(r => r.hire_id === hireId && r.task_id === taskId)[0];
  if (!t) throw new Error(`No task ${taskId} for ${hireId}. Run "Create tasks for new hires" first.`);
  setCell_('Tasks', t._row, 'status', done ? 'Done' : 'Open');
  setCell_('Tasks', t._row, 'done_on', done ? startOfDay_(new Date()) : '');
  return true;
}

// ---------- Helpers ----------
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
function appendRows_(name, rows) {
  if (!rows.length) return;
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
function setCell_(name, row, col, value) {
  const sh = sheet_(name);
  const j = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(col);
  if (j < 0) throw new Error(`Tab ${name} has no column ${col}`);
  sh.getRange(row, j + 1).setValue(value);
}
function byKey_(rows, key) { const o = {}; rows.forEach(r => { o[r[key]] = r; }); return o; }
function settings_() { const o = {}; readTab_('Settings').forEach(r => { o[r.key] = r.value; }); return o; }
function prop_(key) { return PropertiesService.getScriptProperties().getProperty(key); }
function toast_(msg) { try { SpreadsheetApp.getActive().toast(msg); } catch (e) { Logger.log(msg); } }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function fmt_(d, pattern) { return Utilities.formatDate(d, tz_(), pattern); }

function appliesTo_(c, h) {
  const a = String(c.applies_to || 'All');
  return a === 'All' || a.split(',').map(x => x.trim()).indexOf(h.team) >= 0;
}
/** Start date + offset; never on a weekend (prep moves back to Friday, follow-ups forward to Monday). */
function dueDate_(start, offset) {
  const d = startOfDay_(asDate_(start)), n = Number(offset);
  d.setDate(d.getDate() + n);
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + (n < 0 ? -1 : 1));
  return d;
}
/** Manager and Buddy are per hire; everyone else is a named owner. */
function ownerName_(owner, h) { return owner === 'Manager' ? h.manager : owner === 'Buddy' ? h.buddy : owner; }
/** The demo pins "now" in Settings > as_of; blank means the real time. */
function asOf_() { const v = settings_().as_of; return v ? asDate_(v) : new Date(); }
function asDate_(v) {
  if (v instanceof Date) return new Date(v.getTime());
  const m = String(v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (m) return new Date(+m[1], m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
  const d = new Date(v);
  if (isNaN(d.getTime())) throw new Error('Not a date: ' + v);
  return d;
}
function startOfDay_(d) { const x = new Date(d.getTime()); x.setHours(0, 0, 0, 0); return x; }
function addDays_(d, n) { const x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function addWorkdays_(d, n) {
  const x = new Date(d.getTime());
  while (n > 0) { x.setDate(x.getDate() + 1); if (x.getDay() !== 0 && x.getDay() !== 6) n--; }
  return x;
}
/** "9:30", "14:00" or a Sheets time cell, on a given day. */
function atTime_(day, t) {
  const s = t instanceof Date ? fmt_(t, 'H:mm') : String(t);
  const m = s.match(/(\d{1,2}):(\d{2})\s*(am|pm)?/i);
  if (!m) throw new Error('Not a time: ' + t);
  let h = +m[1];
  if (m[3] && /pm/i.test(m[3]) && h < 12) h += 12;
  if (m[3] && /am/i.test(m[3]) && h === 12) h = 0;
  const x = new Date(day.getTime());
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
