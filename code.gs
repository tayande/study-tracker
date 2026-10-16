/**
 * Study Tracker: a Google Apps Script web app.
 *
 * Your data lives in the tabs of this Google Sheet:
 *   Settings   your name and start date
 *   Schedule   your daily time slots (weekday and weekend)
 *   Plan       what to study in each slot, by day number
 *   Resources  links shown on the Timetable page
 *   Log        what you did (written by the app; you can read it any time)
 *
 * To personalise the tracker, edit those tabs. You don't need to change this code.
 */

var TAB = {
  start: 'Start Here',
  settings: 'Settings',
  schedule: 'Schedule',
  plan: 'Plan',
  resources: 'Resources',
  log: 'Log'
};
var LOG_HEADERS = ['Date', 'Key', 'Activity', 'Done', 'Minutes', 'Note', 'Updated'];
var REFLECTION_KEY = '_daynote';

/* ------------------------------------------------------------------ */
/* Web app entry points                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var t = HtmlService.createTemplateFromFile('Index');
  t.initialDate = /^\d{4}-\d{2}-\d{2}$/.test(p.date || '') ? p.date : '';
  t.initialView = p.view === 'timetable' ? 'timetable' : 'today';
  return t.evaluate()
    .setTitle('Study Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Everything the page needs, in one call. */
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(TAB.settings)) setup();
  var tz = ss.getSpreadsheetTimeZone();

  var settings = {};
  rows_(ss, TAB.settings).forEach(function (r) {
    var k = String(r[0] || '').trim();
    if (k) settings[k] = r[1];
  });
  var start = toIsoDate_(settings['Start date'], tz);

  var schedule = { weekday: [], weekend: [] };
  var seen = { weekday: {}, weekend: {} };
  rows_(ss, TAB.schedule).forEach(function (r) {
    var type = String(r[0] || '').trim().toLowerCase();
    if (type !== 'weekday' && type !== 'weekend') return;
    var activity = String(r[2] || '').trim();
    if (!activity) return;
    var key = slug_(r[5] || activity);
    var n = 2, base = key;
    while (seen[type][key]) key = base + '-' + (n++);
    seen[type][key] = true;
    schedule[type].push({
      key: key,
      time: String(r[1] || '').trim(),
      name: activity,
      cat: String(r[3] || 'Other').trim() || 'Other',
      min: Number(r[4]) || 0
    });
  });

  var plan = {};
  rows_(ss, TAB.plan).forEach(function (r) {
    var day = parseInt(r[0], 10);
    var key = slug_(r[1]);
    var task = String(r[2] || '').trim();
    if (!day || !key || !task) return;
    plan[day] = plan[day] || {};
    (plan[day][key] = plan[day][key] || []).push({ t: task, u: String(r[3] || '').trim() });
  });

  var resources = rows_(ss, TAB.resources)
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) {
      return { name: String(r[0]).trim(), url: String(r[1] || '').trim(), note: String(r[2] || '').trim(), group: String(r[3] || '').trim() };
    });

  var log = {};
  rows_(ss, TAB.log).forEach(function (r) {
    var date = toIsoDate_(r[0], tz);
    var key = String(r[1] || '').trim();
    if (!date || !key) return;
    var d = log[date] = log[date] || { blocks: {}, reflection: '' };
    if (key === REFLECTION_KEY) { d.reflection = String(r[5] || ''); return; }
    d.blocks[key] = {
      done: r[3] === true || String(r[3]).toUpperCase() === 'TRUE',
      minutes: r[4] === '' || r[4] === null ? null : Number(r[4]),
      note: String(r[5] || '')
    };
  });

  return {
    name: String(settings['Your name'] || '').trim(),
    start: start,
    schedule: schedule,
    plan: plan,
    resources: resources,
    log: log
  };
}

/**
 * Save one day. payload = {blocks: {key: {done, minutes, note, activity}}, reflection}
 * Rows are matched on Date + Key, updated in place, or added at the end.
 */
function saveDay(date, payload) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error('Bad date: ' + date);
  payload = payload || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(TAB.log) || makeLogSheet_(ss);
    var tz = ss.getSpreadsheetTimeZone();
    var last = sh.getLastRow();
    var index = {};
    if (last >= 2) {
      var vals = sh.getRange(2, 1, last - 1, 2).getValues();
      for (var i = 0; i < vals.length; i++) {
        index[toIsoDate_(vals[i][0], tz) + '|' + String(vals[i][1]).trim()] = i + 2;
      }
    }
    var now = new Date();
    var updates = [];
    var blocks = payload.blocks || {};
    Object.keys(blocks).forEach(function (key) {
      var b = blocks[key] || {};
      var minutes = (b.minutes === null || b.minutes === undefined || b.minutes === '') ? '' : Math.max(0, Math.min(900, Number(b.minutes) || 0));
      updates.push([key, [date, key, String(b.activity || ''), !!b.done, minutes, String(b.note || '').slice(0, 1000), now]]);
    });
    if (typeof payload.reflection === 'string') {
      updates.push([REFLECTION_KEY, [date, REFLECTION_KEY, 'End-of-day note', '', '', payload.reflection.slice(0, 5000), now]]);
    }
    var appends = [];
    updates.forEach(function (u) {
      var row = index[date + '|' + u[0]];
      if (row) sh.getRange(row, 1, 1, LOG_HEADERS.length).setValues([u[1]]);
      else appends.push(u[1]);
    });
    if (appends.length) {
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, appends.length, 1).setNumberFormat('@');
      sh.getRange(startRow, 1, appends.length, LOG_HEADERS.length).setValues(appends);
    }
    return { ok: true, savedAt: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ------------------------------------------------------------------ */
/* Sheet menu                                                          */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Study Tracker')
    .addItem('Show my tracker link', 'showLink')
    .addItem('Create missing tabs', 'setup')
    .addToUi();
}

function showLink() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('No link yet', 'Deploy the tracker first: in Apps Script, click Deploy > New deployment > Web app. See the Start Here tab.', ui.ButtonSet.OK);
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial;font-size:14px">Your tracker link:</p>' +
    '<p style="font-family:Arial;font-size:14px"><a href="' + url + '" target="_blank">' + url + '</a></p>' +
    '<p style="font-family:Arial;font-size:13px;color:#555">Bookmark it, or add it to your phone\'s home screen.</p>'
  ).setWidth(520).setHeight(170);
  ui.showModalDialog(html, 'Study Tracker');
}

/* ------------------------------------------------------------------ */
/* One-time setup: creates and fills the tabs (never overwrites)        */
/* ------------------------------------------------------------------ */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var GH = 'https://github.com/rohitg00/ai-engineering-from-scratch/blob/main/phases/';
  var KG = 'https://www.kaggle.com/learn/python';

  makeTab_(ss, TAB.start, null, [
    ['Study Tracker'],
    [''],
    ['Your tracker is a web page that reads and writes this sheet. Personalise it by editing the tabs; no code changes needed.'],
    [''],
    ['TABS'],
    ['Settings: your name and the date of Day 1.'],
    ['Schedule: your time slots. Day type is Weekday or Weekend. Key is a short id the Plan tab uses; keep it unique within each day type.'],
    ['Plan: what to study. Day 1 is your start date. Key matches a slot Key in Schedule. Add rows for Day 8, 9, ... as you go.'],
    ['Resources: links shown on the Timetable page.'],
    ['Log: written by the app. Read it any time; avoid editing the Date and Key columns.'],
    [''],
    ['FIRST-TIME SETUP (about 10 minutes)'],
    ['1. In this sheet, open Extensions > Apps Script. You should see Code.gs and Index.html. (If you copied this sheet from someone, they are already there.)'],
    ['2. Click Deploy > New deployment. Click the gear icon next to "Select type" and choose Web app.'],
    ['3. Description: Study Tracker. Execute as: Me. Who has access: Anyone. Click Deploy.'],
    ['4. Click Authorize access, choose your Google account. If you see "Google hasn\'t verified this app", click Advanced > Go to Study Tracker (unsafe). This is your own script, so it is safe. Click Allow.'],
    ['5. Copy the Web app URL. That is your tracker link. Bookmark it or add it to your phone\'s home screen.'],
    ['6. Back in this sheet, reload the page. A Study Tracker menu appears; Show my tracker link finds your link again.'],
    [''],
    ['AFTER CHANGING THE CODE'],
    ['Deploy > Manage deployments > pencil icon > Version: New version > Deploy. Your link stays the same.'],
    ['Changes to Settings, Schedule, Plan and Resources need no redeploy: just reload the tracker page.'],
    [''],
    ['SHARING WITH SOMEONE'],
    ['Share this sheet with them as Viewer. They choose File > Make a copy, edit Settings and Schedule in their copy, then follow steps 2 to 6 above. Their data stays in their copy.'],
    ['Note: anyone who has your tracker link can open and log in your tracker, so share your sheet, not your link.']
  ]);

  makeTab_(ss, TAB.settings, ['Setting', 'Value', 'Help'], [
    ['Your name', 'Your name', 'Shown at the top of the tracker.'],
    ['Start date', '2026-10-17', 'Day 1 of your plan, written as YYYY-MM-DD.']
  ], function (sh) {
    sh.getRange('B3').setNumberFormat('@').setValue('2026-10-17');
  });

  makeTab_(ss, TAB.schedule, ['Day type', 'Time', 'Activity', 'Category', 'Planned minutes', 'Key'], [
    ['Weekday', '8:40 - 9:00 am', 'Typing practice', 'Typing practice', 20, 'typing'],
    ['Weekday', '9 am - 12 pm', 'Talent Nation work', 'Talent Nation work', 180, 'tn1'],
    ['Weekday', '1 - 2 pm', 'Talent Nation work', 'Talent Nation work', 60, 'tn2'],
    ['Weekday', '2 - 6 pm', 'AI Engineering from Scratch', 'AI Engineering from Scratch', 240, 'aiefs'],
    ['Weekday', '9 - 10 pm', 'YouTube tutorials', 'Python, YouTube & learning', 60, 'youtube'],
    ['Weekday', '10 pm - 12 am', 'Personal projects', 'Personal projects', 120, 'projects'],
    ['Weekday', '12 - 1 am', 'Python fundamentals', 'Python, YouTube & learning', 60, 'python'],
    ['Weekday', '1 - 2 am', 'Research and other work', 'Research & other work', 60, 'research'],
    ['Weekend', '8:40 - 9:00 am', 'Typing practice', 'Typing practice', 20, 'typing'],
    ['Weekend', '9 am - 12 pm', 'Talent Nation work', 'Talent Nation work', 180, 'tn1'],
    ['Weekend', '1 - 5 pm', 'AI Engineering from Scratch', 'AI Engineering from Scratch', 240, 'aiefs'],
    ['Weekend', '10 pm - 12 am', 'Personal projects', 'Personal projects', 120, 'projects'],
    ['Weekend', '12 - 1 am', 'Other learning and YouTube', 'Python, YouTube & learning', 60, 'learning']
  ]);

  makeTab_(ss, TAB.plan, ['Day', 'Key', 'Task', 'Link'], [
    [1, 'aiefs', 'Read the course README: "How this works" and "The shape of a lesson"', 'https://github.com/rohitg00/ai-engineering-from-scratch'],
    [1, 'aiefs', 'Prerequisites guide', 'https://aiengineeringfromscratch.com/prereqs.html'],
    [1, 'aiefs', '0.01 Dev Environment: install Python, run the setup check', GH + '00-setup-and-tooling/01-dev-environment'],
    [1, 'aiefs', '0.08 Editor Setup', GH + '00-setup-and-tooling/08-editor-setup'],
    [1, 'learning', 'Kaggle Python: Hello Python, then Functions', KG],
    [2, 'aiefs', '0.10 Terminal & Shell', GH + '00-setup-and-tooling/10-terminal-and-shell'],
    [2, 'aiefs', '0.11 Linux for AI', GH + '00-setup-and-tooling/11-linux-for-ai'],
    [2, 'aiefs', '0.02 Git & Collaboration: create a GitHub repo called ai-learning-log', GH + '00-setup-and-tooling/02-git-and-collaboration'],
    [2, 'learning', 'Kaggle Python: Booleans and Conditionals', KG],
    [3, 'aiefs', 'Clone the course repo', 'https://github.com/rohitg00/ai-engineering-from-scratch'],
    [3, 'aiefs', '0.06 Python Environments', GH + '00-setup-and-tooling/06-python-environments'],
    [3, 'aiefs', '0.05 Jupyter Notebooks', GH + '00-setup-and-tooling/05-jupyter-notebooks'],
    [3, 'python', 'Kaggle Python: Lists', KG],
    [3, 'youtube', 'Andrej Karpathy: Intro to Large Language Models (1-hour talk)', 'https://www.youtube.com/@AndrejKarpathy'],
    [4, 'aiefs', '0.09 Data Management', GH + '00-setup-and-tooling/09-data-management'],
    [4, 'aiefs', '0.04 APIs & Keys: keep API keys out of GitHub', GH + '00-setup-and-tooling/04-apis-and-keys'],
    [4, 'python', 'Kaggle Python: Loops and List Comprehensions', KG],
    [4, 'youtube', 'Corey Schafer: Git tutorial for beginners', 'https://www.youtube.com/@coreyms'],
    [5, 'aiefs', '0.12 Debugging & Profiling', GH + '00-setup-and-tooling/12-debugging-and-profiling'],
    [5, 'aiefs', '0.07 Docker for AI (read for concepts if Docker won\'t install)', GH + '00-setup-and-tooling/07-docker-for-ai'],
    [5, 'python', 'Kaggle Python: Strings and Dictionaries', KG],
    [5, 'youtube', 'Corey Schafer: Python virtual environments (venv)', 'https://www.youtube.com/@coreyms'],
    [6, 'aiefs', '0.03 GPU Setup & Cloud (read for concepts)', GH + '00-setup-and-tooling/03-gpu-setup-and-cloud'],
    [6, 'aiefs', '1.01 Linear Algebra Intuition: run vectors.py and save the output', GH + '01-math-foundations/01-linear-algebra-intuition'],
    [6, 'python', 'Kaggle Python: Working with External Libraries', KG],
    [6, 'youtube', '3Blue1Brown: Essence of Linear Algebra, chapter 1', 'https://www.youtube.com/@3blue1brown'],
    [7, 'aiefs', 'Phase 0 quiz with Claude, then revisit your two weakest lessons', ''],
    [7, 'aiefs', '1.02 Vectors, Matrices & Operations, at a relaxed pace', GH + '01-math-foundations/02-vectors-matrices-operations'],
    [7, 'python', 'Review the Kaggle exercises you found hard', KG],
    [7, 'youtube', '3Blue1Brown: Essence of Linear Algebra, chapters 2 - 3', 'https://www.youtube.com/@3blue1brown'],
    [7, 'research', 'Write your weekly reflection and plan Week 2 with Claude', '']
  ]);

  makeTab_(ss, TAB.resources, ['Name', 'Link', 'What it is for', 'Group'], [
    ['AI Engineering from Scratch', 'https://aiengineeringfromscratch.com', 'Main course: read lessons here', 'Courses'],
    ['AI Engineering from Scratch on GitHub', 'https://github.com/rohitg00/ai-engineering-from-scratch', 'Lesson code and the repo to clone', 'Courses'],
    ['Kaggle Learn: Python', 'https://www.kaggle.com/learn/python', 'Free Python course for your Python slot', 'Courses'],
    ['3Blue1Brown', 'https://www.youtube.com/@3blue1brown', 'Visual math: linear algebra, calculus, neural networks', 'YouTube'],
    ['Andrej Karpathy', 'https://www.youtube.com/@AndrejKarpathy', 'How LLMs and neural networks work, from scratch', 'YouTube'],
    ['StatQuest with Josh Starmer', 'https://www.youtube.com/@statquest', 'Statistics and classic ML explained simply', 'YouTube'],
    ['Corey Schafer', 'https://www.youtube.com/@coreyms', 'Clear Python, Git and tooling tutorials', 'YouTube'],
    ['Sentdex', 'https://www.youtube.com/@sentdex', 'Hands-on Python and ML projects', 'YouTube'],
    ['freeCodeCamp', 'https://www.youtube.com/@freecodecamp', 'Long, structured full courses', 'YouTube']
  ]);

  if (!ss.getSheetByName(TAB.log)) makeLogSheet_(ss);

  // Put tabs in a sensible order and remove an empty default "Sheet1".
  [TAB.start, TAB.settings, TAB.schedule, TAB.plan, TAB.resources, TAB.log].forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  ss.setActiveSheet(ss.getSheetByName(TAB.start));
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function makeTab_(ss, name, headers, data, after) {
  if (ss.getSheetByName(name)) return; // never overwrite someone's edits
  var sh = ss.insertSheet(name);
  var r = 1;
  if (headers) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#e6ebfb').setFontColor('#1f2a44');
    sh.setFrozenRows(1);
    r = 2;
  }
  if (data && data.length) {
    var width = data.reduce(function (m, row) { return Math.max(m, row.length); }, 0);
    var padded = data.map(function (row) { while (row.length < width) row.push(''); return row; });
    sh.getRange(r, 1, padded.length, width).setValues(padded);
  }
  if (name === TAB.start) {
    sh.setColumnWidth(1, 900);
    sh.getRange('A1').setFontSize(18).setFontWeight('bold').setFontColor('#1f2a44');
    ['A5', 'A12', 'A20', 'A24'].forEach(function (a) { sh.getRange(a).setFontWeight('bold').setFontColor('#2f4fd1'); });
    sh.getRange(1, 1, sh.getLastRow(), 1).setWrap(true);
  } else {
    sh.autoResizeColumns(1, sh.getLastColumn());
  }
  if (after) after(sh);
}

function makeLogSheet_(ss) {
  var sh = ss.insertSheet(TAB.log);
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
    .setFontWeight('bold').setBackground('#e6ebfb').setFontColor('#1f2a44');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

function rows_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(1, sh.getLastColumn())).getValues();
}

function toIsoDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return '';
}

function slug_(v) {
  return String(v || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
}
