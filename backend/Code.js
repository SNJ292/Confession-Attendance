/** ===== Sheet Helpers ===== **/
function getSheet_(name){
  return SpreadsheetApp.getActive().getSheetByName(name);
}

function getOrCreateSheet_(name, headers){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && sh.getLastRow() === 0){
    sh.appendRow(headers);
  }
  return sh;
}

/** ===== Settings Helper ===== **/
function getSettings_() {
  const sh = getSheet_("Settings");
  const o = {};
  if (!sh) return o;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return o;

  const vals = sh.getRange(1,1,lastRow,2).getValues(); // Key, Value
  vals.forEach(([k,v]) => {
    if (k) o[String(k).trim()] = String(v || "").trim();
  });

  // Defaults
  if (!o.TIMEZONE) {
    o.TIMEZONE = Session.getScriptTimeZone() || "America/New_York";
  }
  if (!o.HISTORY_DEPTH) {
    o.HISTORY_DEPTH = 3;
  } else {
    o.HISTORY_DEPTH = Math.max(1, Number(o.HISTORY_DEPTH));
  }

  return o;
}

/** ===== Add custom menu in the Sheet ===== **/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Confession Attendance')
    .addItem('Build next Saturday roster', 'buildNextSaturdayRoster')
    .addToUi();

  // Ensure basic sheets exist
  getOrCreateSheet_('Roster', ['Date','Name','Email']);
  getOrCreateSheet_('Attendance', ['Date','Name','Email','BaptismalName','Status','Timestamp']);
  getOrCreateSheet_('AttendanceDraft', ['Date','Name','Email','BaptismalName','Status','Timestamp']);
}

// Returns a Date object representing THIS Saturday if today is Saturday,
// otherwise the next upcoming Saturday.
function nextOrThisSaturday_(tz){
  const now = new Date();
  const d = new Date(now);
  const day = d.getDay(); // 0=Sun, 6=Sat

  // days until Saturday (0 if today is Saturday)
  const add = (6 - day + 7) % 7;

  d.setDate(d.getDate() + add);
  d.setHours(0,0,0,0);
  return d;
}

// Returns start and end Date objects for that day
function dayWindow_(d){
  const start = new Date(d);
  start.setHours(0,0,0,0);

  const end = new Date(d);
  end.setHours(23,59,59,999);

  return { start, end };
}

/** ===== Build Roster for a Given Date from Calendar ===== **/
function buildRosterForDateFromCalendar(dateStr) {
  const cfg = getSettings_();
  const tz = cfg.TIMEZONE;

  // 1) Decide which day we're building for
  let targetDate;
  if (dateStr) {
    // dateStr is "yyyy-MM-dd" from the HTML date input
    const [y, m, d] = dateStr.split('-').map(Number);
    targetDate = new Date(y, m - 1, d);
    targetDate.setHours(0, 0, 0, 0);
  } else {
    // fallback: next Saturday
    targetDate = nextOrThisSaturday_(tz);
  }
  const { start, end } = dayWindow_(targetDate);

  // 2) Get calendar
  const cal = cfg.CALENDAR_ID
    ? CalendarApp.getCalendarById(cfg.CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();

  if (!cal) {
    throw new Error('Calendar not found. Check CALENDAR_ID in Settings.');
  }

  const events = cal.getEvents(start, end);
  const filter = (cfg.EVENT_FILTER || '').toLowerCase();
  const priestEmail = (cfg.PRIEST_EMAIL || '').trim().toLowerCase();

  const roster = new Map();

  // 3) Build roster from events that day
  events.forEach(ev => {
    // Optional: filter by title containing "Confession"
    // if (filter && !String(ev.getTitle()).toLowerCase().includes(filter)) {
    //   return;
    // }

    const guests = ev.getGuestList(true);

    // Confessors = all guests EXCEPT the priest (if priestEmail is configured)
    let confessors = guests;
    
    if (priestEmail) {
      confessors = guests.filter(g => {
        const em = (g.getEmail() || '').trim().toLowerCase();
        return em && em !== priestEmail;
      });
    }

    // If we filtered everybody out (bad priestEmail?), fall back to all guests
    if (!confessors.length && guests.length) {
      confessors = guests;
    }

    confessors.forEach(g => {
      const titleName = extractNameFromTitle(ev.getTitle());
      const email = (g.getEmail() || '').trim();

      // If the title gives us a clean name, use it.
      let name = titleName || g.getName() || email || '';
      name = name.trim();
      if (!email && !name) return;

      const key = (email || name).toLowerCase();
      if (!roster.has(key)) {
        roster.set(key, { name, email });
      }
    });
  });

  // 4) Write roster rows for that date into the sheet (overwrite old roster)
  const dateStrOut = Utilities.formatDate(start, tz, 'yyyy-MM-dd');

  const sh = getOrCreateSheet_('Roster', ['Date','Name','Email']);
  sh.clearContents();
  sh.appendRow(['Date','Name','Email']);

  const rows = [...roster.values()]
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(p => [dateStrOut, p.name, p.email]);

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  const result = { date: dateStrOut, count: rows.length };
  return result;
}

function extractNameFromTitle(title) {
  if (!title) return null;

  // Normalize spacing
  title = String(title).trim();

  // Match Confession (...) OR NEW Confession (...)
  // Captures whatever is inside parentheses.
  const match = title.match(/\(([^)]+)\)/);

  if (match && match[1]) {
    return match[1].trim();
  }

  return null; // fallback
}

/** ===== Build Next Saturday Roster from Calendar ===== **/
function buildNextSaturdayRoster() {
    // Just call the generic helper with null → next Saturday
  return buildRosterForDateFromCalendar(null);
}

function getRosterAndHistory(dateStr){
  const cfg = getSettings_();
  const tz = cfg.TIMEZONE;

  // 1) Always (re)build the roster for the requested date from the calendar
  //    If dateStr is null/empty, this defaults to next Saturday.
  const rosterRes = buildRosterForDateFromCalendar(dateStr);
  dateStr = rosterRes.date;  // normalize to "yyyy-MM-dd"

  const rosterSh = getSheet_("Roster");
  const attSh    = getSheet_("Attendance");

  let people = [];
  if (rosterSh && rosterSh.getLastRow() > 1){
    const vals = rosterSh.getRange(2,1,rosterSh.getLastRow()-1,3).getValues();
    // vals: [Date, Name, Email]

    people = vals
      .filter(r => {
        const cell = r[0];
        let cellDateStr;

        if (cell instanceof Date) {
          // Format real Date object to yyyy-MM-dd
          cellDateStr = Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
        } else {
          // Already a string or number
          cellDateStr = String(cell);
        }

        return cellDateStr === dateStr;
      })
      .map(r => ({ name: r[1], email: r[2] }));
  }

  // ✅ NEW: map roster name -> email
  const rosterNameToEmail = {};
  people.forEach(p => {
    const nameKey  = String(p.name || '').trim().toLowerCase();
    const emailKey = String(p.email || '').trim().toLowerCase();
    if (nameKey && emailKey) rosterNameToEmail[nameKey] = emailKey;
  });

  // Build history map from Attendance
  const history = {};
  const depth = cfg.HISTORY_DEPTH || 3;

  if (attSh && attSh.getLastRow() > 1){
    const vals = attSh.getRange(2,1,attSh.getLastRow()-1,6).getValues();
    // vals: [Date, Name, Email, Status, Timestamp]
    vals.reverse(); // newest first

    for (const [d, nm, em, bn, st] of vals){
      const emailKey = String(em || '').trim().toLowerCase();
      const nameKey  = String(nm || '').trim().toLowerCase();

      // Prefer email; if missing, map name -> roster email; else fallback to name
      const key = emailKey || rosterNameToEmail[nameKey] || nameKey;
      if (!key) continue;
      if (!history[key]) history[key] = [];
      if (history[key].length < depth){
        const dateStr = (d instanceof Date)
          ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
          : String(d);

        history[key].push({ date: dateStr, status: st, name: nm });
      }
    }
  }

  const histForPage = {};
  people.forEach(p => {
    const key = String(p.email || p.name || '').trim().toLowerCase();
    histForPage[key] = history[key] || [];
  });

  const result = {
    date: dateStr,
    people: people,
    history: histForPage,
    historyDepth: depth
  };

  return result;
}

/** ===== Save draft (no email) ===== **/
function saveAttendanceDraft(payload) {
  const cfg = getSettings_();
  const tz = cfg.TIMEZONE;

  const dateStr = payload.date;
  const marked  = payload.marked || []; // [{name,email,status}, ...]

  const sh = getOrCreateSheet_('AttendanceDraft', ['Date','Name','Email', 'BaptismalName','Status','Timestamp']);
  const ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  // Remove existing draft rows for this date (so draft is "latest state")
  if (sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues();
    const keep = vals.filter(r => String(r[0]) !== String(dateStr));
    sh.clearContents();
    sh.appendRow(['Date','Name','Email', 'BaptismalName', 'Status','Timestamp']);
    if (keep.length) sh.getRange(2,1,keep.length,6).setValues(keep);
  }

  const rows = marked.map(p => [
    dateStr,
    p.name  || "",
    p.email || "",
    p.baptismalName || "",
    p.status || "",
    ts
  ]);

  if (rows.length) {
    sh.getRange(sh.getLastRow()+1, 1, rows.length, 6).setValues(rows);
  }

  return { ok: true, saved: rows.length };
}

/** ===== Load draft selections for a date ===== **/
function getAttendanceDraft(dateStr) {
  const cfg = getSettings_();
  const tz = cfg.TIMEZONE;

  const sh = getSheet_('AttendanceDraft');
  if (!sh || sh.getLastRow() < 2) return {};

  const vals = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
  // key by email/name

  const out = {};
  for (const [d, nm, em, bn, st] of vals) {
    // Normalize sheet date -> yyyy-MM-dd
    const dStr = (d instanceof Date)
      ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
      : String(d).trim();

    if (dStr !== String(dateStr).trim()) continue;
    const key = (em || nm || "").toLowerCase();
    if (!key) continue;

    // Skip unselected/blank statuses so nothing gets preselected
    const status = String(st || '').trim();
    if (!status) continue;

    const baptismalName = String(bn || '').trim();

    // Optional: normalize to exact casing
    const norm =
      status.toLowerCase() === 'present' ? 'Present' :
      status.toLowerCase() === 'absent'  ? 'Absent'  :
      status;

    out[key] = {status: norm, baptismalName};
  }

  return out;
}

function clearAttendanceDraftForDate_(dateStr) {
  const sh = getSheet_('AttendanceDraft');
  if (!sh || sh.getLastRow() < 2) return { ok: true, removed: 0 };

  const cfg = getSettings_();
  const tz = cfg.TIMEZONE;

  const last = sh.getLastRow();
  const vals = sh.getRange(2, 1, last - 1, 6).getValues();

  // Keep header + rows NOT matching this date
  const keep = [];
  let removed = 0;

  for (const row of vals) {
    const d = row[0];
    const dStr = (d instanceof Date)
      ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
      : String(d).trim();

    if (dStr === String(dateStr).trim()) {
      removed++;
    } else {
      keep.push(row);
    }
  }

  sh.clearContents();
  sh.appendRow(['Date','Name','Email', 'BaptismalName','Status','Timestamp']);
  if (keep.length) {
    sh.getRange(2, 1, keep.length, 6).setValues(keep);
  }

  return { ok: true, removed };
}

/** ===== Save attendance + email Amy & Jamie ===== **/
function saveAttendanceAndEmail(payload){
  const cfg = getSettings_();
  const tz = cfg.TIMEZONE;

  const dateStr = payload.date;
  const marked  = payload.marked || []; // [{name,email,status}, ...]

  // 1) Save to Attendance sheet
  const sh = getOrCreateSheet_('Attendance', ['Date','Name','Email', 'BaptismalName', 'Status','Timestamp']);
  const ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  const rows = marked.map(p => [
    dateStr,
    p.name  || "",
    p.email || "",
    p.baptismalName || "",
    p.status || "",
    ts
  ]);

  if (rows.length > 0){
    sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // 2) Build present/absent lists for email
  const presents = marked.filter(p => p.status === 'Present');
  const absents  = marked.filter(p => p.status === 'Absent');

  // 3) Email Amy (present)
  if (cfg.ROBEL_EMAIL && presents.length){
    const bodyLines = [
      `Confession attendance for ${dateStr}`,
      ``,
      `Present (${presents.length}):`,
      ...presents.map(p => {
        const bn = p.baptismalName ? ` (Baptismal: ${p.baptismalName})` : '';
        return `• ${p.name}${bn}${p.email ? ' <'+p.email+'>' : ''}`;
      })
    ];
    const body = bodyLines.join('\n');

    MailApp.sendEmail({
      to: cfg.ROBEL_EMAIL,
      subject: `Confession Attendance – Present – ${dateStr}`,
      body: body
    });
  }

  // 4) Email Jamie (absent)
  if (cfg.ANNA_EMAIL && absents.length && presents.length){
    const bodyLines = [
      `Confession attendance for ${dateStr}`,
      ``,
      `Absent (${absents.length}):`,
      ...absents.map(p => {
        const bn = p.baptismalName ? ` (Baptismal: ${p.baptismalName})` : '';
        return `• ${p.name}${bn}${p.email ? ' <'+p.email+'>' : ''}`;
      }),
      '',
      `Present (${presents.length}):`,
      ...presents.map(p => {
        const bn = p.baptismalName ? ` (Baptismal: ${p.baptismalName})` : '';
        return `• ${p.name}${bn}${p.email ? ' <'+p.email+'>' : ''}`;
      })
    ];
    const body = bodyLines.join('\n');

    MailApp.sendEmail({
      to: cfg.ANNA_EMAIL,
      subject: `Confession Attendance – Absent – ${dateStr}`,
      body: body
    });
  }

  const result = {
    ok: true,
    saved: rows.length,
    presentCount: presents.length,
    absentCount: absents.length
  };


  // ✅ Clear draft after final submit so UI doesn't preload old draft
  clearAttendanceDraftForDate_(dateStr);
  
  return result;
}

// TEST FUNCTION – DO NOT RUN AGAINST PRODUCTION DATA
// function test_getRosterAndHistory(){
//   const cfg = getSettings_();
//   const tz = cfg.TIMEZONE;
//   const dateStr = Utilities.formatDate(nextOrThisSaturday_(tz), tz, 'yyyy-MM-dd');
//   const res = getRosterAndHistory(dateStr);
// }

// function test_saveAttendanceAndEmail(){
//   const cfg = getSettings_();
//   const tz = cfg.TIMEZONE;
//   const dateStr = Utilities.formatDate(nextOrThisSaturday_(tz), tz, 'yyyy-MM-dd');

//   const payload = {
//     date: dateStr,
//     marked: [
//       { name: 'John Doe',  email: 'john@example.com', status: 'Present' },
//       { name: 'Mary Smith', email: 'mary@example.com', status: 'Absent' }
//     ]
//   };

//   const res = saveAttendanceAndEmail(payload);
// }

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Confession Attendance')
}