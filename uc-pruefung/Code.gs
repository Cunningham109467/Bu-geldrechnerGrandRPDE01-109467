// ============================================================
//  UC Prüfung – CID Portal  |  Bewertungssystem
//  Spreadsheet: 1LXZuycLyk-V7A89cxiKT08y7Or_1dKqOqTQ5NQmw2oQ
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1LXZuycLyk-V7A89cxiKT08y7Or_1dKqOqTQ5NQmw2oQ',
  FORM_SHEET:     'Formularantworten 2',  // ← Sheet-Name der Formular-Antworten
  USERS_SHEET:    'Benutzer',
  LOG_SHEET:      'Änderungslog',
  GRADES_SHEET:   'Bewertungen',
  MAX_POINTS:     72,    // Gesamtpunktzahl
  PASS_POINTS:    60,    // Mindestpunkte zum Bestehen
  CANDIDATE_COL:  2,     // 0-basierter Spaltenindex des Namens (Spalte 0=Timestamp, 1=F1, 2=F2=NAME)
  BAN_FAILS:      3,     // Misserfolge vor Sperrung
  BAN_DAYS:       14     // Sperrdauer in Tagen
};

// ── Kandidaten-Spalte dynamisch (aus ScriptProperties oder CONFIG) ───
function _getCandidateCol() {
  const stored = PropertiesService.getScriptProperties().getProperty('CANDIDATE_COL');
  return stored !== null && stored !== undefined ? Number(stored) : CONFIG.CANDIDATE_COL;
}

// ── App-Konfiguration lesen ──────────────────────────────────
function getAppConfig(token) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };

  const props      = PropertiesService.getScriptProperties();
  const candCol    = _getCandidateCol();
  const mediaConf  = _safeJSON(props.getProperty('MEDIA_CONFIG') || '{}');

  let headers = [];
  try {
    const ss  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sht = ss.getSheetByName(CONFIG.FORM_SHEET);
    if (sht && sht.getLastColumn() > 0)
      headers = sht.getRange(1, 1, 1, sht.getLastColumn()).getValues()[0].map(String);
  } catch (e) { /* ignore */ }

  return { success: true, candidateCol: candCol, mediaConfig: mediaConf, headers };
}

// ── App-Konfiguration speichern (Level 2) ────────────────────
function saveAppConfig(token, candidateCol, mediaJson) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };
  if (session.level < 2) return { success: false, error: 'Keine Berechtigung.' };

  const props = PropertiesService.getScriptProperties();
  props.setProperty('CANDIDATE_COL', String(Number(candidateCol)));
  props.setProperty('MEDIA_CONFIG',  mediaJson || '{}');
  return { success: true };
}

// ── Web-App ──────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('UC Prüfung – CID Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Passwort-Hashing (SHA-256) ───────────────────────────────
function hashPwd(plaintext) {
  const b = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plaintext),
    Utilities.Charset.UTF_8
  );
  return b.map(x => ('0' + (x & 0xff).toString(16)).slice(-2)).join('');
}

// ── Session ──────────────────────────────────────────────────
function _saveSession(token, username, level) {
  PropertiesService.getScriptProperties().setProperty(
    'sess_' + token,
    JSON.stringify({ username, level: Number(level), expiry: new Date(Date.now() + 28800000).toISOString() })
  );
}

function validateSession(token) {
  if (!token) return null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (new Date(s.expiry) < new Date()) {
      PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
      return null;
    }
    return s;
  } catch (e) { return null; }
}

// ── Login / Logout ───────────────────────────────────────────
function login(username, password) {
  if (!username || !password) return { success: false, error: 'Alle Felder ausfüllen.' };
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
    if (!sheet) return { success: false, error: 'Benutzer-Sheet fehlt – bitte setupAdmin() ausführen.' };

    const rows   = sheet.getDataRange().getValues();
    const hashed = hashPwd(password);

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === username.trim().toLowerCase()
          && String(rows[i][1]).trim() === hashed) {
        const token = Utilities.getUuid();
        const level = Number(rows[i][2]) || 1;
        _saveSession(token, String(rows[i][0]).trim(), level);
        return { success: true, token, username: String(rows[i][0]).trim(), level };
      }
    }
    return { success: false, error: 'Benutzername oder Passwort falsch.' };
  } catch (e) {
    return { success: false, error: 'Serverfehler: ' + e.message };
  }
}

function logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
  return { success: true };
}

// ── Alle Bewertungen laden ───────────────────────────────────
function _loadGradesMap(ss) {
  const sheet = ss.getSheetByName(CONFIG.GRADES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  const map  = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]);
    map[key] = {
      kandidat:     String(data[i][1] || ''),
      bewerter:     String(data[i][2] || ''),
      datum:        data[i][3] instanceof Date
        ? Utilities.formatDate(data[i][3], tz, 'dd.MM.yyyy HH:mm')
        : String(data[i][3] || ''),
      gesamtpunkte: Number(data[i][4]) || 0,
      bestanden:    data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE',
      notizen:      String(data[i][6] || ''),
      grades:       _safeJSON(String(data[i][7] || '{}'))
    };
  }
  return map;
}

function _safeJSON(str) {
  try { return JSON.parse(str); } catch (e) { return {}; }
}

// ── Einreichungsliste ────────────────────────────────────────
function getSubmissions(token) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };

  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const fSht  = ss.getSheetByName(CONFIG.FORM_SHEET);
    if (!fSht)  return { success: false, error: 'Formular-Sheet "' + CONFIG.FORM_SHEET + '" nicht gefunden.' };

    const data  = fSht.getDataRange().getValues();
    if (data.length < 2) return { success: true, submissions: [] };

    const tz    = Session.getScriptTimeZone();
    const gMap  = _loadGradesMap(ss);

    const candCol = _getCandidateCol();
    const submissions = data.slice(1).map((row, i) => {
      const tsKey     = String(row[0]);
      const ts        = row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'dd.MM.yyyy HH:mm') : tsKey;
      const candidate = String(row[candCol] !== undefined ? row[candCol] : '–').trim() || '–';
      const g         = gMap[tsKey] || null;
      return {
        formRow:   i + 2,
        ts, tsKey, candidate,
        punkte:    g ? g.gesamtpunkte : null,
        bestanden: g ? g.bestanden    : null,
        bewerter:  g ? g.bewerter     : null,
        datum:     g ? g.datum        : null,
        status:    g ? (g.bestanden ? 'bestanden' : 'nicht_bestanden') : 'unbewertet'
      };
    });

    return { success: true, submissions };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Einzel-Einreichung mit Bewertungs-Detail ─────────────────
function getSubmissionDetail(token, formRow) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };

  try {
    const ss   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const fSht = ss.getSheetByName(CONFIG.FORM_SHEET);
    const data = fSht.getDataRange().getValues();
    const tz   = Session.getScriptTimeZone();

    if (formRow < 2 || formRow > data.length) return { success: false, error: 'Zeile ungültig.' };

    const candCol   = _getCandidateCol();
    const headers   = data[0];
    const row       = data[formRow - 1];
    const tsKey     = String(row[0]);
    const candidate = String(row[candCol] !== undefined ? row[candCol] : '–').trim() || '–';
    const ts        = row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'dd.MM.yyyy HH:mm') : tsKey;

    // Alle Fragen + Antworten (Spalte 0 = Timestamp überspringen)
    const qa = headers.map((h, i) => ({
      colIndex: i,
      question: String(h),
      answer:   String(row[i] === null || row[i] === undefined ? '' : row[i])
    })).filter(q => q.colIndex > 0); // Timestamp überspringen

    // Bestehende Bewertung
    const gMap     = _loadGradesMap(ss);
    const existing = gMap[tsKey] || null;

    // Verlauf des Prüflings
    const history = data.slice(1).map((r, i) => {
      const k = String(r[0]);
      const g = gMap[k];
      return {
        formRow:   i + 2,
        ts:        r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd.MM.yyyy HH:mm') : k,
        candidate: String(r[candCol] !== undefined ? r[candCol] : '').trim(),
        punkte:    g ? g.gesamtpunkte : null,
        status:    g ? (g.bestanden ? 'bestanden' : 'nicht_bestanden') : 'unbewertet'
      };
    }).filter(s => s.candidate.toLowerCase() === candidate.toLowerCase());

    // Bann-Prüfung: X aufeinander folgende Misserfolge
    const bewertet = history
      .filter(s => s.status !== 'unbewertet')
      .sort((a, b) => b.formRow - a.formRow);

    let consecutive = 0;
    for (const s of bewertet) {
      if (s.status === 'nicht_bestanden') consecutive++;
      else break;
    }

    let banInfo = null;
    if (consecutive >= CONFIG.BAN_FAILS && bewertet.length > 0) {
      const lastRow  = data[bewertet[0].formRow - 1][0];
      const lastDate = lastRow instanceof Date ? lastRow : new Date(lastRow);
      const banUntil = new Date(lastDate.getTime() + CONFIG.BAN_DAYS * 86400000);
      if (banUntil > new Date()) {
        banInfo = { banned: true, until: Utilities.formatDate(banUntil, tz, 'dd.MM.yyyy') };
      }
    }

    // Medien-Konfiguration laden
    const mediaConfig = _safeJSON(
      PropertiesService.getScriptProperties().getProperty('MEDIA_CONFIG') || '{}'
    );

    return {
      success: true, formRow, tsKey, ts, candidate, qa, existing, history,
      banInfo, maxPoints: CONFIG.MAX_POINTS, passPoints: CONFIG.PASS_POINTS, mediaConfig
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Bewertung speichern ──────────────────────────────────────
function saveGrading(token, formRow, gradesJson, notes) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };

  try {
    const ss   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const fSht = ss.getSheetByName(CONFIG.FORM_SHEET);
    const data = fSht.getDataRange().getValues();

    if (formRow < 2 || formRow > data.length) return { success: false, error: 'Zeile ungültig.' };

    const row       = data[formRow - 1];
    const tsKey     = String(row[0]);
    const candidate = String(row[_getCandidateCol()] || '–').trim();
    const grades    = _safeJSON(gradesJson);
    const total     = Object.values(grades).reduce((s, v) => s + (Number(v) || 0), 0);
    const passed    = total >= CONFIG.PASS_POINTS;

    // Bewertungen-Sheet anlegen wenn nötig
    let gSht = ss.getSheetByName(CONFIG.GRADES_SHEET);
    if (!gSht) {
      gSht = ss.insertSheet(CONFIG.GRADES_SHEET);
      gSht.appendRow(['Form_Timestamp','Kandidat','Bewerter','Datum','Gesamtpunkte','Bestanden','Notizen','Grades_JSON']);
      gSht.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    // Existierende Zeile finden oder neue anhängen
    const gData   = gSht.getLastRow() > 1 ? gSht.getDataRange().getValues() : [[]];
    let targetRow = -1;
    for (let i = 1; i < gData.length; i++) {
      if (String(gData[i][0]) === tsKey) { targetRow = i + 1; break; }
    }

    const newRow = [tsKey, candidate, session.username, new Date(), total, passed, notes || '', gradesJson];
    if (targetRow > 0) {
      gSht.getRange(targetRow, 1, 1, 8).setValues([newRow]);
    } else {
      gSht.appendRow(newRow);
    }

    // Log
    _log(ss, session.username, formRow, candidate, total + '/' + CONFIG.MAX_POINTS + (passed ? ' BESTANDEN' : ' NICHT BESTANDEN'));

    return { success: true, total, passed, maxPoints: CONFIG.MAX_POINTS };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _log(ss, user, row, candidate, value) {
  let log = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET);
    log.appendRow(['Zeitstempel','Bewerter','Zeile','Kandidat','Ergebnis']);
    log.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  log.appendRow([new Date(), user, row, candidate, value]);
}

// ── Wochenstatistik (Level 2) ────────────────────────────────
function getWeeklyStats(token) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };
  if (session.level < 2) return { success: false, error: 'Keine Berechtigung.' };

  try {
    const ss  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const log = ss.getSheetByName(CONFIG.LOG_SHEET);
    const tz  = Session.getScriptTimeZone();

    const now    = new Date();
    const dow    = now.getDay();
    const monday = new Date(now);
    monday.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1));
    monday.setHours(0, 0, 0, 0);

    const map = {};
    if (log && log.getLastRow() > 1) {
      const rows = log.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const ts   = rows[i][0];
        const user = String(rows[i][1] || '').trim();
        if (!ts || !user) continue;
        const d = ts instanceof Date ? ts : new Date(ts);
        if (d < monday) continue;
        if (!map[user]) map[user] = { count: 0, last: null };
        map[user].count++;
        if (!map[user].last || d > map[user].last) map[user].last = d;
      }
    }

    const stats = Object.entries(map)
      .map(([name, s]) => ({
        name, count: s.count,
        last: s.last ? Utilities.formatDate(s.last, tz, 'dd.MM.yyyy HH:mm') : '–'
      }))
      .sort((a, b) => b.count - a.count);

    return {
      success: true, stats,
      weekStart: Utilities.formatDate(monday, tz, 'dd.MM.yyyy')
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Benutzerverwaltung (Level 2) ─────────────────────────────
function getUsers(token) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };
  if (session.level < 2) return { success: false, error: 'Keine Berechtigung.' };
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.USERS_SHEET);
    if (!sheet) return { success: true, users: [] };
    const users = sheet.getDataRange().getValues().slice(1)
      .map((r, i) => ({ _row: i + 2, name: r[0], level: Number(r[2]) || 1 }));
    return { success: true, users };
  } catch (e) { return { success: false, error: e.message }; }
}

function createUser(token, username, password, level) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };
  if (session.level < 2) return { success: false, error: 'Keine Berechtigung.' };

  username = String(username || '').trim();
  password = String(password || '').trim();
  level    = Number(level) === 2 ? 2 : 1;
  if (!username)         return { success: false, error: 'Benutzername fehlt.' };
  if (password.length < 6) return { success: false, error: 'Passwort min. 6 Zeichen.' };

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sht  = ss.getSheetByName(CONFIG.USERS_SHEET);
    if (!sht) {
      sht = ss.insertSheet(CONFIG.USERS_SHEET);
      sht.appendRow(['Benutzername','Passwort_Hash','Level']);
      sht.getRange(1, 1, 1, 3).setFontWeight('bold');
    }
    const rows = sht.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === username.toLowerCase())
        return { success: false, error: 'Benutzer existiert bereits.' };
    }
    sht.appendRow([username, hashPwd(password), level]);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function deleteUser(token, row) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };
  if (session.level < 2) return { success: false, error: 'Keine Berechtigung.' };
  try {
    SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.USERS_SHEET).deleteRow(row);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function changePassword(token, row, newPwd) {
  const session = validateSession(token);
  if (!session) return { success: false, error: 'Sitzung abgelaufen.' };
  if (session.level < 2) return { success: false, error: 'Keine Berechtigung.' };
  newPwd = String(newPwd || '').trim();
  if (newPwd.length < 6) return { success: false, error: 'Passwort min. 6 Zeichen.' };
  try {
    SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      .getSheetByName(CONFIG.USERS_SHEET).getRange(row, 2).setValue(hashPwd(newPwd));
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Einmalig: Admin erstellen ────────────────────────────────
function setupAdmin() {
  const ss  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.USERS_SHEET);
    sheet.appendRow(['Benutzername','Passwort_Hash','Level']);
    sheet.getRange(1,1,1,3).setFontWeight('bold');
  }
  const pwd = 'Admin1234';
  sheet.appendRow(['admin', hashPwd(pwd), 2]);
  Logger.log('Admin erstellt | User: admin | Passwort: ' + pwd + ' | Bitte sofort ändern!');
}
