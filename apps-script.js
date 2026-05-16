// ============================================================
// Apps Script для игры "Ты не случаен"
// Хранит сырые данные одной игры на строку. Идемпотентные хедеры:
// если COLUMNS расширяется — лист сам дописывает недостающие.
// ============================================================
//
// УСТАНОВКА (для новой Spreadsheet):
// 1. Создай новую Google Spreadsheet, имя на свой вкус
//    (например: "ty-ne-sluchaen-research")
// 2. Extensions → Apps Script
// 3. Удали всё из Code.gs, вставь этот файл целиком
// 4. Cmd+S (сохранить)
// 5. Deploy → New deployment → Type: Web app
//    - Description: "v1 init"
//    - Execute as: Me
//    - Who has access: Anyone
//    - Deploy → Authorize (выберешь свой Google аккаунт)
// 6. Скопируй Web App URL — это новый SHEETS_URL для index.html
// 7. Открой Spreadsheet — лист 'Data' создастся автоматически при первой записи
//
// ОБНОВЛЕНИЕ КОДА (когда меняешь скрипт без смены URL):
// - Замени код, Cmd+S
// - Deploy → Manage deployments → ✏️ (Edit) → Version: New version → Deploy
// - URL остаётся прежним
//
// Эндпоинты:
//  POST /                         → запись одной игры (используется фронтом)
//  GET  /?action=stats            → агрегированные перцентили по уровням (JSON)
//  GET  /?action=stats&callback=X → то же, но JSONP (используется фронтом)
//  GET  /                         → ping
// ============================================================

// ───────── COLUMNS: единый источник правды ─────────
// Имя колонки → функция которая выдёргивает значение из payload.
// Чтобы добавить колонку: добавь сюда строку, скрипт сам допишет header.
const COLUMNS = [
  ['Timestamp',           ()  => new Date().toISOString()],
  ['Session ID',          d   => d.sessionId || ''],
  ['Attempt #',           d   => +d.attemptNumber || 0],
  ['Level',               d   => +d.level || ''],
  ['Total Presses',       d   => +d.totalPresses || 0],
  ['Correct Predictions', d   => +d.correctPredictions || 0],
  ['Accuracy %',          d   => +d.accuracy || 0],
  ['Balance',             d   => +d.balance || 0],
  ['Won',                 d   => d.won ? 'YES' : 'NO'],
  ['Reason',              d   => d.reason || ''],
  ['Ratio 0/1',           d   => d.ratio01 || ''],
  ['Max Run',             d   => +d.maxRun || 0],
  ['Top Trigram',         d   => d.topTrigram || ''],
  ['Sequence',            d   => d.sequence || ''],
  ['Predictions',         d   => d.predictions || ''],   // post-warmup, length = totalPresses - 3
  ['PredScores',          d   => d.predScores || ''],    // CSV, aligned with Predictions
  ['PointerTypes',        d   => d.pointerTypes || ''],  // per-tap chars: m/t/p/k
  ['TapsX',               d   => d.tapsX || ''],         // CSV, % of game width
  ['TapsY',               d   => d.tapsY || ''],         // CSV, % of game height
  ['Timestamps (ms)',     d   => d.timestamps || ''],
  ['Acc Last 10',         d   => valOrEmpty_(d.accLast10)],
  ['Acc Last 20',         d   => valOrEmpty_(d.accLast20)],
  ['Carryover Balance',   d   => valOrEmpty_(d.carryoverBalance)],
  ['Device',              d   => d.device || ''],
  ['Language',            d   => d.lang || ''],
  ['Age',                 d   => d.age || ''],            // opt-in: <18, 18-25, 26-35, 36-50, 50+, or 'skip'
  ['User Agent',          d   => d.userAgent || ''],
];

const DATA_SHEET = 'Data';
const ERR_SHEET  = 'Errors';
const STATS_CACHE_KEY = 'stats_v1';
const STATS_CACHE_TTL = 300; // sec

// Columns whose values are stringy but look numeric to Sheets autoparse
// (CSV of ints, sequences with leading zeros, etc.) — force TEXT format.
const TEXT_COLS = [
  'Session ID', 'Sequence', 'Predictions', 'PredScores', 'PointerTypes',
  'TapsX', 'TapsY', 'Timestamps (ms)', 'Ratio 0/1', 'Top Trigram',
];

// ───────── POST: write a single game ─────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_(DATA_SHEET);
    const headers = COLUMNS.map(c => c[0]);
    ensureHeaders_(sheet, headers);
    const row = COLUMNS.map(c => {
      let v = safeCall_(c[1], data);
      // Prefix CSV/sequence-like fields with apostrophe — bulletproof "this is text"
      // marker for Sheets. Hidden in display, stripped on getValues() read.
      if (TEXT_COLS.indexOf(c[0]) >= 0 && typeof v === 'string' && v.length > 0) {
        v = "'" + v;
      }
      return v;
    });
    sheet.appendRow(row);
    CacheService.getScriptCache().remove(STATS_CACHE_KEY);
    return jsonOut_({ status: 'ok' });
  } catch (err) {
    logError_(err, 'doPost', e);
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

// ───────── GET: stats endpoint or ping ─────────
function doGet(e) {
  const action   = (e && e.parameter && e.parameter.action)   || '';
  const callback = (e && e.parameter && e.parameter.callback) || '';

  if (action === 'stats') {
    let payload;
    try {
      payload = computeStats_();
    } catch (err) {
      logError_(err, 'doGet/stats');
      payload = {};
    }
    const json = JSON.stringify({ status: 'ok', stats: payload });
    if (callback && /^[a-zA-Z_$][\w$]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  return jsonOut_({ status: 'ok', message: 'Endpoint live. POST=write, GET?action=stats=read.' });
}

// ───────── Stats aggregation ─────────
function computeStats_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(STATS_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const sheet = getOrCreateSheet_(DATA_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { accuracy: {}, balance: {} };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lvlCol = headers.indexOf('Level');
  const accCol = headers.indexOf('Accuracy %');
  const balCol = headers.indexOf('Balance');
  const wonCol = headers.indexOf('Won');
  if (lvlCol < 0 || accCol < 0) return { accuracy: {}, balance: {} };

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const acc = { 1: [], 2: [], 3: [], 4: [], 5: [] };
  const bal = { 1: [], 2: [], 3: [], 4: [], 5: [] };
  const endgameBal = [];
  for (let i = 0; i < data.length; i++) {
    const lvl = parseInt(data[i][lvlCol]);
    const a = parseFloat(data[i][accCol]);
    const b = balCol >= 0 ? parseFloat(data[i][balCol]) : NaN;
    const won = wonCol >= 0 ? data[i][wonCol] : '';
    if (lvl >= 1 && lvl <= 5) {
      if (!isNaN(a) && a > 0 && a < 100) acc[lvl].push(a);
      if (!isNaN(b)) bal[lvl].push(b);
      if (won === 'YES' && lvl === 5 && !isNaN(b)) endgameBal.push(b);
    }
  }

  const out = { accuracy: {}, balance: {} };
  for (const lvl in acc) {
    if (acc[lvl].length >= 5) out.accuracy[lvl] = bucket_(acc[lvl]);
  }
  for (const lvl in bal) {
    if (bal[lvl].length >= 5) out.balance[lvl] = bucket_(bal[lvl]);
  }
  if (endgameBal.length >= 5) out.balance.endgame = bucket_(endgameBal);

  cache.put(STATS_CACHE_KEY, JSON.stringify(out), STATS_CACHE_TTL);
  return out;
}

function bucket_(arr) {
  return {
    n: arr.length,
    p: [pct_(arr, 10), pct_(arr, 25), pct_(arr, 50), pct_(arr, 75), pct_(arr, 90)],
  };
}

function pct_(arr, p) {
  const s = arr.slice().sort((a, b) => a - b);
  const k = (s.length - 1) * p / 100;
  const f = Math.floor(k);
  const c = Math.min(f + 1, s.length - 1);
  return Math.round((s[f] + (s[c] - s[f]) * (k - f)) * 10) / 10;
}

// ───────── Helpers ─────────
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet, expected) {
  if (sheet.getLastRow() === 0) {
    // First-ever write: create headers AND apply text format
    sheet.appendRow(expected);
    sheet.getRange(1, 1, 1, expected.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    applyTextFormat_(sheet, expected);
    return;
  }
  // Existing sheet: only sync headers if they drift. DO NOT re-apply text
  // format on every POST — that was rewriting numberFormat on N×10 cells
  // each write, which scaled badly and caused multi-second response times
  // under viral load. Format is set once at creation; new rows inherit it.
  let dirty = false;
  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expected.length)).getValues()[0];
  for (let i = 0; i < expected.length; i++) {
    if ((current[i] || '') !== expected[i]) {
      sheet.getRange(1, i + 1).setValue(expected[i]).setFontWeight('bold');
      dirty = true;
    }
  }
  if (dirty) applyTextFormat_(sheet, expected);
}

// Force TEXT format on columns whose values Sheets might autoparse as numbers.
// Idempotent — safe to call multiple times.
function applyTextFormat_(sheet, headers) {
  headers = headers || sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < TEXT_COLS.length; i++) {
    const idx = headers.indexOf(TEXT_COLS[i]);
    if (idx >= 0) {
      sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
    }
  }
}

// One-time migration: format existing Data sheet columns as TEXT.
// Run manually from Apps Script editor (Select function: migrateFormat → Run).
// Will NOT recover rows already corrupted by autoparse — those need to be re-collected.
function migrateFormat() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sheet) {
    Logger.log('Data sheet not found.');
    return;
  }
  applyTextFormat_(sheet);
  Logger.log('Text format applied to: ' + TEXT_COLS.join(', '));
}

function safeCall_(fn, data) {
  try { return fn(data); } catch (e) { return ''; }
}

function valOrEmpty_(v) {
  return v === null || v === undefined || v === '' ? '' : v;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError_(err, where, e) {
  try {
    const sheet = getOrCreateSheet_(ERR_SHEET);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Where', 'Error', 'Payload']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const payload = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    sheet.appendRow([new Date().toISOString(), where || '', String(err), payload.slice(0, 5000)]);
  } catch (_) { /* swallow */ }
}
