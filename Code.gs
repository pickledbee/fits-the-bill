/**
 * Fits the Bill - shared room for the card game.
 *
 * How to use:
 *   1. Open the Compare Pears sheet, click Extensions > Apps Script, and paste this whole file in.
 *   2. Run setup() once. It builds the Tables tab with all headings.
 *   3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *   4. Copy the web app URL (it ends in /exec) into the game.
 *
 * The game page reads the room with GET and sends each move with POST.
 * Every move is merged into the table's saved state under a lock, so two
 * people clicking at the same moment cannot overwrite each other.
 */

var SHEET_ID = '1fYxP5gVX9sSlASenULD4oicq5LkN8R3xTV5rix0sLSY'; // the Compare Pears sheet
var SHEET_NAME = 'Tables';
var HEADERS = ['Table ID', 'Table Name', 'Last Updated', 'Chairs Taken', 'Round', 'Players', 'State (JSON, do not edit)'];
var TABLES = [['t1', 'Window Table'], ['t2', 'Corner Table'], ['t3', 'Kitchen Table'], ['t4', 'Back Table']];
var CACHE_KEY = 'ftb-tables-v1';
var MAX_BODY = 60000;

/** Run once. Builds (or rebuilds) the Tables tab. Safe to run again to wipe every table. */
function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#1d5f8c').setFontColor('#ffffff');
  var rows = TABLES.map(function (t) { return [t[0], t[1], '', 0, 0, '', '']; });
  sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 2, 120);
  sh.setColumnWidth(3, 160);
  sh.setColumnWidths(4, 2, 100);
  sh.setColumnWidth(6, 260);
  sh.setColumnWidth(7, 420);
  sh.getRange(2, 7, rows.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  CacheService.getScriptCache().remove(CACHE_KEY);
}

/** Run any time to empty every chair and start fresh. */
function clearAllTables() { setup(); }

function doGet() {
  return reply({ ok: true, tables: readTables() });
}

function doPost(e) {
  var req;
  try {
    var raw = e.postData.contents;
    if (raw.length > MAX_BODY) return reply({ ok: false, error: 'too big' });
    req = JSON.parse(raw);
  } catch (err) {
    return reply({ ok: false, error: 'bad request' });
  }
  var row = -1;
  for (var i = 0; i < TABLES.length; i++) if (TABLES[i][0] === req.tid) row = i;
  if (req.a !== 'write' || row < 0 || !isObj(req.patch)) return reply({ ok: false, error: 'bad request' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return reply({ ok: false, error: 'busy' }); }
  try {
    var tables = readTables(true);
    var next = req.replace ? req.patch : merge(tables[req.tid] || {}, req.patch);
    var json = JSON.stringify(next);
    if (json.length > 45000) return reply({ ok: false, error: 'too big' });
    tables[req.tid] = next;
    var seats = next.seats || {}, names = [];
    for (var s in seats) if (seats[s] && seats[s].uid) names.push(String(seats[s].name || 'Guest') + (seats[s].bot ? ' (house)' : ''));
    sheet().getRange(row + 2, 3, 1, 5).setValues([[new Date(), names.length, Number(next.round) || 0, names.join(', '), json]]);
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(tables), 21600);
    return reply({ ok: true, tables: tables });
  } finally {
    lock.releaseLock();
  }
}

function sheet() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Run setup() first.');
  return sh;
}

/** Every table's saved state. Served from cache when possible so polling stays fast. */
function readTables(fresh) {
  var cache = CacheService.getScriptCache();
  if (!fresh) {
    var hit = cache.get(CACHE_KEY);
    if (hit) { try { return JSON.parse(hit); } catch (err) {} }
  }
  var values = sheet().getRange(2, 1, TABLES.length, HEADERS.length).getValues();
  var out = {};
  for (var i = 0; i < TABLES.length; i++) {
    var cell = values[i][6];
    out[TABLES[i][0]] = null;
    if (cell) { try { out[TABLES[i][0]] = JSON.parse(cell); } catch (err) {} }
  }
  if (!fresh) cache.put(CACHE_KEY, JSON.stringify(out), 21600);
  return out;
}

function isObj(v) { return v && typeof v === 'object' && !Array.isArray(v); }

/** Deep merge: nested objects merge key by key, everything else (arrays included) is replaced. */
function merge(a, b) {
  var out = {}, k;
  for (k in a) out[k] = a[k];
  for (k in b) out[k] = isObj(b[k]) ? merge(isObj(out[k]) ? out[k] : {}, b[k]) : b[k];
  return out;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
