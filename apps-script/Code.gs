/**
 * FSS + KSSN Support Workspace — connector for the "FSS + KSSN Support Tracker" Google Sheet.
 * Paste into Extensions > Apps Script of that sheet, add an HTML file named "Index"
 * containing apps-script/Index.html, then Deploy > Web app
 * (Execute as: Me · Who has access: Anyone within Grand Rapids Public Schools).
 */
var FU_SHEET = 'FollowUps';
var KEYS = ['id','created','school','code','barrier','leadRole','leadPerson','action','resource','due','status','outcome','updated','updatedBy'];
var DATE_KEYS = {created:1, due:1, updated:1};
var STATUSES = ['Open','In progress','Done','Stuck'];
var ROLES = ['FSS','KSSN','School','Shared'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FSS + KSSN Support Workspace')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}
function safe_(v) { // keep text from being read as a formula
  v = String(v == null ? '' : v).slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function rows_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

function getData() {
  var schools = rows_('Schools').filter(function (r) { return r[1]; }).map(function (r) {
    return {quad: cell_(r[0]), school: cell_(r[1]), kssn: cell_(r[2]), kssnEmail: cell_(r[3]), kssnPhone: cell_(r[4]), lead: cell_(r[5]),
            leadEmail: cell_(r[6]), leadPhone: cell_(r[7]), fss: cell_(r[8]), fssEmail: cell_(r[9]), fssPhone: cell_(r[10]), fssRole: cell_(r[11]), notes: cell_(r[12])};
  });
  var resources = rows_('Resources').filter(function (r) { return r[0]; }).map(function (r) {
    return {name: cell_(r[0]), needs: cell_(r[1]).split(/\s*,\s*/).filter(String), phone: cell_(r[2]), url: cell_(r[3]),
            how: cell_(r[4]), lang: cell_(r[5]), source: cell_(r[6]), checked: cell_(r[7])};
  });
  var followups = rows_(FU_SHEET).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; KEYS.forEach(function (k, i) { o[k] = cell_(r[i]); }); return o;
  });
  return {schools: schools, resources: resources, followups: followups, user: Session.getActiveUser().getEmail()};
}

function saveFollowUp(fu) {
  var code = String(fu.code || '').trim().toUpperCase();
  if (!/^[A-Z]{0,3}-?\d{1,5}$/.test(code)) throw new Error('Use a family code like A-14, never a name.');
  if (STATUSES.indexOf(fu.status) < 0) fu.status = 'Open';
  if (ROLES.indexOf(fu.leadRole) < 0) fu.leadRole = 'Shared';
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(FU_SHEET);
    var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var row = fu.id ? ids.indexOf(String(fu.id)) : -1;
    if (!fu.id || row < 0) { fu.id = 'F' + Utilities.getUuid().slice(0, 8); fu.created = fu.created || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
    fu.code = code;
    fu.updated = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    fu.updatedBy = Session.getActiveUser().getEmail() || 'GRPS staff';
    var values = KEYS.map(function (k) { return DATE_KEYS[k] ? (fu[k] || '') : safe_(fu[k]); });
    if (row >= 0) sh.getRange(row + 2, 1, 1, KEYS.length).setValues([values]);
    else sh.appendRow(values);
    var out = {}; KEYS.forEach(function (k) { out[k] = String(fu[k] == null ? '' : fu[k]); });
    return out;
  } finally { lock.releaseLock(); }
}
