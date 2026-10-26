/* ===================== ArNote starter — please don't edit =====================
 * This small script never changes. Each time ArNote runs, it loads the ArNote code kept in
 * "ArNote code.js" in your Google Drive (your copy downloads new versions by itself once a day)
 * and hands over to it. Your notes and photos never leave your own Google Drive.
 */
var ARNOTE_HOME = 'https://script.google.com/macros/s/AKfycbyFPAGYJE1bVycoBv8QamFeUHuBjbCAhogOM8mSU7vn20CVO_Ex1XJYfRy7f7lZES-_Kg/exec';
var ARNOTE_LOADED_ = false, ARNOTE_STUBS_ = {};

function arnoteCode_() {
  var cache = CacheService.getScriptCache(), n = +cache.get('arnote_code_n'), src = null;
  if (n) {
    var keys = []; for (var i = 0; i < n; i++) keys.push('arnote_code_' + i);
    var got = cache.getAll(keys); if (keys.every(function (k) { return got[k] !== undefined && got[k] !== null; })) src = keys.map(function (k) { return got[k]; }).join('');
  }
  if (!src) {
    var id = PropertiesService.getScriptProperties().getProperty('codeFile');
    if (!id) return null;
    src = DriveApp.getFileById(id).getBlob().getDataAsString();
    try {
      var parts = {}, size = 90000, k = 0;
      for (var j = 0; j < src.length; j += size) parts['arnote_code_' + (k++)] = src.slice(j, j + size);
      cache.putAll(parts, 21600); cache.put('arnote_code_n', String(k), 21600);
    } catch (e) {}
  }
  return src;
}
function arnoteLoad_() {
  if (ARNOTE_LOADED_) return;
  var src = arnoteCode_();
  if (!src) throw new Error('ArNote isn\'t set up yet: open the ArNote spreadsheet and choose ArNote → Set up ArNote.');
  try { (0, eval)(src); }              // defines ArNote's functions (they replace the stubs below)
  catch (err) {                        // damaged code: fetch a fresh copy from Brian's ArNote and try once more
    var fresh = arnoteFetch_(); if (!fresh) throw err;
    arnoteSaveCode_(fresh); (0, eval)(fresh.code);
  }
  ARNOTE_LOADED_ = true;
}
/** Ask Brian's ArNote for the current release (only answers addresses on his list). */
function arnoteFetch_() {
  var p = PropertiesService.getScriptProperties();
  var inst = p.getProperty('installId'); if (!inst) { inst = Utilities.getUuid(); p.setProperty('installId', inst); }
  var r = UrlFetchApp.fetch(ARNOTE_HOME, { method: 'post', contentType: 'text/plain', muteHttpExceptions: true,
    payload: JSON.stringify({ lic: 'code', email: Session.getEffectiveUser().getEmail(), install: inst, build: 0 }) });
  var body = null; try { body = (JSON.parse(r.getContentText()) || {}).result; } catch (e) {}
  return body ? JSON.parse(body) : null;
}
function arnoteSaveCode_(rel) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('codeFile'), f = null;
  if (id) { try { f = DriveApp.getFileById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (f) f.setContent(rel.code); else { f = DriveApp.getRootFolder().createFile('ArNote code.js', rel.code, 'text/plain'); }
  p.setProperties({ codeFile: f.getId(), releaseInstalled: String(rel.n), updatedAt: new Date().toISOString() });
  try { CacheService.getScriptCache().remove('arnote_code_n'); } catch (e) {}
}
function arnoteRun_(name, args) {
  arnoteLoad_();
  var f = globalThis[name];
  if (typeof f !== 'function' || f === ARNOTE_STUBS_[name]) throw new Error('ArNote code has no ' + name);
  return f.apply(null, args || []);
}

/* the ways in: Google calls these; each hands over to the real ArNote code */
function doGet(e) { return arnoteRun_('doGet', [e]); }
function doPost(e) { return arnoteRun_('doPost', [e]); }
function processDrop() { return arnoteRun_('processDrop', []); }
function importBatch() { return arnoteRun_('importBatch', []); }
function startImport() { return arnoteRun_('startImport', []); }
function showImportStatus() { return arnoteRun_('showImportStatus', []); }
function showAppLink() { return arnoteRun_('showAppLink', []); }
function authorizeMail() { return arnoteRun_('authorizeMail', []); }
function onOpen() {
  try { if (PropertiesService.getScriptProperties().getProperty('codeFile')) return arnoteRun_('onOpen', []); } catch (e) {}
  var m = SpreadsheetApp.getUi().createMenu('ArNote');
  if (arnoteIsSetUp_()) m.addItem('Open the app', 'showAppLink').addItem('Start / resume Evernote import', 'startImport').addItem('Import status', 'showImportStatus');
  else m.addItem('Set up ArNote', 'arnoteFirstSetup');
  m.addToUi();
}
function arnoteIsSetUp_() { try { return !!PropertiesService.getScriptProperties().getProperty('codeFile'); } catch (e) { return false; } }
['doGet', 'doPost', 'processDrop', 'importBatch', 'startImport', 'showImportStatus', 'showAppLink', 'authorizeMail', 'onOpen'].forEach(function (n) { ARNOTE_STUBS_[n] = globalThis[n]; });

/** First time: fetch ArNote from Brian's (only works if you're on his list), then set everything up. */
function arnoteFirstSetup() {
  var ui = null; try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  if (arnoteIsSetUp_()) { var done = 'ArNote is already set up here. Choose ArNote \u2192 Open the app.'; if (ui) ui.alert(done); return done; }
  var rel = arnoteFetch_();
  if (!rel) { var msg = Session.getEffectiveUser().getEmail() + ' isn\'t on the ArNote list yet. Ask the person who invited you to add this exact address, then try again.'; if (ui) ui.alert(msg); return msg; }
  arnoteSaveCode_(rel);
  arnoteLoad_();
  var out = arnoteRun_('arnoteSetupAll_', []);       // folders, sheets, starter tags, private key, 15-minute job; moves the code file into the ArNote folder
  if (ui) ui.alert('ArNote is set up', out, ui.ButtonSet.OK);
  return out;
}
