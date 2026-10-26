var TAG_TREE = '';   // first-time tags now come from STARTER_TAGS_ (Batch.gs)
/**
 * ArNote — personal note system (Google Apps Script, bound to the "ArNote" Sheet).
 * Notes are Google Docs in ArNote/Notes; attachments are files in ArNote/Attachments;
 * this Sheet is the index (Notes, Tags, Attachments tabs).
 */

var NOTE_COLS = ['id', 'title', 'notebook', 'tags', 'created', 'updated', 'docId', 'starred', 'source', 'sourceUrl', 'author', 'attachments'];
var TAG_COLS = ['name', 'parent'];
var ATT_COLS = ['noteId', 'fileId', 'filename', 'mime', 'bytes', 'hash'];
var TAG_SEP = '; ';

/* ---------------- menu & setup ---------------- */

function onOpen() {
  var m = SpreadsheetApp.getUi().createMenu('ArNote');
  if (typeof ARNOTE_STUBS_ === 'undefined') m.addItem('Set up (first time)', 'setup');   // copies are set up by the starter
  m.addItem('Open the app', 'showAppLink')
    .addItem('Start / resume Evernote import', 'startImport')
    .addItem('Import status', 'showImportStatus')
    .addToUi();
}

function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(props_().getProperty('sheetId')); }

function setup() {
  var p = props_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  p.setProperty('sheetId', ss.getId());
  if (ss.getName() !== 'ArNote') ss.rename('ArNote');
  var root = getOrCreateFolder_(p.getProperty('rootId'), 'ArNote', DriveApp.getRootFolder());
  p.setProperty('rootId', root.getId());
  p.setProperty('notesId', getOrCreateFolder_(p.getProperty('notesId'), 'Notes', root).getId());
  p.setProperty('attId', getOrCreateFolder_(p.getProperty('attId'), 'Attachments', root).getId());
  // move the index sheet into the ArNote folder
  try { DriveApp.getFileById(ss.getId()).moveTo(root); } catch (e) {}
  ensureSheet_('Notes', NOTE_COLS);
  var tags = ensureSheet_('Tags', TAG_COLS);
  ensureSheet_('Attachments', ATT_COLS);
  var s1 = ss.getSheetByName('Sheet1'); if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  if (tags.getLastRow() < 2) seedTags_(tags);
  return 'ArNote is set up. Folder: ' + root.getUrl();
}

function getOrCreateFolder_(id, name, parent) {
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function ensureSheet_(name, cols) {
  var ss = ss_();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function seedTags_(sh) {
  if (!TAG_TREE) return;
  var stack = [], rows = [];
  TAG_TREE.split('\n').forEach(function (line) {
    var i = line.indexOf('|'); var depth = +line.slice(0, i); var name = line.slice(i + 1);
    stack[depth] = name; stack.length = depth + 1;
    rows.push([name, depth ? stack[depth - 1] : '']);
  });
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

/* "Open the app": shows the pairing link. Google doesn't reliably tell a script its own published address
 * (it can give a different one, e.g. with /a/yourdomain.com/), so the first time we ask for it once, test it, and remember it. */
var WEBAPP_URL_RE_ = /^https:\/\/script\.google\.com\/(a\/[^\/]+\/)?macros\/s\/[A-Za-z0-9_-]{20,}\/exec$/;
function api_setAppUrl(url) {
  url = String(url || '').trim();
  if (!WEBAPP_URL_RE_.test(url)) throw new Error('That isn\'t a web app address');
  props_().setProperty('appUrl', url);
  return true;
}
function showAppLink() {
  var url = props_().getProperty('appUrl') || '', key = props_().getProperty('apiKey') || '';
  var html = '<div style="font:14px sans-serif" id="w"></div><script>' +
    'var KEY=' + JSON.stringify(key) + ',APP=' + JSON.stringify(APP_URL) + ',URL0=' + JSON.stringify(url) + ',RE=' + WEBAPP_URL_RE_.toString() + ';' +
    'function esc(x){return String(x).replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}[c]})}' +
    'function link(u){return APP+"#pair="+btoa(u+"|"+KEY).replace(/\\+/g,"-").replace(/\\//g,"_").replace(/=+$/,"")}' +
    'function call(u,fn,a){return fetch(u,{method:"POST",body:JSON.stringify({k:KEY,fn:fn,a:a||[]})}).then(function(r){return r.text()}).then(function(t){var j;try{j=JSON.parse(t)}catch(e){throw new Error("no ArNote at that address")}if(!j.ok)throw new Error(j.error);return j.result})}' +
    'function show(u){var l=link(u);document.getElementById("w").innerHTML="<p>Open this <b>pairing link</b> on each phone and computer you\'ll use (email it to yourself). Keep it private \\u2014 it is the key to your notes.</p>' +
      '<textarea readonly style=\\"width:100%;height:70px;font:12px monospace\\" onclick=\\"this.select()\\">"+esc(l)+"</textarea><p><a href=\\""+esc(l)+"\\" target=\\"_blank\\">Open ArNote here</a> &nbsp; <a href=\\"#\\" onclick=\\"ask(\'\');return false\\" style=\\"font-size:12px;color:#666\\">change server address</a></p>"}' +
    'function ask(msg){document.getElementById("w").innerHTML="<p><b>One-time step:</b> paste your ArNote web app address.</p><p style=\\"font-size:13px;color:#444\\">Find it: <b>Extensions \\u2192 Apps Script \\u2192 Deploy \\u2192 Manage deployments</b>, then click <b>Copy</b> under Web app URL. (Not published yet? In that window use <b>Deploy \\u2192 New deployment</b> \\u2192 \\u2699 <b>Web app</b>, Execute as <b>Me</b>, Who has access <b>Anyone</b>.)</p>' +
      '<input id=\\"u\\" style=\\"width:100%;box-sizing:border-box;font:12px monospace;padding:6px\\" placeholder=\\"https://script.google.com/macros/s/\\u2026/exec\\"><p><button id=\\"b\\" style=\\"padding:6px 14px\\">Get my pairing link</button> <span id=\\"m\\" style=\\"color:#b00\\">"+esc(msg||"")+"</span></p>";' +
      'document.getElementById("b").onclick=function(){var u=document.getElementById("u").value.trim(),m=document.getElementById("m");if(!RE.test(u)){m.textContent="That doesn\'t look like a web app address (it ends in /exec).";return}' +
      'm.style.color="#666";m.textContent="Checking\\u2026";call(u,"api_appInfo").then(function(){return call(u,"api_setAppUrl",[u])}).then(function(){show(u)}).catch(function(e){m.style.color="#b00";m.textContent="Couldn\'t reach ArNote there: "+e.message+". Check it is published for Anyone."})}}' +
    'if(URL0)show(URL0);else ask("");</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(280), 'ArNote');
}

/* ---------------- sheet helpers ---------------- */

function readTable_(name) {
  var sh = ss_().getSheetByName(name);
  var v = sh.getDataRange().getValues();
  var head = v.shift();
  return v.map(function (r, i) { var o = { _row: i + 2 }; head.forEach(function (h, j) { o[h] = r[j]; }); return o; });
}

function noteToRow_(n) { return NOTE_COLS.map(function (c) { return n[c] === undefined ? '' : n[c]; }); }

function findNoteRow_(id) {
  var sh = ss_().getSheetByName('Notes');
  var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  throw new Error('Note not found: ' + id);
}

function enDate_(s) { // 20190124T195325Z -> Date
  if (!s) return '';
  var m = /^(\d{4})(\d\d)(\d\d)T(\d\d)(\d\d)(\d\d)Z$/.exec(s);
  return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6])) : '';
}

function headerLine_(notebook, tags) {
  return 'Notebook: ' + notebook + '  |  Tags: ' + (tags && tags.length ? tags.join(TAG_SEP) : '(none)');
}
/* ---------------- Evernote import (one-time) ----------------
 * Expects a Drive folder "ArNote-import" containing notes.jsonl and attachments/<md5>.<ext>
 * (made by arenote_extract.ps1 on the laptop). Runs in ~5-minute batches and re-schedules
 * itself until done. Safe to re-run: it resumes where it stopped.
 */

var IMG_PH = '⟦IMG:';   // ⟦IMG:hash⟧ placeholder
var MAX_INLINE_BYTES = 15 * 1024 * 1024;
var BATCH_MS = 4.5 * 60 * 1000;

function startImport() {
  var p = props_();
  if (!p.getProperty('rootId')) setup();
  if (!p.getProperty('importDone')) {
    clearImportTriggers_();
    importBatch();
  }
  var msg = importStatusText_();
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

function showImportStatus() { SpreadsheetApp.getUi().alert(importStatusText_()); }

function importStatusText_() {
  var p = props_();
  var total = +p.getProperty('importTotal') || 0, next = +p.getProperty('importNext') || 0;
  if (p.getProperty('importDone')) return 'Import finished: ' + total + ' notes.';
  if (!total) return 'Import not started.';
  return 'Import running: ' + next + ' of ' + total + ' notes done. ' + (p.getProperty('importLastError') || '');
}

function clearImportTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'importBatch') ScriptApp.deleteTrigger(t); });
}

function importFolder_() {
  var it = DriveApp.getFoldersByName('ArNote-import');
  while (it.hasNext()) { var f = it.next(); if (!f.isTrashed() && f.getFilesByName('notes.jsonl').hasNext()) return f; }
  throw new Error('Could not find the "ArNote-import" folder (with notes.jsonl) in Drive. Upload it first.');
}

function loadNotesJsonl_() {
  var blob = importFolder_().getFilesByName('notes.jsonl').next().getBlob();
  var text = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  return text.split(/\r?\n/).filter(function (l) { return l.trim(); }).map(function (l) { return JSON.parse(l); });
}

function arr_(x) { // tolerate PowerShell 5.1 JSON shapes
  if (x === null || x === undefined || x === '') return [];
  if (Array.isArray(x)) return x;
  if (x.value && Array.isArray(x.value)) return x.value;
  return [x];
}

/** Index attachment files once: hash -> {id, name}; moves them into ArNote/Attachments. */
function attachmentIndex_() {
  var p = props_();
  var cache = CacheService.getScriptCache();
  var dest = DriveApp.getFolderById(p.getProperty('attId'));
  if (!p.getProperty('attMoved')) {
    // Move the uploaded attachments folder as a whole (fast) instead of file by file.
    var srcIt = importFolder_().getFoldersByName('attachments');
    if (srcIt.hasNext()) {
      var src = srcIt.next(), start = Date.now();
      var count = function (f) { var it = f.getFiles(), n = 0; while (it.hasNext()) { it.next(); n++; } return n; };
      if (count(dest) <= count(src)) {
        // bring back any files an earlier run already moved, then move the folder itself
        var back = dest.getFiles();
        while (back.hasNext()) { back.next().moveTo(src); if (Date.now() - start > 3 * 60 * 1000) return null; }
        src.moveTo(DriveApp.getFolderById(p.getProperty('rootId')));
        src.setName('Attachments');
        dest.setTrashed(true);
        p.setProperty('attId', src.getId());
        dest = src;
      } else {
        var files = src.getFiles();
        while (files.hasNext()) { files.next().moveTo(dest); if (Date.now() - start > 3 * 60 * 1000) return null; }
      }
    }
    p.setProperty('attMoved', '1');
  }
  var map = {}, chunks = +p.getProperty('attMapChunks') || 0;
  if (chunks) {
    var got = cache.getAll(Array.apply(null, Array(chunks)).map(function (_, i) { return 'attmap' + i; }));
    if (Object.keys(got).length === chunks) {
      Object.keys(got).forEach(function (k) { var m = JSON.parse(got[k]); for (var h in m) map[h] = m[h]; });
      return map;
    }
  }
  var it = dest.getFiles();
  while (it.hasNext()) { var f = it.next(); var nm = f.getName(); var h = nm.replace(/\.[^.]*$/, ''); map[h] = f.getId(); }
  // cache in 90KB chunks for 6 hours
  var keys = Object.keys(map), out = {}, cur = {}, size = 0, n = 0;
  keys.forEach(function (h) { cur[h] = map[h]; size += h.length + map[h].length + 6; if (size > 90000) { out['attmap' + n++] = JSON.stringify(cur); cur = {}; size = 0; } });
  out['attmap' + n++] = JSON.stringify(cur);
  cache.putAll(out, 21600); p.setProperty('attMapChunks', String(n));
  return map;
}

function importBatch() {
  var start = Date.now(), p = props_();
  clearImportTriggers_();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    if (p.getProperty('importDone')) return;
    var attMap = attachmentIndex_();
    if (!attMap) { scheduleNext_(); return; }
    var notes = loadNotesJsonl_();
    p.setProperty('importTotal', String(notes.length));
    var i = +p.getProperty('importNext') || 0;
    // crash recovery: a doc created for note i but row not written
    var pending = p.getProperty('importPending');
    if (pending) { var pd = JSON.parse(pending); if (pd.i === i) { try { DriveApp.getFileById(pd.docId).setTrashed(true); } catch (e) {} } p.deleteProperty('importPending'); }
    var notesSh = ss_().getSheetByName('Notes'), attSh = ss_().getSheetByName('Attachments');
    var knownTags = {}; readTable_('Tags').forEach(function (t) { knownTags[t.name] = 1; });
    while (i < notes.length && Date.now() - start < BATCH_MS) {
      var r = importOne_(notes[i], i, attMap, knownTags);
      notesSh.appendRow(noteToRow_(r.note));
      if (r.atts.length) attSh.getRange(attSh.getLastRow() + 1, 1, r.atts.length, ATT_COLS.length).setValues(r.atts);
      p.deleteProperty('importPending');
      i++; p.setProperty('importNext', String(i));
    }
    p.deleteProperty('importLastError');
    if (i >= notes.length) { p.setProperty('importDone', new Date().toISOString()); return; }
    scheduleNext_();
  } catch (e) {
    p.setProperty('importLastError', 'Last error: ' + e.message);
    scheduleNext_();
    throw e;
  } finally { lock.releaseLock(); }
}

function scheduleNext_() { ScriptApp.newTrigger('importBatch').timeBased().after(60 * 1000).create(); }

function importOne_(n, i, attMap, knownTags) {
  var p = props_();
  var tags = arr_(n.tags), resources = arr_(n.resources), attrs = n.attributes || {};
  tags.forEach(function (t) { if (!knownTags[t]) { ss_().getSheetByName('Tags').appendRow([t, '']); knownTags[t] = 1; } });
  var title = (n.title || '').trim() || 'Untitled Note';
  var id = Utilities.getUuid();
  var resByHash = {}; resources.forEach(function (r) { if (r.hash) resByHash[r.hash] = r; });
  var conv = enmlToHtml_(n.enml || '', resByHash, attMap);
  var used = conv.used;
  // attachments not referenced in the body
  var extra = resources.filter(function (r) { return r.hash && !used[r.hash]; });
  var html = '<html><head><meta charset="utf-8"></head><body>' +
    '<p style="color:#888888;font-size:9pt">' + esc_(headerLine_(n.notebook, tags)) + '</p>' +
    (attrs['source-url'] ? '<p style="font-size:9pt"><a href="' + escAttr_(attrs['source-url']) + '">' + esc_(attrs['source-url']) + '</a></p>' : '') +
    conv.html;
  if (extra.length) {
    html += '<p><b>Attachments</b></p>';
    extra.forEach(function (r) { html += mediaHtml_(r, attMap, true); used[r.hash] = 1; });
  }
  html += '</body></html>';

  var docId = createDocFromHtml_(title, html, p.getProperty('notesId'), enDate_(n.updated));
  p.setProperty('importPending', JSON.stringify({ i: i, docId: docId }));
  if (Object.keys(conv.images).length || extra.some(isInlineImage_)) insertImages_(docId, resByHash, attMap);

  var atts = resources.filter(function (r) { return r.hash; }).map(function (r) {
    return [id, attMap[r.hash] || '', r.filename || (r.hash + (r.ext || '')), r.mime || '', r.bytes || 0, r.hash];
  });
  var note = { id: id, title: title, notebook: n.notebook, tags: tags.join(TAG_SEP), created: enDate_(n.created), updated: enDate_(n.updated),
    docId: docId, starred: false, source: attrs.source || '', sourceUrl: attrs['source-url'] || '', author: attrs.author || '', attachments: atts.length };
  return { note: note, atts: atts };
}

function isInlineImage_(r) { return /^image\/(jpeg|png|gif)$/.test(r.mime || '') && (r.bytes || 0) <= MAX_INLINE_BYTES; }

function mediaHtml_(r, attMap, block) {
  if (isInlineImage_(r) && attMap[r.hash]) return (block ? '<p>' : '') + IMG_PH + r.hash + '⟧' + (block ? '</p>' : '');
  var fid = attMap[r.hash];
  var label = '📎 ' + esc_(r.filename || ('attachment' + (r.ext || ''))) + ' (' + Math.round((r.bytes || 0) / 1024) + ' KB)';
  var a = fid ? '<a href="https://drive.google.com/file/d/' + fid + '/view">' + label + '</a>' : label + ' [missing]';
  return block ? '<p>' + a + '</p>' : a;
}

function enmlToHtml_(enml, resByHash, attMap) {
  var used = {}, images = {};
  var s = enml.replace(/<\?xml[^>]*\?>/g, '').replace(/<!DOCTYPE[^>]*>/gi, '');
  var m = /<en-note[^>]*>([\s\S]*)<\/en-note>/i.exec(s);
  s = m ? m[1] : s;
  s = s.replace(/<div[^>]*display:\s*none[^>]*>[\s\S]*?<\/div>/gi, '');           // hidden Evernote metadata divs
  s = s.replace(/<en-media\b([^>]*?)\/?>(\s*<\/en-media>)?/gi, function (_, attrs) {
    var h = (/hash="([0-9a-f]+)"/i.exec(attrs) || [])[1];
    var r = h && resByHash[h];
    if (!r) return '';
    used[h] = 1;
    if (isInlineImage_(r) && attMap[h]) images[h] = 1;
    return mediaHtml_(r, attMap, false);
  });
  s = s.replace(/<en-todo\b[^>]*checked="true"[^>]*\/?>(<\/en-todo>)?/gi, '☑ ')
       .replace(/<en-todo\b[^>]*\/?>(<\/en-todo>)?/gi, '☐ ')
       .replace(/<en-crypt\b[\s\S]*?<\/en-crypt>/gi, '[encrypted text — not imported]');
  return { html: s, used: used, images: images };
}

function createDocFromHtml_(title, html, folderId, modified) {
  var boundary = 'arenote' + Utilities.getUuid().replace(/-/g, '');
  var meta = { name: title, mimeType: 'application/vnd.google-apps.document', parents: [folderId] };
  if (modified) meta.modifiedTime = modified.toISOString();
  var body = Utilities.newBlob(
    '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
    '\r\n--' + boundary + '\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n').getBytes()
    .concat(Utilities.newBlob(html, 'text/html').getBytes())
    .concat(Utilities.newBlob('\r\n--' + boundary + '--').getBytes());
  for (var attempt = 0; attempt < 4; attempt++) {
    var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
      method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload: body,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    var code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText()).id;
    if (code === 429 || code >= 500) { Utilities.sleep(2000 * (attempt + 1)); continue; }
    throw new Error('Doc create failed (' + code + '): ' + res.getContentText().slice(0, 300));
  }
  throw new Error('Doc create failed after retries');
}

function insertImages_(docId, resByHash, attMap) {
  var doc = DocumentApp.openById(docId), body = doc.getBody();
  var pageW = 600, found, guard = 0;
  while ((found = body.findText('⟦IMG:[0-9a-f]+⟧')) && guard++ < 500) {
    var t = found.getElement().asText(), s = found.getStartOffset(), e = found.getEndOffsetInclusive();
    var h = t.getText().slice(s + 5, e);   // text between "⟦IMG:" and "⟧"
    var r = resByHash[h], blob = null;
    try { blob = DriveApp.getFileById(attMap[h]).getBlob(); } catch (err) {}
    var par = t.getParent();
    while (par && par.getType() !== DocumentApp.ElementType.PARAGRAPH && par.getType() !== DocumentApp.ElementType.LIST_ITEM && par.getType() !== DocumentApp.ElementType.TABLE_CELL) par = par.getParent();
    var whole = t.getText().length === (e - s + 1);
    var idx = whole && t.getParent().getType() !== DocumentApp.ElementType.TABLE_CELL ? t.getParent().getChildIndex(t) : -1;
    t.deleteText(s, e);
    if (!blob) continue;
    var img;
    try {
      if (idx >= 0) img = t.getParent().insertInlineImage(idx, blob);
      else if (par && par.appendInlineImage) img = par.appendInlineImage(blob);
      else if (par && par.appendParagraph) img = par.appendParagraph('').appendInlineImage(blob);
      if (img) {
        var w = img.getWidth(), hgt = img.getHeight();
        if (w > pageW) { img.setWidth(pageW).setHeight(Math.round(hgt * pageW / w)); }
      }
    } catch (err) {
      try { t.insertText(Math.min(s, t.getText().length), '📎 ' + ((r && r.filename) || 'image')).setLinkUrl('https://drive.google.com/file/d/' + attMap[h] + '/view'); } catch (e2) {}
    }
  }
  doc.saveAndClose();
}

function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function escAttr_(s) { return esc_(s).replace(/"/g, '&quot;'); }

/** Undo a test import completely (trashes created docs, clears rows, moves attachments back is NOT done). */
function resetImport_DANGER() {
  var p = props_();
  readTable_('Notes').forEach(function (n) { try { DriveApp.getFileById(n.docId).setTrashed(true); } catch (e) {} });
  ['Notes', 'Attachments'].forEach(function (nm) { var sh = ss_().getSheetByName(nm); if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); });
  ['importNext', 'importTotal', 'importDone', 'importPending', 'importLastError'].forEach(function (k) { p.deleteProperty(k); });
}
/* ---------------- web app ---------------- */

function doGet() {
  // The page itself is only for the owner; everyone else (the installable app) talks to doPost with a key.
  if (Session.getActiveUser().getEmail() !== Session.getEffectiveUser().getEmail())
    return ContentService.createTextOutput('ArNote server is running.');
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ArNote')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function iso_(d) { return d instanceof Date ? d.toISOString() : (d || ''); }

/** Everything the client needs to render lists and the tag tree. */
function api_bootstrap() {
  var notes = readTable_('Notes').map(function (n) {
    return { id: n.id, title: String(n.title), notebook: n.notebook, tags: n.tags ? String(n.tags).split(TAG_SEP) : [],
      created: iso_(n.created), updated: iso_(n.updated), docId: n.docId, starred: n.starred === true || n.starred === 'TRUE',
      att: +n.attachments || 0, sourceUrl: n.sourceUrl || '', link: n.source === 'link' };
  });
  var tags = readTable_('Tags').map(function (t) { return { name: String(t.name), parent: String(t.parent || ''), order: +t.order || 0 }; });
  var order = (props_().getProperty('bookmarkOrder') || '').split(',').filter(String);
  return { notes: notes, tags: tags, bookmarkOrder: order, user: Session.getEffectiveUser().getEmail() };
}

function api_getNote(id) {
  var n = readTable_('Notes').filter(function (x) { return x.id === id; })[0];
  if (!n) throw new Error('Note not found');
  var atts = readTable_('Attachments').filter(function (a) { return a.noteId === id; })
    .map(function (a) { return { fileId: a.fileId, filename: String(a.filename), mime: a.mime, bytes: +a.bytes }; });
  return { id: n.id, docId: n.docId, atts: atts };
}

/** Full-text search over note Docs and attachments (Drive search, incl. OCR of PDFs/images). Returns note ids. */
function api_search(q) {
  q = String(q || '').trim(); if (!q) return [];
  var p = props_(), safe = q.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  var query = "fullText contains '" + safe + "' and trashed = false and ('" + p.getProperty('notesId') + "' in parents or '" + p.getProperty('attId') + "' in parents)";
  var ids = {}, byDoc = {}, byFile = {};
  readTable_('Notes').forEach(function (n) { byDoc[n.docId] = n.id; });
  readTable_('Attachments').forEach(function (a) { (byFile[a.fileId] = byFile[a.fileId] || []).push(a.noteId); });
  var it = DriveApp.searchFiles(query), k = 0;
  while (it.hasNext() && k++ < 300) {
    var f = it.next(), fid = f.getId();
    if (byDoc[fid]) ids[byDoc[fid]] = 1;
    (byFile[fid] || []).forEach(function (nid) { ids[nid] = 1; });
  }
  // linked notes point at files elsewhere in Drive: search the whole Drive for those
  var linked = {}; readTable_('Notes').forEach(function (n) { if (n.source === 'link') linked[n.docId] = n.id; });
  if (Object.keys(linked).length) {
    var it2 = DriveApp.searchFiles("fullText contains '" + safe + "' and trashed = false"), k2 = 0;
    while (it2.hasNext() && k2++ < 300) { var f2 = it2.next(); if (linked[f2.getId()]) ids[linked[f2.getId()]] = 1; }
  }
  return Object.keys(ids);
}

/** Update title / notebook / tags / starred. Keeps the Doc name, its first line and Drive star in sync. */
function api_updateNote(id, ch) {
  var row = findNoteRow_(id), sh = ss_().getSheetByName('Notes');
  var vals = sh.getRange(row, 1, 1, NOTE_COLS.length).getValues()[0];
  var n = {}; NOTE_COLS.forEach(function (c, i) { n[c] = vals[i]; });
  var file = DriveApp.getFileById(n.docId), linked = n.source === 'link';   // linked files are never renamed, starred or edited
  if (ch.title !== undefined) { n.title = String(ch.title).trim() || 'Untitled Note'; if (!linked) file.setName(n.title); }
  if (ch.notebook !== undefined) n.notebook = ch.notebook;
  if (ch.tags !== undefined) {
    var known = {}; readTable_('Tags').forEach(function (t) { known[t.name] = 1; });
    ch.tags.forEach(function (t) { if (!known[t]) ss_().getSheetByName('Tags').appendRow([t, '']); });
    n.tags = ch.tags.join(TAG_SEP);
  }
  if (ch.starred !== undefined) { n.starred = !!ch.starred; if (!linked) file.setStarred(n.starred); }
  if (!linked && (ch.notebook !== undefined || ch.tags !== undefined)) {
    try {
      var doc = DocumentApp.openById(n.docId), first = doc.getBody().getChild(0);
      var line = headerLine_(n.notebook, n.tags ? String(n.tags).split(TAG_SEP) : []);
      if (first.getType() === DocumentApp.ElementType.PARAGRAPH && /^Notebook:/.test(first.asParagraph().getText())) first.asParagraph().setText(line);
      else doc.getBody().insertParagraph(0, line);
      doc.saveAndClose();
    } catch (e) {}
  }
  n.updated = new Date();
  sh.getRange(row, 1, 1, NOTE_COLS.length).setValues([noteToRow_(n)]);
  return { ok: true, updated: n.updated.toISOString() };
}

function api_setBookmarkOrder(ids) { props_().setProperty('bookmarkOrder', ids.join(',')); return true; }

/** Quick capture: new note in the Inbox. */
function api_createNote(title, text, tags) {
  tags = tags || [];
  title = String(title || '').trim() || 'Untitled Note';
  var doc = DocumentApp.create(title), body = doc.getBody();
  body.getChild(0).asParagraph().setText(headerLine_('Inbox', tags));
  String(text || '').split('\n').forEach(function (l) { body.appendParagraph(l); });
  doc.saveAndClose();
  DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(props_().getProperty('notesId')));
  var now = new Date(), id = Utilities.getUuid();
  var n = { id: id, title: title, notebook: 'Inbox', tags: tags.join(TAG_SEP), created: now, updated: now, docId: doc.getId(), starred: false, source: 'arenote', sourceUrl: '', author: '', attachments: 0 };
  ss_().getSheetByName('Notes').appendRow(noteToRow_(n));
  return { id: id, docId: doc.getId() };
}

function api_trashNote(id) {
  var row = findNoteRow_(id), sh = ss_().getSheetByName('Notes');
  var docId = sh.getRange(row, NOTE_COLS.indexOf('docId') + 1).getValue();
  try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
  sh.deleteRow(row);
  return true;
}

function api_createTag(name, parent) {
  name = String(name).trim(); if (!name) throw new Error('Empty tag name');
  var exists = readTable_('Tags').some(function (t) { return String(t.name) === name; });
  if (!exists) ss_().getSheetByName('Tags').appendRow([name, parent || '']);
  return true;
}
