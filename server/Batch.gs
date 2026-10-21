/* ---------------- batch actions & tag management ----------------
 * Sheet changes happen immediately; each Doc's first line ("Notebook: … | Tags: …")
 * is refreshed through a small queue so big batches stay fast.
 */

function notesTable_() {
  var sh = ss_().getSheetByName('Notes');
  var vals = sh.getDataRange().getValues();
  vals.shift();
  return { sh: sh, rows: vals };
}

function col_(name) { return NOTE_COLS.indexOf(name); }

function splitTags_(s) { return s ? String(s).split(TAG_SEP).filter(String) : []; }

function ensureTags_(names) {
  var known = {}; readTable_('Tags').forEach(function (t) { known[t.name] = 1; });
  var add = names.filter(function (t) { return t && !known[t]; }).map(function (t) { return [t, '']; });
  if (add.length) { var sh = ss_().getSheetByName('Tags'); sh.getRange(sh.getLastRow() + 1, 1, add.length, 2).setValues(add); }
}

/**
 * ch: { addTags:[], removeTags:[], notebook:'Inbox'|'Cabinet', starred:true|false }
 * Returns the updated notes (id, tags, notebook, starred, updated).
 */
function api_batchUpdate(ids, ch) {
  var want = {}; ids.forEach(function (id) { want[id] = 1; });
  if (ch.addTags && ch.addTags.length) ensureTags_(ch.addTags);
  var t = notesTable_(), now = new Date(), out = [], queue = [];
  t.rows.forEach(function (r, i) {
    var id = r[col_('id')]; if (!want[id]) return;
    var tags = splitTags_(r[col_('tags')]), before = r[col_('tags')] + '|' + r[col_('notebook')];
    (ch.removeTags || []).forEach(function (x) { tags = tags.filter(function (g) { return g !== x; }); });
    (ch.addTags || []).forEach(function (x) { if (tags.indexOf(x) < 0) tags.push(x); });
    r[col_('tags')] = tags.join(TAG_SEP);
    if (ch.notebook && !(ch.notebook === 'Cabinet' && !tags.length)) r[col_('notebook')] = ch.notebook;   // Cabinet needs at least one tag
    if (ch.starred !== undefined && ch.starred !== null) {
      r[col_('starred')] = !!ch.starred;
      if (r[col_('source')] !== 'link') { try { DriveApp.getFileById(r[col_('docId')]).setStarred(!!ch.starred); } catch (e) {} }
    }
    r[col_('updated')] = now;
    t.sh.getRange(i + 2, 1, 1, NOTE_COLS.length).setValues([r]);
    if (r[col_('tags')] + '|' + r[col_('notebook')] !== before && r[col_('source')] !== 'link') queue.push(r[col_('docId')]);
    out.push({ id: id, tags: tags, notebook: r[col_('notebook')], starred: r[col_('starred')] === true, updated: now.toISOString() });
  });
  queueHeaders_(queue);
  flushHeaders_(15000);
  return out;
}

/** Move notes to Drive trash (recoverable for 30 days) and drop them from the index. */
function api_batchTrash(ids) {
  var want = {}; ids.forEach(function (id) { want[id] = 1; });
  var t = notesTable_(), del = [];
  t.rows.forEach(function (r, i) { if (want[r[col_('id')]]) del.push({ row: i + 2, docId: r[col_('docId')], link: r[col_('source')] === 'link' }); });
  // attachments used only by these notes go to trash too
  var attSh = ss_().getSheetByName('Attachments'), att = attSh.getDataRange().getValues(); att.shift();
  var refs = {}; att.forEach(function (a) { if (a[1]) refs[a[1]] = (refs[a[1]] || 0) + (want[a[0]] ? 0 : 1); });
  att.forEach(function (a) { if (want[a[0]] && a[1] && !refs[a[1]] && a[5] !== 'link') { try { DriveApp.getFileById(a[1]).setTrashed(true); } catch (e) {} refs[a[1]] = 1; } });
  for (var k = att.length - 1; k >= 0; k--) if (want[att[k][0]]) attSh.deleteRow(k + 2);
  del.sort(function (a, b) { return b.row - a.row; }).forEach(function (d) {
    if (!d.link) { try { DriveApp.getFileById(d.docId).setTrashed(true); } catch (e) {} }   // linked file stays put
    t.sh.deleteRow(d.row);
  });
  return del.length;
}

/* ---- tag management ---- */

function api_renameTag(oldName, newName) {
  newName = String(newName || '').trim(); if (!newName || newName === oldName) return false;
  var tagSh = ss_().getSheetByName('Tags'), tags = tagSh.getDataRange().getValues();
  var exists = tags.some(function (r) { return String(r[0]) === newName; });   // existing name = merge
  for (var i = tags.length - 1; i >= 1; i--) {
    if (String(tags[i][1]) === oldName) tagSh.getRange(i + 1, 2).setValue(newName);       // children follow
    if (String(tags[i][0]) === oldName) { if (exists) tagSh.deleteRow(i + 1); else tagSh.getRange(i + 1, 1).setValue(newName); }
  }
  retagNotes_(function (tg) { var out = []; tg.forEach(function (g) { g = g === oldName ? newName : g; if (out.indexOf(g) < 0) out.push(g); }); return out; });
  return true;
}

function api_moveTag(name, parent) {
  if (name === parent) return false;
  var tagSh = ss_().getSheetByName('Tags'), tags = tagSh.getDataRange().getValues();
  // refuse to move a tag under its own descendant
  var p = parent, guard = 0, parentOf = {}; tags.forEach(function (r) { parentOf[String(r[0])] = String(r[1]); });
  while (p && guard++ < 20) { if (p === name) throw new Error('Cannot move a tag inside itself'); p = parentOf[p]; }
  for (var i = 1; i < tags.length; i++) if (String(tags[i][0]) === name) { tagSh.getRange(i + 1, 2).setValue(parent || ''); return true; }
  tagSh.appendRow([name, parent || '']);
  return true;
}

/** Delete a tag: removed from all notes; its sub-tags move up to its parent. */
function api_deleteTag(name) {
  var tagSh = ss_().getSheetByName('Tags'), tags = tagSh.getDataRange().getValues(), parent = '';
  tags.forEach(function (r) { if (String(r[0]) === name) parent = String(r[1]); });
  for (var i = tags.length - 1; i >= 1; i--) {
    if (String(tags[i][1]) === name) tagSh.getRange(i + 1, 2).setValue(parent);
    if (String(tags[i][0]) === name) tagSh.deleteRow(i + 1);
  }
  retagNotes_(function (tg) { return tg.filter(function (g) { return g !== name; }); });
  return true;
}

function retagNotes_(fn) {
  var t = notesTable_(), queue = [];
  t.rows.forEach(function (r, i) {
    var before = splitTags_(r[col_('tags')]), after = fn(before.slice());
    if (after.join(TAG_SEP) !== before.join(TAG_SEP)) {
      t.sh.getRange(i + 2, col_('tags') + 1).setValue(after.join(TAG_SEP));
      queue.push(r[col_('docId')]);
    }
  });
  queueHeaders_(queue);
  flushHeaders_(10000);
}

/* ---- Doc first-line refresh queue ---- */

function queueHeaders_(docIds) {
  if (!docIds.length) return;
  var p = props_(), q = (p.getProperty('headerQueue') || '').split(',').filter(String);
  docIds.forEach(function (d) { if (q.indexOf(d) < 0) q.push(d); });
  p.setProperty('headerQueue', q.join(','));
}

function flushHeaders_(budgetMs) {
  var lock = LockService.getUserLock();
  if (!lock.tryLock(100)) return 0;
  try {
    var p = props_(), q = (p.getProperty('headerQueue') || '').split(',').filter(String);
    if (!q.length) return 0;
    var byDoc = {}; notesTable_().rows.forEach(function (r) { byDoc[r[col_('docId')]] = r; });
    var start = Date.now(), done = 0;
    while (q.length && Date.now() - start < budgetMs) {
      var d = q.shift(), r = byDoc[d];
      if (r && r[col_('source')] !== 'link') {   // never write into a linked file
        try {
          var doc = DocumentApp.openById(d), body = doc.getBody(), first = body.getChild(0);
          var line = headerLine_(r[col_('notebook')], splitTags_(r[col_('tags')]));
          if (first.getType() === DocumentApp.ElementType.PARAGRAPH && /^Notebook:/.test(first.asParagraph().getText())) first.asParagraph().setText(line);
          else body.insertParagraph(0, line);
          doc.saveAndClose();
        } catch (e) {}
      }
      done++;
    }
    p.setProperty('headerQueue', q.join(','));
    return q.length;
  } finally { lock.releaseLock(); }
}



/** Called by the app in the background; returns how many Doc header lines are still waiting. */
function api_flushHeaders() { return flushHeaders_(20000); }

/* ---------------- Drop folder & file upload ----------------
 * Anything put in ArNote/Drop becomes an Inbox note:
 *   Google Doc            → becomes the note itself
 *   Word / text / HTML    → converted to a Doc note (original kept as an attachment)
 *   anything else         → a note named after the file, with the file attached (images shown inline)
 *   a sub-folder          → ONE note named after the folder, with every file in it attached
 * processDrop runs every 15 minutes (installDropTrigger), and right away after an in-app upload.
 */

var CONVERTIBLE_ = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 1, 'application/msword': 1,
  'application/rtf': 1, 'text/rtf': 1, 'application/vnd.oasis.opendocument.text': 1, 'text/plain': 1, 'text/html': 1, 'text/markdown': 1
};

function dropFolder_() {
  var p = props_(), root = DriveApp.getFolderById(p.getProperty('rootId'));
  var f = getOrCreateFolder_(p.getProperty('dropId'), 'Drop', root);
  p.setProperty('dropId', f.getId());
  return f;
}

/** Run once: creates ArNote/Drop and the 15-minute check. */
function installDropTrigger() {
  var f = dropFolder_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'processDrop') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('processDrop').timeBased().everyMinutes(15).create();
  Logger.log('Drop folder ready: ' + f.getUrl());
  return f.getUrl();
}

function processDrop() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return [];
  var made = [], start = Date.now();
  try {
    var drop = dropFolder_();
    var fi = drop.getFiles();
    while (fi.hasNext() && Date.now() - start < 240000) {
      var f = fi.next();
      try { made.push(noteFromFile_(f)); } catch (e) { Logger.log('Drop failed for ' + f.getName() + ': ' + e); }
    }
    var di = drop.getFolders();
    while (di.hasNext() && Date.now() - start < 240000) {
      var d = di.next();
      try { var r = noteFromFolder_(d); if (r) made.push(r); } catch (e) { Logger.log('Drop failed for folder ' + d.getName() + ': ' + e); }
    }
    try { made = made.concat(processLinks_(start)); } catch (e) { Logger.log('Links: ' + e); }
  } finally { lock.releaseLock(); }
  try { flushHeaders_(5000); } catch (e) {}
  return made;
}

function baseName_(name) { return String(name).replace(/\.[A-Za-z0-9]{1,5}$/, '') || 'Untitled Note'; }

function addNoteRow_(title, docId, nAtt) {
  var now = new Date(), id = Utilities.getUuid();
  var n = { id: id, title: title, notebook: 'Inbox', tags: '', created: now, updated: now, docId: docId, starred: false, source: 'drop', sourceUrl: '', author: '', attachments: nAtt || 0 };
  ss_().getSheetByName('Notes').appendRow(noteToRow_(n));
  return { id: id, title: title, notebook: 'Inbox', tags: [], created: now.toISOString(), updated: now.toISOString(), docId: docId, starred: false, att: nAtt || 0, sourceUrl: '' };
}

function attachFile_(noteId, f) {
  var p = props_();
  f.moveTo(DriveApp.getFolderById(p.getProperty('attId')));
  ss_().getSheetByName('Attachments').appendRow([noteId, f.getId(), f.getName(), f.getMimeType(), f.getSize(), '']);
}

/** Adds a file to an open Doc body: inline picture for jpeg/png/gif, otherwise a 📎 link. */
function appendFileToBody_(body, f) {
  var mime = f.getMimeType();
  if (/^image\/(jpeg|png|gif)$/.test(mime) && f.getSize() <= MAX_INLINE_BYTES) {
    try {
      var img = body.appendParagraph('').appendInlineImage(f.getBlob());
      var w = img.getWidth(); if (w > 600) img.setWidth(600).setHeight(Math.round(img.getHeight() * 600 / w));
      return;
    } catch (e) {}
  }
  body.appendParagraph('📎 ' + f.getName()).setLinkUrl(f.getUrl());
}

function newNoteDoc_(title) {
  var doc = DocumentApp.create(title);
  doc.getBody().getChild(0).asParagraph().setText(headerLine_('Inbox', []));
  DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(props_().getProperty('notesId')));
  return doc;
}

function convertToDoc_(f, title) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() + '/copy?fields=id', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: title, mimeType: 'application/vnd.google-apps.document', parents: [props_().getProperty('notesId')] }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Convert failed (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 200));
  return JSON.parse(res.getContentText()).id;
}

function noteFromFile_(f) {
  var mime = f.getMimeType(), title = baseName_(f.getName());
  if (mime === MimeType.GOOGLE_DOCS) {
    f.moveTo(DriveApp.getFolderById(props_().getProperty('notesId')));
    var d = DocumentApp.openById(f.getId()); d.getBody().insertParagraph(0, headerLine_('Inbox', [])); d.saveAndClose();
    return addNoteRow_(f.getName(), f.getId(), 0);
  }
  if (CONVERTIBLE_[mime]) {
    var docId = convertToDoc_(f, title);
    var d2 = DocumentApp.openById(docId); d2.getBody().insertParagraph(0, headerLine_('Inbox', [])); d2.saveAndClose();
    var n = addNoteRow_(title, docId, 1);
    attachFile_(n.id, f);
    return n;
  }
  var doc = newNoteDoc_(title);
  appendFileToBody_(doc.getBody(), f);
  doc.saveAndClose();
  var n2 = addNoteRow_(title, doc.getId(), 1);
  attachFile_(n2.id, f);
  return n2;
}

function noteFromFolder_(d) {
  var files = [], it = d.getFiles(); while (it.hasNext()) files.push(it.next());
  if (!files.length) { if (!d.getFolders().hasNext()) d.setTrashed(true); return null; }
  var doc = newNoteDoc_(d.getName()), body = doc.getBody();
  files.forEach(function (f) { appendFileToBody_(body, f); });
  doc.saveAndClose();
  var n = addNoteRow_(d.getName(), doc.getId(), files.length);
  files.forEach(function (f) { attachFile_(n.id, f); });
  if (!d.getFiles().hasNext() && !d.getFolders().hasNext()) d.setTrashed(true);
  return n;
}

/** In-app upload: one file per call (base64). Lands in Drop, then becomes a note right away. */
function api_uploadFile(name, mime, b64) {
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name);
  var f = dropFolder_().createFile(blob);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { return noteFromFile_(f); } finally { lock.releaseLock(); }
}

function api_dropFolderUrl() { return dropFolder_().getUrl(); }



/* ---- Home screen extras: scratch pad + pinned note ---- */
function api_homeExtras() { var p = props_(); return { scratch: p.getProperty('scratch') || '', pinned: p.getProperty('pinnedNote') || '' }; }
function api_setScratch(text) { props_().setProperty('scratch', String(text || '').slice(0, 8500)); return true; }
function api_setPinned(id) { if (id) props_().setProperty('pinnedNote', String(id)); else props_().deleteProperty('pinnedNote'); return true; }

/* ---- Merge notes ----
 * ids[0] is the note that survives (keeps its Doc, created date and notebook).
 * The others' bodies are appended to it under a heading with their title, their tags
 * and attachments move over, and their Docs go to Drive trash (restorable 30 days).
 */
function api_mergeNotes(ids, title) {
  if (!ids || ids.length < 2) throw new Error('Pick at least two notes to merge');
  var t = notesTable_(), byId = {};
  t.rows.forEach(function (r, i) { byId[r[col_('id')]] = { r: r, row: i + 2 }; });
  ids = ids.filter(function (id) { return byId[id]; });
  if (ids.length < 2) throw new Error('Notes not found');
  if (ids.some(function (id) { return byId[id].r[col_('source')] === 'link'; })) throw new Error('Linked notes can\'t be merged (that would change the original files).');
  var keep = byId[ids[0]], kr = keep.r;
  var doc = DocumentApp.openById(kr[col_('docId')]), body = doc.getBody();
  var tags = splitTags_(kr[col_('tags')]), starred = kr[col_('starred')] === true, att = +kr[col_('attachments')] || 0;
  var T = DocumentApp.ElementType;
  ids.slice(1).forEach(function (id) {
    var o = byId[id].r;
    splitTags_(o[col_('tags')]).forEach(function (g) { if (tags.indexOf(g) < 0) tags.push(g); });
    if (o[col_('starred')] === true) starred = true;
    att += +o[col_('attachments')] || 0;
    body.appendHorizontalRule();
    body.appendParagraph(String(o[col_('title')])).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    try {
      var src = DocumentApp.openById(o[col_('docId')]).getBody();
      for (var k = 0; k < src.getNumChildren(); k++) {
        var el = src.getChild(k), ty = el.getType();
        if (k === 0 && ty === T.PARAGRAPH && /^Notebook:/.test(el.asParagraph().getText())) continue;   // skip the header line
        try {
          if (ty === T.PARAGRAPH) body.appendParagraph(el.copy().asParagraph());
          else if (ty === T.LIST_ITEM) body.appendListItem(el.copy().asListItem());
          else if (ty === T.TABLE) body.appendTable(el.copy().asTable());
          else if (ty === T.PAGE_BREAK) body.appendPageBreak();
        } catch (e) {}
      }
    } catch (e) { body.appendParagraph('(could not copy this note\'s text)'); }
  });
  if (title && String(title).trim()) {
    kr[col_('title')] = String(title).trim();
    try { DriveApp.getFileById(kr[col_('docId')]).setName(kr[col_('title')]); } catch (e) {}
  }
  var first = body.getChild(0);
  var line = headerLine_(kr[col_('notebook')], tags);
  if (first.getType() === T.PARAGRAPH && /^Notebook:/.test(first.asParagraph().getText())) first.asParagraph().setText(line);
  doc.saveAndClose();
  // attachments follow the merged note
  var attSh = ss_().getSheetByName('Attachments'), av = attSh.getDataRange().getValues();
  var others = {}; ids.slice(1).forEach(function (id) { others[id] = 1; });
  for (var a = 1; a < av.length; a++) if (others[av[a][0]]) attSh.getRange(a + 1, 1).setValue(ids[0]);
  var now = new Date();
  kr[col_('tags')] = tags.join(TAG_SEP); kr[col_('starred')] = starred; kr[col_('attachments')] = att; kr[col_('updated')] = now;
  t.sh.getRange(keep.row, 1, 1, NOTE_COLS.length).setValues([kr]);
  if (starred) { try { DriveApp.getFileById(kr[col_('docId')]).setStarred(true); } catch (e) {} }
  // remove the merged-away notes (bottom-up so row numbers stay valid)
  ids.slice(1).map(function (id) { return byId[id]; }).sort(function (x, y) { return y.row - x.row; }).forEach(function (o) {
    try { DriveApp.getFileById(o.r[col_('docId')]).setTrashed(true); } catch (e) {}
    t.sh.deleteRow(o.row);
  });
  return { id: ids[0], title: String(kr[col_('title')]), tags: tags, starred: starred, att: att, updated: now.toISOString() };
}

/* ---------------- JSON API for the installable ArNote app ----------------
 * The app (served from GitHub Pages) POSTs {k: key, fn: 'api_x', a: [args]} as plain text.
 * Every call must carry this server's private key (Script Property 'apiKey').
 * Run apiSetup() once in the editor: it creates the key and logs the pairing link.
 */
var APP_URL = 'https://mtnst8.github.io/ArNote/';

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var key = props_().getProperty('apiKey');
    if (!key || typeof req.k !== 'string' || req.k !== key) throw new Error('Not authorized: this device needs to be paired again');
    var fn = String(req.fn || '');
    if (!/^api_[A-Za-z0-9]+$/.test(fn) || typeof globalThis[fn] !== 'function') throw new Error('Unknown action: ' + fn);
    out = { ok: true, result: globalThis[fn].apply(null, Array.isArray(req.a) ? req.a : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor. Creates the private key (if missing) and logs the pairing link. */
function apiSetup() {
  var p = props_(), key = p.getProperty('apiKey');
  if (!key) {
    key = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now())).replace(/=+$/, '');
    p.setProperty('apiKey', key);
  }
  Logger.log('Private key: ' + key);
  Logger.log('Next: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), copy the Web app URL, then open ' + APP_URL +
    ' and choose "Or enter server address and key".');
  return 'ok';
}

/** Change the key (all paired devices must pair again). */
function apiRotateKey() { props_().deleteProperty('apiKey'); return apiSetup(); }

function pairingLink_(serverUrl, key) {
  return APP_URL + '#pair=' + Utilities.base64EncodeWebSafe(serverUrl + '|' + key).replace(/=+$/, '');
}

/** Basic facts the app needs (no hard-coded IDs in the app, so other people can run their own server). */
function api_appInfo() {
  var drop = null; try { drop = dropFolder_().getUrl(); } catch (e) {}
  var link = null; try { link = linkFolder_().getUrl(); } catch (e) {}
  return { user: Session.getEffectiveUser().getEmail(), dropUrl: drop, linkUrl: link, version: 2 };
}

/** Note body as clean HTML for the app's reading view (header line removed), plus attachments. */
function api_noteView(id) {
  var n = readTable_('Notes').filter(function (x) { return x.id === id; })[0];
  if (!n) throw new Error('Note not found');
  if (n.source === 'link') {
    var meta = driveMeta_(n.docId);
    if (!meta) return { html: '<p style="font:14px sans-serif">This linked file is missing or no longer shared with you.</p>', atts: [], docId: n.docId, link: { url: n.sourceUrl } };
    if (meta.mimeType !== 'application/vnd.google-apps.document')
      return { html: linkedPreviewHtml_(n.docId, meta), atts: [], docId: n.docId, link: { url: meta.webViewLink || n.sourceUrl, mime: meta.mimeType, name: meta.name } };
  }
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + n.docId + '/export?mimeType=text/html',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  var html = res.getResponseCode() === 200 ? res.getContentText('UTF-8') : '<p>(Could not load this note: ' + res.getResponseCode() + ')</p>';
  html = html.replace(/<p\b[^>]*>(?:(?!<\/p>)[\s\S])*?Notebook:(?:(?!<\/p>)[\s\S])*?<\/p>/, '');
  var atts = readTable_('Attachments').filter(function (a) { return a.noteId === id; })
    .map(function (a) { return { fileId: a.fileId, filename: String(a.filename), mime: a.mime, bytes: +a.bytes }; });
  return { html: html, atts: atts, docId: n.docId };
}

/** Small page thumbnails for note cards, as data URIs. */
function api_thumbs(ids) {
  var byId = {}; readTable_('Notes').forEach(function (n) { byId[n.id] = n.docId; });
  var out = {}, start = Date.now();
  (ids || []).slice(0, 24).forEach(function (id) {
    if (!byId[id] || Date.now() - start > 20000) return;
    try {
      var b = DriveApp.getFileById(byId[id]).getThumbnail();
      if (b) out[id] = 'data:' + (b.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(b.getBytes());
    } catch (e) {}
  });
  return out;
}


/* ---------------- Linked notes ----------------
 * A linked note points at a file that stays where it is in Drive (Notes row: source = 'link', docId = that file).
 * ArNote never renames, edits, stars or trashes a linked file. Removing the note only removes the link.
 * Ways in: a shortcut put in ArNote/Link (checked with the Drop folder every 15 minutes), pasted links
 * (api_linkFiles), or a linked folder: every file in it (and its sub-folders) gets a note in the Cabinet,
 * tagged with its folder's name, and new files are picked up on each check.
 */
function linkFolder_() {
  var p = props_(), root = DriveApp.getFolderById(p.getProperty('rootId'));
  var f = getOrCreateFolder_(p.getProperty('linkId'), 'Link', root);
  p.setProperty('linkId', f.getId());
  return f;
}

function linkedIndex_() {
  var m = {}; readTable_('Notes').forEach(function (n) { if (n.source === 'link') m[n.docId] = n.id; });
  return m;
}

function linkTitle_(f) {
  var name = f.getName();
  return /^application\/vnd\.google-apps\./.test(f.getMimeType()) ? name : (name.replace(/\.[A-Za-z0-9]{1,5}$/, '') || name);
}

function createLinkedNote_(f, notebook, tags, index) {
  if (index[f.getId()]) return null;
  var now = new Date(), id = Utilities.getUuid();
  var n = { id: id, title: linkTitle_(f), notebook: notebook || 'Inbox', tags: (tags || []).join(TAG_SEP), created: now, updated: now,
    docId: f.getId(), starred: false, source: 'link', sourceUrl: f.getUrl(), author: '', attachments: 0 };
  ss_().getSheetByName('Notes').appendRow(noteToRow_(n));
  index[f.getId()] = id;
  return { id: id, title: n.title, notebook: n.notebook, tags: tags || [], created: now.toISOString(), updated: now.toISOString(),
    docId: n.docId, starred: false, att: 0, sourceUrl: n.sourceUrl, link: true };
}

function ensureTagParent_(name, parent) {
  var exists = readTable_('Tags').some(function (t) { return String(t.name) === name; });
  if (!exists) ss_().getSheetByName('Tags').appendRow([name, parent || '']);
}

function watchList_() { try { return JSON.parse(props_().getProperty('watchFolders') || '[]'); } catch (e) { return []; } }

/** Link every file in a folder (and sub-folders) into the Cabinet, tagged with the folder name. */
function syncFolder_(folder, index, start, parentTag) {
  var tag = folder.getName(), out = [];
  ensureTagParent_(tag, parentTag || '');
  var it = folder.getFiles();
  while (it.hasNext() && Date.now() - start < 240000) {
    var f = it.next();
    if (f.getMimeType() === MimeType.SHORTCUT) continue;
    var n = createLinkedNote_(f, 'Cabinet', [tag], index); if (n) out.push(n);
  }
  var sub = folder.getFolders();
  while (sub.hasNext() && Date.now() - start < 240000) out = out.concat(syncFolder_(sub.next(), index, start, tag));
  return out;
}

function watchFolder_(folder, index, start) {
  var w = watchList_();
  if (w.indexOf(folder.getId()) < 0) { w.push(folder.getId()); props_().setProperty('watchFolders', JSON.stringify(w)); }
  return syncFolder_(folder, index, start, '');
}

/** Called from processDrop (every 15 minutes): shortcuts in ArNote/Link, plus new files in linked folders. */
function processLinks_(start) {
  var index = linkedIndex_(), made = [];
  var it = linkFolder_().getFiles();
  while (it.hasNext() && Date.now() - start < 240000) {
    var s = it.next();
    try {
      if (s.getMimeType() === MimeType.SHORTCUT) {
        var tid = s.getTargetId(), target = null, isFolder = false;
        try { target = DriveApp.getFileById(tid); } catch (e) { try { target = DriveApp.getFolderById(tid); isFolder = true; } catch (e2) {} }
        if (target) {
          if (isFolder) made = made.concat(watchFolder_(target, index, start));
          else { var n = createLinkedNote_(target, 'Inbox', [], index); if (n) made.push(n); }
        }
        s.setTrashed(true);   // only the shortcut; the original is untouched
      } else {
        made.push(noteFromFile_(s));   // a real file put here is imported like the Drop folder
      }
    } catch (e) { Logger.log('Link failed for ' + s.getName() + ': ' + e); }
  }
  watchList_().forEach(function (fid) {
    if (Date.now() - start > 240000) return;
    try { made = made.concat(syncFolder_(DriveApp.getFolderById(fid), index, start, '')); } catch (e) {}
  });
  return made;
}

/** Paste-a-link: Drive file or folder links (one or many). */
function api_linkFiles(text) {
  var ids = String(text || '').match(/[-\w]{25,}/g) || [], seen = {}, index = linkedIndex_(), start = Date.now();
  var out = { notes: [], folders: [], failed: 0, already: 0 };
  ids.forEach(function (id) {
    if (seen[id]) return; seen[id] = 1;
    var f = null, folder = null;
    try { f = DriveApp.getFileById(id); } catch (e) { try { folder = DriveApp.getFolderById(id); } catch (e2) {} }
    if (f) { var n = createLinkedNote_(f, 'Inbox', [], index); if (n) out.notes.push(n); else out.already++; }
    else if (folder) { var made = watchFolder_(folder, index, start); out.folders.push({ name: folder.getName(), count: made.length }); out.notes = out.notes.concat(made); }
    else out.failed++;
  });
  return out;
}

function api_linkedFolders() {
  return watchList_().map(function (id) { try { var f = DriveApp.getFolderById(id); return { id: id, name: f.getName(), url: f.getUrl() }; } catch (e) { return { id: id, name: '(missing folder)', url: '' }; } });
}

function api_unwatchFolder(id) {
  props_().setProperty('watchFolders', JSON.stringify(watchList_().filter(function (x) { return x !== id; })));
  return true;
}

function driveMeta_(id) {
  var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?fields=mimeType,name,thumbnailLink,size,webViewLink',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  return r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null;
}

/** Reading-view HTML for a linked file that isn't a Google Doc: the picture itself, or a large page-1 preview. */
function linkedPreviewHtml_(id, meta) {
  var src = '';
  try {
    if (/^image\//.test(meta.mimeType) && +meta.size < 5 * 1024 * 1024) {
      var b = DriveApp.getFileById(id).getBlob();
      src = 'data:' + meta.mimeType + ';base64,' + Utilities.base64Encode(b.getBytes());
    } else if (meta.thumbnailLink) {
      var r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s1600'), { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) src = 'data:' + (r.getHeaders()['Content-Type'] || 'image/png') + ';base64,' + Utilities.base64Encode(r.getContent());
    }
  } catch (e) {}
  var name = String(meta.name || '').replace(/[<>&"]/g, '');
  return '<html><body style="font:14px sans-serif;color:#444;text-align:center">' +
    (src ? '<img src="' + src + '" style="max-width:100%;box-shadow:0 1px 6px rgba(0,0,0,.25)">' +
      (/^image\//.test(meta.mimeType) ? '' : '<p style="color:#888">Preview of the first page · open the file to see all of it</p>')
      : '<p style="margin-top:40px">No preview for this kind of file.<br>Use “Open in Drive” to see it.</p>') +
    '<p style="color:#888;font-size:12px">' + name + '</p></body></html>';
}

/* ---------------- Print ----------------
 * Google Docs (notes and linked Docs) are combined into one PDF; other files come back as their own PDF/file.
 */
function exportPdf_(id) {
  var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=application/pdf',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('PDF export failed (' + r.getResponseCode() + ')');
  return r.getContent();
}

function api_printPdf(ids) {
  var byId = {}; readTable_('Notes').forEach(function (n) { byId[n.id] = n; });
  var docs = [], others = [], out = { files: [] }, T = DocumentApp.ElementType;
  (ids || []).forEach(function (id) {
    var n = byId[id]; if (!n) return;
    var mt = ''; try { mt = DriveApp.getFileById(n.docId).getMimeType(); } catch (e) { return; }
    (mt === MimeType.GOOGLE_DOCS ? docs : others).push({ n: n, mt: mt });
  });
  if (docs.length) {
    var tmp = DocumentApp.create('ArNote print'), body = tmp.getBody();
    docs.forEach(function (d, i) {
      if (i) body.appendPageBreak();
      body.appendParagraph(String(d.n.title)).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      var src = DocumentApp.openById(d.n.docId).getBody();
      for (var k = 0; k < src.getNumChildren(); k++) {
        var el = src.getChild(k), ty = el.getType();
        if (k === 0 && ty === T.PARAGRAPH && /^Notebook:/.test(el.asParagraph().getText())) continue;
        try {
          if (ty === T.PARAGRAPH) body.appendParagraph(el.copy().asParagraph());
          else if (ty === T.LIST_ITEM) body.appendListItem(el.copy().asListItem());
          else if (ty === T.TABLE) body.appendTable(el.copy().asTable());
          else if (ty === T.PAGE_BREAK) body.appendPageBreak();
        } catch (e) {}
      }
    });
    try { var p0 = body.getChild(0); if (p0.getType() === T.PARAGRAPH && !p0.asParagraph().getText() && body.getNumChildren() > 1) body.removeChild(p0); } catch (e) {}
    tmp.saveAndClose();
    try {
      var pdf = exportPdf_(tmp.getId());
      out.files.push({ name: docs.length > 1 ? 'ArNote - ' + docs.length + ' notes.pdf' : String(docs[0].n.title) + '.pdf', mime: 'application/pdf', count: docs.length, b64: Utilities.base64Encode(pdf) });
    } finally { DriveApp.getFileById(tmp.getId()).setTrashed(true); }
  }
  others.forEach(function (o) {
    var f = DriveApp.getFileById(o.n.docId), name = String(o.n.title), b, mt = o.mt;
    try {
      if (/^application\/vnd\.google-apps\./.test(mt)) { b = exportPdf_(o.n.docId); mt = 'application/pdf'; name += '.pdf'; }
      else { if (f.getSize() > 20 * 1024 * 1024) { out.files.push({ name: f.getName(), tooBig: true, url: f.getUrl() }); return; } b = f.getBlob().getBytes(); mt = f.getMimeType(); name = f.getName(); }
      out.files.push({ name: name, mime: mt, count: 1, b64: Utilities.base64Encode(b) });
    } catch (e) { out.files.push({ name: name, error: String(e.message || e), url: f.getUrl() }); }
  });
  return out;
}
