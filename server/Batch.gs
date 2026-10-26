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

/** Tag order (Tags sheet column C): at the top level a number = pinned to the top in that order; under a parent = manual order. Blank = alphabetical after. */
function tagSheetOrder_() { var sh = ss_().getSheetByName('Tags'); if (String(sh.getRange(1, 3).getValue()) !== 'order') sh.getRange(1, 3).setValue('order').setFontWeight('bold'); return sh; }
function api_tagOrder(parent, names) {
  parent = String(parent || ''); var pos = {}; (names || []).forEach(function (n, i) { pos[String(n)] = i + 1; });
  var sh = tagSheetOrder_(), n = sh.getLastRow() - 1; if (n < 1) return true;
  var v = sh.getRange(2, 1, n, 3).getValues();
  sh.getRange(2, 3, n, 1).setValues(v.map(function (r) { return [String(r[1] || '') === parent ? (pos[String(r[0])] || '') : r[2]]; }));
  return true;
}
/** The tags ArNote makes for Google accounts (not to be mixed up with tags that happen to start with @). */
function acctTags_() { var s = {}, m = accountMap_(); Object.keys(m).forEach(function (e) { if (m[e]) s[m[e]] = 1; }); return s; }

function api_renameTag(oldName, newName) {
  newName = String(newName || '').trim(); if (!newName || newName === oldName) return false;
  var tagSh = ss_().getSheetByName('Tags'), tags = tagSh.getDataRange().getValues();
  var exists = tags.some(function (r) { return String(r[0]) === newName; });   // existing name = merge
  for (var i = tags.length - 1; i >= 1; i--) {
    if (String(tags[i][1]) === oldName) tagSh.getRange(i + 1, 2).setValue(newName);       // children follow
    if (String(tags[i][0]) === oldName) { if (exists) tagSh.deleteRow(i + 1); else tagSh.getRange(i + 1, 1).setValue(newName); }
  }
  retagNotes_(function (tg) { var out = []; tg.forEach(function (g) { g = g === oldName ? newName : g; if (out.indexOf(g) < 0) out.push(g); }); return out; });
  tagRenamed_(oldName, newName);
  return true;
}

function api_moveTag(name, parent) {
  if (name === parent) return false;
  var tagSh = ss_().getSheetByName('Tags'), tags = tagSh.getDataRange().getValues();
  // refuse to move a tag under its own descendant
  var p = parent, guard = 0, parentOf = {}; tags.forEach(function (r) { parentOf[String(r[0])] = String(r[1]); });
  while (p && guard++ < 20) { if (p === name) throw new Error('Cannot move a tag inside itself'); p = parentOf[p]; }
  for (var i = 1; i < tags.length; i++) if (String(tags[i][0]) === name) { tagSh.getRange(i + 1, 2, 1, 2).setValues([[parent || '', '']]); return true; }   // a new parent: its old place in the order no longer applies
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
  tagDeleted_(name);
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

/** One-time fix: store tag names as plain text so names like "2018" aren't turned into numbers. */
function fixTextFormats() {
  ['Tags', 'Notes'].forEach(function (nm) {
    var sh = ss_().getSheetByName(nm), cols = nm === 'Tags' ? [1, 2] : [col_('tags') + 1];
    cols.forEach(function (c) {
      var rg = sh.getRange(1, c, sh.getMaxRows(), 1), v = rg.getValues().map(function (r) { return [r[0] === '' ? '' : String(r[0])]; });
      rg.setNumberFormat('@').setValues(v);
    });
  });
  return 'ok';
}

/** One-time: move Cabinet notes that have no tags back to the Inbox. */
function moveUntaggedCabinetToInbox() {
  var t = notesTable_(), ids = [];
  t.rows.forEach(function (r) { if (String(r[col_('notebook')]) === 'Cabinet' && !splitTags_(r[col_('tags')]).length) ids.push(r[col_('id')]); });
  if (ids.length) api_batchUpdate(ids, { notebook: 'Inbox' });
  Logger.log('Moved ' + ids.length + ' untagged Cabinet notes to the Inbox');
  return ids.length;
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
  try { warmCode_(); } catch (e) {}
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
    try { syncPhotos_(Date.now() - 60000); } catch (e) { Logger.log('Photos: ' + e); }
    try { if (Date.now() - start < 200000) processMail_(Date.now(), 40000); } catch (e) { Logger.log('Mail: ' + e); }
    try { scanQueue_(90000); } catch (e) { Logger.log('Scan: ' + e); }
    try { licenceDaily_(); } catch (e) { Logger.log('Licence: ' + e); }
    try { if (Date.now() - start < 170000) backupDaily_(); } catch (e) { Logger.log('Backup: ' + e); }
    try { buildReceiptReport_(); } catch (e) { Logger.log('Receipt report: ' + e); }
    try { if (Date.now() - start < 230000) geoWork_(Date.now(), 45000); } catch (e) { Logger.log('Geo: ' + e); }
    try { if (Date.now() - start < 250000) photoDedupe_(40000); } catch (e) { Logger.log('Dedupe: ' + e); }
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

/** One-time: trash the uploaded "ArNote-import" working folder in Drive (restorable 30 days). */
function trashImportFolder() {
  var it = DriveApp.getFoldersByName('ArNote-import'), n = 0;
  var attId = props_().getProperty('attId'), rootId = props_().getProperty('rootId');
  while (it.hasNext()) {
    var f = it.next(); if (f.isTrashed() || f.getId() === attId || f.getId() === rootId) continue;
    // safety: only trash it if nothing ArNote uses lives inside (no sub-folders, no files besides the import list)
    var subs = f.getFolders(), full = false; while (subs.hasNext()) { var s = subs.next(); if (s.getFiles().hasNext() || s.getFolders().hasNext()) full = true; }
    if (full) { Logger.log('Skipped (a sub-folder still has files): ' + f.getUrl()); continue; }
    var fi = f.getFiles(), ok = true; while (fi.hasNext()) { if (!/notes\.jsonl$/i.test(fi.next().getName())) ok = false; }
    if (!ok) { Logger.log('Skipped (has other files): ' + f.getUrl()); continue; }
    f.setTrashed(true); n++;
  }
  Logger.log('Trashed ' + n + ' ArNote-import folder(s)');
  return n;
}


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
    if (typeof req.s === 'string' && req.s) {   // read-only share link (see Share section)
      out = { ok: true, result: shareDispatch_(req.s, String(req.fn || ''), Array.isArray(req.a) ? req.a : []) };
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }
    if (req.inv === 'get' || req.inv === 'accept') {   // the invite page (public; only answers a valid invite token)
      out = { ok: true, result: isLicenceServer_() ? inviteReq_(req) : null };
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }
    if (req.lic === 'code') {                      // a copy on the list fetching the released code
      out = { ok: true, result: isLicenceServer_() ? releaseFor_(req) : null };
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }
    if (req.lic === 'check') {                     // another ArNote checking in with Brian's (no key: it only says yes/no)
      out = { ok: true, result: isLicenceServer_() ? licCheck_(req) : { ok: false, reason: 'not a licence server' } };
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }
    var key = props_().getProperty('apiKey');
    if (!key || typeof req.k !== 'string' || req.k !== key) throw new Error('Not authorized: this device needs to be paired again');
    var fn = String(req.fn || '');
    if (fn !== 'api_appInfo' && !licenceState_().ok) throw new Error('ArNote access has ended for this account. Your notes and photos are still in your Google Drive.');
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
function api_appInfo(dev) {
  var drop = null; try { drop = dropFolder_().getUrl(); } catch (e) {}
  var link = null; try { link = linkFolder_().getUrl(); } catch (e) {}
  noteUsed_(); try { noteDevice_(dev); } catch (e) {}
  return { user: Session.getEffectiveUser().getEmail(), dropUrl: drop, linkUrl: link, version: 2, build: ARNOTE_BUILD, licence: licenceState_() };
}

/** Which kinds of device open this ArNote (e.g. "iPhone app", "Windows"), so Brian can see who has it on their phone.
 * Only the kind of device and the day last seen; a new kind is reported to Brian's list straight away. */
function noteDevice_(dev) {
  dev = String(dev || '').replace(/[^A-Za-z ]/g, '').slice(0, 20); if (!dev) return;
  var p = props_(), d = {}; try { d = JSON.parse(p.getProperty('devices') || '{}'); } catch (e) {}
  var isNew = !d[dev], t = today_(); if (d[dev] === t) return;
  d[dev] = t; p.setProperty('devices', JSON.stringify(d));
  if (isNew && !isLicenceServer_()) { try { licenceCheck_(); } catch (e) {} }
}
function devicesText_() {
  var d = {}; try { d = JSON.parse(props_().getProperty('devices') || '{}'); } catch (e) {}
  return Object.keys(d).sort(function (a, b) { return d[b] < d[a] ? -1 : d[b] > d[a] ? 1 : 0; }).join(', ');
}

/* ---------------- People using ArNote (licence) ----------------
 * Brian's own ArNote (this script ID) is the licence server and keeps the "Users" sheet.
 * Every other copy checks in once a day with its owner's email, an install id, its build and the last day the app was opened —
 * nothing else (no notes, tags or photos). Someone removed (or never added) gets ok:false and their app shows
 * "access ended"; their notes and photos in their own Drive are never touched. */
var ARNOTE_BUILD = 58;
var LICENCE_SCRIPT_ID = '1NmFRd6lH63u-JBeTJfg0IyMZtDJj6jxw8WNfbOdh_TSTuOpVZ5JJ2s4B';
var LICENCE_URL = 'https://script.google.com/macros/s/AKfycbyFPAGYJE1bVycoBv8QamFeUHuBjbCAhogOM8mSU7vn20CVO_Ex1XJYfRy7f7lZES-_Kg/exec';
var USER_COLS = ['email', 'name', 'status', 'plan', 'added', 'lastUsed', 'lastCheckin', 'build', 'install', 'note', 'contact', 'invite', 'invited', 'accepted', 'devices'];
var TEMPLATE_COPY_URL = 'https://docs.google.com/spreadsheets/d/1U4jJdCdPe7nZbzf2oSH7BCsAdAevRlnWYheZxxoFUTk/copy';
function isLicenceServer_() { return ScriptApp.getScriptId() === LICENCE_SCRIPT_ID; }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function noteUsed_() { var p = props_(), d = today_(); if (p.getProperty('lastUsed') !== d) p.setProperty('lastUsed', d); }
function usersSheet_() {
  if (!isLicenceServer_()) throw new Error('Only the owner can manage people');
  var sh = ensureSheet_('Users', USER_COLS);
  if (sh.getLastColumn() < USER_COLS.length) sh.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS]).setFontWeight('bold');
  return sh;
}
function userRows_() {
  var sh = usersSheet_(), n = sh.getLastRow() - 1, tz = Session.getScriptTimeZone(); if (n < 1) return { sh: sh, rows: [] };
  var d = function (x) { return x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x || ''); };
  return { sh: sh, rows: sh.getRange(2, 1, n, USER_COLS.length).getValues().map(function (r, i) {
    var o = { row: i + 2 }; USER_COLS.forEach(function (c, j) { o[c] = /^(added|lastUsed|lastCheckin|invited|accepted)$/.test(c) ? d(r[j]) : String(r[j] === null ? '' : r[j]); }); o.email = o.email.toLowerCase(); return o; }) };
}
function api_users() { return userRows_().rows.map(function (u) { delete u.row; delete u.install; u.invite = u.invite ? 1 : 0; return u; }); }
function api_userAdd(email, name, note) {
  email = String(email || '').trim().toLowerCase(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('That isn\'t an email address');
  var t = userRows_(); if (t.rows.some(function (u) { return u.email === email; })) throw new Error(email + ' is already on the list');
  t.sh.appendRow([email, String(name || '').trim().slice(0, 60), 'active', 'free', today_(), '', '', '', '', String(note || '').slice(0, 200), email, '', '', '', '']);
  return api_users();
}
/** ch: {status: 'active'|'removed', name, plan, note} */
function api_userSet(email, ch) {
  email = String(email || '').toLowerCase(); var t = userRows_(), u = t.rows.filter(function (x) { return x.email === email; })[0];
  if (!u) throw new Error('Not on the list'); ch = ch || {};
  if (ch.status !== undefined) t.sh.getRange(u.row, 3).setValue(ch.status === 'removed' ? 'removed' : 'active');
  if (ch.name !== undefined) t.sh.getRange(u.row, 2).setValue(String(ch.name).slice(0, 60));
  if (ch.plan !== undefined) t.sh.getRange(u.row, 4).setValue(String(ch.plan).slice(0, 20));
  if (ch.note !== undefined) t.sh.getRange(u.row, 10).setValue(String(ch.note).slice(0, 200));
  return api_users();
}
/** Take someone off the list entirely (typos, test entries). To stop someone using ArNote, set them to removed instead — that keeps the record. */
function api_userDelete(email) {
  email = String(email || '').toLowerCase(); var t = userRows_(), u = t.rows.filter(function (x) { return x.email === email; })[0];
  if (u) t.sh.deleteRow(u.row);
  return api_users();
}
/* ---------- Invites: Brian types an email and clicks Invite; they confirm their Google account on the invite page ---------- */
function inviteToken_() { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now())).replace(/=+$/, '').slice(0, 32); }
function api_userInvite(contact, name) {
  contact = String(contact || '').trim().toLowerCase(); if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(contact)) throw new Error('That isn\'t an email address');
  var t = userRows_(); if (t.rows.some(function (u) { return u.email === contact || u.contact === contact; })) throw new Error(contact + ' is already on the list (use Resend)');
  var tok = inviteToken_(); name = String(name || '').trim().slice(0, 60);
  t.sh.appendRow([contact, name, 'invited', 'free', today_(), '', '', '', '', '', contact, tok, today_(), '', '']);
  inviteEmail_(contact, name, tok);
  return api_users();
}
function api_userResend(email) {
  email = String(email || '').toLowerCase(); var t = userRows_(), u = t.rows.filter(function (x) { return x.email === email; })[0];
  if (!u) throw new Error('Not on the list');
  var tok = u.invite || inviteToken_();
  t.sh.getRange(u.row, 12, 1, 2).setValues([[tok, today_()]]);
  inviteEmail_(u.contact || u.email, u.name, tok);
  return api_users();
}
/* The invite email is an ordinary ArNote note ("ArNote invite email", in the Cabinet) that Brian can edit like any note.
 * First line "Subject: …"; {first name}, {invite button} and {from} are filled in when it is sent. */
var INVITE_DEFAULT_ = ['Subject: You\'re invited to ArNote', '',
  'Hi {first name},', '',
  'You\'re invited to ArNote, the notes-and-photos app I\'ve been building.', '',
  'What it does:',
  '- Notes, photos, receipts, business cards and documents in one place, organized by tags (Who / What / When / Where to start).',
  '- Everything lives in YOUR own Google Drive. I can\'t see your notes, and nothing is stored anywhere else.',
  '- Photos: one gallery you can drill into by tag or by place on a map.',
  '- Receipts: tag by job and expense category, mark billed / not billed, print a clean report.',
  '- Email in: forward any email (receipts, invoices, anything) to your ArNote address and it lands in your Inbox to tag.',
  '- Daily backup, and new features show up on their own.', '',
  'Start here (about 10 minutes, best on a computer; you\'ll need a Google account, a free Gmail is fine):',
  '{invite button}', '',
  'Then put it on your phone too. That\'s where it shines: snap a receipt or a business card, or send photos from your gallery or a text, and it\'s filed in seconds. The same link shows the phone steps.', '',
  'Any questions, just reply.'].join('\n');
/** Owner: the note holding the invite email (made from the default the first time). */
function api_inviteNote() {
  if (!isLicenceServer_()) throw new Error('Only the owner sends invites');
  var p = props_(), id = p.getProperty('inviteNoteId');
  if (id) { try { findNoteRow_(id); return { id: id }; } catch (e) {} }
  var n = api_createNote('ArNote invite email', INVITE_DEFAULT_, []);
  try { api_updateNote(n.id, { notebook: 'Cabinet' }); } catch (e) {}
  p.setProperty('inviteNoteId', n.id);
  return { id: n.id };
}
function inviteText_() {
  var id = props_().getProperty('inviteNoteId');
  if (id) {
    try {
      var row = findNoteRow_(id), docId = ss_().getSheetByName('Notes').getRange(row, NOTE_COLS.indexOf('docId') + 1).getValue();
      var t = DocumentApp.openById(docId).getBody().getText().replace(/^\s*Notebook:[^\n]*\n?/, '');
      if (t.trim()) return t;
    } catch (e) {}
  }
  return INVITE_DEFAULT_;
}
function inviteEmail_(to, name, tok) {
  var link = APP_URL + '#invite=' + tok, me = Session.getEffectiveUser().getEmail(), first = String(name || '').split(/\s+/)[0];
  var t = inviteText_().replace(/\r/g, ''), subject = 'You\'re invited to ArNote';
  var m = t.match(/^\s*Subject:\s*(.+)\n?/i); if (m) { subject = m[1].trim(); t = t.slice(m[0].length); }
  t = t.replace(/^\s*\n/, '').replace(/[ \t]*\{first name\}/gi, first ? ' ' + first : '').replace(/\{from\}/gi, me);
  if (!/\{invite button\}/i.test(t)) t += '\n\n{invite button}';
  var esc = function (x) { return String(x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var plain = t.replace(/\{invite button\}/gi, link);
  var html = esc(t).replace(/\{invite button\}/gi, '<a href="' + esc(link) + '" style="display:inline-block;background:#2f6f4f;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-weight:bold">Set up ArNote</a>').replace(/\n/g, '<br>');
  GmailApp.sendEmail(to, subject, plain, { htmlBody: '<div style="font:15px/1.5 sans-serif;max-width:560px">' + html + '</div>', replyTo: me });
}
/** Public (no key): the invite page reads and accepts an invite by its token. Says nothing unless the token matches. */
function inviteReq_(req) {
  var tok = String(req.t || ''); if (!/^[A-Za-z0-9_-]{20,}$/.test(tok)) return null;
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) {}
  try {
    var t = userRows_(), u = t.rows.filter(function (x) { return x.invite === tok; })[0];
    if (!u || u.status === 'removed') return null;
    if (req.inv === 'accept' && !u.accepted) {
      var acct = String(req.account || '').trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(acct)) throw new Error('That isn\'t an email address');
      if (t.rows.some(function (x) { return x.row !== u.row && x.email === acct; })) throw new Error(acct + ' is already in use for ArNote');
      t.sh.getRange(u.row, 1).setValue(acct); t.sh.getRange(u.row, 3).setValue('active'); t.sh.getRange(u.row, 14).setValue(today_());
      u.email = acct; u.accepted = today_();
    }
    return { name: u.name, contact: u.contact || u.email, account: u.email, accepted: !!u.accepted, copyUrl: TEMPLATE_COPY_URL, from: Session.getEffectiveUser().getEmail() };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
function licCheck_(req) {
  var email = String(req.email || '').trim().toLowerCase(); if (!email) return { ok: false, status: 'unknown' };
  var lock = LockService.getScriptLock(); try { lock.waitLock(10000); } catch (e) {}
  try {
    var t = userRows_(), u = t.rows.filter(function (x) { return x.email === email; })[0];
    if (!u) return { ok: false, status: 'unknown' };
    var used = /^\d{4}-\d\d-\d\d$/.test(String(req.lastUsed || '')) ? String(req.lastUsed) : '';
    t.sh.getRange(u.row, 6, 1, 4).setValues([[used && used > u.lastUsed ? used : u.lastUsed, today_(), String(+req.build || ''), String(req.install || '').slice(0, 60)]]);
    if (u.status === 'invited') {                 // set up straight from the invite without the form: that counts as accepting
      u.status = 'active'; t.sh.getRange(u.row, 3).setValue('active'); if (!u.accepted) t.sh.getRange(u.row, 14).setValue(today_());
    }
    if (typeof req.devices === 'string' && req.devices) t.sh.getRange(u.row, 15).setValue(req.devices.slice(0, 120));
    return { ok: u.status === 'active', status: u.status, plan: u.plan, release: +props_().getProperty('releaseN') || 0 };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
/* ---------- First-time setup of a friend's ArNote (called by the starter's "Set up ArNote") ---------- */
var STARTER_TAGS_ = [['Who', '', 1], ['What', '', 2], ['When', '', 3], ['Where', '', 4], ['Receipt', '', 5], ['Business Card', '', 6], ['Document', '', 7], ['Photos', '', 8],
  ['Family', 'Who', ''], ['Friends', 'Who', ''],
  ['Projects', 'What', ''], ['Jobs', 'What', ''], ['Example Job', 'Jobs', ''], ['Vehicles', 'What', ''], ['Equipment', 'What', ''], ['Finances', 'What', ''], ['Taxes', 'What', ''],
  ['Insurance', 'What', ''], ['Manuals & Warranties', 'What', ''], ['Recipes', 'What', ''],
  ['Now', 'When', 1], ['Next', 'When', 2], ['Soon', 'When', 3], ['Later', 'When', 4], ['Someday', 'When', 5], ['Waiting', 'When', 6], ['Completed', 'When', 7],
  ['Home', 'Where', ''], ['Work', 'Where', ''], ['Videos', 'Photos', ''],
  ['Expense Categories', '', ''], ['Materials', 'Expense Categories', ''], ['Tools & Equipment', 'Expense Categories', ''], ['Fuel', 'Expense Categories', ''], ['Vehicle', 'Expense Categories', ''],
  ['Meals', 'Expense Categories', ''], ['Office', 'Expense Categories', ''], ['Utilities', 'Expense Categories', ''], ['Subcontractors', 'Expense Categories', ''],
  ['Fees & Permits', 'Expense Categories', ''], ['Other', 'Expense Categories', '']];
function arnoteSetupAll_() {
  var p = props_(), ss = SpreadsheetApp.getActiveSpreadsheet();
  p.setProperty('sheetId', ss.getId());
  var tags = ss.getSheetByName('Tags') || ss.insertSheet('Tags');
  if (tags.getLastRow() < 2) {
    tags.getRange(1, 1, 1, 3).setValues([['name', 'parent', 'order']]).setFontWeight('bold'); tags.setFrozenRows(1);
    tags.getRange(2, 1, STARTER_TAGS_.length, 3).setValues(STARTER_TAGS_);
  }
  setup();
  apiSetup();
  installDropTrigger();
  if (!p.getProperty('receiptCfg')) p.setProperty('receiptCfg', JSON.stringify({ cols: [{ name: 'Job', parent: 'Jobs' }, { name: 'Category', parent: 'Expense Categories' }], cards: {} }));
  try { var cf = p.getProperty('codeFile'); if (cf) DriveApp.getFileById(cf).moveTo(DriveApp.getFolderById(p.getProperty('rootId'))); } catch (e) {}
  noteUsed_(); try { licenceCheck_(); } catch (e) {}
  return 'Your ArNote folder, index and starter tags are ready.\n\nLast step (once): in this spreadsheet choose Extensions → Apps Script. In the window that opens click Deploy (top right) → New deployment → the gear ⚙ → Web app. Set "Execute as: Me" and "Who has access: Anyone", click Deploy and allow the permissions, then click Copy under "Web app URL".\n\nThen come back here, choose ArNote → Open the app and paste that address. You\'ll get your pairing link for your phone and computer.';
}

/* ---------- Updates: Brian releases a version; copies on the list load it ----------
 * A friend's ArNote is a small fixed starter (Loader.gs) that runs the ArNote code kept in a file in their own Drive.
 * Brian's "Release this version" packages the server code from the GitHub repo (server/Code.gs + server/Batch.gs);
 * each copy's daily check downloads a newer release into its code file. No republishing, no Google switches. */
var RELEASE_SOURCES = ['server/Code.gs', 'server/Batch.gs'];
function releaseRaw_(path) {
  var tok = props_().getProperty('ghToken'), h = tok ? { Authorization: 'token ' + tok } : {};
  var r = UrlFetchApp.fetch('https://raw.githubusercontent.com/mtnst8/ArNote/main/' + path + '?t=' + Date.now(), { headers: h, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Could not read ' + path + ' from GitHub (' + r.getResponseCode() + ')');
  return r.getContentText();
}
/** Owner: publish the current server code (as kept in GitHub) as the next release. */
function api_releaseNow(note) {
  if (!isLicenceServer_()) throw new Error('Only the owner can release');
  var code = RELEASE_SOURCES.map(function (p) { return '/* ===== ' + p + ' ===== */\n' + releaseRaw_(p); }).join('\n\n');
  if (code.indexOf('function doPost(') < 0 || code.indexOf('function setup(') < 0) throw new Error('The code from GitHub looks incomplete; nothing was released');
  try { new Function(code); } catch (e) { throw new Error('The code from GitHub has an error (' + e.message + '); nothing was released'); }
  var p = props_(), n = (+p.getProperty('releaseN') || 0) + 1, at = new Date().toISOString();
  var rel = JSON.stringify({ n: n, at: at, note: String(note || '').slice(0, 200), code: code });
  var root = DriveApp.getFolderById(p.getProperty('rootId')), id = p.getProperty('releaseFile'), f = null;
  if (id) { try { f = DriveApp.getFileById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (f) f.setContent(rel); else { f = root.createFile('ArNote release.json', rel, 'application/json'); p.setProperty('releaseFile', f.getId()); }
  p.setProperties({ releaseN: String(n), releaseAt: at, releaseNote: String(note || '').slice(0, 200) });
  return api_releaseInfo();
}
function api_releaseInfo() {
  var p = props_();
  if (isLicenceServer_()) return { owner: true, n: +p.getProperty('releaseN') || 0, at: p.getProperty('releaseAt') || '', note: p.getProperty('releaseNote') || '' };
  var l = {}; try { l = JSON.parse(p.getProperty('licence') || '{}'); } catch (e) {}
  return { owner: false, installed: +p.getProperty('releaseInstalled') || 0, latest: +l.release || 0, at: p.getProperty('updatedAt') || '', error: p.getProperty('updateErr') || '' };
}
function releaseFor_(req) {
  var l = licCheck_(req); if (!l.ok) return null;
  var id = props_().getProperty('releaseFile'); if (!id) return null;
  return DriveApp.getFileById(id).getBlob().getDataAsString();
}
/** A copy (run by the starter): keep the starter's quick copy of the code fresh. When the sheet opens, Google
 * doesn't let the starter read Drive, so the ArNote menu only appears if this quick copy is there. Same keys as Loader.gs. */
function warmCode_(src) {
  if (typeof ARNOTE_STUBS_ === 'undefined') return;           // Brian's own ArNote runs its code directly
  var p = props_(), at = +p.getProperty('codeWarmAt') || 0, cache = CacheService.getScriptCache();
  if (!src && Date.now() - at < 4 * 3600e3 && cache.get('arnote_code_n')) return;
  try {
    if (!src) { var id = p.getProperty('codeFile'); if (!id) return; src = DriveApp.getFileById(id).getBlob().getDataAsString(); }
    var parts = {}, size = 90000, k = 0;
    for (var j = 0; j < src.length; j += size) parts['arnote_code_' + (k++)] = src.slice(j, j + size);
    cache.putAll(parts, 21600); cache.put('arnote_code_n', String(k), 21600);
    p.setProperty('codeWarmAt', String(Date.now()));
  } catch (e) {}
}
/** A copy: download the latest release into this ArNote's code file; the starter uses it from the next run. */
function selfUpdate_() {
  if (isLicenceServer_()) return null;
  var p = props_(), r = UrlFetchApp.fetch(LICENCE_URL, { method: 'post', contentType: 'text/plain', muteHttpExceptions: true,
    payload: JSON.stringify({ lic: 'code', email: Session.getEffectiveUser().getEmail(), install: p.getProperty('installId') || '', build: +p.getProperty('releaseInstalled') || 0 }) });
  var body = (JSON.parse(r.getContentText()) || {}).result; if (!body) throw new Error('No update available for this account');
  var rel = JSON.parse(body);
  if (+rel.n <= (+p.getProperty('releaseInstalled') || 0)) return null;
  if (String(rel.code || '').indexOf('function doPost(') < 0) throw new Error('The update looks incomplete; kept the current version');
  try { new Function(rel.code); } catch (e) { throw new Error('The update has an error; kept the current version'); }
  var id = p.getProperty('codeFile'), f = null;
  if (id) { try { f = DriveApp.getFileById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (f) f.setContent(rel.code);
  else { f = (p.getProperty('rootId') ? DriveApp.getFolderById(p.getProperty('rootId')) : DriveApp.getRootFolder()).createFile('ArNote code.js', rel.code, 'text/plain'); p.setProperty('codeFile', f.getId()); }
  warmCode_(rel.code);
  p.setProperties({ releaseInstalled: String(rel.n), updatedAt: new Date().toISOString(), updateErr: '' });
  return rel.n;
}
function api_updateNow() { licenceCheck_(); var n = selfUpdate_(); return { updated: n, info: api_releaseInfo() }; }

/** This copy asks Brian's ArNote whether it may run. Stored with the time; failures keep the last answer. */
function licenceCheck_() {
  if (isLicenceServer_()) return { ok: true, owner: true };
  var p = props_(), inst = p.getProperty('installId'); if (!inst) { inst = Utilities.getUuid(); p.setProperty('installId', inst); }
  var res = null;
  try {
    var r = UrlFetchApp.fetch(LICENCE_URL, { method: 'post', contentType: 'text/plain', muteHttpExceptions: true,
      payload: JSON.stringify({ lic: 'check', email: Session.getEffectiveUser().getEmail(), install: inst, build: +p.getProperty('releaseInstalled') || 0, lastUsed: p.getProperty('lastUsed') || '', devices: devicesText_() }) });
    res = (JSON.parse(r.getContentText()) || {}).result;
  } catch (e) { Logger.log('Licence check: ' + e); }
  if (res && typeof res.ok === 'boolean') { res.at = new Date().toISOString(); p.setProperty('licence', JSON.stringify(res)); return res; }
  return null;
}
function licenceDaily_() {
  if (isLicenceServer_()) return;
  var l = {}; try { l = JSON.parse(props_().getProperty('licence') || '{}'); } catch (e) {}
  if (!l.at || Date.now() - new Date(l.at).getTime() > 20 * 3600e3) l = licenceCheck_() || l;
  if (l && l.ok && +l.release > (+props_().getProperty('releaseInstalled') || 0)) { try { selfUpdate_(); } catch (e) { props_().setProperty('updateErr', String(e.message || e)); } }
}
/** ok unless Brian's list says removed/unknown. Never checked yet → check now. Can't reach Brian's for 14 days → still ok (grace). */
function licenceState_() {
  if (isLicenceServer_()) return { ok: true, owner: true };
  var l = null; try { l = JSON.parse(props_().getProperty('licence') || 'null'); } catch (e) {}
  if (!l) l = licenceCheck_();
  if (!l) return { ok: true, status: 'unchecked' };
  return { ok: !!l.ok, status: l.status, plan: l.plan, at: l.at };
}

/** Note body as clean HTML for the app's reading view (header line removed), plus attachments. */
function api_noteView(id, o) {
  var n = readTable_('Notes').filter(function (x) { return x.id === id; })[0];
  if (!n) throw new Error('Note not found');
  if (n.source === 'link') {
    var meta = driveMeta_(n.docId);
    if (!meta) return { html: '<p style="font:14px sans-serif">This linked file is missing or no longer shared with you.</p>', atts: [], docId: n.docId, link: { url: n.sourceUrl } };
    if (/^image\//.test(meta.mimeType) && !(o && o.embed))      // photos load straight from Drive in the app (fast, cached); the app falls back to api_photoView
      return { html: photoPreviewHtml_(n.docId, meta), photo: n.docId, atts: [], docId: n.docId, link: { url: meta.webViewLink || n.sourceUrl, mime: meta.mimeType, name: meta.name }, loc: linkLocation_(n.docId) };
    if (meta.mimeType !== 'application/vnd.google-apps.document')
      return { html: linkedPreviewHtml_(n.docId, meta), atts: [], docId: n.docId, link: { url: meta.webViewLink || n.sourceUrl, mime: meta.mimeType, name: meta.name }, loc: linkLocation_(n.docId) };
  }
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + n.docId + '/export?mimeType=text/html',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  var html = res.getResponseCode() === 200 ? res.getContentText('UTF-8') : '<p>(Could not load this note: ' + res.getResponseCode() + ')</p>';
  html = html.replace(/<p\b[^>]*>(?:(?!<\/p>)[\s\S])*?Notebook:(?:(?!<\/p>)[\s\S])*?<\/p>/, '');
  var atts = readTable_('Attachments').filter(function (a) { return a.noteId === id; })
    .map(function (a) { return { fileId: a.fileId, filename: String(a.filename), mime: a.mime, bytes: +a.bytes }; });
  return { html: html, atts: atts, docId: n.docId, loc: n.source === 'link' ? linkLocation_(n.docId) : null };
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

/** One-time: rename the Drive folder and the Sheet to ArNote. */
function renameToArNote() {
  DriveApp.getFolderById(props_().getProperty('rootId')).setName('ArNote');
  ss_().rename('ArNote');
  return 'renamed';
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
  var now = new Date(), id = Utilities.getUuid(), owner = fileOwner_(f), acct = accountTagFor_(owner);
  tags = (tags || []).slice(); if (acct && tags.indexOf(acct) < 0) tags.push(acct);
  var n = { id: id, title: linkTitle_(f), notebook: notebook || 'Inbox', tags: tags.join(TAG_SEP), created: now, updated: now,
    docId: f.getId(), starred: false, source: 'link', sourceUrl: f.getUrl(), author: owner, attachments: 0 };
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
function photoPreviewHtml_(id, meta) {
  var src = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w1600&authuser=' + encodeURIComponent(Session.getEffectiveUser().getEmail());
  return '<html><body style="font:14px sans-serif;color:#444;text-align:center"><img src="' + src + '" style="max-width:100%;box-shadow:0 1px 6px rgba(0,0,0,.25)">' +
    '<p style="color:#888;font-size:12px">' + String(meta.name || '').replace(/[<>&"]/g, '') + '</p></body></html>';
}

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

/* ---------------- Email → ArNote ----------------
 * Two ways in: send/forward to <you>+arnote@<domain> (only from allowed senders), or put the Gmail label "ArNote" on any email.
 * Each email becomes an Inbox note (subject = title, the email = body, attachments attached, #Tags in the subject applied).
 * Done threads get the label "ArNote/Saved" and are archived.
 */
function mailAddress_() { var e = Session.getEffectiveUser().getEmail(), i = e.indexOf('@'); return e.slice(0, i) + '+arnote' + e.slice(i); }
function mailSenders_() {
  var v = []; try { v = JSON.parse(props_().getProperty('mailSenders') || '[]'); } catch (e) {}
  var me = Session.getEffectiveUser().getEmail().toLowerCase(); if (v.indexOf(me) < 0) v.unshift(me);
  return v;
}
function api_mailInfo() {
  var last = null; try { last = JSON.parse(props_().getProperty('mailLast') || 'null'); } catch (e) {}
  return { address: mailAddress_(), senders: mailSenders_(), me: Session.getEffectiveUser().getEmail().toLowerCase(), last: last };
}
function api_setMailSenders(list) {
  var clean = (list || []).map(function (s) { return String(s).trim().toLowerCase(); }).filter(function (s) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s); });
  props_().setProperty('mailSenders', JSON.stringify(clean.filter(function (s, i) { return clean.indexOf(s) === i; })));
  return mailSenders_();
}
function api_mailCheckNow() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) throw new Error('ArNote is busy; try again in a minute');
  try { return processMail_(Date.now(), 150000); } finally { lock.releaseLock(); }
}
/** Run once in the script editor to approve Gmail access. */
function authorizeMail() { GmailApp.getUserLabelByName('ArNote'); return mailAddress_(); }

function mailLabel_(name) { return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }

function processMail_(start, budget) {
  start = start || Date.now(); budget = budget || 60000;
  var addr = mailAddress_(), ok = {}, made = [];
  mailSenders_().forEach(function (s) { ok[s] = 1; });
  var inLabel = mailLabel_('ArNote'), saved = mailLabel_('ArNote/Saved');
  var threads = GmailApp.search('(to:' + addr + ' OR deliveredto:' + addr + ') -label:arnote-saved', 0, 15).concat(GmailApp.search('label:arnote -label:arnote-saved', 0, 15));
  var seen = {};
  threads.forEach(function (th) {
    if (seen[th.getId()] || Date.now() - start > budget) return; seen[th.getId()] = 1;
    var labelled = th.getLabels().some(function (l) { return l.getName() === 'ArNote'; });
    var msgs = th.getMessages().filter(function (m) {
      if (labelled) return true;
      var to = [m.getTo(), m.getCc(), m.getBcc ? m.getBcc() : '', m.getHeader('Delivered-To') || ''].join(',').toLowerCase();
      return to.indexOf(addr.toLowerCase()) >= 0 && ok[senderEmail_(m.getFrom())];
    });
    if (labelled) msgs = msgs.slice(-1);                          // a labelled conversation: keep its latest email
    msgs.forEach(function (m) { try { made.push(noteFromMail_(m, th)); } catch (e) { Logger.log('Mail note failed: ' + e); } });
    th.addLabel(saved); if (labelled) th.removeLabel(inLabel); th.moveToArchive();
  });
  if (made.length) props_().setProperty('mailLast', JSON.stringify({ at: new Date().toISOString(), n: made.length }));
  return { added: made.length, titles: made.map(function (n) { return n.title; }) };
}

function senderEmail_(from) { var m = String(from).match(/<([^>]+)>/); return (m ? m[1] : String(from)).trim().toLowerCase(); }

/** "Boiler quote #HVAC #MSBC Thomas" → title "Boiler quote", tags [HVAC, MSBC Thomas] (only tags that already exist). */
function subjectTags_(subject) {
  var s = String(subject || '').replace(/^\s*((fwd?|fw|re)\s*:\s*)+/i, '').trim(), tags = [];
  var names = readTable_('Tags').map(function (t) { return String(t.name); }).sort(function (a, b) { return b.length - a.length; });
  names.forEach(function (n) {
    var re = new RegExp('#' + n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=\\s|#|$)', 'i');
    if (re.test(s)) { tags.push(n); s = s.replace(re, ' '); }
  });
  return { title: s.replace(/\s+/g, ' ').trim() || 'Email', tags: tags };
}

function noteFromMail_(m, th) {
  var st = subjectTags_(m.getSubject()), esc = function (x) { return String(x || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var body = String(m.getBody() || '').replace(/<img[^>]+src=["']cid:[^"']*["'][^>]*>/gi, '');   // inline pictures are added below as real images
  var html = '<html><body><p>' + esc(headerLine_('Inbox', st.tags)) + '</p>' +
    '<p style="color:#666;font-size:10pt">From: ' + esc(m.getFrom()) + '<br>Date: ' + esc(Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')) +
    (m.getTo() ? '<br>To: ' + esc(m.getTo()) : '') + '</p><hr>' + body + '</body></html>';
  var bnd = 'arnote' + Date.now();
  var payload = '--' + bnd + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify({ name: st.title, mimeType: 'application/vnd.google-apps.document', parents: [props_().getProperty('notesId')] }) +
    '\r\n--' + bnd + '\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n' + html + '\r\n--' + bnd + '--';
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + bnd, payload: Utilities.newBlob(payload).getBytes(),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Could not make the note (' + res.getResponseCode() + ')');
  var docId = JSON.parse(res.getContentText()).id;
  var atts = m.getAttachments({ includeInlineImages: true, includeAttachments: true }).filter(function (a) { return a.getSize() > 0; });
  var files = [];
  if (atts.length) {
    var doc = DocumentApp.openById(docId), b = doc.getBody(), attFolder = DriveApp.getFolderById(props_().getProperty('attId'));
    atts.forEach(function (a) { var f = attFolder.createFile(a.copyBlob().setName(a.getName() || 'attachment')); files.push(f); appendFileToBody_(b, f); });
    doc.saveAndClose();
  }
  var rc = null; try { rc = mailReceipt_(m, files, docId); } catch (e) { Logger.log('Mail receipt: ' + e); }
  if (rc && st.tags.indexOf('Receipt') < 0) st.tags.push('Receipt');
  if (st.tags.length) ensureTags_(st.tags);
  var now = new Date(), id = Utilities.getUuid();
  var n = { id: id, title: st.title, notebook: 'Inbox', tags: st.tags.join(TAG_SEP), created: now, updated: now, docId: docId, starred: false,
    source: 'email', sourceUrl: 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(Session.getEffectiveUser().getEmail()) + '#all/' + th.getId(), author: senderEmail_(m.getFrom()), attachments: files.length };
  ss_().getSheetByName('Notes').appendRow(noteToRow_(n));
  var attSh = ss_().getSheetByName('Attachments');
  files.forEach(function (f) { attSh.appendRow([id, f.getId(), f.getName(), f.getMimeType(), f.getSize(), '']); });
  if (rc) receiptSheet_().appendRow(mailReceiptRow_(rc, id));
  return { id: id, title: st.title, receipt: !!rc };
}

/* ---------------- Backup ----------------
 * Every day (after 3 AM) and on demand: "ArNote Backups/<date>" gets a copy of every tab of the index spreadsheet
 * (notes list + tags + filing, photo tags and places, receipts, attachments) and settings.json.
 * The notes and photos themselves are Drive files (Drive keeps their history and a 30-day trash); the index is what this protects.
 * Keeps the 30 newest daily backups and the 10 newest others. Restore puts a backup's tabs back (after a safety backup of now). */
var BACKUP_KEEP_PROPS_ = ['shares', 'receiptCfg', 'places', 'mailSenders', 'bookmarkOrder', 'scratch', 'pinnedNote', 'watchFolders', 'accountTags'];
function backupRoot_() {
  var p = props_(), root = DriveApp.getFolderById(p.getProperty('rootId'));
  var f = getOrCreateFolder_(p.getProperty('backupId'), 'ArNote Backups', root);
  p.setProperty('backupId', f.getId()); return f;
}
function makeBackup_(label) {
  var tz = Session.getScriptTimeZone(), now = new Date(), ss = ss_();
  var name = Utilities.formatDate(now, tz, 'yyyy-MM-dd') + (label ? ' ' + Utilities.formatDate(now, tz, 'HH-mm') + ' (' + label + ')' : '');
  var folder = backupRoot_().createFolder(name);
  var copy = SpreadsheetApp.create('ArNote index ' + name), first = copy.getSheets()[0];
  ss.getSheets().forEach(function (sh) { if (sh.getName() !== 'Receipt report') sh.copyTo(copy).setName(sh.getName()); });
  copy.deleteSheet(first);
  DriveApp.getFileById(copy.getId()).moveTo(folder);
  var all = props_().getProperties(), keep = {};
  BACKUP_KEEP_PROPS_.forEach(function (k) { if (all[k] !== undefined) keep[k] = all[k]; });
  folder.createFile('settings.json', JSON.stringify({ made: now.toISOString(), version: 1, settings: keep }, null, 1), 'application/json');
  props_().setProperty('backupLast', JSON.stringify({ at: now.toISOString(), name: name }));
  pruneBackups_();
  return { name: name, id: folder.getId(), url: folder.getUrl() };
}
function pruneBackups_() {
  var it = backupRoot_().getFolders(), daily = [], other = [];
  while (it.hasNext()) { var f = it.next(); (/^\d{4}-\d\d-\d\d$/.test(f.getName()) ? daily : other).push(f); }
  [[daily, 30], [other, 10]].forEach(function (x) {
    x[0].sort(function (a, b) { return b.getName().localeCompare(a.getName()); }).slice(x[1]).forEach(function (f) { f.setTrashed(true); });
  });
}
function backupDaily_() {
  var tz = Session.getScriptTimeZone(), now = new Date(), today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  if (+Utilities.formatDate(now, tz, 'H') < 3) return false;
  var last = {}; try { last = JSON.parse(props_().getProperty('backupLast') || '{}'); } catch (e) {}
  var it = backupRoot_().getFoldersByName(today); if (it.hasNext()) return false;          // today's is done
  makeBackup_(''); return true;
}
function api_backupNow() { return makeBackup_('manual'); }
function api_backups() {
  var it = backupRoot_().getFolders(), out = [];
  while (it.hasNext()) { var f = it.next(); out.push({ id: f.getId(), name: f.getName(), url: f.getUrl() }); }
  out.sort(function (a, b) { return b.name.localeCompare(a.name); });
  var last = null; try { last = JSON.parse(props_().getProperty('backupLast') || 'null'); } catch (e) {}
  return { last: last, list: out, folderUrl: backupRoot_().getUrl() };
}
/** Put a backup back: every tab of its index copy replaces the same tab now, plus its settings. A safety backup of now is made first. */
function api_backupRestore(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  if (folder.getParents().next().getId() !== backupRoot_().getId()) throw new Error('Not an ArNote backup');
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS); if (!files.hasNext()) throw new Error('This backup has no index copy');
  var src = SpreadsheetApp.openById(files.next().getId());
  var safety = makeBackup_('before restore');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  var done = [];
  try {
    var ss = ss_();
    src.getSheets().forEach(function (s) {
      var name = s.getName(), dst = ss.getSheetByName(name); if (!dst || name === 'Receipt report') return;
      var v = s.getDataRange().getValues();
      dst.clearContents();
      if (v.length && v[0].length) dst.getRange(1, 1, v.length, v[0].length).setValues(v);
      done.push(name + ' (' + Math.max(0, v.length - 1) + ')');
    });
    var js = folder.getFilesByName('settings.json');
    if (js.hasNext()) {
      var st = (JSON.parse(js.next().getBlob().getDataAsString()) || {}).settings || {};
      BACKUP_KEEP_PROPS_.forEach(function (k) { if (st[k] !== undefined) props_().setProperty(k, st[k]); });
    }
    props_().deleteProperty('rcReportHash');
  } finally { lock.releaseLock(); }
  try { buildReceiptReport_(true); } catch (e) {}
  return { restored: done, safety: safety.name };
}

/* ---------------- Emailed receipts and invoices ----------------
 * A forwarded (or labelled) email that reads as a receipt/invoice gets the Receipt tag and a row in Receipts.
 * Text comes from the email body, then from an attached PDF (Drive reads it, scanned PDFs too).
 * Picture for reports: the PDF if there is one, otherwise the email note itself. */
function fwdHeader_(plain) {
  var i = plain.lastIndexOf('Forwarded message'); if (i < 0) return null;
  var blk = plain.slice(i, i + 2000), from = (blk.match(/\nFrom:\s*(.+)/) || [])[1] || '', date = (blk.match(/\nDate:\s*(.+)/) || [])[1] || '';
  var em = (from.match(/<([^>]+)>/) || from.match(/([\w.+-]+@[\w.-]+)/) || [])[1] || '';
  var name = from.replace(/<[^>]*>/g, '').replace(/["']/g, '').trim(); if (name.indexOf('@') >= 0) name = '';
  var d = date ? new Date(date.replace(/\s+at\s+/, ' ')) : null; if (d && isNaN(d.getTime())) d = null;
  var to = blk.search(/\nTo:[^\n]*\n/), start = to >= 0 ? i + to + blk.slice(to + 1).indexOf('\n') + 2 : i + 200;
  return { name: name, email: em.toLowerCase(), date: d, start: start };
}
function pdfText_(fileId) {
  var hdr = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var cp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy?ocrLanguage=en&fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: 'ArNote read (temporary)', mimeType: 'application/vnd.google-apps.document' }), headers: hdr, muteHttpExceptions: true });
  if (cp.getResponseCode() !== 200) return '';
  var id = JSON.parse(cp.getContentText()).id;
  try {
    var t = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text/plain', { headers: hdr, muteHttpExceptions: true });
    return t.getResponseCode() === 200 ? t.getContentText('UTF-8').replace(/^\uFEFF/, '') : '';
  } finally { UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id, { method: 'delete', headers: hdr, muteHttpExceptions: true }); }
}
function mailVendor_(fh, subj, fromHdr) {
  var generic = /^(no-?reply|noreply|do not reply|info|billing|accounting|accounts?( payable| receivable)?|sales|support|orders?|customer ?service|notifications?|receipts?|invoices?|service|team|hello|admin|office)$/i;
  var name = fh ? fh.name : String(fromHdr || '').replace(/<[^>]*>/g, '').replace(/["']/g, '').trim();
  if (name.indexOf('@') >= 0) name = '';
  if (name && !generic.test(name)) return name.slice(0, 40);
  var m = subj.match(/\b(?:receipt|invoice|order|payment)s? from (.+?)(?:\s*[#(\[].*)?$/i);
  if (m) return m[1].replace(/[.\s]+$/, '').slice(0, 40);
  m = subj.match(/^(?:your\s+)?(.+?)\s+(?:invoice|receipt|order confirmation|statement|bill)\b/i);
  if (m && !/^(your|thank|new|the|an?|order|payment)\b/i.test(m[1]) && m[1].split(' ').length <= 5) return m[1].slice(0, 40);
  var em = fh ? fh.email : senderEmail_(fromHdr), dom = (em.split('@')[1] || '').split('.'), stem = dom.length > 1 ? dom[dom.length - 2] : dom[0] || '';
  return stem ? stem.charAt(0).toUpperCase() + stem.slice(1) : '';
}
function mailReceipt_(m, files, docId) {
  var subj = String(m.getSubject() || '').replace(/^\s*((fwd?|fw|re)\s*:\s*)+/i, '').trim(), plain = String(m.getPlainBody() || ''), fh = fwdHeader_(plain);
  var body = fh ? plain.slice(fh.start) : plain, sig = body.lastIndexOf('\n--');
  if (sig >= 0 && body.length - sig < 600) body = body.slice(0, sig);                      // your own signature
  var pdfs = (files || []).filter(function (f) { return /pdf/i.test(f.getMimeType()) || /\.pdf$/i.test(f.getName()); });
  var hint = /\b(receipt|invoice|inv\b|order|purchase|payment|paid|bill|statement)/i.test(subj + ' ' + pdfs.map(function (f) { return f.getName(); }).join(' '));
  var r = classifyText_(body), text = body, pdfT = '';
  if ((r.kind !== 'receipt' || r.fields.total === null) && pdfs.length) {
    pdfT = pdfText_(pdfs[0].getId());
    var r2 = classifyText_(pdfT);
    if (r2.kind === 'receipt' && r2.fields.total !== null) { r = r2; text = pdfT; }
  }
  if (r.kind !== 'receipt') {
    if (!hint || !moneyIn_(body + '\n' + pdfT).length) return null;
    var both = (body + '\n' + pdfT).split(/\r?\n/).map(function (l) { return l.replace(/\s+/g, ' ').trim(); }).filter(String);
    r = { kind: 'receipt', fields: receiptFields_(both, both.join('\n')) };
  }
  var f = r.fields, sent = (fh && fh.date) || m.getDate(), tz = Session.getScriptTimeZone();
  var date = f.date, dd = date ? new Date(date + 'T12:00:00') : null;
  if (!dd || isNaN(dd.getTime()) || dd > new Date(sent.getTime() + 2 * 864e5) || dd < new Date(sent.getTime() - 400 * 864e5)) date = Utilities.formatDate(sent, tz, 'yyyy-MM-dd');
  var vendor = mailVendor_(fh, subj, m.getFrom()) || f.vendor;
  var pic = pdfs.length ? pdfs[0].getId() : docId;
  return { date: date, vendor: vendor, total: f.total, tax: f.tax, card: f.card || '', pay: f.pay || '', fileId: pic,
    kind: /invoice|\binv\b|statement|bill\b/i.test(subj + ' ' + pdfs.map(function (x) { return x.getName(); }).join(' ')) && !/receipt/i.test(subj) ? 'invoice' : 'receipt',
    check: (!vendor || f.total === null) ? 'check' : '' };
}
function mailReceiptRow_(rc, noteId) {
  return [rc.date, rc.vendor, rc.total === null ? '' : rc.total, rc.tax === null ? '' : rc.tax, rc.card, '', rc.check, noteId, rc.fileId,
    'https://drive.google.com/file/d/' + rc.fileId + '/view', new Date().toISOString(), rc.pay, '', '', rc.kind];
}
/** Read email notes already in ArNote as receipts (also the "It's a receipt" button). Returns what was found for each. */
function api_mailReceiptFor(noteIds, force) {
  var want = {}; (noteIds || []).forEach(function (id) { want[id] = 1; });
  var have = {}, sh = receiptSheet_(), k = RECEIPT_COLS.indexOf('noteId');
  if (sh.getLastRow() > 1) sh.getRange(2, k + 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0])] = 1; });
  var atts = {}; readTable_('Attachments').forEach(function (a) { var id = String(a.noteId); (atts[id] = atts[id] || []).push(a); });
  var out = [], start = Date.now();
  readTable_('Notes').forEach(function (n) {
    if (!want[n.id] || Date.now() - start > 200000) return;
    if (have[n.id]) { out.push({ id: n.id, title: String(n.title), already: true }); return; }
    var tid = (String(n.sourceUrl || '').match(/#all\/([\w-]+)/) || [])[1]; if (!tid) { out.push({ id: n.id, title: String(n.title), error: 'not an email' }); return; }
    try {
      var msgs = GmailApp.getThreadById(tid).getMessages(), m = msgs[msgs.length - 1];
      var files = (atts[n.id] || []).map(function (a) { try { return DriveApp.getFileById(String(a.fileId)); } catch (e) { return null; } }).filter(Boolean);
      var rc = mailReceipt_(m, files, String(n.docId));
      if (!rc && force) rc = { date: Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), vendor: mailVendor_(fwdHeader_(m.getPlainBody() || ''), String(m.getSubject()).replace(/^\s*((fwd?|fw|re)\s*:\s*)+/i, ''), m.getFrom()), total: null, tax: null, card: '', pay: '', fileId: files.length ? files[0].getId() : String(n.docId), kind: 'receipt', check: 'check' };
      if (!rc) { out.push({ id: n.id, title: String(n.title), receipt: false }); return; }
      sh.appendRow(mailReceiptRow_(rc, n.id)); have[n.id] = 1;
      api_batchUpdate([n.id], { addTags: ['Receipt'] });
      out.push({ id: n.id, title: String(n.title), receipt: true, fields: rc });
    } catch (e) { out.push({ id: n.id, title: String(n.title), error: String(e.message || e) }); }
  });
  try { buildReceiptReport_(); } catch (e) {}
  return out;
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

/* ---- Drive browser for the "Link Drive files" dialog ----
 * where: 'root' (My Drive), 'shared' (Shared with me) or a folder id; q: search by name instead.
 */
function driveApi_(path) {
  var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/' + path, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Drive error ' + r.getResponseCode());
  return JSON.parse(r.getContentText());
}

function api_driveList(where, q) {
  where = where || 'root';
  var FOLDER = 'application/vnd.google-apps.folder', p = props_();
  var arnoteRoot = p.getProperty('rootId');
  var myRoot = p.getProperty('myDriveRoot');
  if (!myRoot) { myRoot = driveApi_('files/root?fields=id').id; p.setProperty('myDriveRoot', myRoot); }
  var query;
  if (q) query = "name contains '" + String(q).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "' and trashed = false";
  else if (where === 'shared') query = 'sharedWithMe and trashed = false';
  else query = "'" + (where === 'root' ? myRoot : where) + "' in parents and trashed = false";
  var res = driveApi_('files?pageSize=300&orderBy=folder,name_natural&supportsAllDrives=true&includeItemsFromAllDrives=true' +
    '&fields=files(id,name,mimeType,modifiedTime,shortcutDetails,owners(emailAddress))&q=' + encodeURIComponent(query));
  var linked = {}; readTable_('Notes').forEach(function (n) { if (n.source === 'link') linked[n.docId] = 1; });
  var synced = {}; watchList_().forEach(function (id) { synced[id] = 1; });
  var items = (res.files || []).map(function (f) {
    var id = f.id, mt = f.mimeType;
    if (mt === 'application/vnd.google-apps.shortcut' && f.shortcutDetails) { id = f.shortcutDetails.targetId; mt = f.shortcutDetails.targetMimeType; }
    return { id: id, name: f.name, folder: mt === FOLDER, mime: mt, modified: f.modifiedTime, owner: f.owners && f.owners[0] && f.owners[0].emailAddress,
      linked: !!linked[id], synced: !!synced[id] };
  }).filter(function (x) { return x.id !== arnoteRoot; });   // don't offer ArNote's own folder
  items.sort(function (a, b) { return (b.folder - a.folder) || a.name.localeCompare(b.name, undefined, { numeric: true }); });
  // breadcrumb for a folder: walk up to My Drive (or to the top of a shared folder)
  var path = [];
  if (!q && where !== 'root' && where !== 'shared') {
    var id = where, guard = 0;
    while (id && id !== myRoot && guard++ < 12) {
      var f = null; try { f = driveApi_('files/' + id + '?supportsAllDrives=true&fields=id,name,parents'); } catch (e) {}
      if (!f) break;   // parent not visible to this account (e.g. inside someone else's drive)
      path.unshift({ id: f.id, name: f.name });
      id = f.parents && f.parents[0];
    }
    if (id !== myRoot) path.unshift({ id: 'shared', name: 'Shared with me' });
  }
  return { where: where, q: q || '', path: path, items: items, synced: !!synced[where] };
}

/* ---------------- Account tags ----------------
 * Linked files get a tag for the Google account that owns them (under the parent tag "Accounts").
 * The tag is an ordinary tag: rename it on the Tags page and the account keeps using the new name.
 * Map lives in Script Property 'accountTags' {email: tagName}; '' means "don't tag this account".
 */
function accountMap_() { try { return JSON.parse(props_().getProperty('accountTags') || '{}'); } catch (e) { return {}; } }
function saveAccountMap_(m) { props_().setProperty('accountTags', JSON.stringify(m)); }

function defaultAccountTag_(email) {
  var parts = String(email).toLowerCase().split('@'), local = parts[0], dom = (parts[1] || '').split('.')[0];
  return '@' + local + (dom && dom !== 'gmail' && dom !== 'googlemail' ? '-' + dom : '');
}

function accountTagFor_(email) {
  if (!email) return '';
  var m = accountMap_();
  if (!(email in m)) {
    m[email] = defaultAccountTag_(email);
    saveAccountMap_(m);
    var par = accountParent_(m); if (par) ensureTagParent_(par, ''); ensureTagParent_(m[email], par);
  }
  return m[email];
}

/** New account tags go wherever the existing ones live (so renaming or moving "Accounts" is respected). */
function accountParent_(m) {
  var parent = {}; readTable_('Tags').forEach(function (t) { parent[String(t.name)] = String(t.parent || ''); });
  for (var e in m) if (m[e] && parent.hasOwnProperty(m[e])) return parent[m[e]];
  return 'Accounts';
}

function fileOwner_(f) { try { var o = f.getOwner(); return o ? o.getEmail() : ''; } catch (e) { return ''; } }

/** Settings → Accounts: tags any linked notes that don't have their account tag yet, then lists accounts. */
function api_accounts() {
  var t = notesTable_(), counts = {}, start = Date.now();
  t.rows.forEach(function (r, i) {
    if (r[col_('source')] !== 'link') return;
    var email = String(r[col_('author')] || '');
    if (!email && Date.now() - start < 60000) {
      try { email = fileOwner_(DriveApp.getFileById(r[col_('docId')])); } catch (e) {}
      if (email) t.sh.getRange(i + 2, col_('author') + 1).setValue(email);
    }
    if (!email) return;
    counts[email] = (counts[email] || 0) + 1;
    var tag = accountTagFor_(email), tags = splitTags_(r[col_('tags')]);
    if (tag && tags.indexOf(tag) < 0) { tags.push(tag); t.sh.getRange(i + 2, col_('tags') + 1).setValue(tags.join(TAG_SEP)); }
  });
  var m = accountMap_();
  return Object.keys(m).map(function (e) { return { email: e, tag: m[e], count: counts[e] || 0 }; })
    .sort(function (a, b) { return b.count - a.count; });
}

/** Keep account tags and share links pointing at the right tag after a rename / delete. */
function tagRenamed_(oldName, newName) {
  retagPhotos_(function (tg) { var out = []; tg.forEach(function (g) { g = g === oldName ? newName : g; if (out.indexOf(g) < 0) out.push(g); }); return out; });
  var m = accountMap_(), ch = false;
  Object.keys(m).forEach(function (e) { if (m[e] === oldName) { m[e] = newName; ch = true; } });
  if (ch) saveAccountMap_(m);
  var s = shares_(), ch2 = false;
  s.forEach(function (x) { if (x.tag === oldName) { x.tag = newName; ch2 = true; } });
  if (ch2) saveShares_(s);
  var rc = receiptCfg_(), ch3 = false;
  rc.cols.forEach(function (c) { if (c.parent === oldName) { c.parent = newName; ch3 = true; } });
  if (ch3) props_().setProperty('receiptCfg', JSON.stringify(rc));
}
function tagDeleted_(name) {
  retagPhotos_(function (tg) { return tg.filter(function (g) { return g !== name; }); });
  var m = accountMap_(), ch = false;
  Object.keys(m).forEach(function (e) { if (m[e] === name) { m[e] = ''; ch = true; } });   // stop tagging that account
  if (ch) saveAccountMap_(m);
  var s = shares_(), ch2 = false;
  s.forEach(function (x) { if (x.tag === name && x.active) { x.active = false; ch2 = true; } });
  if (ch2) saveShares_(s);
}

/* ---------------- Share links ----------------
 * A share link shows one tag's notes (including its sub-tags) read-only: list, preview, print, attachments.
 * Each link has its own secret token; the app sends {s: token, fn: 'share_x'} instead of the private key.
 * Script Property 'shares' = [{id, token, tag, label, created, active}]
 */
function shares_() { try { return JSON.parse(props_().getProperty('shares') || '[]'); } catch (e) { return []; } }
function saveShares_(s) { props_().setProperty('shares', JSON.stringify(s)); }

function shareScope_(tag) {
  var kids = {}; readTable_('Tags').forEach(function (t) { var p = String(t.parent || ''); (kids[p] = kids[p] || []).push(String(t.name)); });
  var set = {}; (function walk(x) { set[x] = 1; (kids[x] || []).forEach(walk); })(tag);
  var notes = readTable_('Notes').filter(function (n) { return splitTags_(n.tags).some(function (g) { return set[g]; }); });
  var ids = {}; notes.forEach(function (n) { ids[n.id] = n; });
  return { notes: notes, ids: ids, set: set };
}

function api_shares() {
  var counts = {};
  return shares_().map(function (s) {
    if (!(s.tag in counts)) counts[s.tag] = shareScope_(s.tag).notes.length;
    return { id: s.id, token: s.token, tag: s.tag, label: s.label, created: s.created, active: !!s.active, count: counts[s.tag] };
  });
}

function api_createShare(tag, label) {
  tag = String(tag || '').trim(); if (!tag) throw new Error('Pick a tag to share');
  var s = shares_();
  var token = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now())).replace(/=+$/, '');
  var x = { id: Utilities.getUuid().slice(0, 8), token: token, tag: tag, label: String(label || '').trim() || 'Shared link', created: new Date().toISOString(), active: true };
  s.push(x); saveShares_(s);
  x.count = shareScope_(tag).notes.length;
  return x;
}

function api_updateShare(id, ch) {
  var s = shares_(), x = s.filter(function (y) { return y.id === id; })[0];
  if (!x) throw new Error('Link not found');
  if (ch.label !== undefined) x.label = String(ch.label).trim() || x.label;
  if (ch.active !== undefined) x.active = !!ch.active;
  saveShares_(s); return true;
}

function api_deleteShare(id) { saveShares_(shares_().filter(function (y) { return y.id !== id; })); return true; }

/** Everything a share link is allowed to do. Nothing outside the shared tag is reachable. */
function shareDispatch_(token, fn, a) {
  var sh = shares_().filter(function (x) { return x.token === token; })[0];
  if (!sh || !sh.active) throw new Error('This link has been switched off. Ask the person who shared it for a new one.');
  var sc = shareScope_(sh.tag);
  function inScope(id) { return !!sc.ids[id]; }
  if (fn === 'share_info') return { label: sh.label, tag: sh.tag, count: sc.notes.length };
  if (fn === 'share_list') return sc.notes.map(function (n) {
    return { id: n.id, title: String(n.title), updated: iso_(n.updated), created: iso_(n.created), link: n.source === 'link', att: +n.attachments || 0,
      tags: splitTags_(n.tags).filter(function (g) { return sc.set[g]; }) };
  });
  if (fn === 'share_view') {
    if (!inScope(a[0])) throw new Error('Not part of this shared list');
    var v = api_noteView(a[0], { embed: true });   // guests aren't signed in to Drive: send the picture itself
    return { html: v.html, atts: (v.atts || []).map(function (x) { return { fileId: x.fileId, filename: x.filename, bytes: x.bytes, mime: x.mime }; }) };
  }
  if (fn === 'share_print') return api_printPdf((a[0] || []).filter(inScope));
  if (fn === 'share_thumbs') return api_thumbs((a[0] || []).filter(inScope));
  if (fn === 'share_att') {
    if (!inScope(a[0])) throw new Error('Not part of this shared list');
    var att = readTable_('Attachments').filter(function (x) { return x.noteId === a[0] && x.fileId === a[1]; })[0];
    if (!att) throw new Error('Attachment not found');
    var f = DriveApp.getFileById(att.fileId);
    if (f.getSize() > 20 * 1024 * 1024) throw new Error('This file is too big to open here');
    return { name: f.getName(), mime: f.getMimeType(), b64: Utilities.base64Encode(f.getBlob().getBytes()) };
  }
  throw new Error('Not allowed');
}

/** Where a linked file lives: owning account and folder path (cached 6 hours). */
function linkLocation_(id) {
  var cache = CacheService.getScriptCache(), key = 'loc:' + id, hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  var out = null;
  try {
    var f = driveApi_('files/' + id + '?supportsAllDrives=true&fields=name,parents,driveId,owners(emailAddress,displayName,me)');
    var o = (f.owners || [])[0] || {};
    var path = [], pid = f.parents && f.parents[0], guard = 0, top = null;
    while (pid && guard++ < 15) {
      var p = null; try { p = driveApi_('files/' + pid + '?supportsAllDrives=true&fields=id,name,parents'); } catch (e) {}
      if (!p) break;
      path.unshift({ id: p.id, name: p.name }); top = p;
      pid = p.parents && p.parents[0];
    }
    // Reached the account's own My Drive only if the top folder has no parent and the file is ours; otherwise the
    // folders above are in someone else's Drive and not visible from here.
    var myRoot = props_().getProperty('myDriveRoot');
    var full = !!top && (top.id === myRoot || (!pid && !!o.me));
    out = { owner: o.emailAddress || '', ownerName: o.displayName || '', me: !!o.me, shared: !!f.driveId, path: path, partial: !full };
  } catch (e) { out = null; }
  if (out) cache.put(key, JSON.stringify(out), 21600);
  return out;
}

/** Cabinet → By Drive: for every linked note, the owning account and its folder chain (parallel Drive calls, cached 6 hours). */
function api_linkTree() {
  var notes = readTable_('Notes').filter(function (n) { return n.source === 'link'; });
  var cache = CacheService.getScriptCache(), hdr = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var base = 'https://www.googleapis.com/drive/v3/files/';
  function fetchMany(ids, fields) {
    var out = {};
    for (var i = 0; i < ids.length; i += 50) {
      var part = ids.slice(i, i + 50);
      var res = UrlFetchApp.fetchAll(part.map(function (id) {
        return { url: base + id + '?supportsAllDrives=true&fields=' + encodeURIComponent(fields), headers: hdr, muteHttpExceptions: true };
      }));
      res.forEach(function (r, k) { out[part[k]] = r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null; });
    }
    return out;
  }
  function cachedMany(prefix, ids, fields, shape) {
    var out = {}, need = [], got = {};
    for (var i = 0; i < ids.length; i += 100) {
      var g = cache.getAll(ids.slice(i, i + 100).map(function (id) { return prefix + id; }));
      for (var key in g) got[key] = g[key];
    }
    ids.forEach(function (id) { var v = got[prefix + id]; if (v) out[id] = JSON.parse(v); else need.push(id); });
    if (need.length) {
      var f = fetchMany(need, fields), put = {};
      need.forEach(function (id) { out[id] = shape(f[id]); put[prefix + id] = JSON.stringify(out[id]); });
      try { cache.putAll(put, 21600); } catch (e) {}
    }
    return out;
  }
  var docIds = []; notes.forEach(function (n) { if (docIds.indexOf(n.docId) < 0) docIds.push(n.docId); });
  var files = cachedMany('lt:', docIds, 'parents,owners(emailAddress),driveId', function (f) {
    return f ? { owner: (f.owners && f.owners[0] && f.owners[0].emailAddress) || '', parent: (f.parents || [])[0] || '', drive: f.driveId || '' } : { owner: '', parent: '', drive: '' };
  });
  var folders = {}, level = [];
  docIds.forEach(function (d) { var p = files[d].parent; if (p && level.indexOf(p) < 0) level.push(p); });
  for (var depth = 0; depth < 15 && level.length; depth++) {
    var got = cachedMany('lf:', level, 'name,parents', function (f) { return f ? { name: f.name, parent: (f.parents || [])[0] || '' } : { name: null, parent: '' }; });
    var next = [];
    level.forEach(function (id) { folders[id] = got[id]; var p = got[id].parent; if (p && !folders[p] && next.indexOf(p) < 0) next.push(p); });
    level = next;
  }
  var p = props_(), root = p.getProperty('myDriveRoot');
  if (!root) { try { root = driveApi_('files/root?fields=id').id; p.setProperty('myDriveRoot', root); } catch (e) {} }
  var out = {};
  notes.forEach(function (n) { out[n.id] = files[n.docId]; });
  return { root: root || '', me: Session.getEffectiveUser().getEmail(), notes: out, folders: folders };
}

/* ---------------- Photos ----------------
 * Photos stay where they are in Drive (the "ArNote Photos" folder and everything under it). ArNote keeps only an index
 * in the Photos sheet: tags, inbox/filed, and when each photo was taken. Nothing is copied.
 */
var PHOTO_COLS = ['fileId', 'name', 'taken', 'mime', 'folder', 'tags', 'status', 'added', 'w', 'h', 'scan', 'lat', 'lng', 'town', 'region'];

function photoSheet_() {
  var sh = ensureSheet_('Photos', PHOTO_COLS);
  if (!props_().getProperty('photoFmt')) { sh.getRange('A:G').setNumberFormat('@'); props_().setProperty('photoFmt', '1'); }   // keep dates/ids as plain text
  if (sh.getLastColumn() < PHOTO_COLS.length || sh.getRange(1, PHOTO_COLS.length).getValue() !== PHOTO_COLS[PHOTO_COLS.length - 1]) sh.getRange(1, 1, 1, PHOTO_COLS.length).setValues([PHOTO_COLS]).setFontWeight('bold');
  return sh;
}

function photoRoot_() {
  var p = props_(), id = p.getProperty('photoRootId');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var it = DriveApp.getRootFolder().getFoldersByName('ArNote Photos');
  var folder = it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder('ArNote Photos');
  p.setProperty('photoRootId', folder.getId());
  return folder;
}

/** "2019:07:04 12:34:56" (camera) → ISO; otherwise a date in the file name (IMG_20190704_123456, PXL_…, 2019-07-04 …). */
function photoTaken_(f, path) {
  var t = f.imageMediaMetadata && f.imageMediaMetadata.time;
  var m = t && String(t).match(/^(\d{4}):(\d\d):(\d\d)[ T](\d\d):(\d\d):(\d\d)/);
  if (m && +m[1] > 1980) return m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':' + m[6];
  m = String(f.name).match(/(19[89]\d|20\d\d)[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])(?:[-_ T.]?([01]\d|2[0-3])[-_.:]?([0-5]\d)[-_.:]?([0-5]\d))?/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3] + 'T' + (m[4] || '12') + ':' + (m[5] || '00') + ':' + (m[6] || '00');
  m = String(path || '').match(/Photos from ((?:19|20)\d\d)\b/);     // Google Takeout year folder: no camera date, no date in the name
  if (m) return m[1] + '-01-01T00:00:00';
  return String(f.createdTime || '').slice(0, 19);
}

/** Find new photos/videos under ArNote Photos (all sub-folders) and add them to the index; drop ones that were deleted. */
function syncPhotos_(start) {
  start = start || Date.now();
  var root = photoRoot_().getId(), sh = photoSheet_();
  var vals = sh.getDataRange().getValues(); vals.shift();
  var known = {}; vals.forEach(function (r, i) { known[r[0]] = i; });
  var seen = {}, add = [], fix = [], queue = [{ id: root, path: '' }], complete = true, now = new Date().toISOString();
  while (queue.length) {
    if (Date.now() - start > 240000) { complete = false; break; }
    var dir = queue.shift(), token = '';
    do {
      var res = driveApi_('files?pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&q=' +
        encodeURIComponent("'" + dir.id + "' in parents and trashed = false") +
        '&fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType,createdTime,imageMediaMetadata(time,width,height,location),videoMediaMetadata(width,height))') +
        (token ? '&pageToken=' + token : ''));
      (res.files || []).forEach(function (f) {
        if (f.mimeType === 'application/vnd.google-apps.folder') { queue.push({ id: f.id, path: dir.path ? dir.path + '/' + f.name : f.name }); return; }
        if (!/^(image|video)\//.test(f.mimeType)) return;          // skip Takeout .json sidecars and anything else
        seen[f.id] = 1;
        if (known[f.id] !== undefined) {                         // already indexed: re-date undated photos from their Takeout year folder
          var ki = known[f.id], t2 = ki >= 0 && /Photos from/.test(dir.path) ? photoTaken_(f, dir.path) : '';
          if (t2 && /-01-01T00:00:00$/.test(t2) && t2 !== String(vals[ki][2])) fix.push([ki + 2, t2]);
          return;
        }
        var md = f.imageMediaMetadata || f.videoMediaMetadata || {};
        var tk = photoTaken_(f, dir.path), recent = tk && (Date.now() - new Date(tk).getTime()) < 60 * 864e5;   // only new photos get read for receipts/cards
        var gl = (f.imageMediaMetadata || {}).location, hasL = gl && (gl.latitude || gl.longitude);
        add.push([f.id, f.name, tk, f.mimeType, dir.path, /^video\//.test(f.mimeType) ? VIDEO_TAG : '', '', now, md.width || '', md.height || '', recent && /^image\//.test(f.mimeType) ? 'q' : '',
          hasL ? Math.round(gl.latitude * 1e5) / 1e5 : (/^image\//.test(f.mimeType) ? '-' : ''), hasL ? Math.round(gl.longitude * 1e5) / 1e5 : '', '', '']);
        known[f.id] = -1;
      });
      token = res.nextPageToken || '';
    } while (token && Date.now() - start < 240000);
    if (token) { complete = false; break; }
  }
  if (add.length <= 25) add.forEach(function (a) {      // a few photos added by hand get read whatever their date; bulk imports and Takeout year folders don't
    if (!a[10] && /^image\//.test(a[3]) && !/^Photos from /.test(a[4])) a[10] = 'q';
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, PHOTO_COLS.length).setValues(add);
  fix.forEach(function (x) { sh.getRange(x[0], 3).setValue(x[1]); });
  var removed = 0;
  if (complete) {                                                  // whole tree read: forget photos that are gone
    for (var i = vals.length - 1; i >= 0; i--) if (!seen[vals[i][0]]) { sh.deleteRow(i + 2); removed++; }
  }
  props_().setProperty('photoSync', JSON.stringify({ at: now, complete: complete, added: add.length, removed: removed }));
  return { added: add.length, removed: removed, complete: complete };
}

/**
 * Google Takeout saves every edit as its own file: x.jpg, x-edited.jpg, x-edited(1).jpg …
 * One per shot is kept (the "-edited" version when there is one, as seen in Google Photos);
 * the rest go to Drive trash (recoverable for 30 days) and leave the index. Same shot = same
 * folder, base name, time taken and size, so different photos that share a name are left alone.
 */
function photoDedupe_(budget, dry) {
  var start = Date.now(), rows = photoRows_(), groups = {};
  var byName = {};
  rows.forEach(function (p) {
    var m = p.name.match(/^(.*)(\.[^.]+)$/); if (!m || !p.taken) return;
    var stem = m[1], ed = /-edited/i.test(stem), copies = (stem.match(/\(\d+\)/g) || []).length;
    var base = stem.replace(/-edited/ig, '').replace(/\(\d+\)/g, '').toLowerCase();      // x(0)-edited(1)(1) → x
    var k = [p.folder, base, m[2].toLowerCase()].join('|');
    (byName[k] = byName[k] || []).push({ p: p, ed: ed, rank: (ed ? 0 : 2) * 10 + copies });
  });
  Object.keys(byName).forEach(function (k) {
    var g = byName[k];
    if (g.some(function (x) { return x.ed; })) { groups[k] = g; return; }       // "-edited" means an edit of this shot, even if rotated or cropped
    g.forEach(function (x) {                                                 // plain copies: only when time taken and shape also match
      var k2 = k + '|' + x.p.taken + '|' + [x.p.w, x.p.h].sort().join('x');
      (groups[k2] = groups[k2] || []).push(x);
    });
  });
  var drop = [], keepTags = {}, nGroups = 0;
  Object.keys(groups).forEach(function (k) {
    var g = groups[k]; if (g.length < 2) return;
    nGroups++;
    g.sort(function (a, b) { return a.rank - b.rank || a.p.name.length - b.p.name.length || (a.p.name < b.p.name ? -1 : 1); });
    var keep = g[0].p, tags = keep.tags.slice(), status = keep.status;
    g.slice(1).forEach(function (x) {
      x.p.tags.forEach(function (t) { if (tags.indexOf(t) < 0) tags.push(t); });   // keep any tagging done on a copy
      if (!status && x.p.status) status = x.p.status;
      drop.push(x.p.id);
    });
    if (tags.length !== keep.tags.length || status !== keep.status) keepTags[keep.id] = { tags: tags, status: status };
  });
  // the same shot in two folders (a Takeout copy and a phone upload): same file name, time taken and shape → keep the first one
  var gone1 = {}; drop.forEach(function (id) { gone1[id] = 1; });
  var cross = {};
  rows.forEach(function (p) { if (gone1[p.id] || !p.taken || !p.w) return;
    var k = [p.name.toLowerCase(), p.taken, [p.w, p.h].sort().join('x')].join('|'); (cross[k] = cross[k] || []).push(p); });
  Object.keys(cross).forEach(function (k) {
    var g = cross[k]; if (g.length < 2) return; nGroups++;
    var keep = g[0], kt = keepTags[keep.id] || { tags: keep.tags.slice(), status: keep.status }, ch = !!keepTags[keep.id];
    g.slice(1).forEach(function (p) {
      p.tags.forEach(function (t) { if (kt.tags.indexOf(t) < 0) { kt.tags.push(t); ch = true; } });
      if (!kt.status && p.status) { kt.status = p.status; ch = true; }
      drop.push(p.id);
    });
    if (ch) keepTags[keep.id] = kt;
  });
  if (dry || !drop.length) return { groups: nGroups, extra: drop.length, trashed: 0, left: drop.length };
  var gone = trashFiles_(drop, budget - (Date.now() - start)), n = Object.keys(gone).length;
  dropPhotoRows_(gone, keepTags);
  return { groups: nGroups, extra: drop.length, trashed: n, left: drop.length - n };
}

/** Move Drive files to the trash (restorable for 30 days), 20 at a time. Returns {id:1} for the ones done (or already gone). */
function trashFiles_(ids, budget) {
  var start = Date.now(), gone = {}, tok = ScriptApp.getOAuthToken();
  for (var i = 0; i < ids.length && Date.now() - start < budget; i += 20) {
    var part = ids.slice(i, i + 20);
    UrlFetchApp.fetchAll(part.map(function (id) {
      return { url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?supportsAllDrives=true', method: 'patch', contentType: 'application/json',
        payload: JSON.stringify({ trashed: true }), headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
    })).forEach(function (r, j) { var c = r.getResponseCode(); if (c === 200 || c === 404) gone[part[j]] = 1; });
  }
  return gone;
}

/** Rewrite the Photos sheet without the given photos (and with any merged tags on the ones kept). */
function dropPhotoRows_(gone, keepTags) {
  keepTags = keepTags || {};
  if (!Object.keys(gone).length && !Object.keys(keepTags).length) return;
  var sh = photoSheet_(), v = sh.getDataRange().getValues(), head = v.shift(); if (!v.length) return;
  var out = v.filter(function (r) { return !gone[r[0]]; }).map(function (r) {
    var kt = keepTags[r[0]]; if (kt) { r[5] = kt.tags.join(TAG_SEP); r[6] = kt.status; }
    return r;
  });
  sh.getRange(2, 1, v.length, head.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, head.length).setValues(out);
}

/** Delete photos from the app: files go to Drive trash, rows leave the index. Only files in the Photos index can be deleted. */
function api_photoDelete(ids) {
  var known = {}; photoRows_().forEach(function (p) { known[p.id] = 1; });
  ids = (ids || []).map(String).filter(function (id) { return known[id]; });
  if (!ids.length) return { deleted: 0 };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { var gone = trashFiles_(ids, 240000); dropPhotoRows_(gone); return { deleted: Object.keys(gone).length, failed: ids.length - Object.keys(gone).length }; }
  finally { lock.releaseLock(); }
}

/** Videos get the "Videos" tag (new ones on sync; this catches up the ones already indexed). */
var VIDEO_TAG = 'Videos';
/** Still needs tagging: no tags except the automatic Videos tag. */
function untagged_(p) { return !p.tags.some(function (t) { return t !== VIDEO_TAG; }); }
function api_photoTagVideos() {
  ensureTags_([VIDEO_TAG]);
  var sh = photoSheet_(); if (sh.getLastRow() < 2) return 0;
  var rng = sh.getRange(2, 1, sh.getLastRow() - 1, 6), v = rng.getValues(), n = 0;
  v.forEach(function (r) { var t = splitTags_(r[5]); if (/^video\//.test(r[3]) && t.indexOf(VIDEO_TAG) < 0) { t.push(VIDEO_TAG); r[5] = t.join(TAG_SEP); n++; } });
  if (n) sh.getRange(2, 6, v.length, 1).setValues(v.map(function (r) { return [r[5]]; }));
  return n;
}

function api_photoDedupe(dry) {          // the 15-minute sync holds the same lock while it rewrites the Photos sheet
  var lock = LockService.getScriptLock(); if (!lock.tryLock(30000)) throw new Error('Photos are syncing; try again in a minute');
  try { return photoDedupe_(200000, dry); } finally { lock.releaseLock(); }
}

/* ---------- Adding photos from the app (photo picker or the phone's Share menu) ---------- */
function photoPhoneFolder_() {
  var p = props_(), id = p.getProperty('photoPhoneId');
  if (id) { try { var f0 = DriveApp.getFolderById(id); if (!f0.isTrashed()) return f0; } catch (e) {} }
  var root = photoRoot_(), it = root.getFoldersByName('Phone');
  var f = it.hasNext() ? it.next() : root.createFolder('Phone');
  p.setProperty('photoPhoneId', f.getId()); return f;
}

/**
 * files: [{name, size, type, modified}]. For each: {skip: true} when the same name and size is already in ArNote,
 * otherwise {url} — a Drive upload address the app sends the file to directly (no size limit, nothing passes through here).
 */
function api_photoUploadStart(files, origin) {
  files = (files || []).slice(0, 30);
  var known = {}; photoRows_().forEach(function (p) { known[p.id] = 1; });
  var tok = ScriptApp.getOAuthToken(), folder = photoPhoneFolder_().getId();
  var look = UrlFetchApp.fetchAll(files.map(function (f) {
    var q = "name = '" + String(f.name).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "' and trashed = false";
    return { url: 'https://www.googleapis.com/drive/v3/files?pageSize=50&fields=files(id,size)&q=' + encodeURIComponent(q), headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
  }));
  var need = [], out = files.map(function (f, i) {
    var r = {}; try { r = JSON.parse(look[i].getContentText()); } catch (e) {}
    if ((r.files || []).some(function (x) { return known[x.id] && +x.size === +f.size; })) return { skip: true };
    need.push(i); return null;
  });
  var starts = UrlFetchApp.fetchAll(need.map(function (i) {
    var f = files[i], meta = { name: String(f.name), parents: [folder] };
    if (f.type) meta.mimeType = f.type;
    if (+f.modified) meta.modifiedTime = new Date(+f.modified).toISOString();
    var h = { Authorization: 'Bearer ' + tok, 'X-Upload-Content-Type': f.type || 'application/octet-stream', 'X-Upload-Content-Length': String(+f.size) };
    if (origin) h.Origin = String(origin);
    return { url: 'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true&fields=id', method: 'post',
      contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(meta), headers: h, muteHttpExceptions: true };
  }));
  need.forEach(function (i, j) {
    var hd = starts[j].getAllHeaders ? starts[j].getAllHeaders() : starts[j].getHeaders(), loc = hd.Location || hd.location;
    out[i] = loc ? { url: String(loc) } : { error: 'Drive refused the upload (' + starts[j].getResponseCode() + ')' };
  });
  return out;
}

/** Fallback when the browser can't send to Drive directly: the file comes through here (fine for photos, not big videos). */
function api_photoUploadSmall(name, mime, b64) {
  var f = photoPhoneFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name));
  return f.getId();
}

/** Add just-uploaded files to the Photos index right away. items: [{id, modified}] */
function api_photoUploadDone(items) {
  items = (items || []).filter(function (x) { return x && x.id; });
  if (!items.length) return 0;
  var tok = ScriptApp.getOAuthToken(), tz = Session.getScriptTimeZone(), now = new Date().toISOString();
  var res = UrlFetchApp.fetchAll(items.map(function (x) {
    return { url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(x.id) + '?supportsAllDrives=true&fields=' +
      encodeURIComponent('id,name,mimeType,createdTime,imageMediaMetadata(time,width,height,location),videoMediaMetadata(width,height)'), headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true };
  }));
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = photoSheet_(), known = {}; photoRows_().forEach(function (p) { known[p.id] = 1; });
    var add = [];
    res.forEach(function (r, i) {
      if (r.getResponseCode() !== 200) return;
      var f = JSON.parse(r.getContentText()); if (known[f.id] || !/^(image|video)\//.test(f.mimeType)) return;
      var tk = photoTaken_(f, 'Phone');
      if (tk === String(f.createdTime || '').slice(0, 19) && +items[i].modified)      // no camera date and no date in the name: use the phone's file date
        tk = Utilities.formatDate(new Date(+items[i].modified), tz, "yyyy-MM-dd'T'HH:mm:ss");
      var md = f.imageMediaMetadata || f.videoMediaMetadata || {}, recent = tk && (Date.now() - new Date(tk).getTime()) < 60 * 864e5;
      var gl = (f.imageMediaMetadata || {}).location, hasL = gl && (gl.latitude || gl.longitude);
      add.push([f.id, f.name, tk, f.mimeType, 'Phone', /^video\//.test(f.mimeType) ? VIDEO_TAG : '', '', now, md.width || '', md.height || '', /^image\//.test(f.mimeType) ? 'q' : '',
        hasL ? Math.round(gl.latitude * 1e5) / 1e5 : (/^image\//.test(f.mimeType) ? '-' : ''), hasL ? Math.round(gl.longitude * 1e5) / 1e5 : '', '', '']);   // you picked these: read them whatever their date
      known[f.id] = 1;
    });
    if (add.length) { if (add.some(function (a) { return a[5]; })) ensureTags_([VIDEO_TAG]); sh.getRange(sh.getLastRow() + 1, 1, add.length, PHOTO_COLS.length).setValues(add); }
    return add.length;
  } finally { lock.releaseLock(); }
}

/* ---------------- Photo locations ----------------
 * Each photo keeps its GPS spot (from the photo itself) and a town name looked up once per ~2 km square.
 * Columns: lat, lng ('-' = checked, no location), town, region (US state like "WV", or the country abroad).
 * Your own named places (Script Property 'places': [{name, lat, lng, r (metres)}]) are matched when read, never stored per photo.
 */
function geoKey_(lat, lng) { return (Math.round(lat * 50) / 50).toFixed(2) + ',' + (Math.round(lng * 50) / 50).toFixed(2); }
function geoCache_() {
  var sh = ensureSheet_('GeoCache', ['key', 'town', 'region']), map = {}, add = [];
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) { map[String(r[0])] = { town: String(r[1]), region: String(r[2]) }; });
  return { map: map, add: function (k, c) { map[k] = c; add.push([k, c.town, c.region]); },
    save: function () { if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add); add = []; } };
}
function reverseGeo_(lat, lng) {
  var r = Maps.newGeocoder().reverseGeocode(lat, lng);
  if (r.status === 'OVER_QUERY_LIMIT' || r.status === 'REQUEST_DENIED') throw new Error(r.status);
  var res = r.results || [];
  function find(type, short) { for (var i = 0; i < res.length; i++) { var cs = res[i].address_components || []; for (var j = 0; j < cs.length; j++) if (cs[j].types.indexOf(type) >= 0) return short ? cs[j].short_name : cs[j].long_name; } return ''; }
  var country = find('country', true);
  var town = find('locality') || find('postal_town') || find('administrative_area_level_3') || find('administrative_area_level_2') || '(no town)';
  var region = country === 'US' ? (find('administrative_area_level_1', true) || 'US') : (find('country') || '(at sea)');
  return { town: town, region: region };
}
function placesList_() { try { return JSON.parse(props_().getProperty('places') || '[]'); } catch (e) { return []; } }
function metres_(a, b, c, d) { var R = 6371000, r = Math.PI / 180, x = (d - b) * r * Math.cos((a + c) / 2 * r), y = (c - a) * r; return Math.sqrt(x * x + y * y) * R; }
/** The named place a spot falls in (the closest one whose circle it is inside), or ''. */
function namedAt_(lat, lng, places) {
  var best = '', bd = 1e12;
  (places || []).forEach(function (p) { var m = metres_(lat, lng, +p.lat, +p.lng); if (m <= +p.r && m < bd) { bd = m; best = p.name; } });
  return best;
}
function api_places(list) {
  if (list) {
    var clean = (list || []).map(function (p) { return { name: String(p.name || '').trim().slice(0, 40), lat: Math.round(+p.lat * 1e5) / 1e5, lng: Math.round(+p.lng * 1e5) / 1e5, r: Math.max(50, Math.min(50000, Math.round(+p.r || 1000))) }; })
      .filter(function (p) { return p.name && isFinite(p.lat) && isFinite(p.lng); });
    props_().setProperty('places', JSON.stringify(clean));
  }
  return placesList_();
}
/** Filters shared by the gallery calls: o.region, o.town, o.named, o.nolocation, o.q (search words). */
function placeOk_(p, o, places) {
  if (o.region && p.region !== o.region) return false;
  if (o.town && p.town !== o.town) return false;
  if (o.nolocation && typeof p.lat === 'number') return false;
  if (o.box) { var b = o.box; if (typeof p.lat !== 'number' || p.lat < b[0] || p.lat > b[2] || p.lng < b[1] || p.lng > b[3]) return false; }   // map area [south, west, north, east]
  if (o.named) { if (typeof p.lat !== 'number' || namedAt_(p.lat, p.lng, places) !== o.named) return false; }
  if (o.q) {
    var hay = [p.name, p.tags.join(' '), p.town, p.region, p.folder, typeof p.lat === 'number' ? namedAt_(p.lat, p.lng, places) : ''].join(' ').toLowerCase();
    if (!String(o.q).toLowerCase().split(/\s+/).filter(String).every(function (w) { return hay.indexOf(w) >= 0; })) return false;
  }
  return true;
}
/** Counts for Photos → By place. */
function api_photoPlaces() {
  var rows = photoRows_(), places = placesList_(), regions = {}, named = {}, withLoc = 0, pending = 0, none = 0;
  rows.forEach(function (p) {
    if (p.lat === '') { pending++; return; }
    if (typeof p.lat !== 'number') { none++; return; }
    withLoc++;
    if (p.town) { var g = regions[p.region] = regions[p.region] || {}; g[p.town] = (g[p.town] || 0) + 1; }
    var n = namedAt_(p.lat, p.lng, places); if (n) named[n] = (named[n] || 0) + 1;
  });
  return { regions: regions, named: named, places: places, withLoc: withLoc, none: none, pending: pending };
}
/** Every photo with a location, for the map: [id, lat, lng, 1 if video]. */
function api_photoPoints() {
  var out = [];
  photoRows_().forEach(function (p) { if (typeof p.lat === 'number') out.push([p.id, p.lat, p.lng, /^video\//.test(p.mime) ? 1 : 0]); });
  return out;
}

/* ---------- Video locations: phones write the spot inside the video file (MP4/MOV 'moov' box, e.g. "+39.1234-079.5678/").
 * Only the box headers and the small udta/meta boxes are read with ranged downloads, never the whole video. ---------- */
function rangeBytes_(id, a, b, tok) {
  var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?alt=media&supportsAllDrives=true',
    { headers: { Authorization: 'Bearer ' + tok, Range: 'bytes=' + a + '-' + b }, muteHttpExceptions: true });
  var c = r.getResponseCode(); if (c !== 206 && c !== 200) throw new Error('read ' + c);
  return r.getContent();
}
/** Where a phone video was taken: look for the location tag ("©xyz", or Apple's ISO6709 key) in the last, then first, 256 KB. Two small reads at most. */
function videoLoc_(id, size, tok) {
  var W = 262144, tries = size > W ? [[size - W, size - 1], [0, W - 1]] : [[0, Math.max(0, size - 1)]];
  for (var t = 0; t < tries.length; t++) {
    var b = rangeBytes_(id, tries[t][0], tries[t][1], tok), s = '';
    for (var i = 0; i < b.length; i += 8192) s += String.fromCharCode.apply(null, b.slice(i, i + 8192).map(function (x) { return x & 255; }));
    var k = s.indexOf('©xyz'), seg = k >= 0 ? s.slice(k, k + 60) : '';
    if (!seg) { k = s.indexOf('ISO6709'); seg = k >= 0 ? s.slice(k, k + 400) : ''; }
    var m = seg.match(/([+-]\d{1,2}\.\d{2,})([+-]\d{1,3}\.\d{2,})/);
    if (m && (+m[1] || +m[2]) && Math.abs(+m[1]) <= 90 && Math.abs(+m[2]) <= 180) return { lat: Math.round(+m[1] * 1e5) / 1e5, lng: Math.round(+m[2] * 1e5) / 1e5 };
  }
  return null;
}

/**
 * Background work: find GPS for photos and videos not checked yet, then name the towns. Safe to run repeatedly.
 * Photos: Drive is listed 1,000 files per call with their locations (one outside call per 1,000, not one per photo).
 * Videos: two small reads each. Towns: one lookup per ~2 km square, cached.
 */
function geoWork_(start, budget) {
  start = start || Date.now(); budget = budget || 60000;
  var sh = photoSheet_(), n = sh.getLastRow() - 1; if (n < 1) return { coords: 0, named: 0, left: 0 };
  var C = PHOTO_COLS.indexOf('lat'), v = sh.getRange(2, 1, n, PHOTO_COLS.length).getValues(), tok = ScriptApp.getOAuthToken(), dirty = false, coords = 0, named = 0, pr = props_();
  if (!pr.getProperty('vidGeo1')) {           // once: videos were skipped before; check them now
    v.forEach(function (r) { if (/^video\//.test(r[3]) && r[C] === '-') { r[C] = ''; dirty = true; } });
    pr.setProperty('vidGeo1', '1');
  }
  var want = {}, nWant = 0, vids = [];
  v.forEach(function (r, i) { if (r[C] === '') { if (/^(image|video)\//.test(r[3])) { want[r[0]] = i; nWant++; } else { r[C] = '-'; dirty = true; } } });
  if (nWant) {
    var pt = pr.getProperty('geoPage') || '', fromStart = !pt, q = "(mimeType contains 'image/' or mimeType contains 'video/') and trashed = false";
    do {
      var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?pageSize=1000&supportsAllDrives=true&includeItemsFromAllDrives=true&q=' + encodeURIComponent(q) +
        '&fields=' + encodeURIComponent('nextPageToken,files(id,mimeType,size,imageMediaMetadata(location))') + (pt ? '&pageToken=' + encodeURIComponent(pt) : ''),
        { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) { pt = ''; fromStart = false; break; }
      var js = JSON.parse(res.getContentText());
      (js.files || []).forEach(function (f) {
        var i = want[f.id]; if (i === undefined) return;
        delete want[f.id]; nWant--;
        if (/^video\//.test(f.mimeType)) { vids.push([i, +f.size || 0]); return; }
        var loc = (f.imageMediaMetadata || {}).location;
        if (loc && (loc.latitude || loc.longitude)) { v[i][C] = Math.round(loc.latitude * 1e5) / 1e5; v[i][C + 1] = Math.round(loc.longitude * 1e5) / 1e5; coords++; }
        else v[i][C] = '-';
        dirty = true;
      });
      pt = js.nextPageToken || '';
    } while (pt && nWant > 0 && Date.now() - start < budget * 0.4);
    pr.setProperty('geoPage', pt);                                         // carry on from here next time ('' = from the start)
    if (!pt && fromStart) Object.keys(want).forEach(function (id) {        // a full pass didn't find it (deleted or not readable)
      var i = want[id]; if (!/^video\//.test(v[i][3])) { v[i][C] = '-'; dirty = true; } });
  }
  for (var q2 = 0; q2 < vids.length && Date.now() - start < budget * 0.75; q2++) {
    var vi = vids[q2][0], loc2 = null;
    try { loc2 = vids[q2][1] ? videoLoc_(v[vi][0], vids[q2][1], tok) : null; } catch (e) { if (/invoked too many times/i.test(String(e))) break; continue; }
    if (loc2) { v[vi][C] = loc2.lat; v[vi][C + 1] = loc2.lng; coords++; } else v[vi][C] = '-';
    dirty = true;
  }
  var cache = geoCache_(), miss = {}, stop = false;
  v.forEach(function (r, i) {
    if (typeof r[C] !== 'number' || r[C + 2]) return;
    var key = geoKey_(r[C], r[C + 1]), c = cache.map[key];
    if (c) { r[C + 2] = c.town; r[C + 3] = c.region; dirty = true; named++; } else (miss[key] = miss[key] || []).push(i);
  });
  Object.keys(miss).forEach(function (key) {
    if (stop || Date.now() - start > budget) return;
    var first = v[miss[key][0]], c;
    try { c = reverseGeo_(first[C], first[C + 1]); } catch (e) { stop = true; return; }   // Google's daily limit: carry on tomorrow
    cache.add(key, c);
    miss[key].forEach(function (i) { v[i][C + 2] = c.town; v[i][C + 3] = c.region; named++; });
    dirty = true;
  });
  cache.save();
  if (dirty) sh.getRange(2, C + 1, n, 4).setValues(v.map(function (r) { return [r[C], r[C + 1], r[C + 2], r[C + 3]]; }));
  var left = v.filter(function (r) { return r[C] === ''; }).length + v.filter(function (r) { return typeof r[C] === 'number' && !r[C + 2]; }).length;
  return { coords: coords, named: named, left: left };
}
function api_geoRun() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) throw new Error('ArNote is busy; try again in a minute');
  try { return geoWork_(Date.now(), 200000); } finally { lock.releaseLock(); }
}

function api_photoSyncNow() { var r = syncPhotos_(Date.now() - 180000); return r; }   // ~60 seconds from the app

function photoRows_() {
  var sh = photoSheet_(), v = sh.getDataRange().getValues(); v.shift();
  return v.map(function (r, i) {
    return { id: String(r[0]), name: String(r[1]), taken: r[2] instanceof Date ? r[2].toISOString() : String(r[2]), mime: String(r[3]),
      folder: String(r[4]), tags: splitTags_(r[5]), status: String(r[6]), w: +r[8] || 0, h: +r[9] || 0, scan: String(r[10] || ''),
      lat: typeof r[11] === 'number' ? r[11] : String(r[11] === undefined ? '' : r[11]), lng: typeof r[12] === 'number' ? r[12] : '', town: String(r[13] || ''), region: String(r[14] || ''), row: i + 2 };
  });
}

function tagSet_(tag, exact) {
  if (exact) { var one = {}; one[tag] = 1; return one; }          // just this tag, not its sub-tags
  var kids = {}; readTable_('Tags').forEach(function (t) { var p = String(t.parent || ''); (kids[p] = kids[p] || []).push(String(t.name)); });
  var set = {}; (function walk(x) { set[x] = 1; (kids[x] || []).forEach(walk); })(tag);
  return set;
}

/** One page of photos. o: {view: 'inbox'|'filed'|'all', tag, year, month, folder, offset, limit} */
function api_photos(o) {
  o = o || {};
  var rows = photoRows_(), set = o.tag ? tagSet_(o.tag, o.exact) : null, places = placesList_();
  var list = rows.filter(function (p) {
    if (!placeOk_(p, o, places)) return false;
    if (o.view === 'inbox' && p.status) return false;
    if (o.view === 'filed' && p.status !== 'filed') return false;
    if (o.view === 'untagged' && !untagged_(p)) return false;
    if (set && !p.tags.some(function (g) { return set[g]; })) return false;
    if (o.year && p.taken.slice(0, 4) !== String(o.year)) return false;
    if (o.month && p.taken.slice(5, 7) !== String(o.month)) return false;
    if (o.undated && /^\d{4}-\d\d/.test(p.taken)) return false;
    if (o.folder && p.folder !== o.folder && p.folder.indexOf(o.folder + '/') !== 0) return false;
    return true;
  });
  list.sort(function (a, b) { return a.taken < b.taken ? 1 : a.taken > b.taken ? -1 : 0; });
  var off = +o.offset || 0, lim = Math.min(+o.limit || 120, 1500);
  return { total: list.length, offset: off, items: list.slice(off, off + lim).map(function (p) { delete p.row; return p; }) };
}

/** Counts for the Photos screen: inbox/filed, by year → month, and by tag. */
function api_photoStats() {
  var rows = photoRows_(), years = {}, tags = {}, ftags = {}, folders = {}, inbox = 0, filed = 0;
  rows.forEach(function (p) {
    if (p.status === 'filed') filed++; else if (!p.status) inbox++;
    var y = p.taken.slice(0, 4) || '?', m = p.taken.slice(5, 7) || '?';
    years[y] = years[y] || {}; years[y][m] = (years[y][m] || 0) + 1;
    p.tags.forEach(function (t) { tags[t] = (tags[t] || 0) + 1; if (p.status === 'filed') ftags[t] = (ftags[t] || 0) + 1; });
    folders[p.folder] = (folders[p.folder] || 0) + 1;
  });
  var sync = {}; try { sync = JSON.parse(props_().getProperty('photoSync') || '{}'); } catch (e) {}
  var root = photoRoot_();
  return { total: rows.length, inbox: inbox, filed: filed, untagged: rows.filter(untagged_).length, years: years, tags: tags, ftags: ftags, folders: folders, sync: sync, rootUrl: root.getUrl() };
}

/** Small square thumbnails (cached 6 hours). */
function api_photoThumbs(ids) {
  ids = (ids || []).slice(0, 40);
  var cache = CacheService.getScriptCache(), got = cache.getAll(ids.map(function (id) { return 'pt:' + id; })), out = {}, need = [];
  ids.forEach(function (id) { if (got['pt:' + id]) out[id] = got['pt:' + id]; else need.push(id); });
  if (need.length) {
    var hdr = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    var meta = UrlFetchApp.fetchAll(need.map(function (id) {
      return { url: 'https://www.googleapis.com/drive/v3/files/' + id + '?supportsAllDrives=true&fields=thumbnailLink', headers: hdr, muteHttpExceptions: true };
    }));
    var reqs = [], which = [];
    meta.forEach(function (r, i) {
      if (r.getResponseCode() !== 200) return;
      var l = JSON.parse(r.getContentText()).thumbnailLink; if (!l) return;
      reqs.push({ url: l.replace(/=s\d+(-c)?$/, '') + '=s320', headers: hdr, muteHttpExceptions: true }); which.push(need[i]);
    });
    var imgs = reqs.length ? UrlFetchApp.fetchAll(reqs) : [], put = {};
    imgs.forEach(function (r, i) {
      if (r.getResponseCode() !== 200) return;
      var d = 'data:' + (r.getHeaders()['Content-Type'] || 'image/jpeg') + ';base64,' + Utilities.base64Encode(r.getContent());
      out[which[i]] = d; if (d.length < 95000) put['pt:' + which[i]] = d;
    });
    try { cache.putAll(put, 21600); } catch (e) {}
  }
  return out;
}

/** A big preview for the photo viewer, plus details. */
function api_photoView(id, metaOnly) {
  var f = driveApi_('files/' + id + '?supportsAllDrives=true&fields=' + encodeURIComponent('name,mimeType,size,webViewLink,thumbnailLink,imageMediaMetadata(time,width,height,cameraMake,cameraModel,location),videoMediaMetadata(durationMillis)'));
  var src = '';
  if (!metaOnly) try {
    var hdr = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    if (!f.thumbnailLink && /^image\/(jpeg|png|gif|webp)$/.test(f.mimeType) && +f.size < 2 * 1024 * 1024) {
      var r0 = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '?alt=media&supportsAllDrives=true', { headers: hdr, muteHttpExceptions: true });
      if (r0.getResponseCode() === 200) src = 'data:' + f.mimeType + ';base64,' + Utilities.base64Encode(r0.getContent());
    }
    if (!src && f.thumbnailLink) {
      var r = UrlFetchApp.fetch(f.thumbnailLink.replace(/=s\d+(-c)?$/, '') + '=s1600', { headers: hdr, muteHttpExceptions: true });
      if (r.getResponseCode() === 200) src = 'data:' + (r.getHeaders()['Content-Type'] || 'image/jpeg') + ';base64,' + Utilities.base64Encode(r.getContent());
    }
  } catch (e) {}
  var im = f.imageMediaMetadata || {};
  return { src: src, name: f.name, mime: f.mimeType, url: f.webViewLink, video: /^video\//.test(f.mimeType),
    camera: [im.cameraMake, im.cameraModel].filter(String).join(' '), location: im.location || null };
}

/** Tag / file / unfile photos. ch: {addTags, removeTags, status: 'filed' | ''} */
function api_photoUpdate(ids, ch) {
  ch = ch || {};
  if ((ch.addTags || []).length) ensureTags_(ch.addTags);
  var want = {}; ids.forEach(function (id) { want[id] = 1; });
  var sh = photoSheet_(), v = sh.getDataRange().getValues(), changed = 0;
  for (var i = 1; i < v.length; i++) {
    if (!want[v[i][0]]) continue;
    var tags = splitTags_(v[i][5]);
    (ch.removeTags || []).forEach(function (t) { tags = tags.filter(function (g) { return g !== t; }); });
    (ch.addTags || []).forEach(function (t) { if (tags.indexOf(t) < 0) tags.push(t); });
    v[i][5] = tags.join(TAG_SEP);
    if (ch.status !== undefined) v[i][6] = ch.status;
    changed++;
  }
  if (changed) sh.getRange(1, 6, v.length, 2).setValues(v.map(function (r) { return [r[5], r[6]]; }));
  return changed;
}

/** Tag renamed or deleted: apply the same change to photos. */
function retagPhotos_(fn) {
  var sh = ss_().getSheetByName('Photos'); if (!sh || sh.getLastRow() < 2) return;
  var rng = sh.getRange(2, 6, sh.getLastRow() - 1, 1), v = rng.getValues(), ch = false;
  v.forEach(function (r) { var b = splitTags_(r[0]), a = fn(b.slice()).join(TAG_SEP); if (a !== b.join(TAG_SEP)) { r[0] = a; ch = true; } });
  if (ch) rng.setValues(v);
}

/** Timeline for the Photos grid: how many photos per month (newest first) for the same filters as api_photos. */
function api_photoMonths(o) {
  o = o || {};
  var rows = photoRows_(), set = o.tag ? tagSet_(o.tag, o.exact) : null, m = {}, places = placesList_();
  rows.forEach(function (p) {
    if (!placeOk_(p, o, places)) return;
    if (o.view === 'inbox' && p.status) return;
    if (o.view === 'filed' && p.status !== 'filed') return;
    if (o.view === 'untagged' && !untagged_(p)) return;
    if (set && !p.tags.some(function (g) { return set[g]; })) return;
    if (o.folder && p.folder !== o.folder && p.folder.indexOf(o.folder + '/') !== 0) return;
    var k = /^\d{4}-\d\d/.test(p.taken) ? p.taken.slice(0, 7) : '';
    m[k] = (m[k] || 0) + 1;
  });
  var months = Object.keys(m).sort().reverse().map(function (k) { return [k, m[k]]; });
  return { months: months, total: rows.length };
}

/* ---------------- Receipts, business cards, documents ----------------
 * New photos are read once with Google Drive's free OCR (a temporary Google Doc, deleted right away).
 *  - Receipt: vendor, date, total, tax → a note in the Inbox tagged "Receipt" + a row in the Receipts sheet.
 *  - Business card: person + business → note titled "Name — Business", tagged "Business Card".
 *  - Document: note tagged "Document" (nothing else is kept).
 *  - Anything else stays a normal photo.
 * The note is a linked note pointing at the photo itself, so nothing is copied.
 */
var RECEIPT_COLS = ['date', 'vendor', 'total', 'tax', 'card', 'category', 'check', 'noteId', 'fileId', 'photo', 'added', 'pay', 'checkNo', 'billed', 'kind'];

function receiptSheet_() {
  var sh = ensureSheet_('Receipts', RECEIPT_COLS);
  if (sh.getLastColumn() < RECEIPT_COLS.length) sh.getRange(1, 1, 1, RECEIPT_COLS.length).setValues([RECEIPT_COLS]).setFontWeight('bold');   // newer columns: pay, checkNo
  if (!props_().getProperty('receiptFmt')) { sh.getRange('C:D').setNumberFormat('$#,##0.00'); props_().setProperty('receiptFmt', '1'); }
  return sh;
}

/** Text in a photo, via Drive OCR on a 2000px copy (keeps big phone photos fast and within OCR limits). */
function ocrText_(fileId) {
  var hdr = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var meta = driveApi_('files/' + fileId + '?supportsAllDrives=true&fields=thumbnailLink,mimeType,size');
  var blob = null;
  if (meta.thumbnailLink) {
    var r = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+(-c)?$/, '') + '=s2000', { headers: hdr, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) blob = r.getBlob();
  }
  if (!blob) { if (+meta.size > 4 * 1024 * 1024 || !/^image\//.test(meta.mimeType)) return ''; blob = DriveApp.getFileById(fileId).getBlob(); }
  var b = 'arnote' + Date.now(), head = '--' + b + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify({ name: 'ArNote OCR (temporary)', mimeType: 'application/vnd.google-apps.document' }) +
    '\r\n--' + b + '\r\nContent-Type: ' + (blob.getContentType() || 'image/jpeg') + '\r\n\r\n';
  var payload = Utilities.newBlob(head).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob('\r\n--' + b + '--').getBytes());
  var up = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=en&fields=id',
    { method: 'post', contentType: 'multipart/related; boundary=' + b, payload: payload, headers: hdr, muteHttpExceptions: true });
  if (up.getResponseCode() !== 200) throw new Error('OCR failed (' + up.getResponseCode() + ')');
  var docId = JSON.parse(up.getContentText()).id;
  try {
    var t = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=text/plain', { headers: hdr, muteHttpExceptions: true });
    return t.getResponseCode() === 200 ? t.getContentText('UTF-8').replace(/^﻿/, '') : '';
  } finally {
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId, { method: 'delete', headers: hdr, muteHttpExceptions: true });
  }
}

var MONEY_RE_ = /(?:\$\s?)?(\d{1,3}(?:,\d{3})*|\d+)\.(\d{2})(?!\d)/g;
function moneyIn_(s) { var out = [], m; MONEY_RE_.lastIndex = 0; while ((m = MONEY_RE_.exec(s))) out.push(+(m[1].replace(/,/g, '') + '.' + m[2])); return out; }

/** Decide what a photo is from its text. */
function classifyText_(text) {
  var lines = String(text || '').split(/\r?\n/).map(function (l) { return l.replace(/\s+/g, ' ').trim(); }).filter(String);
  var all = lines.join('\n'), words = (all.match(/[A-Za-z]{2,}/g) || []).length, prices = moneyIn_(all).length;
  var rcptWords = (all.match(/\b(sub\s*-?total|total|tax|change|cash|visa|mastercard|master card|amex|discover|debit|credit|balance due|amount due|tender|qty|item)\b/gi) || []).length;
  var phone = all.match(/\(?\b\d{3}\)?[-.\s]\d{3}[-.\s]\d{4}\b/), email = all.match(/[\w.+-]+@[\w-]+\.[\w.]+/), web = all.match(/\b(?:www\.)?[\w-]+\.(?:com|net|org|biz|co|us|io)\b/i);
  if (/\btotal\b/i.test(all) && prices >= 2 && rcptWords >= 2) return { kind: 'receipt', fields: receiptFields_(lines, all) };
  if (rcptWords >= 4 && prices >= 3) return { kind: 'receipt', fields: receiptFields_(lines, all) };
  if (prices >= 1 && /\b(total|amount)\b/i.test(all) && (/(x{4,}|\*{4,})\s?\d{2,4}\b/i.test(all) || /\b(visa|mastercard|master card|amex|american express|discover|debit)\b/i.test(all)))
    return { kind: 'receipt', fields: receiptFields_(lines, all) };          // a card machine slip: one amount, a card, "Total"/"Amount"
  if (lines.length <= 16 && words <= 70 && (phone || email) && prices <= 1) {     // a card has a phone or email AND a person's name (or both phone and email); a website alone is a sign, not a card
    var cf = cardFields_(lines, email, phone, web);
    if (cf.name || (phone && email)) return { kind: 'card', fields: cf };
  }
  var inSentences = lines.reduce(function (n, l) { var w = (l.match(/[A-Za-z]{2,}/g) || []).length; return n + (w >= 5 ? w : 0); }, 0);   // pages are full lines/paragraphs; screenshots are scattered words
  if (words >= 40 && inSentences >= words * 0.6) return { kind: 'document', fields: {} };
  var tokens = words + (all.match(/\d+/g) || []).length;
  if (lines.length >= 4 && tokens >= 20) return { kind: 'document', fields: {} };   // mostly words and numbers: screenshots, labels, forms, invitations
  if (phone || email || web) return { kind: 'document', fields: {} };              // has an address or number worth keeping (a ticket, a sign with a website)
  return { kind: 'photo', fields: {} };
}

function receiptFields_(lines, all) {
  var skip = /^(welcome|receipt|sales receipt|customer copy|merchant copy|thank|you|store|order|invoice|tel|phone|server|table|guests?|cashier|date|time|expect|visit|www|bank|merchant|term|terminal|sale)\b/i;
  var vendor = '';
  for (var i = 0; i < Math.min(lines.length, 6) && !vendor; i++) {
    var l = lines[i];
    if (skip.test(l) || /^\(?\d{3}\)?[-.\s]\d{3}/.test(l) || /^\d/.test(l) || (l.match(/[A-Za-z]/g) || []).length < 3) continue;
    vendor = l.replace(/[^\w&'.,\- ]/g, '').replace(/\s+#?\d{2,}\b.*$/, '').trim().slice(0, 40);   // "Sheetz 147 5 N 3rd St" → "Sheetz"
  }
  var date = '', m;
  if ((m = all.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/))) {
    var y = +m[3]; if (y < 100) y += 2000; if (+m[1] <= 12 && +m[2] <= 31) date = y + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  } else if ((m = all.match(/\b(20\d\d)[\/\-.](\d{1,2})[\/\-.](\d{1,2})\b/))) {
    date = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  } else if ((m = all.match(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.? (\d{1,2}),? (20\d\d)\b/i))) {
    var mo = 'janfebmaraprmayjunjulaugsepoctnovdec'.indexOf(m[1].toLowerCase().slice(0, 3)) / 3 + 1;
    date = m[3] + '-' + ('0' + mo).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  }
  var isTax = function (low) { return /\btax\b/.test(low) && !/before tax|pre-?tax|tax ?id|taxable|tax ?exempt/.test(low); };
  var isTotal = function (low) { return /(grand total|\btotal\b|amount due|balance due|amount paid)/.test(low) && !/sub\s*-?total|total (tax|savings|discount|items?)|items/.test(low); };
  var total = null, tax = null;
  // labels and amounts on the same line (or the amount on the next line)
  lines.forEach(function (l, i) {
    var low = l.toLowerCase().replace(/\bon \$?\d[\d,]*\.\d{2}/, ''), next = lines[i + 1] || '';   // "TAX 6% on $41.94" — that's the base, not the tax
    var amt = moneyIn_(low); if (!amt.length && (isTax(low) || isTotal(low)) && !/[a-z]{3}/i.test(next)) amt = moneyIn_(next).slice(0, 1);
    if (!amt.length) return;
    if (isTax(low)) tax = amt[amt.length - 1];
    else if (isTotal(low)) { var v = amt[amt.length - 1]; if (total === null || v > total) total = v; }
  });
  // labels in one column and amounts in another (restaurant slips): match them up from the bottom
  if (total === null || tax === null) {
    for (var s = 0; s < lines.length; s++) {
      var labels = [];
      while (s < lines.length && !moneyIn_(lines[s]).length && /(total|tax|due|tip|gratuity|discount|change|cash|balance)/i.test(lines[s])) labels.push(lines[s++].toLowerCase());
      if (labels.length < 2) continue;
      var amts = [];
      for (var j = s; j < lines.length && /^\$?\s?[\d,]+\.\d{2}$/.test(lines[j].replace(/\s/g, '')); j++) amts.push(moneyIn_(lines[j])[0]);
      if (amts.length < labels.length) continue;
      var off = amts.length - labels.length;
      labels.forEach(function (lab, k) { var v = amts[off + k];
        if (tax === null && isTax(lab)) tax = v;
        if (isTotal(lab) && (total === null || v > total)) total = v; });
      break;
    }
  }
  if (total === null) { var allAmt = moneyIn_(all); if (allAmt.length) total = Math.max.apply(null, allAmt); }
  var card = (all.match(/(?:x{2,}|\*+|#{2,}|ending in|acct:?)\s*(\d{4})\b/i) || all.match(/\b(\d{4})\s+(?:visa|mc|mastercard|amex|discover|debit)\b/i) || [])[1] || '';
  var pay = card || /\b(visa|mastercard|master card|amex|american express|discover|debit|credit card)\b/i.test(all) ? 'card' : /\bcash\b/i.test(all) ? 'cash' : '';
  return { vendor: vendor, date: date, total: total, tax: tax, card: card, pay: pay, check: (!vendor || !date || total === null) ? 'check' : '' };
}

function cardFields_(lines, email, phone, web) {
  var biz = /\b(llc|l\.l\.c|inc|co\.?|company|corp|corporation|group|services|service|brewing|brewery|distillery|winery|supply|construction|restaurant|bank|insurance|realty|farms?|equipment|solutions|associates|partners|studio|shop|store|systems|technologies|industries|international|office|department|division|development|university|college|school|county|agency|association|council|foundation|church|hospital|clinic|center|centre|distributing|distributors?|manufacturing|engineering|enterprises|ltd)\b/i;
  var role = /\b(owner|president|manager|director|sales|ceo|cfo|coo|cto|founder|engineer|representative|rep|account|vp|vice|agent|broker|consultant|technician|partner|chef|brewer|officer|inspector|specialist|coordinator|executive|assistant|supervisor|administrator|analyst|general|head|lead|senior|chief)\b/i;
  var stop = /^(cell|mobile|phone|tel|fax|email|e-mail|office|direct|www|web)\b/i;
  var local = email ? email[0].split('@')[0].toLowerCase().replace(/[^a-z]/g, ' ') : '', ltoks = local.split(' ').filter(function (t) { return t.length >= 3; }), lflat = local.replace(/ /g, '');
  function tidy(s) { return s === s.toUpperCase() || s === s.toLowerCase() ? s.toLowerCase().replace(/(^|[\s\-&\/])(\w)/g, function (m, a, c) { return a + c.toUpperCase(); }).replace(/\b(Of|And|The|For)\b/g, function (w) { return w.toLowerCase(); }).replace(/^\w/, function (c) { return c.toUpperCase(); }) : s; }
  // the leading run of name-like words in a line ("DANA WHITE Compliance Officer" → "DANA WHITE")
  function leadName(l) {
    var w = l.replace(/,.*$/, '').split(' '), out = [];
    for (var i = 0; i < w.length && out.length < 4; i++) {
      var x = w[i];
      if (!/^[A-Z][A-Za-z'.\-]*$/.test(x) || role.test(x) || biz.test(x) || stop.test(x)) break;
      if (out.length && (out[0] === out[0].toUpperCase()) !== (x === x.toUpperCase()) && x.length > 2) break;   // "DANA WHITE Compliance…": the name ends where the lettering changes
      out.push(x);
    }
    while (out.length && /^[A-Z]\.?$/.test(out[out.length - 1]) && out.length > 2) out.pop();   // stray trailing initial
    return out.length >= 2 ? out.join(' ') : '';
  }
  var name = '';
  if (lflat.length >= 3) {        // best clue: the line that matches the email (kevin.klaum@ → "Kevin Klaum", shall@ → "Steven Hall")
    lines.some(function (l) {
      var n = leadName(l); if (!n) return false;
      var ws = n.toLowerCase().replace(/[^a-z ]/g, '').split(' ').filter(function (x) { return x.length >= 3; });
      var hit = ws.some(function (x) { return lflat.indexOf(x) >= 0 || ltoks.indexOf(x) >= 0; });
      if (hit) name = n; return hit;
    });
  }
  if (!name) lines.some(function (l) { if (/[@]|www\./i.test(l) || biz.test(l.split(' ').slice(0, 3).join(' '))) return false; var n = leadName(l); if (n && n.split(' ').length <= 4) name = n; return !!name; });
  var business = '';
  var dom = (email && email[0].split('@')[1]) || (web && web[0].replace(/^www\./i, '')) || '';
  var stem = /gmail|yahoo|hotmail|outlook|aol|icloud|comcast|verizon|126\.com|163\.com/i.test(dom) ? '' : dom.split('.')[0].toLowerCase().replace(/[^a-z]/g, '');
  if (stem.length >= 4) {         // a line that spells the email/website domain ("georgfischer" → "Georg Fischer LLC")
    var best = '';
    lines.forEach(function (l) {
      if (/@|www\.|https?:|\d/i.test(l)) return;
      var flat = l.toLowerCase().replace(/[^a-z]/g, '');
      if (flat.length >= 4 && (flat.indexOf(stem) >= 0 || stem.indexOf(flat) >= 0 || flat.indexOf(stem.slice(0, 4)) >= 0) && flat.length > best.replace(/[^a-z]/gi, '').length) best = l;
    });
    business = best;
  }
  if (!business) lines.some(function (l) { if (/@|www\.|\d/.test(l) || (name && l.indexOf(name) >= 0)) return false; if (biz.test(l)) business = l; return !!business; });
  if (!business && stem) business = stem.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  business = tidy(business.replace(/[^\x20-\x7E]+.*$/, '').replace(/\s+/g, ' ').trim());
  if (business.length > 50) business = business.slice(0, 50).replace(/\s+\S*$/, '');
  return { name: tidy(name), business: business, email: email ? email[0] : '', phone: phone ? phone[0] : '' };
}

/** A note in the Inbox that points at the photo (no copy). */
function photoNote_(fileId, title, tags) {
  var now = new Date(), id = Utilities.getUuid();
  ensureTags_(tags);
  var n = { id: id, title: title, notebook: 'Inbox', tags: tags.join(TAG_SEP), created: now, updated: now, docId: fileId, starred: false,
    source: 'link', sourceUrl: 'https://drive.google.com/file/d/' + fileId + '/view', author: Session.getEffectiveUser().getEmail(), attachments: 0 };
  ss_().getSheetByName('Notes').appendRow(noteToRow_(n));
  return id;
}

function money_(v) { return v === null || v === undefined || v === '' ? '' : '$' + (+v).toFixed(2); }

/** Read one photo, make the note / receipt row, and mark the photo. Returns what happened. */
function scanPhoto_(p) {
  var t0 = Date.now(), text = '', r;
  try { text = ocrText_(p.id); } catch (e) { return { id: p.id, kind: 'error', error: String(e.message || e), ms: Date.now() - t0 }; }
  r = classifyText_(text);
  var out = { id: p.id, name: p.name, kind: r.kind, fields: r.fields, ms: Date.now() - t0, chars: text.length };
  if (r.kind === 'receipt') {
    var f = r.fields, d = f.date || String(p.taken).slice(0, 10);
    out.title = 'Receipt — ' + (f.vendor || 'Unknown vendor') + (f.total !== null ? ' — ' + money_(f.total) : '') + (d ? ' — ' + d : '');
    out.noteId = photoNote_(p.id, out.title, ['Receipt']);
    receiptSheet_().appendRow([d, f.vendor, f.total === null ? '' : f.total, f.tax === null ? '' : f.tax, f.card, '', f.check, out.noteId, p.id,
      'https://drive.google.com/file/d/' + p.id + '/view', new Date().toISOString(), f.pay || '', '']);
  } else if (r.kind === 'card') {
    var c = r.fields;
    out.title = [c.name, c.business].filter(String).join(' — ') || 'Business card';
    out.noteId = photoNote_(p.id, out.title, ['Business Card']);
  } else if (r.kind === 'document') {
    out.title = 'Document — ' + String(p.taken).slice(0, 10);
    out.noteId = photoNote_(p.id, out.title, ['Document']);
  }
  return out;
}

function setPhotoScan_(results) {
  var sh = photoSheet_(), v = sh.getDataRange().getValues(), by = {};
  results.forEach(function (r) { by[r.id] = r; });
  var col = PHOTO_COLS.indexOf('scan'), st = PHOTO_COLS.indexOf('status');
  for (var i = 1; i < v.length; i++) {
    var r = by[v[i][0]]; if (!r) continue;
    v[i][col] = r.kind === 'error' ? 'q' : r.kind;
    if (r.noteId) v[i][st] = 'note';                      // now lives in the note Inbox, so it leaves the photo inbox
  }
  sh.getRange(1, st + 1, v.length, 1).setValues(v.map(function (x) { return [x[st]]; }));
  sh.getRange(1, col + 1, v.length, 1).setValues(v.map(function (x) { return [x[col]]; }));
}

/** From the app: scan the selected photos now (a few per call; the app sends the rest in further calls). */
function api_scanPhotos(ids) {
  var start = Date.now(), want = {}, out = [];
  (ids || []).forEach(function (id) { want[id] = 1; });
  var rows = photoRows_().filter(function (p) { return want[p.id]; });
  for (var i = 0; i < rows.length && Date.now() - start < 45000; i++) out.push(scanPhoto_(rows[i]));
  setPhotoScan_(out);
  return { results: out, left: rows.length - out.length };
}

/** Every 15 minutes: scan photos queued by the photo sync (new photos taken in the last 60 days). */
function scanQueue_(budgetMs) {
  var start = Date.now(), q = photoRows_().filter(function (p) { return p.scan === 'q'; }), out = [];
  for (var i = 0; i < q.length && Date.now() - start < budgetMs; i++) out.push(scanPhoto_(q[i]));
  if (out.length) setPhotoScan_(out);
  refreshReceiptCategories_();
  return out.length;
}

/** Receipts sheet "category" = the note's tags (other than Receipt), so tagging the note in the Inbox fills it in. */
function refreshReceiptCategories_() {
  var sh = ss_().getSheetByName('Receipts'); if (!sh || sh.getLastRow() < 2) return;
  var acct = acctTags_(), tags = {}; readTable_('Notes').forEach(function (n) { tags[n.id] = splitTags_(n.tags).filter(function (t) { return t !== 'Receipt' && !acct[t]; }).join(', '); });
  var c = RECEIPT_COLS.indexOf('category'), k = RECEIPT_COLS.indexOf('noteId');
  var rng = sh.getRange(2, 1, sh.getLastRow() - 1, RECEIPT_COLS.length), v = rng.getValues(), out = [], ch = false;
  v.forEach(function (r) { var t = tags[r[k]]; if (t === undefined) t = r[c]; if (t !== r[c]) ch = true; out.push([t]); });
  if (ch) sh.getRange(2, c + 1, out.length, 1).setValues(out);
}

function api_receiptsInfo() {
  var sh = receiptSheet_(), n = Math.max(0, sh.getLastRow() - 1), y = String(new Date().getFullYear()), thisYear = 0, total = 0;
  if (n) sh.getRange(2, 1, n, 3).getValues().forEach(function (r) { var d = r[0] instanceof Date ? r[0].toISOString() : String(r[0]); if (d.slice(0, 4) === y) { thisYear++; total += +r[2] || 0; } });
  return { url: ss_().getUrl() + '#gid=' + sh.getSheetId(), count: n, thisYear: thisYear, totalThisYear: Math.round(total * 100) / 100 };
}

/* ---------- Receipts tab: report columns from your tags, paid-with, fixable details ---------- */
/** Settings: cols = up to 3 report columns [{name, parent}] (any parent tag); cards = {'4013': 'MSBC Visa'}. */
function receiptCfg_() {
  var c = {}; try { c = JSON.parse(props_().getProperty('receiptCfg') || '{}'); } catch (e) {}
  return { cols: (c.cols || []).filter(function (x) { return x && x.name && x.parent; }).slice(0, 3), cards: c.cards || {} };
}
function api_receiptCfg(cfg) {
  if (cfg) {
    var cols = (cfg.cols || []).map(function (x) { return { name: String(x.name || '').trim().slice(0, 30), parent: String(x.parent || '') }; })
      .filter(function (x) { return x.name && x.parent; }).slice(0, 3);
    var cards = {}; Object.keys(cfg.cards || {}).forEach(function (k) { var n = String(cfg.cards[k] || '').trim(); if (/^\d{4}$/.test(k) && n) cards[k] = n.slice(0, 30); });
    props_().setProperty('receiptCfg', JSON.stringify({ cols: cols, cards: cards }));
    try { buildReceiptReport_(true); } catch (e) {}
  }
  return receiptCfg_();
}
function last4_(v) { var s = String(v === null || v === undefined ? '' : v).replace(/\D/g, ''); return s ? ('0000' + s).slice(-4) : ''; }

/* ---------- "Receipt report" tab: one fact per cell, for tax and billing work in the spreadsheet ----------
 * Date | Vendor | <report columns, e.g. Job, Category> | Paid with | Tax | Total | Billed? | Billed on | Other tags | Check | Receipt | Note
 * Rebuilt from ArNote (every 15 minutes and after any receipt change); edits belong in the app. */
function reportSheet_() {
  var ss = ss_(), sh = ss.getSheetByName('Receipt report');
  if (!sh) {
    sh = ss.insertSheet('Receipt report', 0);
    var data = ss.getSheetByName('Receipts');
    if (data) { try { data.protect().setDescription('ArNote data: change receipts in the app').setWarningOnly(true); data.setTabColor('#999999'); } catch (e) {} }
  }
  return sh;
}
function buildReceiptReport_(force) {
  var d = api_receipts(), sh = reportSheet_(), cols = d.cols;
  var head = ['Date', 'Vendor', 'Type'].concat(cols.map(function (c) { return c.name; }), ['Paid with', 'Tax', 'Total', 'Billed?', 'Billed on', 'Other tags', 'Check', 'Receipt', 'Note']);
  var rows = d.rows.slice().sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  var key = JSON.stringify([head, rows.map(function (r) { return [r.date, r.vendor, r.kind, r.cols, r.paid, r.tax, r.total, r.billed, r.other, r.title, r.fileId]; })]);
  var h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key));
  if (!force && props_().getProperty('rcReportHash') === h && sh.getLastRow() === rows.length + 1) return false;   // nothing changed
  var day = function (s) { return s ? new Date(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) : ''; };
  var txt = function (s) { s = String(s || ''); return /^[=+\-@]/.test(s) ? "'" + s : s; };
  var data = rows.map(function (r) {
    var chk = [];
    if (!r.vendor) chk.push('no vendor'); if (!r.date) chk.push('no date'); if (r.total === null) chk.push('no total');
    cols.forEach(function (c, i) { if (r.cols[i].length > 1) chk.push(r.cols[i].length + ' ' + c.name + ' tags: ' + r.cols[i].join(', ')); });
    return [day(r.date), txt(r.vendor), r.kind === 'invoice' ? 'Invoice' : 'Receipt'].concat(r.cols.map(function (x) { return txt(x[0]); }),
      [txt(r.paid), r.tax === null ? '' : r.tax, r.total === null ? '' : r.total, r.billed ? 'Yes' : 'No', day(r.billed), txt(r.other.join(', ')), chk.join('; '),
       r.fileId ? '=HYPERLINK("https://drive.google.com/file/d/' + r.fileId + '/view","Picture")' : '', txt(r.title)]);
  });
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#eeeeee');
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('Built by ArNote. Change receipts, tags and "billed" in the app; this tab is rewritten to match.');
  if (data.length) {
    sh.getRange(2, 1, data.length, head.length).setValues(data);
    var at = function (name) { return head.indexOf(name) + 1; };
    sh.getRange(2, 1, data.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, at('Billed on'), data.length, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(2, at('Tax'), data.length, 2).setNumberFormat('$#,##0.00');
  }
  try { sh.autoResizeColumns(1, head.length); } catch (e) {}
  props_().setProperty('rcReportHash', h);
  return true;
}

/** Every receipt with its tags sorted into the report columns. */
function api_receipts() {
  var cfg = receiptCfg_(), kids = {}, tz = Session.getScriptTimeZone(), sh = receiptSheet_(), acct = acctTags_();
  readTable_('Tags').forEach(function (t) { var p = String(t.parent || ''); (kids[p] = kids[p] || []).push(String(t.name)); });
  var sets = cfg.cols.map(function (c) { var s = {}; (function walk(x) { (kids[x] || []).forEach(function (k) { s[k] = 1; walk(k); }); })(c.parent); return s; });
  var notes = {}; readTable_('Notes').forEach(function (n) { notes[n.id] = n; });
  var n = sh.getLastRow() - 1, rows = [], seenCards = {};
  if (n > 0) sh.getRange(2, 1, n, RECEIPT_COLS.length).getValues().forEach(function (r, i) {
    var o = {}; RECEIPT_COLS.forEach(function (c, j) { o[c] = r[j]; });
    var note = notes[o.noteId]; if (!note) return;                     // note deleted: receipt drops out of the report
    var tags = splitTags_(note.tags), card = last4_(o.card); if (card) seenCards[card] = 1;
    var date = o.date instanceof Date ? Utilities.formatDate(o.date, tz, 'yyyy-MM-dd') : String(o.date || '').slice(0, 10);
    var pay = String(o.pay || '') || (card ? 'card' : '');
    var paid = pay === 'cash' ? 'Cash' : pay === 'check' ? 'Check' + (o.checkNo ? ' #' + o.checkNo : '') : card ? (cfg.cards[card] || '…' + card) : pay === 'card' ? 'Card' : '';
    var total = o.total === '' || o.total === null ? null : +o.total, tax = o.tax === '' || o.tax === null ? null : +o.tax;
    rows.push({ noteId: String(o.noteId), fileId: String(o.fileId), title: String(note.title), date: date, vendor: String(o.vendor || ''), total: total, tax: tax,
      card: card, pay: pay, checkNo: String(o.checkNo || ''), paid: paid,
      cols: sets.map(function (s) { return tags.filter(function (t) { return s[t]; }); }),
      other: tags.filter(function (t) { return t !== 'Receipt' && !acct[t] && !sets.some(function (s) { return s[t]; }); }),
      kind: String(o.kind || 'receipt'),
      billed: o.billed instanceof Date ? Utilities.formatDate(o.billed, tz, 'yyyy-MM-dd') : String(o.billed || '').slice(0, 10),
      check: !o.vendor || !date || total === null });
  });
  return { cols: cfg.cols, cards: cfg.cards, seenCards: Object.keys(seenCards).sort(), rows: rows, url: ss_().getUrl() + '#gid=' + reportSheet_().getSheetId() };
}

/** Not a receipt after all: drop its row from Receipts and take the Receipt tag off the note (the note stays). */
function api_receiptRemove(noteIds) {
  var want = {}; (noteIds || []).forEach(function (id) { want[String(id)] = 1; });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = receiptSheet_(), n = sh.getLastRow() - 1, k = RECEIPT_COLS.indexOf('noteId'), gone = 0;
    if (n > 0) { var ids = sh.getRange(2, k + 1, n, 1).getValues(); for (var i = ids.length - 1; i >= 0; i--) if (want[String(ids[i][0])]) { sh.deleteRow(i + 2); gone++; } }
  } finally { lock.releaseLock(); }
  api_batchUpdate(Object.keys(want), { removeTags: ['Receipt'] });
  try { buildReceiptReport_(); } catch (e) {}
  return gone;
}

/** Mark receipts billed to the customer (date = today) or not billed (on = false). */
function api_receiptBilled(noteIds, on) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = receiptSheet_(), n = sh.getLastRow() - 1; if (n < 1) return 0;
    var k = RECEIPT_COLS.indexOf('noteId'), b = RECEIPT_COLS.indexOf('billed'), want = {}, cnt = 0;
    (noteIds || []).forEach(function (id) { want[String(id)] = 1; });
    var rng = sh.getRange(2, b + 1, n, 1), ids = sh.getRange(2, k + 1, n, 1).getValues(), v = rng.getValues();
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    ids.forEach(function (r, i) { if (want[String(r[0])]) { v[i][0] = on ? (typeof on === 'string' ? on.slice(0, 10) : today) : ''; cnt++; } });
    rng.setNumberFormat('@').setValues(v);
    try { buildReceiptReport_(); } catch (e) {}
    return cnt;
  } finally { lock.releaseLock(); }
}

/** Fix a receipt's details. f: {date, vendor, total, tax, pay: 'card'|'cash'|'check'|'', card, checkNo} */
function api_receiptUpdate(noteId, f) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = receiptSheet_(), n = sh.getLastRow() - 1; if (n < 1) throw new Error('Receipt not found');
    var k = RECEIPT_COLS.indexOf('noteId'), v = sh.getRange(2, 1, n, RECEIPT_COLS.length).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][k]) === String(noteId)) {
      var r = v[i], C = function (c) { return RECEIPT_COLS.indexOf(c); }, num = function (x) { x = String(x === null || x === undefined ? '' : x).replace(/[$,\s]/g, ''); return x === '' || isNaN(+x) ? '' : Math.round(+x * 100) / 100; };
      if (f.date !== undefined) r[C('date')] = String(f.date || '').slice(0, 10);
      if (f.vendor !== undefined) r[C('vendor')] = String(f.vendor || '').slice(0, 60);
      if (f.total !== undefined) r[C('total')] = num(f.total);
      if (f.tax !== undefined) r[C('tax')] = num(f.tax);
      if (f.pay !== undefined) r[C('pay')] = /^(card|cash|check)$/.test(f.pay) ? f.pay : '';
      if (f.card !== undefined) r[C('card')] = last4_(f.card);
      if (f.checkNo !== undefined) r[C('checkNo')] = String(f.checkNo || '').replace(/[^\w-]/g, '').slice(0, 12);
      r[C('check')] = (!r[C('vendor')] || !r[C('date')] || r[C('total')] === '') ? 'check' : '';
      sh.getRange(i + 2, 1, 1, RECEIPT_COLS.length).setValues([r]);
      try { buildReceiptReport_(); } catch (e) {}
      return true;
    }
    throw new Error('Receipt not found');
  } finally { lock.releaseLock(); }
}

/** Test only: read a photo and say what it would be, without changing anything. */
function api_ocrTest(id) { var t0 = Date.now(), text = ocrText_(id), r = classifyText_(text); return { ms: Date.now() - t0, chars: text.length, kind: r.kind, fields: r.fields, sample: text.slice(0, 5000) }; }
