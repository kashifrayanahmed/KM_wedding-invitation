// Google Sheet > Extensions > Apps Script. Paste, set KEY, then Deploy > Web app (Execute as Me, Anyone).
const KEY = 'change-this-secret';   // must match KEY in index.html
const WISH_H = ['Time','Status','Name','Relation','Phone','Message','IP','Device ID','Browser'];

function tab(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let t = ss.getSheetByName(name);
  if (!t) { t = ss.insertSheet(name); t.appendRow(head); }
  return t;
}
function clean(s) { return String(s || '').replace(/^[=+\-@]/, "'$&").slice(0, 500); }
function ok() { return ContentService.createTextOutput('ok'); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== KEY) return ok();
    if (d.type === 'wish') {
      const t = tab('Wishes', WISH_H);
      if (t.getLastRow() > 300) return ok();
      t.appendRow([new Date(), 'pending', clean(d.name), clean(d.relation), clean(d.phone),
                   clean(d.message), clean(d.ip), clean(d.did), clean(d.ua)]);
    } else {
      const t = tab('RSVP', ['Time','Name','Attending','Guests','Message']);
      if (t.getLastRow() > 500) return ok();
      t.appendRow([new Date(), clean(d.name), clean(d.attending), clean(d.guests), clean(d.message)]);
    }
  } finally { lock.release(); }
  return ok();
}

// Public wall: ONLY approved wishes, and only name, relation and message.
function doGet() {
  const rows = tab('Wishes', WISH_H).getDataRange().getValues().slice(1)
    .filter(r => String(r[1]).toLowerCase() === 'approved')
    .map(r => ({ name: r[2], relation: r[3], message: r[5] }));
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}
