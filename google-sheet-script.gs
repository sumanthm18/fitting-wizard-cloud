/**
 * Fitting Wizard Cloud waitlist -> Google Sheet
 *
 * Setup (once):
 * 1. Open the "Fitting Wizard Cloud Waitlist" sheet, then Extensions > Apps Script.
 * 2. Delete what is there, paste this whole file, and save.
 * 3. Deploy > New deployment > type "Web app".
 *    Execute as: Me.  Who has access: Anyone.
 * 4. Copy the web app URL (ends in /exec) and paste it into CONFIG.endpoint in the page.
 */

// Column header -> field sent by the page. Order here is the column order in the sheet.
const FIELDS = [
  ['Submitted at', 'submitted_at'],
  ['Name', 'name'],
  ['Role', 'role'],
  ['Clinic', 'clinic'],
  ['State', 'state'],
  ['Email', 'email'],
  ['Phone', 'phone'],
  ['Clinicians', 'clinicians'],
  ['Locations', 'sites'],
  ['Fitting Wizard runs on', 'hosting'],
  ['Wants most', 'wants'],
  ['Free pilot', 'pilot'],
  ['Source', 'source'],
  ['Page', 'page'],
  ['utm_source', 'utm_source'],
  ['utm_medium', 'utm_medium'],
  ['utm_campaign', 'utm_campaign'],
  ['ref', 'ref']
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (data.website) return reply({ ok: true }); // hidden field only bots fill in
    if (!data.email || !data.clinic) return reply({ ok: false, error: 'missing fields' });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(FIELDS.map(f => f[0]));

    sheet.appendRow(FIELDS.map(([, key]) => {
      let v = data[key];
      if (key === 'submitted_at') v = new Date();
      if (Array.isArray(v)) v = v.join(', ');
      if (typeof v === 'boolean') v = v ? 'Yes' : 'No';
      v = v == null ? '' : v;
      // Stop a typed "=..." from running as a formula in the sheet.
      if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;
      return v;
    }));
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply({ ok: true, service: 'Fitting Wizard Cloud waitlist' });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
