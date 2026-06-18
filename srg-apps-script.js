// ─── CONFIG ────────────────────────────────────────────────────────────────
const ADMIN_KEY  = 'srg-admin';   // Must match ADMIN_PASSWORD in admin.html
const RULES_URL  = 'https://stevencase243.github.io/SRG-/srg-rules.html';
const FROM_NAME  = 'SRG League';  // Sender name shown in email
// ───────────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'addDriver') {
      return handleAddDriver(data);
    } else {
      return handleAcknowledgment(data);
    }

  } catch (err) {
    return respond({ result: 'error', message: err.message });
  }
}


// ── Add driver: validate admin key, add to Approved sheet, send email ──────
function handleAddDriver(data) {
  if (data.adminKey !== ADMIN_KEY) {
    return respond({ result: 'unauthorized' });
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var approved = getOrCreateSheet(ss, 'Approved');

  // Add headers if sheet is empty
  if (approved.getLastRow() === 0) {
    approved.appendRow(['Name', 'Email', 'iRacing ID', 'Invited On']);
  }

  approved.appendRow([
    data.name  || '',
    data.email || '',
    data.irid  || '',
    new Date().toISOString()
  ]);

  // Build their personalized link
  var link = RULES_URL
    + '?name='  + encodeURIComponent(data.name  || '')
    + '&email=' + encodeURIComponent(data.email || '')
    + '&irid='  + encodeURIComponent(data.irid  || '');

  // Send the email
  sendInviteEmail(data.name, data.email, link);

  return respond({ result: 'success' });
}


// ── Acknowledgment submission: validate email, log to Sheet1 ───────────────
function handleAcknowledgment(data) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var approved = getOrCreateSheet(ss, 'Approved');
  var sheet    = getOrCreateSheet(ss, 'Sheet1');

  // Check submitted email is in Approved list (column B)
  var emails = approved.getRange('B:B').getValues().flat().map(function(e) {
    return String(e).toLowerCase().trim();
  });

  if (!emails.includes(String(data.email).toLowerCase().trim())) {
    return respond({ result: 'unauthorized' });
  }

  // Add headers to Sheet1 if empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Date', 'Name', 'Email', 'iRacing ID', 'Acknowledged']);
  }

  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.name       || '',
    data.email      || '',
    data.irid       || '',
    data.acknowledged ? 'Yes' : 'No'
  ]);

  return respond({ result: 'success' });
}


// ── Email template ──────────────────────────────────────────────────────────
function sendInviteEmail(name, email, link) {
  var firstName = name ? name.split(' ')[0] : 'Driver';

  var html = `
    <div style="background:#0c0c0e;padding:40px 20px;font-family:'Segoe UI',sans-serif;">
      <div style="max-width:520px;margin:0 auto;">

        <div style="background:#e02020;color:#fff;font-size:11px;font-weight:700;
                    letter-spacing:2px;text-transform:uppercase;padding:4px 12px;
                    border-radius:2px;display:inline-block;margin-bottom:20px;">
          SRG League
        </div>

        <h1 style="color:#ffffff;font-size:26px;font-weight:800;margin:0 0 8px;">
          Hey ${firstName}, welcome to SRG!
        </h1>

        <p style="color:#888898;font-size:14px;margin:0 0 28px;line-height:1.6;">
          Before you hit the track, you need to read and acknowledge the SRG league rules.
          This keeps everyone on the same page and the racing clean.
        </p>

        <div style="background:#141418;border:1px solid #2a2a34;border-radius:8px;
                    padding:24px;margin-bottom:28px;">
          <p style="color:#c8c8d8;font-size:13px;margin:0 0 16px;line-height:1.6;">
            Click the button below to open your personalized rules page.
            Read through everything, check the box at the bottom, and hit submit.
            It only takes a few minutes.
          </p>
          <a href="${link}"
             style="display:block;background:#e02020;color:#fff;text-align:center;
                    padding:13px 24px;border-radius:6px;font-size:13px;font-weight:700;
                    letter-spacing:1px;text-transform:uppercase;text-decoration:none;">
            📋 Read &amp; Acknowledge Rules
          </a>
        </div>

        <p style="color:#555560;font-size:12px;line-height:1.6;margin:0;">
          If the button doesn't work, copy and paste this link into your browser:<br/>
          <a href="${link}" style="color:#888898;word-break:break-all;">${link}</a>
        </p>

        <hr style="border:none;border-top:1px solid #2a2a34;margin:24px 0;" />
        <p style="color:#333340;font-size:11px;margin:0;">
          SRG League &nbsp;·&nbsp; This link was generated specifically for ${email}
        </p>

      </div>
    </div>
  `;

  MailApp.sendEmail({
    to:       email,
    subject:  'SRG League Rules — Action Required',
    htmlBody: html,
    name:     FROM_NAME
  });
}


// ── Helpers ─────────────────────────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
