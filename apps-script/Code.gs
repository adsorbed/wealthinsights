/**
 * Your Pension Insights — website lead handler
 *
 * Receives enquiries from the modal on pensioninsights.co.uk, appends each one
 * to this spreadsheet, and forwards it to the GoHighLevel sub-account.
 *
 * SETUP
 *  1. Create a Google Sheet called "Pension Insights - Website Leads".
 *  2. Extensions > Apps Script, delete anything there, paste this file in.
 *  3. Put your GoHighLevel inbound webhook URL in GHL_WEBHOOK_URL below.
 *     (In GHL: Automation > Workflows > new workflow > trigger "Inbound Webhook",
 *      copy the URL it gives you. Leave the string empty to skip GHL for now.)
 *  4. Put the address you want new-lead alerts sent to in NOTIFY_EMAIL.
 *  5. Deploy > New deployment > type "Web app".
 *       Execute as:        Me
 *       Who has access:    Anyone
 *     Copy the /exec URL it gives you and send it to Claude.
 *
 * Re-deploying after an edit: Deploy > Manage deployments > pencil icon >
 * Version "New version" > Deploy. The URL stays the same.
 */

var GHL_WEBHOOK_URL = '';
var NOTIFY_EMAIL = 'info@pensioninsights.co.uk';
var SHEET_NAME = 'Leads';

var HEADERS = [
  'Timestamp',
  'First name',
  'Last name',
  'Email',
  'Phone',
  'Investable assets',
  'Requested',
  'Page',
  'Sent to CRM'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Reject anything missing the fields the form marks as required, so junk
    // does not silently fill the sheet.
    if (!data.email || !data.firstName) {
      return json({ ok: false, error: 'Missing required fields' });
    }

    var crmStatus = sendToGhl(data);
    appendRow(data, crmStatus);
    notify(data);

    return json({ ok: true });
  } catch (err) {
    // Log so failures are visible under Executions rather than lost.
    console.error('Lead handler failed: ' + err);
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, message: 'Your Pension Insights lead handler is running.' });
}

function appendRow(data, crmStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    data.firstName || '',
    data.lastName || '',
    data.email || '',
    data.phone || '',
    data.assets || '',
    data.resource || '',
    data.page || '',
    crmStatus
  ]);
}

function sendToGhl(data) {
  if (!GHL_WEBHOOK_URL) return 'not configured';

  var payload = {
    firstName: data.firstName || '',
    lastName: data.lastName || '',
    email: data.email || '',
    phone: data.phone || '',
    source: 'Pension Insights website',
    investableAssets: data.assets || '',
    requestedResource: data.resource || '',
    landingPage: data.page || ''
  };

  try {
    var res = UrlFetchApp.fetch(GHL_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    return (code >= 200 && code < 300) ? 'yes' : 'failed (' + code + ')';
  } catch (err) {
    // Never let a CRM outage lose the lead. The sheet row still gets written.
    console.error('GHL forward failed: ' + err);
    return 'failed';
  }
}

function notify(data) {
  if (!NOTIFY_EMAIL) return;

  var subject = 'New website lead: ' + (data.firstName || '') + ' ' + (data.lastName || '');
  var body = [
    'A new enquiry came in from pensioninsights.co.uk.',
    '',
    'Name:               ' + (data.firstName || '') + ' ' + (data.lastName || ''),
    'Email:              ' + (data.email || ''),
    'Phone:              ' + (data.phone || ''),
    'Investable assets:  ' + (data.assets || ''),
    'Requested:          ' + (data.resource || ''),
    'Page:               ' + (data.page || '')
  ].join('\n');

  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) {
    console.error('Notification email failed: ' + err);
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
