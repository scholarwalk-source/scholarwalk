/**
 * ScholarWalk data receiver — handles three submission types:
 *   1. Full sign-up / application submissions (from signup.html)
 *   2. Eligibility quiz leads (from the homepage quiz widget)
 *   3. Community Q&A questions (from study-lounge.html)
 *
 * Paste this into Extensions → Apps Script in your Google Sheet,
 * then deploy as a Web App (see instructions). All three forms
 * point at the SAME deployed URL — this script routes each
 * submission to the right tab AND emails you a notification.
 *
 * Required tabs in your spreadsheet:
 *   "Signups"   — Timestamp | Full Name | Email | WhatsApp | Citizenship |
 *                 Degree Level | GPA | Target Country | Field of Study |
 *                 Package | Price | Scholarship Interest
 *   "Leads"     — Timestamp | Email | WhatsApp | GPA | Budget |
 *                 Field of Study | Matches Shown
 *   "Questions" — Timestamp | Name | Question
 */

// Change this if you ever want notifications to go elsewhere.
const NOTIFY_EMAIL = 'Info@scholarwalk.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (data.formType === 'question') {
      // Community Q&A submission
      const sheet = ss.getSheetByName('Questions') || ss.getActiveSheet();
      sheet.appendRow([
        new Date(),
        data.name || 'Anonymous',
        data.question || ''
      ]);
      sendNotification(
        'New question on the Study Lounge',
        [
          ['Name', data.name || 'Anonymous'],
          ['Question', data.question || '']
        ]
      );

    } else if (data.packageLabel || data.tierLabel || data.price) {
      // Full sign-up / application submission
      const sheet = ss.getSheetByName('Signups') || ss.getActiveSheet();
      const packageLabel = data.packageLabel || data.tierLabel || data.serviceLabel || '';
      const price = data.price || '';
      sheet.appendRow([
        new Date(),
        data.fullName || '',
        data.email || '',
        data.whatsapp || data.phone || '',
        data.citizenship || data.nationality || '',
        data.degreeLevel || data.educationLevel || '',
        data.gpa || '',
        data.targetCountry || data.destinationCountries || data.country || '',
        data.fieldOfStudy || data.major || '',
        packageLabel,
        price,
        data.scholarshipInterest || data.interest || ''
      ]);
      sendNotification(
        'New sign-up: ' + (data.fullName || 'Unknown') + ' — ' + packageLabel,
        [
          ['Full name', data.fullName || ''],
          ['Email', data.email || ''],
          ['WhatsApp', data.whatsapp || data.phone || ''],
          ['Citizenship', data.citizenship || data.nationality || ''],
          ['Degree level', data.degreeLevel || data.educationLevel || ''],
          ['GPA', data.gpa || ''],
          ['Target country', data.targetCountry || data.destinationCountries || data.country || ''],
          ['Field of study', data.fieldOfStudy || data.major || ''],
          ['Package', packageLabel],
          ['Price', price ? ('$' + price) : ''],
          ['Scholarship interest', data.scholarshipInterest || data.interest || '(none)']
        ]
      );

    } else {
      // Eligibility quiz lead
      const sheet = ss.getSheetByName('Leads') || ss.getActiveSheet();
      sheet.appendRow([
        new Date(),
        data.email || '',
        data.whatsapp || '',
        data.gpa || '',
        data.budget || '',
        data.fieldOfStudy || '',
        data.matchCount || ''
      ]);
      sendNotification(
        'New eligibility quiz lead: ' + (data.email || 'Unknown'),
        [
          ['Email', data.email || ''],
          ['WhatsApp', data.whatsapp || ''],
          ['GPA', data.gpa || ''],
          ['Budget', data.budget || ''],
          ['Field of study', data.fieldOfStudy || ''],
          ['Matches shown', data.matchCount || '']
        ]
      );
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sends a plain-text email to NOTIFY_EMAIL summarizing one
 * submission. `rows` is an array of [label, value] pairs.
 * Failures here are caught so a mail hiccup never blocks the
 * Sheet row from being saved (that part already succeeded above).
 */
function sendNotification(subject, rows) {
  try {
    const body = rows
      .map(function(pair) { return pair[0] + ': ' + pair[1]; })
      .join('\n');
    MailApp.sendEmail(NOTIFY_EMAIL, '[ScholarWalk] ' + subject, body);
  } catch (mailErr) {
    // Don't let an email failure break the submission — the Sheet
    // row above already saved successfully regardless.
    console.error('Notification email failed: ' + mailErr);
  }
}
