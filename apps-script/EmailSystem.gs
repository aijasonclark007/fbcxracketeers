/**
 * Emails a player when Status (column F) is set to
 * "Approved", "Declined", "Move Up", "Remind" or "Waitlist".
 * (New registrations to a full level are set to "Waitlist" automatically by the registration script.)
 * Setup: run setupStatusColumn() once, then Triggers > Add Trigger:
 *   Function: onStatusEdit | From spreadsheet | On edit
 */
const STATUS_COL = 6;   // F
const EMAILED_COL = 7;  // G
const PROOF_COL = 9;    // I (Payment Proof link, filled when a player uploads a screenshot)

const PAYMENT_PAGE_URL = 'https://aijasonclark007.github.io/fbcxracketeers/payment.html';
const SITE_URL = 'https://aijasonclark007.github.io/fbcxracketeers/';
const DEADLINE = 'October 30, 2026';

function setupStatusColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sheet.getRange(1, STATUS_COL, 1, 5)
    .setValues([['Status', 'Emailed', 'Token', 'Payment Proof', 'Paid At']]).setFontWeight('bold');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Declined', 'Move Up', 'Remind', 'Waitlist'], true)
    .setAllowInvalid(false).build();
  sheet.getRange(2, STATUS_COL, 999, 1).setDataValidation(rule);
}

function onStatusEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (range.getColumn() > STATUS_COL || range.getLastColumn() < STATUS_COL) return;

  let skipped = 0;
  for (let row = range.getRow(); row <= range.getLastRow(); row++) {
    if (row === 1) continue;
    const [, name, email, , skill, status, emailed, token, proof] =
      sheet.getRange(row, 1, 1, PROOF_COL).getValues()[0];
    if (!email) continue;

    const mark = String(emailed);
    const first = String(name).split(' ')[0];
    const retry = mark === '' || mark.startsWith('Failed') || mark.startsWith('Not sent');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

    try {
      if (status === 'Approved' && (retry || mark.startsWith('Declined') || mark.startsWith('Move Up') || mark.startsWith('Waitlisted'))) {
        const fromWaitlist = mark.startsWith('Waitlisted');
        const newToken = Utilities.getUuid();
        sheet.getRange(row, TOKEN_COL).setValue(newToken); // links the upload page to this player
        const link = PAYMENT_PAGE_URL + '?t=' + newToken;
        const intro = fromWaitlist
          ? 'Good news! A slot opened up in the ' + skill + ' level, and your waitlisted entry for the FBC x Racketeers Badminton Tournament has been approved.'
          : 'Good news! Your entry (' + skill + ' level) for the FBC x Racketeers Badminton Tournament has been approved.';
        const howTo = 'To secure your slot, pay using the QR code on the payment page, then upload a screenshot of your payment. ' +
          'Please pay on or before ' + DEADLINE + ', because slots without payment by then may be given to another player.';
        MailApp.sendEmail({
          to: email,
          name: 'FBC x Racketeers',
          subject: "You're approved: FBC x Racketeers Badminton Tournament",
          body:
            'Hi ' + first + ',\n\n' + intro + '\n\n' +
            'Date: November 8, 2026\n' +
            'Venue: IBT Sports Center\n' +
            'Games: 2PM - 9PM (please be on time)\n' +
            'Entry fee: PHP 500 per player\n' +
            'Payment deadline: ' + DEADLINE + '\n' +
            'Food & drinks are included.\n\n' +
            howTo + '\n\nPay and upload your payment: ' + link + '\n\n' +
            'See you on the court!\nFBC x Racketeers',
          htmlBody:
            '<p>Hi ' + esc_(first) + ',</p>' +
            '<p>' + esc_(intro) + '</p>' +
            '<p><b>Date:</b> November 8, 2026<br><b>Venue:</b> IBT Sports Center<br>' +
            '<b>Games:</b> 2PM - 9PM (please be on time)<br><b>Entry fee:</b> PHP 500 per player<br>' +
            '<b>Payment deadline:</b> ' + DEADLINE + '<br>Food &amp; drinks are included.</p>' +
            '<p>' + esc_(howTo) + '</p>' + button_(link, 'Click to pay & upload payment') +
            '<p>See you on the court!<br>FBC x Racketeers</p>'
        });
        sheet.getRange(row, EMAILED_COL).setValue(new Date());

      } else if (status === 'Declined' && !mark.startsWith('Declined')) {
        MailApp.sendEmail({
          to: email,
          name: 'FBC x Racketeers',
          subject: 'Update on your entry: FBC x Racketeers Badminton Tournament',
          body:
            'Hi ' + first + ',\n\n' +
            'Thank you for registering for the FBC x Racketeers Badminton Tournament.\n\n' +
            "We're sorry, but we're unable to approve your entry this time. " +
            'Entries are approved for players with several recorded attendances at our regular schedule games.\n\n' +
            'We hope to see you at our regular games and at our future events.\n\n' +
            'FBC x Racketeers'
        });
        sheet.getRange(row, EMAILED_COL).setValue('Declined ' + stamp);

      } else if (status === 'Move Up' && !mark.startsWith('Move Up')) {
        // Low players are invited to Mid or High, Mid players to High
        const higher = skill === 'Low' ? 'the Mid or High level' : skill === 'Mid' ? 'the High level' : '';
        if (!higher) {
          sheet.getRange(row, EMAILED_COL).setValue('Not sent: High is already the top level');
          continue;
        }
        MailApp.sendEmail({
          to: email,
          name: 'FBC x Racketeers',
          subject: 'Please register at a higher level: FBC x Racketeers Badminton Tournament',
          body:
            'Hi ' + first + ',\n\n' +
            'Thank you for registering for the FBC x Racketeers Badminton Tournament in the ' + skill + ' level.\n\n' +
            'We believe your skills are above the ' + skill + ' level, so we are unable to approve your entry there. ' +
            'To keep every match fair for all players, we invite you to register again and choose ' + higher + ':\n' +
            SITE_URL + '\n\n' +
            'Thank you for understanding. We hope to see you on the court!\n\n' +
            'FBC x Racketeers'
        });
        sheet.getRange(row, EMAILED_COL).setValue('Move Up ' + stamp);

      } else if (status === 'Waitlist' && !mark.startsWith('Waitlisted')) {
        sendWaitlistEmail_(email, first, skill);
        sheet.getRange(row, EMAILED_COL).setValue('Waitlisted ' + stamp);

      } else if (status === 'Remind') {
        // Only for approved players (they have a token) who have not uploaded proof of payment yet
        if (!token || proof) { skipped++; continue; }
        const link = PAYMENT_PAGE_URL + '?t=' + token;
        const msg = 'This is a friendly reminder that your entry (' + skill + ' level) is approved, but we have not received your payment yet. ' +
          'Please pay on or before ' + DEADLINE + ', because slots without payment by then may be given to another player.';
        MailApp.sendEmail({
          to: email,
          name: 'FBC x Racketeers',
          subject: 'Payment reminder: FBC x Racketeers Badminton Tournament',
          body:
            'Hi ' + first + ',\n\n' + msg + '\n\n' +
            'Entry fee: PHP 500 per player\n' +
            'Payment deadline: ' + DEADLINE + '\n\n' +
            'Pay and upload your payment: ' + link + '\n\n' +
            'If you have already paid, please upload your screenshot using the same link.\n\n' +
            'FBC x Racketeers',
          htmlBody:
            '<p>Hi ' + esc_(first) + ',</p>' +
            '<p>' + esc_(msg) + '</p>' +
            '<p><b>Entry fee:</b> PHP 500 per player<br><b>Payment deadline:</b> ' + DEADLINE + '</p>' +
            button_(link, 'Click to pay & upload payment') +
            '<p>If you have already paid, please upload your screenshot using the same button.</p>' +
            '<p>FBC x Racketeers</p>'
        });
        sheet.getRange(row, EMAILED_COL).setValue('Reminded ' + stamp);
      }
    } catch (err) {
      sheet.getRange(row, EMAILED_COL).setValue('Failed: ' + err.message);
    }
  }
  if (skipped) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      skipped + ' row(s) skipped: not approved yet, or proof of payment already uploaded.', 'Reminder', 8);
  }
}

// "You're on the waitlist" email (sent automatically at registration, or when you set Status to Waitlist)
function sendWaitlistEmail_(email, first, skill) {
  if (!email) return;
  const msg = 'Thank you for registering for the FBC x Racketeers Badminton Tournament in the ' + skill + ' level. ' +
    'All ' + skill + ' slots are taken right now, so you have been added to the waitlist.';
  MailApp.sendEmail({
    to: email,
    name: 'FBC x Racketeers',
    subject: "You're on the waitlist: FBC x Racketeers Badminton Tournament",
    body:
      'Hi ' + first + ',\n\n' + msg + '\n\n' +
      "If a slot opens up, we'll email you right away. You don't need to pay anything now.\n\n" +
      'FBC x Racketeers',
    htmlBody:
      '<p>Hi ' + esc_(first) + ',</p><p>' + esc_(msg) + '</p>' +
      "<p>If a slot opens up, we'll email you right away. You don't need to pay anything now.</p><p>FBC x Racketeers</p>"
  });
}

function button_(link, label) {
  return '<p><a href="' + link + '" style="display:inline-block;background:#8DC63F;color:#0B1B33;font-weight:bold;' +
    'text-decoration:none;padding:12px 22px;border-radius:999px;">' + esc_(label) + '</a></p>';
}

// Escapes text from the registration form before putting it in the HTML email
function esc_(t) {
  return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}