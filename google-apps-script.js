// Google Apps Script backend for the beat store
// 1) Create a new Apps Script project at https://script.google.com
// 2) Paste this code into the editor
// 3) Deploy as a Web App: Execute as me, Who has access: Anyone
// 4) Copy the Web App URL into the site config as google_apps_script_url
// 5) Replace PASTE_YOUR_SPREADSHEET_ID_HERE with your Google Sheet ID

function doGet(e) {
  return HtmlService.createHtmlOutput('Beat Store Apps Script is running.');
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function hashPassword(password) {
  const value = password || '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
  return Utilities.base64Encode(digest);
}

function getAccountSheet(ss) {
  const sheet = getOrCreateSheet(ss, 'Accounts');
  ensureHeaders(sheet, ['timestamp', 'type', 'name', 'email', 'password', 'status']);
  return sheet;
}

function findAccountByEmail(sheet, email) {
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return null;
  }

  const normalizedEmail = (email || '').toString().trim().toLowerCase();
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const rowEmail = (row[3] || '').toString().trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      return {
        rowIndex: i + 1,
        name: row[2] || '',
        email: row[3] || '',
        password: row[4] || '',
        status: row[5] || ''
      };
    }
  }

  return null;
}

function updateAccountPassword(sheet, email, newPassword) {
  const account = findAccountByEmail(sheet, email);
  if (!account || !account.rowIndex) {
    return false;
  }

  sheet.getRange(account.rowIndex, 5).setValue(hashPassword(newPassword));
  return true;
}

function sendNotificationEmail(options) {
  if (!options || !options.to) {
    return { ok: false, reason: 'missing-recipient' };
  }

  MailApp.sendEmail({
    to: options.to,
    replyTo: options.replyTo || 'omoluwajames2024@gmail.com',
    subject: options.subject,
    htmlBody: options.htmlBody
  });

  return { ok: true, to: options.to };
}

function doPost(e) {
  const spreadsheetId = '11tNO0MmDQPQvATk8KcifXfkKUjLKcIjW1o2BKQc-bwk';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const params = e.parameter || {};
  const type = (params.type || 'payment').toString().toLowerCase();
  const timestamp = params.timestamp || new Date().toISOString();

  let sheetName = 'Store Responses';
  let headers = [
    'timestamp',
    'type',
    'name',
    'email',
    'message',
    'beatTitle',
    'beatGenre',
    'beatBpm',
    'beatKey',
    'offerPrice',
    'offerMessage',
    'paymentReference',
    'amount',
    'currency',
    'status',
    'orderItems',
    'orderSummary',
    'rawResponse',
    'phone'
  ];

  if (type === 'contact') {
    sheetName = 'Contacts';
    headers = ['timestamp', 'type', 'name', 'email', 'message', 'phone'];
  } else if (type === 'exclusive_offer') {
    sheetName = 'Offers';
    headers = ['timestamp', 'type', 'name', 'email', 'beatTitle', 'beatGenre', 'beatBpm', 'beatKey', 'offerPrice', 'offerMessage'];
  } else {
    sheetName = 'Payments';
    headers = ['timestamp', 'type', 'name', 'email', 'paymentReference', 'amount', 'currency', 'status', 'orderItems', 'orderSummary', 'downloadLinks', 'rawResponse'];
  }

  if (type === 'account_create' || type === 'account_signin' || type === 'account_forgot_password') {
    const accountSheet = getAccountSheet(ss);
    const email = (params.email || '').toString().trim().toLowerCase();
    const password = (params.password || '').toString();
    const name = (params.name || '').toString().trim();

    if (type === 'account_create') {
      if (!name || !email || !password) {
        return ContentService.createTextOutput(JSON.stringify({
          ok: false,
          message: 'Please provide your name, email, and password.'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const existingAccount = findAccountByEmail(accountSheet, email);
      if (existingAccount) {
        return ContentService.createTextOutput(JSON.stringify({
          ok: false,
          message: 'An account with this email already exists.'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      accountSheet.appendRow([timestamp, type, name, email, hashPassword(password), 'active']);
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        message: 'Account created successfully.',
        user: { name: name, email: email }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (type === 'account_forgot_password') {
      if (!email || !password) {
        return ContentService.createTextOutput(JSON.stringify({
          ok: false,
          message: 'Please provide your email and new password.'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const updated = updateAccountPassword(accountSheet, email, password);
      if (!updated) {
        return ContentService.createTextOutput(JSON.stringify({
          ok: false,
          message: 'No account found with that email.'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        message: 'Password reset successfully.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (!email || !password) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false,
        message: 'Please provide your email and password.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const account = findAccountByEmail(accountSheet, email);
    if (!account || account.password !== hashPassword(password)) {
      return ContentService.createTextOutput(JSON.stringify({
        ok: false,
        message: 'Invalid email or password.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      message: 'Signed in successfully.',
      user: { name: account.name, email: account.email }
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = getOrCreateSheet(ss, sheetName);
  ensureHeaders(sheet, headers);

  const data = {
    timestamp: timestamp,
    type: type,
    name: params.name || '',
    email: params.email || '',
    message: params.message || '',
    beatTitle: params.beatTitle || '',
    beatGenre: params.beatGenre || '',
    beatBpm: params.beatBpm || '',
    beatKey: params.beatKey || '',
    offerPrice: params.offerPrice || '',
    offerMessage: params.offerMessage || '',
    paymentReference: params.paymentReference || '',
    amount: params.amount || '',
    currency: params.currency || '',
    status: params.status || '',
    orderItems: params.orderItems || '',
    orderSummary: params.orderSummary || '',
    downloadLinks: params.downloadLinks || '',
    rawResponse: params.rawResponse || '',
    phone: params.phone || ''
  };

  const row = headers.map(function (header) {
    return data[header] || '';
  });

  sheet.appendRow(row);

  const customerEmail = (params.customerEmail || params.email || '').toString().trim();
  const adminEmail = (params.adminEmail || params.admin_email || '').toString().trim();
  const customerName = (params.name || 'Customer').toString().trim() || 'Customer';
  const customerMessage = (params.message || '').toString();
  const notificationResults = [];

  if (type === 'contact' && customerEmail) {
    const customerBody = [
      `<p>Hi ${customerName},</p>`,
      '<p>Thanks for contacting De Beat Chef. We have received your message and will get back to you shortly.</p>',
      '<p><strong>Your message:</strong></p>',
      `<p>${customerMessage.replace(/\n/g, '<br>') || 'No message provided.'}</p>`,
      '<p>Best regards,<br>De Beat Chef</p>'
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: customerEmail,
      replyTo: adminEmail || customerEmail,
      subject: 'We received your message',
      htmlBody: customerBody
    }));
  }

  if (type === 'contact' && adminEmail) {
    const adminBody = [
      `<p>You have a new contact message from ${customerName}.</p>`,
      `<p><strong>Email:</strong> ${customerEmail || 'Not provided'}</p>`,
      '<p><strong>Message:</strong></p>',
      `<p>${customerMessage.replace(/\n/g, '<br>') || 'No message provided.'}</p>`,
      '<p>Please reply to the customer directly.</p>'
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: adminEmail,
      replyTo: customerEmail || '',
      subject: `New contact message from ${customerName}`,
      htmlBody: adminBody
    }));
  }

  if (type === 'payment' && customerEmail) {
    let downloadHtml = '';
    try {
      const parsedLinks = JSON.parse(params.downloadLinks || '[]');
      if (Array.isArray(parsedLinks) && parsedLinks.length) {
        downloadHtml = parsedLinks.map(function(item) {
          const links = item.downloadLinks || {};
          const allowed = Array.isArray(item.allowedFiles) && item.allowedFiles.length ? item.allowedFiles : Object.keys(links);
          const files = allowed.map(function(key) {
            if (!links[key]) return null;
            return `<li><strong>${String(key).toUpperCase()}:</strong> <a href="${links[key]}" target="_blank">Download</a></li>`;
          }).filter(Boolean).join('');
          return `<div style="margin-bottom:16px;"><p><strong>${item.beat}</strong> (${item.license})</p><ul style="margin:0;padding-left:18px;">${files}</ul></div>`;
        }).join('');
      }
    } catch (e) {
      downloadHtml = `<p>Unable to parse download links automatically. Please contact support with your payment reference.</p>`;
    }

    const customerBody = [
      `<p>Hi ${customerName},</p>`,
      '<p>Thank you for your purchase from De Beat Chef!</p>',
      `<p><strong>Order reference:</strong> ${params.paymentReference || 'N/A'}</p>`,
      `<p><strong>Total paid:</strong> ${params.amount ? '₦' + params.amount : 'N/A'}</p>`,
      '<p>Here are your download links:</p>',
      `${downloadHtml || '<p>No valid download links were included in the order.</p>'}`,
      '<p>If you have any trouble downloading the files, reply to this email and we will help you.</p>',
      '<p>Best regards,<br>De Beat Chef</p>'
    ].join('');

    notificationResults.push(sendNotificationEmail({
      to: customerEmail,
      replyTo: adminEmail || customerEmail,
      subject: 'Your Beat Store download links',
      htmlBody: customerBody
    }));

    if (adminEmail) {
      const adminPaymentBody = [
        `<p>New purchase completed by ${customerName} (${customerEmail}).</p>`,
        `<p><strong>Reference:</strong> ${params.paymentReference || 'N/A'}</p>`,
        `<p><strong>Amount:</strong> ₦${params.amount || 'N/A'}</p>`,
        `<p><strong>Purchased items:</strong></p>`,
        `<pre style="white-space:pre-wrap;">${params.orderItems || 'N/A'}</pre>`
      ].join('');
      notificationResults.push(sendNotificationEmail({
        to: adminEmail,
        replyTo: customerEmail || '',
        subject: `New beat purchase completed by ${customerName}`,
        htmlBody: adminPaymentBody
      }));
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: 'Saved to sheet and sent notifications',
    type: type,
    sheetName: sheetName,
    notifications: notificationResults
  })).setMimeType(ContentService.MimeType.JSON);
}
