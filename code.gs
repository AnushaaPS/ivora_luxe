const SHEET_NAME = 'Products';
const ADMIN_EMAIL = 'ivoraluxen@gmail.com';

const REQUIRED_COLUMNS = [
  'id',
  'name',
  'price',
  'category',
  'description',
  'image',
  'badge'
];

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// MENU
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('🌟 Ivora Luxe')
    .addItem('✓ Validate Sheet Structure', 'validateSheetStructure')
    .addItem('📊 View Statistics', 'showStats')
    .addItem('✉️ Test Email Settings', 'testEmail')
    .addItem('🔄 Clear Old Products', 'clearSheet')
    .addSeparator()
    .addItem('📋 Sheet Instructions', 'showInstructions')
    .addToMenu();
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// GET API
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function doGet(e) {

  const action = e.parameter.action || 'test';

  if (action === 'products') {
    return getProductsJSON();
  }

  return ContentService
    .createTextOutput('Ivora Luxe Backend Active')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// POST API
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function doPost(e) {

  try {

    let data = {};

    // JSON requests
    if (e.postData && e.postData.contents) {

      try {
        data = JSON.parse(e.postData.contents);
      } catch(err) {
        data = e.parameter;
      }

    } else {
      data = e.parameter;
    }

    const action = data.action || '';
    const type = data.type || '';

    // ───── PRODUCT ACTIONS ─────

    if (action === 'addProduct') {
      return addProduct(data);
    }

    if (action === 'updateProduct') {
      return updateProduct(data);
    }

    if (action === 'deleteProduct') {
      return deleteProduct(data.id);
    }

    // ───── CONTACT / ORDER ─────

    if (type === 'contact') {
      return handleContactForm(data);
    }

    if (type === 'order') {
      return handleOrderEmail(data);
    }

    if (type === 'validate') {
      return handleValidation();
    }

    return jsonResponse({
      success:false,
      error:'Invalid request'
    });

  } catch(error) {

    Logger.log(error);

    return jsonResponse({
      success:false,
      error:error.toString()
    });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// PRODUCT CRUD
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function addProduct(data) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  sheet.appendRow([
    data.id,
    data.name,
    data.price,
    data.category,
    data.description,
    data.image,
    data.badge
  ]);

  return jsonResponse({
    success:true,
    message:'Product added'
  });
}

function updateProduct(data) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {

    if (rows[i][0] == data.id) {

      sheet.getRange(i + 1, 1, 1, 7).setValues([[
        data.id,
        data.name,
        data.price,
        data.category,
        data.description,
        data.image,
        data.badge
      ]]);

      return jsonResponse({
        success:true,
        message:'Product updated'
      });
    }
  }

  return jsonResponse({
    success:false,
    error:'Product not found'
  });
}

function deleteProduct(id) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {

    if (rows[i][0] == id) {

      sheet.deleteRow(i + 1);

      return jsonResponse({
        success:true,
        message:'Product deleted'
      });
    }
  }

  return jsonResponse({
    success:false,
    error:'Product not found'
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// CONTACT FORM
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function handleContactForm(params) {

  const name = params.name || '';
  const email = params.email || '';
  const subject = params.subject || 'General Enquiry';
  const message = params.message || '';

  if (!name || !message) {

    return jsonResponse({
      success:false,
      error:'Name and message required'
    });
  }

  try {

    const adminSubject =
      `[Ivora Luxe] New Enquiry from ${name}`;

    const adminBody = `
New enquiry received:

Name: ${name}
Email: ${email}
Subject: ${subject}

Message:
${message}
    `.trim();

    GmailApp.sendEmail(
      ADMIN_EMAIL,
      adminSubject,
      adminBody
    );

    // Confirmation mail

    if (email) {

      GmailApp.sendEmail(
        email,
        'We received your message - Ivora Luxe',
`
Dear ${name},

Thank you for contacting Ivora Luxe.

We received your enquiry and our team will contact you shortly.

Best regards,
Ivora Luxe Team
ivoraluxen@gmail.com
`
      );
    }

    logContactToSheet({
      timestamp:new Date(),
      name,
      email,
      subject,
      message
    });

    return jsonResponse({
      success:true,
      message:'Enquiry sent successfully'
    });

  } catch(error) {

    return jsonResponse({
      success:false,
      error:error.toString()
    });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ORDER EMAIL
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function handleOrderEmail(params) {

  const customerName =
    params.customerName || 'Customer';

  const customerEmail =
    params.customerEmail || '';

  const orderDetails =
    params.orderDetails || '';

  const total =
    params.total || '0';

  try {

    GmailApp.sendEmail(
      ADMIN_EMAIL,
      `[Ivora Luxe Order] ${customerName}`,
`
Customer: ${customerName}
Email: ${customerEmail}

Order:
${orderDetails}

Total: ₹${total}
`
    );

    if (customerEmail) {

      GmailApp.sendEmail(
        customerEmail,
        'Order Confirmation - Ivora Luxe',
`
Dear ${customerName},

Thank you for your order!

Order Details:
${orderDetails}

Total: ₹${total}

We will contact you shortly.

Regards,
Ivora Luxe
`
      );
    }

    return jsonResponse({
      success:true,
      message:'Order email sent'
    });

  } catch(error) {

    return jsonResponse({
      success:false,
      error:error.toString()
    });
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// PRODUCTS JSON
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getProductsJSON() {

  const sheet =
    SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) {

    return jsonResponse({
      products:[]
    });
  }

  const data = sheet.getDataRange().getValues();

  const headers =
    data[0].map(h =>
      String(h).toLowerCase().trim()
    );

  const products = [];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const product = {};

    headers.forEach((header, idx) => {
      product[header] = row[idx] || '';
    });

    if (product.name) {
      products.push(product);
    }
  }

  return jsonResponse({
    products:products
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// VALIDATION
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function validateSheetStructure() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(SHEET_NAME);

  if (!sheet) {

    sheet = ss.insertSheet(SHEET_NAME);

    initializeSheet(sheet);

    SpreadsheetApp.getUi()
      .alert('Products sheet created');

    return;
  }

  const headers =
    sheet.getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    ).getValues()[0];

  const lower =
    headers.map(h =>
      String(h).toLowerCase().trim()
    );

  const missing =
    REQUIRED_COLUMNS.filter(
      col => !lower.includes(col)
    );

  if (missing.length > 0) {

    SpreadsheetApp.getUi()
      .alert(
        'Missing columns: ' +
        missing.join(', ')
      );

  } else {

    SpreadsheetApp.getUi()
      .alert('Sheet structure valid');
  }
}

function initializeSheet(sheet) {

  sheet.appendRow(REQUIRED_COLUMNS);

  sheet.getRange(
    1,
    1,
    1,
    REQUIRED_COLUMNS.length
  )
  .setBackground('#0e0c0a')
  .setFontColor('#c9a84c')
  .setFontWeight('bold');
}

function handleValidation() {

  validateSheetStructure();

  return jsonResponse({
    success:true
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// STATS
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function showStats() {

  const sheet =
    SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi()
      .alert('No products found');
    return;
  }

  const data =
    sheet.getDataRange().getValues();

  SpreadsheetApp.getUi().alert(
`
Total Products: ${data.length - 1}
Last Updated: ${new Date().toLocaleString()}
`
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// CONTACT LOG
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function logContactToSheet(data) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let logSheet =
    ss.getSheetByName('Contact Log');

  if (!logSheet) {

    logSheet = ss.insertSheet('Contact Log');

    logSheet.appendRow([
      'Timestamp',
      'Name',
      'Email',
      'Subject',
      'Message'
    ]);
  }

  logSheet.appendRow([
    data.timestamp,
    data.name,
    data.email,
    data.subject,
    data.message
  ]);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// UTILITIES
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function testEmail() {

  GmailApp.sendEmail(
    ADMIN_EMAIL,
    '[Ivora Luxe] Test Email',
    'Email system working correctly.'
  );

  SpreadsheetApp.getUi()
    .alert('Test email sent');
}

function clearSheet() {

  const ui = SpreadsheetApp.getUi();

  if (
    ui.alert(
      'Delete all products?',
      ui.ButtonSet.YES_NO
    ) === ui.Button.YES
  ) {

    const sheet =
      SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

    if (sheet && sheet.getLastRow() > 1) {

      sheet.deleteRows(
        2,
        sheet.getLastRow() - 1
      );
    }

    ui.alert('Products cleared');
  }
}

function showInstructions() {

  SpreadsheetApp.getUi().alert(
`
Required columns:

id
name
price
category
description
image
badge

Make sure:
✓ Sheet is public
✓ Web app deployed
✓ Access = Anyone
`
  );
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// JSON RESPONSE HELPER
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function jsonResponse(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}