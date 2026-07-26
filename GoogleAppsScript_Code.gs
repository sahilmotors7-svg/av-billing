// =================================================================
// AV TECHNOLOGY - Complete Google Apps Script (Google Drive PDF Link)
// Paste this entire file into Extensions > Apps Script in Google Sheets
// =================================================================

// Helper: Target the first sheet tab guaranteed
function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets()[0];
}

// Helper: Delete invoice by searching bottom-up across all columns
function deleteInvoiceByNumber(invoiceNo) {
  if (!invoiceNo) return false;
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  
  const targetClean = invoiceNo.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  if (!targetClean) return false;
  
  let deletedAny = false;
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    for (let j = 0; j < Math.min(row.length, 5); j++) {
      if (row[j]) {
        const cellClean = row[j].toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
        if (cellClean === targetClean) {
          sheet.deleteRow(i + 1);
          deletedAny = true;
          break;
        }
      }
    }
  }
  return deletedAny;
}

// Helper: Convert Number to Words (Indian Rupees)
function numberToWordsGS(num) {
  if (num === 0) return "Zero Rupees Only";
  const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  function convertLessThanThousand(n) {
    if (n === 0) return "";
    if (n < 20) return ones[n] + " ";
    if (n < 100) return tens[Math.floor(n/10)] + " " + (n%10 !== 0 ? ones[n%10] + " " : "");
    return ones[Math.floor(n/100)] + " Hundred " + (n%100 !== 0 ? convertLessThanThousand(n%100) : "");
  }
  let rupees = Math.floor(num);
  let paise = Math.round((num - rupees) * 100);
  let result = "";
  if (rupees >= 10000000) result += convertLessThanThousand(Math.floor(rupees/10000000)) + "Crore ";
  rupees %= 10000000;
  if (rupees >= 100000) result += convertLessThanThousand(Math.floor(rupees/100000)) + "Lakh ";
  rupees %= 100000;
  if (rupees >= 1000) result += convertLessThanThousand(Math.floor(rupees/1000)) + "Thousand ";
  rupees %= 1000;
  if (rupees >= 100) result += convertLessThanThousand(Math.floor(rupees/100)) + "Hundred ";
  rupees %= 100;
  if (rupees > 0) result += convertLessThanThousand(rupees);
  result = result.trim() + " Rupees";
  if (paise > 0) result += " and " + convertLessThanThousand(paise).trim() + " Paise";
  return result + " Only";
}

// Helper: Build Printable HTML for PDF Generation
function buildPdfHtmlString(inv) {
  let itemsRows = '';
  if (Array.isArray(inv.items)) {
    inv.items.forEach((it, idx) => {
      itemsRows += `<tr>
        <td style="border:1px solid #cfdfed; padding:6px; text-align:center;">${idx + 1}</td>
        <td style="border:1px solid #cfdfed; padding:6px; text-align:left;">${it.desc || ''}</td>
        <td style="border:1px solid #cfdfed; padding:6px; text-align:center;">${it.hsn || ''}</td>
        <td style="border:1px solid #cfdfed; padding:6px; text-align:center;">${it.qty || 1}</td>
        <td style="border:1px solid #cfdfed; padding:6px; text-align:right;">₹ ${parseFloat(it.rate || 0).toFixed(2)}</td>
        <td style="border:1px solid #cfdfed; padding:6px; text-align:right;">₹ ${parseFloat(it.amount || 0).toFixed(2)}</td>
      </tr>`;
    });
  }

  let gstRows = '';
  if (inv.isDelhi) {
    gstRows = `<tr><td style="padding:4px; text-align:left;">CGST @9%</td><td style="text-align:right;">₹ ${(inv.cgst||0).toFixed(2)}</td></tr>
               <tr><td style="padding:4px; text-align:left;">SGST @9%</td><td style="text-align:right;">₹ ${(inv.sgst||0).toFixed(2)}</td></tr>`;
  } else {
    gstRows = `<tr><td style="padding:4px; text-align:left;">IGST @18%</td><td style="text-align:right;">₹ ${(inv.igst||0).toFixed(2)}</td></tr>`;
  }

  const words = numberToWordsGS(inv.grandTotal || 0);

  return `<!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; padding: 20px; font-size: 12px; }
      .title { text-align: center; font-size: 20px; font-weight: bold; color: #1f3b4c; border-bottom: 2px solid #cfdfed; padding-bottom: 4px; display: inline-block; margin-bottom: 10px; }
      .brand-header { text-align: center; border-bottom: 1px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 15px; }
      .brand-name { font-size: 24px; font-weight: bold; color: #003153; }
      .brand-tag { font-size: 9px; color: #2c5a7a; }
      .company-details { font-size: 9px; color: #334155; margin-top: 4px; }
      .info-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; background: #fafcff; border: 1px solid #e2e8f0; border-radius: 8px; }
      .info-table td { padding: 8px 12px; vertical-align: top; font-size: 11px; }
      .items-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 11px; }
      .items-table th { background: #1e4a76; color: white; padding: 8px; border: 1px solid #1e4a76; }
      .total-box { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px; }
      .words { background: #eef2ff; padding: 8px 12px; border-radius: 6px; margin: 10px 0; font-weight: bold; }
      .terms { font-size: 9px; background: #fef7e0; padding: 8px; border-left: 3px solid #f59e0b; margin-top: 15px; }
      .sign-box { text-align: right; margin-top: 40px; font-weight: bold; font-size: 11px; }
    </style>
  </head>
  <body>
    <div style="text-align:center;"><div class="title">TAX INVOICE</div></div>
    <div class="brand-header">
      <div class="brand-name">AV TECHNOLOGY</div>
      <div class="brand-tag">PROJECTORS | VIDEO CAMERA | LED VIDEO WALLS | SPARE PARTS</div>
      <div class="company-details">20 PART 2 SAINIK VIHAR MOHAN GARDEN UTTAM NAGAR NEW DELHI -59<br>☎ +91 9711166056 ✉ avtechnologyy@gmail.com</div>
    </div>
    <table class="info-table">
      <tr>
        <td style="width:55%;">
          <strong>M/S (Customer):</strong> ${inv.customerName || ''}<br>
          <strong>GSTIN NO:</strong> ${inv.custGstin || '—'}<br>
          <strong>State:</strong> ${inv.stateName || 'Delhi'} (Code: ${inv.stateCode || '07'})<br>
          <strong>Place of Supply:</strong> ${inv.placeSupply || '—'}<br>
          <strong>Vehicle No:</strong> ${inv.vehicleNo || '—'}
        </td>
        <td style="width:45%; border-left: 1px solid #e2e8f0;">
          <strong>GSTIN:</strong> 07ABIFA3151F1ZS<br>
          <strong>STATE CODE:</strong> 07<br>
          <strong>INVOICE NO:</strong> ${inv.invoiceNo || ''}<br>
          <strong>DATE:</strong> ${inv.invoiceDate || ''}<br>
          <strong>Transport Mode:</strong> ${inv.transportMode || '—'}
        </td>
      </tr>
    </table>
    <table class="items-table">
      <thead>
        <tr><th>S.No</th><th>PARTICULAR</th><th>HSN</th><th>QTY</th><th>RATE</th><th>AMOUNT</th></tr>
      </thead>
      <tbody>
        ${itemsRows}
      </tbody>
    </table>
    <div style="text-align:right; font-weight:bold; margin-bottom:6px;">TAXABLE VALUE: ₹ ${(inv.taxable||0).toFixed(2)}</div>
    <div class="words">AMOUNT IN WORDS: ${words}</div>
    <table class="total-box">
      ${gstRows}
      <tr><td style="padding:4px; text-align:left;">CARTAGE:</td><td style="text-align:right;">₹ ${(inv.cartage||0).toFixed(2)}</td></tr>
      <tr style="background:#fef9e3; font-size:13px; font-weight:bold;"><td style="padding:6px; text-align:left;">GRAND TOTAL:</td><td style="text-align:right;">₹ ${(inv.grandTotal||0).toFixed(2)}</td></tr>
    </table>
    <div class="terms">
      <strong>TERMS & CONDITIONS:</strong><br>
      1. IF THE BILL IS NOT PAID ON DUE DATE INTEREST 18% P.A WILL BE CHARGED.<br>
      2. ALL DISPUTES SUBJECT TO DELHI JURISDICTION ONLY.<br>
      3. GOODS ONCE SOLD WILL NOT BE TAKEN BACK.
    </div>
    <div class="sign-box">
      FOR AV TECHNOLOGY<br><br><br>
      AUTHORISED SIGNATORY
    </div>
  </body>
  </html>`;
}

// Generate Google Drive PDF Link (No Gmail login required for client!)
function getGoogleDrivePdfUrl(invoiceNo) {
  if (!invoiceNo) return null;
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  
  const targetClean = invoiceNo.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  let row = null;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0]) {
      const cellClean = data[i][0].toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
      if (cellClean === targetClean) {
        row = data[i];
        break;
      }
    }
  }
  if (!row) return null;

  let itemsList = [];
  try { itemsList = JSON.parse(row[19] || "[]"); } catch(ex) { itemsList = []; }

  const inv = {
    invoiceNo: row[0],
    invoiceDate: row[1],
    customerName: row[2],
    custGstin: row[3],
    stateCode: row[4],
    stateName: row[5],
    vehicleNo: row[6],
    deliveryDate: row[7],
    placeSupply: row[8],
    transportMode: row[9],
    chequeNo: row[10],
    bankBranch: row[11],
    cartage: parseFloat(row[12]) || 0,
    taxable: parseFloat(row[13]) || 0,
    cgst: parseFloat(row[14]) || 0,
    sgst: parseFloat(row[15]) || 0,
    igst: parseFloat(row[16]) || 0,
    grandTotal: parseFloat(row[17]) || 0,
    isDelhi: row[18] === true || row[18] === "true",
    items: itemsList
  };

  const htmlContent = buildPdfHtmlString(inv);
  const blob = Utilities.newBlob(htmlContent, 'text/html', 'invoice.html').getAs('application/pdf');
  const cleanNo = inv.invoiceNo.toString().replace(/[^a-zA-Z0-9]/g, '_');
  blob.setName('Invoice_' + cleanNo + '.pdf');

  let folder;
  const folders = DriveApp.getFoldersByName("AV_Technology_Invoices");
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder("AV_Technology_Invoices");
  }

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

// ========== DO POST - Save, Update, or Delete Invoice ==========
function doPost(e) {
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(ex) {}
    }
    
    const urlAction = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
    const bodyAction = data.action || "";
    const isDelete = (urlAction === "delete" || bodyAction === "delete" || data.isDelete === true);
    
    const targetInvoiceNo = data.invoiceNo || (e && e.parameter && e.parameter.invoiceNo) || "";
    const sheet = getTargetSheet();
    
    // Ensure Header Row Exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "invoiceNo", "invoiceDate", "customerName", "custGstin", 
        "stateCode", "stateName", "vehicleNo", "deliveryDate", 
        "placeSupply", "transportMode", "chequeNo", "bankBranch",
        "cartage", "taxable", "cgst", "sgst", "igst", "grandTotal",
        "isDelhi", "items", "timestamp"
      ]);
    }

    // 1. DELETE ACTION
    if (isDelete || (e.postData && e.postData.contents && e.postData.contents.indexOf('"delete"') !== -1)) {
      const deleted = deleteInvoiceByNumber(targetInvoiceNo);
      return ContentService.createTextOutput("DELETED: " + deleted);
    }

    // 2. SAFEGUARD: STOPS BLANK ROWS
    if (!data.customerName || !data.items) {
      return ContentService.createTextOutput("IGNORED: Missing customerName or items");
    }

    // Find existing row index for UPDATE
    const allSheetData = sheet.getDataRange().getValues();
    let existingRowIndex = -1;
    const targetClean = targetInvoiceNo.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
    if (targetClean) {
      for (let i = allSheetData.length - 1; i >= 1; i--) {
        if (allSheetData[i][0]) {
          const cellClean = allSheetData[i][0].toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          if (cellClean === targetClean) {
            existingRowIndex = i + 1;
            break;
          }
        }
      }
    }

    const rowData = [
      data.invoiceNo || "", data.invoiceDate || "", data.customerName || "", data.custGstin || "",
      data.stateCode || "", data.stateName || "", data.vehicleNo || "", data.deliveryDate || "",
      data.placeSupply || "", data.transportMode || "", data.chequeNo || "", data.bankBranch || "",
      data.cartage || 0, data.taxable || 0, data.cgst || 0, data.sgst || 0, data.igst || 0, data.grandTotal || 0,
      data.isDelhi !== undefined ? data.isDelhi : true,
      JSON.stringify(data.items || []), new Date().toISOString()
    ];

    // 3. UPDATE ACTION
    if (existingRowIndex > 1 && (bodyAction === "update" || data.isUpdate)) {
      sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
      return ContentService.createTextOutput("UPDATED");
    }

    // 4. CREATE ACTION
    sheet.appendRow(rowData);
    return ContentService.createTextOutput("CREATED");

  } catch (err) {
    return ContentService.createTextOutput("ERROR: " + err.toString());
  }
}

// ========== DO GET - Retrieve Data & Delete / PDF Link ==========
function doGet(e) {
  const sheet = getTargetSheet();
  
  // Action: GET DELETE
  if (e && e.parameter && e.parameter.action === "delete" && e.parameter.invoiceNo) {
    const deleted = deleteInvoiceByNumber(e.parameter.invoiceNo);
    return ContentService.createTextOutput("DELETED: " + deleted);
  }

  // Action: GET PDF LINK (Public Google Drive Link)
  if (e && e.parameter && e.parameter.action === "getPdfLink" && e.parameter.invoiceNo) {
    try {
      const pdfUrl = getGoogleDrivePdfUrl(e.parameter.invoiceNo);
      return ContentService.createTextOutput(JSON.stringify({ pdfUrl: pdfUrl })).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    if (e && e.parameter && e.parameter.action === "getLast") {
      return ContentService.createTextOutput(JSON.stringify({ lastNo: "INV-26/000" })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  
  const rows = data.slice(1);

  // Action: Get all invoices
  if (e && e.parameter && e.parameter.action === "getAll") {
    const invoices = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0]) continue;
      let itemsList = [];
      try { itemsList = JSON.parse(row[19] || "[]"); } catch(ex) { itemsList = []; }
      
      invoices.push({
        invoiceNo: row[0],
        invoiceDate: row[1],
        customerName: row[2],
        custGstin: row[3],
        stateCode: row[4],
        stateName: row[5],
        vehicleNo: row[6],
        deliveryDate: row[7],
        placeSupply: row[8],
        transportMode: row[9],
        chequeNo: row[10],
        bankBranch: row[11],
        cartage: parseFloat(row[12]) || 0,
        taxable: parseFloat(row[13]) || 0,
        cgst: parseFloat(row[14]) || 0,
        sgst: parseFloat(row[15]) || 0,
        igst: parseFloat(row[16]) || 0,
        grandTotal: parseFloat(row[17]) || 0,
        isDelhi: row[18] === true || row[18] === "true",
        items: itemsList
      });
    }
    return ContentService.createTextOutput(JSON.stringify(invoices)).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Action: Get last invoice number
  if (e && e.parameter && e.parameter.action === "getLast") {
    if (rows.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ lastNo: "INV-26/000" })).setMimeType(ContentService.MimeType.JSON);
    }
    const lastInvoice = rows[rows.length-1][0];
    return ContentService.createTextOutput(JSON.stringify({ lastNo: lastInvoice })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ error: "Invalid request" })).setMimeType(ContentService.MimeType.JSON);
}

// ONE-CLICK AUTHORIZATION TEST FOR GOOGLE DRIVE (Run once in Apps Script Editor!)
function testDrivePdf() {
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    const lastInvoiceNo = data[data.length - 1][0];
    Logger.log("Generating Drive PDF Link for: " + lastInvoiceNo);
    const link = getGoogleDrivePdfUrl(lastInvoiceNo);
    Logger.log("Result PDF URL: " + link);
  } else {
    Logger.log("Please create at least 1 invoice in sheet first.");
  }
}
