// =================================================================
// AV TECHNOLOGY - Complete Google Apps Script (Google Drive PDF & Sandbox E-Way Bill)
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

// Helper: Authenticate with Sandbox.co.in GSP API
function getSandboxAuthToken() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GSP_API_KEY') || "";
  const apiSecret = props.getProperty('GSP_API_SECRET') || "";
  if (!apiKey || !apiSecret) {
    throw new Error("GSP_API_KEY or GSP_API_SECRET missing in Script Properties");
  }
  
  const options = {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'x-api-secret': apiSecret,
      'x-api-version': '1.0',
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  const res = UrlFetchApp.fetch('https://api.sandbox.co.in/authenticate', options);
  const json = JSON.parse(res.getContentText());
  if (json.access_token) {
    return json.access_token;
  }
  throw new Error(json.message || "Failed to authenticate with Sandbox GSP");
}

// Generate Official E-Way Bill via Sandbox GSP (Multi-Endpoint Auto-Routing & Exact Diagnostics)
function generateEWayBillDirectGSP(invoiceNo, invData) {
  try {
    const token = getSandboxAuthToken();
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GSP_API_KEY') || "";

    const itemList = (invData.items || []).map(it => ({
      productName: it.desc,
      hsnCode: parseInt(it.hsn) || 8528,
      quantity: parseFloat(it.qty) || 1,
      qtyUnit: "NOS",
      taxableAmount: parseFloat(it.amount),
      cgstRate: invData.isDelhi ? 9 : 0,
      sgstRate: invData.isDelhi ? 9 : 0,
      igstRate: invData.isDelhi ? 0 : 18
    }));

    const ewbPayload = {
      userGstin: "07ABIFA3151F1ZS",
      supplyType: "O",
      subSupplyType: "1",
      docType: "INV",
      docNo: invData.invoiceNo,
      docDate: invData.invoiceDate,
      fromGstin: "07ABIFA3151F1ZS",
      fromStateCode: 7,
      toGstin: invData.custGstin || "URP",
      toStateCode: parseInt(invData.stateCode) || 7,
      totalValue: invData.taxable,
      cgstValue: invData.cgst,
      sgstValue: invData.sgst,
      igstValue: invData.igst,
      totInvValue: invData.grandTotal,
      transMode: "1",
      transDistance: "15",
      vehicleNo: invData.vehicleNo || "",
      vehicleType: "R",
      itemList: itemList
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': token,
        'x-api-key': apiKey,
        'x-api-version': '1.0',
        'x-user-gstin': '07ABIFA3151F1ZS',
        'gstin': '07ABIFA3151F1ZS',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(ewbPayload),
      muteHttpExceptions: true
    };

    const endpoints = [
      'https://api.sandbox.co.in/gsp/ewaybill',
      'https://api.sandbox.co.in/gst/ewaybill',
      'https://api.sandbox.co.in/gsp/v1.0/ewaybill',
      'https://api.sandbox.co.in/gsp/v1.0/ewaybill/generate',
      'https://api.sandbox.co.in/gst/v1.0/ewaybill',
      'https://api.sandbox.co.in/ewaybill/v1.0/generate',
      'https://api.sandbox.co.in/gsp/ewaybill/v1.03/generate',
      'https://api.sandbox.co.in/gst/ewaybill/v1.03/generate'
    ];

    let diagnosticLogs = [];

    for (let i = 0; i < endpoints.length; i++) {
      const res = UrlFetchApp.fetch(endpoints[i], options);
      const code = res.getResponseCode();
      const rawText = res.getContentText();
      diagnosticLogs.push(`[${code}] ${endpoints[i]} -> ${rawText}`);

      if (code !== 404) {
        let json = {};
        try { json = JSON.parse(rawText); } catch(ex) {}
        if (json.ewayBillNo || (json.data && json.data.ewayBillNo) || (json.result && json.result.ewayBillNo)) {
          const ewbNo = json.ewayBillNo || (json.data && json.data.ewayBillNo) || (json.result && json.result.ewayBillNo);
          const vUpto = json.validUpto || (json.data && json.data.validUpto);
          return { status: "SUCCESS", ewayBillNo: ewbNo, validUpto: vUpto };
        }
        return { status: "INFO", message: json.message || json.error || rawText };
      }
    }

    return { status: "INFO", message: "Sandbox Response:\n" + diagnosticLogs.join("\n") };
  } catch(e) {
    return { status: "ERROR", message: e.toString() };
  }
}

// ========== DO POST - Save, Update, Delete Invoice or Generate E-Way Bill ==========
function doPost(e) {
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); } catch(ex) {}
    }
    
    const urlAction = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
    const bodyAction = data.action || "";
    
    // E-Way Bill Direct Generation Action
    if (bodyAction === "generateEWayBill" || urlAction === "generateEWayBill") {
      const ewbRes = generateEWayBillDirectGSP(data.invoiceNo, data.invoiceData || data);
      return ContentService.createTextOutput(JSON.stringify(ewbRes)).setMimeType(ContentService.MimeType.JSON);
    }

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

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
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

  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}
