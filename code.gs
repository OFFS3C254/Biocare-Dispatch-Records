/**
 * BIOCARE DISPATCH PORTAL — Backend (code.gs)
 * ======================================================================
 * Biocare Health Systems Ltd. — Daily Dispatch Database & API
 * 
 * Supports:
 * 1. Native Google Apps Script Web App (direct browser UI)
 * 2. External Cloud / Vercel Webhook REST API (JSON + CORS)
 * 3. Multi-date navigation across daily sheets (DD-MM-YYYY)
 * 4. Safe sheet duplication from 'Template'
 * 5. Dynamic Remarks & Configurations from 'Settings'
 * 6. Historical customer/facility directory scanning
 * 7. Excel (.xlsx) batch order importation
 * 8. PDF Evening Report generation & Reports_Log logging
 * 9. Automated executive PDF report generation & team email dispatch
 * ======================================================================
 */

const COMPANY_NAME = "Biocare Health Systems Ltd.";

// ======================================================================
// CONFIGURATION VARIABLES: Recipient Emails for Automated PDF Reports
// ======================================================================
// Add or edit emails below to customize who receives the daily PDF dispatch report:
const REPORT_RECIPIENT_EMAILS = [
  "biocarehealthsystems@gmail.com",
  "alexandremuithya@gmail.com"
  // Add additional team emails here, e.g. "operations@biocare.co.ke"
];
const SALES_TEAM_EMAILS = REPORT_RECIPIENT_EMAILS; // alias for backwards compatibility
const LOGO_DRIVE_FILE_ID = "1RyzfF7YwM7lLREkfCl96KupxSBZmGlNP"; // Optional Google Drive file ID

const TEMPLATE_SHEET_NAME = "Template";
const REPORTS_LOG_SHEET_NAME = "Reports_Log";
const SETTINGS_SHEET_NAME = "Settings";
const DATA_START_ROW = 4;
const DATA_END_ROW = 103;       // 100 order slots (rows 4 to 103)
const NUM_COLUMNS = 5;          // Date, Order No, Customer, Status, Remarks

// Minimalist Modern Brand Tokens
const COLOR_ACCENT = "#0052FF";       // Electric Blue
const COLOR_ACCENT_LIGHT = "#EBF2FF";
const COLOR_FOREGROUND = "#0F172A";   // Deep Slate
const COLOR_MUTED = "#64748B";
const COLOR_GREEN = "#059669";
const COLOR_GREEN_LIGHT = "#D1FAE5";

// ======================================================================
// WEB APP & REST API ROUTING (doGet & doPost)
// ======================================================================

function doGet(e) {
  // Check if called as a REST API endpoint (e.g. from Vercel or curl)
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  if (action) {
    return handleApiGet_(e);
  }

  // Otherwise, serve the HTML frontend
  let html;
  try {
    html = HtmlService.createTemplateFromFile('Index').evaluate();
  } catch (err) {
    try {
      html = HtmlService.createHtmlOutputFromFile('Index');
    } catch (err2) {
      html = HtmlService.createHtmlOutputFromFile('index');
    }
  }

  return html
    .setTitle('Biocare Dispatch Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return handleApiPost_(e);
}

function createJsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiGet_(e) {
  const action = e.parameter.action;
  const targetDate = e.parameter.targetDate || e.parameter.date || "";

  try {
    if (action === "getData") {
      const data = getAppData(targetDate);
      return createJsonResponse_({ success: true, ...data });
    }
    if (action === "ping") {
      return createJsonResponse_({ success: true, timestamp: new Date().toISOString() });
    }
    return createJsonResponse_({ success: false, error: "Unknown action: " + action });
  } catch (err) {
    return createJsonResponse_({ success: false, error: err.toString() });
  }
}

function handleApiPost_(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action || "";
    const targetDate = payload.targetDate || payload.date || "";

    if (action === "addRecord") {
      const result = addDispatchRecord(payload.record, targetDate);
      return createJsonResponse_({ success: true, ...result });
    }
    if (action === "updateRecord") {
      const result = updateDispatchRecord(payload.rowNumber, payload.record, targetDate);
      return createJsonResponse_({ success: true, ...result });
    }
    if (action === "deleteRecord") {
      const result = deleteDispatchRecord(payload.rowNumber, targetDate);
      return createJsonResponse_({ success: true, ...result });
    }
    if (action === "batchAdd") {
      const result = batchAddDispatchRecords(payload.records, targetDate);
      return createJsonResponse_({ success: true, ...result });
    }
    if (action === "generateReport" || action === "emailReport") {
      const customRecipients = payload.recipients || payload.emails || null;
      const result = generateEveningReport(targetDate, customRecipients);
      return createJsonResponse_({ success: true, ...result });
    }
    if (action === "saveSettings") {
      const result = saveConfigurations(payload.configs || payload.config || {});
      return createJsonResponse_({ success: true, ...result });
    }
    if (action === "addRemark") {
      const result = addRemarkOption(payload.remark);
      return createJsonResponse_({ success: true, remarks: result });
    }

    return createJsonResponse_({ success: false, error: "Unrecognized POST action: " + action });
  } catch (err) {
    return createJsonResponse_({ success: false, error: err.toString() });
  }
}

// ======================================================================
// DATE UTILITIES
// ======================================================================

function formatDateLabel_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || "GMT+3", "dd-MM-yyyy");
}

function normalizeDateLabel_(dateParam) {
  if (!dateParam) {
    return formatDateLabel_(new Date());
  }
  if (dateParam instanceof Date) {
    return formatDateLabel_(dateParam);
  }
  const str = String(dateParam).trim();
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const parts = str.split("-");
    return parts[2] + "-" + parts[1] + "-" + parts[0];
  }
  // DD-MM-YYYY
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) {
    return str;
  }
  return formatDateLabel_(new Date());
}

// ======================================================================
// SETTINGS & CONFIGURATIONS
// ======================================================================

function ensureSettingsSheet_(ss) {
  let settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settings) {
    settings = ss.insertSheet(SETTINGS_SHEET_NAME);
    settings.getRange(1, 1).setValue("Remarks Options")
      .setFontWeight("bold").setBackground(COLOR_NAVY).setFontColor("#FFFFFF");
    const defaultRemarks = [
      ["Fargo"],
      ["Bolt"],
      ["Naekana"],
      ["Customer Picked From Office"],
      ["Courier Service Used"]
    ];
    settings.getRange(2, 1, defaultRemarks.length, 1).setValues(defaultRemarks);
    settings.setColumnWidth(1, 260);
  }

  // Ensure WhatsApp and Sales Team settings columns exist (Cols C & D)
  if (settings.getLastColumn() < 4 || settings.getRange(1, 3).getValue() === "") {
    settings.getRange(1, 3, 1, 2).setValues([["Configuration Key", "Configuration Value"]])
      .setFontWeight("bold").setBackground(COLOR_PINK).setFontColor("#FFFFFF");
    const defaultConfigs = [
      ["WhatsApp_Gateway_Url", ""],
      ["WhatsApp_Token", ""],
      ["WhatsApp_Chat_ID", ""],
      ["WhatsApp_Auto_Send", "FALSE"],
      ["Sales_Team_Emails", SALES_TEAM_EMAILS.join(", ")]
    ];
    settings.getRange(2, 3, defaultConfigs.length, 2).setValues(defaultConfigs);
    settings.setColumnWidth(3, 200);
    settings.setColumnWidth(4, 320);
  }
  return settings;
}

function getRemarksOptions_(ss) {
  const settings = ensureSettingsSheet_(ss);
  const lastRow = settings.getLastRow();
  if (lastRow <= 1) {
    return ["Fargo", "Bolt", "Naekana", "Customer Picked From Office", "Courier Service Used"];
  }
  const opts = settings.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(String).map(s => s.trim()).filter(Boolean);
  return opts.length ? Array.from(new Set(opts)) : ["Fargo", "Bolt", "Naekana", "Customer Picked From Office", "Courier Service Used"];
}

function addRemarkOption(remark) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ensureSettingsSheet_(ss);
  remark = String(remark || "").trim();
  if (!remark) return getRemarksOptions_(ss);

  const existing = getRemarksOptions_(ss);
  if (existing.includes(remark)) return existing;

  const nextRow = settings.getLastRow() + 1;
  settings.getRange(nextRow, 1).setValue(remark);
  return getRemarksOptions_(ss);
}

function getConfigurations_(ss) {
  const settings = ensureSettingsSheet_(ss);
  const lastRow = settings.getLastRow();
  const config = {
    whatsappGatewayUrl: "",
    whatsappToken: "",
    whatsappChatId: "",
    whatsappAutoSend: false,
    salesEmails: SALES_TEAM_EMAILS
  };
  if (lastRow > 1) {
    const vals = settings.getRange(2, 3, lastRow - 1, 2).getValues();
    vals.forEach(row => {
      const k = String(row[0] || "").trim();
      const v = String(row[1] || "").trim();
      if (!k) return;
      if (k === "WhatsApp_Gateway_Url") config.whatsappGatewayUrl = v;
      if (k === "WhatsApp_Token") config.whatsappToken = v;
      if (k === "WhatsApp_Chat_ID") config.whatsappChatId = v;
      if (k === "WhatsApp_Auto_Send") config.whatsappAutoSend = (v.toUpperCase() === "TRUE");
      if (k === "Sales_Team_Emails" && v) {
        config.salesEmails = v.split(",").map(e => e.trim()).filter(Boolean);
      }
    });
  }
  return config;
}

function saveConfigurations(configObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ensureSettingsSheet_(ss);
  const rows = [
    ["WhatsApp_Gateway_Url", configObj.whatsappGatewayUrl || ""],
    ["WhatsApp_Token", configObj.whatsappToken || ""],
    ["WhatsApp_Chat_ID", configObj.whatsappChatId || ""],
    ["WhatsApp_Auto_Send", configObj.whatsappAutoSend ? "TRUE" : "FALSE"],
    ["Sales_Team_Emails", Array.isArray(configObj.salesEmails) ? configObj.salesEmails.join(", ") : (configObj.salesEmails || SALES_TEAM_EMAILS.join(", "))]
  ];
  settings.getRange(2, 3, rows.length, 2).setValues(rows);
  return { success: true };
}

// ======================================================================
// SHEET DISCOVERY & MANAGEMENT
// ======================================================================

function getAvailableDates_(ss) {
  const sheets = ss.getSheets();
  const dateSheets = [];
  const dateRegex = /^\d{2}-\d{2}-\d{4}$/;
  sheets.forEach(s => {
    const name = s.getName();
    if (dateRegex.test(name)) {
      dateSheets.push(name);
    }
  });
  // Sort descending (most recent date first)
  dateSheets.sort((a, b) => {
    const pa = a.split("-");
    const pb = b.split("-");
    const da = new Date(Number(pa[2]), Number(pa[1]) - 1, Number(pa[0]));
    const db = new Date(Number(pb[2]), Number(pb[1]) - 1, Number(pb[0]));
    return db.getTime() - da.getTime();
  });
  return dateSheets;
}

function getFacilityDirectory_(ss) {
  const facilitiesSet = {};

  // Check dedicated Facilities or Customers sheet
  const dedicatedSheet = ss.getSheetByName("Facilities") || ss.getSheetByName("Customers");
  if (dedicatedSheet) {
    const lr = dedicatedSheet.getLastRow();
    if (lr >= 1) {
      const names = dedicatedSheet.getRange(1, 1, lr, 1).getValues().flat();
      names.forEach(n => {
        const str = String(n || "").trim();
        if (str && !/^(facility|customer|name|client)$/i.test(str)) {
          facilitiesSet[str] = true;
        }
      });
    }
  }

  // Scan recent 6 date sheets
  const dateSheets = getAvailableDates_(ss).slice(0, 6);
  dateSheets.forEach(sheetName => {
    const s = ss.getSheetByName(sheetName);
    if (!s) return;
    const colValues = s.getRange(DATA_START_ROW, 3, DATA_END_ROW - DATA_START_ROW + 1, 1).getValues().flat();
    colValues.forEach(val => {
      const str = String(val || "").trim();
      if (str) facilitiesSet[str] = true;
    });
  });

  return Object.keys(facilitiesSet).sort((a, b) => a.localeCompare(b));
}

function ensureDateSheet_(label) {
  label = normalizeDateLabel_(label);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Could not access active Google Spreadsheet. Please ensure this script is open inside your Biocare Dispatch Tracker sheet.");
  }
  let sheet = ss.getSheetByName(label);
  if (sheet) {
    if (sheet.getMaxRows() < DATA_END_ROW) {
      sheet.insertRowsAfter(sheet.getMaxRows(), DATA_END_ROW - sheet.getMaxRows());
    }
    return sheet;
  }

  const template = ss.getSheetByName(TEMPLATE_SHEET_NAME);
  if (template) {
    sheet = template.copyTo(ss);
    sheet.setName(label);
    sheet.showSheet();
  } else {
    sheet = ss.insertSheet(label);
    sheet.getRange("A1").setValue("Daily Dispatch Tracker — " + label)
      .setFontSize(14).setFontWeight("bold").setFontColor(COLOR_NAVY);
    sheet.getRange("A2").setValue("Biocare Health Systems Ltd. — Daily Dispatch Database")
      .setFontStyle("italic").setFontColor("#666666");
    const headers = [["Date", "Order No", "Customer / Facility / Individual", "Status", "Remarks"]];
    sheet.getRange(3, 1, 1, NUM_COLUMNS).setValues(headers)
      .setFontWeight("bold").setBackground(COLOR_NAVY).setFontColor("#FFFFFF");
  }

  sheet.getRange("A1").setValue("Daily Dispatch Tracker — " + label);
  if (sheet.getMaxRows() < DATA_END_ROW) {
    sheet.insertRowsAfter(sheet.getMaxRows(), DATA_END_ROW - sheet.getMaxRows());
  }
  sheet.getRange(DATA_START_ROW, 1, DATA_END_ROW - DATA_START_ROW + 1, 1).setValue(label);
  applyConditionalFormatting_(sheet);

  return sheet;
}

function applyConditionalFormatting_(sheet) {
  const statusRange = sheet.getRange(DATA_START_ROW, 4, DATA_END_ROW - DATA_START_ROW + 1, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Dispatched")
      .setBackground(COLOR_NAVY_LIGHT).setFontColor(COLOR_NAVY)
      .setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Pending")
      .setBackground(COLOR_PINK_LIGHT).setFontColor(COLOR_PINK)
      .setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("To Be Dispatched Tomorrow")
      .setBackground("#F3D9E8").setFontColor("#8E1550")
      .setRanges([statusRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

// ======================================================================
// CORE DATA API (Used by google.script.run & REST API)
// ======================================================================

function getAppData(targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Could not access active Google Spreadsheet. Please ensure this script is created inside your Biocare Dispatch Tracker sheet.");
  }
  const label = normalizeDateLabel_(targetDate);
  const todayLabel = formatDateLabel_(new Date());
  const sheet = ensureDateSheet_(label);

  const remarksOptions = getRemarksOptions_(ss);
  const availableDates = getAvailableDates_(ss);
  const facilities = getFacilityDirectory_(ss);
  const configs = getConfigurations_(ss);

  // Read fixed rows 4..103
  const rawData = sheet.getRange(DATA_START_ROW, 1, DATA_END_ROW - DATA_START_ROW + 1, NUM_COLUMNS).getValues();

  const records = [];
  rawData.forEach((row, i) => {
    if (row[1] !== "" && row[1] !== null) {
      records.push({
        row: DATA_START_ROW + i,
        date: String(row[0] || label),
        orderNo: String(row[1]).trim(),
        customer: String(row[2]).trim(),
        status: String(row[3]).trim(),
        remarks: String(row[4]).trim()
      });
    }
  });

  const statusValues = records.map(r => r.status);
  const count = (arr, val) => arr.filter(v => v === val).length;
  const stats = {
    total: records.length,
    dispatched: count(statusValues, "Dispatched"),
    pending: count(statusValues, "Pending"),
    toDispatchTomorrow: count(statusValues, "To Be Dispatched Tomorrow"),
    availableSlots: (DATA_END_ROW - DATA_START_ROW + 1) - records.length
  };

  return {
    label: label,
    todayLabel: todayLabel,
    isToday: (label === todayLabel),
    stats: stats,
    records: records,
    remarksOptions: remarksOptions,
    availableDates: availableDates,
    facilities: facilities,
    configs: configs
  };
}

function addDispatchRecord(record, targetDate) {
  const label = normalizeDateLabel_(targetDate);
  const sheet = ensureDateSheet_(label);

  const colValues = sheet.getRange(DATA_START_ROW, 2, DATA_END_ROW - DATA_START_ROW + 1, 1).getValues();
  let rowToInsert = -1;
  for (let i = 0; i < colValues.length; i++) {
    if (colValues[i][0] === "" || colValues[i][0] === null) {
      rowToInsert = DATA_START_ROW + i;
      break;
    }
  }
  if (rowToInsert === -1) {
    throw new Error(`Sheet for ${label} is full (100 rows maximum).`);
  }

  sheet.getRange(rowToInsert, 1, 1, NUM_COLUMNS).setValues([[
    label,
    String(record.orderNo || record.orderRef || record.order_no || record.ref || "").trim(),
    String(record.customer || record.facility || record.client || record.name || "").trim(),
    String(record.status || "Dispatched").trim(),
    String(record.remarks || "").trim()
  ]]);
  applyConditionalFormatting_(sheet);

  return getAppData(label);
}

function updateDispatchRecord(rowNumber, record, targetDate) {
  const label = normalizeDateLabel_(targetDate);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss ? ss.getSheetByName(label) : null;
  if (!sheet) throw new Error(`Sheet for ${label} not found.`);

  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < DATA_START_ROW || rowNumber > DATA_END_ROW) {
    throw new Error("Invalid row reference: " + rowNumber);
  }

  sheet.getRange(rowNumber, 1, 1, NUM_COLUMNS).setValues([[
    label,
    String(record.orderNo || record.orderRef || record.order_no || record.ref || "").trim(),
    String(record.customer || record.facility || record.client || record.name || "").trim(),
    String(record.status || "Dispatched").trim(),
    String(record.remarks || "").trim()
  ]]);
  applyConditionalFormatting_(sheet);
  return getAppData(label);
}

function deleteDispatchRecord(rowNumber, targetDate) {
  const label = normalizeDateLabel_(targetDate);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss ? ss.getSheetByName(label) : null;
  if (!sheet) throw new Error(`Sheet for ${label} not found.`);

  rowNumber = Number(rowNumber);
  if (!rowNumber || rowNumber < DATA_START_ROW || rowNumber > DATA_END_ROW) {
    throw new Error("Invalid row reference: " + rowNumber);
  }

  sheet.getRange(rowNumber, 2, 1, NUM_COLUMNS - 1).clearContent();
  return getAppData(label);
}

function batchAddDispatchRecords(records, targetDate) {
  if (!records || !records.length) {
    throw new Error("No order records provided to import.");
  }
  const label = normalizeDateLabel_(targetDate);
  const sheet = ensureDateSheet_(label);

  const colValues = sheet.getRange(DATA_START_ROW, 2, DATA_END_ROW - DATA_START_ROW + 1, 1).getValues();
  const emptyRowIndices = [];
  for (let i = 0; i < colValues.length; i++) {
    const val = String(colValues[i][0] || "").trim();
    if (val === "") {
      emptyRowIndices.push(DATA_START_ROW + i);
    }
  }

  if (records.length > emptyRowIndices.length) {
    throw new Error(`Only ${emptyRowIndices.length} slots available on ${label}. You are trying to import ${records.length} records.`);
  }

  for (let k = 0; k < records.length; k++) {
    const r = records[k];
    const targetRow = emptyRowIndices[k];
    sheet.getRange(targetRow, 1, 1, NUM_COLUMNS).setValues([[
      label,
      String(r.orderNo || r.orderRef || r.order_no || r.ref || r["Order Reference"] || r["Order No"] || "").trim(),
      String(r.customer || r.facility || r.client || r.name || r["Customer"] || "").trim(),
      String(r.status || "Dispatched").trim(),
      String(r.remarks || "").trim()
    ]]);
  }

  applyConditionalFormatting_(sheet);
  return getAppData(label);
}

// ======================================================================
// EVENING REPORT & HIGH-FIDELITY PDF GENERATION (MINIMALIST MODERN)
// ======================================================================

const BIOCARE_LOGO_BASE64 = "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bAZQB07fRUREREREREREREQRMJBiIAd3Hth4+OTgI12e8ZuinJIKQCCwcs38DBQtyMTiOTOwcPaMsFOB5V67CDkrFiLHnpeQ0SrjQ6MY7hlE/8kL6O/sweDJ0FMPuvvH0T/qx9mBi3D1+/DRBR+iCTZJOHEleNLFEkdEREREREREREREchhIMaCDOw/MGxjzf+vo2ZHvtJ4escU44gLAleCKPScD82alYfGcDMyZmTYtwDLn6lzMts3DzHmzkGPPQ/rMGZg1f44qxzPg9mB8YBRj/aPoP9GDUc/wlAXje4cvwuefQLdnHADQ3juGc0P+aKfoCucAgCYweEJEREREREREREREMWAgxeAO7jwwz+Ud+5tzgxe/1nlh7GOvnhydoda2l89Nx+LZ6Zg1IxVLczIAAPOy0pCTdSXIsmRRNvLyr4xYyV4yD+mZoXdhvG8EY95hjF68hP6BcVzouAAA6B/1o2/0EgDAO+pH73AgQPK2e0yLJOuGEDgB0MRpu4iIiIiIiIiIiIgoHgykmMzBnQdWv3Giv6p32F/a7fUtO+a5mJHkSeIE0IZA8KSJo06IiIiIiIiIiIiISE0MpJjcrs1PLfeM+O+7NIGy0wMXrzt6wbfYwofbjUDQJBg4aeOIEyIiIiIiIiIiIiLSEgMpFlS8Zu9qAKuXzUlfPe6fLB2+OHHViH/STAGWbgBdCARMuhAImDQxZ4mIiIiIiIiIiIhIbwykJJHiNXtLAcwDsFr0XwBYr/OuNAv/7RL+9SEQNOlztla1MaeIiIiIiIiIiIiIyCgYSKHLitfsFQdXgCsBl2gFAyOX/2aAhIiIiIiIiIiIiIjMiIEUIiIiIiIiIiIiIiKiMFKZBERERERERERERERERKExkEJERERERERERERERBQGAylERERERERERERERERhMJBCREREREREREREREQUBgMpREREREREREREREREYTCQQkREREREREREREREFAYDKURERERERERERERERGEwkEJERERERERERERERBQGAylERERERERERERERERhMJBCREREREREREREREQUBgMpREREREREREREREREYTCQQkREREREREREREREFAYDKURERERERERERERERGEwkEJERERERERERERERBQGAylERERERERERERERERhMJBCREREREREREREREQUBgMpREREREREREREREREYTCQQkREREREREREREREFAYDKURERERERERERERERGEwkEJERERERERERERERBQGAylERERERERERERERERhMJBCREREREREREREREQUBgMpREREREREREREREREYTCQQkREREREREREREREFAYDKURERERERERERERERGEwkEJERERERERERERERBQGAylERERERERERERERERhMJBCREREREREREREREQUBgMpREREREREREREREREYfxffdeCdCtbWjkAAAAASUVORK5CYII=";

function generateEveningReport(targetDate, customRecipients) {
  const label = normalizeDateLabel_(targetDate);
  const appData = getAppData(label);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configs = getConfigurations_(ss);

  // Parse recipient emails
  let recipients = [];
  if (customRecipients) {
    if (Array.isArray(customRecipients)) {
      recipients = customRecipients.map(String).map(e => e.trim()).filter(Boolean);
    } else if (typeof customRecipients === 'string') {
      recipients = customRecipients.split(',').map(e => e.trim()).filter(Boolean);
    }
  }
  if (!recipients.length) {
    recipients = (configs.salesEmails && configs.salesEmails.length) ? configs.salesEmails : REPORT_RECIPIENT_EMAILS;
  }

  const remarksTally = {};
  appData.records.forEach(r => {
    const key = r.remarks || "(blank)";
    remarksTally[key] = (remarksTally[key] || 0) + 1;
  });
  const remarksSummaryText = Object.keys(remarksTally).length
    ? Object.keys(remarksTally).map(k => k + ": " + remarksTally[k]).join(", ")
    : "—";

  let logSheet = ss.getSheetByName(REPORTS_LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(REPORTS_LOG_SHEET_NAME);
    logSheet.appendRow(["Date", "Total Orders", "Dispatched", "Pending", "To Be Dispatched Tomorrow", "Remarks Breakdown", "Report PDF URL", "Generated At"]);
  }

  let pdfUrl = "";
  let pdfDownloadUrl = "";
  let emailNotice = "";

  try {
    const pdfBlob = buildReportPDF_(appData.label, appData.stats, appData.records, remarksTally);
    const pdfFile = DriveApp.createFile(pdfBlob);
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (permErr) {
      Logger.log("Sharing permission warning: " + permErr);
    }
    pdfUrl = pdfFile.getUrl();
    pdfDownloadUrl = "https://drive.google.com/uc?export=download&id=" + pdfFile.getId();

    // Send email with PDF attachment to specified recipients
    if (recipients.length > 0) {
      try {
        const plainBody = "Hello Team,\n\n" +
          "Please find attached the official Daily Dispatch Report for " + appData.label + " from " + COMPANY_NAME + ".\n\n" +
          "Summary of Today's Dispatches:\n" +
          "• Total Orders Logged: " + appData.stats.total + "\n" +
          "• Dispatched: " + appData.stats.dispatched + "\n" +
          "• Pending: " + appData.stats.pending + "\n" +
          "• To Be Dispatched Tomorrow: " + appData.stats.toDispatchTomorrow + "\n\n" +
          "Courier & Remarks Breakdown: " + remarksSummaryText + "\n\n" +
          (pdfUrl ? "View Report Online: " + pdfUrl + "\n\n" : "") +
          "Best regards,\n" + COMPANY_NAME + " Dispatch Portal";

        const htmlBody = buildEmailHtmlBody_(appData.label, appData.stats, remarksTally, remarksSummaryText, pdfUrl);

        MailApp.sendEmail({
          to: recipients.join(","),
          subject: COMPANY_NAME + " — Daily Dispatch Operational Report — " + appData.label,
          body: plainBody,
          htmlBody: htmlBody,
          attachments: [pdfBlob]
        });
        emailNotice = "Emailed PDF to: " + recipients.join(", ");
      } catch (mailErr) {
        emailNotice = "Email delivery notice: " + (mailErr.message || mailErr);
        Logger.log("Email notification notice: " + mailErr);
      }
    }
  } catch (pdfErr) {
    emailNotice = "PDF compilation notice: " + (pdfErr.message || pdfErr);
    Logger.log("Notice: PDF generation error: " + pdfErr);
  }

  logSheet.appendRow([
    appData.label,
    appData.stats.total,
    appData.stats.dispatched,
    appData.stats.pending,
    appData.stats.toDispatchTomorrow,
    remarksSummaryText,
    pdfUrl || "Logged in Portal",
    new Date()
  ]);

  return {
    status: "Success",
    label: appData.label,
    pdfUrl: pdfUrl,
    pdfDownloadUrl: pdfDownloadUrl,
    recipients: recipients,
    emailNotice: emailNotice,
    stats: appData.stats
  };
}

function buildReportPDF_(label, stats, records, remarksTally) {
  const html = buildReportHtmlDocument_(label, stats, records, remarksTally);
  const pdfBlob = Utilities.newBlob(html, "text/html", COMPANY_NAME + " - Dispatch Manifest - " + label + ".pdf").getAs("application/pdf");
  return pdfBlob;
}

function buildReportHtmlDocument_(label, stats, records, remarksTally) {
  const remarksKeys = Object.keys(remarksTally || {});
  const totalRemarks = remarksKeys.reduce((acc, k) => acc + (remarksTally[k] || 0), 0);

  const completionPct = stats.total > 0 ? Math.round((stats.dispatched / stats.total) * 100) : 0;
  const pendingPct = stats.total > 0 ? Math.round((stats.pending / stats.total) * 100) : 0;
  const tomorrowPct = stats.total > 0 ? Math.round((stats.toDispatchTomorrow / stats.total) * 100) : 0;

  const now = new Date();
  const generatedDateStr = Utilities.formatDate(now, Session.getScriptTimeZone() || "GMT+3", "dd MMM yyyy");
  const generatedTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone() || "GMT+3", "HH:mm");

  const courierBadges = remarksKeys.map(k => {
    const count = remarksTally[k];
    const pct = totalRemarks > 0 ? Math.round((count / totalRemarks) * 100) : 0;
    return '<div style="display:inline-block; margin: 3px 6px 3px 0; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 4px 10px; font-size: 10px; color: #1E293B;">' +
      '<span style="font-weight: 700; color: #0052FF;">' + escapeXml_(k) + '</span> ' +
      '<span style="background: #EFF6FF; color: #0052FF; font-weight: 700; padding: 1px 6px; border-radius: 9999px; margin-left: 4px; font-size: 9px;">' + count + '</span>' +
      '<span style="color: #94A3B8; font-size: 8.5px; margin-left: 2px;">(' + pct + '%)</span>' +
      '</div>';
  }).join('');

  const orderRows = records.map((r, idx) => {
    const status = String(r.status || 'Dispatched').trim();
    let badgeBg = '#DCFCE7';
    let badgeColor = '#15803D';
    let badgeBorder = '#86EFAC';
    let badgeIcon = '✓';
    let badgeText = 'DISPATCHED';

    if (status === 'Pending') {
      badgeBg = '#FEF3C7';
      badgeColor = '#B45309';
      badgeBorder = '#FCD34D';
      badgeIcon = '⏳';
      badgeText = 'PENDING';
    } else if (status === 'To Be Dispatched Tomorrow') {
      badgeBg = '#EDE9FE';
      badgeColor = '#6D28D9';
      badgeBorder = '#C4B5FD';
      badgeIcon = '➔';
      badgeText = 'TOMORROW';
    }

    const rowBg = idx % 2 === 1 ? '#F8FAFC' : '#FFFFFF';

    return '<tr style="background: ' + rowBg + ';">' +
      '<td style="width: 32px; text-align: center; color: #64748B; font-weight: 600; font-size: 9.5px; padding: 7px 6px; border-bottom: 1px solid #E2E8F0;">' + (idx + 1) + '</td>' +
      '<td style="width: 140px; font-family: monospace; font-weight: 700; color: #0052FF; font-size: 10.5px; padding: 7px 8px; border-bottom: 1px solid #E2E8F0; letter-spacing: 0.02em;">' + escapeXml_(r.orderNo) + '</td>' +
      '<td style="padding: 7px 8px; font-size: 10px; color: #0F172A; font-weight: 600; border-bottom: 1px solid #E2E8F0;">' + escapeXml_(r.customer) + '</td>' +
      '<td style="width: 125px; text-align: center; padding: 7px 6px; border-bottom: 1px solid #E2E8F0;">' +
        '<span style="display:inline-block; padding:3px 8px; border-radius:9999px; font-size:8.5px; font-weight:700; letter-spacing:0.04em; white-space:nowrap; background:' + badgeBg + '; color:' + badgeColor + '; border:1px solid ' + badgeBorder + ';">' + badgeIcon + ' ' + badgeText + '</span>' +
      '</td>' +
      '<td style="width: 155px; padding: 7px 8px; font-size: 9.5px; color: #475569; border-bottom: 1px solid #E2E8F0;">' + escapeXml_(r.remarks || '—') + '</td>' +
    '</tr>';
  }).join('');

  return '<!DOCTYPE html>' +
'<html><head><meta charset="utf-8">' +
'<title>Daily Dispatch Report — ' + label + '</title>' +
'<style>' +
'  @page { size: A4 portrait; margin: 10mm 12mm 12mm 12mm; }' +
'  * { box-sizing: border-box; margin: 0; padding: 0; }' +
'  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #0F172A; background: #FFFFFF; font-size: 10px; line-height: 1.35; max-width: 900px; margin: 0 auto; padding: 16px; }' +
'  .header-table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }' +
'  .brand-logo-img { height: 40px; max-width: 280px; object-fit: contain; display: block; }' +
'  .header-tagline { font-size: 9.5px; color: #64748B; font-weight: 500; margin-top: 4px; letter-spacing: 0.02em; }' +
'  .meta-badge { display: inline-block; background: #0F172A; color: #FFFFFF; padding: 4px 10px; border-radius: 9999px; font-size: 9px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 4px; }' +
'  .meta-date { font-family: monospace; font-size: 13px; font-weight: 800; color: #0052FF; }' +
'  .meta-info-sub { font-size: 8.5px; color: #94A3B8; margin-top: 1px; }' +
'  .accent-bar { height: 3px; background: #0052FF; border-radius: 3px; margin: 8px 0 12px 0; }' +
'  .kpi-table { width: 100%; border-collapse: separate; border-spacing: 6px; margin-bottom: 12px; }' +
'  .kpi-cell { border-radius: 8px; padding: 8px 10px; vertical-align: top; width: 25%; }' +
'  .kpi-total { background: #0F172A; color: #FFFFFF; }' +
'  .kpi-dispatched { background: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; }' +
'  .kpi-pending { background: #FEF3C7; border: 1px solid #FDE68A; color: #B45309; }' +
'  .kpi-tomorrow { background: #F5F3FF; border: 1px solid #DDD6FE; color: #6D28D9; }' +
'  .kpi-label { font-size: 8px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; opacity: 0.9; margin-bottom: 2px; }' +
'  .kpi-val { font-size: 20px; font-weight: 800; line-height: 1.1; }' +
'  .kpi-sub { font-size: 8px; margin-top: 2px; opacity: 0.85; font-weight: 500; }' +
'  .progress-wrap { width: 100%; background: #E2E8F0; height: 6px; border-radius: 9999px; overflow: hidden; margin: 2px 0 12px 0; }' +
'  .section-title { font-size: 10.5px; font-weight: 800; color: #0F172A; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 6px; }' +
'  .courier-panel { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 6px 10px; margin-bottom: 12px; }' +
'  .orders-table { width: 100%; border-collapse: collapse; margin-top: 4px; border-radius: 6px; overflow: hidden; }' +
'  .orders-table th { background: #0F172A; color: #FFFFFF; font-weight: 700; font-size: 8.5px; text-transform: uppercase; letter-spacing: 0.06em; padding: 8px 8px; text-align: left; }' +
'  .audit-table { width: 100%; border-collapse: separate; border-spacing: 12px 0; margin-top: 18px; }' +
'  .audit-box { border: 1px solid #E2E8F0; border-radius: 6px; padding: 8px 12px; background: #F8FAFC; vertical-align: top; width: 50%; }' +
'  .audit-role { font-size: 8.5px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 22px; }' +
'  .audit-line { border-bottom: 1px dashed #94A3B8; margin-bottom: 4px; }' +
'  .doc-footer { margin-top: 14px; padding-top: 8px; border-top: 1px solid #E2E8F0; text-align: center; font-size: 8px; color: #94A3B8; }' +
'</style></head><body>' +
'  <table class="header-table">' +
'    <tr>' +
'      <td style="width:55%; vertical-align:middle;">' +
'        <img src="data:image/png;base64,' + BIOCARE_LOGO_BASE64 + '" alt="' + COMPANY_NAME + '" class="brand-logo-img" />' +
'        <div class="header-tagline">' + COMPANY_NAME + ' • Daily Operational Dispatch Manifest</div>' +
'      </td>' +
'      <td style="width:45%; vertical-align:middle; text-align:right;">' +
'        <div class="meta-badge">OFFICIAL DISPATCH MANIFEST</div>' +
'        <div class="meta-date">' + label + '</div>' +
'        <div class="meta-info-sub">Generated: ' + generatedDateStr + ' at ' + generatedTimeStr + '</div>' +
'        <div class="meta-info-sub">Manifest ID: #BHS-DISP-' + label.replace(/-/g, '') + '</div>' +
'      </td>' +
'    </tr>' +
'  </table>' +
'  <div class="accent-bar"></div>' +
'  <table class="kpi-table">' +
'    <tr>' +
'      <td class="kpi-cell kpi-total">' +
'        <div class="kpi-label">Total Logged</div>' +
'        <div class="kpi-val">' + stats.total + '</div>' +
'        <div class="kpi-sub">Total Consignments Logged</div>' +
'      </td>' +
'      <td class="kpi-cell kpi-dispatched">' +
'        <div class="kpi-label">Dispatched</div>' +
'        <div class="kpi-val" style="color:#0052FF;">' + stats.dispatched + '</div>' +
'        <div class="kpi-sub">' + completionPct + '% completed delivery</div>' +
'      </td>' +
'      <td class="kpi-cell kpi-pending">' +
'        <div class="kpi-label">Pending</div>' +
'        <div class="kpi-val">' + stats.pending + '</div>' +
'        <div class="kpi-sub">' + pendingPct + '% processing hold</div>' +
'      </td>' +
'      <td class="kpi-cell kpi-tomorrow">' +
'        <div class="kpi-label">For Tomorrow</div>' +
'        <div class="kpi-val">' + stats.toDispatchTomorrow + '</div>' +
'        <div class="kpi-sub">' + tomorrowPct + '% scheduled next cycle</div>' +
'      </td>' +
'    </tr>' +
'  </table>' +
'  <div class="section-title">Delivery Courier &amp; Distribution Breakdown</div>' +
'  <div class="courier-panel">' + (courierBadges || '<span style="color:#94A3B8; font-size:9px;">No courier remarks logged for this shift.</span>') + '</div>' +
'  <div class="section-title">Itemized Dispatch Records (' + records.length + ' items)</div>' +
'  <table class="orders-table">' +
'    <thead><tr><th style="width:32px; text-align:center;">#</th><th style="width:140px;">Order Reference</th><th>Customer / Facility / Client</th><th style="width:125px; text-align:center;">Status</th><th style="width:155px;">Courier / Remarks</th></tr></thead>' +
'    <tbody>' + (orderRows || '<tr><td colspan="5" style="text-align:center; padding:16px; color:#94A3B8;">No records logged for this date.</td></tr>') + '</tbody>' +
'  </table>' +
'  <table class="audit-table">' +
'    <tr>' +
'      <td class="audit-box"><div class="audit-role">Prepared By — Dispatch Operations</div><div class="audit-line"></div><div style="font-size:8px; color:#94A3B8; display:flex; justify-content:space-between;"><span>Officer Signature</span><span>Date: ________________</span></div></td>' +
'      <td class="audit-box"><div class="audit-role">Verified By — Warehouse Logistics Lead</div><div class="audit-line"></div><div style="font-size:8px; color:#94A3B8; display:flex; justify-content:space-between;"><span>Manager Signature</span><span>Date: ________________</span></div></td>' +
'    </tr>' +
'  </table>' +
'  <div class="doc-footer"><strong>' + COMPANY_NAME + '</strong> — Confidential Operational Logistics Record.<br>System Verified • Biocare Cloud Dispatch Tracker</div>' +
'</body></html>';
}

function buildEmailHtmlBody_(label, stats, remarksTally, remarksSummaryText, pdfUrl) {
  const completionPct = stats.total > 0 ? Math.round((stats.dispatched / stats.total) * 100) : 0;
  return '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden;">' +
    '<div style="background: #0F172A; padding: 24px; text-align: center; color: #FFFFFF;">' +
      '<div style="font-size: 11px; text-transform: uppercase; letter-spacing: 0.15em; color: #4D7CFF; font-weight: 700; margin-bottom: 6px;">Daily Dispatch Operational Report</div>' +
      '<h1 style="margin: 0; font-size: 22px; font-weight: 800; color: #FFFFFF;">' + COMPANY_NAME + '</h1>' +
      '<div style="margin-top: 8px; display: inline-block; background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 9999px; font-size: 12px; font-weight: 600; color: #E2E8F0;">Shift Date: ' + label + '</div>' +
    '</div>' +
    '<div style="padding: 24px;">' +
      '<p style="font-size: 14px; color: #334155; line-height: 1.5; margin-top: 0;">Hello Team,</p>' +
      '<p style="font-size: 14px; color: #334155; line-height: 1.5;">Please find attached the official <strong>Daily Dispatch Operational Report PDF</strong> for <strong>' + label + '</strong>. Below is the executive shift summary:</p>' +
      '<table style="width: 100%; border-collapse: separate; border-spacing: 8px; margin: 20px 0;">' +
        '<tr>' +
          '<td style="background: #0F172A; color: #FFFFFF; border-radius: 8px; padding: 12px; text-align: center; width: 25%;">' +
            '<div style="font-size: 9px; text-transform: uppercase; letter-spacing: 0.08em; opacity: 0.8;">Total Orders</div>' +
            '<div style="font-size: 22px; font-weight: 800; margin-top: 2px;">' + stats.total + '</div>' +
          '</td>' +
          '<td style="background: #EFF6FF; border: 1px solid #BFDBFE; color: #1D4ED8; border-radius: 8px; padding: 12px; text-align: center; width: 25%;">' +
            '<div style="font-size: 9px; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 700;">Dispatched</div>' +
            '<div style="font-size: 22px; font-weight: 800; color: #0052FF; margin-top: 2px;">' + stats.dispatched + '</div>' +
          '</td>' +
          '<td style="background: #FEF3C7; border: 1px solid #FDE68A; color: #B45309; border-radius: 8px; padding: 12px; text-align: center; width: 25%;">' +
            '<div style="font-size: 9px; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 700;">Pending</div>' +
            '<div style="font-size: 22px; font-weight: 800; margin-top: 2px;">' + stats.pending + '</div>' +
          '</td>' +
          '<td style="background: #F5F3FF; border: 1px solid #DDD6FE; color: #6D28D9; border-radius: 8px; padding: 12px; text-align: center; width: 25%;">' +
            '<div style="font-size: 9px; text-transform: uppercase; letter-spacing: 0.08em; font-weight: 700;">Tomorrow</div>' +
            '<div style="font-size: 22px; font-weight: 800; margin-top: 2px;">' + stats.toDispatchTomorrow + '</div>' +
          '</td>' +
        '</tr>' +
      '</table>' +
      '<div style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 14px; margin-bottom: 20px;">' +
        '<div style="font-size: 11px; font-weight: 700; text-transform: uppercase; color: #475569; letter-spacing: 0.05em; margin-bottom: 6px;">Courier / Remarks Breakdown</div>' +
        '<div style="font-size: 13px; color: #1E293B;">' + escapeXml_(remarksSummaryText) + '</div>' +
      '</div>' +
      (pdfUrl ? '<div style="text-align: center; margin: 24px 0;"><a href="' + pdfUrl + '" style="display: inline-block; background: #0052FF; color: #FFFFFF; font-weight: 700; font-size: 14px; text-decoration: none; padding: 12px 28px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,82,255,0.3);">View Report PDF on Google Drive &rarr;</a></div>' : '') +
      '<p style="font-size: 12px; color: #64748B; margin-top: 20px; line-height: 1.5;">The complete itemized operational manifest is attached as an official PDF document to this email.<br>For questions or dispatch discrepancies, contact logistics operations.</p>' +
    '</div>' +
    '<div style="background: #F1F5F9; padding: 14px; text-align: center; font-size: 11px; color: #94A3B8; border-top: 1px solid #E2E8F0;">' +
      COMPANY_NAME + ' • Automated Daily Dispatch Gateway' +
    '</div>' +
  '</div>';
}

function escapeXml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
