// Google Apps Script — Production Log
// Paste this into your Apps Script editor (Extensions > Apps Script)
// Then deploy as Web App (Execute as: Me, Access: Anyone)

const SS = SpreadsheetApp.getActiveSpreadsheet();
const LOG_SHEET = 'Log';
const CONFIG_SHEET = 'Config';

// ── GET: return config data to the HTML form ──
function doGet(e) {
  try {
    const config = getConfig();
    return ContentService.createTextOutput(JSON.stringify(config))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST: handle form submissions and config additions ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'submit') {
      return handleSubmit(data);
    } else if (data.action === 'add_config') {
      return handleAddConfig(data);
    } else {
      return jsonResponse({ success: false, error: 'Unknown action: ' + data.action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── Submit a production log entry ──
function handleSubmit(data) {
  const sheet = SS.getSheetByName(LOG_SHEET) || SS.insertSheet(LOG_SHEET);
  const id = sheet.getLastRow(); // row number as entry ID

  const startMin = timeToMinutes(data.start_time);
  const endMin = timeToMinutes(data.end_time);
  const taskTime = endMin - startMin;
  const qty = parseInt(data.quantity) || 1;
  const timePer = (taskTime / qty).toFixed(2);

  sheet.appendRow([
    id,
    new Date(),
    data.name,
    data.date,
    data.item,
    data.operation,
    data.start_time,
    data.end_time,
    taskTime,
    data.quantity,
    timePer,
    data.issues,
    data.comments
  ]);

  return jsonResponse({
    success: true,
    id: id,
    calculated: {
      task_time_min: taskTime,
      time_per_part: timePer
    }
  });
}

// ── Add a new config entry (employee, item, or operation) ──
function handleAddConfig(data) {
  const sheet = SS.getSheetByName(CONFIG_SHEET);
  if (!sheet) return jsonResponse({ success: false, error: 'Config sheet not found' });

  const type = data.config_type;
  const value = (data.value || '').trim();
  if (!value) return jsonResponse({ success: false, error: 'Empty value' });

  if (type === 'employee') {
    const col = getConfigColumn(sheet, 'Employees');
    if (col === -1) return jsonResponse({ success: false, error: 'Employees column not found in Config sheet' });
    const existing = getColumnValues(sheet, col);
    if (existing.indexOf(value) === -1) {
      sheet.getRange(existing.length + 2, col).setValue(value); // +2: header + next empty row
    }
    return jsonResponse({ success: true, config_type: type, value: value });

  } else if (type === 'item') {
    // Add to Items column; also used by operations lookup
    const col = getConfigColumn(sheet, 'Items');
    if (col === -1) return jsonResponse({ success: false, error: 'Items column not found in Config sheet' });
    const existing = getColumnValues(sheet, col);
    if (existing.indexOf(value) === -1) {
      sheet.getRange(existing.length + 2, col).setValue(value);
    }
    return jsonResponse({ success: true, config_type: type, value: value });

  } else if (type === 'operation') {
    // Operations are stored as two paired columns: Items | Operations
    const itemCol = getConfigColumn(sheet, 'Op_Item');
    const opCol = getConfigColumn(sheet, 'Op_Operation');
    if (itemCol === -1 || opCol === -1) return jsonResponse({ success: false, error: 'Op_Item / Op_Operation columns not found in Config sheet' });

    const itemName = (data.item || '').trim();
    if (!itemName) return jsonResponse({ success: false, error: 'Item name required for operation' });

    // Check for duplicates
    const items = getColumnValues(sheet, itemCol);
    const ops = getColumnValues(sheet, opCol);
    for (let i = 0; i < items.length; i++) {
      if (items[i] === itemName && ops[i] === value) {
        return jsonResponse({ success: true, config_type: type, value: value, note: 'Already exists' });
      }
    }

    const nextRow = Math.max(items.length, ops.length) + 2;
    sheet.getRange(nextRow, itemCol).setValue(itemName);
    sheet.getRange(nextRow, opCol).setValue(value);
    return jsonResponse({ success: true, config_type: type, value: value, item: itemName });

  } else {
    return jsonResponse({ success: false, error: 'Unknown config_type: ' + type });
  }
}

// ── Build config object from Config sheet ──
function getConfig() {
  const sheet = SS.getSheetByName(CONFIG_SHEET);
  if (!sheet) throw new Error('Config sheet not found');

  const employees = getColumnValues(sheet, getConfigColumn(sheet, 'Employees'));
  const issueTypes = getColumnValues(sheet, getConfigColumn(sheet, 'Issues'));
  const items = getColumnValues(sheet, getConfigColumn(sheet, 'Items'));

  // Build operations list from paired columns
  const opItemCol = getConfigColumn(sheet, 'Op_Item');
  const opOpCol = getConfigColumn(sheet, 'Op_Operation');
  const operations = [];
  if (opItemCol !== -1 && opOpCol !== -1) {
    const opItems = getColumnValues(sheet, opItemCol);
    const opOps = getColumnValues(sheet, opOpCol);
    for (let i = 0; i < opItems.length; i++) {
      if (opItems[i] && opOps[i]) {
        operations.push({ item: opItems[i], operation: opOps[i] });
      }
    }
  }

  return { employees, operations, issue_types: issueTypes };
}

// ── Helpers ──
function getConfigColumn(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === headerName) return i + 1;
  }
  return -1;
}

function getColumnValues(sheet, col) {
  if (col === -1 || col === undefined) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, col, lastRow - 1, 1).getValues()
    .map(r => r[0].toString().trim())
    .filter(v => v !== '');
}

function timeToMinutes(t) {
  const parts = t.split(':');
  return parseInt(parts[0]) * 60 + parseInt(parts[1]);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
