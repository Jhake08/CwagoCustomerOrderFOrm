// Load the login page initially
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('LOGIN PAGE')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getOrderForm() {
  return HtmlService.createHtmlOutputFromFile('Index').getContent();
}

// Get all sheet names
function getSheetNames() {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  return ss.getSheets().map(sheet => sheet.getName());
}

// Get agent names from 'Agents' sheet or use fallback list
function getAgents() {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName('Agents');
  if (!sheet) return ['Harley', 'Gab', 'Pia', 'Mase', 'Benjie'];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(row => row[0]).filter(Boolean);
}

function getProductCombo() {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName('Remarks');
  if (!sheet) return [
    '6 PCS WOODEN PLATE=198',
    '12 PCS WOODEN PLATE=396',
    '1 PC. STORAGE BAG = 199',
    '2 PCS. STORAGE BAG = 398',
    '3 PCS. STORAGE BAG = 597',
    '6 PCS. STORAGE BAG  = 1,194',
    '4 REF MAT = 198',
    '2 PCS SEALANT TAPE=398'
  ];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(row => row[0]).filter(Boolean);
}

function getFacebookPages(sheetName) {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const pageNames = [];

  // Get background colors for column A only
  const bgColors = sheet.getRange(1, 1, lastRow, 1).getBackgrounds();

  for (let i = 0; i < lastRow; i++) {
    const color = bgColors[i][0];
    // Check if the cell in column A is green (assuming green is '#00ff00' or similar)
    if (color === '#00ff00' || color === '#00FF00' || color === '#008000' || color === '#008000ff') {
      const cellValue = sheet.getRange(i + 1, 1).getValue();
      if (cellValue) {
        pageNames.push(cellValue);
      }
    }
  }

  return pageNames;
}

// Add order data to selected sheet with encoding logic
function addDataToSheet(sheetName, facebookPageName, agent, customerName, address, phoneNo, productCombo) {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} not found`);

  const lastRow = sheet.getLastRow();
  const bgColors = sheet.getRange(1, 1, lastRow, 1).getBackgrounds();

  // Find the row of the selected Facebook Page Name by scanning for green-highlighted row in column A
  let targetRow = -1;
  for (let i = 0; i < lastRow; i++) {
    const color = bgColors[i][0];
    const cellValue = sheet.getRange(i + 1, 1).getValue();
    if ((color === '#00ff00' || color === '#00FF00' || color === '#008000' || color === '#008000ff') && cellValue === facebookPageName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error(`Facebook Page Name "${facebookPageName}" not found in green-highlighted rows.`);
  }

  // Insert a new row directly below the target row
  sheet.insertRowAfter(targetRow);

  // Extract price from productCombo by last '=' occurrence
  let price = 0;
  const lastEqualIndex = productCombo.lastIndexOf('=');
  if (lastEqualIndex !== -1) {
    const priceStr = productCombo.substring(lastEqualIndex + 1).trim().replace(/,/g, '');
    price = parseInt(priceStr, 10) || 0;
  }

  // Current time in 12-hour format with AM/PM
  const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "hh:mm:ss a");

  // Quantity is always 1
  const quantity = 1;

  // Prepare the row data according to mapping
  const rowData = [
    agent,          // Column A
    time,           // Column B
    customerName,   // Column C
    address,        // Column D
    phoneNo,        // Column E
    quantity,       // Column F
    price,          // Column G
    productCombo    // Column H
  ];

  // Set the values in the new row
  const newRowRange = sheet.getRange(targetRow + 1, 1, 1, rowData.length);
  newRowRange.setValues([rowData]);

  // Remove any background color (clear formatting) from the new row to avoid green highlight
  // Extend to clear entire row width (all columns) to ensure full row cleared
  const totalColumns = sheet.getMaxColumns();
  const newRowIndex = targetRow + 1;
  // Clear background for entire row (all columns)
  sheet.getRange(newRowIndex, 1, 1, totalColumns).setBackground(null);
}

// Add incomplete entry to backup sheet
function addDataToIncompleteSheet(agent, customerName, address, phoneNo, remarks) {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName('Incomplete Entries') || ss.insertSheet('Incomplete Entries');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Agent', 'Customer Name', 'Address', 'Phone No', 'Remarks', 'Time']);
  }

  const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
  sheet.appendRow([agent, customerName, address, phoneNo, remarks, time]);
}

// Delete last entry from a specific sheet
function deleteLastRecord(sheetName) {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet ${sheetName} not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No records to delete');

  sheet.deleteRow(lastRow);
}

function getSheetByName(sheetName, facebookPageName) {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues();
  const bgColors = sheet.getRange(1, 1, sheet.getLastRow(), 1).getBackgrounds();

  let totalQuantity = 0;
  let totalPrice = 0;

  // Find the start row of the chosen Facebook Page (green-highlighted row with matching name)
  let startRow = -1;
  for (let i = 0; i < data.length; i++) {
    const color = bgColors[i][0];
    const cellValue = data[i][0];
    if ((color === '#00ff00' || color === '#00FF00' || color === '#008000' || color === '#008000ff') && cellValue === facebookPageName) {
      startRow = i;
      break;
    }
  }
  if (startRow === -1) {
    throw new Error(`Facebook Page Name "${facebookPageName}" not found in green-highlighted rows.`);
  }

  // Find the next green-highlighted row after startRow or end of data
  let endRow = data.length;
  for (let i = startRow + 1; i < data.length; i++) {
    const color = bgColors[i][0];
    if (color === '#00ff00' || color === '#00FF00' || color === '#008000' || color === '#008000ff') {
      endRow = i;
      break;
    }
  }

  // Sum quantity and price for rows between startRow+1 and endRow-1
  for (let i = startRow + 1; i < endRow; i++) {
    totalQuantity += Number(data[i][5]) || 0; // Quantity in column F (index 5)
    totalPrice += Number(data[i][6]) || 0;    // Price in column G (index 6)
  }

  return {
    facebookPage: facebookPageName,
    totalQuantity: totalQuantity,
    totalPrice: totalPrice
  };
}

function generateReport(sheetName, facebookPageName) {
  return getSheetByName(sheetName, facebookPageName);
}

// Scan Facebook page sections in selected sheet tab
function scanFacebookPageSections(sheetName) {
  const ss = SpreadsheetApp.openById('100usuKPUFHFrx8UXvFoy6G-k-2vYHKVxWkzFfLlfVSs');
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const lastRow = sheet.getLastRow();
  const pageSections = [];

  // Get background colors for column A only
  const bgColors = sheet.getRange(1, 1, lastRow, 1).getBackgrounds();

  for (let i = 0; i < lastRow; i++) {
    const color = bgColors[i][0];
    // Check if the cell in column A is green (assuming green is '#00ff00' or similar)
    if (color === '#00ff00' || color === '#00FF00' || color === '#008000' || color === '#008000ff') {
      const pageName = sheet.getRange(i + 1, 1).getValue();
      if (pageName) {
        pageSections.push({ pageName: pageName, startRow: i + 1 });
      }
    }
  }

  return pageSections;
}
