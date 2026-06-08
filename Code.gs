// ============================================================
// SISTEM MANAJEMEN BARANG - Google Apps Script Backend
// Deploy as: Web App → Execute as Me → Anyone can access
// ============================================================

const SPREADSHEET_ID = '1V2Ufn6jG2oxuHDOy_9nCkcDekgTtNinyIG_I7FYcxRc'; // PASTE YOUR SPREADSHEET ID HERE
const SHEET_NAMES = {
  users: 'users',
  itemNames: 'itemNames',
  incomingGoods: 'incomingGoods',
  outgoingGoods: 'outgoingGoods'
};

// Column definitions per sheet
const COLUMNS = {
  users: ['id', 'username', 'password', 'role', 'createdAt'],
  itemNames: ['id', 'owner', 'itemName', 'itemCode', 'createdAt', 'createdBy'],
  incomingGoods: ['id', 'date', 'owner', 'itemName', 'itemCode', 'batchLot', 'expiredDate',
    'quantity', 'unit', 'condition', 'location', 'supplier', 'officer', 'notes',
    'status', 'createdAt', 'createdBy', 'approvedBy', 'approvedAt', 'updatedAt', 'updatedBy'],
  outgoingGoods: ['id', 'shipmentDate', 'owner', 'itemName', 'itemCode', 'batchLot',
    'deliveryNo', 'poNo', 'quantityOut', 'unit', 'destination', 'sendingOfficer',
    'salesName', 'shipmentStatus', 'notes', 'status', 'createdAt', 'createdBy',
    'updatedAt', 'updatedBy']
};

// ── Entry point ──────────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };

  try {
    const params = e.parameter || {};
let postData = {};
try {
  postData = e.postData
    ? JSON.parse(e.postData.contents || '{}')
    : {};
} catch(err) {
  postData = {};
}
const requestData =
  params.data
    ? JSON.parse(params.data)
    : postData.data || {};
	
    const action = params.action || postData.action;
    const sheet  = params.sheet  || postData.sheet;

    ensureSheetsExist();

    let result;
    switch (action) {
      case 'getAll':    result = getAll(sheet);               break;
      case 'create':    result = create(sheet, requestData); break;
      case 'update':    result = update(sheet, requestData); break;
      case 'delete':    result = deleteRow(sheet, params.id || postData.id);break;
	  case 'login':
  result = login(
    params.username || postData.username,
    params.password || postData.password
  );
  break;
      case 'seed':      result = seedDefaults();               break;
      default:          result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Sheet helpers ────────────────────────────────────────────
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheetsExist() {
  const ss = getSpreadsheet();
  for (const [key, name] of Object.entries(SHEET_NAMES)) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, COLUMNS[key].length).setValues([COLUMNS[key]]);
      sh.getRange(1, 1, 1, COLUMNS[key].length)
        .setBackground('#2c3e50').setFontColor('#ffffff').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet, colDef) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map((row, idx) => {
    const obj = { __rowIndex: idx + 2 };
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── CRUD ─────────────────────────────────────────────────────
function getAll(sheetName) {
  const sh = getSheet(sheetName);
  if (!sh) return { success: false, error: 'Sheet not found: ' + sheetName };
  const rows = sheetToObjects(sh, COLUMNS[sheetName]);
  return { success: true, data: rows };
}

	function create(sheetName, data) {
	  const sh = getSheet(sheetName);
	  if (!sh) return { success: false, error: 'Sheet not found: ' + sheetName };
	  const cols = COLUMNS[sheetName];
	  const id = Utilities.getUuid();
	  data.id = id;
	  const row = cols.map(c => data[c] !== undefined ? data[c] : '');
	  sh.appendRow(row);
	  return { success: true, id };
	}

function update(sheetName, data) {
  const sh = getSheet(sheetName);
  if (!sh) return { success: false, error: 'Sheet not found: ' + sheetName };
  const cols   = COLUMNS[sheetName];
  const allRows = sh.getDataRange().getValues();
  const idCol  = cols.indexOf('id');
  for (let i = 1; i < allRows.length; i++) {
    if (allRows[i][idCol] === data.id) {
      const newRow = cols.map(c => data[c] !== undefined ? data[c] : allRows[i][cols.indexOf(c)]);
      sh.getRange(i + 1, 1, 1, cols.length).setValues([newRow]);
      return { success: true };
    }
  }
  return { success: false, error: 'Row not found for id: ' + data.id };
}

function deleteRow(sheetName, id) {
  const sh = getSheet(sheetName);
  if (!sh) return { success: false, error: 'Sheet not found: ' + sheetName };
  const cols   = COLUMNS[sheetName];
  const idCol  = cols.indexOf('id');
  const allRows = sh.getDataRange().getValues();
  for (let i = 1; i < allRows.length; i++) {
    if (allRows[i][idCol] === id) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Row not found' };
}

// ── Auth ─────────────────────────────────────────────────────
function login(username, password) {
  const sh = getSheet(SHEET_NAMES.users);
  if (!sh) return { success: false, error: 'Users sheet not found' };
  const rows = sheetToObjects(sh, COLUMNS.users);
  const user = rows.find(u => u.username === username && u.password === password);
  if (user) {
    return { success: true, user: { id: user.id, username: user.username, role: user.role } };
  }
  return { success: false, error: 'Invalid credentials' };
}

// ── Seed default data ────────────────────────────────────────
function seedDefaults() {
  ensureSheetsExist();
  const sh = getSheet(SHEET_NAMES.users);
  const existing = sheetToObjects(sh, COLUMNS.users);

  const defaults = [
    { username: 'admin',     password: 'admin123',     role: 'admin'     },
    { username: 'manager',   password: 'manager123',   role: 'manager'   },
    { username: 'warehouse', password: 'warehouse123', role: 'warehouse' },
    { username: 'guest',     password: 'guest123',     role: 'guest'     }
  ];

  let added = 0;
  for (const u of defaults) {
    if (!existing.some(e => e.username === u.username)) {
      create(SHEET_NAMES.users, { ...u, createdAt: new Date().toISOString() });
      added++;
    }
  }

  // Seed sample item names
  const itemSh = getSheet(SHEET_NAMES.itemNames);
  const existingItems = sheetToObjects(itemSh, COLUMNS.itemNames);
  const sampleItems = [
    { owner: 'PT. Rena Haniem Mulia',     itemName: 'Catly 318',                    itemCode: 'CTL318'  },
    { owner: 'PT. Rena Haniem Mulia',     itemName: 'Dmast 2101',                   itemCode: 'DMS2101' },
    { owner: 'PT. Rena Haniem Mulia',     itemName: 'Spraysphere SC-GR-VR-1401-S',  itemCode: 'SPR1401' },
    { owner: 'PT. Sumber Phoenix Makmur', itemName: 'Plexiglas',                    itemCode: 'PLX001'  },
    { owner: 'PT. Sumber Phoenix Makmur', itemName: 'Powercast',                    itemCode: 'PWC001'  },
    { owner: 'PT. Monca Wijaya Makmur',   itemName: 'Matcha Powder',                itemCode: 'MTP001'  },
    { owner: 'PT. Monca Wijaya Makmur',   itemName: 'Egg yolk powder',              itemCode: 'EYP001'  }
  ];
  for (const item of sampleItems) {
    if (!existingItems.some(e => e.owner === item.owner && e.itemName === item.itemName)) {
      create(SHEET_NAMES.itemNames, { ...item, createdAt: new Date().toISOString(), createdBy: 'system' });
    }
  }

  return { success: true, message: `Seeded ${added} users and sample items` };
}
