// ================================================
// Warehouse Pro — Google Apps Script Backend
// Paste this entire file into Google Apps Script
// ================================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const action = e.parameter.action || (e.postData ? JSON.parse(e.postData.contents).action : null);
    const body = e.postData ? JSON.parse(e.postData.contents) : {};
    let result = {};
    switch(action) {
      case 'load':  result = loadAll(); break;
      case 'save':  result = saveAll(body.data); break;
      case 'ping':  result = {ok:true, msg:'Connected!'}; break;
      default:      result = {ok:false, msg:'Unknown action: '+action};
    }
    output.setContent(JSON.stringify(result));
  } catch(err) {
    output.setContent(JSON.stringify({ok:false, msg:err.toString()}));
  }
  return output;
}

const SHEETS = {
  products:   '📦 Products',
  inbound:    '📥 Inbound',
  outbound:   '📤 Outbound',
  transfers:  '🔄 Transfers',
  invoices:   '🧾 Invoices',
  adjustments:'📝 Adjustments',
  settings:   '⚙ Settings',
  company:    '⚙ Company'
};

const HEADERS = {
  products:    ['id','sku','name','cat','unit','price','opening','min','branch','supplier','notes'],
  inbound:     ['id','no','date','prodId','prodName','branch','qty','price','supplier','notes'],
  outbound:    ['id','no','date','prodId','prodName','branch','qty','price','dept','notes'],
  transfers:   ['id','no','date','prodId','prodName','qty','fromBranch','toBranch','notes'],
  invoices:    ['id','no','date','due','type','status','buyer','babn','bcon','btel','baddr','discType','discVal','discAmt','sub','gst','total','gstRate','terms','notes','itemsJson'],
  adjustments: ['id','no','date','prodId','prodName','type','qty','reason','branch'],
  settings:    ['key','value'],
  company:     ['key','value']
};

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#0f172a')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToArray(sheet, headers) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(row => row[0] !== '' && row[0] !== null).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const val = row[i];
      if (val === null || val === undefined || val === '') {
        obj[h] = '';
      } else if (val instanceof Date) {
        const y = val.getFullYear();
        const mo = String(val.getMonth()+1).padStart(2,'0');
        const d = String(val.getDate()).padStart(2,'0');
        const hr = String(val.getHours()).padStart(2,'0');
        const mn = String(val.getMinutes()).padStart(2,'0');
        obj[h] = (hr==='00'&&mn==='00') ? `${y}-${mo}-${d}` : `${y}-${mo}-${d} ${hr}:${mn}`;
      } else if (typeof val === 'number') {
        obj[h] = Number.isInteger(val) ? String(val) : String(Math.round(val*10000)/10000);
      } else {
        obj[h] = String(val);
      }
    });
    return obj;
  });
}

function arrayToSheet(sheet, headers, rows) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow-1, headers.length).clearContent();
  if (!rows || rows.length === 0) return;
  const values = rows.map(row => headers.map(h => row[h] !== undefined ? row[h] : ''));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function loadAll() {
  const result = {ok: true};
  Object.keys(SHEETS).forEach(key => {
    const sheet = getOrCreateSheet(SHEETS[key], HEADERS[key]);
    if (key === 'settings' || key === 'company') {
      const rows = sheetToArray(sheet, HEADERS[key]);
      const obj = {};
      rows.forEach(r => { if(r.key) obj[r.key] = r.value; });
      result[key] = obj;
    } else {
      result[key] = sheetToArray(sheet, HEADERS[key]);
    }
  });
  return result;
}

function saveAll(data) {
  if (!data) return {ok:false, msg:'No data received'};
  Object.keys(SHEETS).forEach(key => {
    if (!data[key]) return;
    const sheet = getOrCreateSheet(SHEETS[key], HEADERS[key]);
    if (key === 'settings' || key === 'company') {
      const rows = Object.keys(data[key]).map(k => ({key:k, value:data[key][k]}));
      arrayToSheet(sheet, HEADERS[key], rows);
    } else {
      arrayToSheet(sheet, HEADERS[key], data[key]);
    }
  });
  return {ok:true, msg:'Saved', timestamp:new Date().toISOString()};
}

function setup() {
  Object.keys(SHEETS).forEach(key => getOrCreateSheet(SHEETS[key], HEADERS[key]));
  SpreadsheetApp.getUi().alert('Setup complete!');
}
