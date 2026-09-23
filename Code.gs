/**
 * Backend for "ระบบลาและจัดสอนแทน - โรงเรียนบ้านป่าฝาง"
 * Deploy this as a Google Apps Script Web App.
 * It uses a Google Sheet as the database, with two sheets: "Leaves" and "Substitutions".
 *
 * SETUP:
 * 1. Create a new Google Sheet (any name).
 * 2. Extensions > Apps Script, paste this whole file in, replacing any starter code.
 * 3. Run `setupSheets` once (select it in the function dropdown, click ▶ Run) to create
 *    the "Leaves" and "Substitutions" tabs with the right headers.
 * 4. Set an admin PIN: Project Settings (gear icon) > Script Properties > Add property
 *    key "ADMIN_PIN" value = whatever PIN you want (e.g. 2569). This PIN is never stored
 *    in the public GitHub repo — only here in your own Apps Script project.
 * 5. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    Click Deploy, authorize the requested permissions, and copy the Web App URL —
 *    it ends in /exec. Paste that URL into SCRIPT_URL near the top of index.html.
 * 6. Whenever you edit this file, use Deploy > Manage deployments > ✎ > New version,
 *    otherwise the live URL keeps serving the old code.
 */

const LEAVES_SHEET = 'Leaves';
const SUBS_SHEET = 'Substitutions';
const LEAVES_HEADERS = ['id','teacherName','position','leaveType','dateFrom','dateTo','reason','contactPlace','contactPhone','createdAt'];
const SUBS_HEADERS = ['leaveId','rowsJson','updatedAt'];

function setupSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[LEAVES_SHEET, LEAVES_HEADERS], [SUBS_SHEET, SUBS_HEADERS]].forEach(([name, headers])=>{
    let sh = ss.getSheetByName(name);
    if(!sh){ sh = ss.insertSheet(name); }
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  });
  // remove the default "Sheet1" if it's still empty and unused
  const def = ss.getSheetByName('Sheet1');
  if(def && ss.getSheets().length > 2) ss.deleteSheet(def);
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheetToObjects(sh){
  const values = sh.getDataRange().getValues();
  if(values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(row=>row[0]!=="").map(row=>{
    const obj = {};
    headers.forEach((h,i)=> obj[h]=row[i]);
    return obj;
  });
}

function findRowById(sh, idCol, id){
  const values = sh.getDataRange().getValues();
  for(let r=1; r<values.length; r++){
    if(String(values[r][idCol]) === String(id)) return r+1; // 1-indexed sheet row
  }
  return -1;
}

function doGet(e){
  try{
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if(action === 'listLeaves'){
      const sh = ss.getSheetByName(LEAVES_SHEET);
      const leaves = sheetToObjects(sh).map(l=>{
        // dates come back as Date objects from the sheet; normalize to YYYY-MM-DD strings
        ['dateFrom','dateTo'].forEach(k=>{
          if(l[k] instanceof Date) l[k] = Utilities.formatDate(l[k], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        });
        if(l.createdAt instanceof Date) l.createdAt = l.createdAt.toISOString();
        return l;
      });
      leaves.sort((a,b)=> (b.createdAt||'').localeCompare(a.createdAt||''));
      return jsonOut({ leaves });
    }

    if(action === 'getAssignments'){
      const sh = ss.getSheetByName(SUBS_SHEET);
      const row = findRowById(sh, 0, e.parameter.leaveId);
      if(row < 0) return jsonOut({ rows: null });
      const rowsJson = sh.getRange(row, 2).getValue();
      return jsonOut({ rows: rowsJson ? JSON.parse(rowsJson) : null });
    }

    if(action === 'saveLeave'){
      const data = JSON.parse(e.parameter.data);
      const sh = ss.getSheetByName(LEAVES_SHEET);
      const existingRow = findRowById(sh, 0, data.id);
      const rowValues = LEAVES_HEADERS.map(h=> data[h]!==undefined ? data[h] : '');
      if(existingRow > 0){
        sh.getRange(existingRow,1,1,rowValues.length).setValues([rowValues]);
      } else {
        sh.appendRow(rowValues);
      }
      return jsonOut({ ok:true });
    }

    if(action === 'saveAssignments'){
      const leaveId = e.parameter.leaveId;
      const rows = e.parameter.data; // already a JSON string
      const sh = ss.getSheetByName(SUBS_SHEET);
      const existingRow = findRowById(sh, 0, leaveId);
      const rowValues = [leaveId, rows, new Date().toISOString()];
      if(existingRow > 0){
        sh.getRange(existingRow,1,1,rowValues.length).setValues([rowValues]);
      } else {
        sh.appendRow(rowValues);
      }
      return jsonOut({ ok:true });
    }

    if(action === 'checkAdmin'){
      const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
      return jsonOut({ isAdmin: !!pin && String(e.parameter.pin) === String(pin) });
    }

    return jsonOut({ error: 'unknown action' });
  }catch(err){
    return jsonOut({ error: String(err) });
  }
}

// Writes also go through doGet above (simplest, avoids CORS preflight issues when
// calling from a different domain like GitHub Pages). doPost is kept as a harmless
// alias in case anything ever calls it directly.
function doPost(e){ return doGet(e); }
