// ════════════════════════════════════════════════════════════
//  YOUNGS CAPITAL — Apps Script (Fresh Build)
//  Step 1: Paste this in Extensions → Apps Script → Code.gs
//  Step 2: Run setupSheets() once
//  Step 3: Deploy → New Deployment → Web App
//          Execute as: Me | Access: Anyone
//  Step 4: Copy Web App URL → paste in dashboard
// ════════════════════════════════════════════════════════════

var SS = SpreadsheetApp.getActiveSpreadsheet();

// ── Helpers ──────────────────────────────────────────────────
function ok(data)  { return out({ status:'ok',  data: data  }); }
function err(msg)  { return out({ status:'error', message: msg }); }
function out(obj)  {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  return SS.getSheetByName(name) || SS.insertSheet(name);
}

function headerStyle(range) {
  range.setBackground('#0d1f0d')
       .setFontColor('#d6ae52')
       .setFontWeight('bold');
}

// ── GET ───────────────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e.parameter.action || 'load';

    if (action === 'load') {
      // Report header
      var rVals = sheet('Report').getDataRange().getValues();
      var r = {};
      rVals.forEach(function(row){ if(row[0]) r[row[0]] = String(row[1]||''); });

      // Stocks
      var sVals = sheet('Stocks').getDataRange().getValues();
      var stocks = [];
      for (var i = 1; i < sVals.length; i++) {
        var s = sVals[i];
        if (!s[0]) continue;
        stocks.push({ symbol:s[0], strategy:s[1], close:String(s[2]||''), target1:String(s[3]||''), target2:String(s[4]||''), stoploss:String(s[5]||''), note:s[6], islamic: s[7]==='Yes' });
      }

      // Records
      var recVals = sheet('Records').getDataRange().getValues();
      var records = [];
      for (var j = 1; j < recVals.length; j++) {
        var rec = recVals[j];
        if (!rec[0]) continue;
        records.push({ date:rec[0], kseClosing:String(rec[1]||''), kseTarget1:String(rec[2]||''), kseTarget2:String(rec[3]||''), kseStoploss:String(rec[4]||''), symbol:rec[5], strategy:rec[6], stockClose:String(rec[7]||''), stockT1:String(rec[8]||''), stockT2:String(rec[9]||''), stockSL:String(rec[10]||''), note:rec[11] });
      }

      return ok({
        report: { date:r.date||'', ref:r.ref||'', kse:{ closing:r.closing||'', target1:r.target1||'', target2:r.target2||'', stoploss:r.stoploss||'' }, marketView:r.mv||'' },
        stocks: stocks,
        records: records
      });
    }

    return err('Unknown action');
  } catch(e) { return err(e.message); }
}

// ── POST ──────────────────────────────────────────────────────
// Receives: ?action=X  with body = JSON string of payload
function doPost(e) {
  try {
    var action  = e.parameter.action;
    var payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch(x) {}
    }

    Logger.log('doPost action=' + action + ' payload=' + JSON.stringify(payload));

    if (action === 'saveReport')  return saveReport(payload);
    if (action === 'saveStocks')  return saveStocks(payload);
    if (action === 'saveRecord')  return saveRecord(payload);

    return err('Unknown action: ' + action);
  } catch(e) { return err(e.message); }
}

function saveReport(p) {
  var sh = sheet('Report');
  sh.clearContents();
  var hdr = sh.getRange(1,1,1,2);
  hdr.setValues([['Field','Value']]);
  headerStyle(hdr);
  sh.getRange(2,1,7,2).setValues([
    ['date',     p.date||''],
    ['ref',      p.ref||''],
    ['closing',  p.closing||''],
    ['target1',  p.t1||''],
    ['target2',  p.t2||''],
    ['stoploss', p.sl||''],
    ['mv',       p.mv||'']
  ]);
  sh.autoResizeColumns(1,2);
  return ok({ message:'Report saved' });
}

function saveStocks(p) {
  var stocks = Array.isArray(p) ? p : [];
  var sh = sheet('Stocks');
  sh.clearContents();
  var hdr = sh.getRange(1,1,1,8);
  hdr.setValues([['Symbol','Strategy','Close','Target 1','Target 2','Stoploss','Note','Islamic']]);
  headerStyle(hdr);
  if (stocks.length) {
    var rows = stocks.map(function(s){ return [s.symbol,s.strategy,s.close,s.target1,s.target2,s.stoploss,s.note,s.islamic?'Yes':'No']; });
    sh.getRange(2,1,rows.length,8).setValues(rows);
    rows.forEach(function(row, i){
      var c = sh.getRange(i+2,2);
      var st = String(row[1]).toLowerCase();
      if (st.indexOf('buy')>-1)  c.setBackground('#d4edda').setFontColor('#155724');
      if (st.indexOf('sell')>-1) c.setBackground('#f8d7da').setFontColor('#721c24');
      if (st.indexOf('hold')>-1) c.setBackground('#fff3cd').setFontColor('#856404');
    });
  }
  sh.autoResizeColumns(1,8);
  return ok({ message:'Stocks saved' });
}

function saveRecord(p) {
  var sh = sheet('Records');
  if (sh.getLastRow() === 0) {
    var hdr = sh.getRange(1,1,1,12);
    hdr.setValues([['Date','KSE Close','KSE T1','KSE T2','KSE SL','Symbol','Strategy','Stock Close','Stock T1','Stock T2','Stock SL','Note']]);
    headerStyle(hdr);
  }
  // Remove existing rows for this date (upsert)
  var last = sh.getLastRow();
  if (last > 1) {
    var col = sh.getRange(2,1,last-1,1).getValues();
    for (var i = col.length-1; i >= 0; i--) {
      if (String(col[i][0]).trim() === String(p.date).trim()) sh.deleteRow(i+2);
    }
  }
  var kse = p.kse || {};
  sh.appendRow([p.date, kse.closing||'', kse.target1||'', kse.target2||'', kse.stoploss||'', 'KSE100','Index','','','','','']);
  (p.stocks||[]).forEach(function(s){
    sh.appendRow([p.date, kse.closing||'', kse.target1||'', kse.target2||'', kse.stoploss||'', s.symbol||'', s.strategy||'', s.close||'', s.target1||'', s.target2||'', s.stoploss||'', s.note||'']);
  });
  sh.autoResizeColumns(1,12);
  return ok({ message:'Record saved for ' + p.date });
}

// ── One-time setup ────────────────────────────────────────────
function setupSheets() {
  sheet('Report');
  sheet('Stocks');
  sheet('Records');
  saveReport({ date:'', ref:'', closing:'', t1:'', t2:'', sl:'', mv:'' });
  saveStocks([]);
  SpreadsheetApp.getUi().alert('✅ Done! Now deploy as Web App.\n\nDeploy → New Deployment → Web App\nExecute as: Me\nWho has access: Anyone\n\nCopy the URL into your dashboard.');
}
