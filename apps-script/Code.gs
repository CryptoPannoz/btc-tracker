/**
 * BTC Tracker — backend Google Apps Script (v2)
 * =============================================
 *   1) setup()  → crea/ripulisce il foglio "REGISTRO" dal vecchio "TX Ledger".
 *   2) Web App  → API usata dalla dashboard: list (GET/JSONP), add/update/delete (POST).
 *
 * ⚠️ ATTENZIONE: NON rieseguire setup() se hai già aggiunto/modificato acquisti dalla
 *    dashboard: setup() RICOSTRUISCE REGISTRO da TX Ledger e perderesti quelle righe.
 *    La colonna FEES e le nuove azioni funzionano senza rieseguire setup(): basta
 *    aggiornare questo codice e rifare il Deploy (Gestisci deployment → nuova versione).
 *
 * DEPLOY: Deploy → App web → Esegui come: Me · Chi ha accesso: Chiunque → copia URL /exec.
 */

// >>>>>> la tua password segreta (uguale a quella nella dashboard) <<<<<<
var TOKEN = 'CAMBIA-QUESTO-TOKEN';

var SHEET = 'REGISTRO';
var SOURCE = 'TX Ledger';
var HEADERS = ['DATA', 'WALLET', 'BTC', 'PREZZO_USD', 'COSTO_USD', 'SATS', 'FEES'];
// indici colonna (1-based): A=DATA B=WALLET C=BTC D=PREZZO E=COSTO F=SATS G=FEES

/** Crea/ripulisce REGISTRO partendo da TX Ledger. Eseguire UNA SOLA volta all'inizio. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SOURCE);
  var rows = [];
  if (src) {
    var vals = src.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (r[0] === '' || r[2] === '') continue;
      rows.push([fmtDate_(r[0]), r[1], num_(r[2]), num_(r[3])]);
    }
  }
  var sh = ss.getSheetByName(SHEET);
  if (sh) sh.clear(); else sh = ss.insertSheet(SHEET);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#f7931a').setFontColor('#1a1203');
  sh.setFrozenRows(1);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 4).setValues(rows);
    setFormulas_(sh, 2, rows.length);
  }
  formatSheet_(sh);
  SpreadsheetApp.getUi().alert('REGISTRO creato: ' + rows.length + ' acquisti importati.');
}

/** GET → lista acquisti. Supporta JSONP (?callback=fn). */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.token !== TOKEN) throw new Error('token non valido');
    out = { ok: true, purchases: readRows_(), updated: new Date().toISOString() };
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  var body = JSON.stringify(out);
  if (p.callback) return ContentService.createTextOutput(p.callback + '(' + body + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST → add / update / delete.
 *   add:    {action:'add', token, purchase:{date,ledger,amount,price,fee}}
 *   update: {action:'update', token, match:{date,ledger,amount}, set:{date,ledger,amount,price,fee}}
 *   delete: {action:'delete', token, match:{date,ledger,amount}}
 */
function doPost(e) {
  var out;
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) throw new Error('token non valido');
    var sh = sheet_();
    if (data.action === 'add') {
      var p = data.purchase, rn = sh.getLastRow() + 1;
      writeRow_(sh, rn, p);
      out = { ok: true, added: rn - 1 };
    } else if (data.action === 'update') {
      var rowU = matchRow_(sh, data.match);
      if (rowU < 0) throw new Error('riga non trovata');
      writeRow_(sh, rowU, data.set);
      out = { ok: true, updated: rowU };
    } else if (data.action === 'delete') {
      var rowD = matchRow_(sh, data.match);
      if (rowD < 0) throw new Error('riga non trovata');
      sh.deleteRow(rowD);
      out = { ok: true, deleted: rowD };
    } else throw new Error('azione sconosciuta');
  } catch (err) { out = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- helper ----------
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); }
  if (sh.getRange(1, 7).getValue() !== 'FEES') sh.getRange(1, 7).setValue('FEES'); // assicura colonna FEES
  return sh;
}
function writeRow_(sh, rn, p) {
  sh.getRange(rn, 1, 1, 4).setValues([[String(p.date), p.ledger, num_(p.amount), num_(p.price)]]);
  sh.getRange(rn, 7).setValue(num_(p.fee || 0));
  sh.getRange(rn, 1).setNumberFormat('@');
  setFormulas_(sh, rn, 1);
}
function setFormulas_(sh, startRow, n) {
  var f = [];
  for (var k = 0; k < n; k++) { var rn = startRow + k; f.push(['=C' + rn + '*D' + rn, '=ROUND(C' + rn + '*100000000,0)']); }
  sh.getRange(startRow, 5, n, 2).setFormulas(f);
}
function matchRow_(sh, m) {
  var last = sh.getLastRow(); if (last < 2) return -1;
  var v = sh.getRange(2, 1, last - 1, 3).getValues(); // A,B,C
  for (var i = 0; i < v.length; i++) {
    if (fmtDate_(v[i][0]) === String(m.date) && v[i][1] === m.ledger && Math.abs(num_(v[i][2]) - num_(m.amount)) < 1e-8)
      return i + 2;
  }
  return -1;
}
function readRows_() {
  var sh = sheet_(); var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 7).getValues();
  return v.filter(function (r) { return r[0] !== '' && r[2] !== ''; })
    .map(function (r) { return { date: fmtDate_(r[0]), ledger: r[1], amount: num_(r[2]), price: num_(r[3]), fee: num_(r[6]) }; });
}
function formatSheet_(sh) {
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('C:C').setNumberFormat('0.00000000');
  sh.getRange('D:E').setNumberFormat('$#,##0.00');
  sh.getRange('G:G').setNumberFormat('€#,##0.00');
  sh.autoResizeColumns(1, HEADERS.length);
}
function num_(x) { if (typeof x === 'number') return x; return parseFloat(String(x).replace(/[$€\s,]/g, '').replace(',', '.')) || 0; }
function fmtDate_(d) {
  if (d instanceof Date) { var p = function (n) { return ('0' + n).slice(-2); }; return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear(); }
  return String(d);
}
