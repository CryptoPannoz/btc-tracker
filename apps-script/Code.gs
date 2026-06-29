/**
 * BTC Tracker — backend Google Apps Script
 * =========================================
 * Fa due cose:
 *   1) setup()  → crea/ripulisce il foglio "REGISTRO" (fonte unica e pulita)
 *                 leggendo i dati dal vecchio foglio "TX Ledger".
 *   2) Web App  → API JSONP/POST usata dalla dashboard per leggere e scrivere.
 *
 * ISTRUZIONI RAPIDE
 * -----------------
 *  A) Apri il tuo Google Sheet → Estensioni → Apps Script.
 *  B) Incolla questo file (sostituisci tutto), poi imposta il TOKEN qui sotto.
 *  C) Esegui una volta la funzione setup()  (autorizza quando richiesto).
 *  D) Deploy → Nuovo deployment → tipo "App web":
 *        - Esegui come:  Me (tuo account)
 *        - Chi ha accesso:  Chiunque
 *     Copia l'URL che finisce con /exec.
 *  E) Nella dashboard premi "☁ Collega…" e incolla URL + TOKEN.
 */

// >>>>>> CAMBIA QUESTO con una password tua a piacere (deve combaciare con la dashboard) <<<<<<
var TOKEN = 'CAMBIA-QUESTO-TOKEN';

var SHEET = 'REGISTRO';
var SOURCE = 'TX Ledger';   // foglio di origine da cui costruire il REGISTRO
var HEADERS = ['DATA', 'WALLET', 'BTC', 'PREZZO_USD', 'COSTO_USD', 'SATS'];

/** Crea/ripulisce il foglio REGISTRO partendo da TX Ledger. Esegui una volta. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SOURCE);
  var rows = [];
  if (src) {
    var vals = src.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {            // salta intestazione
      var r = vals[i];
      if (r[0] === '' || r[2] === '') continue;        // serve data + amount
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
    // colonne calcolate
    var f = [];
    for (var k = 0; k < rows.length; k++) {
      var rn = k + 2;
      f.push(['=C' + rn + '*D' + rn, '=ROUND(C' + rn + '*100000000,0)']);
    }
    sh.getRange(2, 5, f.length, 2).setFormulas(f);
  }
  sh.getRange('A:A').setNumberFormat('@');            // data come testo
  sh.getRange('C:C').setNumberFormat('0.00000000');
  sh.getRange('D:E').setNumberFormat('$#,##0.00');
  sh.autoResizeColumns(1, HEADERS.length);
  SpreadsheetApp.getUi().alert('REGISTRO creato: ' + rows.length + ' acquisti importati.');
}

/** GET → lista acquisti. Supporta JSONP (?callback=fn). */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.token !== TOKEN) throw new Error('token non valido');
    out = { ok: true, purchases: readRows_(), updated: new Date().toISOString() };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  var body = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

/** POST → aggiunge un acquisto. Body JSON: {action:'add', token, purchase:{date,ledger,amount,price}} */
function doPost(e) {
  var out;
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) throw new Error('token non valido');
    if (data.action === 'add') {
      var p = data.purchase;
      var sh = sheet_();
      var rn = sh.getLastRow() + 1;
      sh.getRange(rn, 1, 1, 4).setValues([[String(p.date), p.ledger, num_(p.amount), num_(p.price)]]);
      sh.getRange(rn, 5, 1, 2).setFormulas([['=C' + rn + '*D' + rn, '=ROUND(C' + rn + '*100000000,0)']]);
      sh.getRange(rn, 1).setNumberFormat('@');
      out = { ok: true, added: rn - 1 };
    } else {
      throw new Error('azione sconosciuta');
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- helper ----------
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); }
  return sh;
}
function readRows_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 4).getValues();
  return v.filter(function (r) { return r[0] !== '' && r[2] !== ''; })
          .map(function (r) {
            return { date: fmtDate_(r[0]), ledger: r[1], amount: num_(r[2]), price: num_(r[3]) };
          });
}
function num_(x) {
  if (typeof x === 'number') return x;
  return parseFloat(String(x).replace(/[$€\s,]/g, '').replace(',', '.')) || 0;
}
function fmtDate_(d) {
  if (d instanceof Date) {
    var p = function (n) { return ('0' + n).slice(-2); };
    return p(d.getDate()) + '/' + p(d.getMonth() + 1) + '/' + d.getFullYear();
  }
  return String(d);
}
