/**
 * Lernstand-Check · Ablage in Google Sheets
 * BBW Winterthur · AFP 5. Semester
 *
 * 1. In Google Sheets: Erweiterungen → Apps Script → diesen Code einfügen.
 * 2. Unten bei SCHLUESSEL ein eigenes Passwort setzen (nur für das Dashboard).
 * 3. Bereitstellen → Neue Bereitstellung → Web-App
 *    Ausführen als: Ich · Zugriff: Jeder → Bereitstellen → URL kopieren.
 */
const SCHLUESSEL = 'HIER-EIGENES-PASSWORT';

const KOEPFE = {
  Lernstand: ['ID', 'Zeit', 'Datum', 'Klasse', 'Kürzel', 'Spur', 'LZ', 'Thema', 'K', 'Ziel', 'Note', 'Hilfe'],
  Feedback:  ['ID', 'Datum', 'Klasse', 'Kürzel', 'Spur', 'Thema', 'Erklärungen', 'Unterlagen', 'Praxisbezug', 'Abwechslung', 'Lernklima', 'Tempo'],
  Texte:     ['ID', 'Datum', 'Klasse', 'Kürzel', 'Top', 'Kritik', 'Wunsch']
};

function blatt_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(KOEPFE[name]); s.setFrozenRows(1); }
  return s;
}

/* Text säubern: Länge begrenzen, keine Formeln einschleusen */
function t_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max || 200).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function n_(v) { const x = Number(v); return (x >= 1 && x <= 6) ? Math.round(x * 2) / 2 : ''; }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = t_(d.id, 40);
    const cache = CacheService.getScriptCache();
    if (!id || cache.get(d.type + id)) return antwort_({ ok: true, doppelt: true });

    if (d.type === 'lern') {
      const rows = (d.rows || []).slice(0, 20).map(r => [
        id, t_(d.zeit, 30), t_(d.datum, 12), t_(d.klasse, 20), t_(d.kz, 4),
        t_(r[0], 2), t_(r[1], 10), t_(r[2], 60), t_(r[3], 2), t_(r[4], 200), n_(r[5]), t_(d.hilfe, 300)]);
      if (rows.length) blatt_('Lernstand').getRange(blatt_('Lernstand').getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } else if (d.type === 'fb') {
      const kz = d.anon ? 'anonym' : t_(d.kz, 4);
      const rows = (d.rows || []).slice(0, 4).map(r => [
        id, t_(d.datum, 12), t_(d.klasse, 20), kz, t_(r[0], 2), t_(r[1], 60),
        n_(r[2]), n_(r[3]), n_(r[4]), n_(r[5]), n_(r[6]), t_(r[7], 20)]);
      if (rows.length) blatt_('Feedback').getRange(blatt_('Feedback').getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      if (d.top || d.kritik || d.wunsch)
        blatt_('Texte').appendRow([id, t_(d.datum, 12), t_(d.klasse, 20), kz, t_(d.top, 1000), t_(d.kritik, 1000), t_(d.wunsch, 1000)]);
    } else {
      return antwort_({ ok: false });
    }
    cache.put(d.type + id, '1', 21600);
    return antwort_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* Dashboard holt die Daten nur mit dem richtigen Schlüssel */
function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== SCHLUESSEL) return antwort_({ ok: false, fehler: 'Schlüssel falsch' });
  const out = { ok: true };
  Object.keys(KOEPFE).forEach(name => {
    const tz = Session.getScriptTimeZone();
    const v = blatt_(name).getDataRange().getValues();
    const kopf = v.shift();
    out[name] = v.map(r => Object.fromEntries(kopf.map((k, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, tz, k === 'Zeit' ? "yyyy-MM-dd'T'HH:mm" : 'yyyy-MM-dd');
      return [k, x];
    })));
  });
  return antwort_(out);
}

function antwort_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* Einmal von Hand ausführen: legt die drei Tabellenblätter an */
function einrichten() { Object.keys(KOEPFE).forEach(blatt_); }
