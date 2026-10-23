/**
 * Smits letzte Mission – gemeinsame Bestenliste in einer Google-Tabelle.
 *
 * Einrichtung (einmalig):
 *  1. Neue Google-Tabelle anlegen (z. B. "Smits Abschiedsspiel – Bestenliste").
 *  2. Erweiterungen → Apps Script. Den vorhandenen Code löschen und diesen Code einfügen. Speichern.
 *  3. Bereitstellen → Neue Bereitstellung → Typ "Web-App".
 *       Ausführen als: Ich
 *       Zugriff: Jeder
 *     → Bereitstellen, Zugriff erlauben, die Web-App-URL (endet auf /exec) kopieren.
 *  4. Diese URL im Spiel bei LEADERBOARD_CONFIG.sheetUrl eintragen.
 *
 * Die Ergebnisse erscheinen im Tabellenblatt "Bestenliste". Zeilen dort zu löschen
 * entfernt sie auch aus dem Leaderboard im Spiel.
 * Abschiedsworte für Smit landen im Tabellenblatt "Abschiedsworte". Sie werden nie
 * an das Spiel zurückgegeben, nur du siehst sie in der Tabelle.
 *
 * Nach einer Änderung an diesem Code: Bereitstellen → Bereitstellungen verwalten →
 * Stift-Symbol → Version: "Neue Version" → Bereitstellen (die URL bleibt gleich).
 */

const SHEET_NAME = 'Bestenliste';
const TEAMS = {
  triathlon: 'Triathlon',
  wiwisb: 'WiWiSB',
  career: 'Career Center',
  kontor: 'Service Kontor Andreas Müller',
};
const MAX_SCORE = 5000;
const NOTES_SHEET_NAME = 'Abschiedsworte';
const MAX_NOTE_LENGTH = 500;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Zeitpunkt', 'Name', 'Bereich', 'Punkte']);
    sh.setFrozenRows(1);
    sh.getRange('A1:D1').setFontWeight('bold');
  }
  return sh;
}

function getNotesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOTES_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(NOTES_SHEET_NAME);
    sh.appendRow(['Zeitpunkt', 'Name', 'Bereich', 'Punkte', 'Abschiedsworte']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
    sh.setColumnWidth(5, 480);
    sh.getRange('E:E').setWrap(true);
  }
  return sh;
}

// Eingaben nie als Formel ausführen
function safeText_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function readScores_() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, 4).getValues()
    .map(function (r, i) {
      return {
        id: 'r' + (i + 2),
        createdAt: r[0] instanceof Date ? r[0].getTime() : Number(r[0]) || 0,
        name: String(r[1]).replace(/^'/, '').trim().slice(0, 20),
        team: String(r[2]),
        score: Number(r[3]) || 0,
      };
    })
    .filter(function (e) { return e.name; });
}

// Liest die Bestenliste
function doGet() {
  return json_({ ok: true, scores: readScores_() });
}

// Trägt ein Ergebnis oder Abschiedsworte ein
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let name = String(d.name || '').trim().slice(0, 20);
    const score = Math.round(Number(d.score));
    const team = TEAMS[d.team] || '';

    if (d.type === 'note') {
      const text = String(d.text || '').trim().slice(0, MAX_NOTE_LENGTH);
      if (!name || !text) return json_({ ok: false, error: 'invalid' });
      getNotesSheet_().appendRow([new Date(), safeText_(name), team, isFinite(score) ? score : '', safeText_(text)]);
      return json_({ ok: true });
    }

    if (!name || !isFinite(score) || score < 0 || score > MAX_SCORE) {
      return json_({ ok: false, error: 'invalid' });
    }
    const sh = getSheet_();
    sh.appendRow([new Date(), safeText_(name), team, score]);
    return json_({ ok: true, id: 'r' + sh.getLastRow() });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
