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
 * entfernt sie auch aus der Bestenliste im Spiel.
 */

const SHEET_NAME = 'Bestenliste';
const TEAMS = {
  triathlon: 'Triathlon',
  wiwisb: 'WiWiSB',
  career: 'Career Center',
  kontor: 'Service Kontor Andreas Müller',
};
const MAX_SCORE = 5000;

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

// Trägt ein Ergebnis ein
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let name = String(d.name || '').trim().slice(0, 20);
    const score = Math.round(Number(d.score));
    const team = TEAMS[d.team] || '';
    if (!name || !isFinite(score) || score < 0 || score > MAX_SCORE) {
      return json_({ ok: false, error: 'invalid' });
    }
    // Schutz: Eingaben nie als Formel ausführen
    if (/^[=+\-@]/.test(name)) name = "'" + name;
    const sh = getSheet_();
    sh.appendRow([new Date(), name, team, score]);
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
