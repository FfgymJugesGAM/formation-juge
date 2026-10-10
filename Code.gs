const SHEET_NAME = 'Inscriptions';
const SPREADSHEET_ID = '1f3rSByAS6W2JEsXyj66eIh1iS5v6mxeBQsD2JP36oaw';
const HEADERS = [
  'Date', 'Nom', 'Prénom', 'Agrès', 'Rôle',
  ...[1, 2, 3, 4].flatMap(m => [`Mouvement ${m} - A`, `Mouvement ${m} - B`, `Mouvement ${m} - C`, `Mouvement ${m} - D`, `Mouvement ${m} - E`, `Mouvement ${m} - F et +`, `Mouvement ${m} - Déductions`])
];
function doGet() {
  try {
    SpreadsheetApp.openById(SPREADSHEET_ID);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(String(err.stack || err));
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Google Sheet access failed' })).setMimeType(ContentService.MimeType.JSON);
  }
}
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const nom = String(data.nom || '').trim();
    const prenom = String(data.prenom || '').trim();
    if (!nom || !prenom || !Array.isArray(data.inscriptions) || data.inscriptions.length === 0) throw new Error('Nom, prénom et au moins un agrès sont obligatoires.');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    const now = new Date();
    const rows = data.inscriptions.map(item => {
      if (!item.agre || !['Jury D', 'Jury E'].includes(item.role) || !Array.isArray(item.mouvements) || item.mouvements.length !== 4) throw new Error('Données d’agrès ou de mouvements invalides.');
      const values = item.mouvements.flatMap(m => {
        const counts = ['A', 'B', 'C', 'D', 'E', 'F et +'].map(letter => {
          const value = Number(m[letter]);
          if (!Number.isInteger(value) || value < 0) throw new Error('Les nombres d’éléments doivent être des entiers positifs ou nuls.');
          return value;
        });
        const deduction = item.role === 'Jury E' ? Number(m['Déductions']) : 0;
        if (!Number.isFinite(deduction) || deduction < 0) throw new Error('Les déductions doivent être un nombre positif ou nul.');
        return [...counts, deduction];
      });
      return [now, nom, prenom, String(item.agre), item.role, ...values];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(String(err.stack || err));
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err.message || err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

