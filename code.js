// @jamsamjam

function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dbSheet = ss.getSheetByName("db_actif");

  const values = e.values;

  const nom = values[1];
  const prenom = values[2];
  const tel_priv = values[5];
  const email = values[7];
  const statut = values[8];
  const instit = values[9];
  const sciper = values[10];

  const id = getNextId(dbSheet);
  const today = new Date();

  const newRow = [
    id,
    nom,
    prenom,
    tel_priv,
    email,
    statut,
    instit,
    sciper,
    "", // salles_piano_paiement
    "", // participation
    today
  ];

  dbSheet.appendRow(newRow);
}

function getNextId(sheet) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return 1;
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1)
    .getValues()
    .flat()
    .map(id => Number(id))
    .filter(id => !isNaN(id));

  if (ids.length === 0) {
    return 1;
  }

  const maxId = Math.max(...ids);

  return maxId + 1;
}