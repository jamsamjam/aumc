// @jamsamjam

function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dbSheet = ss.getSheetByName("db_actif");

  Logger.log("Event object: " + JSON.stringify(e))

  const values = e.values;

  Logger.log(" values length: " + values.length);
  Logger.log(" values content: " + values);

  const nom = values[1];
  const prenom = values[2];
  let tel_priv = values[5];

  if (tel_priv && tel_priv.startsWith("+")) {
    tel_priv = "'" + tel_priv;
  }

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

  Logger.log("Writing to sheet: " + dbSheet.getName());
  Logger.log("Current last row in db_actif: " + dbSheet.getLastRow());

  dbSheet.appendRow(newRow);

  Logger.log("New last row after append: " + dbSheet.getLastRow());
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

function onEdit(e) {

  const sheetName = "db_actif";
  const paymentColumn = 9; // salles_piano_paiement
  const mailSentColumn = 12; // mail_sent

  const sheet = e.range.getSheet();
  if (sheet.getName() !== sheetName) return;

  const row = e.range.getRow();
  const column = e.range.getColumn();

  // exclude header row
  if (row <= 1 || column !== paymentColumn) return;

  const newValue = e.value;
  const oldValue = e.oldValue;

  if (!oldValue && newValue) {

    const uni = sheet.getRange(row, 7).getValue();
    if (uni !== "EPFL") return;

    const alreadySent = sheet.getRange(row, mailSentColumn).getValue();
    if (alreadySent === true) return;

    const email = sheet.getRange(row, 5).getValue();
    const nom = sheet.getRange(row, 2).getValue();
    const prenom = sheet.getRange(row, 3).getValue();

    if (!email) return;

    const today = new Date();
    const formattedDate = Utilities.formatDate(
      today,
      Session.getScriptTimeZone(),
      "d MMMM yyyy"
    );

    const subject = "Accréditations AUMC";
    
    const htmlBody =
  "Bonjour " + prenom + " " + nom + ",<br><br>" +
  "You have been granted access to the piano rooms starting from today, " + formattedDate + ", for the duration you paid for.<br><br>" +
  "You can find the procedure on how to update the access rights on your Camipro card and where to find the rooms on our " +
  "<a href='https://www.epfl.ch/campus/associations/aumc/fr/salles-piano/'>website</a>.<br><br>" +
  "Please also read the rules on the same site regarding booking and using the rooms.<br><br>" +
  "If you have any questions or encounter problems accessing the piano rooms, do not hesitate to write to aumc@epfl.ch or ask " + 
  "<a href='https://www.epfl.ch/campus/associations/aumc/fr/contact/'>here</a>.<br><br>" +
  "Musical greetings,<br>AUMC";

    try{
      GmailApp.sendEmail(email, subject, "", {
        htmlBody: htmlBody
      });

      sheet.getRange(row, mailSentColumn).setValue(true);
    }
    catch(err) {
      Logger.log("Error sending email: " + err);
    }
  }
}

function sendBiMonthlyReport() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("db_actif");

  const data = sheet.getDataRange().getValues();

  const reportColumn = 13; // reported
  const nameColumn = 2;
  const prenomColumn = 3;
  const sciperColumn = 8;

  let list = [];
  let rowsToUpdate = [];

  for (let i = 1; i < data.length; i++) {

    const reported = data[i][reportColumn - 1];
    const uni = data[i][6];

    if (uni !== "EPFL") continue;

    const nom = data[i][nameColumn - 1];
    const prenom = data[i][prenomColumn - 1];
    const sciper = data[i][sciperColumn - 1];

    if (!nom && !prenom) continue;

    if (reported !== true) {
      list.push(prenom + " " + nom + " - " + sciper);
      rowsToUpdate.push(i + 1);
    }
  }

  if (list.length === 0) return;

  const subject = "New piano room members";

  const body =
    "Bonjour," +
    "Newly added members are:\n\n" +
    list.join("\n");

  try {
      GmailApp.sendEmail(
      "responsible@gmail.com", // TODO
      subject,
      body
    );

    rowsToUpdate.forEach(row => {
      sheet.getRange(row, reportColumn).setValue(true);
    });
  }
  catch(err) {
    Logger.log("Error sending email: " + err);
  }
}