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

    const email = sheet.getRange(row, 5).getValue(); // email
    const nom = sheet.getRange(row, 2).getValue(); // nom
    const prenom = sheet.getRange(row, 3).getValue(); // prenom

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