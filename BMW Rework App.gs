const SPREADSHEET_ID = "17GQW-mUQPpk5Y0hrUTJE3yQTx9UJA8ditdOOd8lqrnU";
const SHEET_NAME = "REWORK_DATABASE";
const DEFECT_SHEET = "DEFECT_CODES";
const PHOTO_FOLDER_ID = "1FK5LMO7Yo6MEKcoAhqkjxA3a39pZXc1T";

function doGet() {
  return ContentService.createTextOutput("BMW APP TEST OK");
}

function getDefects() {

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName(DEFECT_SHEET);

  const data = sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    2
  ).getValues();

  return data;
}

function saveRecord(data) {

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName(SHEET_NAME);

  let photoUrl = "";

  if (data.image && data.image !== "") {

    const base64Data = data.image.split(",")[1];

    const bytes =
      Utilities.base64Decode(base64Data);

    const fileName =
      data.serial +
      "_" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyyMMdd_HHmmss"
      ) +
      ".jpg";

    const blob = Utilities.newBlob(
      bytes,
      "image/jpeg",
      fileName
    );

    const folder =
      DriveApp.getFolderById(PHOTO_FOLDER_ID);

    const file =
      folder.createFile(blob);

    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    photoUrl = file.getUrl();
  }

  const id =
    "RW-" +
    Utilities.formatString(
      "%06d",
      Math.max(sheet.getLastRow(), 1)
    );

  sheet.appendRow([
    id,
    new Date(),
    data.serial,
    data.defectCode,
    data.defectName,
    data.comment,
    data.qty,
    data.shift,
    photoUrl,
    data.user,
    "OPEN"
  ]);

  return {
    status: "OK",
    id: id
  };
}