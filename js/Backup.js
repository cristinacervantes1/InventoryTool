const BACKUP_AUTHORIZED_USERS = [
  "maribel.nieves@essigq.com",
  "cecilia.dolores@essig.com"
];

function downloadSpreadsheetBackup(userEmail) {
  const user = getUserByEmail(userEmail);

  if (!user) {
    throw new Error("User not found.");
  }

  const normalizedEmail = normalize(user.Email);

  const authorized = BACKUP_AUTHORIZED_USERS.some(email =>
    normalize(email) === normalizedEmail
  );

  if (!authorized) {
    throw new Error(
      "You are not authorized to download local backups."
    );
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();

  const exportUrl =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?format=xlsx";

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization:
        "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      "Unable to generate the Excel backup."
    );
  }

  const blob = response.getBlob();

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_HH-mm"
  );

  const fileName =
    "Asset_Management_Backup_" +
    timestamp +
    ".xlsx";

  return {
    fileName: fileName,
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    base64: Utilities.base64Encode(
      blob.getBytes()
    )
  };
}

function canDownloadBackup(userEmail) {
  const user = getUserByEmail(userEmail);

  if (!user) {
    return false;
  }

  const normalizedEmail =
    normalize(user.Email);

  return BACKUP_AUTHORIZED_USERS.some(email =>
    normalize(email) === normalizedEmail
  );
}
