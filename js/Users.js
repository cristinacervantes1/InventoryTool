function getUserByEmail(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  const users = rows.map(row => {
    let user = {};
    headers.forEach((header, index) => {
      user[header] = row[index];
    });
    return user;
  });

  return users.find(user =>
    String(user.Email).toLowerCase().trim() === String(email).toLowerCase().trim()
  ) || null;
}

function getUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  return rows.map(row => {
    let item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  });
}

function getVisibleUsers(email) {
  const currentUser = getUserByEmail(email);
  if (!currentUser) return [];

  const users = getUsers();
  const role = String(currentUser.Role || "").trim().toLowerCase();
  const team = String(currentUser.Team || "").trim().toLowerCase();

  if (role === "system_admin") {
    return users;
  }

  if (role === "team_admin") {
    return users.filter(user =>
      String(user.Team || "").trim().toLowerCase() === team
    );
  }

  return [];
}