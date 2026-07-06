function doGet() {
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Inventory Tool")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAssets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assets");
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

function getMyInventory(email) {
  const user = getUserByEmail(email);
  if (!user) return [];

  const assets = getAssets();
  const userName = String(user.Name || "").trim().toLowerCase();

  return assets.filter(asset =>
    String(asset.Assignee || "").trim().toLowerCase() === userName
  );
}

function getTeamInventory(email) {
  const user = getUserByEmail(email);
  if (!user) return [];

  const role = String(user.Role || "").trim().toLowerCase();

  if (role !== "team_admin" && role !== "system_admin") {
    return [];
  }

  const assets = getAssets();
  const team = String(user.Team || "").trim().toLowerCase();

  return assets.filter(asset =>
    String(asset.Team || "").trim().toLowerCase() === team
  );
}

function getAllInventory(email) {
  const user = getUserByEmail(email);
  if (!user) return [];

  const role = String(user.Role || "").trim().toLowerCase();

  if (role !== "system_admin") {
    return [];
  }

  return getAssets();
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