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

function createRequest(request) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");

  const nextNumber = sheet.getLastRow();
  const requestId = "REQ-" + String(nextNumber).padStart(6, "0");

  sheet.appendRow([
    requestId,
    request.Request_type || "",
    request.Requested_by || "",
    request.Requested_name || "",
    request.Team || "",
    "pending",
    request.Device || "",
    request.Brand || "",
    request.Model || "",
    request.SN || "",
    request["Internal SN"] || "",
    request.Quantity || 1,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"), "", "",
    request.Comments || ""
  ]);

  return {
    success: true,
    request_id: requestId
  };
}

function getRequests() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");
  const data = sheet.getDataRange().getDisplayValues();

  const rows = data.slice(1);

  return rows
    .filter(row => row[0])
    .map(row => ({
      Request_id: row[0] || "",
      Request_type: row[1] || "",
      Requested_by: row[2] || "",
      Requested_name: row[3] || "",
      Team: row[4] || "",
      Status: row[5] || "",
      Device: row[6] || "",
      Brand: row[7] || "",
      Model: row[8] || "",
      SN: row[9] || "",
      "Internal SN": row[10] || "",
      Quantity: row[11] || "",
      Created_at: row[12] || "",
      Reviewed_by: row[13] || "",
      Reviewed_at: row[14] || "",
      Comments: row[15] || ""
    }));
}

function getMyRequests(email) {
  const currentUser = getUserByEmail(email);
  if (!currentUser) return [];

  const requests = getRequests();

  const userEmail = String(currentUser.Email || "").trim().toLowerCase();
  const userName = String(currentUser.Name || "").trim().toLowerCase();

  return requests.filter(request => {
    const requestedBy = String(request.Requested_by || "").trim().toLowerCase();
    const requestedName = String(request.Requested_name || "").trim().toLowerCase();

    return requestedBy === userEmail || requestedName === userName;
  });
}

function getVisibleRequests(email) {
  const currentUser = getUserByEmail(email);
  if (!currentUser) return [];

  const requests = getRequests();
  const role = String(currentUser.Role || "").trim().toLowerCase();
  const team = String(currentUser.Team || "").trim().toLowerCase();

  if (role === "system_admin") {
    return requests;
  }

  if (role === "team_admin") {
    return requests.filter(request =>
      String(request.Team || "").trim().toLowerCase() === team
    );
  }

  return [];
}