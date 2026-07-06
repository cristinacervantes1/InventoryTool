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

function openRequestForm(type) {
  document.getElementById("requestModal").style.display = "block";

  document.getElementById("requestBrand").style.display = "block";
  document.getElementById("requestModel").style.display = "block";
  document.getElementById("requestSN").style.display = "block";
  document.getElementById("requestInternalSN").style.display = "block";

  if (type === "request_equipment") {
    document.getElementById("requestTitle").textContent = "Request Equipment";
    document.getElementById("requestBrand").style.display = "none";
    document.getElementById("requestModel").style.display = "none";
    document.getElementById("requestSN").style.display = "none";
    document.getElementById("requestInternalSN").style.display = "none";
  }

  if (type === "register_equipment") {
    document.getElementById("requestTitle").textContent = "Register Equipment";
  }

  if (type === "damaged_equipment") {
    document.getElementById("requestTitle").textContent = "Report Damaged Equipment";
    document.getElementById("requestBrand").style.display = "none";
    document.getElementById("requestModel").style.display = "none";
    document.getElementById("requestSN").style.display = "none";
    document.getElementById("requestInternalSN").style.display = "none";
  }

  if (type === "return_equipment") {
    document.getElementById("requestTitle").textContent = "Return Equipment";
    document.getElementById("requestBrand").style.display = "none";
    document.getElementById("requestModel").style.display = "none";
    document.getElementById("requestSN").style.display = "none";
    document.getElementById("requestInternalSN").style.display = "none";
  }
}

function closeRequestForm() {
  document.getElementById("requestModal").style.display = "none";
}

function createRequest(request) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Requests");

  const requestId = "REQ-" + new Date().getTime();

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
    new Date(),
    "",
    "",
    request.Comments || ""
  ]);

  return {
    success: true,
    request_id: requestId
  };
}