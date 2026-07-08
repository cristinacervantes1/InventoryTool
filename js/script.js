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

function reviewRequest(requestId, decision, reviewedByEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();

  const status = decision === "approve" ? "approved" : "rejected";
  const reviewedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm"
  );

  for (let i = 1; i < data.length; i++) {
    const rowRequestId = String(data[i][0] || "").trim();

    if (rowRequestId === String(requestId).trim()) {
      const requestType = String(data[i][1] || "").trim();

      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 14).setValue(reviewedByEmail);
      sheet.getRange(i + 1, 15).setValue(reviewedAt);

      if (decision === "approve") {
        if (requestType === "return_equipment") {
          updateAssetAssignee(data[i], "Available");
        }

        if (requestType === "damaged_equipment") {
          updateAssetAssignee(data[i], "Damaged");
        }
      }

      return {
        success: true,
        request_id: requestId,
        status: status
      };
    }
  }

  throw new Error("Request not found: " + requestId);
}

function updateAssetAssignee(requestRow, newAssignee) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assets");
  const data = sheet.getDataRange().getValues();

  const headers = data[0].map(h => String(h).trim());

  const deviceCol = headers.indexOf("Device");
  const snCol = headers.indexOf("SN");
  const internalSnCol = headers.indexOf("Internal SN");
  const assigneeCol = headers.indexOf("Assignee");
  const requestOwner = String(requestRow[3] || "").trim().toLowerCase();

  const rawDevice = String(requestRow[6] || "").trim();

  let requestDevice = rawDevice.toLowerCase();
  let requestSN = String(requestRow[9] || "").trim().toLowerCase();
  let requestInternalSN = String(requestRow[10] || "").trim().toLowerCase();

  if (rawDevice.includes("|")) {
    const parts = rawDevice.split("|").map(part => part.trim());

    requestDevice = String(parts[0] || "").toLowerCase();

    const possibleId = String(parts[parts.length - 1] || "").toLowerCase();

    if (!requestSN) requestSN = possibleId;
    if (!requestInternalSN) requestInternalSN = possibleId;
  }

  function findMatches(matchFn) {
    let matches = [];

    Logger.log("========== REQUEST ==========");
    Logger.log("Device: " + requestDevice);
    Logger.log("SN: " + requestSN);
    Logger.log("Internal SN: " + requestInternalSN);

    for (let i = 1; i < data.length; i++) {

      const assetDevice = String(data[i][deviceCol] || "").trim().toLowerCase();
      const assetSN = String(data[i][snCol] || "").trim().toLowerCase();
      const assetInternalSN = String(data[i][internalSnCol] || "").trim().toLowerCase();
      const assetAssignee = String(data[i][assigneeCol] || "").trim().toLowerCase();

      if (assetAssignee !== requestOwner) {
        continue;
      }

      if (matchFn(assetDevice, assetSN, assetInternalSN)) {
        matches.push(i);
      }
    }

    return matches;
  }

  let matches = [];

  if (requestInternalSN && requestDevice) {
    matches = findMatches((assetDevice, assetSN, assetInternalSN) =>
      assetInternalSN === requestInternalSN && assetDevice === requestDevice
    );
  }

  if (matches.length === 0 && requestSN && requestDevice) {
    matches = findMatches((assetDevice, assetSN, assetInternalSN) =>
      assetSN === requestSN && assetDevice === requestDevice
    );
  }

  if (matches.length === 0 && requestInternalSN) {
    matches = findMatches((assetDevice, assetSN, assetInternalSN) =>
      assetInternalSN === requestInternalSN
    );
  }

  if (matches.length === 0 && requestSN) {
    matches = findMatches((assetDevice, assetSN, assetInternalSN) =>
      assetSN === requestSN
    );
  }

  if (matches.length === 0 && requestDevice) {
    matches = findMatches((assetDevice, assetSN, assetInternalSN) =>
      assetDevice === requestDevice
    );
  }

  if (matches.length === 1) {
    sheet.getRange(matches[0] + 1, assigneeCol + 1).setValue(newAssignee);
    return;
  }

  if (matches.length > 1) {
    throw new Error("Multiple assets found. Please make the request more specific.");
  }

  throw new Error(
  "Asset not found. Looking for -> Device: " + requestDevice +
  " | SN: " + requestSN +
  " | Internal SN: " + requestInternalSN
);
}