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

function registerEquipmentFromRequest(requestRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Assets");
  const data = sheet.getDataRange().getValues();

  const headers = data[0].map(h => String(h).trim());

  const deviceCol = headers.indexOf("Device");
  const brandCol = headers.indexOf("Brand");
  const modelCol = headers.indexOf("Model");
  const snCol = headers.indexOf("SN");
  const teamCol = headers.indexOf("Team");
  const internalSnCol = headers.indexOf("Internal SN");
  const assigneeCol = headers.indexOf("Assignee");

  const requestedName = String(requestRow[3] || "").trim();
  const team = String(requestRow[4] || "").trim();
  const device = String(requestRow[6] || "").trim();
  const brand = String(requestRow[7] || "").trim();
  const model = String(requestRow[8] || "").trim();
  const sn = String(requestRow[9] || "").trim();
  const internalSN = String(requestRow[10] || "").trim();

  let matches = [];

  for (let i = 1; i < data.length; i++) {
    const assetDevice = String(data[i][deviceCol] || "").trim().toLowerCase();
    const assetBrand = String(data[i][brandCol] || "").trim().toLowerCase();
    const assetModel = String(data[i][modelCol] || "").trim().toLowerCase();
    const assetSN = String(data[i][snCol] || "").trim().toLowerCase();
    const assetInternalSN = String(data[i][internalSnCol] || "").trim().toLowerCase();

    if (internalSN && assetInternalSN === internalSN.toLowerCase()) {
      matches.push(i);
      continue;
    }

    if (sn && assetSN === sn.toLowerCase()) {
      matches.push(i);
      continue;
    }

    if (
      !sn &&
      !internalSN &&
      device &&
      brand &&
      model &&
      assetDevice === device.toLowerCase() &&
      assetBrand === brand.toLowerCase() &&
      assetModel === model.toLowerCase()
    ) {
      matches.push(i);
    }
  }

  if (matches.length === 1) {
    sheet.getRange(matches[0] + 1, assigneeCol + 1).setValue(requestedName);
    return;
  }

  if (matches.length > 1) {
    throw new Error("Multiple matching assets found. Please review inventory manually.");
  }

  const labeled = internalSN ? "Yes" : "No";

  const newRow = [
    device,
    brand,
    model,
    sn,
    team,
    internalSN,
    requestedName,
    labeled
  ];

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();
  return;
}