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
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Assets");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const deviceCol = headers.indexOf("Device");
  const brandCol = headers.indexOf("Brand");
  const modelCol = headers.indexOf("Model");
  const snCol = headers.indexOf("SN");
  const internalSnCol = headers.indexOf("Internal SN");
  const assigneeCol = headers.indexOf("Assignee");
  const teamCol = headers.indexOf("Team");

  const requestOwner = String(requestRow[3] || "")
    .trim()
    .toLowerCase();

  const requestDevice = String(requestRow[6] || "")
    .trim()
    .toLowerCase();

  const requestBrand = String(requestRow[7] || "")
    .trim()
    .toLowerCase();

  const requestModel = String(requestRow[8] || "")
    .trim()
    .toLowerCase();

  const requestSN = String(requestRow[9] || "")
    .trim()
    .toLowerCase();

  const requestInternalSN = String(requestRow[10] || "")
    .trim()
    .toLowerCase();

  const ownerAssets = [];

  for (let i = 1; i < data.length; i++) {
    const assetAssignee = String(data[i][assigneeCol] || "")
      .trim()
      .toLowerCase();

    if (assetAssignee !== requestOwner) {
      continue;
    }

    ownerAssets.push({
      rowIndex: i,
      device: String(data[i][deviceCol] || "").trim().toLowerCase(),
      brand: String(data[i][brandCol] || "").trim().toLowerCase(),
      model: String(data[i][modelCol] || "").trim().toLowerCase(),
      sn: String(data[i][snCol] || "").trim().toLowerCase(),
      internalSN: String(data[i][internalSnCol] || "").trim().toLowerCase()
    });
  }

  let matches = [];

  // 1. Internal SN + Device
  if (requestInternalSN) {
    matches = ownerAssets.filter(asset =>
      asset.internalSN === requestInternalSN &&
      asset.device === requestDevice
    );
  }

  // 2. SN + Device
  if (matches.length === 0 && requestSN) {
    matches = ownerAssets.filter(asset =>
      asset.sn === requestSN &&
      asset.device === requestDevice
    );
  }

  // 3. Device + Brand + Model
  if (
    matches.length === 0 &&
    requestDevice &&
    (requestBrand || requestModel)
  ) {
    matches = ownerAssets.filter(asset =>
      asset.device === requestDevice &&
      asset.brand === requestBrand &&
      asset.model === requestModel
    );
  }

  // 4. Device solamente, como último recurso
  if (
    matches.length === 0 &&
    requestDevice &&
    !requestBrand &&
    !requestModel &&
    !requestSN &&
    !requestInternalSN
  ) {
    matches = ownerAssets.filter(asset =>
      asset.device === requestDevice
    );
  }

  if (matches.length === 1) {
    const rowNumber = matches[0].rowIndex + 1;

    sheet
      .getRange(rowNumber, assigneeCol + 1)
      .setValue(newAssignee);

    const normalizedNewAssignee = String(newAssignee || "")
      .trim()
      .toLowerCase();

    if (
      normalizedNewAssignee === "available" ||
      normalizedNewAssignee === "damaged"
    ) {
      sheet
        .getRange(rowNumber, teamCol + 1)
        .setValue("");
    }

    SpreadsheetApp.flush();
    return;
  }

  if (matches.length > 1) {
    throw new Error(
      "Multiple matching assets found for " +
      requestOwner +
      ". Device: " + requestDevice +
      " | Brand: " + requestBrand +
      " | Model: " + requestModel
    );
  }

  throw new Error(
    "Asset not found. Device: " + requestDevice +
    " | Brand: " + requestBrand +
    " | Model: " + requestModel +
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

function getAvailableAssetsByDevice(device) {
  const assets = getAssets();

  const requestedDevice = String(device || "").trim().toLowerCase();

  return assets.filter(asset =>
    String(asset.Device || "").trim().toLowerCase() === requestedDevice &&
    String(asset.Assignee || "").trim().toLowerCase() === "available"
  );
}

function assignSelectedAssetToUser(assetData, requestedName, requestedTeam) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Assets");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  const deviceCol = headers.indexOf("Device");
  const brandCol = headers.indexOf("Brand");
  const modelCol = headers.indexOf("Model");
  const snCol = headers.indexOf("SN");
  const teamCol = headers.indexOf("Team");
  const internalSnCol = headers.indexOf("Internal SN");
  const assigneeCol = headers.indexOf("Assignee");

  const device = String(assetData.Device || "").trim().toLowerCase();
  const brand = String(assetData.Brand || "").trim().toLowerCase();
  const model = String(assetData.Model || "").trim().toLowerCase();
  const sn = String(assetData.SN || "").trim().toLowerCase();
  const internalSN = String(
    assetData["Internal SN"] || ""
  ).trim().toLowerCase();

  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const assetDevice = String(data[i][deviceCol] || "").trim().toLowerCase();
    const assetBrand = String(data[i][brandCol] || "").trim().toLowerCase();
    const assetModel = String(data[i][modelCol] || "").trim().toLowerCase();
    const assetSN = String(data[i][snCol] || "").trim().toLowerCase();
    const assetInternalSN = String(
      data[i][internalSnCol] || ""
    ).trim().toLowerCase();
    const assetAssignee = String(
      data[i][assigneeCol] || ""
    ).trim().toLowerCase();

    if (assetAssignee !== "available") {
      continue;
    }

    const sameAsset =
      assetDevice === device &&
      assetBrand === brand &&
      assetModel === model &&
      assetSN === sn &&
      assetInternalSN === internalSN;

    if (sameAsset) {
      matches.push(i);
    }
  }

  if (matches.length === 0) {
    throw new Error("The selected asset is no longer available.");
  }

  if (matches.length > 1) {
    throw new Error("Multiple identical available assets were found.");
  }

  const rowNumber = matches[0] + 1;

  sheet
    .getRange(rowNumber, assigneeCol + 1)
    .setValue(requestedName);

  sheet
    .getRange(rowNumber, teamCol + 1)
    .setValue(requestedTeam);

  SpreadsheetApp.flush();

  return {
    success: true
  };
}