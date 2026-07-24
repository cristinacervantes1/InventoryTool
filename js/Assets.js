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
  const userName = normalize(user.Name);

  return assets.filter(asset =>
    normalize(asset.Assignee) === userName
  );
}

function getTeamInventory(email) {
  const user = getUserByEmail(email);
  if (!user) return [];

  const role = normalize(user.Role);

  if (role !== "team_admin" && role !== "system_admin") {
    return [];
  }

  const assets = getAssets();
  const team = normalize(user.Team);

  return assets.filter(asset =>
    normalize(asset.Team) === team
  );
}

function getAllInventory(email) {
  const user = getUserByEmail(email);
  if (!user) return [];

  const role = normalize(user.Role);

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

  const requestOwner = normalize(requestRow[3]);
  const requestDevice = normalize(requestRow[6]);
  const requestBrand = normalize(requestRow[7]);
  const requestModel = normalize(requestRow[8]);
  const requestSN = normalize(requestRow[9]);
  const requestInternalSN = normalize(requestRow[10]);

  const ownerAssets = [];

  for (let i = 1; i < data.length; i++) {
    const assetAssignee = normalize(data[i][assigneeCol]);

    if (assetAssignee !== requestOwner) {
      continue;
    }

    ownerAssets.push({
      rowIndex: i,
      device: normalize(data[i][deviceCol]),
      brand: normalize(data[i][brandCol]),
      model: normalize(data[i][modelCol]),
      sn: normalize(data[i][snCol]),
      internalSN: normalize(data[i][internalSnCol])
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

    const normalizedNewAssignee = normalize(newAssignee);

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
  Logger.log("registerEquipmentFromRequest ejecutada");
  Logger.log(requestRow);
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

  const normalizedDevice = normalize(device);
  const normalizedBrand = normalize(brand);
  const normalizedModel = normalize(model);
  const normalizedSN = normalize(sn);
  const normalizedInternalSN = normalize(internalSN);

  for (let i = 1; i < data.length; i++) {
    const assetDevice = normalize(data[i][deviceCol]);
    const assetBrand = normalize(data[i][brandCol]);
    const assetModel = normalize(data[i][modelCol]);
    const assetSN = normalize(data[i][snCol]);
    const assetInternalSN = normalize(data[i][internalSnCol]);

    if (
      normalizedInternalSN &&
      assetInternalSN === normalizedInternalSN
    ) {
      matches.push(i);
      continue;
    }

    if (
      normalizedSN &&
      assetSN === normalizedSN
    ) {
      matches.push(i);
      continue;
    }

    if (
      !normalizedSN &&
      !normalizedInternalSN &&
      normalizedDevice &&
      normalizedBrand &&
      normalizedModel &&
      assetDevice === normalizedDevice &&
      assetBrand === normalizedBrand &&
      assetModel === normalizedModel
    ) {
      matches.push(i);
    }
  }

  if (matches.length === 1) {

      const rowNumber = matches[0] + 1;

      sheet
        .getRange(rowNumber, assigneeCol + 1)
        .setValue(requestedName);

      sheet
        .getRange(rowNumber, teamCol + 1)
        .setValue(team);

      SpreadsheetApp.flush();

      return;
  }

  if (matches.length > 1) {
    throw new Error("Multiple matching assets found. Please review inventory manually.");
  }

  const labeled = internalSN ? "Yes" : "No";

  const newRow = [device, brand, model, sn, team, internalSN, requestedName, labeled];

  const nextRow = sheet.getLastRow() + 1;

  Logger.log("Hoja: " + sheet.getName());
  Logger.log("Fila donde se insertará: " + nextRow);
  Logger.log("Datos a insertar: " + JSON.stringify(newRow));

  sheet
    .getRange(nextRow, 1, 1, newRow.length)
    .setValues([newRow]);

  SpreadsheetApp.flush();

  const insertedValues = sheet
    .getRange(nextRow, 1, 1, newRow.length)
    .getDisplayValues()[0];

  Logger.log("Datos leídos después de insertar:");
  Logger.log(insertedValues);
  Logger.log(
    "Spreadsheet ID: " +
    SpreadsheetApp.getActiveSpreadsheet().getId()
  );
  Logger.log(
    "Spreadsheet name: " +
    SpreadsheetApp.getActiveSpreadsheet().getName()
  );
  return;
}

function getAvailableAssetsByDevice(device) {
  const assets = getAssets();

  const requestedDevice = normalize(device);

  return assets.filter(asset =>
    normalize(asset.Device) === requestedDevice &&
    normalize(asset.Assignee) === "available"
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

  const device = normalize(assetData.Device);
  const brand = normalize(assetData.Brand);
  const model = normalize(assetData.Model);
  const sn = normalize(assetData.SN);
  const internalSN = normalize(assetData["Internal SN"]);

  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const assetDevice = normalize(data[i][deviceCol]);
    const assetBrand = normalize(data[i][brandCol]);
    const assetModel = normalize(data[i][modelCol]);
    const assetSN = normalize(data[i][snCol]);
    const assetInternalSN = normalize(data[i][internalSnCol]);
    const assetAssignee = normalize(data[i][assigneeCol]);

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

function releaseAllAssetsForUser(userName) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Assets");

  if (!sheet) {
    throw new Error("Assets sheet was not found.");
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return {
      releasedAssets: 0
    };
  }

  const headers = data[0].map(header =>
    String(header || "").trim()
  );

  const assigneeCol = headers.indexOf("Assignee");
  const teamCol = headers.indexOf("Team");

  if (assigneeCol === -1) {
    throw new Error(
      'The "Assignee" column was not found in Assets.'
    );
  }

  if (teamCol === -1) {
    throw new Error(
      'The "Team" column was not found in Assets.'
    );
  }

  const normalizedUserName = normalize(userName);

  let releasedAssets = 0;

  for (let i = 1; i < data.length; i++) {
    const currentAssignee = normalize(
      data[i][assigneeCol]
    );

    if (currentAssignee !== normalizedUserName) {
      continue;
    }

    sheet
      .getRange(i + 1, assigneeCol + 1)
      .setValue("Available");

    sheet
      .getRange(i + 1, teamCol + 1)
      .setValue("");

    releasedAssets++;
  }

  SpreadsheetApp.flush();

  return {releasedAssets: releasedAssets};
}