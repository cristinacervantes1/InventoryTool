function createRequest(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Requests");

    const lastRow = sheet.getLastRow();
    let highestNumber = 0;

    if (lastRow >= 2) {
      const existingIds = sheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues()
        .flat();

      existingIds.forEach(id => {
        const match = String(id || "")
          .trim()
          .match(/^REQ-(\d+)$/i);

        if (match) {
          highestNumber = Math.max(
            highestNumber,
            Number(match[1])
          );
        }
      });
    }

    const nextNumber = highestNumber + 1;

    const requestId =
      "REQ-" +
      String(nextNumber).padStart(6, "0");

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
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm"
      ),
      "",
      "",
      request.Comments || ""
    ]);

    SpreadsheetApp.flush();

    return {success: true, request_id: requestId};
  } finally {
    lock.releaseLock();
  }
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

  const userEmail = normalize(currentUser.Email);
  const userName = normalize(currentUser.Name);

  return requests.filter(request => {
    const requestedBy = normalize(request.Requested_by);
    const requestedName = normalize(request.Requested_name);

    return requestedBy === userEmail || requestedName === userName;
  });
}

function getVisibleRequests(email) {
  const currentUser = getUserByEmail(email);
  if (!currentUser) return [];

  const requests = getRequests();
  const role = normalize(currentUser.Role);
  const team = normalize(currentUser.Team);

  if (role === "system_admin") {
    return requests;
  }

  if (role === "team_admin") {
    return requests.filter(request =>
      normalize(request.Team) === team
    );
  }

  return [];
}

function reviewRequest(requestId, decision, reviewedByEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();

  if (
    decision !== "approve" &&
    decision !== "reject"
  ) {
    throw new Error("Invalid review decision.");
  }
  const status = decision === "approve"
      ? "approved"
      : "rejected";
  const reviewedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm"
  );

  const normalizedRequestId = String(requestId || "").trim();

  const matchingRows = [];

  for (let i = 1; i < data.length; i++) {
    const rowRequestId = String(data[i][0] || "").trim();

    if (rowRequestId === normalizedRequestId) {
      matchingRows.push(i);
    }
  }

  if (matchingRows.length === 0) {
    throw new Error("Request not found: " + requestId);
  }

  if (matchingRows.length > 1) {
    throw new Error(
      "Duplicate Request ID found: " + requestId +
      ". Please correct the duplicate IDs in the Requests sheet."
    );
  }

  for (const i of matchingRows) {
    const rowRequestId = String(data[i][0] || "").trim();

    if (rowRequestId === String(requestId).trim()) {
      const currentStatus = normalize(data[i][5]);

      if (currentStatus !== "pending") {
        throw new Error(
          "This request has already been reviewed."
        );
      }
      const requestType = String(data[i][1] || "").trim();

      if (decision === "approve") {
        if (requestType === "return_equipment") {
          updateAssetAssignee(data[i], "Available");
        }

        if (requestType === "damaged_equipment") {
          updateAssetAssignee(data[i], "Damaged");
        }

        if (requestType === "register_equipment") {
          registerEquipmentFromRequest(data[i]);
        }

        if (requestType === "edit_equipment") {
          updateAssetInformation(data[i]);
        }
      }
      
      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 14).setValue(reviewedByEmail);
      sheet.getRange(i + 1, 15).setValue(reviewedAt);

      return {
        success: true,
        request_id: requestId,
        status: status
      };
    }
  }

  throw new Error("Request not found: " + requestId);
}

function updateAssetInformation(requestRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const assetsSheet = ss.getSheetByName("Assets");

  if (!assetsSheet) {
    throw new Error('The "Assets" sheet was not found.');
  }

  const range = assetsSheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  if (values.length < 2) {
    throw new Error("No assets were found.");
  }

  const headers = displayValues[0].map(header =>
    String(header || "").trim()
  );

  function getColumnIndex(headerName) {
    const index = headers.indexOf(headerName);

    if (index === -1) {
      throw new Error(
        'The "' + headerName + '" column was not found in Assets.'
      );
    }

    return index;
  }

  const columns = {
    Device: getColumnIndex("Device"),
    Brand: getColumnIndex("Brand"),
    Model: getColumnIndex("Model"),
    SN: getColumnIndex("SN"),
    InternalSN: getColumnIndex("Internal SN")
  };

  /*
   * Original values stored in the Requests row.
   *
   * 6  = Device
   * 7  = Brand
   * 8  = Model
   * 9  = SN
   * 10 = Internal SN
   * 15 = Comments
   */
  const original = {
    Device: String(requestRow[6] || "").trim(),
    Brand: String(requestRow[7] || "").trim(),
    Model: String(requestRow[8] || "").trim(),
    SN: String(requestRow[9] || "").trim(),
    "Internal SN": String(requestRow[10] || "").trim()
  };

  let commentsData;

  try {
    commentsData = JSON.parse(
      String(requestRow[15] || "{}")
    );
  } catch (error) {
    throw new Error(
      "The edit request contains invalid JSON in Comments."
    );
  }

  const updated = commentsData.updated;

  if (
    !updated ||
    typeof updated !== "object" ||
    Array.isArray(updated)
  ) {
    throw new Error(
      "The edit request does not contain updated asset information."
    );
  }

  const assetRows = displayValues.slice(1).map(
    (row, index) => ({
      row: row,
      sheetRow: index + 2
    })
  );

  function isSame(valueA, valueB) {
    return normalize(valueA) === normalize(valueB);
  }

  function findUniqueAsset(predicate, searchDescription) {
    const matches = assetRows.filter(asset =>
      predicate(asset.row)
    );

    if (matches.length > 1) {
      throw new Error(
        "Multiple assets matched using " +
        searchDescription +
        ". The request cannot be approved safely."
      );
    }

    return matches.length === 1
      ? matches[0]
      : null;
  }

  let assetMatch = null;

  /*
   * Priority 1:
   * Internal SN + Device
   */
  if (
    original["Internal SN"] &&
    original.Device
  ) {
    assetMatch = findUniqueAsset(
      row =>
        isSame(
          row[columns.InternalSN],
          original["Internal SN"]
        ) &&
        isSame(
          row[columns.Device],
          original.Device
        ),
      "Internal SN and Device"
    );
  }

  /*
   * Priority 2:
   * SN + Device
   */
  if (
    !assetMatch &&
    original.SN &&
    original.Device
  ) {
    assetMatch = findUniqueAsset(
      row =>
        isSame(
          row[columns.SN],
          original.SN
        ) &&
        isSame(
          row[columns.Device],
          original.Device
        ),
      "SN and Device"
    );
  }

  /*
   * Priority 3:
   * Device + Brand + Model
   */
  if (
    !assetMatch &&
    original.Device &&
    original.Brand &&
    original.Model
  ) {
    assetMatch = findUniqueAsset(
      row =>
        isSame(
          row[columns.Device],
          original.Device
        ) &&
        isSame(
          row[columns.Brand],
          original.Brand
        ) &&
        isSame(
          row[columns.Model],
          original.Model
        ),
      "Device, Brand and Model"
    );
  }

  /*
   * Priority 4:
   * Device only
   */
  if (
    !assetMatch &&
    original.Device
  ) {
    assetMatch = findUniqueAsset(
      row =>
        isSame(
          row[columns.Device],
          original.Device
        ),
      "Device"
    );
  }

  if (!assetMatch) {
    throw new Error(
      "The original asset could not be found. " +
      "No information was updated."
    );
  }

  const editableFields = [
    {
      requestKey: "Device",
      assetColumn: columns.Device
    },
    {
      requestKey: "Brand",
      assetColumn: columns.Brand
    },
    {
      requestKey: "Model",
      assetColumn: columns.Model
    },
    {
      requestKey: "SN",
      assetColumn: columns.SN
    },
    {
      requestKey: "Internal SN",
      assetColumn: columns.InternalSN
    }
  ];

  let fieldsUpdated = 0;

  editableFields.forEach(field => {
    if (
      Object.prototype.hasOwnProperty.call(
        updated,
        field.requestKey
      )
    ) {
      const newValue =
        updated[field.requestKey] === null ||
        updated[field.requestKey] === undefined
          ? ""
          : updated[field.requestKey];

      assetsSheet
        .getRange(
          assetMatch.sheetRow,
          field.assetColumn + 1
        )
        .setValue(newValue);

      fieldsUpdated++;
    }
  });

  if (fieldsUpdated === 0) {
    throw new Error(
      "The request does not contain any editable asset fields."
    );
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    asset_row: assetMatch.sheetRow,
    fields_updated: fieldsUpdated
  };
}

function approveAssignedRequest(requestId, reviewedByEmail) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Requests");

  const data = sheet.getDataRange().getValues();

  const reviewedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm"
  );

  const matches = [];

  for (let i = 1; i < data.length; i++) {
    const rowRequestId = String(data[i][0] || "").trim();

    if (rowRequestId === String(requestId || "").trim()) {
      matches.push(i);
    }
  }

  if (matches.length === 0) {
    throw new Error("Request not found: " + requestId);
  }

  if (matches.length > 1) {
    throw new Error("Duplicate Request ID found: " + requestId);
  }

  const rowIndex = matches[0];

  sheet.getRange(rowIndex + 1, 6).setValue("approved");
  sheet.getRange(rowIndex + 1, 14).setValue(reviewedByEmail);
  sheet.getRange(rowIndex + 1, 15).setValue(reviewedAt);

  SpreadsheetApp.flush();

  return {success: true, request_id: requestId, status: "approved"};
}