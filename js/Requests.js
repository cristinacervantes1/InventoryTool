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