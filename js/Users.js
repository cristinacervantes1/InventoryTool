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

function createUser(userData, createdByEmail) {
  const currentUser = getUserByEmail(createdByEmail);

  if (!currentUser) {
    throw new Error("The current user was not found.");
  }

  const currentRole = String(currentUser.Role || "")
    .trim()
    .toLowerCase();

  if (currentRole !== "system_admin") {
    throw new Error(
      "Only a System Admin can create users."
    );
  }

  const name = String(userData.Name || "").trim();

  const email = String(userData.Email || "")
    .trim()
    .toLowerCase();

  const employeeId = String(
    userData["ID"] || ""
  ).trim();

  const team = String(userData.Team || "").trim();

  const role = String(userData.Role || "")
    .trim()
    .toLowerCase();

  if (!name) {
    throw new Error("Name is required.");
  }

  if (!email) {
    throw new Error("Email is required.");
  }

  if (!email.includes("@")) {
    throw new Error("Enter a valid email address.");
  }

  if (!employeeId) {
    throw new Error("Employee ID is required.");
  }

  if (!team) {
    throw new Error("Team is required.");
  }

  const allowedRoles = ["user", "team_admin", "system_admin"];

  if (!allowedRoles.includes(role)) {
    throw new Error("Invalid user role.");
  }

  const users = getUsers();

  const existingEmail = users.find(user =>
    String(user.Email || "")
      .trim()
      .toLowerCase() === email
  );

  if (existingEmail) {
    throw new Error(
      "A user with this email already exists."
    );
  }

  const existingEmployeeId = users.find(user =>
    String(user["ID"] || "")
      .trim()
      .toLowerCase() ===
    employeeId.toLowerCase()
  );

  if (existingEmployeeId) {
    throw new Error(
      "A user with this Employee ID already exists."
    );
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Users");

  if (!sheet) {
    throw new Error("Users sheet was not found.");
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    throw new Error(
      "The Users sheet does not contain headers."
    );
  }

  const headers = data[0].map(header =>
    String(header || "").trim()
  );

  const requiredHeaders = ["ID", "Name", "Email", "Team", "Role"];

  const missingHeaders = requiredHeaders.filter(
    header => !headers.includes(header)
  );

  if (missingHeaders.length > 0) {
    throw new Error(
      "Missing columns in Users sheet: " +
      missingHeaders.join(", ")
    );
  }

  const newUser = {
    ID: employeeId,
    Name: name,
    Email: email,
    Team: team,
    Role: role
  };

  const newRow = headers.map(header =>
    Object.prototype.hasOwnProperty.call(newUser, header)
      ? newUser[header]
      : ""
  );

  sheet.appendRow(newRow);
  SpreadsheetApp.flush();

  return {success: true, user: newUser};
}

function getUserDeletionPreview(userEmail, adminEmail) {
  const admin = getUserByEmail(adminEmail);

  if (!admin) {
    throw new Error("The current user was not found.");
  }

  const adminRole = String(admin.Role || "")
    .trim()
    .toLowerCase();

  if (adminRole !== "system_admin") {
    throw new Error(
      "Only a System Admin can delete users."
    );
  }

  const user = getUserByEmail(userEmail);

  if (!user) {
    throw new Error("The user was not found.");
  }

  const normalizedTargetEmail = String(user.Email || "")
    .trim()
    .toLowerCase();

  const normalizedAdminEmail = String(admin.Email || "")
    .trim()
    .toLowerCase();

  if (normalizedTargetEmail === normalizedAdminEmail) {
    throw new Error(
      "You cannot delete your own user account."
    );
  }

  const users = getUsers();

  const sameNameUsers = users.filter(existingUser =>
    String(existingUser.Name || "")
      .trim()
      .toLowerCase() ===
    String(user.Name || "")
      .trim()
      .toLowerCase()
  );

  if (sameNameUsers.length > 1) {
    throw new Error(
      "This user cannot be deleted because another user has the same name. Assets currently use the user name as the assignee."
    );
  }

  const assets = getAssets();

  const normalizedUserName = String(user.Name || "")
    .trim()
    .toLowerCase();

  const assignedAssets = assets
    .filter(asset =>
      String(asset.Assignee || "")
        .trim()
        .toLowerCase() === normalizedUserName
    )
    .map(asset => ({
      Device: asset.Device || "",
      Brand: asset.Brand || "",
      Model: asset.Model || "",
      SN: asset.SN || "",
      "Internal SN": asset["Internal SN"] || ""
    }));

  return {Name: user.Name || "", Email: user.Email || "", Team: user.Team || "", Role: user.Role || "", Assets: assignedAssets, AssetCount: assignedAssets.length};
}

function deleteUserAndReleaseAssets(userEmail, adminEmail) {
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const admin = getUserByEmail(adminEmail);

    if (!admin) {
      throw new Error("The current user was not found.");
    }

    const adminRole = String(admin.Role || "")
      .trim()
      .toLowerCase();

    if (adminRole !== "system_admin") {
      throw new Error(
        "Only a System Admin can delete users."
      );
    }

    const user = getUserByEmail(userEmail);

    if (!user) {
      throw new Error("The user was not found.");
    }

    const normalizedTargetEmail = String(user.Email || "")
      .trim()
      .toLowerCase();

    const normalizedAdminEmail = String(admin.Email || "")
      .trim()
      .toLowerCase();

    if (normalizedTargetEmail === normalizedAdminEmail) {
      throw new Error(
        "You cannot delete your own user account."
      );
    }

    const users = getUsers();

    const sameNameUsers = users.filter(existingUser =>
      String(existingUser.Name || "")
        .trim()
        .toLowerCase() ===
      String(user.Name || "")
        .trim()
        .toLowerCase()
    );

    if (sameNameUsers.length > 1) {
      throw new Error(
        "This user cannot be deleted because another user has the same name."
      );
    }

    const targetRole = String(user.Role || "")
      .trim()
      .toLowerCase();

    if (targetRole === "system_admin") {
      const systemAdmins = users.filter(existingUser =>
        String(existingUser.Role || "")
          .trim()
          .toLowerCase() === "system_admin"
      );

      if (systemAdmins.length <= 1) {
        throw new Error(
          "The last System Admin cannot be deleted."
        );
      }
    }

    const usersSheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Users");

    if (!usersSheet) {
      throw new Error("Users sheet was not found.");
    }

    const data = usersSheet
      .getDataRange()
      .getValues();

    const headers = data[0].map(header =>
      String(header || "").trim()
    );

    const emailCol = headers.indexOf("Email");

    if (emailCol === -1) {
      throw new Error(
        'The "Email" column was not found in Users.'
      );
    }

    const matchingRows = [];

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailCol] || "")
        .trim()
        .toLowerCase();

      if (rowEmail === normalizedTargetEmail) {
        matchingRows.push(i + 1);
      }
    }

    if (matchingRows.length === 0) {
      throw new Error("The user was not found.");
    }

    if (matchingRows.length > 1) {
      throw new Error(
        "Multiple users with the same email were found."
      );
    }

    const releaseResult =
      releaseAllAssetsForUser(user.Name);

    usersSheet.deleteRow(matchingRows[0]);

    SpreadsheetApp.flush();

    return {
      success: true,
      deletedUser: user.Name || "",
      deletedEmail: user.Email || "",
      releasedAssets:
        releaseResult.releasedAssets || 0
    };
  } finally {
    lock.releaseLock();
  }
}

function updateUser(originalEmail, userData, updatedByEmail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const currentAdmin = getUserByEmail(updatedByEmail);

    if (!currentAdmin) {
      throw new Error("The current user was not found.");
    }

    const currentRole = String(currentAdmin.Role || "")
      .trim()
      .toLowerCase();

    if (currentRole !== "system_admin") {
      throw new Error(
        "Only a System Admin can edit users."
      );
    }

    const existingUser = getUserByEmail(originalEmail);

    if (!existingUser) {
      throw new Error("The user was not found.");
    }

    const employeeId = String(userData.ID || "").trim();
    const name = String(userData.Name || "").trim();

    const email = String(userData.Email || "")
      .trim()
      .toLowerCase();

    const team = String(userData.Team || "").trim();

    const role = String(userData.Role || "")
      .trim()
      .toLowerCase();

    if (!employeeId || !name || !email || !team || !role) {
      throw new Error("Complete all required fields.");
    }

    if (!email.includes("@")) {
      throw new Error("Enter a valid email address.");
    }

    const allowedRoles = [
      "user",
      "team_admin",
      "system_admin"
    ];

    if (!allowedRoles.includes(role)) {
      throw new Error("Invalid user role.");
    }

    const users = getUsers();

    const normalizedOriginalEmail =
      String(originalEmail || "")
        .trim()
        .toLowerCase();

    const duplicateEmail = users.find(user =>
      String(user.Email || "")
        .trim()
        .toLowerCase() === email &&
      String(user.Email || "")
        .trim()
        .toLowerCase() !== normalizedOriginalEmail
    );

    if (duplicateEmail) {
      throw new Error(
        "A user with this email already exists."
      );
    }

    const duplicateEmployeeId = users.find(user =>
      String(user.ID || "")
        .trim()
        .toLowerCase() === employeeId.toLowerCase() &&
      String(user.Email || "")
        .trim()
        .toLowerCase() !== normalizedOriginalEmail
    );

    if (duplicateEmployeeId) {
      throw new Error(
        "A user with this Employee ID already exists."
      );
    }

    const oldRole = String(existingUser.Role || "")
      .trim()
      .toLowerCase();

    if (
      oldRole === "system_admin" &&
      role !== "system_admin"
    ) {
      const systemAdmins = users.filter(user =>
        String(user.Role || "")
          .trim()
          .toLowerCase() === "system_admin"
      );

      if (systemAdmins.length <= 1) {
        throw new Error(
          "The last System Admin cannot be changed to another role."
        );
      }
    }

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Users");

    if (!sheet) {
      throw new Error("Users sheet was not found.");
    }

    const data = sheet.getDataRange().getValues();

    const headers = data[0].map(header =>
      String(header || "").trim()
    );

    const idCol = headers.indexOf("ID");
    const nameCol = headers.indexOf("Name");
    const emailCol = headers.indexOf("Email");
    const teamCol = headers.indexOf("Team");
    const roleCol = headers.indexOf("Role");

    if (
      idCol === -1 ||
      nameCol === -1 ||
      emailCol === -1 ||
      teamCol === -1 ||
      roleCol === -1
    ) {
      throw new Error(
        "One or more required columns are missing in Users."
      );
    }

    let matchingRow = -1;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailCol] || "")
        .trim()
        .toLowerCase();

      if (rowEmail === normalizedOriginalEmail) {
        matchingRow = i + 1;
        break;
      }
    }

    if (matchingRow === -1) {
      throw new Error("The user was not found.");
    }

    const oldName = String(existingUser.Name || "").trim();
    const oldEmail = String(existingUser.Email || "")
      .trim()
      .toLowerCase();

    sheet.getRange(matchingRow, idCol + 1)
      .setValue(employeeId);

    sheet.getRange(matchingRow, nameCol + 1)
      .setValue(name);

    sheet.getRange(matchingRow, emailCol + 1)
      .setValue(email);

    sheet.getRange(matchingRow, teamCol + 1)
      .setValue(team);

    sheet.getRange(matchingRow, roleCol + 1)
      .setValue(role);

    updateUserReferences(
      oldName,
      name,
      oldEmail,
      email,
      team
    );

    SpreadsheetApp.flush();

    return {
      success: true,
      user: {
        ID: employeeId,
        Name: name,
        Email: email,
        Team: team,
        Role: role
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function updateUserReferences(
  oldName,
  newName,
  oldEmail,
  newEmail,
  newTeam
) {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const normalizedOldName = String(oldName || "")
    .trim()
    .toLowerCase();

  const normalizedOldEmail = String(oldEmail || "")
    .trim()
    .toLowerCase();

  const assetsSheet =
    spreadsheet.getSheetByName("Assets");

  if (assetsSheet) {
    const assetData =
      assetsSheet.getDataRange().getValues();

    if (assetData.length > 0) {
      const assetHeaders = assetData[0].map(header =>
        String(header || "").trim()
      );

      const assigneeCol =
        assetHeaders.indexOf("Assignee");

      const teamCol =
        assetHeaders.indexOf("Team");

      if (assigneeCol !== -1) {
        for (let i = 1; i < assetData.length; i++) {
          const assignee = String(
            assetData[i][assigneeCol] || ""
          )
            .trim()
            .toLowerCase();

          if (assignee === normalizedOldName) {
            assetsSheet
              .getRange(i + 1, assigneeCol + 1)
              .setValue(newName);

            if (teamCol !== -1) {
              assetsSheet
                .getRange(i + 1, teamCol + 1)
                .setValue(newTeam);
            }
          }
        }
      }
    }
  }

  const requestsSheet =
    spreadsheet.getSheetByName("Requests");

  if (requestsSheet) {
    const requestData =
      requestsSheet.getDataRange().getValues();

    if (requestData.length > 0) {
      const requestHeaders = requestData[0].map(header =>
        String(header || "").trim()
      );

      const requestedByCol =
        requestHeaders.indexOf("Requested_by");

      const requestedNameCol =
        requestHeaders.indexOf("Requested_name");

      const requestTeamCol =
        requestHeaders.indexOf("Team");

      for (let i = 1; i < requestData.length; i++) {
        if (requestedByCol !== -1) {
          const requestedBy = String(
            requestData[i][requestedByCol] || ""
          )
            .trim()
            .toLowerCase();

          if (requestedBy === normalizedOldEmail) {
            requestsSheet
              .getRange(i + 1, requestedByCol + 1)
              .setValue(newEmail);
          }
        }

        if (requestedNameCol !== -1) {
          const requestedName = String(
            requestData[i][requestedNameCol] || ""
          )
            .trim()
            .toLowerCase();

          if (requestedName === normalizedOldName) {
            requestsSheet
              .getRange(i + 1, requestedNameCol + 1)
              .setValue(newName);

            if (requestTeamCol !== -1) {
              requestsSheet
                .getRange(i + 1, requestTeamCol + 1)
                .setValue(newTeam);
            }
          }
        }
      }
    }
  }
}