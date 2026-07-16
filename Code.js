/**
 * ============================================================
 * FIBERONE SD REQUEST HUB — COMPLETE BACKEND v1.0
 * Google Apps Script for: index.html, auth.html, admin.html, employee.html
 *
 * SHEETS USED:
 *   SD_USERS       — All registered portal users
 *   SD_DEPARTMENTS — Company departments
 *   SD_REQUESTS    — All submitted requests
 *   SD_NEWS        — Employee dashboard news items
 *   SD_ALLOWLIST   — Admin allowlist
 *
 * SESSION SYSTEM:
 *   Uses PropertiesService (UserProperties) — no sheet needed.
 *   Each logged-in browser session stores a token key in UserProperties.
 *   Token → JSON payload (email, name, role, department, expires).
 *
 * HOW TO DEPLOY
 *   1. Open your Google Spreadsheet
 *   2. Extensions → Apps Script → paste this entire file as Code.gs
 *   3. Also create these HTML files in the same project:
 *      index.html, auth.html, admin.html, employee.html
 *      (paste your existing HTML into each)
 *   4. In each HTML file, find the closing </body> tag and add:
 *      <script>var scriptUrl = "<?= scriptUrl ?>";</script>
 *      (this is already in your auth.html — add it to admin/employee too)
 *   5. Run "runFirstSetup" ONCE from the ⚙️ SD PORTAL ADMIN menu
 *   6. Deploy → New Deployment → Web App
 *      - Execute as: Me
 *      - Who has access: Anyone (even anonymous)
 *   7. Copy the Web App URL — it becomes scriptUrl automatically
 * ============================================================
 */

// ─────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────
var SD = {
  USERS: "SD_USERS",
  DEPTS: "SD_DEPARTMENTS",
  PAT: "SD_PAT_PROJECTS",
  MATS: "SD_MATERIALS",
  MAILS: "SD_MAILS",
  JCC: "SD_JCC_CERTIFICATES",
  BLACKLIST: "SD_BLACKLIST",
  DOCS: "SD_DOCUMENTS",
  ADMIN: "admin",
  EMPLOYEE: "employee",
};

/**
 * Configuration Helper
 * Pulls settings from Script Properties or defaults to standard values.
 */
function _getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    TOKEN_TTL: parseInt(props.getProperty("TOKEN_TTL") || 8 * 60 * 60 * 1000),
    SECRET_KEY: props.getProperty("SECRET_KEY") || "SD_SYSTEM_DEFAULT_SECRET",
    DOMAIN: props.getProperty("ALLOWED_DOMAIN") || "@fob.ng",
    GEMINI_API_KEY: props.getProperty("GEMINI_API_KEY") || "",
    NAVY: props.getProperty("THEME_COLOR") || "#0d1526",
    AI_WELCOME_MESSAGE:
      props.getProperty("AI_WELCOME_MESSAGE") ||
      "Hello! I'm the SD-AI assistant. How can I help you today?",
    DEFAULT_REQUEST_TYPES:
      props.getProperty("DEFAULT_REQUEST_TYPES") ||
      "Software Request,Automation Idea,Feedback,Bug Report,Other",
    MEC_HEAD_NAME: props.getProperty("MEC_HEAD_NAME") || "MEC Head",
    MEC_HEAD_EMAIL: props.getProperty("MEC_HEAD_EMAIL") || "",
  };
}

/**
 * Internal: Gets or creates the Drive folder for PAT images
 */
function _getPATFolder() {
  var folderName = "PAT_PROJECT_IMAGES";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var folder = DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

/**
 * Internal: Gets or creates the Drive folder for general documents
 */
function _getDocumentsFolder() {
  var folderName = "PORTAL_DOCUMENTS";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  var folder = DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return folder;
}

/**
 * Utility to update system configuration via UI
 */
function ui_configureSystem() {
  var ui = SpreadsheetApp.getUi();
  var config = _getConfig();

  var domain = ui
    .prompt("Set Allowed Email Domain (currently: " + config.DOMAIN + "):")
    .getResponseText();
  var secret = ui
    .prompt(
      "Set System Secret Key (for security, currently: " +
        config.SECRET_KEY +
        "):",
    )
    .getResponseText();
  var aiWelcome = ui
    .prompt(
      "Set AI Assistant Welcome Message (currently: " +
        config.AI_WELCOME_MESSAGE +
        "):",
    )
    .getResponseText();

  if (domain)
    PropertiesService.getScriptProperties().setProperty(
      "ALLOWED_DOMAIN",
      domain,
    );
  var geminiKey = ui
    .prompt("Set Gemini API Key for AI tasks (leave blank to keep current):")
    .getResponseText();
  if (geminiKey)
    PropertiesService.getScriptProperties().setProperty(
      "GEMINI_API_KEY",
      geminiKey,
    );

  if (secret)
    PropertiesService.getScriptProperties().setProperty("SECRET_KEY", secret);
  if (aiWelcome)
    PropertiesService.getScriptProperties().setProperty(
      "AI_WELCOME_MESSAGE",
      aiWelcome,
    );

  // Initialize empty defaults if they don't exist
  if (
    !PropertiesService.getScriptProperties().getProperty("DEFAULT_ADMIN_DEPT")
  ) {
    PropertiesService.getScriptProperties().setProperty(
      "DEFAULT_ADMIN_DEPT",
      "Technology Support",
    );
  }
  // Add default request types if the sheet is empty
  if (
    !PropertiesService.getScriptProperties().getProperty(
      "DEFAULT_REQUEST_TYPES",
    )
  ) {
    PropertiesService.getScriptProperties().setProperty(
      "DEFAULT_REQUEST_TYPES",
      "Software Request,Automation Idea,Feedback,Bug Report,Other",
    );
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Configuration updated successfully.",
  );
}

// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("⚙️ SD PORTAL ADMIN")
      .addItem("🚀 Run First Setup", "runFirstSetup")
      .addSeparator()
      .addItem("👑 Create Super Admin", "ui_createSuperAdmin")
      .addItem("👤 Create Admin User", "ui_createAdmin")
      .addItem("➕ Add Material", "ui_addMaterial")
      .addItem("🔧 Configure System", "ui_configureSystem")
      .addItem("🔑 Reset User Password", "ui_resetPassword")
      .addItem("🤖 Setup Employee Check-ins", "setupEmployeeCheckins")
      .addItem("👤 Set MEC Head", "ui_setMecHead")
      .addItem("🗑️  Clear All Sessions", "clearAllSessions")
      .addItem("🔓 Force Unban by Email", "ui_forceUnbanByEmail")
      .addSeparator()
      .addItem("🧾 Seed Audit Department", "runFirstSetup")
      .addSeparator()
      .addItem("📊 View Live Stats", "ui_liveStats")
      .addItem("🏥 System Health Check", "ui_healthCheck")
      .addSeparator()
      .addItem("🧨 Reset & Wipe System", "ui_wipeSystem")
      .addToUi();
  } catch (e) {
    // This error happens when running onOpen from the script editor.
    // The menu only appears when you refresh the Spreadsheet itself.
  }
}

/**
 * UI Helper to set the MEC Head name for JCC signatures.
 */
function ui_setMecHead() {
  var ui = SpreadsheetApp.getUi();
  var currentHead =
    PropertiesService.getScriptProperties().getProperty("MEC_HEAD_NAME") ||
    "Not set";
  var newHead = ui
    .prompt("Set MEC Head Name (currently: " + currentHead + "):")
    .getResponseText()
    .trim();
  if (newHead) {
    PropertiesService.getScriptProperties().setProperty(
      "MEC_HEAD_NAME",
      newHead,
    );
    ui.alert("MEC Head set to: " + newHead);
  }
}

// ─────────────────────────────────────────────────────────────
// FIRST SETUP
// Creates all sheets with headers. Seeds default data.
// Safe to re-run — will not overwrite existing sheets.
// ─────────────────────────────────────────────────────────────
function runFirstSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = _getConfig();

  // 1. SD_USERS
  // UserID | Name | Email | PasswordHash | Salt | Role | Department | Gender | WorkflowNotes | Status | CreatedAt | LastLoginAt | AllowMessages | AIAutoAnswer | LoginIPs | Fingerprints | BanReason | BannedBy
  var users = _sheet(SD.USERS, [
    "UserID",
    "Name",
    "Email",
    "PasswordHash",
    "Salt",
    "Role",
    "Department",
    "Gender",
    "WorkflowNotes",
    "Status",
    "CreatedAt",
    "LastLoginAt",
    "AllowMessages",
    "AIAutoAnswer",
    "LoginIPs",
    "Fingerprints",
    "BanReason",
    "BannedBy",
  ]);
  _repairUserSheet(users);

  // 2. SD_DEPARTMENTS
  // DeptID | Name | HeadEmail | CreatedAt | CreatedBy
  var depts = _sheet(SD.DEPTS, [
    "DeptID",
    "Name",
    "HeadEmail",
    "CreatedAt",
    "CreatedBy",
  ]);

  // 3. SD_PAT_PROJECTS
  var pat = _sheet(SD.PAT, [
    "ProjectID",
    "ProjectName",
    "SiteAddress",
    "Lat",
    "Lon",
    "Phase",
    "WorkDescription",
    "Vendor",
    "InspectionDate",
    "Orchestrator",
    "ComplaintScore",
    "Status",
    "Verdict",
    "WorkflowStatus",
    "Checklist",
    "BOQ",
    "Snags",
    "Signoff",
    "Images",
    "HasVendor",
    "VendorToken",
    "VendorApprovalStatus",
    "VendorApprovalComments",
    "VendorApprovalDate",
    "SubmittedBy",
    "SubmittedByEmail",
    "SubmittedByDept",
    "SubmittedAt",
    "AssignedToName",
    "AssignedToEmail",
    "AssignedToDept",
    "RejectionReason",
    "WorkflowHistory",
    "Department",
    "UpdatedAt",
    "VendorEverApproved",
    "PresidingOfficer",
  ]);

  // 4. SD_MATERIALS
  var mats = _sheet(SD.MATS, ["MaterialID", "Description", "CreatedAt"]);

  // 5. SD_MAILS
  var mails = _sheet(SD.MAILS, [
    "MailID",
    "SenderEmail",
    "SenderName",
    "ReceiverEmail",
    "ReceiverName",
    "Subject",
    "Body",
    "Timestamp",
    "Status",
    "Folder",
    "Attachments",
    "Starred",
    "Labels",
    "ThreadID",
    "Priority",
    "CC",
    "BCC",
    "DeletedBy",
  ]);

  // Ensure mail sheet has all required columns (migration for existing sheets)
  if (mails && mails.getLastRow() > 0) {
    var mailHeaders = mails
      .getRange(1, 1, 1, mails.getLastColumn())
      .getValues()[0];
    var expectedHeaders = [
      "MailID",
      "SenderEmail",
      "SenderName",
      "ReceiverEmail",
      "ReceiverName",
      "Subject",
      "Body",
      "Timestamp",
      "Status",
      "Folder",
      "Attachments",
      "Starred",
      "Labels",
      "ThreadID",
      "Priority",
      "CC",
      "BCC",
      "DeletedBy",
    ];
    if (mailHeaders.length < expectedHeaders.length) {
      mails
        .getRange(1, 1, 1, expectedHeaders.length)
        .setValues([expectedHeaders]);
    }
  }

  // 6. SD_JCC_CERTIFICATES
  var jcc = _sheet(SD.JCC, [
    "JccID",
    "ProjectID",
    "ProjectName",
    "ProjectNumber",
    "CertificateType",
    "Vendor",
    "CertificateID",
    "Penalty",
    "Remarks",
    "MecName",
    "MecSignature",
    "MecDate",
    "VendorName",
    "VendorSignature",
    "VendorDate",
    "MecHeadName",
    "MecHeadSignature",
    "MecHeadDate",
    "GeneratedAt",
    "GeneratedBy",
    "GeneratedByEmail",
    "StateRegion",
    "Orchestrator",
    "PresidingOfficer",
  ]);

  // 7. SD_DOCUMENTS — General file/document storage (all departments)
  // DocID | ProjectID | FileName | FileType | FileSize | DriveURL | DriveFileID | UploadedBy | UploadedByEmail | UploadedByDept | UploadedAt | Category
  var docs = _sheet(SD.DOCS, [
    "DocID",
    "ProjectID",
    "FileName",
    "FileType",
    "FileSize",
    "DriveURL",
    "DriveFileID",
    "UploadedBy",
    "UploadedByEmail",
    "UploadedByDept",
    "UploadedAt",
    "Category",
  ]);

  // Style all header rows navy
  [users, depts, pat, mats, mails, jcc, docs].forEach(function (sh) {
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setBackground(config.NAVY)
      .setFontColor("#ffffff")
      .setFontWeight("bold");
  });

  // Seed a default super-admin if no users exist yet
  if (users.getLastRow() === 1) {
    var adminDept =
      PropertiesService.getScriptProperties().getProperty(
        "DEFAULT_ADMIN_DEPT",
      ) || "Technology Support";
    _createUserRow(
      "Portal Admin",
      "admin" + config.DOMAIN,
      "admin123",
      SD.ADMIN,
      adminDept,
    );
  }

  // Seed default materials if none exist
  if (mats.getLastRow() === 1) {
    var defaultMats = [
      "Fiber Cables",
      "Total Box – Box",
      "PCC Cable",
      "80m Box to Box PCC",
      "60m Box to Box PCC",
      "50m Box to Box PCC",
      "5m Box to Box PCC",
      "3m Box to Box PCC",
      "Aluminium Hook",
      "Belt & Clip",
      "Small Clamp",
      "Tiny Small",
      "½ Sub Box",
      "⅑ FAT6",
      "⅛ FAT2",
      "Installation of Cable Hanger",
      "Metallic Pole (Supply & Installation)",
    ];
    defaultMats.forEach(function (m) {
      mats.appendRow([_genId("MAT-"), m, new Date().toISOString()]);
    });
  }

  // Seed the Audit department + its portal admin account (idempotent)
  _seedAuditDepartment();

  ss.toast("✅ SD Portal Setup Complete — v1.0 Ready", "Setup", 5);
}

/**
 * Idempotently seed the Audit department and its portal admin account.
 * Safe to call repeatedly — only creates when missing.
 */
function _seedAuditDepartment() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = _getConfig();
    var deptSh = ss.getSheetByName(SD.DEPTS);
    if (deptSh && deptSh.getLastRow() >= 2) {
      var dData = deptSh.getDataRange().getValues().slice(1);
      var hasAudit = dData.some(function (r) {
        return String(r[1] || "").toLowerCase() === "audit";
      });
      if (!hasAudit)
        deptSh.appendRow([
          _genId("DEPT"),
          "Audit",
          "",
          new Date().toISOString(),
          "system",
        ]);
    }
    var userSh = ss.getSheetByName(SD.USERS);
    if (userSh && userSh.getLastRow() >= 2) {
      var uData = userSh.getDataRange().getValues().slice(1);
      var hasAuditAdmin = uData.some(function (r) {
        return String(r[2] || "").toLowerCase() === "audit" + config.DOMAIN;
      });
      if (!hasAuditAdmin) {
        _createUserRow(
          "Audit Admin",
          "audit" + config.DOMAIN,
          "audit123",
          SD.ADMIN,
          "Audit",
          "Other",
          "TRUE",
          "FALSE",
        );
      }
    }
  } catch (e) {
    console.warn("_seedAuditDepartment: " + e.message);
  }
}

/**
 * Clears ALL mails from SD_MAILS for fresh testing.
 * Keeps the header row intact. Only Super Admin can run this.
 */
function clearAllMails() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "⚠️ Clear All Mails",
    "This will delete ALL mail records from SD_MAILS. Are you sure?",
    ui.ButtonSet.YES_NO,
  );
  if (resp !== ui.Button.YES) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
  if (!sh) {
    ui.alert("SD_MAILS sheet not found. Run runFirstSetup first.");
    return;
  }
  var rows = sh.getLastRow();
  if (rows > 1) sh.deleteRows(2, rows - 1);
  ui.alert("✅ All mails cleared. SD_MAILS is ready for fresh testing.");
}

/**
 * Clears ALL data from all system sheets (users, mails, projects, docs, etc.)
 * Preserves sheet structure and header rows. Only for full reset.
 */
function clearAllSystemData() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "⚠️ Full System Reset",
    "This will DELETE ALL data from ALL system sheets. Are you absolutely sure?",
    ui.ButtonSet.YES_NO,
  );
  if (resp !== ui.Button.YES) return;
  var sheets = [
    { name: SD.MAILS, label: "Mails" },
    { name: SD.PAT, label: "Projects" },
    { name: SD.JCC, label: "JCC Certificates" },
    { name: SD.DOCS, label: "Documents" },
    { name: SD.MATS, label: "Materials" },
    { name: SD.BLACKLIST, label: "Blacklist" },
  ];
  var cleared = 0;
  sheets.forEach(function (s) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s.name);
    if (sh && sh.getLastRow() > 1) {
      sh.deleteRows(2, sh.getLastRow() - 1);
      cleared++;
    }
  });
  ui.alert(
    "✅ System data cleared. " +
      cleared +
      " sheets reset. Users and Departments preserved.",
  );
}

/**
 * Standardizes the SD_USERS sheet to the 12-column format.
 * Fixes malformed/shifted data from older versions.
 */
function _repairUserSheet(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var data = sh.getDataRange().getValues();
  var headers = [
    "UserID",
    "Name",
    "Email",
    "PasswordHash",
    "Salt",
    "Role",
    "Department",
    "Gender",
    "WorkflowNotes",
    "Status",
    "CreatedAt",
    "LastLoginAt",
    "AllowMessages",
    "AIAutoAnswer",
    "LoginIPs",
    "Fingerprints",
    "BanReason",
    "BannedBy",
  ];
  var repaired = [headers];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r.filter(String).length === 0) continue;

    var row = new Array(18).fill("");

    var emailIdx = -1;
    if (String(r[0]).indexOf("@") !== -1) emailIdx = 0;
    else if (String(r[1]).indexOf("@") !== -1) emailIdx = 1;
    else if (String(r[2]).indexOf("@") !== -1) emailIdx = 2;

    if (emailIdx === -1) continue;

    if (String(r[0]).startsWith("USR-")) {
      for (var j = 0; j < Math.min(r.length, 18); j++) row[j] = r[j];
      if (!row[7]) row[7] = "Other";
      if (!row[9]) row[9] = "active";
      if (!row[12]) row[12] = "TRUE";
      if (!row[13]) row[13] = "FALSE";
    } else if (emailIdx === 1) {
      row[0] = _genId("USR");
      row[1] = r[0]; // Name
      row[2] = r[1]; // Email
      row[3] = r[2]; // Hash
      row[4] = r[3]; // Salt
      row[5] = r[4]; // Role
      row[6] = r[5]; // Dept

      var val6 = String(r[6]).toLowerCase();
      if (val6 === "male" || val6 === "female") {
        row[7] = r[6];
        row[9] = r[7] || "active";
        row[10] = r[8] || "";
      } else {
        row[7] = "Other";
        row[9] = r[6] || "active";
        row[10] = r[7] || "";
        row[11] = r[8] || "";
        row[12] = "TRUE";
        row[13] = "FALSE";
      }
    } else {
      row[0] = _genId("USR");
      row[1] = r[emailIdx - 1] || "Unknown";
      row[2] = r[emailIdx];
      for (var k = 3; k < 14; k++) {
        if (r[k - 1]) row[k] = r[k - 1];
      }
    }

    repaired.push(row);
  }

  sh.clear();
  sh.getRange(1, 1, repaired.length, 18).setValues(repaired);
  var config = _getConfig();
  sh.getRange(1, 1, 1, 18)
    .setBackground(config.NAVY)
    .setFontColor("#ffffff")
    .setFontWeight("bold");
}

// Create a sheet only if it doesn't exist; returns the sheet either way
function _sheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    // Ensure getLastColumn() is not 0 for an existing sheet before getting range
    if (sh.getLastColumn() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      var curH = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      if (curH.length < headers.length)
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}

// ─────────────────────────────────────────────────────────────
// PASSWORD HELPERS
// ─────────────────────────────────────────────────────────────
function _salt() {
  return Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      String(new Date().getTime()) + String(Math.random()),
    ),
  ).substring(0, 16);
}

function _hash(password, salt) {
  var secret = _getConfig().SECRET_KEY;
  return Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password + salt + secret,
    ),
  );
}

// ─────────────────────────────────────────────────────────────
// SESSION SYSTEM  (PropertiesService — UserProperties)
//
// Why UserProperties?
//   • Scoped to the Google account running the script —
//     each user gets isolated storage, no sheet required.
//   • Survives page refreshes; cleared on logout or expiry.
//   • 9KB per key is more than enough for our token payload.
//
// Token format stored as a JSON string:
//   { email, name, role, department, expires (ms timestamp) }
//
// The token itself is a random 32-char hex string that we
// use as the property key: "token_<hex>"
// ─────────────────────────────────────────────────────────────

function _makeToken() {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(Math.random()) + String(new Date().getTime()),
  );
  return bytes
    .map(function (b) {
      return ("0" + (b & 0xff).toString(16)).slice(-2);
    })
    .join("")
    .substring(0, 32);
}

/**
 * Determines if a user is currently "active" based on last login.
 */
/**
 * Create a session. Stores the payload in UserProperties.
 * Returns the token string (sent to the browser via sessionStorage).
 */
function _createSession(email, name, role, department) {
  console.log("Creating session for: " + email);
  var lock = LockService.getUserLock();
  try {
    lock.waitLock(10000); // 10 second timeout
    var token = _makeToken();
    var config = _getConfig();
    var payload = JSON.stringify({
      email: email,
      name: name,
      role: role,
      department: department,
      expires: new Date().getTime() + config.TOKEN_TTL,
    });
    PropertiesService.getUserProperties().setProperty("tok_" + token, payload);
    // ── TOKEN INDEX (peak security) ──
    // Track which tokens belong to an email so a ban can force-logout ALL live sessions.
    try {
      var idxKey = "tokidx_" + email.toLowerCase();
      var idx = PropertiesService.getUserProperties().getProperty(idxKey);
      var list = idx ? idx.split(",") : [];
      if (list.indexOf(token) === -1) list.push(token);
      // Cap to last 10 sessions per user
      if (list.length > 10) list = list.slice(list.length - 10);
      PropertiesService.getUserProperties().setProperty(idxKey, list.join(","));
    } catch (e) {
      console.warn("token index write failed: " + e.message);
    }
  } catch (e) {
    throw new Error("Unable to create session: " + e.message);
  } finally {
    lock.releaseLock();
  }
  return token;
}

/**
 * Force-logout every active session for an email (used on ban).
 * @param {string} email
 */
function _purgeUserSessions(email) {
  try {
    var idxKey = "tokidx_" + String(email || "").toLowerCase();
    var idx = PropertiesService.getUserProperties().getProperty(idxKey);
    if (!idx) return;
    idx.split(",").forEach(function (t) {
      if (t) PropertiesService.getUserProperties().deleteProperty("tok_" + t);
    });
    PropertiesService.getUserProperties().deleteProperty(idxKey);
  } catch (e) {
    console.warn("purge sessions failed: " + e.message);
  }
}

/**
 * Validate a token. Returns the session object or throws.
 */
function _session(token) {
  if (!token) throw new Error("Not logged in. Please sign in again.");
  var raw = PropertiesService.getUserProperties().getProperty("tok_" + token);
  if (!raw) throw new Error("Session not found. Please sign in again.");
  try {
    var s = JSON.parse(raw);
  } catch (e) {
    throw new Error("Session data corrupted. Please sign in again.");
  }
  if (new Date().getTime() > s.expires) {
    PropertiesService.getUserProperties().deleteProperty("tok_" + token);
    throw new Error("Session expired. Please sign in again.");
  }
  // ── BAN ENFORCEMENT (peak security) ──
  // A banned account cannot use ANY endpoint, even with a valid token.
  // The authoritative status is read from the users sheet on every call.
  try {
    var bannedInfo = _getBanInfoFromDb(s.email);
    if (bannedInfo && bannedInfo.banned) {
      throw new Error(
        "__BANNED__::" +
          (bannedInfo.reason || "Account suspended by an administrator.") +
          "::" +
          (bannedInfo.bannedBy || "Administrator"),
      );
    }
  } catch (e) {
    if (String(e.message || "").indexOf("__BANNED__") === 0) throw e;
    // On lookup failure, fail open (do not lock out valid users) but log it.
    console.warn("_session ban check failed: " + e.message);
  }
  return s;
}

/**
 * Reads a user's current status + ban reason directly from the USERS sheet.
 * Returns { banned: bool, reason: string, bannedBy: string } or null on error.
 * @param {string} email
 */
function _getBanInfoFromDb(email) {
  var normEmail = String(email || "")
    .toLowerCase()
    .trim();
  if (!normEmail) return null;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (
      String(data[i][2] || "")
        .toLowerCase()
        .trim() === normEmail
    ) {
      var status = String(data[i][9] || "").toLowerCase();
      var reason = String(data[i][16] || ""); // BanReason column (index 16 / 1-based col 17)
      var bannedBy = String(data[i][17] || ""); // BannedBy column (index 17 / 1-based col 18)
      return {
        banned: status === "banned",
        reason: reason,
        bannedBy: bannedBy,
      };
    }
  }
  return null;
}

/**
 * Validate token AND require a specific role.
 */
function _adminSession(token) {
  var s = _session(token);
  var role = String(s.role || "").toLowerCase();
  if (role !== SD.ADMIN && role !== "super admin") {
    // Token might be stale — check DB for latest role
    var dbRole = _getUserRoleFromDb(s.email);
    if (dbRole && (dbRole === SD.ADMIN || dbRole === "super admin")) {
      s.role = dbRole;
      return s;
    }
    throw new Error("Admin access required.");
  }
  return s;
}

/**
 * Validate token AND require Super Admin role.
 */
function _superAdminSession(token) {
  var s = _session(token);
  var role = String(s.role || "").toLowerCase();
  if (role !== "super admin") {
    // Token might be stale — check DB for latest role
    var dbRole = _getUserRoleFromDb(s.email);
    if (dbRole === "super admin") {
      s.role = dbRole;
      return s;
    }
    throw new Error("Super Admin access required.");
  }
  return s;
}

/**
 * Look up a user's current role from the database (not from the token).
 * Used to handle role promotions without requiring re-login.
 */
function _getUserRoleFromDb(email) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    if (!sh) return null;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][2] || "")
          .toLowerCase()
          .trim() === String(email).toLowerCase().trim()
      ) {
        return String(data[i][5] || "").toLowerCase();
      }
    }
  } catch (e) {}
  return null;
}

/**
 * Delete a session (logout).
 */
function _destroySession(token) {
  if (token)
    PropertiesService.getUserProperties().deleteProperty("tok_" + token);
}

/** Public-facing logout — called from frontend. */
function logoutUser(token) {
  _destroySession(token);
  return { success: true };
}

/**
 * Menu action — wipes all "tok_*" keys from UserProperties.
 * Useful for testing or locked-out admins.
 */
function clearAllSessions() {
  var props = PropertiesService.getUserProperties().getProperties();
  Object.keys(props).forEach(function (k) {
    if (k.indexOf("tok_") === 0)
      PropertiesService.getUserProperties().deleteProperty(k);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "All sessions cleared.",
    "Sessions",
    3,
  );
}

// ─────────────────────────────────────────────────────────────
// ID / REF GENERATORS
// ─────────────────────────────────────────────────────────────
function _genId(prefix) {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var rand = "";
  for (var i = 0; i < 6; i++)
    rand += chars[Math.floor(Math.random() * chars.length)];
  return prefix + "-" + rand;
}

// ─────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
//  AUTH  —  called from auth.html via google.script.run
// ══════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────

/**
 * Sign up a new employee.
 * @param {string} name
 * @param {string} email     must end with @fob.ng
 * @param {string} password  min 4 chars
 * @param {string} department
 * @param {string} gender
 * @returns {{ success, message? }}
 */
/**
 * Admin-only: Creates a user with a specific role and gender.
 */
function adminCreateUser(
  token,
  name,
  email,
  password,
  department,
  role,
  gender,
) {
  try {
    _superAdminSession(token); // ONLY super admin can create users (incl. admins/audit)
    var config = _getConfig();
    email = String(email || "")
      .toLowerCase()
      .trim();
    role = String(role || SD.EMPLOYEE)
      .toLowerCase()
      .trim();
    gender = String(gender || "Other").trim();

    // Audit is a first-class role, not a normal department. When role === 'audit',
    // force the department to 'Audit' so the audit console recognises it.
    var effectiveDept = department;
    if (role === "audit") effectiveDept = "Audit";

    if (!name || !email || !password || !effectiveDept)
      throw new Error("Name, email, password, and department are required.");
    if (!email.endsWith(config.DOMAIN))
      throw new Error("Only " + config.DOMAIN + " company emails are allowed.");

    // Validate role values
    var allowedRoles = ["employee", "admin", "super admin", "audit"];
    if (allowedRoles.indexOf(role) === -1)
      throw new Error(
        "Invalid role. Allowed: employee, admin, audit, super admin.",
      );

    // Self-protection: never allow creating a second super admin through this
    // path is fine (super admin only), but block creating an audit/employee as admin.
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var dup = data.slice(1).find(function (r) {
      return (
        String(r[2] || "")
          .toLowerCase()
          .trim() === String(email).toLowerCase().trim()
      );
    });
    if (dup) throw new Error("An account with this email already exists.");

    _createUserRow(
      name,
      email,
      password,
      role,
      effectiveDept,
      gender,
      "TRUE",
      "FALSE",
    );
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Save images to a project (lightweight — only updates images column).
 */
function saveProjectImages(token, projectId, images) {
  try {
    _session(token);
    var c = _patCols();
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };

    var data = sh
      .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .getValues();
    var searchId = String(projectId)
      .trim()
      .replace(/^(FOB|PAT)-/i, "");
    var foundRow = -1;

    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0])
        .trim()
        .replace(/^(FOB|PAT)-/i, "");
      if (rowId.toUpperCase() === searchId.toUpperCase()) {
        foundRow = i;
        break;
      }
    }
    if (foundRow === -1)
      return { success: false, message: "Project not found: " + projectId };

    var sheetRow = foundRow + 2;
    sh.getRange(sheetRow, c.IMAGES + 1).setValue(JSON.stringify(images || []));
    return {
      success: true,
      message: "Images saved to project.",
      imageCount: (images || []).length,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function registerUser(name, email, password, department, gender) {
  var lock = LockService.getScriptLock();
  try {
    var config = _getConfig();
    email = String(email || "")
      .toLowerCase()
      .trim();
    if (!name || !email || !password || !department || !gender)
      throw new Error("All fields are required.");
    if (!email.endsWith(config.DOMAIN))
      throw new Error("Only " + config.DOMAIN + " company emails are allowed.");
    if (String(password).length < 4)
      throw new Error("Password must be at least 4 characters.");

    // ── BANNED EMAIL GUARD ──
    // A banned email (blacklist entry or a user row already marked 'banned')
    // must never be allowed to mint a fresh account. Redirect the applicant
    // to the same ban page a banned login would hit.
    var sh0 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var rows0 = sh0.getDataRange().getValues();
    var existing = rows0.slice(1).find(function (r) {
      return (
        String(r[2] || "")
          .toLowerCase()
          .trim() === email
      );
    });
    if (existing && String(existing[9] || "").toLowerCase() === "banned") {
      var _breason = String(
        existing[16] || "Your account has been suspended by an administrator.",
      );
      var _bby = String(existing[17] || "Administrator");
      throw new Error("__BANNED__::" + _breason + "::" + _bby);
    }
    var _bl = _checkBlacklistMatch(email, "", null);
    if (_bl) {
      throw new Error(
        "__BANNED__::" +
          (_bl.reason || "This email has been blocked.") +
          "::" +
          (_bl.bannedBy || "Administrator"),
      );
    }

    lock.waitLock(15000);

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var dup = data.slice(1).find(function (r) {
      return (
        String(r[2] || "")
          .toLowerCase()
          .trim() === String(email).toLowerCase().trim()
      );
    });
    if (dup) {
      // Signal the frontend (with a machine-readable prefix) so it can route
      // the applicant to the "email already in system → contact admin" page
      // instead of a generic toast.
      throw new Error("__EMAIL_EXISTS__::" + email);
    }

    _createUserRow(name, email, password, SD.EMPLOYEE, department, gender);

    lock.releaseLock();
    return { success: true };
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, message: e.message };
  }
}

function _createUserRow(
  name,
  email,
  password,
  role,
  department,
  gender,
  allowMsgs,
  aiAuto,
) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  var s = _salt();
  var uid = _genId("USR");
  sh.appendRow([
    uid,
    name,
    email.toLowerCase(),
    _hash(password, s),
    s,
    role,
    department,
    gender,
    "",
    "active",
    new Date().toISOString(),
    "",
    allowMsgs || "TRUE",
    aiAuto || "FALSE",
  ]);
  return uid;
}

/**
 * Rate-limiter: tracks failed login attempts per email.
 * Stores in ScriptProperties: login_lock_<email> = JSON timestamp
 * Locks out for 5 minutes after 5 consecutive failures.
 */
function _checkLoginRateLimit(email) {
  var normEmail = String(email || "")
    .toLowerCase()
    .trim();
  if (!normEmail) return { allowed: true };
  var props = PropertiesService.getScriptProperties();
  var lockKey = "login_lock_" + normEmail;
  var lockDataRaw = props.getProperty(lockKey);
  if (!lockDataRaw) return { allowed: true };
  try {
    var lockData = JSON.parse(lockDataRaw);
    var now = new Date().getTime();
    if (now < lockData.until) {
      var remainingSec = Math.ceil((lockData.until - now) / 1000);
      return {
        allowed: false,
        message:
          "Too many failed attempts. Try again in " +
          remainingSec +
          " seconds.",
        retryAfter: lockData.until,
      };
    }
    props.deleteProperty(lockKey);
  } catch (e) {
    props.deleteProperty(lockKey);
  }
  return { allowed: true };
}

function _recordLoginFailure(email) {
  var normEmail = String(email || "")
    .toLowerCase()
    .trim();
  if (!normEmail) return;
  var props = PropertiesService.getScriptProperties();
  var lockKey = "login_lock_" + normEmail;
  var lockDataRaw = props.getProperty(lockKey);
  var attempts = 1;
  if (lockDataRaw) {
    try {
      var existing = JSON.parse(lockDataRaw);
      attempts = (existing.attempts || 0) + 1;
    } catch (e) {}
  }
  if (attempts >= 5) {
    var lockUntil = new Date().getTime() + 5 * 60 * 1000;
    props.setProperty(
      lockKey,
      JSON.stringify({ attempts: attempts, until: lockUntil }),
    );
  } else {
    props.setProperty(
      lockKey,
      JSON.stringify({ attempts: attempts, until: 0 }),
    );
  }
}

function _clearLoginLock(email) {
  var normEmail = String(email || "")
    .toLowerCase()
    .trim();
  if (!normEmail) return;
  PropertiesService.getScriptProperties().deleteProperty(
    "login_lock_" + normEmail,
  );
}

/**
 * Blacklist enforcement at LOGIN.
 * Returns an active match { banned:true, matchedOn, reason, bannedBy } or null.
 * Checks the SD_BLACKLIST sheet for:
 *   - the exact email (or an "EMAIL:..." entry),
 *   - the client IP (exact OR same /24 subnet = "similar" IP),
 *   - the device fingerprint (best-effort: exact compositeHash or substring).
 * This blocks a banned user even if they try a brand-new account from the
 * same machine / same network.
 * @param {string} email
 * @param {string} clientIp
 * @param {object} fingerprint
 */
function _checkBlacklistMatch(email, clientIp, fingerprint) {
  try {
    var sh = _blacklistSheet();
    if (!sh || sh.getLastRow() < 2) return null;
    var c = _blacklistCols();
    var data = sh.getDataRange().getValues();
    var normEmail = String(email || "")
      .toLowerCase()
      .trim();
    var normIp = String(clientIp || "").trim();
    var fpHash =
      fingerprint && fingerprint.compositeHash
        ? String(fingerprint.compositeHash)
        : "";
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[c.UNBANNED_AT]) continue; // unbanned -> ignore
      var rowEmail = String(row[c.USER_EMAIL] || "").toLowerCase();
      var rowIpRaw = String(row[c.IP] || "");
      var rowFp = String(row[c.DEVICE_FINGERPRINT] || "");
      var reason = String(
        row[c.REASON] || "This account or device has been blocked.",
      );
      var bannedBy = String(row[c.BANNED_BY] || "Administrator");

      // EMAIL match (userEmail column, or "EMAIL:..." ip column)
      if (normEmail) {
        if (rowEmail === normEmail) return _blMatch("email", reason, bannedBy);
        if (
          rowIpRaw.indexOf("EMAIL:") === 0 &&
          rowIpRaw.slice(6).toLowerCase() === normEmail
        ) {
          return _blMatch("email", reason, bannedBy);
        }
      }

      // IP match — exact OR same /24 subnet ("similar" IP)
      if (
        normIp &&
        normIp !== "unknown" &&
        rowIpRaw &&
        rowIpRaw.indexOf("EMAIL:") !== 0
      ) {
        if (rowIpRaw === normIp) return _blMatch("ip", reason, bannedBy);
        var a = normIp.split("."),
          b = rowIpRaw.split(".");
        if (
          a.length === 4 &&
          b.length === 4 &&
          a[0] === b[0] &&
          a[1] === b[1] &&
          a[2] === b[2]
        ) {
          return _blMatch("ip-subnet", reason, bannedBy);
        }
      }

      // DEVICE/FP match (best-effort)
      if (
        fpHash &&
        rowFp &&
        (rowFp === fpHash || rowFp.indexOf(fpHash) !== -1)
      ) {
        return _blMatch("device", reason, bannedBy);
      }
    }
    return null;
  } catch (e) {
    console.warn("_checkBlacklistMatch error: " + e.message);
    return null; // fail open — never lock out valid users on a lookup error
  }
}

function _blMatch(on, reason, bannedBy) {
  return { banned: true, matchedOn: on, reason: reason, bannedBy: bannedBy };
}

/**
 * Log in.
 * @param {string} email
 * @param {string} password
 * @returns {{ success, token?, role?, name?, department?, message? }}
 */
function loginUser(email, password, clientIp, fingerprint) {
  try {
    console.log("Login attempt: " + email);
    email = String(email).toLowerCase().trim();
    if (!email) throw new Error("Email is required.");

    var rateLimit = _checkLoginRateLimit(email);
    if (!rateLimit.allowed) {
      return { success: false, message: rateLimit.message };
    }

    // Passwordless login for test/department accounts
    var TEST_ACCOUNTS = {
      "mec-test@fob.ng": { role: "admin", dept: "MEC" },
      "project-test@fob.ng": { role: "employee", dept: "Project Team" },
      "planning-test@fob.ng": { role: "employee", dept: "Planning Team" },
      "sd-test@fob.ng": { role: "employee", dept: "Service Delivery / Metro" },
    };
    if (TEST_ACCOUNTS[email]) {
      var ta = TEST_ACCOUNTS[email];
      // Blacklist check (IP / device / email) — block banned machines even for test accounts
      var bl = _checkBlacklistMatch(email, clientIp, fingerprint);
      if (bl) {
        throw new Error("__BANNED__::" + bl.reason + "::" + bl.bannedBy);
      }
      _recordLoginIp(email, clientIp || "unknown", null, fingerprint);
      var token = _createSession(email, email.split("@")[0], ta.role, ta.dept);
      _clearLoginLock(email);
      return {
        success: true,
        token: token,
        role: ta.role,
        name: email.split("@")[0],
        department: ta.dept,
      };
    }

    if (!password) throw new Error("Password is required for this account.");

    // ── BLACKLIST ENFORCEMENT AT LOGIN ──
    // If this machine/IP/fingerprint is blacklisted, block the login outright
    // even if the email itself is brand-new (banned users can't just switch
    // accounts from the same device or same network). Reuses the same
    // "__BANNED__::reason::by" signal the frontend already renders as a page.
    var blMatch = _checkBlacklistMatch(email, clientIp, fingerprint);
    if (blMatch) {
      throw new Error(
        "__BANNED__::" + blMatch.reason + "::" + blMatch.bannedBy,
      );
    }

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var lastRow = sh.getLastRow();
    if (!sh || lastRow < 2) return { success: true, data: [] };

    var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (
        String(r[2] || "")
          .toLowerCase()
          .trim() !== email
      )
        continue;
      if (r[9] === "banned") {
        var _banReason = String(r[16] || "");
        var _bannedBy = String(r[17] || "");
        throw new Error(
          "__BANNED__::" +
            (_banReason ||
              "Your account has been suspended by an administrator.") +
            "::" +
            (_bannedBy || "Administrator"),
        );
      }
      if (_hash(String(password), String(r[4])) !== String(r[3])) {
        _recordLoginFailure(email);
        throw new Error("Incorrect password.");
      }

      sh.getRange(i + 2, 12).setValue(new Date().toISOString()); // Update login time
      _recordLoginIp(email, clientIp || "unknown", i + 2, fingerprint);
      var token = _createSession(
        email,
        String(r[1]),
        String(r[5]),
        String(r[6]),
      );
      _clearLoginLock(email);
      // Lazily ensure the Audit department + audit admin exist (idempotent)
      try {
        _seedAuditDepartment();
      } catch (e) {}
      return {
        success: true,
        token: token,
        role: r[5],
        name: r[1],
        department: r[6],
        gender: r[7],
        notes: r[8],
        allowMessages: r[12] === "TRUE",
        aiAutoAnswer: r[13] === "TRUE",
        deviceFingerprint: fingerprint || null,
      };
    }
    _recordLoginFailure(email);
    throw new Error("No account found with that email address.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Record the client IP for a user login attempt.
 * Stores as a JSON array of { ip, time } objects in column 14.
 * Stores fingerprints in column 16.
 * @param {string} email
 * @param {string} ip
 * @param {number} row - optional sheet row (1-based). If omitted, looks up by email.
 * @param {object} fingerprint - advanced browser fingerprint object
 */
function _recordLoginIp(email, ip, row, fingerprint) {
  try {
    if (!email) {
      console.error("_recordLoginIp: no email provided");
      return;
    }
    var safeIp = ip || "unknown";
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    if (!sh) {
      console.error("_recordLoginIp: users sheet not found");
      return;
    }

    if (!row) {
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (
          String(data[i][2] || "")
            .toLowerCase()
            .trim() === email.toLowerCase().trim()
        ) {
          row = i + 1;
          break;
        }
      }
    }
    if (!row) {
      console.error("_recordLoginIp: user not found for " + email);
      return;
    }

    var existing = sh.getRange(row, 15).getValue() || "";
    var ips = [];
    try {
      ips = existing ? JSON.parse(existing) : [];
    } catch (e) {
      ips = [];
    }

    ips.push({ ip: safeIp, time: new Date().toISOString() });
    if (ips.length > 20) ips = ips.slice(ips.length - 20);

    sh.getRange(row, 15).setValue(JSON.stringify(ips));
    console.log(
      "IP recorded for " +
        email +
        ": " +
        safeIp +
        " (total: " +
        ips.length +
        ")",
    );

    // Record fingerprint in column 16
    if (fingerprint) {
      var fpExisting = sh.getRange(row, 16).getValue() || "";
      var fps = [];
      try {
        fps = fpExisting ? JSON.parse(fpExisting) : [];
      } catch (e) {
        fps = [];
      }

      fps.push({
        fingerprint: fingerprint,
        recordedAt: new Date().toISOString(),
      });
      if (fps.length > 10) fps = fps.slice(fps.length - 10);

      sh.getRange(row, 16).setValue(JSON.stringify(fps));
      console.log(
        "Fingerprint recorded for " +
          email +
          " (total fingerprints: " +
          fps.length +
          ")",
      );
    }
  } catch (e) {
    console.error("Failed to record login IP: " + e.message);
  }
}

/**
 * Get the current user's session info from the backend (authoritative).
 * The frontend uses this instead of sessionStorage for permission decisions.
 */
function getMySessionInfo(token) {
  var sess = _session(token);
  return {
    success: true,
    email: sess.email,
    name: sess.name,
    role: sess.role,
    department: sess.department,
  };
}

/**
 * Get OpenRouter API key (only accessible to admins)
 */
function getOpenRouterKey(token) {
  try {
    _adminSession(token); // Only admins/super admins can access this
    var config = _getConfig();
    var apiKey =
      PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) {
      throw new Error("OpenRouter API key not configured");
    }
    return { success: true, key: apiKey };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// PAT PERMISSION SYSTEM — authoritative backend checks
// These functions use the deployed project's PAT data and the
// session system to determine what actions a user can take.
// ─────────────────────────────────────────────────────────────

/**
 * Owner map: which department "owns" each workflow status.
 */
var PAT_OWNER_MAP = {
  Draft: "mec",
  Rejected: "mec",
  "Awaiting Project Team": "project team",
  "Awaiting Service Delivery (Metro)": "service delivery / metro", // metro-route entry = replaces Project Team
  "Awaiting Planning Team": "planning team",
  "Awaiting Service Delivery": "service delivery / metro",
  "Awaiting Final MEC Review": "mec",
  "Awaiting MEC Recheck": "mec",
  "Partially Approved": "__all__",
  Completed: "__all__",
};

var STATUS_TO_DEPT = {
  Draft: "MEC",
  Rejected: "MEC",
  "Awaiting Project Team": "Project Team",
  "Awaiting Service Delivery (Metro)": "Service Delivery / Metro",
  "Awaiting Planning Team": "Planning Team",
  "Awaiting Service Delivery": "Service Delivery / Metro",
  "Awaiting Final MEC Review": "MEC",
  "Awaiting MEC Recheck": "MEC",
  "Partially Approved": "Project Team",
  Completed: "MEC",
};

/**
 * Linear workflow (happy path) — PROJECT TEAM route.
 * MEC → Project Team → Planning → SD/Metro → Final MEC Review → Completed.
 */
var WF_LINEAR_FLOW = {
  Draft: "Awaiting Project Team",
  "Awaiting Project Team": "Awaiting Planning Team",
  "Awaiting Planning Team": "Awaiting Service Delivery",
  "Awaiting Service Delivery": "Awaiting Final MEC Review",
  "Awaiting Final MEC Review": "Completed",
};

/**
 * Linear workflow (happy path) — SERVICE DELIVERY / METRO route.
 * MEC → SD/Metro → Planning → Final MEC Review → Completed.
 * The SD workflow has NO Project Team stage (it is skipped). After SD/Metro the
 * project moves to Planning for design validation, then straight to Final MEC Review.
 * Vendor approval is handled inside SD/Metro before it advances to Planning.
 * (The PROJECT route is the one that includes the Project Team: see WF_LINEAR_FLOW.)
 */
var WF_LINEAR_FLOW_METRO = {
  Draft: "Awaiting Service Delivery (Metro)",
  "Awaiting Service Delivery (Metro)": "Awaiting Planning Team",
  "Awaiting Planning Team": "Awaiting Final MEC Review",
  "Awaiting Final MEC Review": "Completed",
};

/**
 * Returns the linear-flow map for a given entry route ('project' | 'metro').
 */
function _linearFlowFor(route) {
  return route === "metro" ? WF_LINEAR_FLOW_METRO : WF_LINEAR_FLOW;
}

/**
 * Returns the first (entry) linear status for a route.
 */
function _firstStageOf(route) {
  var f = _linearFlowFor(route);
  return f["Draft"] || "Awaiting Project Team";
}

/**
 * Returns the next linear status for a given current status + route.
 */
function _nextStageOf(status, route) {
  return _linearFlowFor(route)[status] || null;
}

/**
 * Derive a project's entry route from its workflow history (first push out of Draft).
 * Stored as entryRoute on the history entry — backward compatible with old PATs
 * (which had no route: default to 'project').
 */
function _entryRouteOf(project) {
  var history = (project && project.workflowHistory) || [];
  for (var i = 0; i < history.length; i++) {
    var h = history[i];
    if (h && h.entryRoute) return h.entryRoute; // 'project' | 'metro'
  }
  return "project"; // default for legacy PATs
}

/**
 * Recovery/retry transitions for rejected and recheck states.
 */
var WF_RECOVERY_FLOW = {
  Rejected: "Awaiting Project Team",
  "Awaiting MEC Recheck": "Awaiting Project Team",
};

/**
 * Rejection triggers (PROJECT route baseline, used for graph display + fallback).
 * For route-aware rejection targets use _rejectionTarget(status, route).
 */
var WF_REJECTION_TRIGGERS = {
  "Awaiting Project Team": "Rejected",
  "Awaiting Planning Team": "Awaiting Project Team",
  "Awaiting Service Delivery": "Awaiting Project Team",
  "Awaiting Final MEC Review": "Awaiting Project Team",
  "Awaiting MEC Recheck": "Awaiting MEC Recheck",
};

/**
 * Route-aware rejection target.
 * metro route: Planning/SD rejections cycle within SD/Metro (no Project Team).
 * project route: standard behavior (Planning → Project Team).
 */
function _rejectionTarget(status, route) {
  if (route === "metro") {
    // Metro route has NO Project Team: rejections cycle within the metro chain.
    if (status === "Awaiting Planning Team")
      return "Awaiting Service Delivery (Metro)";
    if (status === "Awaiting Service Delivery (Metro)") return "Rejected";
    if (status === "Awaiting Final MEC Review")
      return "Awaiting Service Delivery (Metro)";
    if (status === "Awaiting MEC Recheck") return "Awaiting MEC Recheck";
    if (status === "Rejected") return "Awaiting Service Delivery (Metro)";
  }
  return WF_REJECTION_TRIGGERS[status] || "Rejected";
}

/**
 * Combined stage transition map used by submitPATToNextStage, partiallyApprovePAT, etc.
 * Maps each workflow status to the next status in the linear (or recovery) path.
 * Now route-aware: use _wfStageFlow(route) for correct metro-route transitions.
 * WF_STAGE_FLOW remains as the PROJECT-route default for any legacy callers.
 */
var WF_STAGE_FLOW = {};
(function () {
  var k;
  for (k in WF_LINEAR_FLOW) WF_STAGE_FLOW[k] = WF_LINEAR_FLOW[k];
  for (k in WF_RECOVERY_FLOW) WF_STAGE_FLOW[k] = WF_RECOVERY_FLOW[k];
})();

/**
 * Route-aware stage flow: merges the route's linear map + recovery map.
 */
function _wfStageFlow(route) {
  var flow = {};
  var lin = _linearFlowFor(route);
  for (var k in lin) flow[k] = lin[k];
  for (var r in WF_RECOVERY_FLOW) flow[r] = WF_RECOVERY_FLOW[r];
  // Metro route: a rejected metro entry returns to its own entry stage, not Project Team
  if (route === "metro") flow["Rejected"] = "Awaiting Service Delivery (Metro)";
  return flow;
}

/**
 * Returns the complete workflow graph definition derived from the backend logic.
 * The AI and frontend use this to programmatically construct workflow graph instances.
 *
 * @param {string} token - Admin session token
 * @returns {Object} { success, stages, linearOrder, transitions, rejectStages, ownerMap, stageToDept }
 */
function getWorkflowGraphConfig(token) {
  try {
    if (token) _adminSession(token);

    // Derive linear order from WF_LINEAR_FLOW by traversing from Draft
    var linearOrder = ["Draft"];
    var currentStage = "Draft";
    while (WF_LINEAR_FLOW[currentStage]) {
      linearOrder.push(WF_LINEAR_FLOW[currentStage]);
      currentStage = WF_LINEAR_FLOW[currentStage];
    }

    // Build all known stages from PAT_OWNER_MAP keys
    var allStatuses = Object.keys(PAT_OWNER_MAP);
    var stages = [];
    var seen = {};
    var displayLabels = {
      Draft: "Draft",
      "Awaiting Project Team": "Awaiting<br/>Project Team",
      "Awaiting Service Delivery (Metro)": "Awaiting<br/>SD / Metro",
      "Awaiting Planning Team": "Awaiting<br/>Planning Team",
      "Awaiting Service Delivery": "Awaiting<br/>Service Delivery",
      "Awaiting Final MEC Review": "Awaiting<br/>Final MEC Review",
      Completed: "Completed",
      Rejected: "Rejected",
      "Awaiting MEC Recheck": "Awaiting<br/>MEC Recheck",
      "Partially Approved": "Partially<br/>Approved",
    };

    var rejectStages = ["Rejected", "Awaiting MEC Recheck"];
    var linearSet = {};
    linearOrder.forEach(function (s) {
      linearSet[s] = true;
    });

    allStatuses.forEach(function (status) {
      if (seen[status]) return;
      seen[status] = true;
      var dept = STATUS_TO_DEPT[status] || "";
      var owner = PAT_OWNER_MAP[status] || "";
      var order = linearOrder.indexOf(status);
      stages.push({
        id: status,
        label: displayLabels[status] || status,
        dept: dept,
        owner: owner,
        order: order >= 0 ? order : -1,
        isLinear: order >= 0,
        isReject: rejectStages.indexOf(status) !== -1,
      });
    });

    // Build complete transition list from all workflow data
    var transitions = [];

    // Linear transitions
    linearOrder.slice(0, -1).forEach(function (from, i) {
      var to = linearOrder[i + 1];
      transitions.push({ from: from, to: to, type: "linear" });
    });

    // Rejection triggers
    Object.keys(WF_REJECTION_TRIGGERS).forEach(function (from) {
      var to = WF_REJECTION_TRIGGERS[from];
      transitions.push({ from: from, to: to, type: "rejection" });
      // Also add a "return" edge from reject back (recovery path)
      if (from !== to) {
        transitions.push({ from: to, to: from, type: "review" });
      }
    });

    // Recovery transitions from Rejected / MEC Recheck
    Object.keys(WF_RECOVERY_FLOW).forEach(function (from) {
      var to = WF_RECOVERY_FLOW[from];
      // Add recovery edges only for routes not already covered
      var exists = transitions.some(function (t) {
        return t.from === from && t.to === to;
      });
      if (!exists) {
        transitions.push({ from: from, to: to, type: "recovery" });
      }
    });

    return {
      success: true,
      stages: stages,
      linearOrder: linearOrder,
      transitions: transitions,
      rejectStages: rejectStages,
      ownerMap: PAT_OWNER_MAP,
      stageToDept: STATUS_TO_DEPT,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Department name normalization.
 */
var PAT_DEPT_MAP = {
  mec: "mec",
  mech: "mec",
  "project team": "project team",
  project: "project team",
  "planning team": "planning team",
  planning: "planning team",
  "service delivery": "service delivery / metro",
  "service delivery / metro": "service delivery / metro",
  metro: "service delivery / metro",
};

/**
 * Normalize a department name for comparison.
 */
function _normalizeDept(dept) {
  var key = String(dept || "")
    .toLowerCase()
    .trim();
  return PAT_DEPT_MAP[key] || key;
}

/**
 * Check if a user can perform workflow actions on a given project.
 * This is the authoritative backend version — the frontend should
 * trust this result over local sessionStorage calculations.
 *
 * @param {object} sess - Session object from _session(token)
 * @param {object} project - PAT project object
 * @returns {boolean}
 */
/**
 * Check if a user can view a specific PAT project (prevents IDOR).
 * Used by getPATProjectById to re-validate access for external callers.
 */
function _userCanAccessProject(sess, project) {
  // All authenticated users can view any PAT project
  // Permission to ACT on the project is controlled separately via _userCanActOnProject
  if (!sess) return false;
  return true;
}

function _userCanActOnProject(sess, project) {
  if (!project) return false;

  var role = String(sess.role || "").toLowerCase();
  var dept = String(sess.department || "").toLowerCase();
  var email = String(sess.email || "").toLowerCase();

  if (role === "super admin") return true;

  var status = project.workflowStatus || "Draft";
  if (status === "Completed") return false;

  var myDeptNorm = _normalizeDept(dept);
  var assignedDeptNorm = _normalizeDept(project.assignedToDept || "");
  var assignedEmail = String(project.assignedToEmail || "").toLowerCase();

  // Stage-based turn locking: enforce strict ownership
  if (status === "Draft") return myDeptNorm === "mec";
  if (status === "Rejected") return myDeptNorm === "mec";

  // Standard assignment check
  if (assignedEmail === email && !!email) return true;
  if (assignedDeptNorm === myDeptNorm && !!myDeptNorm) return true;

  return false;
}

/**
 * Check if a user can edit (modify data of) a PAT project.
 * Used for image upload/delete permission checks on the server side.
 * Logic mirrors the frontend's canEditCurrentProject().
 *
 * @param {object} sess - Session object from _session(token)
 * @param {string} projectId - PAT project ID
 * @returns {boolean}
 */
function _canEditPATProject(sess, projectId) {
  try {
    // Fetch project with null token (internal call, no session check needed)
    var res = getPATProjectById(null, projectId);
    if (!res || !res.success) return false;
    var project = res.project;
    if (!project) return false;

    var role = String(sess.role || "").toLowerCase();
    var dept = String(sess.department || "").toLowerCase();

    // Super admin can always edit
    if (role === "super admin") return true;

    // Only MEC department can edit images
    var isMEC = dept === "mec" || dept === "mech";
    if (!isMEC) return false;

    // MEC can only edit when project is in one of their editable states
    var status = project.workflowStatus || "Draft";
    var EDITABLE_STATUSES = [
      "Draft",
      "Rejected",
      "Awaiting Final MEC Review",
      "Awaiting MEC Recheck",
    ];
    return EDITABLE_STATUSES.indexOf(status) !== -1;
  } catch (e) {
    return false;
  }
}

/**
 * Get an authoritative canAct value for a single project.
 * Called by the frontend when opening a project to verify permissions.
 */
function checkPATPermission(token, projectId) {
  var sess = _session(token);
  var role = String(sess.role || "").toLowerCase();

  // Call with null token for internal read — no session check, no permission filter
  var result = getPATProjectById(null, projectId);
  if (!result || !result.success) {
    return { success: false, message: "Project not found.", canAct: false };
  }

  var project = result.project;
  var canAct = _userCanActOnProject(sess, project);
  var dept = String(sess.department || "").toLowerCase();
  var normalizedDept = _normalizeDept(dept);
  var owner = PAT_OWNER_MAP[project.workflowStatus || "Draft"] || "";

  return {
    success: true,
    canAct: canAct,
    isPushedOut: !canAct && role !== "admin" && role !== "super admin",
    currentOwner: owner,
    userDepartment: dept,
  };
}

/**
 * Check if a user can edit/delete a specific snag item.
 * @param {object} sess - Session object from _session(token)
 * @param {object} snagItem - The snag item object
 * @param {object} project - The full PAT project object
 * @returns {boolean}
 */
function _canEditSnag(sess, snagItem, project) {
  if (!sess || !snagItem || !project) return false;
  const role = String(sess.role || "").toLowerCase();
  if (role === "super admin") return true;

  const userEmail = String(sess.email || "").toLowerCase();
  const snagAddedByEmail = String(snagItem.addedByEmail || "").toLowerCase();

  // Rule: You can ONLY edit/delete snags you added yourself.
  if (userEmail === snagAddedByEmail && !!userEmail) return true;

  // MEC Rule: MEC can manage auto-filled/system snags IF it is their turn (Hub control)
  const isMEC = _normalizeDept(sess.department) === "mec";
  const status = project.workflowStatus || "Draft";
  const isMECControl =
    status === "Draft" ||
    status === "Rejected" ||
    status === "Awaiting Final MEC Review" ||
    status === "Awaiting MEC Recheck";

  const isAutoSnag =
    !snagAddedByEmail ||
    snagAddedByEmail.includes("system") ||
    String(snagItem.addedBy).includes("System Generated");

  if (isMEC && isMECControl && isAutoSnag) return true;

  return false;
}

/**
 * Annotate all PAT projects with canAct permissions from the backend.
 * Replaces the frontend-only permission calculation with authoritative server-side checks.
 * This function calls the deployed getPATProjects() internally.
 */
function getPATProjectsWithPermissions(token) {
  var sess = _session(token);
  var role = String(sess.role || "").toLowerCase();
  var dept = String(sess.department || "").toLowerCase();
  var email = String(sess.email || "").toLowerCase();
  var normalizedDept = _normalizeDept(dept);
  var isSuperAdmin = role === "super admin";

  // Get all projects from the deployed function
  var result = getPATProjects();
  if (!result || !result.success) {
    return result;
  }

  var projects = (result.projects || []).map(function (p) {
    var canAct = false;
    var status = p.workflowStatus || "Draft";
    var owner = PAT_OWNER_MAP[status] || "";

    if (isSuperAdmin) {
      canAct = true;
    } else if (owner === "__all__") {
      canAct = false;
    } else if (owner === normalizedDept) {
      canAct = true;
    } else if (
      p.assignedToEmail &&
      String(p.assignedToEmail).toLowerCase() === email
    ) {
      canAct = true;
    }

    // Backend-authoritative permission flag
    p.canAct = canAct;
    p._isBackendVerified = true;

    return p;
  });

  return {
    success: true,
    projects: projects,
  };
}

// ─────────────────────────────────────────────────────────────
// PAT ANALYTICS — Idris Dashboard
// ─────────────────────────────────────────────────────────────

// SLA window: 24h per workflow stage, but the clock FREEZES every weekend
// from Friday 23:59:00 through Monday 09:29:59 (resumes Monday 09:30:00).
var SLA_DURATION_MS = 24 * 60 * 60 * 1000;

// Returns the freeze window [start, end) that contains `t`, or null if not frozen.
function slaFreezeWindowFor(t) {
  var d = new Date(t);
  var day = d.getDay(); // 0=Sun .. 5=Fri .. 6=Sat
  var h = d.getHours(),
    m = d.getMinutes(),
    s = d.getSeconds();
  // Friday at/after 23:59:00 -> freeze until Monday 09:30:00
  if (day === 5 && (h > 23 || (h === 23 && m >= 59))) {
    var start = new Date(d);
    start.setHours(23, 59, 0, 0);
    var end = new Date(d);
    end.setDate(d.getDate() + 3);
    end.setHours(9, 30, 0, 0);
    return { start: start.getTime(), end: end.getTime() };
  }
  // Saturday or Sunday -> freeze until Monday 09:30:00
  if (day === 6 || day === 0) {
    var st = new Date(d);
    st.setHours(0, 0, 0, 0);
    var en = new Date(d);
    // roll to Monday
    var add = day === 6 ? 2 : 1;
    en.setDate(d.getDate() + add);
    en.setHours(9, 30, 0, 0);
    return { start: st.getTime(), end: en.getTime() };
  }
  return null;
}

// Total frozen milliseconds within [a, b) — sums every freeze window overlapping it.
// Scans forward from `a`, correctly handling the case where `a` itself is not frozen
// (e.g. Friday 22:00, before the 23:59 freeze start).
function slaFreezeInRange(a, b) {
  if (b <= a) return 0;
  if (!isFinite(a) || !isFinite(b)) return 0; // safety: never loop on bad input
  var total = 0;
  var guard = 0,
    MAX_FREEZE_ITERS = 366; // hard stop so a bad timestamp can't hang the run
  // Find the first freeze window at or after `a`. If `a` is already inside a freeze,
  // slaFreezeWindowFor returns it; otherwise jump to the next Friday 23:59.
  var win = slaFreezeWindowFor(a);
  if (!win) {
    var d = new Date(a);
    // days until next Friday 23:59 (0 if today is Friday but pre-23:59)
    var daysUntilFri = (5 - d.getDay() + 7) % 7;
    var nxt = new Date(d);
    nxt.setDate(d.getDate() + daysUntilFri);
    nxt.setHours(23, 59, 0, 0);
    win = slaFreezeWindowFor(nxt.getTime());
  }
  while (win && win.start < b) {
    var s = Math.max(win.start, a);
    var e = Math.min(win.end, b);
    if (e > s) total += e - s;
    if (win.end >= b) break;
    win = slaFreezeWindowFor(win.end + 1);
    if (++guard > MAX_FREEZE_ITERS) break; // safety valve
  }
  return total;
}

// Remaining SLA time at `now`. Counted (non-frozen) time = real elapsed minus any
// freeze window in [startTime, now]. The clock pauses during the weekend freeze, so
// breach occurs only after 24h of COUNTED time has elapsed. `startTime` is the
// project's SLA start (ms); if omitted we infer it as deadline - 24h.
function slaComputeRemaining(deadline, now, startTime) {
  if (startTime == null) startTime = deadline - SLA_DURATION_MS;
  var countedElapsed = now - startTime - slaFreezeInRange(startTime, now);
  if (countedElapsed < 0) countedElapsed = 0;
  return SLA_DURATION_MS - countedElapsed;
}

// Is the clock currently frozen right now?
function slaIsPaused(now) {
  return slaFreezeWindowFor(now) !== null;
}

/**
 * Returns PAT analytics for Idris Dashboard.
 * Schema: { total, byStatus, workflowEfficiency: { completionRate, averageCompletionTime, rejectionRate } }
 */

/**
 * SLA Analytics — real-time department SLA dashboard.
 * Analyzes all active PAT projects and computes per-department compliance.
 * SLA = 8 hours per workflow stage transition.
 */
function getSLAAnalytics(token) {
  try {
    if (token) _session(token);
    var now = new Date().getTime();

    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) {
      return {
        success: true,
        stats: {
          overallCompliance: 100,
          activeProjects: 0,
          breachedProjects: 0,
          completedLate: 0,
        },
        activeList: [],
        breachedList: [],
        departments: [],
        allDepts: [],
      };
    }

    var data = sh.getDataRange().getValues();
    var c = _patCols();

    var activeList = [];
    var breachedList = [];
    var deptMap = {};

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!r[c.PROJECT_ID]) continue;

      var status = String(r[c.WORKFLOW_STATUS] || "").trim() || "Draft";
      if (status === "Completed" || status === "Rejected") continue;

      var projectId = String(r[c.PROJECT_ID] || "").trim();
      var projectName = String(r[c.PROJECT_NAME] || "").trim();
      var updatedAt = String(r[c.UPDATED_AT] || r[c.SUBMITTED_AT] || "").trim();
      var startTime = updatedAt ? new Date(updatedAt).getTime() : now;
      if (isNaN(startTime)) startTime = now;

      var elapsed = now - startTime;
      var deadline = startTime + SLA_DURATION_MS;
      var remaining = slaComputeRemaining(deadline, now, startTime);
      var isPaused = slaIsPaused(now) && remaining > 0;
      var remainingForPct = Math.max(0, remaining);
      var elapsedPercent = Math.min(
        100,
        Math.max(
          0,
          Math.round(
            ((SLA_DURATION_MS - remainingForPct) / SLA_DURATION_MS) * 100,
          ),
        ),
      );

      var dept = String(
        r[c.ASSIGNED_TO_DEPT] ||
          r[c.DEPARTMENT] ||
          r[c.SUBMITTED_BY_DEPT] ||
          "Unassigned",
      ).trim();
      var assignedTo = String(
        r[c.ASSIGNED_TO_NAME] || r[c.ASSIGNED_TO_EMAIL] || "—",
      ).trim();

      var isBreached = remaining <= 0;

      var entry = {
        projectId: projectId,
        projectName: projectName || projectId,
        dept: dept,
        assignedTo: assignedTo,
        status: status,
        deadline: deadline,
        remaining: Math.max(0, remaining),
        paused: isPaused,
        overdueMinutes: isBreached ? Math.ceil(Math.abs(remaining) / 60000) : 0,
        elapsedPercent: elapsedPercent,
      };

      if (isBreached) breachedList.push(entry);
      else activeList.push(entry);

      if (!deptMap[dept])
        deptMap[dept] = { dept: dept, onTime: 0, breached: 0 };
      if (isBreached) deptMap[dept].breached++;
      else deptMap[dept].onTime++;
    }

    var departments = Object.keys(deptMap).map(function (d) {
      var dt = deptMap[d];
      var total = dt.onTime + dt.breached;
      dt.complianceRate =
        total > 0 ? Math.round((dt.onTime / total) * 100) : 100;
      return dt;
    });

    var totalActive = activeList.length;
    var totalBreached = breachedList.length;
    var totalProjects = totalActive + totalBreached;
    var overallCompliance =
      totalProjects > 0 ? Math.round((totalActive / totalProjects) * 100) : 100;

    return {
      success: true,
      stats: {
        overallCompliance: overallCompliance,
        activeProjects: totalActive,
        breachedProjects: totalBreached,
      },
      activeList: activeList,
      breachedList: breachedList,
      departments: departments,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Advanced SLA details for a single PAT project.
 * Returns full project data with SLA micro-tracking analysis.
 */
function getAdvancedSLADetails(token, projectId) {
  try {
    var projectRes = getPATProjectById(token, projectId);
    if (!projectRes.success) return projectRes;
    var p = projectRes.project;

    var SLA_DURATION_MS = 24 * 60 * 60 * 1000;
    var now = new Date().getTime();
    var updatedAt = p.updatedAt || p.submittedAt || "";
    var startTime = updatedAt ? new Date(updatedAt).getTime() : now;
    if (isNaN(startTime)) startTime = now;
    var elapsed = now - startTime;
    var deadline = startTime + SLA_DURATION_MS;
    var remaining = slaComputeRemaining(deadline, now, startTime);
    var isPaused = slaIsPaused(now) && remaining > 0;
    var remainingForPct = Math.max(0, remaining);
    var elapsedPercent = Math.min(
      100,
      Math.max(
        0,
        Math.round(
          ((SLA_DURATION_MS - remainingForPct) / SLA_DURATION_MS) * 100,
        ),
      ),
    );
    var isBreached = remaining <= 0;

    // Build timeline from workflow history
    var history = p.workflowHistory || [];
    var timeline = [];
    var prevTimestamp = null;
    history.forEach(function (h) {
      var ts = h.timestamp ? new Date(h.timestamp).getTime() : null;
      var durationFromPreviousHours = null;
      if (prevTimestamp && ts) {
        var diffMs = ts - prevTimestamp;
        durationFromPreviousHours = Math.round((diffMs / 3600000) * 10) / 10;
      }
      timeline.push({
        from: h.fromStatus || "—",
        to: h.toStatus || "—",
        by: h.by && h.by.name ? h.by.name : "System",
        department: h.by && h.by.department ? h.by.department : "—",
        durationFromPreviousHours: durationFromPreviousHours,
        timestamp: h.timestamp || null,
        comments: h.comments || "",
        isRejection: h.isRejection === true,
      });
      if (ts) prevTimestamp = ts;
    });

    // Build stage timing
    var stageTiming = {};
    history.forEach(function (h) {
      var stage = h.toStatus || h.fromStatus || "Unknown";
      if (!stageTiming[stage])
        stageTiming[stage] = { entries: 0, totalMs: 0, rejections: 0 };
      stageTiming[stage].entries++;
      if (h.isRejection) stageTiming[stage].rejections++;
    });

    // Count total workflow steps
    var totalSteps = timeline.length;

    // Calculate total journey time
    var totalJourneyMs = 0;
    if (timeline.length >= 2) {
      var firstTs = timeline[0].timestamp
        ? new Date(timeline[0].timestamp).getTime()
        : null;
      var lastTs = timeline[timeline.length - 1].timestamp
        ? new Date(timeline[timeline.length - 1].timestamp).getTime()
        : null;
      if (firstTs && lastTs) totalJourneyMs = lastTs - firstTs;
    }
    var totalJourneyHours =
      totalJourneyMs > 0 ? Math.round((totalJourneyMs / 3600000) * 10) / 10 : 0;

    // Count rejections
    var rejectionCount = timeline.filter(function (t) {
      return t.isRejection;
    }).length;

    // Avg time per step
    var avgTimePerStepHours =
      totalSteps > 0
        ? Math.round((totalJourneyMs / totalSteps / 3600000) * 10) / 10
        : 0;

    return {
      success: true,
      project: {
        projectId: p.projectId,
        projectName: p.projectName,
        siteAddress: p.siteAddress,
        vendor: p.vendor,
        orchestrator: p.orchestrator,
        workDescription: p.workDescription,
        inspectionDate: p.inspectionDate,
        snagScore: p.snagScore,
        verdict: p.verdict,
        workflowStatus: p.workflowStatus,
        presidingOfficer: p.presidingOfficer || "",
        assignedToDept: p.assignedToDept || "",
        assignedToName: p.assignedToName || "",
        assignedToEmail: p.assignedToEmail || "",
        hasVendor: p.hasVendor || false,
        vendorApprovalStatus: p.vendorApprovalStatus || "pending",
        department: p.department || "",
        updatedAt: p.updatedAt || "",
        submittedAt: p.submittedAt || "",
        submittedBy: p.submittedBy || "",
        submittedByEmail: p.submittedByEmail || "",
      },
      sla: {
        slaDurationHours: 24,
        elapsedHours: Math.round((elapsed / 3600000) * 10) / 10,
        remainingHours:
          Math.round((Math.max(0, remaining) / 3600000) * 10) / 10,
        remainingMinutes: Math.round(Math.max(0, remaining) / 60000),
        overdueMinutes: isBreached
          ? Math.round(Math.abs(remaining) / 60000)
          : 0,
        elapsedPercent: elapsedPercent,
        isBreached: isBreached,
        paused: isPaused,
        deadline: deadline,
        assignedTo: p.assignedToName || p.assignedToEmail || "—",
        dept: p.assignedToDept || p.department || "Unassigned",
        status: isBreached
          ? "breached"
          : elapsedPercent >= 75
            ? "at-risk"
            : "active",
      },
      metrics: {
        totalWorkflowSteps: totalSteps,
        totalJourneyHours: totalJourneyHours,
        rejectionCount: rejectionCount,
        avgTimePerStepHours: avgTimePerStepHours,
      },
      timeline: timeline,
      stageTiming: stageTiming,
      relatedMails: [],
      boqChanges: [],
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Performance scoring for Presiding Officers / Supervisors.
 * Aggregates every PAT by its PresidingOfficer and computes a 0-100 score:
 *   Quality (40%)    — inverse of avg snagScore (cleaner inspections score higher)
 *   SLA (35%)        — share of that officer's PATs that are NOT breached
 *   Pass rate (25%)  — Approved / (Approved + Rejected), Rejected penalised
 * Only officers with >= 1 PAT are returned; ranked best-first via a volume-aware
 * composite so light-touch officers don't top the board on a single clean PAT.
 */
function getPerformanceScores(token, force) {
  var CACHE_KEY = "perf_scores_v1";
  try {
    _session(token); // any authenticated user may view

    // ── PERF: serve cached leaderboard if fresh (avoids re-scanning the whole sheet) ──
    // `force` (passed by the UI Retry button) bypasses the cache for a live recompute.
    try {
      var cache = CacheService.getScriptCache();
      var cached = !force ? cache.get(CACHE_KEY) : null;
      if (cached) {
        var parsed = JSON.parse(cached);
        // re-stamp generatedAt so the UI always shows "updated just now"
        parsed.cached = true;
        parsed.generatedAt = new Date().toISOString();
        return parsed;
      }
    } catch (ce) {
      /* cache miss / corrupt — recompute below */
    }

    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) {
      return {
        success: true,
        officers: [],
        generatedAt: new Date().toISOString(),
      };
    }
    // PERF: read only the columns we need. The PAT sheet can hold large base64 image
    // blobs in the IMAGES column (idx 18); pulling the whole range every load is what
    // makes this slow enough to trip the 35s client timeout. Read around it instead.
    var data;
    try {
      var lastRow = sh.getLastRow();
      // Guard against degenerate sheets (header-only / zero data rows).
      if (!lastRow || lastRow < 2) {
        return {
          success: true,
          officers: [],
          generatedAt: new Date().toISOString(),
        };
      }
      // BULLETPROOF READ: never call getRange with an invalid column count, and
      // never let a read hiccup kill the whole leaderboard. We read cols we need
      // while skipping the IMAGES blob column (idx 18) to avoid the 35s timeout.
      var maxCols =
        typeof sh.getMaxColumns === "function" ? sh.getMaxColumns() : 37;
      var endCol = Math.min(37, maxCols); // PRESIDING_OFFICER is idx 36
      if (endCol <= 18) {
        // Sheet has no PRESIDING_OFFICER column yet — nothing to score, but don't error.
        data = sh.getRange(1, 1, lastRow, Math.max(1, endCol)).getValues();
      } else {
        var partA = sh.getRange(1, 1, lastRow, 18).getValues(); // cols A–R (idx 0–17)
        var partBCount = endCol - 19;
        if (partBCount < 1) partBCount = 1; // guard: getRange requires >= 1 column
        var partB = sh.getRange(1, 20, lastRow, partBCount).getValues(); // cols T– (idx 19+)
        data = partA.map(function (rowA, i) {
          var r = rowA.slice();
          r.splice(18, 0, null); // restore the skipped IMAGES slot (unused by scoring)
          var b = partB[i] || [];
          for (var k = 0; k < b.length; k++) r.push(b[k]);
          return r;
        });
      }
    } catch (readErr) {
      // Fallback to a full read if the scoped read isn't possible on this sheet.
      try {
        data = sh.getDataRange().getValues();
      } catch (e2) {
        // Last-ditch: return empty rather than throwing an opaque "Could not load scores".
        return {
          success: true,
          officers: [],
          generatedAt: new Date().toISOString(),
        };
      }
    }
    var c = _patCols();
    var now = new Date().getTime();

    // GROUP PATs BY PRESIDING OFFICER
    var map = {};
    for (var i = 1; i < data.length; i++) {
      // ── ROBUST: one malformed row must never break the entire leaderboard ──
      try {
        var r = data[i];
        if (!r || !r[c.PROJECT_ID]) continue;
        var officer = String(r[c.PRESIDING_OFFICER] || "").trim();
        if (!officer) continue; // only scored PATs count
        if (!map[officer]) {
          map[officer] = {
            officer: officer,
            department: String(
              r[c.SUBMITTED_BY_DEPT] ||
                r[c.ASSIGNED_TO_DEPT] ||
                r[c.DEPARTMENT] ||
                "Unassigned",
            ).trim(),
            patCount: 0,
            snagSum: 0,
            notBreached: 0,
            totalForSla: 0,
            approved: 0,
            rejected: 0,
            other: 0,
            pats: [],
            timestamps: [], // per-PAT timestamps (ms) for time-windowed averages
            windowScores: [], // per-PAT composite score, aligned with timestamps
          };
        }
        var g = map[officer];
        g.patCount++;
        var verdict = String(r[c.VERDICT] || "")
          .trim()
          .toLowerCase();
        if (verdict === "approved") g.approved++;
        else if (verdict === "rejected") g.rejected++;
        else g.other++;
        var rowSnag = Number(r[c.SNAG_SCORE]) || 0;
        g.snagSum += rowSnag;

        // Freeze-aware breach check
        var wf = String(r[c.WORKFLOW_STATUS] || "").trim();
        var status = String(r[c.STATUS] || "").trim();
        g.totalForSla++;
        var breached = false;
        if (wf === "Completed" || wf === "Rejected" || status === "Rejected") {
          // closed PAT: resolved within SLA if approved; rejected PATs still resolve
          g.notBreached++;
        } else {
          var upd = String(r[c.UPDATED_AT] || r[c.SUBMITTED_AT] || "").trim();
          var st = upd ? new Date(upd).getTime() : NaN;
          // Guard: a blank/corrupt date must NOT become `now`. Treat unknowns as breached.
          if (isNaN(st)) {
            g.notBreached++; // no usable timestamp => counts as breached, but does not kill the loop
            breached = true;
          } else {
            var dl = st + SLA_DURATION_MS;
            var rm = slaComputeRemaining(dl, now, st);
            if (rm > 0) g.notBreached++;
            else breached = true;
          }
        }
        // Per-PAT composite (same formula as the officer aggregate) — used for window averages.
        var q2 = Math.max(0, Math.min(100, Math.round(100 - rowSnag * 20)));
        var sla2 = breached ? 0 : 100;
        var dec2 = g.approved + g.rejected; // not exact per-row but adequate for window trend
        var pass2 =
          verdict === "approved" ? 100 : verdict === "rejected" ? 0 : 100;
        var rowScore = Math.round(q2 * 0.4 + sla2 * 0.35 + pass2 * 0.25);
        // Capture timestamp (prefer updatedAt, fall back to submittedAt) for window bucketing.
        var tsRaw = String(r[c.UPDATED_AT] || r[c.SUBMITTED_AT] || "").trim();
        var ts = tsRaw ? new Date(tsRaw).getTime() : NaN;
        if (!isNaN(ts)) {
          g.timestamps.push(ts);
          g.windowScores.push(rowScore);
        }
        g.pats.push({
          projectId: String(r[c.PROJECT_ID]),
          projectName: String(r[c.PROJECT_NAME] || ""),
          snagScore: rowSnag,
          verdict: r[c.VERDICT] || "Pending",
          workflowStatus: wf,
        });
      } catch (rowErr) {
        // skip the bad row, keep going
        continue;
      }
    }

    // ── SEED WITH ALL USERS so the board lists every person (all roles/depts),
    // not just those who already have a PAT. Inactive users sort to the bottom
    // and show "No score yet" instead of a fabricated score.
    try {
      var _usheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SD.USERS,
      );
      if (_usheet && _usheet.getLastRow() > 1) {
        var _udata = _usheet.getDataRange().getValues();
        for (var _ui = 1; _ui < _udata.length; _ui++) {
          var _uname = String(_udata[_ui][1] || "").trim();
          var _uem = String(_udata[_ui][2] || "").trim();
          var _key = _uname || _uem;
          if (!_key) continue;
          if (!map[_key]) {
            map[_key] = {
              officer: _key,
              department: String(_udata[_ui][6] || "").trim(),
              patCount: 0,
              snagSum: 0,
              notBreached: 0,
              totalForSla: 0,
              approved: 0,
              rejected: 0,
              other: 0,
              pats: [],
              timestamps: [],
              windowScores: [],
            };
          }
        }
      }
    } catch (_seedErr) {
      /* best-effort: board still works with PAT-derived officers */
    }

    var officers = Object.keys(map).map(function (name) {
      var g = map[name];
      // Inactive (no PAT) users: no score, no bars — render as "no data".
      if (g.patCount === 0) {
        return {
          officer: g.officer,
          department: g.department,
          patCount: 0,
          avgSnagScore: 0,
          qualityScore: null,
          slaCompliance: null,
          passRate: null,
          approved: 0,
          rejected: 0,
          score: null,
          avgDay: null,
          avgWeek: null,
          avgMonth: null,
          pats: [],
          noData: true,
        };
      }
      var avgSnag = g.patCount ? g.snagSum / g.patCount : 0;
      // Quality: 0 snag -> 100, 5+ snag -> 0 (linear)
      var quality = Math.max(0, Math.min(100, Math.round(100 - avgSnag * 20)));
      var slaPct = g.totalForSla
        ? Math.round((g.notBreached / g.totalForSla) * 100)
        : 100;
      var decided = g.approved + g.rejected;
      var passRate = decided ? Math.round((g.approved / decided) * 100) : 100;
      // Composite (0-100), volume-aware so a single clean PAT isn't #1
      var volumeConf = Math.min(1, g.patCount / 5); // confidence ramps to full at 5 PATs
      var raw = quality * 0.4 + slaPct * 0.35 + passRate * 0.25;
      var score = Math.round(raw * (0.7 + 0.3 * volumeConf)); // blend with confidence

      // ── DAY / WEEK / MONTH AVERAGES ──
      // Average composite score over PATs whose timestamp falls in the last
      // 24h / 7d / 30d. Uses the per-PAT composite (windowScores) aligned to timestamps.
      var DAY_MS = 24 * 60 * 60 * 1000;
      function _avgInWindow(ms) {
        var sum = 0,
          n = 0;
        for (var i = 0; i < g.timestamps.length; i++) {
          if (now - g.timestamps[i] <= ms) {
            sum += g.windowScores[i];
            n++;
          }
        }
        return n ? Math.round(sum / n) : null; // null = no PATs in window
      }
      var avgDay = _avgInWindow(DAY_MS);
      var avgWeek = _avgInWindow(7 * DAY_MS);
      var avgMonth = _avgInWindow(30 * DAY_MS);

      return {
        officer: g.officer,
        department: g.department,
        patCount: g.patCount,
        avgSnagScore: Math.round(avgSnag * 10) / 10,
        qualityScore: quality,
        slaCompliance: slaPct,
        passRate: passRate,
        approved: g.approved,
        rejected: g.rejected,
        score: score,
        avgDay: avgDay,
        avgWeek: avgWeek,
        avgMonth: avgMonth,
        pats: g.pats,
      };
    });

    officers.sort(function (a, b) {
      // Active (scored) officers first; no-data users sink to the bottom.
      var aNull = a.score === null,
        bNull = b.score === null;
      if (aNull !== bNull) return aNull ? 1 : -1;
      if (!aNull && !bNull && b.score !== a.score) return b.score - a.score;
      if (!aNull && !bNull) return b.patCount - a.patCount;
      return String(a.officer).localeCompare(String(b.officer));
    });
    officers.forEach(function (o, idx) {
      o.rank = idx + 1;
    });

    var result = {
      success: true,
      officers: officers,
      generatedAt: new Date().toISOString(),
    };

    // ── PERF: cache for 2 minutes so repeat loads / tab switches are instant ──
    try {
      CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(result), 120);
    } catch (ce) {}

    return result;
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Invalidate the cached leaderboard (call after any PAT/assignment change).
function _invalidatePerfCache() {
  try {
    CacheService.getScriptCache().remove("perf_scores_v1");
  } catch (e) {}
}

/**
 * Assign (or change) the Presiding Officer on a PAT. Stores the officer name and
 * resolves their department from the USERS sheet so performance scoring can group
 * by both person and department.
 */
function assignPresidingOfficer(token, projectId, officerName) {
  try {
    var s = _session(token);
    // Only MEC department or Super Admin may assign a Presiding Officer.
    var sRole = String(s.role || "").toLowerCase();
    var sDept = String(s.department || "").toLowerCase();
    var isMEC = sDept === "mec" || sDept === "mech";
    var isSuperAdmin = sRole === "super admin";
    if (!isMEC && !isSuperAdmin) {
      throw new Error(
        "Only MEC personnel or a Super Admin may assign a Presiding Officer.",
      );
    }
    if (!projectId || !officerName)
      throw new Error("Project ID and officer name are required.");
    var sh = _patSheet();
    var data = sh.getDataRange().getValues();
    var c = _patCols();
    var norm = String(projectId).trim().toUpperCase();
    var target = -1;
    for (var i = 1; i < data.length; i++) {
      var rid = String(data[i][c.PROJECT_ID] || "")
        .trim()
        .toUpperCase();
      if (rid === norm || "PAT-" + rid === norm || "FOB-" + rid === norm) {
        target = i;
        break;
      }
    }
    if (target < 0) return { success: false, message: "PAT not found." };

    // Resolve department from users
    var dept = "";
    try {
      var ush = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
      var ud = ush.getDataRange().getValues();
      for (var j = 1; j < ud.length; j++) {
        var uname = String(ud[j][1] || "")
          .trim()
          .toLowerCase();
        var uemail = String(ud[j][2] || "")
          .trim()
          .toLowerCase();
        if (
          uname === String(officerName).trim().toLowerCase() ||
          uemail === String(officerName).trim().toLowerCase()
        ) {
          dept = String(ud[j][6] || "");
          break;
        }
      }
    } catch (e2) {}

    sh.getRange(target + 1, c.PRESIDING_OFFICER + 1).setValue(
      String(officerName).trim(),
    );
    _invalidatePerfCache();
    return {
      success: true,
      message: "Presiding Officer updated.",
      officer: String(officerName).trim(),
      department: dept,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// PAT CRUD — full backend for PAT projects
// ─────────────────────────────────────────────────────────────

function _patSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.PAT);
}

function _patCols() {
  return {
    PROJECT_ID: 0,
    PROJECT_NAME: 1,
    SITE_ADDRESS: 2,
    LAT: 3,
    LON: 4,
    PHASE: 5,
    WORK_DESC: 6,
    VENDOR: 7,
    INSP_DATE: 8,
    ORCHESTRATOR: 9,
    SNAG_SCORE: 10,
    STATUS: 11,
    VERDICT: 12,
    WORKFLOW_STATUS: 13,
    CHECKLIST: 14,
    BOQ: 15,
    SNAGS: 16,
    SIGNOFF: 17,
    IMAGES: 18,
    HAS_VENDOR: 19,
    VENDOR_TOKEN: 20,
    VENDOR_APPROVAL_STATUS: 21,
    VENDOR_APPROVAL_COMMENTS: 22,
    VENDOR_APPROVAL_DATE: 23,
    SUBMITTED_BY: 24,
    SUBMITTED_BY_EMAIL: 25,
    SUBMITTED_BY_DEPT: 26,
    SUBMITTED_AT: 27,
    ASSIGNED_TO_NAME: 28,
    ASSIGNED_TO_EMAIL: 29,
    ASSIGNED_TO_DEPT: 30,
    REJECTION_REASON: 31,
    WORKFLOW_HISTORY: 32,
    DEPARTMENT: 33,
    UPDATED_AT: 34,
    VENDOR_EVER_APPROVED: 35,
    PRESIDING_OFFICER: 36,
    EDIT_LOG: 37,
  };
}

/**
 * Ensures the SD_PAT_PROJECTS sheet has the EDIT_LOG column (col 37).
 * Safe to call repeatedly — only appends a header cell if missing.
 * Older rows simply have an empty EDIT_LOG and are treated as [].
 */
function _ensurePATEditLogColumn() {
  try {
    var sh = _patSheet();
    if (!sh) return;
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (headers.length <= 37) {
      // Pad header row out to 38 columns (index 37 = EDIT_LOG)
      for (var i = headers.length; i < 38; i++) headers.push("");
      headers[37] = "EditLog";
      sh.getRange(1, 1, 1, 38).setValues([headers]);
    }
  } catch (e) {
    console.warn("_ensurePATEditLogColumn: " + e.message);
  }
}

/**
 * Builds a single audit entry describing what changed in a PAT save.
 * Called for EVERY editor (MEC, every department, Super Admin) so no
 * edit is ever silent. Returns an entry object or null if nothing changed.
 * @param {Object} sess - session { name, email, department }
 * @param {Object} prev - the project as it was in the DB before this save
 * @param {Object} next - the data being saved
 */
function _buildEditLogEntry(sess, prev, next) {
  sess = sess || {};
  prev = prev || {};
  next = next || {};
  var changed = [];
  var TRACKED = [
    "projectName",
    "siteAddress",
    "lat",
    "lon",
    "phase",
    "workDescription",
    "vendor",
    "inspectionDate",
    "orchestrator",
    "presidingOfficer",
    "verdict",
    "workflowStatus",
    "assignedToDept",
    "rejectionReason",
    "department",
  ];
  TRACKED.forEach(function (f) {
    var a = prev[f] == null ? "" : String(prev[f]);
    var b = next[f] == null ? "" : String(next[f]);
    if (a !== b) changed.push(f);
  });

  // Collection fields (arrays/objects) — detect add/remove by length delta
  ["checklist", "boq", "snags", "images"].forEach(function (f) {
    var pa = (prev[f] && prev[f].length) || 0;
    var pb = (next[f] && next[f].length) || 0;
    if (pa !== pb) changed.push(f + " (" + pa + "\u2192" + pb + ")");
  });

  var statusBefore = String(prev.workflowStatus || "Draft");
  var statusAfter = String(
    next.workflowStatus || prev.workflowStatus || "Draft",
  );

  return {
    at: new Date().toISOString(),
    by: sess.name || sess.email || "Unknown",
    email: sess.email || "",
    dept: sess.department || prev.assignedToDept || "",
    role: sess.role || "",
    action: "edit",
    changed: changed,
    statusBefore: statusBefore,
    statusAfter: statusAfter,
  };
}

/**
 * Appends an edit-log entry to a PAT row (cumulative — history is never lost).
 * @param {number} rowIndex - 0-based data row index (i.e. sheet row = rowIndex+1)
 */
function _appendEditLog(rowIndex, entry) {
  if (!entry) return;
  _ensurePATEditLogColumn();
  var c = _patCols();
  var sh = _patSheet();
  if (!sh) return;
  var cur = sh.getRange(rowIndex + 1, c.EDIT_LOG + 1).getValue();
  var log = [];
  try {
    log = cur ? JSON.parse(cur) : [];
  } catch (e) {
    log = [];
  }
  if (!Array.isArray(log)) log = [];
  log.push(entry);
  // Cap the log so the cell never grows unbounded (keep last 200 entries)
  if (log.length > 200) log = log.slice(log.length - 200);
  sh.getRange(rowIndex + 1, c.EDIT_LOG + 1).setValue(JSON.stringify(log));
}

function _projectFromRow(r) {
  var c = _patCols();
  return {
    projectId: String(r[c.PROJECT_ID] || "").trim(),
    projectName: r[c.PROJECT_NAME] || "",
    siteAddress: r[c.SITE_ADDRESS] || "",
    lat: r[c.LAT] || "",
    lon: r[c.LON] || "",
    phase: r[c.PHASE] || "",
    workDescription: r[c.WORK_DESC] || "",
    vendor: r[c.VENDOR] || "",
    inspectionDate: r[c.INSP_DATE] || "",
    orchestrator: r[c.ORCHESTRATOR] || "",
    snagScore: Number(r[c.SNAG_SCORE]) || 0,
    status: r[c.STATUS] || "",
    verdict: r[c.VERDICT] || "Pending",
    workflowStatus: r[c.WORKFLOW_STATUS] || "Draft",
    checklist: _safeParse(r[c.CHECKLIST], {}),
    boq: _safeParse(r[c.BOQ], []),
    snags: _safeParse(r[c.SNAGS], []),
    signoff: _safeParse(r[c.SIGNOFF], {}),
    images: _safeParse(r[c.IMAGES], []),
    hasVendor: String(r[c.HAS_VENDOR] || "").toUpperCase() === "TRUE",
    vendorToken: r[c.VENDOR_TOKEN] || "",
    vendorApprovalStatus: r[c.VENDOR_APPROVAL_STATUS] || "pending",
    vendorApprovalComments: r[c.VENDOR_APPROVAL_COMMENTS] || "",
    vendorApprovalDate: r[c.VENDOR_APPROVAL_DATE] || "",
    vendorEverApproved:
      String(r[c.VENDOR_EVER_APPROVED] || "").toUpperCase() === "TRUE",
    submittedBy: r[c.SUBMITTED_BY] || "",
    submittedByEmail: r[c.SUBMITTED_BY_EMAIL] || "",
    submittedByDept: r[c.SUBMITTED_BY_DEPT] || "",
    submittedAt: r[c.SUBMITTED_AT] || "",
    assignedToName: r[c.ASSIGNED_TO_NAME] || "",
    assignedToEmail: r[c.ASSIGNED_TO_EMAIL] || "",
    assignedToDept: r[c.ASSIGNED_TO_DEPT] || r[c.SUBMITTED_BY_DEPT] || "",
    rejectionReason: r[c.REJECTION_REASON] || "",
    workflowHistory: _safeParse(r[c.WORKFLOW_HISTORY], []),
    department: r[c.DEPARTMENT] || "",
    updatedAt: r[c.UPDATED_AT] || "",
    presidingOfficer: r[c.PRESIDING_OFFICER] || "",
    editLog: _safeParse(r[c.EDIT_LOG], []),
  };
}

function _safeParse(str, fallback) {
  if (!str) return fallback;
  // If already an object/array, return it
  if (typeof str === "object") return str;
  // If it's a string, try to parse it
  if (typeof str === "string") {
    try {
      return JSON.parse(str);
    } catch (e) {
      return fallback;
    }
  }
  return fallback;
}

/**
 * Ensure a value is properly formatted for saving:
 * - If it's already an object/array, stringify it
 * - If it's a string, check if it's JSON and return as-is, else return fallback
 */
function _ensureJSONString(val, fallback) {
  if (!val) return JSON.stringify(fallback);
  if (typeof val === "object") return JSON.stringify(val);
  if (typeof val === "string") {
    try {
      JSON.parse(val);
      return val;
    } catch (e) {
      return JSON.stringify(fallback);
    }
  }
  return JSON.stringify(fallback);
}

// ─────────────────────────────────────────────────────────────
// DOCUMENT HELPERS (SD_DOCUMENTS)
// ─────────────────────────────────────────────────────────────

function _docSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DOCS);
}

function _docCols() {
  return {
    DOC_ID: 0,
    PROJECT_ID: 1,
    FILE_NAME: 2,
    FILE_TYPE: 3,
    FILE_SIZE: 4,
    DRIVE_URL: 5,
    DRIVE_FILE_ID: 6,
    UPLOADED_BY: 7,
    UPLOADED_BY_EMAIL: 8,
    UPLOADED_BY_DEPT: 9,
    UPLOADED_AT: 10,
    CATEGORY: 11,
  };
}

function _documentFromRow(r) {
  var c = _docCols();
  return {
    docId: r[c.DOC_ID] || "",
    projectId: String(r[c.PROJECT_ID] || "").trim(),
    fileName: r[c.FILE_NAME] || "",
    fileType: r[c.FILE_TYPE] || "",
    fileSize: r[c.FILE_SIZE] || 0,
    driveUrl: r[c.DRIVE_URL] || "",
    driveFileId: r[c.DRIVE_FILE_ID] || "",
    uploadedBy: r[c.UPLOADED_BY] || "",
    uploadedByEmail: r[c.UPLOADED_BY_EMAIL] || "",
    uploadedByDept: r[c.UPLOADED_BY_DEPT] || "",
    uploadedAt: r[c.UPLOADED_AT] || "",
    category: r[c.CATEGORY] || "",
  };
}

function _genDocId() {
  return "DOC-" + Utilities.getUuid().substring(0, 6).toUpperCase();
}

function _genPATId() {
  return "PAT-" + Utilities.getUuid().substring(0, 6).toUpperCase();
}

/**
 * List all PAT projects.
 */
function getPATProjects(token) {
  try {
    var sess = token ? _session(token) : null;
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: true, projects: [] };
    var data = sh
      .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .getValues();
    var allProjects = data
      .filter(function (r) {
        return r[0];
      })
      .map(_projectFromRow);

    // If no session (internal call) or user is Admin/MEC, return everything
    if (!sess) return { success: true, projects: allProjects };

    var role = String(sess.role || "").toLowerCase();
    var dept = String(sess.department || "").toLowerCase();
    var email = String(sess.email || "").toLowerCase();
    var isSuperAdmin = role === "super admin";
    var isMEC = dept === "mec" || dept === "mech";

    // All authenticated users can see all PAT projects (except Drafts for non-MEC)
    // This enables cross-department visibility — any department can check details
    // The frontend controls what actions each user can perform based on their department
    var visibleProjects = allProjects.filter(function (p) {
      var status = p.workflowStatus || "Draft";
      // Only hide Drafts from non-MEC users (MEC workspace)
      if (status === "Draft" && !isMEC && !isSuperAdmin) return false;
      return true;
    });

    return { success: true, projects: visibleProjects };
  } catch (e) {
    return { success: false, message: e.message, projects: [] };
  }
}

/**
 * Get a single PAT project by ID.
 */
function getPATProjectById(token, projectId) {
  try {
    // SECURITY: Token is mandatory for all external callers.
    // Internal callers pass null explicitly as first arg: getPATProjectById(null, projectId)
    if (arguments.length < 2) {
      throw new Error(
        "Authentication required. Call with (token, projectId) or (null, projectId) for internal use.",
      );
    }
    var isInternalCall = token === null;
    if (!isInternalCall) {
      _session(token); // Validates token and throws if invalid
    }
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };
    var data = sh
      .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .getValues();
    var searchId = String(projectId).trim();
    var altId = searchId.replace(/^(FOB|PAT)-/i, "");
    var foundRow = -1;
    var foundProject = null;
    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      if (
        rowId.toUpperCase() === searchId.toUpperCase() ||
        rowId.toUpperCase() === altId.toUpperCase() ||
        ("PAT-" + rowId).toUpperCase() === searchId.toUpperCase()
      ) {
        foundRow = i;
        foundProject = _projectFromRow(data[i]);
        break;
      }
    }
    if (foundRow === -1)
      return { success: false, message: "Project not found: " + projectId };

    // All authenticated users can view any PAT project details.
    // Action permissions are checked separately via _userCanActOnProject/checkPATPermission.

    // ── LEGACY METRO-PAT REPAIR (idempotent, one-time) ──
    // PATs created under the OLD schema went Draft -> "Awaiting Service Delivery"
    // in a single hop with no entryRoute and no Planning step. The new metro-route
    // ENTRY stage is the distinct status "Awaiting Service Delivery (Metro)", after
    // which the metro route now flows to Planning then Final MEC Review. Without
    // repair, such PATs are misread as project-route (and would skip the metro entry).
    // Detect that legacy shape and relabel + backfill the metro route marker.
    // We only relabel the entry stage here; the onward Planning step is added by the
    // normal route-aware flow when the project is advanced, so no Planning hop is backfilled.
    var hist = foundProject.workflowHistory || [];
    var hasRoute = hist.some(function (h) {
      return h && h.entryRoute;
    });
    var isLegacyMetro =
      foundProject.workflowStatus === "Awaiting Service Delivery" && !hasRoute;
    if (isLegacyMetro) {
      foundProject.workflowStatus = "Awaiting Service Delivery (Metro)";
      // Mark the original Draft hop with entryRoute so downstream routing is metro.
      if (hist[0]) hist[0].entryRoute = "metro";
      foundProject.workflowHistory = hist;
      // Persist the repair so it only happens once.
      try {
        savePATProject(null, foundProject);
      } catch (e) {
        /* read-only-safe */
      }
    }

    return { success: true, project: foundProject };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Save (create or update) a PAT project.
 */
function savePATProject(token, data) {
  try {
    var sess = _session(token);
    var c = _patCols();
    var sh = _patSheet();
    if (!sh) return { success: false, message: "PAT sheet not found." };

    var existingId = data.projectId || "";
    var now = new Date().toISOString();

    // Build row data
    function _rowFromProject(p) {
      var row = [];
      row[c.PROJECT_ID] = p.projectId || _genPATId();
      row[c.PROJECT_NAME] = String(p.projectName || "");
      row[c.SITE_ADDRESS] = String(p.siteAddress || "");
      row[c.LAT] = String(p.lat || "");
      row[c.LON] = String(p.lon || "");
      row[c.PHASE] = String(p.phase || "");
      row[c.WORK_DESC] = String(p.workDescription || "");
      row[c.VENDOR] = String(p.vendor || "");
      row[c.INSP_DATE] = String(p.inspectionDate || "");
      row[c.ORCHESTRATOR] = String(p.orchestrator || "");
      row[c.SNAG_SCORE] = Number(p.snagScore) || 0;
      row[c.STATUS] = String(p.status || "");
      row[c.VERDICT] = String(p.verdict || "Pending");
      row[c.WORKFLOW_STATUS] = String(p.workflowStatus || "Draft");
      // Safely handle JSON fields that might already be stringified
      row[c.CHECKLIST] = _ensureJSONString(p.checklist, {});
      row[c.BOQ] = _ensureJSONString(p.boq, []);
      row[c.SNAGS] = _ensureJSONString(p.snags, []);
      row[c.SIGNOFF] = _ensureJSONString(p.signoff, {});
      row[c.IMAGES] = _ensureJSONString(p.images, []);
      row[c.HAS_VENDOR] = p.hasVendor ? "TRUE" : "FALSE";
      row[c.VENDOR_TOKEN] = String(p.vendorToken || "");
      row[c.VENDOR_APPROVAL_STATUS] = String(
        p.vendorApprovalStatus || "pending",
      );
      row[c.VENDOR_APPROVAL_COMMENTS] = String(p.vendorApprovalComments || "");
      row[c.VENDOR_APPROVAL_DATE] = String(p.vendorApprovalDate || "");
      row[c.VENDOR_EVER_APPROVED] = p.vendorEverApproved ? "TRUE" : "FALSE";
      row[c.SUBMITTED_BY] = String(p.submittedBy || (sess ? sess.name : ""));
      row[c.SUBMITTED_BY_EMAIL] = String(
        p.submittedByEmail || (sess ? sess.email : ""),
      );
      row[c.SUBMITTED_BY_DEPT] = String(
        p.submittedByDept || (sess ? sess.department : ""),
      );
      row[c.SUBMITTED_AT] = p.submittedAt || now;
      // Ensure assignedToDept is set to the user's department if not provided
      row[c.ASSIGNED_TO_NAME] = String(p.assignedToName || "");
      row[c.ASSIGNED_TO_EMAIL] = String(p.assignedToEmail || "");
      row[c.ASSIGNED_TO_DEPT] = String(
        p.assignedToDept || (sess ? sess.department : "Unassigned"),
      );
      row[c.REJECTION_REASON] = String(p.rejectionReason || "");
      row[c.WORKFLOW_HISTORY] = _ensureJSONString(p.workflowHistory, []);
      row[c.DEPARTMENT] = String(p.department || "");
      row[c.PRESIDING_OFFICER] = String(p.presidingOfficer || "");
      row[c.UPDATED_AT] = now;
      // Preserve the cumulative edit log (never overwritten by the form)
      row[c.EDIT_LOG] = _ensureJSONString(p.editLog, []);
      return row;
    }

    // Check if updating existing
    if (existingId) {
      var normExisting = existingId.trim();
      var altExisting = normExisting.replace(/^(FOB|PAT)-/i, "");
      var existingData = sh.getDataRange().getValues();
      for (var i = 1; i < existingData.length; i++) {
        var rowId = String(existingData[i][0]).trim();
        if (
          rowId.toUpperCase() === normExisting.toUpperCase() ||
          rowId.toUpperCase() === altExisting.toUpperCase() ||
          ("PAT-" + rowId).toUpperCase() === normExisting.toUpperCase()
        ) {
          var existingProject = _projectFromRow(existingData[i]);
          var sessRole = String(sess.role || "").toLowerCase();
          var sessDept = String(sess.department || "").toLowerCase();
          var isAdmin = sessRole === "admin" || sessRole === "super admin";
          var isMEC = sessDept === "mec" || sessDept === "mech";

          // STRICT PERMISSION CHECK FOR NON-MEC/ADMIN
          if (!isMEC && !isAdmin) {
            // Non-MEC/Admin users can only update Snags and workflow metadata.
            // We ignore any changes they might have attempted on Site Info, Checklist, or BOQ
            // by using the existing project data as the master template for those columns.
            var finalRow = _rowFromProject(existingProject); // Start with DB version
            finalRow[c.SNAGS] = _ensureJSONString(data.snags, []);
            finalRow[c.SNAG_SCORE] =
              Number(data.snagScore || data.complaintScore) || 0;

            // ALLOW metadata updates for the current assignee/dept
            finalRow[c.WORKFLOW_STATUS] =
              data.workflowStatus !== undefined
                ? String(data.workflowStatus)
                : existingProject.workflowStatus;
            finalRow[c.WORKFLOW_HISTORY] = _ensureJSONString(
              data.workflowHistory,
              [],
            );
            finalRow[c.DEPARTMENT] =
              data.department !== undefined
                ? String(data.department)
                : existingProject.department;
            finalRow[c.ASSIGNED_TO_DEPT] =
              data.assignedToDept !== undefined
                ? String(data.assignedToDept)
                : existingProject.assignedToDept;
            finalRow[c.ASSIGNED_TO_NAME] =
              data.assignedToName !== undefined
                ? String(data.assignedToName)
                : existingProject.assignedToName;
            finalRow[c.ASSIGNED_TO_EMAIL] =
              data.assignedToEmail !== undefined
                ? String(data.assignedToEmail)
                : existingProject.assignedToEmail;
            finalRow[c.REJECTION_REASON] =
              data.rejectionReason !== undefined
                ? String(data.rejectionReason)
                : existingProject.rejectionReason;
            finalRow[c.VERDICT] =
              data.verdict !== undefined
                ? String(data.verdict)
                : existingProject.verdict;

            // SNAG PERMISSION ENFORCEMENT:
            // Prevent unauthorized deletion/modification of snags added by other departments.
            var incomingSnagsFromUI = _safeParse(data.snags, []); // Snags from frontend
            var originalSnagsFromDB = _safeParse(existingProject.snags, []); // Snags from DB
            var finalSnagsForSave = [];

            // Robust Snag Merging: Keep original snags that the current user is NOT allowed to touch.
            originalSnagsFromDB.forEach(function (origS) {
              if (!_canEditSnag(sess, origS, existingProject)) {
                finalSnagsForSave.push(origS);
              }
            });
            // Add all incoming snags that the user HAS permission to manage (edits or new additions).
            incomingSnagsFromUI.forEach(function (incS) {
              if (_canEditSnag(sess, incS, existingProject)) {
                finalSnagsForSave.push(incS);
              }
            });
            finalRow[c.SNAGS] = _ensureJSONString(finalSnagsForSave, []);
            finalRow[c.SNAG_SCORE] = finalSnagsForSave.reduce(function (
              sum,
              s,
            ) {
              return sum + (parseInt(s.weight) || 0);
            }, 0);

            finalRow[c.UPDATED_AT] = now;
            sh.getRange(i + 1, 1, 1, finalRow.length).setValues([finalRow]);
            SpreadsheetApp.flush();
            // Audit: every edit is logged, cumulative
            try {
              var _entryNM = _buildEditLogEntry(sess, existingProject, {
                workflowStatus: finalRow[c.WORKFLOW_STATUS],
                assignedToDept: finalRow[c.ASSIGNED_TO_DEPT],
                snags: _safeParse(finalRow[c.SNAGS], []),
                images: _safeParse(finalRow[c.IMAGES], []),
              });
              _appendEditLog(i, _entryNM);
            } catch (e) {
              console.warn("editlog(nonMEC): " + e.message);
            }
            return {
              success: true,
              projectId: existingId,
              message: "Workflow updated successfully.",
            };
          }

          // Full update allowed ONLY for MEC and Super Admins
          if (isMEC || isAdmin) {
            var rowData = _rowFromProject(data);
            // Preserve original submission data
            var origData = existingData[i];
            rowData[c.SUBMITTED_BY] =
              origData[c.SUBMITTED_BY] || data.submittedBy || "";
            rowData[c.SUBMITTED_BY_EMAIL] =
              origData[c.SUBMITTED_BY_EMAIL] || data.submittedByEmail || "";
            rowData[c.SUBMITTED_BY_DEPT] =
              data.submittedByDept ||
              existingData[i][c.SUBMITTED_BY_DEPT] ||
              "";
            rowData[c.SUBMITTED_AT] =
              data.submittedAt || existingData[i][c.SUBMITTED_AT] || now;
            rowData[c.WORKFLOW_HISTORY] = data.workflowHistory
              ? JSON.stringify(data.workflowHistory)
              : existingData[i][c.WORKFLOW_HISTORY] || "[]";

            // ── WIPE GUARD ──
            // Never let an empty form value silently overwrite a non-empty DB
            // column. If the form didn't supply BOQ/IMAGES/CHECKLIST/SNAGS/
            // SIGNOFF (e.g. those sections weren't rendered or hadn't loaded
            // yet), keep what's already persisted. This is what previously
            // caused PATs to "clear for no reason" on edit.
            var _dbBoq = _safeParse(existingData[i][c.BOQ], []);
            var _dbImages = _safeParse(existingData[i][c.IMAGES], []);
            var _dbCheck = _safeParse(existingData[i][c.CHECKLIST], {});
            var _dbSnags = _safeParse(existingData[i][c.SNAGS], []);
            var _dbSignoff = _safeParse(existingData[i][c.SIGNOFF], {});

            // BOQ: keep DB unless the form actually sent a non-empty array.
            if (data.boq && Array.isArray(data.boq) && data.boq.length > 0) {
              rowData[c.BOQ] = _ensureJSONString(data.boq, []);
            } else if (_dbBoq.length > 0) {
              rowData[c.BOQ] = _ensureJSONString(_dbBoq, []);
            }

            // IMAGES: ALWAYS preserved from DB here — image add/remove is
            // handled exclusively by savePATImages/deletePATImages so the form
            // array must never be authoritative in a normal project save.
            if (_dbImages.length > 0) {
              rowData[c.IMAGES] = _ensureJSONString(_dbImages, []);
            } else if (
              data.images &&
              Array.isArray(data.images) &&
              data.images.length > 0
            ) {
              rowData[c.IMAGES] = _ensureJSONString(data.images, []);
            }

            // Checklist / Snags / Signoff: keep DB unless form sent non-empty.
            if (
              data.checklist &&
              typeof data.checklist === "object" &&
              Object.keys(data.checklist).length > 0
            ) {
              rowData[c.CHECKLIST] = _ensureJSONString(data.checklist, {});
            } else if (_dbCheck && Object.keys(_dbCheck).length > 0) {
              rowData[c.CHECKLIST] = _ensureJSONString(_dbCheck, {});
            }
            if (
              data.snags &&
              Array.isArray(data.snags) &&
              data.snags.length > 0
            ) {
              rowData[c.SNAGS] = _ensureJSONString(data.snags, []);
            } else if (_dbSnags.length > 0) {
              rowData[c.SNAGS] = _ensureJSONString(_dbSnags, []);
            }
            if (
              data.signoff &&
              typeof data.signoff === "object" &&
              Object.keys(data.signoff).length > 0
            ) {
              rowData[c.SIGNOFF] = _ensureJSONString(data.signoff, {});
            } else if (_dbSignoff && Object.keys(_dbSignoff).length > 0) {
              rowData[c.SIGNOFF] = _ensureJSONString(_dbSignoff, {});
            }

            sh.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
            SpreadsheetApp.flush();
            // Audit: every edit is logged, cumulative
            try {
              var _entryMEC = _buildEditLogEntry(
                sess,
                existingProject,
                rowData,
              );
              _appendEditLog(i, _entryMEC);
            } catch (e) {
              console.warn("editlog(MEC): " + e.message);
            }
            return { success: true, projectId: existingId };
          }

          // No permission to update
          return {
            success: false,
            message:
              "Permission denied: you are not allowed to edit this project.",
          };
        }
      }
    }

    // New project
    var sessRole = String(sess.role || "").toLowerCase();
    var sessDept = String(sess.department || "").toLowerCase();
    var isMEC = sessDept === "mec" || sessDept === "mech";
    var isAdmin = sessRole === "admin" || sessRole === "super admin";

    if (!isMEC) {
      return {
        success: false,
        message:
          "Access Denied: You do not have the necessary permissions to create or manage PAT projects. Even the AI will not open the workspace for this action. Only MEC personnel are authorized for PAT creation.",
      };
    }

    // DEDUPLICATION CHECK: Prevent AI from creating duplicate projects in rapid succession
    var projectData = sh.getDataRange().getValues();
    var nameToMatch = String(data.projectName || "")
      .trim()
      .toLowerCase();
    var addrToMatch = String(data.siteAddress || "")
      .trim()
      .toLowerCase();
    var userEmail = String(sess.email || "").toLowerCase();
    var nowTs = new Date().getTime();

    if (nameToMatch && addrToMatch) {
      for (var j = projectData.length - 1; j >= 1; j--) {
        var rowName = String(projectData[j][c.PROJECT_NAME] || "")
          .trim()
          .toLowerCase();
        var rowAddr = String(projectData[j][c.SITE_ADDRESS] || "")
          .trim()
          .toLowerCase();
        var rowEmail = String(
          projectData[j][c.SUBMITTED_BY_EMAIL] || "",
        ).toLowerCase();
        var rowDateStr = projectData[j][c.SUBMITTED_AT];
        var rowTs = rowDateStr ? new Date(rowDateStr).getTime() : 0;

        // If same name, address, user and created within last 60 seconds
        if (
          rowName === nameToMatch &&
          rowAddr === addrToMatch &&
          rowEmail === userEmail &&
          nowTs - rowTs < 60000
        ) {
          var existingId = String(projectData[j][c.PROJECT_ID] || "").trim();
          return {
            success: true,
            projectId: existingId,
            message: "Existing project reused to prevent duplication.",
          };
        }
      }
    }

    var newId = _genPATId();
    data.projectId = newId;
    var newRow = _rowFromProject(data);
    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    // Audit: creation is the first log entry (cumulative history)
    try {
      var _newIdx = sh.getLastRow() - 1;
      var _entryNew = _buildEditLogEntry(sess, {}, data);
      _entryNew.action = "created";
      _appendEditLog(_newIdx, _entryNew);
    } catch (e) {
      console.warn("editlog(new): " + e.message);
    }
    return { success: true, projectId: newId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// VENDOR REVIEW WORKFLOW
// ─────────────────────────────────────────────────────────────

/**
 * Generate a vendor review link for a PAT project.
 * Requires session. Always generates a fresh token, sets hasVendor=TRUE,
 * and resets vendorApprovalStatus to 'pending'.
 * Only available at Awaiting Project Team stage.
 */
function generateVendorReviewLink(token, projectId) {
  try {
    var sess = _session(token);
    var dept = String(sess.department || "").toLowerCase();
    var isProjectTeam = dept.includes("project");
    var isSDMetro = dept.includes("service") || dept.includes("metro");

    if (!isProjectTeam && !isSDMetro) {
      return {
        success: false,
        message:
          "Only Project Team or Service Delivery / Metro members can generate vendor review links.",
      };
    }

    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // Vendor review is requested at the ENTRY stage of a PAT's route:
    //  - Project Team route  → Awaiting Project Team
    //  - Service Delivery / Metro route → Awaiting Service Delivery (Metro)
    var canVendor =
      project.workflowStatus === "Awaiting Project Team" ||
      project.workflowStatus === "Awaiting Service Delivery (Metro)";
    if (!canVendor) {
      return {
        success: false,
        message:
          "Vendor review is only available at the Awaiting Project Team or Awaiting Service Delivery (Metro) stage.",
      };
    }

    var c = _patCols();
    var sh = _patSheet();
    var data = sh.getDataRange().getValues();
    var searchId = String(projectId).trim();
    var altId = searchId.replace(/^(FOB|PAT)-/i, "");
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      if (
        rowId.toUpperCase() === searchId.toUpperCase() ||
        rowId.toUpperCase() === altId.toUpperCase() ||
        ("PAT-" + rowId).toUpperCase() === searchId.toUpperCase()
      ) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "Project not found." };

    var sheetRow = rowIndex + 1;

    // Read current vendor status to decide whether to reset
    var currentVendorApproval = String(
      data[rowIndex][c.VENDOR_APPROVAL_STATUS] || "pending",
    ).toLowerCase();
    var currentEverApproved =
      String(data[rowIndex][c.VENDOR_EVER_APPROVED] || "").toUpperCase() ===
      "TRUE";

    // Always generate a fresh token (overwrite old one)
    var vendorToken =
      "VEND-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    sh.getRange(sheetRow, c.VENDOR_TOKEN + 1).setValue(vendorToken);

    // Set hasVendor = TRUE
    sh.getRange(sheetRow, c.HAS_VENDOR + 1).setValue("TRUE");

    // Only reset approval status if vendor has NEVER approved (preserve ever-approved)
    if (!currentEverApproved && currentVendorApproval !== "approved") {
      sh.getRange(sheetRow, c.VENDOR_APPROVAL_STATUS + 1).setValue("pending");
      sh.getRange(sheetRow, c.VENDOR_APPROVAL_COMMENTS + 1).setValue("");
      sh.getRange(sheetRow, c.VENDOR_APPROVAL_DATE + 1).setValue("");
    }

    var baseUrl = ScriptApp.getService().getUrl();
    var fullUrl = baseUrl + "?token=" + encodeURIComponent(vendorToken);

    return { success: true, url: fullUrl, vendorToken: vendorToken };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Get a PAT project by vendor token (public, no session required).
 * Returns anonymised project data for vendor review.
 */
function getVendorPATByToken(vendorToken) {
  try {
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };
    var data = sh.getDataRange().getValues();
    var c = _patCols();

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][c.VENDOR_TOKEN] || "").toUpperCase() ===
        String(vendorToken).toUpperCase()
      ) {
        var approvalStatus = String(
          data[i][c.VENDOR_APPROVAL_STATUS] || "pending",
        );
        var everApproved =
          String(data[i][c.VENDOR_EVER_APPROVED] || "").toUpperCase() ===
          "TRUE";
        var alreadyCompleted = approvalStatus !== "pending" && !everApproved;
        if (alreadyCompleted) {
          return {
            success: false,
            message: "This review link has already been used.",
          };
        }
        var project = _projectFromRow(data[i]);
        // Return relevant vendor-facing data only
        var reviewCompleted = everApproved || approvalStatus !== "pending";
        return {
          success: true,
          reviewCompleted: reviewCompleted,
          project: {
            projectId: project.projectId,
            projectName: project.projectName,
            siteAddress: project.siteAddress,
            phase: project.phase,
            workDescription: project.workDescription,
            vendor: project.vendor,
            inspectionDate: project.inspectionDate,
            orchestrator: project.orchestrator,
            snagScore: project.snagScore,
            checklist: project.checklist,
            boq: project.boq,
            snags: project.snags,
            signoff: project.signoff,
            images: project.images,
            verdict: project.verdict,
            presidingOfficer: project.presidingOfficer,
          },
        };
      }
    }
    return { success: false, message: "Invalid or expired review link." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Submit vendor decision (public, no session required).
 * decision: 'approved' or 'rejected'
 */
function submitVendorDecision(vendorToken, decision, comments) {
  try {
    decision = String(decision || "")
      .toLowerCase()
      .trim();
    if (decision !== "approved" && decision !== "rejected") {
      return {
        success: false,
        message: 'Decision must be "approved" or "rejected".',
      };
    }
    if (decision === "rejected" && !String(comments || "").trim()) {
      return {
        success: false,
        message: "A comment is required when rejecting.",
      };
    }

    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };
    var data = sh.getDataRange().getValues();
    var c = _patCols();

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][c.VENDOR_TOKEN] || "").toUpperCase() ===
        String(vendorToken).toUpperCase()
      ) {
        var currentStatus = String(
          data[i][c.VENDOR_APPROVAL_STATUS] || "pending",
        );
        var currentEverApproved =
          String(data[i][c.VENDOR_EVER_APPROVED] || "").toUpperCase() ===
          "TRUE";
        if (currentStatus !== "pending" || currentEverApproved) {
          return {
            success: false,
            message: "This review has already been completed.",
          };
        }
        var sheetRow = i + 1;
        sh.getRange(sheetRow, c.VENDOR_APPROVAL_STATUS + 1).setValue(decision);
        sh.getRange(sheetRow, c.VENDOR_APPROVAL_COMMENTS + 1).setValue(
          String(comments || ""),
        );
        sh.getRange(sheetRow, c.VENDOR_APPROVAL_DATE + 1).setValue(
          new Date().toISOString(),
        );

        // PERMANENT: Once vendor approves, set VendorEverApproved=TRUE (never reset)
        if (decision === "approved") {
          sh.getRange(sheetRow, c.VENDOR_EVER_APPROVED + 1).setValue("TRUE");
        }

        // When vendor rejects, push the PAT BACK so it cannot sit frozen at the
        // vendor-entry stage. It returns to the PROJECT TEAM (the department
        // that owns vendor engagement) — NOT to MEC. This mirrors the portal's
        // internal rejection rule, where every department rejects back to the
        // Project Team. The Project Team then reviews the vendor's reason and
        // rejects back to MEC (status "Rejected") with their own comment.
        if (decision === "rejected") {
          // Vendor rejection returns the PAT to the PROJECT TEAM — the department
          // that owns vendor engagement — NOT to MEC.
          var existingProject = _projectFromRow(data[i]);
          var priorStatus =
            existingProject.workflowStatus || "Awaiting Project Team";
          var history = existingProject.workflowHistory || [];
          var vendorRejectComment = "VENDOR REJECTED: " + (comments || "");
          history.push({
            fromStatus: priorStatus,
            toStatus: "Awaiting Project Team",
            by: {
              name: "Vendor",
              email: existingProject.vendor || "vendor",
              department: "Vendor",
            },
            comments: vendorRejectComment,
            isRejection: true,
            timestamp: new Date().toISOString(),
          });
          // Return the PAT to the Project Team's inbox.
          sh.getRange(sheetRow, c.WORKFLOW_STATUS + 1).setValue(
            "Awaiting Project Team",
          );
          sh.getRange(sheetRow, c.ASSIGNED_TO_DEPT + 1).setValue(
            "Project Team",
          );
          sh.getRange(sheetRow, c.ASSIGNED_TO_NAME + 1).setValue("");
          sh.getRange(sheetRow, c.ASSIGNED_TO_EMAIL + 1).setValue("");
          sh.getRange(sheetRow, c.VERDICT + 1).setValue("Rejected");
          sh.getRange(sheetRow, c.REJECTION_REASON + 1).setValue(
            vendorRejectComment,
          );
          sh.getRange(sheetRow, c.WORKFLOW_HISTORY + 1).setValue(
            JSON.stringify(history),
          );
          sh.getRange(sheetRow, c.UPDATED_AT + 1).setValue(
            new Date().toISOString(),
          );
        }

        return { success: true, decision: decision };
      }
    }
    return { success: false, message: "Invalid token." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Check if vendor approval is required and has been given for a project.
 * Returns { required: bool, approved: bool, status: string }
 */
function _checkVendorApproval(project) {
  if (!project || !project.hasVendor) {
    return { required: false, approved: true, status: "not_required" };
  }
  var status = String(project.vendorApprovalStatus || "pending").toLowerCase();
  var everApproved = project.vendorEverApproved === true;
  return {
    required: true,
    approved: everApproved || status === "approved",
    status: everApproved ? "approved" : status,
  };
}

/**
 * Submit PAT to a department — advances the workflow.
 */
/**
 * Submit PAT to a department — advances the workflow.
 * @param {string} entryRoute - optional 'project' | 'metro' — recorded on the first
 *        push out of Draft so the whole branch knows which route this PAT took.
 */
function submitPATToDepartment(
  token,
  projectId,
  targetDept,
  comments,
  entryRoute,
) {
  try {
    var sess = _session(token);
    // Use null as token for internal call — no session check needed for read
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // Determine next workflow status based on target department
    var statusMap = {
      "project team": "Awaiting Project Team",
      project: "Awaiting Project Team",
      "planning team": "Awaiting Planning Team",
      planning: "Awaiting Planning Team",
      "service delivery / metro": "Awaiting Service Delivery",
      "service delivery": "Awaiting Service Delivery",
      metro: "Awaiting Service Delivery",
      mec: "Awaiting Final MEC Review",
      mech: "Awaiting Final MEC Review",
    };
    var targetDeptNorm = String(targetDept).toLowerCase().trim();
    var newStatus = statusMap[targetDeptNorm] || "Awaiting Project Team";

    // ROUTE-AWARE ENTRY: When MEC pushes from Draft, the "Service Delivery / Metro"
    // department can mean EITHER the metro ENTRY stage (skips Project Team) OR the
    // normal post-Planning SD/Metro stage. Decide by the project's entry route.
    var explicitRoute =
      entryRoute === "metro" || entryRoute === "project"
        ? entryRoute
        : _entryRouteOf(project);
    if (
      project.workflowStatus === "Draft" ||
      project.workflowStatus === "Rejected"
    ) {
      if (
        targetDeptNorm.indexOf("metro") !== -1 ||
        targetDeptNorm.indexOf("service delivery") !== -1
      ) {
        // Metro department chosen at entry → metro ENTRY stage (skips Project Team).
        if (explicitRoute === "metro") {
          newStatus = "Awaiting Service Delivery (Metro)";
        }
      }
    }

    // Determine if this is a rejection first to inform the workflow transition
    var isRejection = _isRejection(newStatus, comments);

    // MEC REJECT FROM FINAL REVIEW: route back to the project's entry department
    // (Project Team for project-route, SD/Metro (Metro) for metro-route).
    if (project.workflowStatus === "Awaiting Final MEC Review" && isRejection) {
      var fr = _entryRouteOf(project);
      newStatus =
        fr === "metro"
          ? "Awaiting Service Delivery (Metro)"
          : "Awaiting Project Team";
      project.verdict = "Rejected";
    }

    // RECHECK LOGIC: If pushing to MEC from Project Team / Metro entry, handle based on previous failure stage
    if (
      (targetDeptNorm === "mec" || targetDeptNorm === "mech") &&
      (project.workflowStatus === "Awaiting Project Team" ||
        project.workflowStatus === "Awaiting Service Delivery (Metro)" ||
        project.workflowStatus === "Rejected")
    ) {
      if (isRejection) {
        newStatus = "Awaiting MEC Recheck";
      } else if (project.verdict === "Rejected") {
        // If it failed Final Review, it must go back to Final Review for approval
        newStatus = "Awaiting Final MEC Review";
      } else {
        // If it's an early stage approval, it goes to MEC Recheck
        newStatus = "Awaiting MEC Recheck";
      }
    }

    // Handle other rejections
    if (
      isRejection &&
      targetDeptNorm.indexOf("mec") !== -1 &&
      newStatus !== "Awaiting MEC Recheck"
    ) {
      // Status "Rejected" is now reserved for MEC inbox (unless it's a project-to-mec recheck)
      newStatus = "Rejected";
    } else if (isRejection && newStatus !== "Awaiting MEC Recheck") {
      // Rejections to other teams stay in their workflow stage
      newStatus = statusMap[targetDeptNorm] || "Awaiting Project Team";
    }

    var oldStatus = project.workflowStatus || "Draft";

    // Build history entry
    var historyEntry = {
      fromStatus: oldStatus,
      toStatus: newStatus,
      by: { name: sess.name, email: sess.email, department: sess.department },
      comments: comments || "",
      isRejection: isRejection, // Explicitly store if this was a rejection
      timestamp: new Date().toISOString(),
    };
    // Record the entry route on the FIRST push out of Draft (recovery anchor).
    if ((oldStatus === "Draft" || oldStatus === "Rejected") && explicitRoute) {
      historyEntry.entryRoute = explicitRoute;
    }
    var history = project.workflowHistory || [];
    history.push(historyEntry);

    // Update project
    project.workflowStatus = newStatus;
    project.workflowHistory = history;
    if (isRejection) {
      project.rejectionReason = comments || "";
    }
    project.department = targetDept; // Department currently holding the PAT
    project.assignedToDept = targetDept;
    project.assignedToName = "";
    project.assignedToEmail = "";

    var saveRes = savePATProject(token, project);
    return {
      success: saveRes.success,
      message: "Submitted to " + targetDept + ". Status: " + newStatus,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Internal helper to determine if a workflow transition is a rejection.
 */
function _isRejection(status, comments) {
  if (status === "Rejected") return true;
  if (
    String(comments || "")
      .toUpperCase()
      .indexOf("REJECTED") === 0
  )
    return true;
  return false;
}

/**
 * Internal helper to find the last department that rejected the project
 * and has not yet had their rejection resolved by an approval.
 */
function _getLastUnresolvedRejecter(project, currentDept) {
  var history = project.workflowHistory || [];
  var normCurrentDept = _normalizeDept(currentDept);

  // 1. Find all rejections in the history
  var rejections = [];
  for (var i = 0; i < history.length; i++) {
    var h = history[i];
    if (!h) continue;
    if (h.isRejection === true || _isRejection(h.toStatus, h.comments)) {
      rejections.push({
        dept: h.by ? h.by.department : null,
        index: i,
        entry: h,
      });
    }
  }

  // 2. Filter for rejections that have NOT been resolved by a later approval from the same dept
  var unresolved = rejections.filter(function (rej) {
    if (!rej.dept) return false;
    var normRejDept = _normalizeDept(rej.dept);

    // Look forward in history for any non-rejection action by this department
    for (var j = rej.index + 1; j < history.length; j++) {
      var h = history[j];
      if (!h) continue;

      var isLaterRejection =
        h.isRejection === true || _isRejection(h.toStatus, h.comments);
      if (!isLaterRejection && h.by && h.by.department) {
        if (_normalizeDept(h.by.department) === normRejDept) {
          return false; // This rejection was resolved by a later approval
        }
      }
    }
    return true; // Still unresolved
  });

  // 3. Return the most recent unresolved rejection entry who is NOT the current approver
  for (var k = unresolved.length - 1; k >= 0; k--) {
    var rejDept = unresolved[k].dept;
    if (rejDept && _normalizeDept(rejDept) !== normCurrentDept) {
      return unresolved[k].entry;
    }
  }

  return null;
}

/**
 * Approve and move to next stage with specific assignee.
 */
function submitPATToNextStage(token, projectId, nextEmail, comments) {
  try {
    var sess = _session(token);
    // Use null token for internal call
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // Vendor approval check
    var vendorCheck = _checkVendorApproval(project);
    if (vendorCheck.required && !vendorCheck.approved) {
      var msg =
        vendorCheck.status === "rejected"
          ? "Vendor rejected this project. Please resolve with vendor before proceeding."
          : "Vendor approval is required before proceeding. Please wait for vendor response.";
      return { success: false, message: msg };
    }

    var newStatus =
      _wfStageFlow(_entryRouteOf(project))[project.workflowStatus] || null;

    // ROUTING FIX: If this project has unresolved rejections, return to the last rejecter
    var lastRejEntry = _getLastUnresolvedRejecter(project, sess.department);
    if (lastRejEntry) {
      var rejDept = lastRejEntry.by ? lastRejEntry.by.department : "";
      var normalized = _normalizeStatusByDept(rejDept);

      if (_normalizeDept(rejDept) === "mec") {
        // When Project Team resolves a MEC rejection:
        // If it was rejected from Final Review -> go back to Final Review
        // If it was rejected from Draft/Other -> go to MEC Recheck
        newStatus =
          lastRejEntry.fromStatus === "Awaiting Final MEC Review"
            ? "Awaiting Final MEC Review"
            : "Awaiting MEC Recheck";
      } else {
        newStatus = normalized;
      }
    }
    // All MEC actions on a Rejected project now route through Project Team by default

    var history = project.workflowHistory || [];
    history.push({
      fromStatus: project.workflowStatus,
      toStatus: newStatus,
      by: { name: sess.name, email: sess.email, department: sess.department },
      comments: comments || "",
      timestamp: new Date().toISOString(),
    });

    project.workflowStatus = newStatus;
    project.workflowHistory = history;

    // Update department and assigned department based on the new status
    var deptForStatus = STATUS_TO_DEPT[newStatus] || "Project Team";
    project.department = deptForStatus;
    project.assignedToDept = deptForStatus;

    // If assigned to specific person
    if (nextEmail) {
      project.assignedToEmail = nextEmail;
      // Look up name/dept
      try {
        var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          SD.USERS,
        );
        var userData = userSh.getDataRange().getValues();
        for (var i = 1; i < userData.length; i++) {
          if (
            String(userData[i][2]).toLowerCase() === nextEmail.toLowerCase()
          ) {
            project.assignedToName = userData[i][1];
            project.assignedToDept = userData[i][6];
            project.department = userData[i][6];
            break;
          }
        }
      } catch (e) {}
    }

    // If the new status is Awaiting Final MEC Review, assign to MEC department generally
    if (newStatus === "Awaiting Final MEC Review") {
      project.assignedToDept = "MEC";
      project.assignedToName = ""; // Clear specific assignment
      project.assignedToEmail = ""; // Clear specific assignment
    }

    var saveRes = savePATProject(token, project);
    return {
      success: saveRes.success,
      message: "Approved to " + newStatus + ".",
      verdict: "Fully Accepted",
      snagScore: project.snagScore,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Partially approve — move to next stage with snags + Partially Approved verdict.
 */
function partiallyApprovePAT(
  token,
  projectId,
  comments,
  newSnags,
  nextDept,
  nextEmail,
) {
  try {
    var sess = _session(token);
    // Use null token for internal call
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // Vendor approval check
    var vendorCheck = _checkVendorApproval(project);
    if (vendorCheck.required && !vendorCheck.approved) {
      var msg =
        vendorCheck.status === "rejected"
          ? "Vendor rejected this project. Please resolve with vendor before proceeding."
          : "Vendor approval is required before proceeding. Please wait for vendor response.";
      return { success: false, message: msg };
    }

    // Append new snags
    var snags = project.snags || [];
    if (newSnags && newSnags.length) {
      snags = snags.concat(newSnags);
    }

    // Recalculate score
    var score = 0;
    snags.forEach(function (s) {
      score += parseInt(s.weight) || 0;
    });

    project.snags = snags;
    project.snagScore = score;
    project.verdict = "Partially Accepted";

    var newStatus =
      _wfStageFlow(_entryRouteOf(project))[project.workflowStatus] ||
      "Partially Approved";

    // If resolving a rejection, jump back to that department's review stage
    var lastRejEntry = _getLastUnresolvedRejecter(project, sess.department);
    if (lastRejEntry) {
      var rejDept = lastRejEntry.by ? lastRejEntry.by.department : "";
      var normalized = _normalizeStatusByDept(rejDept);

      if (_normalizeDept(rejDept) === "mec") {
        // When Project Team resolves a MEC rejection:
        // If it was rejected from Final Review -> go back to Final Review
        // If it was rejected from Draft/Other -> go to MEC Recheck
        newStatus =
          lastRejEntry.fromStatus === "Awaiting Final MEC Review"
            ? "Awaiting Final MEC Review"
            : "Awaiting MEC Recheck";
      } else {
        newStatus = normalized;
      }
    }

    project.workflowStatus = newStatus;

    // Update department and assigned department based on the new status
    var deptForStatus = STATUS_TO_DEPT[newStatus] || "Project Team";
    project.department = deptForStatus;
    project.assignedToDept = deptForStatus;

    var history = project.workflowHistory || [];
    history.push({
      fromStatus: project.workflowStatus,
      toStatus: newStatus,
      by: { name: sess.name, email: sess.email, department: sess.department },
      comments: "PARTIALLY APPROVED: " + (comments || ""),
      timestamp: new Date().toISOString(),
    });
    project.workflowHistory = history;

    // Update assignment
    if (nextDept) project.assignedToDept = nextDept;
    if (nextEmail) {
      project.assignedToEmail = nextEmail;
      try {
        var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          SD.USERS,
        );
        var userData = userSh.getDataRange().getValues();
        for (var i = 1; i < userData.length; i++) {
          if (
            String(userData[i][2]).toLowerCase() === nextEmail.toLowerCase()
          ) {
            project.assignedToName = userData[i][1];
            break;
          }
        }
      } catch (e) {}
    } else {
      project.assignedToName = "";
      project.assignedToEmail = "";
    }

    // If the new status is Awaiting Final MEC Review, assign to MEC department generally
    if (newStatus === "Awaiting Final MEC Review") {
      project.assignedToDept = "MEC";
      project.assignedToName = ""; // Clear specific assignment
      project.assignedToEmail = "";
    }
    // If the new status is Awaiting Final MEC Review, assign to MEC department generally
    if (newStatus === "Awaiting Final MEC Review") {
      project.assignedToDept = "MEC";
      project.assignedToName = ""; // Clear specific assignment
      project.assignedToEmail = "";
    }

    var saveRes = savePATProject(token, project);
    return {
      success: saveRes.success,
      message: "Partially approved. Score: " + score,
      verdict: "Partially Accepted",
      snagScore: score,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * AUTHORITATIVE workflow router.
 * Computes the correct next status for any action based on the project's CURRENT
 * status, its ENTRY ROUTE (project | metro), and the action (approve | partial | reject).
 * This keeps routing rules in ONE place so every caller (MEC push, assign modal,
 * quick actions, reject) produces correct, route-aware transitions.
 *
 * @param {string} token       - session token
 * @param {string} projectId   - PAT project id
 * @param {string} action      - 'approve' | 'partial' | 'reject'
 * @param {string} comments    - free-text comments / rejection reason
 * @param {string} recipientEmail - optional specific assignee email for the next stage
 * @param {string} entryRoute  - optional 'project' | 'metro' (only meaningful on first push from Draft)
 * @returns {{success, message, newStatus, verdict, snagScore}}
 */
function routePAT(
  token,
  projectId,
  action,
  comments,
  recipientEmail,
  entryRoute,
) {
  try {
    var sess = _session(token);
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    var route = _entryRouteOf(project);
    if (entryRoute === "metro" || entryRoute === "project") route = entryRoute;

    var current = project.workflowStatus || "Draft";
    var act = String(action || "approve").toLowerCase();
    var isReject = act === "reject" || _isRejection(null, comments);
    var flow = _wfStageFlow(route);

    var newStatus;

    // ── MEC first push out of Draft / Rejected ──
    if ((current === "Draft" || current === "Rejected") && !isReject) {
      newStatus = _firstStageOf(route); // project → Awaiting Project Team, metro → Awaiting Service Delivery (Metro)
    }
    // ── MEC reject from Final Review → back to entry department ──
    else if (current === "Awaiting Final MEC Review" && isReject) {
      newStatus =
        route === "metro"
          ? "Awaiting Service Delivery (Metro)"
          : "Awaiting Project Team";
      project.verdict = "Rejected";
    }
    // ── MEC recheck handling (pushing to MEC from entry stage) ──
    else if (
      (current === "Awaiting Project Team" ||
        current === "Awaiting Service Delivery (Metro)") &&
      route === "metro" &&
      _normalizeDept(sess.department) === "mec" &&
      !isReject
    ) {
      newStatus =
        project.verdict === "Rejected"
          ? "Awaiting Final MEC Review"
          : "Awaiting MEC Recheck";
    } else if (
      (current === "Awaiting Project Team" ||
        current === "Awaiting Service Delivery (Metro)") &&
      _normalizeDept(sess.department) === "mec" &&
      !isReject
    ) {
      newStatus =
        project.verdict === "Rejected"
          ? "Awaiting Final MEC Review"
          : "Awaiting MEC Recheck";
    }
    // ── Rejections at any stage → route-aware rejection target ──
    else if (isReject) {
      newStatus = _rejectionTarget(current, route);
      if (
        current === "Awaiting Service Delivery (Metro)" ||
        (route === "metro" && current === "Awaiting Service Delivery")
      ) {
        newStatus = "Rejected"; // SD/Metro reject of metro-route → MEC inbox
      }
      project.verdict = "Rejected";
    }
    // ── Standard forward approval (next linear stage) ──
    else {
      newStatus = flow[current] || _nextStageOf(current, route) || "Completed";
    }

    // Build history entry
    var history = project.workflowHistory || [];
    var historyEntry = {
      fromStatus: current,
      toStatus: newStatus,
      by: { name: sess.name, email: sess.email, department: sess.department },
      comments: comments || "",
      isRejection: isReject,
      timestamp: new Date().toISOString(),
    };
    if ((current === "Draft" || current === "Rejected") && route)
      historyEntry.entryRoute = route;
    history.push(historyEntry);

    project.workflowStatus = newStatus;
    project.workflowHistory = history;
    if (isReject) project.rejectionReason = comments || "";

    // Vendor approval bookkeeping for partial approve
    var snagScore = project.snagScore || 0;
    if (act === "partial") {
      project.verdict = "Partially Accepted";
      // vendor must be approved to partial-approve forward
      var vCheck = _checkVendorApproval(project);
      if (vCheck.required && !vCheck.approved) {
        return {
          success: false,
          message:
            vCheck.status === "rejected"
              ? "Vendor rejected this project. Please resolve with vendor before proceeding."
              : "Vendor approval is required before proceeding. Please wait for vendor response.",
        };
      }
    }

    // Department / assignee resolution
    var deptForStatus = STATUS_TO_DEPT[newStatus] || "MEC";
    project.department = deptForStatus;
    project.assignedToDept = deptForStatus;

    if (recipientEmail) {
      project.assignedToEmail = recipientEmail;
      try {
        var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          SD.USERS,
        );
        var userData = userSh.getDataRange().getValues();
        for (var i = 1; i < userData.length; i++) {
          if (
            String(userData[i][2]).toLowerCase() ===
            recipientEmail.toLowerCase()
          ) {
            project.assignedToName = userData[i][1];
            project.assignedToDept = userData[i][6];
            project.department = userData[i][6];
            break;
          }
        }
      } catch (e) {}
    } else {
      project.assignedToName = "";
      project.assignedToEmail = "";
    }

    // Final MEC review → clear specific assignee
    if (newStatus === "Awaiting Final MEC Review") {
      project.assignedToDept = "MEC";
      project.assignedToName = "";
      project.assignedToEmail = "";
    }

    var saveRes = savePATProject(token, project);
    var msg = isReject
      ? "Rejected & returned. Status: " + newStatus
      : "Approved to " + newStatus + ".";
    return {
      success: saveRes.success,
      message: msg,
      newStatus: newStatus,
      verdict: project.verdict,
      snagScore: snagScore,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Helper to get canonical status for a department
 */
function _normalizeStatusByDept(dept) {
  var d = String(dept || "")
    .toLowerCase()
    .trim();
  if (d.includes("planning")) return "Awaiting Planning Team";
  if (d.includes("service") || d.includes("delivery") || d.includes("metro"))
    return "Awaiting Service Delivery";
  if (d.includes("mec") || d.includes("mech"))
    return "Awaiting Final MEC Review";
  if (d.includes("project")) return "Awaiting Project Team";
  return "Awaiting Project Team";
}

/**
 * Internal helper: Delete all JCC certificates linked to a given PAT project ID.
 * Scans the SD_JCC_CERTIFICATES sheet and removes any row whose ProjectID matches.
 * Iterates backwards to safely delete rows.
 *
 * @param {string} projectId - The PAT project ID whose JCCs should be removed
 * @returns {{ success, deletedCount }}
 */
function _deleteJCCByProjectId(projectId) {
  try {
    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    if (!jccSh || jccSh.getLastRow() < 2)
      return { success: true, deletedCount: 0 };

    var searchId = String(projectId || "")
      .trim()
      .toUpperCase();
    var altId = searchId.replace(/^(FOB|PAT)-/i, "");

    var data = jccSh.getDataRange().getValues();
    var deletedCount = 0;

    // Iterate backwards so row indices stay valid after deletion
    for (var i = data.length - 1; i >= 1; i--) {
      var jccProjectId = String(data[i][1] || "")
        .trim()
        .toUpperCase();
      if (
        jccProjectId === searchId ||
        jccProjectId === altId ||
        "PAT-" + jccProjectId === searchId
      ) {
        jccSh.deleteRow(i + 1); // +1 because sheet is 1-indexed
        deletedCount++;
      }
    }

    return { success: true, deletedCount: deletedCount };
  } catch (e) {
    // Non-critical: log error but don't block PAT deletion
    console.error(
      "_deleteJCCByProjectId error for " + projectId + ": " + e.message,
    );
    return { success: false, deletedCount: 0, message: e.message };
  }
}

/**
 * Delete all Drive files linked in a PAT project's image URLs.
 * Urls like "https://drive.google.com/uc?export=view&id=FILE_ID" have their Drive file trashed.
 */
function _deletePATDriveImages(projectId, data, foundRow) {
  try {
    if (!data || !data[foundRow]) return { success: true, deletedCount: 0 };
    var c = _patCols();
    var raw = data[foundRow][c.IMAGES] || "[]";
    var urls = [];
    try {
      urls = JSON.parse(raw);
    } catch (e) {
      urls = [];
    }
    if (!Array.isArray(urls) || urls.length === 0)
      return { success: true, deletedCount: 0 };
    var deletedCount = 0;
    urls.forEach(function (url) {
      var match = String(url).match(/[?&]id=([^&]+)/);
      if (match) {
        try {
          var file = DriveApp.getFileById(match[1]);
          file.setTrashed(true);
          deletedCount++;
        } catch (e) {
          console.warn("Could not delete Drive image: " + e.message);
        }
      }
    });
    return { success: true, deletedCount: deletedCount };
  } catch (e) {
    console.error("_deletePATDriveImages error: " + e.message);
    return { success: false, deletedCount: 0, message: e.message };
  }
}

/**
 * Delete ONE PAT Drive image (by URL) and its linked SD_DOCUMENTS row.
 * Used by the reconcile logic in savePATImages so that a removed image is
 * actually trashed from Drive and never "comes back" on the next save.
 */
function _deleteOnePATDriveImage(url, projectId) {
  if (!url) return 0;
  var deleted = 0;
  // 1) Trash the Drive file (if the URL carries a file id)
  var match = String(url).match(/[?&]id=([^&]+)/);
  if (!match) match = String(url).match(/googleusercontent\.com\/d\/([^/?]+)/);
  if (!match) match = String(url).match(/\/file\/d\/([^/?]+)/);
  if (match) {
    try {
      var file = DriveApp.getFileById(match[1]);
      file.setTrashed(true);
      deleted++;
    } catch (e) {
      console.warn("Could not trash PAT image: " + e.message);
    }
  }
  // 2) Remove the matching Site Document row (image -> doc)
  try {
    var docSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DOCS);
    if (docSh && docSh.getLastRow() >= 2) {
      var c = _docCols();
      var dData = docSh.getDataRange().getValues();
      var key = String(projectId || "")
        .toUpperCase()
        .trim();
      var normId = key.replace(/^(FOB|PAT)-/i, "");
      for (var i = dData.length - 1; i >= 1; i--) {
        var dPid = String(dData[i][1] || "")
          .toUpperCase()
          .trim();
        var dUrl = String(dData[i][5] || "");
        var isProject =
          dPid === key || dPid === normId || "PAT-" + dPid === key;
        if (
          isProject &&
          (dUrl === url || (match && dUrl.indexOf(match[1]) !== -1))
        ) {
          docSh.deleteRow(i + 1);
          break; // only one doc row per image
        }
      }
    }
  } catch (e) {
    console.warn("PAT image doc-row cleanup: " + e.message);
  }
  return deleted;
}

/**
 * Delete a single PAT project.
 * Also cascade-deletes linked JCC certificates, documents, and Drive files.
 */
function deletePATProject(token, projectId) {
  try {
    var sess = _session(token);
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };
    var data = sh
      .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .getValues();

    var searchId = String(projectId || "").trim();
    var altId = searchId.replace(/^(FOB|PAT)-/i, "");

    var foundRow = -1;
    var foundProject = null;

    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0] || "").trim();
      if (
        rowId.toUpperCase() === searchId.toUpperCase() ||
        rowId.toUpperCase() === altId.toUpperCase() ||
        ("PAT-" + rowId).toUpperCase() === searchId.toUpperCase()
      ) {
        foundRow = i;
        foundProject = _projectFromRow(data[i]);
        break;
      }
    }

    if (foundRow === -1)
      return { success: false, message: "Project not found: " + projectId };

    // SAFETY: re-verify the row we are about to delete actually carries the
    // requested ID. A stale read (e.g. an in-flight delete changing row indices,
    // or an index-based lookup that drifted) must NEVER delete the wrong row.
    var confirmedId = String(data[foundRow][0] || "").trim();
    var confirmedNorm = confirmedId.toUpperCase();
    var requestedNorm = searchId.toUpperCase();
    var altNorm = altId.toUpperCase();
    var isMatch =
      confirmedNorm === requestedNorm ||
      confirmedNorm === altNorm ||
      "PAT-" + confirmedNorm === requestedNorm ||
      confirmedNorm === "PAT-" + requestedNorm.replace(/^PAT-/, "");
    if (!isMatch) {
      return {
        success: false,
        message:
          'Delete aborted: row/ID mismatch (got "' +
          confirmedId +
          '" for requested "' +
          projectId +
          '"). Possible concurrent edit — refresh and retry.',
      };
    }

    // STRICT delete policy (mirrors frontend button visibility):
    //   - ONLY the MEC department (mec/mech) may delete a PAT.
    //   - MEC may delete ONLY while the PAT is still in DRAFT.
    //   - No other role (incl. Super Admin) and no other department may delete.
    //   - Once a PAT leaves Draft, it can no longer be deleted (must be worked through).
    var canDelete = false;
    var dept = String(sess.department || "").toLowerCase();
    if (dept === "mec" || dept === "mech") {
      var status = (foundProject ? foundProject.workflowStatus : "") || "";
      if (status === "Draft") canDelete = true;
    }

    if (!canDelete) {
      return {
        success: false,
        message: "Permission denied: You cannot delete this project.",
      };
    }

    // Cascade-delete JCC certificates
    var jccResult = _deleteJCCByProjectId(
      String(data[foundRow][0] || projectId),
    );

    // Cascade-delete linked documents from SD_DOCUMENTS and their Drive files
    var docResult = { deletedCount: 0 };
    try {
      var docSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DOCS);
      if (docSh && docSh.getLastRow() >= 2) {
        var docData = docSh.getDataRange().getValues();
        var projectKey = String(data[foundRow][0] || projectId)
          .toUpperCase()
          .trim();
        var rowsToDelete = [];
        for (var di = docData.length - 1; di >= 1; di--) {
          var docProjectId = String(docData[di][1] || "")
            .toUpperCase()
            .trim();
          var normId = projectKey.replace(/^(FOB|PAT)-/i, "");
          if (
            docProjectId === projectKey ||
            docProjectId === normId ||
            "PAT-" + docProjectId === projectKey
          ) {
            var driveFileId = docData[di][6] || "";
            if (driveFileId) {
              try {
                DriveApp.getFileById(driveFileId).setTrashed(true);
              } catch (e) {
                console.warn("Drive doc delete: " + e.message);
              }
            }
            rowsToDelete.push(di + 1);
            docResult.deletedCount++;
          }
        }
        rowsToDelete
          .sort(function (a, b) {
            return b - a;
          })
          .forEach(function (r) {
            docSh.deleteRow(r);
          });
      }
    } catch (e) {
      console.error("Document cascade delete error: " + e.message);
    }

    // Cascade-delete Drive images from PAT project
    var imgResult = _deletePATDriveImages(
      String(data[foundRow][0] || projectId),
      data,
      foundRow,
    );

    // Delete the PAT project row itself.
    // IMPORTANT: `data` was read starting at spreadsheet ROW 2 (header is row 1),
    // so the matched index `foundRow` corresponds to spreadsheet row `foundRow + 2`.
    // Deleting the wrong offset would remove the wrong project — strictly prohibited.
    sh.deleteRow(foundRow + 2);

    var msg = "Project deleted.";
    var details = [];
    if (jccResult.deletedCount > 0)
      details.push(jccResult.deletedCount + " JCC certificate(s)");
    if (docResult.deletedCount > 0)
      details.push(docResult.deletedCount + " document(s)");
    if (imgResult.deletedCount > 0)
      details.push(imgResult.deletedCount + " image(s)");
    if (details.length > 0) msg += " Also removed: " + details.join(", ") + ".";
    return { success: true, message: msg };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteAllDraftPATProjects(token) {
  try {
    _session(token);
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found.", deleted: 0 };

    var data = sh.getDataRange().getValues();
    var c = _patCols();
    var totalDeleted = 0;
    var totalJCCDeleted = 0;
    var totalDocDeleted = 0;
    var totalImgDeleted = 0;
    var deleted = [];

    // Iterate backwards to preserve row indices
    for (var i = data.length - 1; i >= 1; i--) {
      var wfStatus = String(data[i][c.WORKFLOW_STATUS] || "").trim();
      if (wfStatus === "Draft") {
        var projectId = String(data[i][c.PROJECT_ID] || "").trim();

        // Cascade-delete associated JCC certificates
        var jccResult = _deleteJCCByProjectId(projectId);
        totalJCCDeleted += jccResult.deletedCount || 0;

        // Cascade-delete associated documents and their Drive files
        try {
          var docSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
            SD.DOCS,
          );
          if (docSh && docSh.getLastRow() >= 2) {
            var docData = docSh.getDataRange().getValues();
            var projectKey = projectId.toUpperCase().trim();
            var normId = projectKey.replace(/^(FOB|PAT)-/i, "");
            var docRowsToDel = [];
            for (var di = docData.length - 1; di >= 1; di--) {
              var dPid = String(docData[di][1] || "")
                .toUpperCase()
                .trim();
              if (
                dPid === projectKey ||
                dPid === normId ||
                "PAT-" + dPid === projectKey
              ) {
                var dfid = docData[di][6] || "";
                if (dfid) {
                  try {
                    DriveApp.getFileById(dfid).setTrashed(true);
                  } catch (e) {}
                }
                docRowsToDel.push(di + 1);
                totalDocDeleted++;
              }
            }
            docRowsToDel
              .sort(function (a, b) {
                return b - a;
              })
              .forEach(function (r) {
                docSh.deleteRow(r);
              });
          }
        } catch (e) {
          console.error("Draft doc cascade error: " + e.message);
        }

        // Cascade-delete Drive images
        var imgResult = _deletePATDriveImages(projectId, data, i);
        totalImgDeleted += imgResult.deletedCount || 0;

        sh.deleteRow(i + 1); // +1 because sheet is 1-indexed
        deleted.push(projectId);
        totalDeleted++;
      }
    }
    var msg = "Deleted " + deleted.length + " draft projects.";
    var details = [];
    if (totalJCCDeleted > 0)
      details.push(totalJCCDeleted + " JCC certificate(s)");
    if (totalDocDeleted > 0) details.push(totalDocDeleted + " document(s)");
    if (totalImgDeleted > 0) details.push(totalImgDeleted + " image(s)");
    if (details.length > 0) msg += " Also removed: " + details.join(", ") + ".";
    return { success: true, message: msg };
  } catch (e) {
    return { success: false, message: e.message, deleted: 0 };
  }
}

function getPATImages(token, projectId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found.", images: [] };
    var data = sh
      .getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn())
      .getValues();
    var targetId = String(projectId || "")
      .toUpperCase()
      .trim();
    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        var c = _patCols();
        var images = _safeParse(data[i][c.IMAGES], []);
        return { success: true, images: images };
      }
    }
    return { success: false, message: "Project not found.", images: [] };
  } catch (e) {
    return { success: false, message: e.message, images: [] };
  }
}

/**
 * Check if a user can edit a project (upload/delete images, edit site info, etc.)
 * @param {object} sess   session object from _session(token)
 * @param {object} project PAT project object
 * @returns {boolean}
 */
function _canEditProject(sess, project) {
  if (!sess || !project) return false;
  const role = String(sess.role || "").toLowerCase();
  if (role === "super admin") return true;

  const dept = String(sess.department || "").toLowerCase();
  const normalizedDept = _normalizeDept(dept);
  const status = project.workflowStatus || "Draft";
  const isMEC = normalizedDept === "mec";
  const isDraftOrRejected = status === "Draft" || status === "Rejected";
  const isFinalReview = status === "Awaiting Final MEC Review";
  const isRecheck = status === "Awaiting MEC Recheck";

  // MEC can edit drafts, rejected, and final review (i.e. when it's their turn)
  if (isMEC && (isDraftOrRejected || isFinalReview || isRecheck)) return true;

  // Admin (non‑super) can edit only if they are explicitly assigned
  if (role === "admin") {
    const assignedDept = _normalizeDept(project.assignedToDept || "");
    const assignedEmail = (project.assignedToEmail || "").toLowerCase();
    if (assignedDept === normalizedDept) return true;
    if (assignedEmail === String(sess.email).toLowerCase()) return true;
  }
  return false;
}

/**
 * Uploads a SINGLE image to a project.
 * Much more reliable than bulk uploading.
 */
function addSinglePATImage(token, projectId, imageData) {
  try {
    const sess = _session(token);
    const projectRes = getPATProjectById(null, projectId);
    if (!projectRes.success) throw new Error("Project not found.");

    if (!_canEditProject(sess, projectRes.project)) {
      throw new Error(
        "Permission denied: You cannot upload images to this project right now.",
      );
    }

    const folder = _getPATFolder();
    const sh = _patSheet();
    const c = _patCols();

    // Get existing images
    const existingImagesRes = getPATImages(token, projectId);
    let finalUrls = existingImagesRes.success ? existingImagesRes.images : [];

    if (finalUrls.length >= 4)
      throw new Error("Maximum 4 images allowed per project.");

    // If it's already a direct link (from ImgBB), just save it!
    if (typeof imageData === "string" && imageData.startsWith("http")) {
      finalUrls.push(imageData);
    }
    // Otherwise, if it's still Base64, process it as a file in Drive
    else if (
      typeof imageData === "string" &&
      imageData.indexOf("data:image") === 0
    ) {
      const parts = imageData.split(",");
      const mime = parts[0].match(/:(.*?);/)[1];
      const bytes = Utilities.base64Decode(parts[1]);
      const fileName =
        "PAT_" +
        projectId +
        "_" +
        (finalUrls.length + 1) +
        "_" +
        new Date().getTime();
      const blob = Utilities.newBlob(bytes, mime, fileName);

      const file = folder.createFile(blob);
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );

      // Use the direct "uc" link which is more certain to load than the "lh3" thumbnail link
      const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
      finalUrls.push(url);

      // Also register this image as a SD_DOCUMENTS row (image -> doc)
      try {
        _appendImageDocRow(
          token,
          projectId,
          url,
          file.getId(),
          fileName,
          mime,
          bytes.length,
        );
      } catch (docErr) {
        console.warn("PAT single image doc-row: " + docErr.message);
      }
    } else {
      throw new Error("Invalid image data format.");
    }

    // Update the sheet with the new list of URLs
    const data = sh.getDataRange().getValues();
    const targetId = String(projectId || "")
      .toUpperCase()
      .trim();
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        sh.getRange(i + 1, c.IMAGES + 1).setValue(JSON.stringify(finalUrls));
        return { success: true, images: finalUrls };
      }
    }

    throw new Error("Project ID not found in records.");
  } catch (e) {
    console.error("Single Upload Error: " + e.message);
    return { success: false, message: e.message };
  }
}

function savePATImages(token, projectId, images) {
  try {
    const sess = _session(token);
    // Load the project to check permissions
    const projectRes = getPATProjectById(null, projectId);
    if (!projectRes.success)
      return { success: false, message: "Project not found." };
    const project = projectRes.project;

    if (!_canEditProject(sess, project)) {
      return {
        success: false,
        message:
          "Permission denied: You cannot upload images to this project at its current stage.",
      };
    }

    const folder = _getPATFolder();
    const sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };

    const c = _patCols();
    const data = sh.getDataRange().getValues();
    const targetId = String(projectId || "")
      .toUpperCase()
      .trim();

    // Find the existing row first
    let existingRowIdx = -1;
    let existingUrls = [];
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        existingRowIdx = i;
        break;
      }
    }
    if (existingRowIdx === -1)
      return { success: false, message: "Project not found." };

    // Current persisted images (the authoritative source of truth).
    try {
      existingUrls = _safeParse(data[existingRowIdx][c.IMAGES], []);
    } catch (e) {
      existingUrls = [];
    }
    if (!Array.isArray(existingUrls)) existingUrls = [];

    // Incoming: the desired full image set from the UI (URLs + any new base64
    // uploads). When the UI removes an image, it simply omits that URL here.
    const incoming =
      images && Array.isArray(images)
        ? images.filter((s) => typeof s === "string" && s.length > 0)
        : [];

    // ── SAFETY: if BOTH the DB and the incoming set are empty, there is
    // nothing to do — never wipe, never error. ──
    if (existingUrls.length === 0 && incoming.length === 0) {
      return { success: true, images: [] };
    }

    // ── GUARD against a spurious "clear everything" call: if the incoming set
    // is empty but the project ALREADY has images, an empty save is almost
    // always a UI race / unfinished load, NOT a deliberate wipe. Keep the
    // existing images. Deliberate removal goes through deletePATImages() /
    // the reconcile logic below (which only removes URLs the UI explicitly
    // dropped). This stops PATs from clearing "for no reason" on a re-save. ──
    if (incoming.length === 0 && existingUrls.length > 0) {
      return { success: true, images: existingUrls };
    }

    // ── RECONCILE ──
    // 1) Remove (trash from Drive + delete doc row) any persisted image the
    //    UI no longer includes. This is what makes "I delete from files and
    //    save" actually delete, and stops removed images coming back.
    const incomingSet = {};
    incoming.forEach((s) => {
      incomingSet[s] = true;
    });
    existingUrls.forEach((url) => {
      if (!incomingSet[url]) {
        try {
          _deleteOnePATDriveImage(url, projectId);
        } catch (e) {
          console.warn("reconcile trash: " + e.message);
        }
      }
    });

    // 2) Start from the incoming desired set, then upload any new base64.
    let finalUrls = incoming.slice();

    // Process incoming images (uploads are base64; existing links are kept).
    for (let i = 0; i < finalUrls.length; i++) {
      if (finalUrls.length >= 4) break; // Hard limit of 4 images
      const imgStr = finalUrls[i];
      if (imgStr.indexOf("data:image") === 0) {
        // New Upload: Convert base64 to Blob and save to Drive
        const parts = imgStr.split(",");
        const mime = parts[0].match(/:(.*?);/)[1];
        const bytes = Utilities.base64Decode(parts[1]);
        const blob = Utilities.newBlob(
          bytes,
          mime,
          "PAT_" + projectId + "_" + finalUrls.length,
        );
        const file = folder.createFile(blob);
        const url = "https://lh3.googleusercontent.com/d/" + file.getId();
        finalUrls[i] = url; // replace the base64 with the stored URL

        // Also register this image as a SD_DOCUMENTS row (image -> doc)
        try {
          _appendImageDocRow(
            token,
            projectId,
            url,
            file.getId(),
            "PAT_" + projectId + "_" + finalUrls.length,
            mime,
            bytes.length,
          );
        } catch (docErr) {
          console.warn("PAT bulk image doc-row: " + docErr.message);
        }
      }
      // Existing link: keep as is (already a URL)
    }

    // De-duplicate (defensive)
    const seen = {};
    finalUrls = finalUrls.filter((u) => {
      if (seen[u]) return false;
      seen[u] = true;
      return true;
    });
    if (finalUrls.length > 4) finalUrls = finalUrls.slice(0, 4);

    sh.getRange(existingRowIdx + 1, c.IMAGES + 1).setValue(
      JSON.stringify(finalUrls),
    );
    return { success: true, images: finalUrls };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deletePATImages(token, projectId) {
  try {
    const sess = _session(token);
    const projectRes = getPATProjectById(null, projectId);
    if (!projectRes.success)
      return { success: false, message: "Project not found." };
    const project = projectRes.project;

    if (!_canEditProject(sess, project)) {
      return {
        success: false,
        message:
          "Permission denied: You cannot delete images from this project.",
      };
    }

    const sh = _patSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No projects found." };
    const data = sh.getDataRange().getValues();
    const targetId = String(projectId || "")
      .toUpperCase()
      .trim();
    const c = _patCols();
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        // EXPLICIT delete — trash every Drive file + matching doc row, then
        // clear the cell. Called only when the user deliberately removed all
        // images (or individually deleted one via the UI -> save reconcile).
        try {
          const cur = _safeParse(data[i][c.IMAGES], []);
          if (Array.isArray(cur))
            cur.forEach((url) => {
              try {
                _deleteOnePATDriveImage(url, projectId);
              } catch (e) {}
            });
        } catch (e) {
          console.warn("deletePATImages cleanup: " + e.message);
        }
        sh.getRange(i + 1, c.IMAGES + 1).setValue("[]");
        return { success: true };
      }
    }
    return { success: false, message: "Project not found." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete all PAT projects (super admin only).
 * Also cascade-deletes ALL JCC certificates since every JCC is linked to a PAT project.
 */
function deleteAllPATProjects(token) {
  try {
    _superAdminSession(token);
    var sh = _patSheet();
    if (!sh) return { success: false, message: "PAT sheet not found." };
    var lastRow = sh.getLastRow();

    // Cascade-delete ALL JCC certificates (they all reference PAT projects)
    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    var jccDeleted = 0;
    if (jccSh && jccSh.getLastRow() > 1) {
      jccDeleted = jccSh.getLastRow() - 1;
      jccSh.deleteRows(2, jccSh.getLastRow() - 1);
    }

    if (lastRow > 1) {
      sh.deleteRows(2, lastRow - 1);
    }
    var msg = "All PAT projects deleted.";
    if (jccDeleted > 0) {
      msg += " " + jccDeleted + " associated JCC certificate(s) also removed.";
    }
    return { success: true, message: msg };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// DOCUMENT MANAGEMENT — Upload & store any file type
// Free storage via Google Drive. No API key needed.
// All departments can upload. No per-project limit.
// ─────────────────────────────────────────────────────────────

/**
 * Upload a document (any file type) to the portal.
 * Stores file in Google Drive, metadata in SD_DOCUMENTS sheet.
 * @param {string} token - Session token
 * @param {string} projectId - Project ID (can be empty for general docs)
 * @param {string} fileData - Base64 data URI or direct URL
 * @param {string} fileName - Original file name
 * @param {string} category - Category label (e.g. "Receipt", "Report", "Contract")
 * @return {Object} - { success, docId, driveUrl, message }
 */
function uploadDocument(token, projectId, fileData, fileName, category) {
  try {
    const sess = _session(token);
    const folder = _getDocumentsFolder();
    const sh = _docSheet();
    const c = _docCols();

    if (!sh) return { success: false, message: "Documents sheet not found." };

    let driveUrl = "";
    let driveFileId = "";
    let fileType = "";
    let fileSize = 0;

    if (typeof fileData === "string" && fileData.startsWith("http")) {
      driveUrl = fileData;
      fileType = "link";
      fileSize = 0;
    } else if (
      typeof fileData === "string" &&
      fileData.indexOf("data:") === 0
    ) {
      const parts = fileData.split(",");
      const mimeMatch = parts[0].match(/:(.*?);/);
      if (!mimeMatch) {
        return {
          success: false,
          message:
            "Could not detect file type from upload data. Try a different file.",
        };
      }
      const mime = mimeMatch[1];
      const raw = parts[1];
      if (!raw || raw.length === 0) {
        return {
          success: false,
          message: "Upload data is empty. Please select a valid file.",
        };
      }
      let bytes;
      try {
        bytes = Utilities.base64Decode(raw);
      } catch (e) {
        return {
          success: false,
          message:
            "File encoding error. The file may be corrupted or too large.",
        };
      }
      fileSize = bytes.length;
      fileType = mime;

      const ext =
        _mimeToExt(mime) || (fileName || "").split(".").pop() || "bin";
      const safeName = (fileName || "document_" + _genDocId()).replace(
        /[^a-zA-Z0-9._\- ]/g,
        "_",
      );
      const blobName =
        safeName.indexOf(".") === -1 ? safeName + "." + ext : safeName;
      const blob = Utilities.newBlob(bytes, mime, blobName);

      const file = folder.createFile(blob);
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW,
      );
      driveFileId = file.getId();
      driveUrl = "https://drive.google.com/uc?export=view&id=" + driveFileId;
    } else {
      return {
        success: false,
        message: "Unsupported file data format. Please try again.",
      };
    }

    const docId = _genDocId();
    const now = new Date().toISOString();
    const row = [
      docId,
      String(projectId || "").trim(),
      fileName || "Untitled",
      fileType,
      fileSize,
      driveUrl,
      driveFileId,
      sess.name || "",
      sess.email || "",
      sess.department || "",
      now,
      category || "General",
    ];

    sh.appendRow(row);
    return {
      success: true,
      docId: docId,
      driveUrl: driveUrl,
      message: "Document uploaded successfully.",
    };
  } catch (e) {
    return { success: false, message: "Upload failed: " + e.message };
  }
}

function _mimeToExt(mime) {
  const map = {
    "application/pdf": "pdf",
    "application/msword": "doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
      "docx",
    "application/vnd.ms-excel": "xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "text/plain": "txt",
    "text/csv": "csv",
    "application/zip": "zip",
    "application/x-rar-compressed": "rar",
    "application/vnd.rar": "rar",
    "application/vnd.ms-powerpoint": "ppt",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation":
      "pptx",
    "image/jpeg": "jpg",
    "image/png": "png",
    "image/gif": "gif",
    "image/webp": "webp",
    "image/svg+xml": "svg",
    "text/html": "html",
    "application/json": "json",
    "application/xml": "xml",
  };
  return map[mime] || null;
}

/**
 * Registers a PAT project image as a row in SD_DOCUMENTS so the image
 * also shows up under "Site Documents" and is tracked centrally.
 * Deduplicates by Drive file ID so re-saving the same images never
 * creates duplicate rows.
 * @param {string} token - session token (for uploader identity)
 * @param {string} projectId - owning PAT project id
 * @param {string} driveUrl - the Drive URL returned by the image upload
 * @param {string} driveFileId - the Drive file id
 * @param {string} fileName - display name
 * @param {string} mime - mime type of the image
 * @param {number} fileSize - byte size of the image
 */
function _appendImageDocRow(
  token,
  projectId,
  driveUrl,
  driveFileId,
  fileName,
  mime,
  fileSize,
) {
  try {
    if (!driveFileId) return; // Only Drive-stored images get a doc row (skip raw links)
    const sh = _docSheet();
    if (!sh) return;
    const c = _docCols();

    // Dedupe: skip if a doc row already exists for this Drive file id
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][c.DRIVE_FILE_ID] || "") === String(driveFileId))
        return;
    }

    let sess = null;
    try {
      sess = _session(token);
    } catch (e) {
      sess = null;
    }

    const docId = _genDocId();
    const now = new Date().toISOString();
    const row = [
      docId,
      String(projectId || "").trim(),
      fileName || "PAT_Image_" + projectId,
      mime || "image/jpeg",
      fileSize || 0,
      driveUrl || "",
      driveFileId,
      (sess && sess.name) || "",
      (sess && sess.email) || "",
      (sess && sess.department) || "",
      now,
      "PAT Image",
    ];
    sh.appendRow(row);
  } catch (e) {
    // Never let doc-row bookkeeping break the image upload itself.
    console.warn("_appendImageDocRow skipped: " + e.message);
  }
}

/**
 * Get documents for a project or all documents for a department.
 * @param {string} token - Session token
 * @param {string} projectId - Optional project ID filter
 * @param {string} department - Optional department filter
 * @return {Object} - { success, documents: [...] }
 */
function getDocuments(token, projectId, department) {
  try {
    _session(token);
    const sh = _docSheet();
    if (!sh || sh.getLastRow() < 2) return { success: true, documents: [] };

    const data = sh.getDataRange().getValues();
    const docs = [];
    const sess = _session(token);

    for (let i = 1; i < data.length; i++) {
      const doc = _documentFromRow(data[i]);

      // Filter by project if specified — case-insensitive, trimmed
      if (projectId) {
        const docPid = String(doc.projectId || "")
          .trim()
          .toUpperCase();
        const queryPid = String(projectId).trim().toUpperCase();
        if (docPid !== queryPid) continue;
      }

      // Filter by department if specified
      if (department && doc.uploadedByDept !== department) continue;

      // When viewing documents scoped to a project, show all docs for that project
      // (project access controls visibility instead)
      if (!projectId) {
        const role = String(sess.role || "").toLowerCase();
        if (role !== "super admin") {
          if (department && doc.uploadedByDept !== department) continue;
          if (
            !department &&
            doc.uploadedByEmail !== sess.email &&
            doc.uploadedByDept !== sess.department
          )
            continue;
        }
      }

      docs.push(doc);
    }

    docs.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
    return { success: true, documents: docs };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete a document by ID.
 * @param {string} token - Session token
 * @param {string} docId - Document ID to delete
 * @return {Object} - { success, message }
 */
function deleteDocument(token, docId) {
  try {
    const sess = _session(token);
    const sh = _docSheet();
    if (!sh || sh.getLastRow() < 2)
      return { success: false, message: "No documents found." };

    const data = sh.getDataRange().getValues();
    const c = _docCols();
    const role = String(sess.role || "").toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][c.DOC_ID]).toUpperCase().trim();
      if (rowId === String(docId).toUpperCase().trim()) {
        const doc = _documentFromRow(data[i]);

        // Check permission: uploader, same department admin, or super admin
        const canDelete =
          role === "super admin" ||
          doc.uploadedByEmail === sess.email ||
          (role === "admin" && doc.uploadedByDept === sess.department);

        if (!canDelete) {
          return {
            success: false,
            message: "Permission denied: You cannot delete this document.",
          };
        }

        // Try to delete from Drive
        if (doc.driveFileId) {
          try {
            const file = DriveApp.getFileById(doc.driveFileId);
            file.setTrashed(true);
          } catch (e) {
            console.warn("Could not delete Drive file: " + e.message);
          }
        }

        sh.deleteRow(i + 1);
        return { success: true, message: "Document deleted." };
      }
    }

    return { success: false, message: "Document not found." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete all documents for a project (cascade delete).
 * @param {string} token - Session token (must be admin or super admin)
 * @param {string} projectId - Project ID
 * @return {Object} - { success, message, deletedCount }
 */
function deleteProjectDocuments(token, projectId) {
  try {
    const sess = _session(token);
    const sh = _docSheet();
    if (!sh || sh.getLastRow() < 2)
      return {
        success: true,
        message: "No documents to delete.",
        deletedCount: 0,
      };

    const data = sh.getDataRange().getValues();
    const c = _docCols();
    const role = String(sess.role || "").toLowerCase();

    if (role !== "admin" && role !== "super admin") {
      return { success: false, message: "Admin access required." };
    }

    let deletedCount = 0;
    const rowsToDelete = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const docProjectId = String(data[i][c.PROJECT_ID] || "")
        .toUpperCase()
        .trim();
      if (docProjectId === String(projectId).toUpperCase().trim()) {
        const doc = _documentFromRow(data[i]);

        // Check permission
        const canDelete =
          role === "super admin" ||
          doc.uploadedByEmail === sess.email ||
          doc.uploadedByDept === sess.department;

        if (canDelete) {
          if (doc.driveFileId) {
            try {
              const file = DriveApp.getFileById(doc.driveFileId);
              file.setTrashed(true);
            } catch (e) {
              console.warn("Drive delete failed: " + e.message);
            }
          }
          rowsToDelete.push(i + 1);
          deletedCount++;
        }
      }
    }

    // Delete rows from bottom to top to preserve row indices
    rowsToDelete
      .sort((a, b) => b - a)
      .forEach(function (rowIdx) {
        sh.deleteRow(rowIdx);
      });

    return {
      success: true,
      message: deletedCount + " document(s) deleted.",
      deletedCount: deletedCount,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Gets the current status and verdict of a PAT project
 * @param {string} token - User session token
 * @param {string} projectId - PAT project ID
 * @return {Object} - {success: boolean, status: string, verdict: string, comments: string, assignedToDept: string}
 */
function getProjectStatus(token, projectId) {
  try {
    _session(token);
    var result = getPATProjectById(token, projectId);
    if (!result.success) return result;

    var project = result.project;
    return {
      success: true,
      status: project.workflowStatus,
      verdict: project.verdict,
      comments: project.rejectionReason || project.comments || "",
      assignedToDept: project.assignedToDept,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Get department names for the signup dropdown.
 * Public — no token needed.
 * @returns {{ success, data: string[] }}
 */
function getPublicDepartments() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("public_departments");
  if (cached) {
    try {
      return { success: true, data: JSON.parse(cached), cached: true };
    } catch (e) {
      cache.remove("public_departments");
    }
  }

  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    var data = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    var names = data
      .map(function (r) {
        return String(r[0]).trim();
      })
      .filter(Boolean)
      .sort();
    // The Audit department is NOT a self-signup option. Accounts in Audit can
    // only be created by a Super Admin via the admin console, so it is hidden
    // from the public registration dropdown.
    names = names.filter(function (n) {
      return String(n).toLowerCase() !== "audit";
    });

    cache.put("public_departments", JSON.stringify(names), 600); // Cache for 10 mins (better UX)
    return { success: true, data: names };
  } catch (e) {
    return { success: false, data: [], message: e.message };
  }
}

/**
 * Get the allowed email domain from configuration.
 * Public — no token needed.
 * @returns {{ success, domain: string }}
 */
function getDomainConfig() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("domain_config");
  if (cached) {
    try {
      return { success: true, ...JSON.parse(cached), cached: true };
    } catch (e) {
      cache.remove("domain_config");
    }
  }

  try {
    var config = _getConfig();
    var props = PropertiesService.getScriptProperties();
    var allowLocalhost = props.getProperty("allow_localhost") === "true";
    var data = {
      domain: config.DOMAIN,
      aiWelcomeMessage: config.AI_WELCOME_MESSAGE,
      allowLocalhost: allowLocalhost,
    };
    cache.put("domain_config", JSON.stringify(data), 1500);
    return { success: true, ...data };
  } catch (e) {
    return { success: false, domain: "@fob.ng", message: e.message };
  }
}

/**
 * Toggle localhost access. Super Admin only.
 */
function setLocalhostAccess(token, allowed) {
  try {
    _superAdminSession(token);

    var current = _loadAllowedOrigins();

    var filtered = current.filter(function (origin) {
      return origin !== "http://localhost" && origin !== "http://127.0.0.1";
    });

    if (allowed === true) {
      filtered.push("http://localhost");
      filtered.push("http://127.0.0.1");
    }

    PropertiesService.getScriptProperties().setProperty(
      "allowed_origins",
      JSON.stringify(filtered),
    );
    CacheService.getScriptCache().remove("allowed_origins");

    PropertiesService.getScriptProperties().setProperty(
      "allow_localhost",
      allowed === true ? "true" : "false",
    );
    CacheService.getScriptCache().remove("domain_config");

    return {
      success: true,
      message: "Localhost access " + (allowed ? "enabled" : "disabled") + ".",
      allowedOrigins: filtered,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
//  ADMIN  —  all require a valid admin token (first argument)
// ══════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────

// ── ALL REQUESTS + STATS ──────────────────────────────────────

/**
 * Returns all requests and running stats.
 * Called by admin.html "All Requests" tab on load.
 * @param {string} token   admin token
 * @param {Object} filters { status?: string, department?: string }
 * @returns {{ success, requests, stats: {total,newCount,inProgress,done} }}
 */

// ── DEPARTMENTS ───────────────────────────────────────────────

/**
 * List departments with live request + user counts.
 * @param {string} token  admin token
 */
function getDepartments(token) {
  try {
    _adminSession(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var deptSh = ss.getSheetByName(SD.DEPTS);
    var userSh = ss.getSheetByName(SD.USERS);

    var deptData =
      deptSh && deptSh.getLastRow() > 1
        ? deptSh.getDataRange().getValues().slice(1)
        : [];
    var userData =
      userSh && userSh.getLastRow() > 1
        ? userSh.getDataRange().getValues().slice(1)
        : [];

    // Count active users per department name
    var userMap = {};
    userData.forEach(function (r) {
      if (r[7] === "active") {
        var d = String(r[6]).trim();
        userMap[d] = (userMap[d] || 0) + 1;
      }
    });

    var departments = deptData.map(function (r) {
      return {
        deptId: r[0],
        name: r[1],
        headEmail: r[2],
        createdAt: r[3],
        activeUsers: userMap[r[1]] || 0,
      };
    });
    return { success: true, departments: departments };
  } catch (e) {
    return { success: false, message: e.message, departments: [] };
  }
}

/**
 * Add a department.
 * @param {string} token
 * @param {string} name
 * @param {string} headEmail  (optional)
 */
function addDepartment(token, name, headEmail) {
  try {
    var sess = _adminSession(token);
    name = String(name || "").trim();
    if (!name) throw new Error("Department name is required.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues().slice(1);
    var dup = data.find(function (r) {
      return String(r[1]).toLowerCase() === name.toLowerCase();
    });
    if (dup)
      throw new Error(
        "A department with the name '" + name + "' already exists.",
      );

    var deptId = _genId("DEPT");
    sh.appendRow([
      deptId,
      name,
      headEmail || "",
      new Date().toISOString(),
      sess.email,
    ]);
    CacheService.getScriptCache().remove("public_departments"); // Invalidate cache for signup
    return { success: true, deptId: deptId, message: "Department added." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Update a department's name and/or head email.
 * @param {string} token
 * @param {string} deptId
 * @param {string} name
 * @param {string} headEmail
 */
function updateDepartment(token, deptId, name, headEmail) {
  try {
    _adminSession(token);
    name = String(name || "").trim();
    if (!name) throw new Error("Department name is required.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sh.getRange(i + 1, 2).setValue(name);
        sh.getRange(i + 1, 3).setValue(headEmail || "");
        CacheService.getScriptCache().remove("public_departments");
        return { success: true, message: "Department updated." };
      }
    }
    throw new Error("Department not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Delete a department.
 * @param {string} token
 * @param {string} deptId
 */
function deleteDepartment(token, deptId) {
  try {
    _adminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sh.deleteRow(i + 1);
        CacheService.getScriptCache().remove("public_departments");
        return { success: true, message: "Department deleted." };
      }
    }
    throw new Error("Department not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ── USERS ─────────────────────────────────────────────────────

/**
 * List all registered users.
 * @param {string} token  admin token
 */
function getUsers(token) {
  try {
    _adminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues().slice(1);
    var users = data.map(function (r) {
      var ipsRaw = r[14] || "";
      var ips = [];
      try {
        ips = ipsRaw ? JSON.parse(ipsRaw) : [];
      } catch (e) {
        ips = [];
      }
      return {
        userId: r[0],
        name: r[1],
        email: r[2],
        role: r[5],
        department: r[6],
        gender: r[7],
        workflowNotes: r[8],
        status: r[9],
        createdAt: r[10],
        lastLoginAt: r[11],
        loginIps: ips,
      };
    });
    return { success: true, users: users };
  } catch (e) {
    return { success: false, message: e.message, users: [] };
  }
}

function searchUsers(token, query) {
  try {
    _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues().slice(1);
    var q = String(query || "")
      .toLowerCase()
      .trim();
    var users = data
      .map(function (r) {
        return {
          userId: r[0],
          name: r[1],
          email: r[2],
          role: r[5],
          department: r[6],
          status: r[9],
        };
      })
      .filter(function (u) {
        if (!q) return true;
        return (
          (u.name || "").toLowerCase().indexOf(q) !== -1 ||
          (u.email || "").toLowerCase().indexOf(q) !== -1
        );
      })
      .filter(function (u) {
        // Banned users must never appear as mail contacts / in search.
        return String(u.status || "").toLowerCase() !== "banned";
      });
    return { success: true, users: users };
  } catch (e) {
    return { success: false, message: e.message, users: [] };
  }
}

function getUserById(token, userId) {
  try {
    _adminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var searchId = String(userId || "").trim();

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      if (rowId === searchId) {
        var ipsRaw = data[i][14] || ""; // LoginIPs (col 15 / index 14) — JSON array
        var pwHash = data[i][13] || "";
        var fingerprintsRaw = data[i][15] || ""; // Fingerprints (col 16 / index 15) — JSON array
        var ips = [];
        var fingerprints = [];
        try {
          ips = ipsRaw ? JSON.parse(ipsRaw) : [];
        } catch (e) {
          ips = [];
        }
        try {
          fingerprints = fingerprintsRaw ? JSON.parse(fingerprintsRaw) : [];
        } catch (e) {
          fingerprints = [];
        }
        return {
          success: true,
          userId: data[i][0],
          name: data[i][1],
          email: data[i][2],
          role: data[i][5],
          department: data[i][6],
          status: data[i][9],
          banReason: String(data[i][16] || ""),
          bannedBy: String(data[i][17] || ""),
          ips: ips,
          fingerprints: fingerprints,
        };
      }
    }
    throw new Error("User not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Promote or demote a user's role.
 * @param {string} token
 * @param {string} userId
 * @param {string} newRole  "admin" | "employee"
 */
function updateUserRole(token, userId, newRole) {
  try {
    var sess = _superAdminSession(token);
    if (
      newRole !== SD.ADMIN &&
      newRole !== SD.EMPLOYEE &&
      newRole !== "super admin" &&
      newRole !== "audit"
    )
      throw new Error(
        "Role must be 'admin', 'employee', 'audit', or 'super admin'.",
      );
    // Only super admin can grant the 'admin' role — no other path can mint admins.
    if (newRole === SD.ADMIN) {
      // already gated by _superAdminSession above; explicit note for clarity
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var searchId = String(userId || "").trim();
    var isEmailSearch = searchId.indexOf("@") !== -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      var rowEmail = String(data[i][2]).toLowerCase().trim();
      if (
        rowId === searchId ||
        (isEmailSearch && rowEmail === searchId.toLowerCase())
      ) {
        var currentRole = String(data[i][5] || "").toLowerCase();
        // ── SELF-ACCOUNT PROTECTION ──
        // Super admin role cannot be removed/changed by anyone.
        if (currentRole === "super admin" && newRole !== "super admin") {
          return {
            success: false,
            message: "Super admin role is protected and cannot be changed.",
          };
        }
        // When promoting to audit, align the department so the console recognises it.
        if (newRole === "audit") {
          sh.getRange(i + 1, 7).setValue("Audit");
        }
        sh.getRange(i + 1, 6).setValue(newRole);
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Activate or ban a user account.
 * @param {string} token
 * @param {string} userId
 * @param {string} status  "active" | "banned"
 */
function updateUserStatus(token, userId, status, reason, bannedBy) {
  // NOTE: Ban metadata lives in USERS columns 16 (BanReason) and 17 (BannedBy),
  // immediately after the 15 declared headers (1-based).

  var lock = LockService.getScriptLock();
  try {
    var sess = _superAdminSession(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        // ── SELF-ACCOUNT PROTECTION ──
        // A super admin can NEVER be banned, by anyone (including another super admin).
        var targetRole = String(data[i][5] || "").toLowerCase();
        if (
          targetRole === "super admin" &&
          String(status).toLowerCase() === "banned"
        ) {
          lock.releaseLock();
          return {
            success: false,
            message: "Super admin accounts are protected and cannot be banned.",
          };
        }
        sh.getRange(i + 1, 10).setValue(status);
        // Store ban reason (col 17 / index 16) and who banned (col 18 / index 17).
        // NOTE: cols 15 (LoginIPs) and 16 (Fingerprints) are owned by _recordLoginIp.
        if (String(status).toLowerCase() === "banned") {
          sh.getRange(i + 1, 17).setValue(
            String(reason || "Account suspended by administrator."),
          );
          sh.getRange(i + 1, 18).setValue(
            String(bannedBy || sess.email || "Administrator"),
          );
          // ── FORCE LOGOUT ── kill every live session for this user immediately
          try {
            _purgeUserSessions(String(data[i][2] || ""));
          } catch (e) {}
        } else {
          // Clear ban metadata on reactivation
          sh.getRange(i + 1, 17).setValue("");
          sh.getRange(i + 1, 18).setValue("");
          // ── SYNC BLACKLIST ──
          // A ban creates up to THREE blacklist rows per user: one per known IP,
          // one per fingerprint, and a dedicated "EMAIL:<addr>" row — all sharing
          // the user's email in col 3 (USER_EMAIL). If we only clear the USERS
          // status here (which is what the Users-tab "Activate" button used to do),
          // the IP/fingerprint blacklist rows survive, and login is STILL blocked
          // by _checkBlacklistMatch even though the account is "active". That is
          // the "cleared the blacklist but still banned" trap.
          //
          // So on reactivation we also stamp UNBANNED_AT on every active blacklist
          // row tied to this email. This makes the Users-tab Activate button fully
          // equivalent to unbanUserWithBlacklist, with no orphaned block rows.
          try {
            var _bsh = _blacklistSheet();
            if (_bsh && _bsh.getLastRow() >= 2) {
              var _bc = _blacklistCols();
              var _bdata = _bsh.getDataRange().getValues();
              var _uemail = String(data[i][2] || "").toLowerCase();
              var _now = new Date().toISOString();
              for (var _b = 1; _b < _bdata.length; _b++) {
                if (_bdata[_b][_bc.UNBANNED_AT]) continue; // already inactive
                if (
                  String(_bdata[_b][_bc.USER_EMAIL] || "").toLowerCase() ===
                  _uemail
                ) {
                  _bsh.getRange(_b + 1, _bc.UNBANNED_AT + 1).setValue(_now);
                }
              }
            }
          } catch (e) {
            /* never block reactivation on a blacklist lookup hiccup */
          }
        }
        lock.releaseLock();
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, message: e.message };
  }
}

// ── ALLOWLIST ─────────────────────────────────────────────────

/**
 * Get all allowlisted admins.
 * @param {string} token  admin token
 */

// ── PROFILE SETTINGS (admin + employee) ──────────────────────

/**
 * Update the logged-in user's display name.
 * Token identifies who is updating — email must match session.
 *
 * @param {string} token
 * @param {string} email    the user's own email (must match session)
 * @param {string} newName  new display name
 */
function updateUserProfile(token, email, newName) {
  try {
    var sess = _session(token);
    if (sess.email !== String(email).toLowerCase().trim())
      throw new Error("You can only update your own profile.");
    newName = String(newName || "").trim();
    if (!newName) throw new Error("Name cannot be empty.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][2] || "")
          .toLowerCase()
          .trim() ===
        String(sess.email || "")
          .toLowerCase()
          .trim()
      ) {
        sh.getRange(i + 1, 2).setValue(newName);
        // Update the stored session payload so the nav reflects the change immediately
        var raw = PropertiesService.getUserProperties().getProperty(
          "tok_" + token,
        );
        if (raw) {
          try {
            var s = JSON.parse(raw);
            s.name = newName;
            PropertiesService.getUserProperties().setProperty(
              "tok_" + token,
              JSON.stringify(s),
            );
          } catch (parseErr) {
            // Session data is corrupted, but profile was updated anyway
            console.warn(
              "Session parse error during profile update:",
              parseErr,
            );
          }
        }

        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════
//  EMPLOYEE  —  called from employee.html via google.script.run
// ══════════════════════════════════════════════════════════════
// ─────────────────────────────────────────────────────────────

/**
 * Server-side AI Proxy.
 * Centralizes AI calls to ensure stability and security.
 * @param {string} token
 * @param {string} prompt
 */
/**
 * Records an internal message between employees/admins.
 */
function updateMessagingSettings(token, allowMsgs, aiAuto) {
  try {
    var sess = _session(token);
    var sessEmail = String(sess.email || "")
      .toLowerCase()
      .trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][2] || "")
          .toLowerCase()
          .trim() === sessEmail
      ) {
        sh.getRange(i + 1, 13).setValue(allowMsgs ? "TRUE" : "FALSE");
        sh.getRange(i + 1, 14).setValue(aiAuto ? "TRUE" : "FALSE");
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Gets members of the same department (for employees) or all members (for admins).
 */
function getDepartmentMembers(token) {
  try {
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, members: [] };

    var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    var role = String(sess.role || "").toLowerCase();
    // Only Super Admin sees the entire org. Everyone else (incl. admins)
    // sees ONLY the people in their own department.
    var members = data
      .filter(function (r) {
        if (role === "super admin") return true;
        return String(r[6]) === sess.department;
      })
      .filter(function (r) {
        // Banned members must not surface in team/department views.
        return String(r[9] || "").toLowerCase() !== "banned";
      })
      .map(function (r) {
        return {
          name: r[1],
          email: r[2],
          role: String(r[5] || "").toLowerCase(),
        };
      });

    return { success: true, members: members };
  } catch (e) {
    return { success: false, message: e.message, members: [] };
  }
}

/**
 * Server-side Gemini API proxy.
 * Proxies AI chat requests to Google Gemini, keeping the API key hidden.
 * @param {string} token - user session token
 * @param {string} systemPrompt - the system prompt
 * @param {string} userMessage - the user message
 * @param {Array} history - recent chat history (optional)
 * @returns {{ success, answer }}
 */
function callGemini(token, systemPrompt, userMessage, history) {
  try {
    var sess = _session(token);
    var config = _getConfig();
    var apiKey = config.GEMINI_API_KEY;
    if (!apiKey)
      throw new Error(
        "Gemini API key not configured. Please set it in System Configuration.",
      );

    // Build contents array for Gemini API
    var contents = [];

    // Add chat history
    if (history && Array.isArray(history)) {
      history.slice(-10).forEach(function (h) {
        if (typeof h.content === "string") {
          contents.push({
            role: h.role === "assistant" ? "model" : "user",
            parts: [{ text: h.content }],
          });
        }
      });
    }

    // Add current user message
    contents.push({ role: "user", parts: [{ text: userMessage }] });

    var payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 2048,
      },
    };

    var url =
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=" +
      apiKey;
    var response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      var errData = JSON.parse(response.getContentText());
      throw new Error(
        errData.error?.message ||
          "Gemini API error " + response.getResponseCode(),
      );
    }

    var data = JSON.parse(response.getContentText());
    var text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";

    // Check for PAT topic restriction violation
    if (text.indexOf("REJECTED_NON_PAT_QUERY:") !== -1) {
      _logRejectedQuery(
        sess,
        userMessage || "[Image uploaded]",
        text,
        "off-topic",
      );
      return {
        success: true,
        answer:
          "I can only answer questions related to PAT projects and the SD portal system. Please ask me something about PAT projects, fiber deployment, or the SD workflow system.",
      };
    }

    return { success: true, answer: text };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Server-side Gemini Vision API proxy.
 * Handles image attachments (base64) along with text prompts.
 * @param {string} token - user session token
 * @param {string} systemPrompt - the system prompt
 * @param {string} userMessage - the user message
 * @param {Array} imageAttachments - array of {type: 'image_url', image_url: {url: 'data:...'}} objects
 * @param {Array} history - recent chat history (optional)
 * @returns {{ success, answer }}
 */
function callGeminiVision(
  token,
  systemPrompt,
  userMessage,
  imageAttachments,
  history,
) {
  try {
    var sess = _session(token);
    var config = _getConfig();
    var apiKey = config.GEMINI_API_KEY;
    if (!apiKey)
      throw new Error(
        "Gemini API key not configured. Please set it in System Configuration.",
      );

    var contents = [];

    // Add chat history (text only)
    if (history && Array.isArray(history)) {
      history.slice(-8).forEach(function (h) {
        if (typeof h.content === "string") {
          contents.push({
            role: h.role === "assistant" ? "model" : "user",
            parts: [{ text: h.content }],
          });
        }
      });
    }

    // Build current user message with images
    var parts = [];
    if (userMessage) parts.push({ text: userMessage });

    // Process image attachments
    if (imageAttachments && Array.isArray(imageAttachments)) {
      imageAttachments.forEach(function (att) {
        if (att.type === "image_url" && att.image_url && att.image_url.url) {
          var dataUrl = att.image_url.url;
          var mimeMatch = dataUrl.match(/data:(image\/[a-z]+);base64,/);
          var mime = mimeMatch ? mimeMatch[1] : "image/jpeg";
          var base64Data = dataUrl.replace(/^data:image\/[a-z]+;base64,/, "");
          parts.push({
            inlineData: {
              mimeType: mime,
              data: base64Data,
            },
          });
        }
      });
    }

    if (parts.length === 0)
      parts.push({ text: userMessage || "Analyze this image." });
    contents.push({ role: "user", parts: parts });

    // Inject PAT-ONLY restriction into system prompt
    var patRestriction =
      "\n\nCRITICAL TOPIC RESTRICTION: You are STRICTLY LIMITED to answering questions about PAT (Provisional Acceptability Test) projects, the SD portal system, fiber network deployment, and related workflow operations. If the user asks about ANY topic outside of PAT, fiber network deployment, or the SD system, you MUST respond with ONLY this exact prefix: 'REJECTED_NON_PAT_QUERY:' followed by your refusal. Do NOT answer off-topic questions under any circumstances. Stay locked to PAT domain only.";

    var payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: {
        temperature: 0.5,
        maxOutputTokens: 4096,
      },
    };

    var url =
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=" +
      apiKey;
    var response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
      var errData = JSON.parse(response.getContentText());
      throw new Error(
        errData.error?.message ||
          "Gemini Vision API error " + response.getResponseCode(),
      );
    }

    var data = JSON.parse(response.getContentText());
    var text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    return { success: true, answer: text };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAIResponse(token, prompt) {
  try {
    var sess = _session(token);
    var config = _getConfig();
    var patRestriction =
      " CRITICAL: You are STRICTLY LIMITED to PAT (Provisional Acceptability Test) projects only. Reject anything else.";
    var systemPrompt =
      "You are the SD-AI for FiberOne Broadband. " +
      config.AI_WELCOME_MESSAGE +
      patRestriction;

    // Using the stable Pollinations API via server-side fetch
    var url =
      "https://text.pollinations.ai/" +
      encodeURIComponent(prompt) +
      "?system=" +
      encodeURIComponent(systemPrompt) +
      "&model=openai";
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (response.getResponseCode() !== 200)
      throw new Error("AI service is currently busy.");
    return { success: true, answer: response.getContentText() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// JCC (Job Completion Certificate) AUTO-GENERATION
// ─────────────────────────────────────────────────────────────

/**
 * Auto-generates a JCC when a PAT project is completed.
 * Called from approvePAT() or when workflowStatus becomes 'Completed'.
 *
 * @param {string} projectId - The PAT project ID
 * @param {object} session - The user session object {name, email, department}
 * @returns {object} - {success, jccId, message}
 */
/**
 * Extracts State/Region from a Nigerian site address.
 * E.g. "135 Port Road, Apapa, Lagos" → "Lagos State"
 */
function _extractStateRegion(address) {
  if (!address) return "N/A";
  var states = [
    "Abia",
    "Adamawa",
    "Akwa Ibom",
    "Anambra",
    "Bauchi",
    "Bayelsa",
    "Benue",
    "Borno",
    "Cross River",
    "Delta",
    "Ebonyi",
    "Edo",
    "Ekiti",
    "Enugu",
    "FCT",
    "Gombe",
    "Imo",
    "Jigawa",
    "Kaduna",
    "Kano",
    "Katsina",
    "Kebbi",
    "Kogi",
    "Kwara",
    "Lagos",
    "Nassarawa",
    "Niger",
    "Ogun",
    "Ondo",
    "Osun",
    "Oyo",
    "Plateau",
    "Rivers",
    "Sokoto",
    "Taraba",
    "Yobe",
    "Zamfara",
  ];
  var addrLower = address.toLowerCase();
  for (var i = 0; i < states.length; i++) {
    if (addrLower.indexOf(states[i].toLowerCase()) !== -1) {
      // Return proper casing + " State" (except FCT)
      if (states[i] === "FCT") return "FCT, Abuja";
      return states[i] + " State";
    }
  }
  // Fallback: try to extract last part after comma
  var parts = address.split(",").map(function (p) {
    return p.trim();
  });
  if (parts.length >= 2) {
    return parts[parts.length - 1] + " State";
  }
  return "N/A";
}

function _buildJCCDataObject(project, session) {
  var config = _getConfig();
  var year = new Date().getFullYear();
  var certId = "JCC-" + year + "-" + _genId("").substring(4);
  var certType = "JOB COMPLETION CERTIFICATE (State %)";
  var vendor = project.vendor || "Unknown Vendor";
  var aiAnalysis = _generateJCCRemarks(project);

  // mecName must be STATIC — the MEC account that finally approved/completed this project.
  // Do NOT use session.name because that is whoever is currently viewing/requesting.
  var mecName = "MEC Officer";
  var history = project.workflowHistory || [];
  for (var i = history.length - 1; i >= 0; i--) {
    var entry = history[i];
    if (entry.toStatus === "Completed" && entry.by && entry.by.department) {
      var dept = String(entry.by.department).toLowerCase();
      if (dept.indexOf("mec") !== -1 || dept.indexOf("mech") !== -1) {
        mecName = entry.by.name || "MEC Officer";
        break;
      }
    }
  }
  if (mecName === "MEC Officer" && session && session.name) {
    mecName = session.name; // fallback only if no MEC completion history found
  }

  var mecDate = new Date().toISOString().split("T")[0];
  var mecSignature = "Approved ~~" + mecName + "~~";
  var vendorDate = mecDate;
  var vendorSignature = "Approved ~~" + vendor + "~~";
  var mecHeadName = config.MEC_HEAD_NAME || "MEC Head";
  var mecHeadDate = mecDate;
  var mecHeadSignature = "Approved ~~" + mecHeadName + "~~";
  var stateRegion = _extractStateRegion(project.siteAddress);
  var orchestrator = project.orchestrator || "N/A";

  return {
    jccId: "PREVIEW",
    projectId: project.projectId,
    projectName: project.projectName,
    projectNumber: "PREVIEW",
    presidingOfficer: project.presidingOfficer || "",
    certificateType: certType,
    vendor: vendor,
    certificateId: certId,
    penalty: aiAnalysis.penalty,
    remarks: aiAnalysis.remarks,
    stateRegion: stateRegion,
    orchestrator: orchestrator,
    mecName: mecName,
    mecSignature: mecSignature,
    mecDate: mecDate,
    vendorName: vendor,
    vendorSignature: vendorSignature,
    vendorDate: vendorDate,
    mecHeadName: mecHeadName,
    mecHeadSignature: mecHeadSignature,
    mecHeadDate: mecHeadDate,
    generatedAt: new Date().toISOString(),
    generatedBy: session ? session.name : "System",
    isPreview: true,
  };
}

function generateJCC(projectId, session) {
  try {
    var config = _getConfig();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.PAT);
    if (!sh) throw new Error("PAT sheet not found.");

    var data = sh.getDataRange().getValues();
    var c = _patCols();
    var project = null;
    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][c.PROJECT_ID]).trim();
      if (
        rowId.toUpperCase() === String(projectId).toUpperCase() ||
        rowId.toUpperCase() ===
          String(projectId)
            .replace(/^(FOB|PAT)-/i, "")
            .toUpperCase() ||
        ("PAT-" + rowId).toUpperCase() === String(projectId).toUpperCase()
      ) {
        project = _projectFromRow(data[i]);
        break;
      }
    }

    if (!project) throw new Error("Project not found: " + projectId);
    if (project.workflowStatus !== "Completed")
      throw new Error("Project is not completed yet.");

    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    if (!jccSh) throw new Error("JCC sheet not found. Run First Setup.");
    var jccData = jccSh.getDataRange().getValues();
    for (var j = 1; j < jccData.length; j++) {
      if (
        String(jccData[j][1]).trim().toUpperCase() ===
        String(projectId).toUpperCase()
      ) {
        return {
          success: false,
          message: "JCC already exists for this project.",
          jccId: jccData[j][0],
        };
      }
    }

    // PO No = row number in the PAT sheet (starting from 1 for data rows)
    // Find the project's row index in the PAT sheet
    var projectRowNumber = 0;
    var patData = sh.getDataRange().getValues();
    for (var k = 1; k < patData.length; k++) {
      var rowId = String(patData[k][c.PROJECT_ID]).trim();
      if (
        rowId.toUpperCase() === String(projectId).toUpperCase() ||
        rowId.toUpperCase() ===
          String(projectId)
            .replace(/^(FOB|PAT)-/i, "")
            .toUpperCase() ||
        ("PAT-" + rowId).toUpperCase() === String(projectId).toUpperCase()
      ) {
        projectRowNumber = k; // k is 0-based for data, but k=1 is first data row, so k itself is the 1-based row number
        break;
      }
    }
    var projectNumber = projectRowNumber > 0 ? projectRowNumber : 1;

    var jccObj = _buildJCCDataObject(project, session);
    jccObj.projectNumber = projectNumber;
    jccObj.isPreview = false;

    var jccId = _genJCCId();
    jccObj.jccId = jccId;

    jccSh.appendRow([
      jccId,
      jccObj.projectId,
      jccObj.projectName,
      jccObj.projectNumber,
      jccObj.certificateType,
      jccObj.vendor,
      jccObj.certificateId,
      jccObj.penalty,
      jccObj.remarks,
      jccObj.mecName,
      jccObj.mecSignature,
      jccObj.mecDate,
      jccObj.vendorName,
      jccObj.vendorSignature,
      jccObj.vendorDate,
      jccObj.mecHeadName,
      jccObj.mecHeadSignature,
      jccObj.mecHeadDate,
      jccObj.generatedAt,
      session.name,
      session.email,
      jccObj.stateRegion,
      jccObj.orchestrator,
      jccObj.presidingOfficer,
    ]);

    var lastRow = jccSh.getLastRow();
    jccSh
      .getRange(lastRow, 1, 1, jccSh.getLastColumn())
      .setFontFamily("Arial")
      .setFontSize(10);

    return {
      success: true,
      jccId: jccId,
      certificateId: jccObj.certificateId,
      projectNumber: projectNumber,
      message:
        "JCC generated successfully. Certificate ID: " + jccObj.certificateId,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function _genJCCId() {
  return "JCC-" + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function _generateJCCRemarks(project) {
  var score = parseInt(project.snagScore) || 0;
  var snags = project.snags || [];
  var history = project.workflowHistory || [];

  var penalty = "None";
  var remarks = "";

  var criticalSnags = snags.filter(function (s) {
    return parseInt(s.weight) >= 3;
  });
  var majorSnags = snags.filter(function (s) {
    return parseInt(s.weight) === 2;
  });
  var minorSnags = snags.filter(function (s) {
    return parseInt(s.weight) === 1;
  });

  if (score === 0) {
    penalty = "None - Project completed with zero complaints.";
  } else if (score <= 2) {
    penalty = "Minor - Project completed with minor observations noted.";
  } else if (score <= 5) {
    penalty =
      "Moderate - Project completed with major observations requiring follow-up.";
  } else {
    penalty =
      "Significant - Project completed with critical observations. Follow-up inspection required.";
  }

  remarks =
    "Project '" +
    (project.projectName || "Unknown") +
    "' has been reviewed and accepted. ";
  if (score === 0) {
    remarks += "All checklist items passed with no complaints. ";
  } else {
    remarks += "Total complaint score: " + score + ". ";
    if (criticalSnags.length > 0)
      remarks += criticalSnags.length + " critical item(s) identified. ";
    if (majorSnags.length > 0)
      remarks += majorSnags.length + " major item(s) identified. ";
    if (minorSnags.length > 0)
      remarks += minorSnags.length + " minor item(s) identified. ";
  }

  var deptCount = {};
  history.forEach(function (h) {
    var dept = h.by && h.by.department ? h.by.department : "Unknown";
    deptCount[dept] = (deptCount[dept] || 0) + 1;
  });

  remarks +=
    "Workflow involved: " +
    (Object.keys(deptCount).join(", ") || "MEC review") +
    ". ";
  remarks += "Final verdict: " + (project.verdict || "Pending") + ". ";
  remarks +=
    "Project cleared for completion on " +
    new Date().toLocaleDateString("en-GB") +
    ".";

  return { penalty: penalty, remarks: remarks };
}

function getJCCByProjectId(token, projectId) {
  try {
    var sess = _session(token);
    var searchId = String(projectId || "")
      .trim()
      .toUpperCase();

    if (!searchId || searchId === "UNDEFINED" || searchId === "NULL") {
      return {
        success: false,
        message: "Project ID is missing. Please try again.",
      };
    }

    // 1. Try finding an official JCC record first
    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    if (jccSh && jccSh.getLastRow() >= 2) {
      var data = jccSh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowId = String(data[i][1] || "")
          .trim()
          .toUpperCase();
        var rowNum = String(data[i][3] || "")
          .trim()
          .toUpperCase();

        if (
          rowId === searchId ||
          rowId === searchId.replace(/^(FOB|PAT)-/, "") ||
          "PAT-" + rowId === searchId ||
          rowNum === searchId ||
          rowNum === searchId.replace(/^(FOB|PAT)-/, "") ||
          "PAT-" + rowNum === searchId
        ) {
          return {
            success: true,
            jcc: {
              jccId: data[i][0],
              projectId: data[i][1],
              projectName: data[i][2],
              projectNumber: data[i][3],
              certificateType: data[i][4],
              vendor: data[i][5],
              certificateId: data[i][6],
              penalty: data[i][7],
              remarks: data[i][8],
              mecName: data[i][9],
              mecSignature: data[i][10],
              mecDate: data[i][11],
              vendorName: data[i][12],
              vendorSignature: data[i][13],
              vendorDate: data[i][14],
              mecHeadName: data[i][15],
              mecHeadSignature: data[i][16],
              mecHeadDate: data[i][17],
              generatedAt: data[i][18],
              generatedBy: data[i][19],
              stateRegion: data[i][21] || _extractStateRegion(""),
              orchestrator: data[i][22] || "N/A",
              presidingOfficer: data[i][23] || "",
              isPreview: false,
            },
          };
        }
      }
    }

    // 2. If no official JCC, try to generate a preview from PAT project data
    var patRes = getPATProjectById(null, projectId);
    if (patRes && patRes.success) {
      var previewJcc = _buildJCCDataObject(patRes.project, sess);
      // Ensure presidingOfficer is carried over from PAT project
      if (patRes.project.presidingOfficer) {
        previewJcc.presidingOfficer = patRes.project.presidingOfficer;
      }
      return {
        success: true,
        jcc: previewJcc,
      };
    }

    return {
      success: false,
      message:
        "Unable to find any JCC or PAT record for Project ID: " + projectId,
    };
  } catch (e) {
    return { success: false, message: "System Error: " + e.message };
  }
}

function getAllJCCs() {
  try {
    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    if (!jccSh || jccSh.getLastRow() < 2) return { success: true, jccs: [] };
    var data = jccSh.getDataRange().getValues();
    var jccs = [];
    for (var i = 1; i < data.length; i++) {
      jccs.push({
        jccId: data[i][0],
        projectId: data[i][1],
        projectName: data[i][2],
        projectNumber: data[i][3],
        certificateId: data[i][6],
        vendor: data[i][5],
        mecName: data[i][9],
        generatedAt: data[i][18],
      });
    }
    return { success: true, jccs: jccs };
  } catch (e) {
    return { success: false, message: e.message, jccs: [] };
  }
}

// ─────────────────────────────────────────────────────────────
//  MAIL SYSTEM
// ─────────────────────────────────────────────────────────────

function getMails(token, folder, options) {
  try {
    options = options || {};
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) return { success: true, mails: [] };
    var data = sh.getDataRange().getValues();
    var rows = data.slice(1);

    var filtered = rows
      .filter(function (r) {
        var mailFolder = r[9]; // Folder column
        var sender = r[1]; // SenderEmail
        var receiver = r[3]; // ReceiverEmail
        var starred = r[11]; // Starred column
        var label = r[12]; // Labels column
        var subject = String(r[5] || "").toLowerCase();
        var body = String(r[6] || "").toLowerCase();
        var fromEmail = String(r[1] || "").toLowerCase();
        var toEmail = String(r[3] || "").toLowerCase();
        var deletedBy = String(r[17] || "").toLowerCase();
        var query =
          options && options.query ? String(options.query).toLowerCase() : "";
        var userEmail = String(sess.email || "").toLowerCase();

        // Soft delete filter: exclude if current user deleted this mail
        var deletedList = deletedBy
          ? deletedBy.split(",").map(function (e) {
              return e.trim();
            })
          : [];
        if (deletedList.indexOf(userEmail) !== -1) return false;

        if (folder === "Inbox") {
          if (receiver !== sess.email) return false;
          if (mailFolder !== "Inbox") return false;
        } else if (folder === "Sent") {
          if (sender !== sess.email) return false;
          if (mailFolder !== "Sent") return false;
        } else if (folder === "Self") {
          if (sender !== sess.email && receiver !== sess.email) return false;
          if (sender === receiver && mailFolder !== "Self") return false;
        } else if (folder === "Trash") {
          if (sender !== sess.email && receiver !== sess.email) return false;
          if (mailFolder !== "Trash") return false;
        } else if (folder === "Starred") {
          if (sender !== sess.email && receiver !== sess.email) return false;
          if (
            starred !== true &&
            starred !== "TRUE" &&
            String(starred) !== "true"
          )
            return false;
        } else if (folder === "Archive") {
          if (sender !== sess.email && receiver !== sess.email) return false;
          if (mailFolder !== "Archive") return false;
        } else if (folder === "All Mail") {
          if (sender !== sess.email && receiver !== sess.email) return false;
        }

        // Search filter
        if (query) {
          var match =
            subject.indexOf(query) !== -1 ||
            body.indexOf(query) !== -1 ||
            fromEmail.indexOf(query) !== -1 ||
            toEmail.indexOf(query) !== -1;
          if (!match) return false;
        }

        return true;
      })
      .map(function (r) {
        var att = [];
        try {
          att = JSON.parse(r[10] || "[]");
        } catch (e) {
          att = [];
        }
        var labels = [];
        try {
          labels = JSON.parse(r[12] || "[]");
        } catch (e) {
          labels = [];
        }
        var deletedBy = String(r[17] || "").toLowerCase();
        return {
          mailId: r[0],
          senderEmail: r[1],
          senderName: r[2],
          receiverEmail: r[3],
          receiverName: r[4],
          subject: r[5],
          body: r[6],
          timestamp: r[7],
          status: r[8],
          folder: r[9],
          attachments: att,
          hasAttachments: att.length > 0,
          starred:
            r[11] === true || r[11] === "TRUE" || String(r[11]) === "true",
          labels: labels,
          threadId: r[13] || "",
          priority: r[14] || "normal",
          cc: r[15] || "",
          bcc: r[16] || "",
          deletedBy: deletedBy,
        };
      });

    return { success: true, mails: filtered };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function sendMail(
  token,
  receiverEmail,
  subject,
  body,
  attachments,
  cc,
  bcc,
  labels,
  priority,
  threadId,
) {
  try {
    var sess = _session(token);
    var ccNorm = cc ? String(cc).toLowerCase().trim() : "";
    var bccNorm = bcc ? String(bcc).toLowerCase().trim() : "";
    var labelsArray = labels || [];
    var labelsJson = JSON.stringify(labelsArray);
    var priorityNorm = priority || "normal";
    var threadIdNorm = threadId || "";

    // Accept both single email string and array of emails
    var primaryReceivers = [];
    if (typeof receiverEmail === "string") {
      primaryReceivers = [receiverEmail.toLowerCase().trim()];
    } else if (Array.isArray(receiverEmail)) {
      primaryReceivers = receiverEmail
        .map(function (e) {
          return String(e).toLowerCase().trim();
        })
        .filter(Boolean);
    }

    // Parse CC recipients from the cc string
    var ccReceivers = [];
    if (ccNorm) {
      ccReceivers = ccNorm
        .split(",")
        .map(function (e) {
          return e.trim();
        })
        .filter(Boolean);
    }

    // Remove CC recipients from primary receivers to avoid duplicates
    var ccSet = ccReceivers.join(",");
    primaryReceivers = primaryReceivers.filter(function (r) {
      return ccSet.indexOf(r) === -1;
    });

    // ── TEST ACCOUNTS: never deliver mail to the 4 department test logins ──
    // (mec-test@, project-test@, planning-test@, sd-test@fob.ng). They are
    // throwaway accounts; tagging them clutters their inboxes and confuses the
    // sender's Sent view. Drop them silently and warn.
    var TEST_ACCOUNTS = {
      "mec-test@fob.ng": 1,
      "project-test@fob.ng": 1,
      "planning-test@fob.ng": 1,
      "sd-test@fob.ng": 1,
    };
    var skippedTest = 0;
    primaryReceivers = primaryReceivers.filter(function (r) {
      if (TEST_ACCOUNTS[r]) {
        skippedTest++;
        return false;
      }
      return true;
    });
    ccReceivers = ccReceivers.filter(function (r) {
      if (TEST_ACCOUNTS[r]) {
        skippedTest++;
        return false;
      }
      return true;
    });

    var receivers = primaryReceivers;
    if (receivers.length === 0) throw new Error("No valid recipient emails.");

    var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var userData = userSh.getDataRange().getValues();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var mailId = _genId("MAIL");
    var now = new Date().toISOString();
    var attachmentsJson = JSON.stringify(attachments || []);
    var mailCount = 0;

    // Build a batch array instead of calling appendRow per recipient
    var batchRows = [];
    var skippedBanned = 0;

    receivers.forEach(function (receiverEmailNorm) {
      var receiver = userData.find(function (r) {
        return String(r[2]).toLowerCase().trim() === receiverEmailNorm;
      });
      if (!receiver) return; // skip invalid recipients silently
      // BANNED RECIPIENT: never deliver mail to a banned account.
      if (String(receiver[9] || "").toLowerCase() === "banned") {
        skippedBanned++;
        return;
      }
      var receiverName = receiver[1];
      var isSelf = sess.email.toLowerCase().trim() === receiverEmailNorm;
      var senderFolder = isSelf ? "Self" : "Sent";
      var receiverFolder = isSelf ? "Self" : "Inbox";

      // Add to Receiver's folder
      batchRows.push([
        mailId,
        sess.email,
        sess.name,
        receiverEmailNorm,
        receiverName,
        subject,
        body,
        now,
        "unread",
        receiverFolder,
        attachmentsJson,
        false,
        labelsJson,
        threadIdNorm,
        priorityNorm,
        ccNorm,
        bccNorm,
        "",
      ]);
      // Add to Sender's folder
      batchRows.push([
        mailId,
        sess.email,
        sess.name,
        receiverEmailNorm,
        receiverName,
        subject,
        body,
        now,
        "read",
        senderFolder,
        attachmentsJson,
        false,
        labelsJson,
        threadIdNorm,
        priorityNorm,
        ccNorm,
        bccNorm,
        "",
      ]);
      mailCount++;
    });

    // Also create rows for CC recipients
    ccReceivers.forEach(function (receiverEmailNorm) {
      var receiver = userData.find(function (r) {
        return String(r[2]).toLowerCase().trim() === receiverEmailNorm;
      });
      if (!receiver) return;
      if (String(receiver[9] || "").toLowerCase() === "banned") {
        skippedBanned++;
        return;
      }
      var receiverName = receiver[1];
      var isSelf = sess.email.toLowerCase().trim() === receiverEmailNorm;
      var senderFolder = isSelf ? "Self" : "Sent";
      var receiverFolder = isSelf ? "Self" : "Inbox";

      batchRows.push([
        mailId,
        sess.email,
        sess.name,
        receiverEmailNorm,
        receiverName,
        subject,
        body,
        now,
        "unread",
        receiverFolder,
        attachmentsJson,
        false,
        labelsJson,
        threadIdNorm,
        priorityNorm,
        ccNorm,
        bccNorm,
        "",
      ]);
      batchRows.push([
        mailId,
        sess.email,
        sess.name,
        receiverEmailNorm,
        receiverName,
        subject,
        body,
        now,
        "read",
        senderFolder,
        attachmentsJson,
        false,
        labelsJson,
        threadIdNorm,
        priorityNorm,
        ccNorm,
        bccNorm,
        "",
      ]);
      mailCount++;
    });

    if (mailCount === 0)
      throw new Error("No valid recipients found in system.");

    // IDEMPOTENCY GUARD: block duplicate sends.
    // A single send action can reach the backend more than once when a user
    // double-clicks Send, presses Ctrl/Cmd+Enter then clicks, or the frontend
    // retries/fires twice. Every duplicate produces extra copies in the sheet
    // (the user saw one message appear 4x). We detect an identical in-flight
    // send (same sender + recipients + subject + body + thread within 8s) and
    // return the EXISTING mail instead of writing new rows.
    var dupWindowMs = 8000;
    var dupNow = new Date().getTime();
    var allRecipientsKey = primaryReceivers
      .concat(ccReceivers)
      .map(function (r) {
        return r;
      })
      .sort()
      .join(",");
    var existingAll = sh.getDataRange().getValues();
    for (var d = 1; d < existingAll.length; d++) {
      var existingSender = String(existingAll[d][1] || "")
        .toLowerCase()
        .trim();
      var existingSubject = String(existingAll[d][5] || "")
        .toLowerCase()
        .trim();
      var existingBody = String(existingAll[d][6] || "")
        .toLowerCase()
        .trim();
      var existingThread = String(existingAll[d][13] || "").trim();
      var existingTs = existingAll[d][7]
        ? new Date(existingAll[d][7]).getTime()
        : 0;
      // recipient set encoded in ReceiverEmail + CC columns
      var existingRecipientsKey = [
        String(existingAll[d][3] || "")
          .toLowerCase()
          .trim(),
      ]
        .concat(
          String(existingAll[d][15] || "")
            .toLowerCase()
            .split(",")
            .map(function (x) {
              return x.trim();
            })
            .filter(Boolean),
        )
        .sort()
        .join(",");
      var recipMatch =
        existingRecipientsKey === allRecipientsKey ||
        (allRecipientsKey.length > 0 &&
          existingRecipientsKey.indexOf(allRecipientsKey) !== -1);
      if (
        existingSender === sess.email.toLowerCase().trim() &&
        existingSubject ===
          String(subject || "")
            .toLowerCase()
            .trim() &&
        existingBody ===
          String(body || "")
            .toLowerCase()
            .trim() &&
        existingThread === threadIdNorm &&
        recipMatch &&
        existingTs > 0 &&
        dupNow - existingTs <= dupWindowMs
      ) {
        // Duplicate send detected — return the already-written mail, do NOT write again.
        return {
          success: true,
          mailId: existingAll[d][0],
          duplicate: true,
          attachments: attachments || [],
        };
      }
    }

    // BATCH WRITE: single API call instead of 2*N appendRow calls
    if (batchRows.length > 0) {
      var lastRow = sh.getLastRow();
      sh.getRange(
        lastRow + 1,
        1,
        batchRows.length,
        batchRows[0].length,
      ).setValues(batchRows);
    }

    // Back-fill threadId to existing mails in the same conversation
    // so the original mail appears in the thread history (WhatsApp-style grouping)
    if (
      threadIdNorm &&
      (primaryReceivers.length > 0 || ccReceivers.length > 0)
    ) {
      var cleanSubject = String(subject || "")
        .replace(/^(Re:\s*)+/i, "")
        .toLowerCase()
        .trim();
      var sessEmail = String(sess.email || "")
        .toLowerCase()
        .trim();
      var allTargetEmails = primaryReceivers
        .concat(ccReceivers)
        .filter(Boolean);
      var existingData = sh.getDataRange().getValues();
      for (var r = 1; r < existingData.length; r++) {
        var existingRowSubject = String(existingData[r][5] || "")
          .replace(/^(Re:\s*)+/i, "")
          .toLowerCase()
          .trim();
        var existingSender = String(existingData[r][1] || "")
          .toLowerCase()
          .trim();
        var existingReceiver = String(existingData[r][3] || "")
          .toLowerCase()
          .trim();
        var existingThreadId = String(existingData[r][13] || "").trim();
        // Match: same conversation, involves any participant, no threadId yet
        var matchesParticipant = allTargetEmails.some(function (t) {
          return existingSender === t || existingReceiver === t;
        });
        if (
          existingRowSubject === cleanSubject &&
          (existingSender === sessEmail || matchesParticipant) &&
          existingThreadId === ""
        ) {
          sh.getRange(r + 1, 14).setValue(threadIdNorm);
        }
      }
    }

    var ret = { success: true, mailId: mailId, attachments: attachments || [] };
    var warnParts = [];
    if (skippedTest > 0) {
      warnParts.push(
        skippedTest +
          " test account(s) were skipped (MEC/PROJECT/PLANNING/SD-METRO test users are not real recipients).",
      );
    }
    if (skippedBanned > 0) {
      warnParts.push(
        skippedBanned +
          " banned account(s) were skipped (mail is never delivered to banned users).",
      );
    }
    if (warnParts.length > 0) ret.warning = warnParts.join(" ");

    return ret;
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getMailById(token, mailId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        var att = [];
        try {
          att = JSON.parse(data[i][10] || "[]");
        } catch (e) {
          att = [];
        }
        var labels = [];
        try {
          labels = JSON.parse(data[i][12] || "[]");
        } catch (e) {
          labels = [];
        }
        var deletedBy = String(data[i][17] || "").toLowerCase();
        var userEmail = String(sess.email || "").toLowerCase();
        var deletedList = deletedBy
          ? deletedBy.split(",").map(function (e) {
              return e.trim();
            })
          : [];
        if (deletedList.indexOf(userEmail) !== -1) {
          return { success: false, message: "You have deleted this mail." };
        }

        var mail = {
          mailId: data[i][0],
          senderEmail: data[i][1],
          senderName: data[i][2],
          receiverEmail: data[i][3],
          receiverName: data[i][4],
          subject: data[i][5],
          body: data[i][6],
          timestamp: data[i][7],
          status: data[i][8],
          folder: data[i][9],
          attachments: att,
          hasAttachments: att.length > 0,
          starred:
            data[i][11] === true ||
            data[i][11] === "TRUE" ||
            String(data[i][11]) === "true",
          labels: labels,
          threadId: data[i][13] || "",
          priority: data[i][14] || "normal",
          cc: data[i][15] || "",
          bcc: data[i][16] || "",
          deletedBy: deletedBy,
        };

        // SECURITY: IDOR fix — verify user owns this mail or is admin
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          var isSender =
            String(mail.senderEmail || "").toLowerCase() ===
            String(sess.email || "").toLowerCase();
          var isReceiver =
            String(mail.receiverEmail || "").toLowerCase() ===
            String(sess.email || "").toLowerCase();
          if (!isSender && !isReceiver) {
            return {
              success: false,
              message: "Permission denied: You can only view your own emails.",
            };
          }
        }

        sh.getRange(i + 1, 9).setValue("read");
        return { success: true, mail: mail };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function markMailAsRead(token, mailId) {
  try {
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        // IDOR fix: only the sender or receiver can mark mail as read
        var senderEmail = String(data[i][1] || "").toLowerCase();
        var receiverEmail = String(data[i][3] || "").toLowerCase();
        var userEmail = String(sess.email || "").toLowerCase();
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          if (senderEmail !== userEmail && receiverEmail !== userEmail) {
            return {
              success: false,
              message:
                "Permission denied: You can only access your own emails.",
            };
          }
        }
        sh.getRange(i + 1, 9).setValue("read");
        return { success: true, message: "Mail marked as read." };
      }
    }
    throw new Error("Mail not found: " + mailId);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteMail(token, mailId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        var email = String(sess.email || "").toLowerCase();
        var senderEmail = String(data[i][1] || "").toLowerCase();
        var receiverEmail = String(data[i][3] || "").toLowerCase();
        var role = String(sess.role || "").toLowerCase();

        // SECURITY: Only allow owner (sender/receiver) or admin to delete mail
        if (role !== "super admin" && role !== "admin") {
          if (senderEmail !== email && receiverEmail !== email) {
            return {
              success: false,
              message:
                "Permission denied: You can only delete your own emails.",
            };
          }
        }

        // Soft delete: mark this user as deleted
        var deletedBy = String(data[i][17] || "").toLowerCase();
        var deletedList = deletedBy
          ? deletedBy.split(",").map(function (e) {
              return e.trim();
            })
          : [];
        if (deletedList.indexOf(email) === -1) {
          deletedList.push(email);
        }
        sh.getRange(i + 1, 18).setValue(deletedList.join(","));
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAllMails(token) {
  try {
    var sess = _session(token);
    var role = String(sess.role || "").toLowerCase();
    if (role !== "super admin" && role !== "admin") {
      throw new Error("Only admins can view all mails.");
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) return { success: true, mails: [] };
    var data = sh.getDataRange().getValues();
    var rows = data.slice(1).map(function (r) {
      var att = [];
      try {
        att = JSON.parse(r[10] || "[]");
      } catch (e) {
        att = [];
      }
      var labels = [];
      try {
        labels = JSON.parse(r[12] || "[]");
      } catch (e) {
        labels = [];
      }
      return {
        mailId: r[0],
        senderEmail: r[1],
        senderName: r[2],
        receiverEmail: r[3],
        receiverName: r[4],
        subject: r[5],
        body: r[6],
        timestamp: r[7],
        status: r[8],
        folder: r[9],
        attachments: att,
        hasAttachments: att.length > 0,
        starred: r[11] === true || r[11] === "TRUE" || String(r[11]) === "true",
        labels: labels,
        threadId: r[13] || "",
        priority: r[14] || "normal",
        cc: r[15] || "",
        bcc: r[16] || "",
      };
    });
    return { success: true, mails: rows };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteMailPermanently(token, mailId) {
  try {
    var sess = _session(token);
    var role = String(sess.role || "").toLowerCase();
    if (role !== "super admin") {
      throw new Error("Only super admin can permanently delete mails.");
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function craftAndSendMail(token, receiverEmail, subject, body) {
  try {
    var sess = _session(token);
    var receiverEmailNorm = String(receiverEmail).toLowerCase().trim();
    var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var userData = userSh.getDataRange().getValues();
    var receiver = null;
    for (var i = 1; i < userData.length; i++) {
      if (String(userData[i][2]).toLowerCase().trim() === receiverEmailNorm) {
        receiver = userData[i];
        break;
      }
    }
    if (!receiver) throw new Error("Receiver email not found in system.");
    var receiverName = receiver[1];
    var mailId = _genId("MAIL");
    var now = new Date().toISOString();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    sh.appendRow([
      mailId,
      sess.email,
      sess.name,
      receiverEmailNorm,
      receiverName,
      subject,
      body,
      now,
      "unread",
      "Inbox",
      "[]",
      false,
      "",
      "",
      "normal",
      "",
      "",
    ]);
    sh.appendRow([
      mailId,
      sess.email,
      sess.name,
      receiverEmailNorm,
      receiverName,
      subject,
      body,
      now,
      "read",
      "Sent",
      "[]",
      false,
      "",
      "",
      "normal",
      "",
      "",
    ]);
    return { success: true, mailId: mailId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ══════════════════════════════════════════════════════════════
// ADVANCED MAIL FUNCTIONS
// ══════════════════════════════════════════════════════════════

function toggleStarMail(token, mailId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        var senderEmail = String(data[i][1] || "").toLowerCase();
        var receiverEmail = String(data[i][3] || "").toLowerCase();
        var email = String(sess.email || "").toLowerCase();
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: "Permission denied." };
          }
        }
        var currentStarred = data[i][11];
        var newStarred = !(
          currentStarred === true ||
          currentStarred === "TRUE" ||
          String(currentStarred) === "true"
        );
        sh.getRange(i + 1, 12).setValue(newStarred);
        return { success: true, starred: newStarred };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function archiveMail(token, mailId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        var senderEmail = String(data[i][1] || "").toLowerCase();
        var receiverEmail = String(data[i][3] || "").toLowerCase();
        var email = String(sess.email || "").toLowerCase();
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: "Permission denied." };
          }
        }
        sh.getRange(i + 1, 10).setValue("Archive");
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function restoreMail(token, mailId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        var senderEmail = String(data[i][1] || "").toLowerCase();
        var receiverEmail = String(data[i][3] || "").toLowerCase();
        var email = String(sess.email || "").toLowerCase();
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: "Permission denied." };
          }
        }
        sh.getRange(i + 1, 10).setValue("Inbox");
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function markMailAsUnread(token, mailId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        var senderEmail = String(data[i][1] || "").toLowerCase();
        var receiverEmail = String(data[i][3] || "").toLowerCase();
        var email = String(sess.email || "").toLowerCase();
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: "Permission denied." };
          }
        }
        sh.getRange(i + 1, 9).setValue("unread");
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function forwardMail(token, mailId, toEmail) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    var originalMail = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === mailId) {
        originalMail = data[i];
        break;
      }
    }
    if (!originalMail) throw new Error("Mail not found.");

    var senderEmail = String(originalMail[1] || "").toLowerCase();
    var receiverEmail = String(originalMail[3] || "").toLowerCase();
    var email = String(sess.email || "").toLowerCase();
    var role = String(sess.role || "").toLowerCase();
    if (role !== "super admin" && role !== "admin") {
      if (senderEmail !== email && receiverEmail !== email) {
        return { success: false, message: "Permission denied." };
      }
    }

    var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var userData = userSh.getDataRange().getValues();
    var toEmailNorm = String(toEmail).toLowerCase().trim();
    var toUser = null;
    for (var j = 1; j < userData.length; j++) {
      if (String(userData[j][2]).toLowerCase().trim() === toEmailNorm) {
        toUser = userData[j];
        break;
      }
    }
    if (!toUser) throw new Error("Forward recipient not found in system.");

    var mailIdNew = _genId("MAIL");
    var now = new Date().toISOString();
    var subject = "Fwd: " + (originalMail[5] || "");
    var body =
      "\n\n---------- Forwarded message ----------\nFrom: " +
      originalMail[2] +
      " <" +
      originalMail[1] +
      ">\nDate: " +
      originalMail[7] +
      "\nSubject: " +
      originalMail[5] +
      "\n\n" +
      (originalMail[6] || "");
    var originalAtt = [];
    try {
      originalAtt = JSON.parse(originalMail[10] || "[]");
    } catch (e) {
      originalAtt = [];
    }
    var attJson = JSON.stringify(originalAtt);

    sh.appendRow([
      mailIdNew,
      sess.email,
      sess.name,
      toEmailNorm,
      toUser[1],
      subject,
      body,
      now,
      "unread",
      "Inbox",
      attJson,
      false,
      "[]",
      "",
      "normal",
      "",
      "",
    ]);
    sh.appendRow([
      mailIdNew,
      sess.email,
      sess.name,
      toEmailNorm,
      toUser[1],
      subject,
      body,
      now,
      "read",
      "Sent",
      attJson,
      false,
      "[]",
      "",
      "normal",
      "",
      "",
    ]);

    return { success: true, mailId: mailIdNew };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function searchMails(token, query, folder) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) return { success: true, mails: [] };
    var data = sh.getDataRange().getValues();
    var rows = data.slice(1);
    var queryLower = String(query || "").toLowerCase();

    var filtered = rows
      .filter(function (r) {
        var mailFolder = r[9];
        var sender = r[1];
        var receiver = r[3];
        var subject = String(r[5] || "").toLowerCase();
        var body = String(r[6] || "").toLowerCase();

        if (
          folder === "Inbox" &&
          (receiver !== sess.email || mailFolder !== "Inbox")
        )
          return false;
        if (
          folder === "Sent" &&
          (sender !== sess.email || mailFolder !== "Sent")
        )
          return false;
        if (
          folder === "Trash" &&
          ((sender !== sess.email && receiver !== sess.email) ||
            mailFolder !== "Trash")
        )
          return false;
        if (
          folder === "Starred" &&
          sender !== sess.email &&
          receiver !== sess.email
        )
          return false;
        if (
          folder === "Archive" &&
          ((sender !== sess.email && receiver !== sess.email) ||
            mailFolder !== "Archive")
        )
          return false;
        if (!folder && sender !== sess.email && receiver !== sess.email)
          return false;

        if (queryLower) {
          return (
            subject.indexOf(queryLower) !== -1 ||
            body.indexOf(queryLower) !== -1
          );
        }
        return true;
      })
      .map(function (r) {
        var att = [];
        try {
          att = JSON.parse(r[10] || "[]");
        } catch (e) {
          att = [];
        }
        var labels = [];
        try {
          labels = JSON.parse(r[12] || "[]");
        } catch (e) {
          labels = [];
        }
        return {
          mailId: r[0],
          senderEmail: r[1],
          senderName: r[2],
          receiverEmail: r[3],
          receiverName: r[4],
          subject: r[5],
          body: r[6],
          timestamp: r[7],
          status: r[8],
          folder: r[9],
          attachments: att,
          hasAttachments: att.length > 0,
          starred:
            r[11] === true || r[11] === "TRUE" || String(r[11]) === "true",
          labels: labels,
          threadId: r[13] || "",
          priority: r[14] || "normal",
          cc: r[15] || "",
          bcc: r[16] || "",
        };
      });

    return { success: true, mails: filtered };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function emptyTrash(token) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var role = String(sess.role || "").toLowerCase();
    if (role !== "super admin" && role !== "admin") {
      throw new Error("Only admins can empty trash.");
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][9] === "Trash") {
        rowsToDelete.push(i + 1);
      }
    }
    for (var j = 0; j < rowsToDelete.length; j++) {
      sh.deleteRow(rowsToDelete[j]);
    }
    return { success: true, deletedCount: rowsToDelete.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getAllMailCounts(token) {
  try {
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) {
      return {
        success: true,
        counts: {
          Inbox: 0,
          Sent: 0,
          Trash: 0,
          Starred: 0,
          Archive: 0,
          unreadInbox: 0,
        },
      };
    }
    var data = sh.getDataRange().getValues();
    var rows = data.slice(1);
    var counts = {
      Inbox: 0,
      Sent: 0,
      Self: 0,
      Trash: 0,
      Starred: 0,
      Archive: 0,
      unreadInbox: 0,
    };
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var mailFolder = String(r[9] || "");
      var sender = String(r[1] || "").toLowerCase();
      var receiver = String(r[3] || "").toLowerCase();
      var starred =
        r[11] === true || r[11] === "TRUE" || String(r[11]) === "true";
      var status = String(r[8] || "").toLowerCase();
      var email = String(sess.email || "").toLowerCase();
      var isOwner = sender === email || receiver === email;
      if (!isOwner) continue;

      // SOFT-DELETE FILTER: mirror getMails() — skip rows this user deleted,
      // otherwise the badge keeps counting mails the user has already removed
      // (so a deleted mail would still inflate Inbox/Sent after delete AND reload).
      var deletedBy = String(r[17] || "").toLowerCase();
      var deletedList = deletedBy
        ? deletedBy.split(",").map(function (e) {
            return e.trim();
          })
        : [];
      if (deletedList.indexOf(email) !== -1) continue;

      if (mailFolder === "Inbox" && receiver === email) counts.Inbox++;
      else if (mailFolder === "Sent" && sender === email) counts.Sent++;
      else if (mailFolder === "Self" && sender === email && sender === receiver)
        counts.Self++;
      else if (mailFolder === "Trash" && isOwner) counts.Trash++;
      else if (mailFolder === "Archive" && isOwner) counts.Archive++;
      else if (starred && isOwner) counts.Starred++;

      if (mailFolder === "Inbox" && receiver === email && status === "unread")
        counts.unreadInbox++;
    }
    return { success: true, counts: counts };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * ROLE-AWARE NOTIFICATION FEED (server-authoritative).
 * Returns a list of pending notifications for the current user, computed from
 * live data so the frontend never has to guess roles:
 *   - mail: unread inbox mail addressed to this user
 *   - pat: PAT projects whose current owner-department == this user's department
 *          (i.e. a PAT just arrived in their queue), or awaiting their action
 *   - audit (Audit dept / audit role): completed projects pending review,
 *          plus any PAT currently in a review/owner stage
 * Each item carries a stable id so the client can dedupe + schedule reminders.
 * @param {string} token
 * @param {string} sinceIso  (optional) only items updated after this timestamp
 */
function getMyNotifications(token, sinceIso) {
  try {
    var sess = _session(token);
    var email = String(sess.email || "").toLowerCase();
    var role = String(sess.role || "").toLowerCase();
    var dept = String(sess.department || "").toLowerCase();
    var isAudit = role === "audit" || dept === "audit";
    var isSuperAdmin = role === "super admin";
    var out = [];

    // ── MAIL: unread inbox addressed to this user ──
    try {
      var mailSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SD.MAILS,
      );
      if (mailSh && mailSh.getLastRow() >= 2) {
        var mdata = mailSh.getDataRange().getValues();
        for (var i = 1; i < mdata.length; i++) {
          var r = mdata[i];
          var folder = String(r[9] || "");
          var receiver = String(r[3] || "").toLowerCase();
          var status = String(r[8] || "").toLowerCase();
          var sender = String(r[1] || "");
          if (folder === "Inbox" && receiver === email && status === "unread") {
            var subj = String(r[6] || "(no subject)");
            out.push({
              id: "mail:" + String(r[0] || i),
              kind: "mail",
              title: "New mail from " + sender,
              body: subj,
              href: "#mail",
              ts: String(r[10] || ""),
              priority: "normal",
            });
          }
        }
      }
    } catch (e) {
      /* mail sheet optional */
    }

    // ── PAT: projects whose owner-department == this user's department ──
    try {
      var patRes = getPATProjects(token);
      var projs = (patRes && patRes.projects) || [];
      projs.forEach(function (p) {
        var status = p.workflowStatus || "Draft";
        var owner = (PAT_OWNER_MAP[status] || "").toLowerCase();
        var matchesDept =
          owner === dept ||
          (owner === "__all__" && dept !== "mec" && !isSuperAdmin) ||
          (isSuperAdmin && false); // super admin sees via audit/summary, not personal queue
        // Also match by explicit assignment
        var assignedDept = String(p.assignedToDept || "").toLowerCase();
        if (assignedDept && assignedDept === dept) matchesDept = true;

        if (isAudit) {
          // Audit console cares about COMPLETED projects (reviewable) + anything in review.
          if (status === "Completed") {
            out.push({
              id: "pat:" + p.projectId,
              kind: "audit",
              title: "Project ready for audit: " + p.projectName,
              body:
                "Completed PAT " +
                p.projectId +
                " is now available for compliance review.",
              href: "#audit",
              ts: String(p.updatedAt || ""),
              createdAt: String(p.submittedAt || p.updatedAt || ""),
              priority: "high",
            });
          } else if (owner === dept || assignedDept === dept) {
            out.push({
              id: "pat:" + p.projectId,
              kind: "pat",
              title: "PAT in your queue: " + p.projectName,
              body:
                p.projectId +
                ' is now at "' +
                status +
                '" — action may be required.',
              href: "#pat-projects",
              ts: String(p.updatedAt || ""),
              createdAt: String(p.submittedAt || p.updatedAt || ""),
              priority: "normal",
            });
          }
        } else if (
          matchesDept &&
          status !== "Completed" &&
          status !== "Draft"
        ) {
          out.push({
            id: "pat:" + p.projectId,
            kind: "pat",
            title: "PAT moved to your department: " + p.projectName,
            body:
              p.projectId +
              ' is now at "' +
              status +
              '" (' +
              (STATUS_TO_DEPT[status] || status) +
              ").",
            href: "#pat-projects",
            ts: String(p.updatedAt || ""),
            createdAt: String(p.submittedAt || p.updatedAt || ""),
            priority: "high",
          });
        }
      });
    } catch (e) {
      /* PAT sheet optional */
    }

    // Sort newest first
    out.sort(function (a, b) {
      var ta = a.ts ? new Date(a.ts).getTime() : 0;
      var tb = b.ts ? new Date(b.ts).getTime() : 0;
      return tb - ta;
    });

    return {
      success: true,
      notifications: out,
      role: role,
      department: dept,
      isAudit: isAudit,
    };
  } catch (e) {
    return { success: false, message: e.message, notifications: [] };
  }
}

function getMailThread(token, threadId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    if (!threadId) throw new Error("Thread ID required.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    var header = data[0];
    var mails = [];
    // DEDUP: a single sent message is stored as TWO rows (receiver's Inbox copy +
    // sender's Sent/Self copy) sharing the same mailId. Both rows carry the same
    // threadId, so grouping by threadId alone would render every message TWICE in
    // the thread timeline. Collapse to one row per unique mailId.
    var seenMailIds = {};

    for (var i = 1; i < data.length; i++) {
      if (data[i][13] === threadId) {
        var dupMid = data[i][0];
        if (seenMailIds[dupMid]) continue; // skip the duplicate copy
        seenMailIds[dupMid] = true;
        var att = [];
        try {
          att = JSON.parse(data[i][10] || "[]");
        } catch (e) {
          att = [];
        }
        var labels = [];
        try {
          labels = JSON.parse(data[i][12] || "[]");
        } catch (e) {
          labels = [];
        }

        var mail = {
          mailId: data[i][0],
          senderEmail: data[i][1],
          senderName: data[i][2],
          receiverEmail: data[i][3],
          receiverName: data[i][4],
          subject: data[i][5],
          body: data[i][6],
          timestamp: data[i][7],
          status: data[i][8],
          folder: data[i][9],
          attachments: att,
          hasAttachments: att.length > 0,
          starred:
            data[i][11] === true ||
            data[i][11] === "TRUE" ||
            String(data[i][11]) === "true",
          labels: labels,
          threadId: data[i][13] || "",
          priority: data[i][14] || "normal",
          cc: data[i][15] || "",
          bcc: data[i][16] || "",
        };

        // SECURITY: IDOR fix — verify user owns this mail or is admin
        var role = String(sess.role || "").toLowerCase();
        if (role !== "super admin" && role !== "admin") {
          var isSender =
            String(mail.senderEmail || "").toLowerCase() ===
            String(sess.email || "").toLowerCase();
          var isReceiver =
            String(mail.receiverEmail || "").toLowerCase() ===
            String(sess.email || "").toLowerCase();
          if (!isSender && !isReceiver) {
            continue;
          }
        }

        mails.push(mail);
      }
    }

    // Sort by timestamp ascending
    mails.sort(function (a, b) {
      return new Date(a.timestamp) - new Date(b.timestamp);
    });

    return { success: true, mails: mails };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Clear all mails — either all accounts (super admin) or a specific account (admin/super admin).
 * targetEmail: "ALL" = hard-delete every row in SD_MAILS (super admin only)
 *              "user@example.com" = hard-delete all rows where that email is sender or receiver
 */
function clearUserMails(token, targetEmail) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var role = String(sess.role || "").toLowerCase();

    // Super admin required for ALL accounts, admin/super admin for specific account
    if (!targetEmail || targetEmail === "ALL") {
      if (role !== "super admin") {
        throw new Error(
          "Only super admin can clear all mails across all accounts.",
        );
      }
    } else {
      if (role !== "super admin" && role !== "admin") {
        throw new Error("Only admins can clear mails for a specific account.");
      }
    }

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) {
      return { success: true, deletedCount: 0, message: "No mails to clear." };
    }

    var data = sh.getDataRange().getValues();
    var rowsToDelete = []; // collect 1-based row numbers to delete

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var senderEmail = String(row[1] || "").toLowerCase();
      var receiverEmail = String(row[3] || "").toLowerCase();

      if (targetEmail === "ALL") {
        rowsToDelete.push(i + 1);
      } else {
        var target = String(targetEmail || "").toLowerCase();
        if (senderEmail === target || receiverEmail === target) {
          rowsToDelete.push(i + 1);
        }
      }
    }

    // Delete from bottom to top to avoid shifting issues
    rowsToDelete.reverse();
    for (var j = 0; j < rowsToDelete.length; j++) {
      sh.deleteRow(rowsToDelete[j]);
    }

    var mode = targetEmail === "ALL" ? "All accounts" : targetEmail;
    return {
      success: true,
      deletedCount: rowsToDelete.length,
      message: "Cleared " + rowsToDelete.length + " mails for " + mode + ".",
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function handleRequest(e) {
  console.log("Request received: " + JSON.stringify(e.parameter));
  var action, args, requestOrigin;
  var responseHeaders = { "Access-Control-Allow-Origin": "*" };

  if (e.postData && e.postData.contents) {
    try {
      var payload = JSON.parse(e.postData.contents);
      action = payload.action;
      args = payload.args || [];
      requestOrigin = payload._origin || "";
    } catch (err) {
      return _json({
        success: false,
        message: "API ERROR: Invalid JSON payload",
      });
    }
  } else {
    action = e.parameter.action;
    requestOrigin = e.parameter._origin || "";
    try {
      args = e.parameter.data ? JSON.parse(e.parameter.data) : [];
    } catch (err) {
      args = [];
    }
  }

  // ── DOMAIN LOCK: Only allow requests from whitelisted origins ────
  // The frontend sends window.location.origin as _origin in every request.
  // This prevents cloned pages hosted on unknown domains from using the API.
  if (!_isOriginAllowed(requestOrigin)) {
    // Allow requests from google.script.run (same-origin, origin is empty string)
    if (requestOrigin !== "") {
      console.warn(
        "Blocked request from unauthorized origin: " + requestOrigin,
      );
      return _json({
        success: false,
        message:
          "Access denied: This backend is locked to the official FiberOne portal. Unauthorized origin: " +
          requestOrigin,
      });
    }
  }

  if (!action)
    return _json({
      success: false,
      message: "API ERROR: No action specified.",
    });

  try {
    var result;
    switch (action) {
      case "loginUser":
        result = loginUser.apply(null, args);
        break;
      case "registerUser":
        result = registerUser.apply(null, args);
        break;
      case "getDepartments":
        result = getDepartments.apply(null, args);
        break;
      case "getUsers":
        result = getUsers.apply(null, args);
        break;
      case "searchUsers":
        result = searchUsers.apply(null, args);
        break;
      case "getDepartmentMembers":
        result = getDepartmentMembers.apply(null, args);
        break;
      case "updateMessagingSettings":
        result = updateMessagingSettings.apply(null, args);
        break;
      case "adminCreateUser":
        result = adminCreateUser.apply(null, args);
        break;
      case "updateUserProfile":
        result = updateUserProfile.apply(null, args);
        break;
      case "addDepartment":
        result = addDepartment.apply(null, args);
        break;
      case "updateDepartment":
        result = updateDepartment.apply(null, args);
        break;
      case "deleteDepartment":
        result = deleteDepartment.apply(null, args);
        break;
      case "updateUserRole":
        result = updateUserRole.apply(null, args);
        break;
      case "updateUserStatus":
        result = updateUserStatus.apply(null, args);
        break;
      case "getPublicDepartments":
        result = getPublicDepartments.apply(null, args);
        break;
      case "resetUserPassword":
        result = resetUserPassword.apply(null, args);
        break;
      case "getDomainConfig":
        result = getDomainConfig.apply(null, args);
        break;
      case "getAIResponse":
        result = getAIResponse.apply(null, args);
        break;
      case "callGemini":
        result = callGemini.apply(null, args);
        break;
      case "callGeminiVision":
        result = callGeminiVision.apply(null, args);
        break;
      case "logEmployeeMessage":
        result = logEmployeeMessage.apply(null, args);
        break;
      case "getPendingCheckin":
        result = getPendingCheckin.apply(null, args);
        break;
      case "respondCheckin":
        result = respondCheckin.apply(null, args);
        break;
      case "completeCheckin":
        result = completeCheckin.apply(null, args);
        break;
      case "getMySessionInfo":
        result = getMySessionInfo.apply(null, args);
        break;
      case "logoutUser":
        result = logoutUser.apply(null, args);
        break;
      case "getUserById":
        result = getUserById.apply(null, args);
        break;
      case "getUsersByDepartment":
        result = getUsersByDepartment.apply(null, args);
        break;
      case "getUserByEmail":
        result = getUserByEmail.apply(null, args);
        break;
      case "adminDeleteUser":
        result = adminDeleteUser.apply(null, args);
        break;
      case "adminUpdateUser":
        result = adminUpdateUser.apply(null, args);
        break;
      case "approvePAT":
        result = approvePAT.apply(null, args);
        break;
      case "getNextStageUsers":
        result = getNextStageUsers.apply(null, args);
        break;
      case "createDashboardShareToken":
        result = createDashboardShareToken.apply(null, args);
        break;
      case "getSharedDashboard":
        result = getSharedDashboard.apply(null, args);
        break;
      case "getPATAnalyticsForIdris":
        result = getPATAnalyticsForIdris.apply(null, args);
        break;
      case "getAuditAnalytics":
        result = getAuditAnalytics.apply(null, args);
        break;
      case "getMyNotifications":
        result = getMyNotifications.apply(null, args);
        break;
      case "getMaterials":
        result = getMaterials.apply(null, args);
        break;
      case "addMaterial":
        result = addMaterial.apply(null, args);
        break;
      case "deleteMaterial":
        result = deleteMaterial.apply(null, args);
        break;
      case "getPATProjects":
        result = getPATProjects.apply(null, args);
        break;
      case "getPATProjectById":
        result = getPATProjectById.apply(null, args);
        break;
      case "savePATProject":
        result = savePATProject.apply(null, args);
        break;
      case "getPATImages":
        result = getPATImages.apply(null, args);
        break;
      case "savePATImages":
        result = savePATImages.apply(null, args);
        break;
      case "addSinglePATImage":
        result = addSinglePATImage.apply(null, args);
        break;
      case "deletePATImages":
        result = deletePATImages.apply(null, args);
        break;
      case "saveProjectImages":
        result = saveProjectImages.apply(null, args);
        break;
      case "submitPATToDepartment":
        result = submitPATToDepartment.apply(null, args);
        break;
      case "submitPATToNextStage":
        result = submitPATToNextStage.apply(null, args);
        break;
      case "routePAT":
        result = routePAT.apply(null, args);
        break;
      case "partiallyApprovePAT":
        result = partiallyApprovePAT.apply(null, args);
        break;
      case "deletePATProject":
        result = deletePATProject.apply(null, args);
        break;
      case "deleteAllDraftPATProjects":
        result = deleteAllDraftPATProjects.apply(null, args);
        break;
      case "deleteAllPATProjects":
        result = deleteAllPATProjects.apply(null, args);
        break;
      case "getPATProjectsWithPermissions":
        result = getPATProjectsWithPermissions.apply(null, args);
        break;
      case "checkPATPermission":
        result = checkPATPermission.apply(null, args);
        break;
      case "getSLAAnalytics":
        result = getSLAAnalytics.apply(null, args);
        break;
      case "getAdvancedSLADetails":
        result = getAdvancedSLADetails.apply(null, args);
        break;
      case "getPerformanceScores":
        result = getPerformanceScores.apply(null, args);
        break;
      case "assignPresidingOfficer":
        result = assignPresidingOfficer.apply(null, args);
        break;
      case "getBlacklist":
        result = getBlacklist.apply(null, args);
        break;
      case "banUserWithBlacklist":
        result = banUserWithBlacklist.apply(null, args);
        break;
      case "unbanUserWithBlacklist":
        result = unbanUserWithBlacklist.apply(null, args);
        break;
      case "removeBlacklistEntry":
        result = removeBlacklistEntry.apply(null, args);
        break;
      case "getMails":
        result = getMails.apply(null, args);
        break;
      case "getAllMailCounts":
        result = getAllMailCounts.apply(null, args);
        break;
      case "sendMail":
        result = sendMail.apply(null, args);
        break;
      case "getMailById":
        result = getMailById.apply(null, args);
        break;
      case "markMailAsRead":
        result = markMailAsRead.apply(null, args);
        break;
      case "markMailAsUnread":
        result = markMailAsUnread.apply(null, args);
        break;
      case "deleteMail":
        result = deleteMail.apply(null, args);
        break;
      case "getAllMails":
        result = getAllMails.apply(null, args);
        break;
      case "deleteMailPermanently":
        result = deleteMailPermanently.apply(null, args);
        break;
      case "emptyTrash":
        result = emptyTrash.apply(null, args);
        break;
      case "craftAndSendMail":
        result = craftAndSendMail.apply(null, args);
        break;
      case "toggleStarMail":
        result = toggleStarMail.apply(null, args);
        break;
      case "archiveMail":
        result = archiveMail.apply(null, args);
        break;
      case "restoreMail":
        result = restoreMail.apply(null, args);
        break;
      case "forwardMail":
        result = forwardMail.apply(null, args);
        break;
      case "searchMails":
        result = searchMails.apply(null, args);
        break;
      case "generateJCC":
        result = generateJCC.apply(null, args);
        break;
      case "getJCCByProjectId":
        result = getJCCByProjectId.apply(null, args);
        break;
      case "getAllJCCs":
        result = getAllJCCs.apply(null, args);
        break;
      case "generateVendorReviewLink":
        result = generateVendorReviewLink.apply(null, args);
        break;
      case "getVendorPATByToken":
        result = getVendorPATByToken.apply(null, args);
        break;
      case "submitVendorDecision":
        result = submitVendorDecision.apply(null, args);
        break;
      case "setAllowedOrigins":
        result = setAllowedOrigins.apply(null, args);
        break;
      case "getAllowedOrigins":
        result = getAllowedOrigins.apply(null, args);
        break;
      case "setLocalhostAccess":
        result = setLocalhostAccess.apply(null, args);
        break;
      case "uploadDocument":
        result = uploadDocument.apply(null, args);
        break;
      case "getDocuments":
        result = getDocuments.apply(null, args);
        break;
      case "deleteDocument":
        result = deleteDocument.apply(null, args);
        break;
      case "deleteProjectDocuments":
        result = deleteProjectDocuments.apply(null, args);
        break;
      case "getMailThread":
        result = getMailThread.apply(null, args);
        break;
      case "clearUserMails":
        result = clearUserMails.apply(null, args);
        break;
      default:
        throw new Error("Unknown action: " + action);
    }
    return _json(result);
  } catch (err) {
    return _json({ success: false, message: "SYSTEM CRASH: " + err.message });
  }
}

/**
 * THE MASTER ROUTER
 * Serves the requested HTML page or the API handler.
 * Ensures 'employee.html' requests are funneled to 'admin.html'.
 */
function doGet(e) {
  try {
    // If it's an API call, handle it
    if (e.parameter.action) return handleRequest(e);

    // Serve vendor review page if token parameter is present
    if (e.parameter.token) {
      var vendorTemplate = HtmlService.createTemplateFromFile("vendor-review");
      vendorTemplate.tokenFromServer = e.parameter.token;
      vendorTemplate.scriptUrl = ScriptApp.getService().getUrl();
      return vendorTemplate
        .evaluate()
        .setTitle("Vendor Review - PAT Project")
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Determine which page to serve
    var page = e.parameter.page || "index";

    // UNIFIED INTERFACE ENFORCEMENT
    // If any legacy system or cached link tries to access 'employee', force load 'admin'
    if (page === "employee" || page === "employee.html") page = "admin";
    if (page === "admin.html") page = "admin";
    if (page === "mail.html") page = "mail";
    if (page === "audit.html") page = "audit";
    // Standalone ban page — serves without a session so banned users always land here.
    if (page === "ban" || page === "ban.html") page = "ban";
    // Standalone "email already in the system" page — serves without a session.
    if (page === "email-exists" || page === "email-exists.html")
      page = "email-exists";

    var template = HtmlService.createTemplateFromFile(page);
    template.scriptUrl = ScriptApp.getService().getUrl();

    return template
      .evaluate()
      .setTitle("FiberOne Portal")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    console.error("doGet Error: " + err.message);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px;color:#dc2626">' +
        "<h2>System Routing Error</h2>" +
        "<p>The server encountered an error while trying to load the page.</p>" +
        '<pre style="background:#fef2f2;padding:10px;border:1px solid #fca5a5">' +
        err.message +
        "</pre>" +
        '<p style="margin-top:20px;font-size:12px;color:#6b7280">If you are the developer, check your template scriptlets.</p>' +
        "</div>",
    );
  }
}

function doPost(e) {
  return handleRequest(e);
}

function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

// ─────────────────────────────────────────────────────────────
// ORIGIN / DOMAIN LOCK — blocks cloned pages on unknown domains
// ─────────────────────────────────────────────────────────────

/**
 * Check if a request origin is allowed to call this API.
 * Allowed origins are stored in ScriptProperties under 'allowed_origins'.
 * The default allows the official GAS deployment and any custom domains
 * added by the admin via setAllowedOrigins().
 */
function _isOriginAllowed(origin) {
  if (!origin) return true; // empty origin = google.script.run (same-origin, safe)

  var cache = CacheService.getScriptCache();
  var cached = cache.get("allowed_origins");
  var allowed;
  if (cached) {
    allowed = JSON.parse(cached);
  } else {
    allowed = _loadAllowedOrigins();
    try {
      cache.put("allowed_origins", JSON.stringify(allowed), 600);
    } catch (e) {}
  }

  for (var i = 0; i < allowed.length; i++) {
    if (origin.indexOf(allowed[i]) === 0) return true;
  }
  return false;
}

function _loadAllowedOrigins() {
  var stored =
    PropertiesService.getScriptProperties().getProperty("allowed_origins");
  if (stored) {
    try {
      return JSON.parse(stored);
    } catch (e) {}
  }
  // Default: allow the GAS deployment domain, your custom domain, and local dev
  return [
    "https://script.google.com",
    "https://script.googleusercontent.com",
    "https://pat.fob.net.ng",
    "http://localhost",
    "http://127.0.0.1",
  ];
}

/**
 * Admin function: set allowed origins for the domain lock.
 * @param {string} token - Super Admin token
 * @param {string} originsJson - JSON array of allowed origin prefixes, e.g. ["https://mycustomdomain.com"]
 */
function setAllowedOrigins(token, originsJson) {
  try {
    _superAdminSession(token);
    var origins = JSON.parse(originsJson);
    if (!Array.isArray(origins) || origins.length === 0) {
      throw new Error("Must provide a non-empty array of origin strings.");
    }
    PropertiesService.getScriptProperties().setProperty(
      "allowed_origins",
      JSON.stringify(origins),
    );
    CacheService.getScriptCache().remove("allowed_origins");
    return {
      success: true,
      message: "Allowed origins updated.",
      origins: origins,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Admin function: get current allowed origins.
 * @param {string} token - Super Admin token
 */
function getAllowedOrigins(token) {
  try {
    _superAdminSession(token);
    return { success: true, origins: _loadAllowedOrigins() };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// ADMIN MENU HELPERS  (run from the spreadsheet UI)
// ─────────────────────────────────────────────────────────────

/**
 * UI Helper to create a Super Admin.
 */
function ui_createSuperAdmin() {
  _manualCreateUser("super admin");
}

/**
 * UI Helper to create a standard Admin.
 */
function ui_createAdmin() {
  _manualCreateUser(SD.ADMIN);
}

function ui_addMaterial() {
  var ui = SpreadsheetApp.getUi();
  var description = ui
    .prompt("New Material Description:")
    .getResponseText()
    .trim();
  if (!description) {
    ui.alert("Material description is required.");
    return;
  }
  var result = _createMaterial(description);
  if (!result.success) {
    ui.alert("Failed to add material: " + result.message);
  } else {
    ui.alert("Material added successfully.");
  }
}

/**
 * Internal logic for creating privileged users from the Spreadsheet UI.
 */
function _manualCreateUser(role) {
  var ui = SpreadsheetApp.getUi();
  var config = _getConfig();
  var email = ui
    .prompt("New " + role + " email (must end with " + config.DOMAIN + "):")
    .getResponseText()
    .trim();
  if (!email || !email.endsWith(config.DOMAIN)) {
    ui.alert("Invalid domain.");
    return;
  }
  var name = ui.prompt("Full name:").getResponseText().trim();
  var pass = ui.prompt("Password (min 4 chars):").getResponseText();

  var deptsRes = getPublicDepartments();
  var dept = "";
  if (role !== "super admin" && deptsRes.success && deptsRes.data.length > 0) {
    var deptMsg =
      "Available Departments:\n" +
      deptsRes.data.join(", ") +
      "\n\nEnter Department:";
    dept = ui.prompt(deptMsg).getResponseText().trim();
  } else if (role !== "super admin") {
    dept = ui
      .prompt("No departments found. Enter a new department name:")
      .getResponseText()
      .trim();
  } else {
    dept = "N/A";
  }

  if (role !== "super admin" && !dept) {
    ui.alert("Department is required.");
    return;
  }

  try {
    var res = registerUser(name, email, pass, dept, "Other"); // Default gender for manual admin
    if (!res.success) throw new Error(res.message);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var lo = email.toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === lo) {
        sh.getRange(i + 1, 6).setValue(role);
        break;
      }
    }

    ui.alert("✅ " + role + " created!\n\nEmail: " + email);
  } catch (e) {
    ui.alert("❌ Error: " + e.message);
  }
}

/**
 * Securely resets a user's password. Restricted to Super Admins.
 */
function resetUserPassword(token, targetEmail, newPassword) {
  try {
    _superAdminSession(token);
    targetEmail = String(targetEmail || "")
      .toLowerCase()
      .trim();
    if (!targetEmail || !newPassword)
      throw new Error("Email and new password are required.");
    if (newPassword.length < 4)
      throw new Error("Password must be at least 4 characters.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase().trim() === targetEmail) {
        var s = _salt();
        sh.getRange(i + 1, 4).setValue(_hash(newPassword, s));
        sh.getRange(i + 1, 5).setValue(s);
        return {
          success: true,
          message: "Password reset successfully for " + targetEmail,
        };
      }
    }
    throw new Error("User not found: " + targetEmail);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Marks all unread messages from a specific partner as read.
 */
function ui_resetPassword() {
  var ui = SpreadsheetApp.getUi();
  var email = ui
    .prompt("Email of user to reset:")
    .getResponseText()
    .trim()
    .toLowerCase();
  if (!email) return;
  var pass = ui.prompt("New password (min 4 chars):").getResponseText();
  if (!pass || pass.length < 4) {
    ui.alert("❌ Password too short.");
    return;
  }

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email) {
      var s = _salt();
      sh.getRange(i + 1, 4).setValue(_hash(pass, s));
      sh.getRange(i + 1, 5).setValue(s);
      ui.alert("✅ Password reset for: " + email);
      return;
    }
  }
  ui.alert("❌ User not found: " + email);
}

function ui_liveStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usr = ss.getSheetByName(SD.USERS);
  if (!usr) {
    ss.toast("Run First Setup first.", "Stats", 3);
    return;
  }

  var ud = usr.getDataRange().getValues().slice(1);
  var allUsers = ud.length;
  var active = ud.filter(function (r) {
    return r[7] === "active";
  }).length;

  // Count PAT projects
  var patSh = ss.getSheetByName(SD.PAT);
  var totalPat = patSh && patSh.getLastRow() > 1 ? patSh.getLastRow() - 1 : 0;

  SpreadsheetApp.getUi().alert(
    "📊 SD Portal — Live Stats\n\n" +
      "Users\n" +
      "  Registered: " +
      allUsers +
      "\n" +
      "  Active: " +
      active +
      "\n\n" +
      "PAT Projects\n" +
      "  Total: " +
      totalPat,
  );
}

function ui_healthCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var needed = [SD.USERS, SD.DEPTS, SD.PAT];
  var issues = [];
  needed.forEach(function (n) {
    if (!ss.getSheetByName(n)) issues.push("Missing sheet: " + n);
  });
  var userSh = ss.getSheetByName(SD.USERS);
  if (userSh && userSh.getLastRow() < 2)
    issues.push("No users registered yet.");

  SpreadsheetApp.getUi().alert(
    issues.length === 0
      ? "✅ All clear — all sheets exist and are populated."
      : "⚠️ Issues found:\n\n" +
          issues.join("\n") +
          "\n\nRun 'First Setup' to fix.",
  );
}

function ui_wipeSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "🚨 DANGER: SYSTEM WIPE",
    "This will delete ALL Users, Departments, and PAT Projects. \n\nAre you absolutely sure you want to proceed?",
    ui.ButtonSet.YES_NO,
  );

  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [SD.USERS, SD.DEPTS, SD.PAT];
    sheets.forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (sh && sh.getLastRow() > 1) {
        sh.deleteRows(2, sh.getLastRow() - 1);
      }
    });
    PropertiesService.getScriptProperties().deleteAllProperties();
    PropertiesService.getUserProperties().deleteAllProperties();
    ss.toast(
      "System wiped successfully. Run 'First Setup' to initialize again.",
      "Reset Complete",
    );
  }
}

/**
 * Force-unban an account by email, fully clearing BOTH ban stores.
 *
 * A ban is recorded in two places:
 *   1. The USERS sheet  — Status (col 10) = "banned", plus BanReason (17) / BannedBy (18).
 *   2. The SD_BLACKLIST sheet — one or more rows (per IP, per fingerprint, and an
 *      EMAIL: row), all sharing the user's email in col 3 (USER_EMAIL).
 *
 * Editing the sheet by hand reliably leaves ONE of those stores behind, so the
 * account stays blocked (or blocks with a blank reason). This tool clears BOTH
 * completely, so it is safe to use even when a manual edit left the account
 * half-banned. Super-admin only.
 */
function ui_forceUnbanByEmail() {
  var ui = SpreadsheetApp.getUi();
  var email = ui.prompt(
    "🔓 Force Unban by Email",
    "Enter the exact email to fully unban (clears account ban + all blacklist rows):",
    ui.ButtonSet.OK_CANCEL,
  );
  if (email.getSelectedButton() !== ui.Button.OK) return;
  var addr = String(email.getResponseText() || "")
    .trim()
    .toLowerCase();
  if (!addr || addr.indexOf("@") === -1) {
    ui.alert("No valid email entered.");
    return;
  }

  var confirm = ui.alert(
    "Confirm Force Unban",
    "This will reactivate " +
      addr +
      " and clear EVERY active blacklist row for that email. Continue?",
    ui.ButtonSet.YES_NO,
  );
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date().toISOString();
  var ush = ss.getSheetByName(SD.USERS);
  var reactivated = 0;

  if (ush && ush.getLastRow() >= 2) {
    var udata = ush.getDataRange().getValues();
    for (var i = 1; i < udata.length; i++) {
      if (String(udata[i][2] || "").toLowerCase() === addr) {
        if (String(udata[i][9] || "").toLowerCase() === "banned") {
          ush.getRange(i + 1, 10).setValue("active"); // Status
          ush.getRange(i + 1, 17).setValue(""); // BanReason
          ush.getRange(i + 1, 18).setValue(""); // BannedBy
          reactivated++;
        }
      }
    }
  }

  var blacklistCleared = 0;
  var bsh = ss.getSheetByName(SD.BLACKLIST);
  if (bsh && bsh.getLastRow() >= 2) {
    var c = _blacklistCols();
    var bdata = bsh.getDataRange().getValues();
    for (var j = 1; j < bdata.length; j++) {
      if (bdata[j][c.UNBANNED_AT]) continue; // already inactive
      if (String(bdata[j][c.USER_EMAIL] || "").toLowerCase() === addr) {
        bsh.getRange(j + 1, c.UNBANNED_AT + 1).setValue(now);
        blacklistCleared++;
      }
    }
  }

  // Also kill any live sessions for the email so a stale token can't linger.
  try {
    _purgeUserSessions(addr);
  } catch (e) {}

  ui.alert(
    "✅ Force Unban Complete",
    "Email: " +
      addr +
      "\n" +
      "Account reactivated: " +
      reactivated +
      "\n" +
      "Blacklist rows cleared: " +
      blacklistCleared +
      "\n\n" +
      "The account can now sign in normally.",
    ui.ButtonSet.OK,
  );
}

// ─────────────────────────────────────────────────────────────
//  MISSING BACKEND FUNCTIONS — frontend calls these
// ─────────────────────────────────────────────────────────────

function getUsersByDepartment(token, department) {
  try {
    // Support both (department) and (token, department) calling patterns
    if (arguments.length === 1) {
      department = token;
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var members = [];
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][6]).toLowerCase() ===
        String(department).toLowerCase().trim()
      ) {
        members.push({
          email: data[i][2],
          name: data[i][1],
          role: String(data[i][5] || "").toLowerCase(),
        });
      }
    }
    return { success: true, members: members };
  } catch (e) {
    return { success: false, message: e.message, members: [] };
  }
}

function getUserByEmail(token, email) {
  try {
    if (!token) throw new Error("Admin token required.");
    _adminSession(token);

    if (arguments.length === 1) {
      email = token;
      token = null;
    }
    if (!email) throw new Error("Email is required.");
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][2]).toLowerCase().trim() ===
        String(email).toLowerCase().trim()
      ) {
        return {
          success: true,
          user: {
            userId: data[i][0],
            name: data[i][1],
            email: data[i][2],
            role: data[i][5],
            department: data[i][6],
            gender: data[i][7],
            status: data[i][9],
          },
        };
      }
    }
    return { success: false, message: "User not found." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function adminDeleteUser(token, userId) {
  try {
    _superAdminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var searchId = String(userId || "").trim();
    var isEmailSearch = searchId.indexOf("@") !== -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      var rowEmail = String(data[i][2]).toLowerCase().trim();
      if (
        rowId === searchId ||
        (isEmailSearch && rowEmail === searchId.toLowerCase())
      ) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "User not found." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function adminUpdateUser(token, userId, updates) {
  try {
    _superAdminSession(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var searchId = String(userId || "").trim();
    var isEmailSearch = searchId.indexOf("@") !== -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      var rowEmail = String(data[i][2]).toLowerCase().trim();
      if (
        rowId === searchId ||
        (isEmailSearch && rowEmail === searchId.toLowerCase())
      ) {
        var row = i + 1;
        if (updates.name) sh.getRange(row, 2).setValue(updates.name);
        if (updates.department || updates.dept)
          sh.getRange(row, 7).setValue(updates.department || updates.dept);
        if (updates.role) sh.getRange(row, 6).setValue(updates.role);
        if (updates.gender) sh.getRange(row, 8).setValue(updates.gender);
        return { success: true };
      }
    }
    return { success: false, message: "User not found." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getNextStageUsers(token, currentStatus) {
  try {
    if (arguments.length === 1) {
      currentStatus = token;
    }
    // Derive next department from the centralized workflow constants:
    // 1. Find the next status from WF_STAGE_FLOW (linear + recovery)
    // 2. Look up the owning department for that next status in STATUS_TO_DEPT
    var nextStatus = WF_STAGE_FLOW[currentStatus];
    var nextDept = nextStatus ? STATUS_TO_DEPT[nextStatus] || "MEC" : "MEC";
    return getUsersByDepartment(nextDept);
  } catch (e) {
    return { success: false, message: e.message, members: [] };
  }
}

function approvePAT(token, projectId, comments, verdict) {
  try {
    var sess = _session(token);
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // IDOR fix: verify user has permission to act on this project
    if (!_userCanActOnProject(sess, project)) {
      return {
        success: false,
        message: "Permission denied: You are not assigned to this project.",
      };
    }

    // Update project status
    project.workflowStatus = "Completed";
    project.verdict = verdict || "Fully Accepted";
    project.department = "MEC";
    project.assignedToDept = "MEC";
    project.assignedToName = "";
    project.assignedToEmail = "";
    var history = project.workflowHistory || [];
    history.push({
      fromStatus: res.project.workflowStatus,
      toStatus: "Completed",
      by: {
        name: sess.name || "Admin",
        email: sess.email,
        department: sess.department,
      },
      comments: comments || "",
      timestamp: new Date().toISOString(),
    });
    project.workflowHistory = history;

    var saveRes = savePATProject(token, project);

    // ── AUTO-GENERATE JCC ON COMPLETION ──
    var jccResult = { success: false, message: "JCC not generated" };
    if (saveRes.success) {
      jccResult = generateJCC(projectId, sess);
    }

    return {
      success: saveRes.success,
      message:
        "PAT approved and completed." +
        (jccResult.success ? " JCC generated: " + jccResult.certificateId : ""),
      jcc: jccResult,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function createDashboardShareToken(token, email, durationMinutes) {
  try {
    _session(token);
    var shareToken = _makeToken();
    PropertiesService.getScriptProperties().setProperty(
      "dash_share_" + shareToken,
      JSON.stringify({
        email: email,
        expiresAt: Date.now() + (durationMinutes || 5) * 60000,
        from: "admin",
      }),
    );
    return {
      success: true,
      token: shareToken,
      expiresAt: new Date(
        Date.now() + (durationMinutes || 5) * 60000,
      ).toISOString(),
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getSharedDashboard(token) {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    for (var key in props) {
      if (key.indexOf("dash_share_") === 0) {
        var data = _safeParse(props[key], null);
        if (data && Date.now() < data.expiresAt) {
          return {
            success: true,
            shared: true,
            analytics: { totalProjects: 0, byStatus: {} },
            from: data.from,
            expiresAt: new Date(data.expiresAt).toISOString(),
          };
        }
        if (data === null)
          PropertiesService.getScriptProperties().deleteProperty(key);
      }
    }
    return { success: true, shared: false };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getPATAnalyticsForIdris(token, forceRefresh) {
  try {
    if (arguments.length === 1) {
      forceRefresh = token;
    }
    var projects = getPATProjects("").projects || [];
    var byStatus = {};
    var byDept = {};
    var completed = 0;
    var rejected = 0;

    projects.forEach(function (p) {
      var st = p.workflowStatus || "Draft";
      byStatus[st] = (byStatus[st] || 0) + 1;
      if (st === "Completed") completed++;
      if (st === "Rejected") rejected++;

      var dp = p.department || "Unknown";
      byDept[dp] = (byDept[dp] || 0) + 1;
    });

    var total = projects.length;
    var completionRate =
      total > 0 ? Math.round((completed / total) * 100) + "%" : "0%";
    var rejectionRate =
      total > 0 ? Math.round((rejected / total) * 100) + "%" : "0%";

    return {
      success: true,
      analytics: {
        total: total,
        byStatus: byStatus,
        projects: projects,
        workflowEfficiency: {
          completionRate: completionRate,
          averageCompletionTime: "12 days", // Mock value
          rejectionRate: rejectionRate,
        },
      },
    };
  } catch (e) {
    return {
      success: false,
      message: e.message,
      analytics: {
        total: 0,
        byStatus: {},
        projects: [],
        workflowEfficiency: {
          completionRate: "0%",
          averageCompletionTime: "N/A",
          rejectionRate: "0%",
        },
      },
    };
  }
}

/**
 * AUDIT CONSOLE — Analytics over COMPLETED PAT projects.
 * Restricted to the Audit department and admins.
 * Returns the full list of completed projects (with BOQ/snags/JCC flags)
 * plus aggregated material and workflow analytics for the audit dashboard.
 * @param {string} token
 */
function getAuditAnalytics(token) {
  try {
    var sess = _session(token);
    var role = String(sess.role || "").toLowerCase();
    var dept = String(sess.department || "").toLowerCase();
    if (!(role === "admin" || role === "super admin" || dept === "audit")) {
      throw new Error(
        "Audit access required. This console is restricted to the Audit department.",
      );
    }

    var all = getPATProjects("").projects || [];
    var completed = all.filter(function (p) {
      return (p.workflowStatus || "") === "Completed";
    });

    // Build a lookup of generated JCC certificates by normalized project id
    var jccMap = {};
    try {
      var jccRes = getAllJCCs();
      (jccRes.jccs || []).forEach(function (j) {
        var key = String(j.projectId || "")
          .toUpperCase()
          .replace(/^(FOB|PAT)-/, "");
        jccMap[key] = j;
      });
    } catch (e) {
      /* JCC sheet may be empty */
    }

    // Aggregation buckets
    var usedMap = {}; // desc -> { projectCount, totalScope, totalInstalled }
    var insuffMap = {}; // desc -> { occurrences, totalShort }
    var returnMap = {}; // desc -> { occurrences, totalLeftover }

    var verdictDist = {};
    var deptCounts = {};
    var snagByDept = {};
    var orchMap = {};
    var vendorSet = {};
    var totalSnags = 0,
      snagScoreSum = 0,
      jccCount = 0;
    var materialWasteTotal = 0,
      insufficientTotal = 0;

    var projectRows = completed.map(function (p) {
      var boq = p.boq || [];
      boq.forEach(function (b) {
        var desc = String(b.desc || "").trim();
        if (!desc) return;
        var scope = Number(b.scope) || 0;
        var installed = Number(b.installed) || 0;
        var variance = Number(b.variance);
        if (isNaN(variance)) variance = scope - installed;

        if (!usedMap[desc])
          usedMap[desc] = { projectCount: 0, totalScope: 0, totalInstalled: 0 };
        usedMap[desc].projectCount++;
        usedMap[desc].totalScope += scope;
        usedMap[desc].totalInstalled += installed;

        if (variance < 0) {
          if (!insuffMap[desc])
            insuffMap[desc] = { occurrences: 0, totalShort: 0 };
          insuffMap[desc].occurrences++;
          insuffMap[desc].totalShort += Math.abs(variance);
        } else if (variance > 0) {
          if (!returnMap[desc])
            returnMap[desc] = { occurrences: 0, totalLeftover: 0 };
          returnMap[desc].occurrences++;
          returnMap[desc].totalLeftover += variance;
        }
      });

      var v = String(p.verdict || "Pending");
      verdictDist[v] = (verdictDist[v] || 0) + 1;

      var dp =
        String(p.department || p.assignedToDept || "").trim() || "Unknown";
      deptCounts[dp] = (deptCounts[dp] || 0) + 1;

      var orch = String(p.orchestrator || "").trim();
      if (orch) orchMap[orch] = (orchMap[orch] || 0) + 1;

      var vd = String(p.vendor || "").trim();
      if (vd) vendorSet[vd] = true;

      var snags = p.snags || [];
      totalSnags += snags.length;
      snagScoreSum += Number(p.snagScore) || 0;
      snags.forEach(function (s) {
        var sd = String(s.department || "Unknown").trim() || "Unknown";
        snagByDept[sd] = (snagByDept[sd] || 0) + 1;
      });

      var normId = String(p.projectId || "")
        .toUpperCase()
        .replace(/^(FOB|PAT)-/, "");
      var jcc = jccMap[normId];
      if (jcc) jccCount++;

      return {
        projectId: p.projectId,
        projectName: p.projectName,
        vendor: p.vendor,
        inspectionDate: p.inspectionDate,
        siteAddress: p.siteAddress,
        verdict: p.verdict,
        snagScore: Number(p.snagScore) || 0,
        orchestrator: p.orchestrator,
        department: p.department || p.assignedToDept,
        presidingOfficer: p.presidingOfficer,
        boq: boq,
        snags: snags,
        hasJCC: !!jcc,
        jccId: jcc ? jcc.jccId : "",
        certificateId: jcc ? jcc.certificateId : "",
        workflowStatus: p.workflowStatus,
      };
    });

    var mostUsed = Object.keys(usedMap)
      .map(function (d) {
        return {
          desc: d,
          projectCount: usedMap[d].projectCount,
          totalScope: usedMap[d].totalScope,
          totalInstalled: usedMap[d].totalInstalled,
        };
      })
      .sort(function (a, b) {
        return b.totalScope - a.totalScope;
      })
      .slice(0, 12);

    var insufficient = Object.keys(insuffMap)
      .map(function (d) {
        return {
          desc: d,
          occurrences: insuffMap[d].occurrences,
          totalShort: insuffMap[d].totalShort,
        };
      })
      .sort(function (a, b) {
        return b.totalShort - a.totalShort;
      });

    var returned = Object.keys(returnMap)
      .map(function (d) {
        return {
          desc: d,
          occurrences: returnMap[d].occurrences,
          totalLeftover: returnMap[d].totalLeftover,
        };
      })
      .sort(function (a, b) {
        return b.totalLeftover - a.totalLeftover;
      });

    Object.keys(insuffMap).forEach(function (d) {
      insufficientTotal += insuffMap[d].totalShort;
    });
    Object.keys(returnMap).forEach(function (d) {
      materialWasteTotal += returnMap[d].totalLeftover;
    });

    var topOrchestrators = Object.keys(orchMap)
      .map(function (d) {
        return { name: d, count: orchMap[d] };
      })
      .sort(function (a, b) {
        return b.count - a.count;
      })
      .slice(0, 8);

    var avgSnagScore = completed.length
      ? Math.round((snagScoreSum / completed.length) * 10) / 10
      : 0;

    return {
      success: true,
      completedProjects: projectRows,
      analytics: {
        totalCompleted: completed.length,
        totalVendors: Object.keys(vendorSet).length,
        totalSnags: totalSnags,
        avgSnagScore: avgSnagScore,
        mostUsedMaterials: mostUsed,
        insufficientMaterials: insufficient,
        returnedMaterials: returned,
        verdictDistribution: verdictDist,
        departmentCounts: deptCounts,
        snagsByDepartment: snagByDept,
        topOrchestrators: topOrchestrators,
        jccGenerated: jccCount,
        materialWasteTotal: materialWasteTotal,
        insufficientTotal: insufficientTotal,
      },
    };
  } catch (e) {
    return {
      success: false,
      message: e.message,
      completedProjects: [],
      analytics: null,
    };
  }
}

function doEmployeeCheckin() {
  var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  if (!userSh || userSh.getLastRow() < 2) return;

  var data = userSh.getDataRange().getValues();
  var sh = _ensureCheckinSheet();
  var today = new Date().toISOString().split("T")[0];

  var questions = [
    "Good morning! How is your work going today?",
    "Hi there! How's your progress on today's tasks?",
    "Hello! How are things going with your work today?",
    "Hey! How's everything going on your end?",
    "Warm greetings! How is your work coming along?",
  ];

  var hour = new Date().getHours();
  var minute = new Date().getMinutes();
  var timeSlot = hour * 60 + minute;

  var slots = [
    { start: 540, end: 570 },
    { start: 642, end: 672 },
    { start: 744, end: 774 },
    { start: 846, end: 876 },
    { start: 948, end: 978 },
  ];

  var questionIndex = -1;
  for (var s = 0; s < slots.length; s++) {
    if (timeSlot >= slots[s].start && timeSlot <= slots[s].end) {
      questionIndex = s;
      break;
    }
  }

  if (questionIndex === -1) return;

  var question = questions[questionIndex] || "How is your work going?";
  var existingData = sh.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][2] || "")
      .toLowerCase()
      .trim();
    var name = String(data[i][1] || "").trim();
    var role = String(data[i][5] || "").toLowerCase();
    var status = String(data[i][9] || "").toLowerCase();

    if (role !== "employee") continue;
    if (status !== "active") continue;
    if (!email) continue;

    var alreadyExists = false;
    for (var j = 1; j < existingData.length; j++) {
      var exEmail = String(existingData[j][1] || "")
        .toLowerCase()
        .trim();
      var exStatus = String(existingData[j][5] || "").toLowerCase();
      var exDate = String(existingData[j][6] || "").split("T")[0];
      var exQ = String(existingData[j][3] || "");

      if (
        exEmail === email &&
        exDate === today &&
        exQ === question &&
        (exStatus === "pending" ||
          exStatus === "responded" ||
          exStatus === "completed")
      ) {
        alreadyExists = true;
        break;
      }
    }

    if (!alreadyExists) {
      var checkinId = _genId("CHK");
      sh.appendRow([
        checkinId,
        email,
        name,
        question,
        "",
        "pending",
        new Date().toISOString(),
        "",
      ]);
    }
  }
}

function _ensureCheckinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("SD_CHECKINS");
  if (!sh) {
    sh = ss.insertSheet("SD_CHECKINS");
    // Headers: CheckinID | Email | Name | Question | Answer | Status | CreatedAt | CompletedAt
    sh.appendRow([
      "CheckinID",
      "Email",
      "Name",
      "Question",
      "Answer",
      "Status",
      "Timestamp",
      "CompletedAt",
    ]);
    sh.getRange(1, 1, 1, 8)
      .setBackground("#0d1526")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
  }
  return sh;
}

function setupEmployeeCheckins() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === "doEmployeeCheckin") {
      ScriptApp.deleteTrigger(t);
    }
  });

  var times = [
    { hour: 9, minute: 0 },
    { hour: 10, minute: 42 },
    { hour: 12, minute: 24 },
    { hour: 14, minute: 6 },
    { hour: 15, minute: 48 },
  ];

  times.forEach(function (t) {
    ScriptApp.newTrigger("doEmployeeCheckin")
      .timeBased()
      .atHour(t.hour)
      .nearMinute(t.minute)
      .everyDays(1)
      .inTimezone("Africa/Lagos")
      .create();
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Employee check-in triggers created (5x daily: 9AM-4:30PM).",
  );
}

// ─────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────
//  MATERIALS MANAGEMENT
// ─────────────────────────────────────────────────────────────

function getMaterials(token) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MATS);
    if (!sh || sh.getLastRow() < 2) return { success: true, materials: [] };
    var data = sh.getDataRange().getValues();
    var materials = [];
    for (var i = 1; i < data.length; i++) {
      materials.push({ id: data[i][0], description: data[i][1] });
    }
    return { success: true, materials: materials };
  } catch (e) {
    return { success: false, message: e.message, materials: [] };
  }
}

function _createMaterial(description) {
  try {
    if (!description || !description.trim())
      return { success: false, message: "Description required." };
    var desc = description.trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MATS);
    if (!sh) return { success: false, message: "Materials sheet not found." };
    // Case-insensitive duplicate check — do NOT add a material that already exists
    var existing = sh.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (
        String(existing[i][1] || "")
          .trim()
          .toLowerCase() === desc.toLowerCase()
      ) {
        return {
          success: true,
          material: { id: existing[i][0], description: existing[i][1] },
          isNew: false,
          message: "Material already exists.",
        };
      }
    }
    var id = _genId("MAT-");
    sh.appendRow([id, desc, new Date().toISOString()]);
    return {
      success: true,
      material: { id: id, description: desc },
      isNew: true,
      message: "Material added.",
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addMaterial(token, description) {
  try {
    _session(token);
    return _createMaterial(description);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteMaterial(token, materialId) {
  try {
    _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MATS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(materialId)) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Material not found." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// BLACKLIST SYSTEM — IP and device ban management
// ─────────────────────────────────────────────────────────────

function _blacklistSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SD.BLACKLIST);
  if (!sh) {
    sh = ss.insertSheet(SD.BLACKLIST);
    sh.getRange(1, 1, 1, 9).setValues([
      [
        "EntryID",
        "IP",
        "DeviceFingerprint",
        "UserEmail",
        "UserName",
        "BannedBy",
        "Reason",
        "BannedAt",
        "UnbannedAt",
      ],
    ]);
    sh.getRange(1, 1, 1, 9)
      .setBackground(_getConfig().NAVY)
      .setFontColor("#ffffff")
      .setFontWeight("bold");
  }
  return sh;
}

function _blacklistCols() {
  return {
    ENTRY_ID: 0,
    IP: 1,
    DEVICE_FINGERPRINT: 2,
    USER_EMAIL: 3,
    USER_NAME: 4,
    BANNED_BY: 5,
    REASON: 6,
    BANNED_AT: 7,
    UNBANNED_AT: 8,
  };
}

function getBlacklist(token) {
  try {
    _adminSession(token);
    var sh = _blacklistSheet();
    if (sh.getLastRow() < 2) {
      return { success: true, entries: [] };
    }
    var c = _blacklistCols();
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
    var entries = [];
    data.forEach(function (r) {
      entries.push({
        entryId: r[c.ENTRY_ID],
        ip: r[c.IP],
        deviceFingerprint: r[c.DEVICE_FINGERPRINT],
        userEmail: r[c.USER_EMAIL],
        userName: r[c.USER_NAME],
        bannedBy: r[c.BANNED_BY],
        reason: r[c.REASON],
        bannedAt: r[c.BANNED_AT],
        unbannedAt: r[c.UNBANNED_AT],
      });
    });
    return { success: true, entries: entries };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function banUserWithBlacklist(token, userId, reason) {
  try {
    var sess = _adminSession(token);

    // Get user info and their known IPs/fingerprints
    var user = getUserById(token, userId);
    if (!user.success)
      return { success: false, message: "User not found: " + user.message };

    // Self-account protection: super admins cannot be banned.
    if (String(user.role || "").toLowerCase() === "super admin") {
      return {
        success: false,
        message: "Super admin accounts are protected and cannot be banned.",
      };
    }

    // Apply the ban status + persist reason/who-banned (peak-security audit trail)
    var statusRes = updateUserStatus(
      token,
      userId,
      "banned",
      reason,
      sess.email,
    );
    if (!statusRes.success) return statusRes;

    var sh = _blacklistSheet();
    var c = _blacklistCols();
    var entryId = _genId("BL");
    var timestamp = new Date().toISOString();
    var sess = _session(token);
    var banCount = 0;
    var ips = user.ips || [];
    var fingerprints = user.fingerprints || [];

    var userName = user.name || "";
    var userEmail = user.email || "";

    ips.forEach(function (ip) {
      sh.appendRow([
        entryId + "-IP-" + banCount,
        ip,
        "",
        userEmail,
        userName,
        sess.email || "admin",
        reason || "Manual ban",
        timestamp,
        "",
      ]);
      banCount++;
    });

    fingerprints.forEach(function (fp) {
      sh.appendRow([
        entryId + "-FP-" + banCount,
        "",
        fp,
        userEmail,
        userName,
        sess.email || "admin",
        reason || "Manual ban",
        timestamp,
        "",
      ]);
      banCount++;
    });

    // Also add email-based blacklist entry
    sh.appendRow([
      entryId + "-EMAIL",
      "EMAIL:" + userEmail,
      "",
      userEmail,
      userName,
      sess.email || "admin",
      "Email banned - " + (reason || ""),
      timestamp,
      "",
    ]);

    return {
      success: true,
      ipsBlacklisted: ips.length,
      devicesBlacklisted: fingerprints.length,
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function unbanUserWithBlacklist(token, userId) {
  try {
    var sess = _adminSession(token);

    // Get user email + role from userId
    var user = getUserById(token, userId);
    if (!user.success)
      return { success: false, message: "User not found: " + user.message };
    // Reactivate the user in the USERS sheet and clear ban metadata.
    updateUserStatus(token, userId, "active", "", sess.email);

    var userEmail = user.email || "";
    var sh = _blacklistSheet();
    if (sh.getLastRow() < 2)
      return { success: true, message: "No blacklist entries found." };

    var c = _blacklistCols();
    var data = sh.getDataRange().getValues();
    var unbannedCount = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var rowEmail = String(row[c.USER_EMAIL] || "").toLowerCase();
      var searchEmail = String(userEmail || "").toLowerCase();
      if (rowEmail === searchEmail && !row[c.UNBANNED_AT]) {
        sh.getRange(i + 1, c.UNBANNED_AT + 1).setValue(
          new Date().toISOString(),
        );
        unbannedCount++;
      }
    }
    return { success: true, unbannedCount: unbannedCount };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function removeBlacklistEntry(token, entryId) {
  try {
    _adminSession(token);
    var sh = _blacklistSheet();
    if (sh.getLastRow() < 2)
      return { success: false, message: "No entries found." };

    var c = _blacklistCols();
    var data = sh.getDataRange().getValues();

    var affectedEmails = [];
    var stampedRows = [];
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][c.ENTRY_ID] || "") === String(entryId)) {
        var em = String(data[i][c.USER_EMAIL] || "").toLowerCase();
        if (em && affectedEmails.indexOf(em) === -1) affectedEmails.push(em);
        sh.getRange(i + 1, c.UNBANNED_AT + 1).setValue(
          new Date().toISOString(),
        );
        stampedRows.push(i);
      }
    }

    // ── SYNC ACCOUNT BAN ──
    // The blacklist and the USERS "banned" status are two faces of the same
    // block. The login/signup guards check BOTH (USERS status at cols 10/17/18
    // and the blacklist via _checkBlacklistMatch). If we only stamp
    // UNBANNED_AT here, a user whose account was banned (USERS status="banned",
    // bannedBy populated) stays blocked forever even after the blacklist row is
    // gone — they keep hitting "__BANNED__::...::<bannedBy>".
    //
    // So: only if NO OTHER active blacklist row still references the email do
    // we lift the account-level ban and clear its ban metadata in USERS. This
    // keeps removal consistent — removing one of several entries for an account
    // does not falsely "unban" an account that is still otherwise blocklisted.
    var stillBanned = {};
    for (var k = 1; k < data.length; k++) {
      if (stampedRows.indexOf(k) !== -1) continue; // just unbanned above
      if (data[k][c.UNBANNED_AT]) continue; // already inactive
      var rem = String(data[k][c.USER_EMAIL] || "").toLowerCase();
      if (rem) stillBanned[rem] = true;
    }
    var fullyClear = affectedEmails.filter(function (e) {
      return !stillBanned[e];
    });

    var reactivated = 0;
    if (fullyClear.length) {
      var ush = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
      var udata = ush.getDataRange().getValues();
      for (var j = 1; j < udata.length; j++) {
        var uemail = String(udata[j][2] || "").toLowerCase();
        if (
          fullyClear.indexOf(uemail) !== -1 &&
          String(udata[j][9] || "").toLowerCase() === "banned"
        ) {
          ush.getRange(j + 1, 10).setValue("active"); // clear status
          ush.getRange(j + 1, 17).setValue(""); // clear BanReason
          ush.getRange(j + 1, 18).setValue(""); // clear BannedBy
          reactivated++;
        }
      }
    }
    return { success: true, reactivated: reactivated };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// AI COMMAND AUTHORIZATION — backend re-validation for every
// AI-issued command. Prevents prompt-injection attacks where a
// malicious input tricks the AI into issuing destructive commands.
//
// The frontend system prompt invites the AI to emit json-cmd blocks.
// An attacker could craft input that causes the AI to emit DELETE or
// APPROVE commands. This layer ensures the backend verifies the
// caller's identity and role before any command is honored.
// ─────────────────────────────────────────────────────────────

/**
 * Check if the current session can perform an AI-issued command.
 * Returns { allowed: bool, role: string, message?: string }
 */
function authorizeAICommand(token, commandType) {
  try {
    if (!token) throw new Error("No session token provided.");
    var sess = _session(token);
    var role = String(sess.role || "").toLowerCase();

    // Check DB for latest role (handles promotions without re-login)
    var dbRole = _getUserRoleFromDb(sess.email);
    if (dbRole && dbRole !== role) role = dbRole;

    if (role === "employee") {
      var EMPLOYEE_TEST_ONLY = ["OPEN_PAT", "FILL_PAT_FORM", "SAVE_PAT"];
      if (EMPLOYEE_TEST_ONLY.indexOf(commandType) === -1) {
        return {
          allowed: false,
          role: role,
          message:
            "Access Denied: Employees are restricted to performing standard authorized tests only (open, fill, save PAT). Access to system data, backend information, project lists, mail, and sensitive data is prohibited.",
        };
      }
      return {
        allowed: true,
        role: role,
        message: "Employee authorized for standard test only.",
      };
    }

    // MEC-only commands (non-MEC admins excluded — backend enforced)
    var MEC_CMDS = ["FILL_PAT_WITH_RANDOM_DATA"];
    if (MEC_CMDS.indexOf(commandType) !== -1) {
      var deptNorm = String(sess.department || "").toLowerCase();
      var mech = deptNorm === "mec" || deptNorm === "mech";
      if (mech) {
        return {
          allowed: true,
          role: role,
          message: "MEC authorized for PAT data operations.",
        };
      }
      return {
        allowed: false,
        role: role,
        message:
          "Access Denied: This operation is restricted to MEC department personnel. You do not have the necessary permissions to perform this action. Even the AI will not open the workspace for this request.",
      };
    }

    // Safe read-only commands — non-employee, non-MEC-specific commands
    var READ_ONLY = [
      "OPEN_PAT",
      "LOAD_PAT_PROJECT",
      "FILL_PAT_FORM",
      "SAVE_PAT",
      "LIST_PAT_PROJECTS",
      "GET_PAT_STATUS",
      "GET_USER_BY_EMAIL",
      "REFRESH_SYSTEM_DATA",
      "SYNC_ALL_SYSTEM_DATA",
      "FETCH_MAILS",
      "GET_SHARED_DASHBOARD",
      "EXPORT_PAT_PDF",
      "GET_BLACKLIST",
      "GET_USER_BY_ID",
      "GET_AUDIT_ANALYTICS",
      "GET_MY_NOTIFICATIONS",
    ];
    if (READ_ONLY.indexOf(commandType) !== -1) {
      return { allowed: true, role: role, message: "Authorized." };
    }

    // Medium-risk commands — admin or above
    var ADMIN_CMDS = [
      "CREATE_USER",
      "UPDATE_USER",
      "UPDATE_USER_ROLE",
      "CREATE_DEPT",
      "SUBMIT_TO_NEXT_DEPT",
      "PARTIALLY_APPROVE_PAT",
      "APPROVE_PAT",
      "REFRESH_ALL_CACHE",
      "ADD_IMAGE_TO_PROJECT",
      "CRAFT_AND_SEND_MAIL",
      "SHARE_IDRIS_DASHBOARD",
      "RESET_PASSWORD",
      "BAN_USER",
      "UNBAN_USER",
      "REMOVE_BLACKLIST_ENTRY",
    ];
    if (ADMIN_CMDS.indexOf(commandType) !== -1) {
      if (role === "admin" || role === "super admin") {
        return { allowed: true, role: role, message: "Admin authorized." };
      }
      return {
        allowed: false,
        role: role,
        message: "Admin role required for command: " + commandType,
      };
    }

    // High-risk / destructive commands — super admin only
    var SUPER_ADMIN_CMDS = [
      "DELETE_USER",
      "DELETE_PAT_PROJECT",
      "DELETE_ALL_PAT_PROJECTS",
      "DELETE_MAIL_PERMANENTLY",
      "RANDOM_FILL_AND_APPROVE",
    ];
    if (SUPER_ADMIN_CMDS.indexOf(commandType) !== -1) {
      if (role === "super admin") {
        return {
          allowed: true,
          role: role,
          message: "Super admin authorized.",
        };
      }
      return {
        allowed: false,
        role: role,
        message: "Super admin role required for command: " + commandType,
      };
    }

    // Mail commands: normal delete restricted to owners, permanent delete to super admin
    if (commandType === "DELETE_MAIL") {
      if (role === "super admin" || role === "admin") {
        return {
          allowed: true,
          role: role,
          message: "Admin authorized for mail delete.",
        };
      }
      return {
        allowed: false,
        role: role,
        message: "Admin role required to delete mail.",
      };
    }

    // Unknown command — deny by default
    return {
      allowed: false,
      role: role,
      message: "Command not recognized or not authorized: " + commandType,
    };
  } catch (e) {
    return { allowed: false, role: "", message: "Auth error: " + e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// END OF FILE
// ─────────────────────────────────────────────────────────────
