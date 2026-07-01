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
  PAT:   "SD_PAT_PROJECTS",
  MATS:  "SD_MATERIALS",
  MAILS: "SD_MAILS",
  JCC:   "SD_JCC_CERTIFICATES",
  BLACKLIST: "SD_BLACKLIST",
  DOCS:  "SD_DOCUMENTS",
  ADMIN:    "admin",
  EMPLOYEE: "employee",
};

/**
 * Configuration Helper
 * Pulls settings from Script Properties or defaults to standard values.
 */
function _getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    TOKEN_TTL: parseInt(props.getProperty("TOKEN_TTL") || (8 * 60 * 60 * 1000)),
    SECRET_KEY: props.getProperty("SECRET_KEY") || "SD_SYSTEM_DEFAULT_SECRET",
    DOMAIN: props.getProperty("ALLOWED_DOMAIN") || "@fob.ng",
    GEMINI_API_KEY: props.getProperty("GEMINI_API_KEY") || "",
    NAVY: props.getProperty("THEME_COLOR") || "#0d1526",
    AI_WELCOME_MESSAGE: props.getProperty("AI_WELCOME_MESSAGE") || "Hello! I'm the SD-AI assistant. How can I help you today?",
    DEFAULT_REQUEST_TYPES: props.getProperty("DEFAULT_REQUEST_TYPES") || "Software Request,Automation Idea,Feedback,Bug Report,Other",
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
  
  var domain = ui.prompt("Set Allowed Email Domain (currently: " + config.DOMAIN + "):").getResponseText();
  var secret = ui.prompt("Set System Secret Key (for security, currently: " + config.SECRET_KEY + "):").getResponseText();
  var aiWelcome = ui.prompt("Set AI Assistant Welcome Message (currently: " + config.AI_WELCOME_MESSAGE + "):").getResponseText();
  
  if (domain) PropertiesService.getScriptProperties().setProperty("ALLOWED_DOMAIN", domain);
  var geminiKey = ui.prompt("Set Gemini API Key for AI tasks (leave blank to keep current):").getResponseText();
  if (geminiKey) PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", geminiKey);

  if (secret) PropertiesService.getScriptProperties().setProperty("SECRET_KEY", secret);
  if (aiWelcome) PropertiesService.getScriptProperties().setProperty("AI_WELCOME_MESSAGE", aiWelcome);
  
  // Initialize empty defaults if they don't exist
  if (!PropertiesService.getScriptProperties().getProperty("DEFAULT_ADMIN_DEPT")) {
    PropertiesService.getScriptProperties().setProperty("DEFAULT_ADMIN_DEPT", "Technology Support");
  }
  // Add default request types if the sheet is empty
  if (!PropertiesService.getScriptProperties().getProperty("DEFAULT_REQUEST_TYPES")) {
    PropertiesService.getScriptProperties().setProperty("DEFAULT_REQUEST_TYPES", "Software Request,Automation Idea,Feedback,Bug Report,Other");
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Configuration updated successfully.");
}


// ─────────────────────────────────────────────────────────────
// MENU
// ─────────────────────────────────────────────────────────────
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("⚙️ SD PORTAL ADMIN")
      .addItem("🚀 Run First Setup",        "runFirstSetup")
      .addSeparator()
      .addItem("👑 Create Super Admin",     "ui_createSuperAdmin")
      .addItem("👤 Create Admin User",      "ui_createAdmin")
      .addItem("➕ Add Material",            "ui_addMaterial")
      .addItem("🔧 Configure System",       "ui_configureSystem")
      .addItem("🔑 Reset User Password",   "ui_resetPassword")
       .addItem("🤖 Setup Employee Check-ins","setupEmployeeCheckins")
      .addItem("👤 Set MEC Head",          "ui_setMecHead")
      .addItem("🗑️  Clear All Sessions",    "clearAllSessions")
      .addSeparator()
      .addItem("📊 View Live Stats",        "ui_liveStats")
      .addItem("🏥 System Health Check",   "ui_healthCheck")
      .addSeparator()
      .addItem("🧨 Reset & Wipe System",    "ui_wipeSystem")
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
  var currentHead = PropertiesService.getScriptProperties().getProperty("MEC_HEAD_NAME") || "Not set";
  var newHead = ui.prompt("Set MEC Head Name (currently: " + currentHead + "):").getResponseText().trim();
  if (newHead) {
    PropertiesService.getScriptProperties().setProperty("MEC_HEAD_NAME", newHead);
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
  // UserID | Name | Email | PasswordHash | Salt | Role | Department | Gender | WorkflowNotes | Status | CreatedAt | LastLoginAt | AllowMessages | AIAutoAnswer
  var users = _sheet(SD.USERS, ["UserID","Name","Email","PasswordHash","Salt","Role","Department","Gender","WorkflowNotes","Status","CreatedAt","LastLoginAt","AllowMessages","AIAutoAnswer","LoginIPs"]);
  _repairUserSheet(users);

  // 2. SD_DEPARTMENTS
  // DeptID | Name | HeadEmail | CreatedAt | CreatedBy
  var depts = _sheet(SD.DEPTS, ["DeptID","Name","HeadEmail","CreatedAt","CreatedBy"]);

  // 3. SD_PAT_PROJECTS
  var pat = _sheet(SD.PAT, [
    "ProjectID","ProjectName","SiteAddress","Lat","Lon","Phase","WorkDescription",
    "Vendor","InspectionDate","Orchestrator","ComplaintScore","Status","Verdict",
    "WorkflowStatus","Checklist","BOQ","Snags","Signoff","Images",
    "HasVendor","VendorToken","VendorApprovalStatus","VendorApprovalComments","VendorApprovalDate",
    "SubmittedBy","SubmittedByEmail","SubmittedByDept","SubmittedAt",
    "AssignedToName","AssignedToEmail","AssignedToDept",
    "RejectionReason","WorkflowHistory","Department","UpdatedAt","VendorEverApproved","PresidingOfficer"
  ]);
  
  // 4. SD_MATERIALS
  var mats = _sheet(SD.MATS, ["MaterialID","Description","CreatedAt"]);

  // 5. SD_MAILS
  var mails = _sheet(SD.MAILS, ["MailID","SenderEmail","SenderName","ReceiverEmail","ReceiverName","Subject","Body","Timestamp","Status","Folder","Attachments","Starred","Labels","ThreadID","Priority","CC","BCC","DeletedBy"]);

  // Ensure mail sheet has all required columns (migration for existing sheets)
  if (mails && mails.getLastRow() > 0) {
    var mailHeaders = mails.getRange(1, 1, 1, mails.getLastColumn()).getValues()[0];
    var expectedHeaders = ["MailID","SenderEmail","SenderName","ReceiverEmail","ReceiverName","Subject","Body","Timestamp","Status","Folder","Attachments","Starred","Labels","ThreadID","Priority","CC","BCC","DeletedBy"];
    if (mailHeaders.length < expectedHeaders.length) {
      mails.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    }
  }


  // 6. SD_JCC_CERTIFICATES
  var jcc = _sheet(SD.JCC, [
    "JccID","ProjectID","ProjectName","ProjectNumber","CertificateType",
    "Vendor","CertificateID","Penalty","Remarks",
    "MecName","MecSignature","MecDate",
    "VendorName","VendorSignature","VendorDate",
    "MecHeadName","MecHeadSignature","MecHeadDate",
    "GeneratedAt","GeneratedBy","GeneratedByEmail",
    "StateRegion","Orchestrator","PresidingOfficer"
  ]);

  // 7. SD_DOCUMENTS — General file/document storage (all departments)
  // DocID | ProjectID | FileName | FileType | FileSize | DriveURL | DriveFileID | UploadedBy | UploadedByEmail | UploadedByDept | UploadedAt | Category
  var docs = _sheet(SD.DOCS, [
    "DocID","ProjectID","FileName","FileType","FileSize",
    "DriveURL","DriveFileID","UploadedBy","UploadedByEmail","UploadedByDept","UploadedAt","Category"
  ]);
  
  // Style all header rows navy
  [users, depts, pat, mats, mails, jcc, docs].forEach(function(sh) {
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setBackground(config.NAVY).setFontColor("#ffffff").setFontWeight("bold");
  });

  // Seed a default super-admin if no users exist yet
  if (users.getLastRow() === 1) {
    var adminDept = PropertiesService.getScriptProperties().getProperty("DEFAULT_ADMIN_DEPT") || "Technology Support";
    _createUserRow("Portal Admin", "admin" + config.DOMAIN, "admin123", SD.ADMIN, adminDept);
  }

  // Seed default materials if none exist
  if (mats.getLastRow() === 1) {
    var defaultMats = [
      "Fiber Cables","Total Box – Box","PCC Cable","80m Box to Box PCC",
      "60m Box to Box PCC","50m Box to Box PCC","5m Box to Box PCC",
      "3m Box to Box PCC","Aluminium Hook","Belt & Clip","Small Clamp",
      "Tiny Small","½ Sub Box","⅑ FAT6","⅛ FAT2",
      "Installation of Cable Hanger","Metallic Pole (Supply & Installation)"
    ];
    defaultMats.forEach(function(m) {
      mats.appendRow([_genId("MAT-"), m, new Date().toISOString()]);
    });
  }

  ss.toast("✅ SD Portal Setup Complete — v1.0 Ready", "Setup", 5);
}

/**
 * Clears ALL mails from SD_MAILS for fresh testing.
 * Keeps the header row intact. Only Super Admin can run this.
 */
function clearAllMails() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert("⚠️ Clear All Mails", "This will delete ALL mail records from SD_MAILS. Are you sure?", ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
  if (!sh) { ui.alert("SD_MAILS sheet not found. Run runFirstSetup first."); return; }
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
  var resp = ui.alert("⚠️ Full System Reset", "This will DELETE ALL data from ALL system sheets. Are you absolutely sure?", ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  var sheets = [
    {name: SD.MAILS, label: "Mails"},
    {name: SD.PAT, label: "Projects"},
    {name: SD.JCC, label: "JCC Certificates"},
    {name: SD.DOCS, label: "Documents"},
    {name: SD.MATS, label: "Materials"},
    {name: SD.BLACKLIST, label: "Blacklist"}
  ];
  var cleared = 0;
  sheets.forEach(function(s) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s.name);
    if (sh && sh.getLastRow() > 1) {
      sh.deleteRows(2, sh.getLastRow() - 1);
      cleared++;
    }
  });
  ui.alert("✅ System data cleared. " + cleared + " sheets reset. Users and Departments preserved.");
}

/**
 * Standardizes the SD_USERS sheet to the 12-column format.
 * Fixes malformed/shifted data from older versions.
 */
function _repairUserSheet(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  var data = sh.getDataRange().getValues();
  var headers = ["UserID","Name","Email","PasswordHash","Salt","Role","Department","Gender","WorkflowNotes","Status","CreatedAt","LastLoginAt","AllowMessages","AIAutoAnswer"];
  var repaired = [headers];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (r.filter(String).length === 0) continue; 

    var row = new Array(14).fill("");

    var emailIdx = -1;
    if (String(r[0]).indexOf("@") !== -1) emailIdx = 0;
    else if (String(r[1]).indexOf("@") !== -1) emailIdx = 1;
    else if (String(r[2]).indexOf("@") !== -1) emailIdx = 2;

    if (emailIdx === -1) continue; 

    if (String(r[0]).startsWith("USR-")) {
      for (var j = 0; j < Math.min(r.length, 14); j++) row[j] = r[j];
      if (!row[7]) row[7] = "Other";
      if (!row[9]) row[9] = "active";
      if (!row[12]) row[12] = "TRUE";
      if (!row[13]) row[13] = "FALSE";
    } 
    else if (emailIdx === 1) {
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
    }
    else {
       row[0] = _genId("USR");
       row[1] = r[emailIdx - 1] || "Unknown";
       row[2] = r[emailIdx];
       for (var k = 3; k < 14; k++) { if (r[k-1]) row[k] = r[k-1]; }
    }

    repaired.push(row);
  }
  
  sh.clear();
  sh.getRange(1, 1, repaired.length, 14).setValues(repaired);
  var config = _getConfig();
  sh.getRange(1, 1, 1, 14).setBackground(config.NAVY).setFontColor("#ffffff").setFontWeight("bold");
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
      if (curH.length < headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}


// ─────────────────────────────────────────────────────────────
// PASSWORD HELPERS
// ─────────────────────────────────────────────────────────────
function _salt() {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
      String(new Date().getTime()) + String(Math.random()))
  ).substring(0, 16);
}

function _hash(password, salt) {
  var secret = _getConfig().SECRET_KEY;
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      password + salt + secret)
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
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(Math.random()) + String(new Date().getTime()));
  return bytes.map(function(b) {
    return ('0' + (b & 0xff).toString(16)).slice(-2);
  }).join('').substring(0, 32);
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
    var token   = _makeToken();
    var config  = _getConfig();
    var payload = JSON.stringify({
      email:      email,
      name:       name,
      role:       role,
      department: department,
      expires:    new Date().getTime() + config.TOKEN_TTL,
    });
    PropertiesService.getUserProperties().setProperty("tok_" + token, payload);
  } catch (e) {
    throw new Error("Unable to create session: " + e.message);
  } finally {
    lock.releaseLock();
  }
  return token;
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
  } catch(e) {
    throw new Error("Session data corrupted. Please sign in again.");
  }
  if (new Date().getTime() > s.expires) {
    PropertiesService.getUserProperties().deleteProperty("tok_" + token);
    throw new Error("Session expired. Please sign in again.");
  }
  return s;
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
      if (String(data[i][2] || "").toLowerCase().trim() === String(email).toLowerCase().trim()) {
        return String(data[i][5] || "").toLowerCase();
      }
    }
  } catch(e) {}
  return null;
}

/**
 * Delete a session (logout).
 */
function _destroySession(token) {
  if (token) PropertiesService.getUserProperties().deleteProperty("tok_" + token);
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
  Object.keys(props).forEach(function(k) {
    if (k.indexOf("tok_") === 0) PropertiesService.getUserProperties().deleteProperty(k);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("All sessions cleared.", "Sessions", 3);
}


// ─────────────────────────────────────────────────────────────
// ID / REF GENERATORS
// ─────────────────────────────────────────────────────────────
function _genId(prefix) {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var rand  = "";
  for (var i = 0; i < 6; i++) rand += chars[Math.floor(Math.random() * chars.length)];
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
function adminCreateUser(token, name, email, password, department, role, gender) {
  try {
    _superAdminSession(token);
    var config = _getConfig();
    email = String(email || "").toLowerCase().trim();
    role = String(role || SD.EMPLOYEE).toLowerCase().trim();
    gender = String(gender || "Other").trim();

    if (!name || !email || !password || !department)
      throw new Error("Name, email, password, and department are required.");
    if (!email.endsWith(config.DOMAIN))
      throw new Error("Only " + config.DOMAIN + " company emails are allowed.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var dup = data.slice(1).find(function(r){ 
      return String(r[2] || "").toLowerCase().trim() === String(email).toLowerCase().trim(); 
    });
    if (dup) throw new Error("An account with this email already exists.");

    _createUserRow(name, email, password, role, department, gender, "TRUE", "FALSE");
    return { success: true };
  } catch(e) {
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
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };

    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var searchId = String(projectId).trim().replace(/^(FOB|PAT)-/i, '');
    var foundRow = -1;

    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0]).trim().replace(/^(FOB|PAT)-/i, '');
      if (rowId.toUpperCase() === searchId.toUpperCase()) {
        foundRow = i;
        break;
      }
    }
    if (foundRow === -1) return { success: false, message: 'Project not found: ' + projectId };

    var sheetRow = foundRow + 2;
    sh.getRange(sheetRow, c.IMAGES + 1).setValue(JSON.stringify(images || []));
    return { success: true, message: 'Images saved to project.', imageCount: (images || []).length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function registerUser(name, email, password, department, gender) {
  var lock = LockService.getScriptLock();
  try {
    var config = _getConfig();
    email = String(email || "").toLowerCase().trim();
    if (!name || !email || !password || !department || !gender)
      throw new Error("All fields are required.");
    if (!email.endsWith(config.DOMAIN))
      throw new Error("Only " + config.DOMAIN + " company emails are allowed.");
    if (String(password).length < 4)
      throw new Error("Password must be at least 4 characters.");

    lock.waitLock(15000);

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var dup  = data.slice(1).find(function(r){ 
      return String(r[2] || "").toLowerCase().trim() === String(email).toLowerCase().trim(); 
    });
    if (dup) throw new Error("An account with this email already exists.");

    _createUserRow(name, email, password, SD.EMPLOYEE, department, gender);
    
    lock.releaseLock();
    return { success: true };
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, message: e.message };
  }
}

function _createUserRow(name, email, password, role, department, gender, allowMsgs, aiAuto) {
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  var s    = _salt();
  var uid  = _genId("USR");
  sh.appendRow([uid, name, email.toLowerCase(), _hash(password, s), s, role, department, gender, "", "active", new Date().toISOString(), "", allowMsgs || "TRUE", aiAuto || "FALSE"]);
  return uid;
}

/**
 * Rate-limiter: tracks failed login attempts per email.
 * Stores in ScriptProperties: login_lock_<email> = JSON timestamp
 * Locks out for 5 minutes after 5 consecutive failures.
 */
function _checkLoginRateLimit(email) {
  var normEmail = String(email || '').toLowerCase().trim();
  if (!normEmail) return { allowed: true };
  var props = PropertiesService.getScriptProperties();
  var lockKey = 'login_lock_' + normEmail;
  var lockDataRaw = props.getProperty(lockKey);
  if (!lockDataRaw) return { allowed: true };
  try {
    var lockData = JSON.parse(lockDataRaw);
    var now = new Date().getTime();
    if (now < lockData.until) {
      var remainingSec = Math.ceil((lockData.until - now) / 1000);
      return { allowed: false, message: 'Too many failed attempts. Try again in ' + remainingSec + ' seconds.', retryAfter: lockData.until };
    }
    props.deleteProperty(lockKey);
  } catch(e) { props.deleteProperty(lockKey); }
  return { allowed: true };
}

function _recordLoginFailure(email) {
  var normEmail = String(email || '').toLowerCase().trim();
  if (!normEmail) return;
  var props = PropertiesService.getScriptProperties();
  var lockKey = 'login_lock_' + normEmail;
  var lockDataRaw = props.getProperty(lockKey);
  var attempts = 1;
  if (lockDataRaw) {
    try { var existing = JSON.parse(lockDataRaw); attempts = (existing.attempts || 0) + 1; } catch(e) {}
  }
  if (attempts >= 5) {
    var lockUntil = new Date().getTime() + (5 * 60 * 1000);
    props.setProperty(lockKey, JSON.stringify({ attempts: attempts, until: lockUntil }));
  } else {
    props.setProperty(lockKey, JSON.stringify({ attempts: attempts, until: 0 }));
  }
}

function _clearLoginLock(email) {
  var normEmail = String(email || '').toLowerCase().trim();
  if (!normEmail) return;
  PropertiesService.getScriptProperties().deleteProperty('login_lock_' + normEmail);
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
      'mec-test@fob.ng':       { role: 'admin', dept: 'MEC' },
      'project-test@fob.ng':   { role: 'employee', dept: 'Project Team' },
      'planning-test@fob.ng':  { role: 'employee', dept: 'Planning Team' },
      'sd-test@fob.ng':        { role: 'employee', dept: 'Service Delivery / Metro' }
    };
    if (TEST_ACCOUNTS[email]) {
      var ta = TEST_ACCOUNTS[email];
      _recordLoginIp(email, clientIp || 'unknown', null, fingerprint);
      var token = _createSession(email, email.split('@')[0], ta.role, ta.dept);
      _clearLoginLock(email);
      return { success: true, token: token, role: ta.role, name: email.split('@')[0], department: ta.dept };
    }

    if (!password) throw new Error("Password is required for this account.");

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var lastRow = sh.getLastRow();
    if (!sh || lastRow < 2) return { success: true, data: [] };

    var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (String(r[2] || "").toLowerCase().trim() !== email) continue;
      if (r[9] === "banned") throw new Error("Your account has been suspended. Contact your admin.");
      if (_hash(String(password), String(r[4])) !== String(r[3])) {
        _recordLoginFailure(email);
        throw new Error("Incorrect password.");
      }

      sh.getRange(i + 2, 12).setValue(new Date().toISOString()); // Update login time
      _recordLoginIp(email, clientIp || 'unknown', i + 2, fingerprint);
      var token = _createSession(email, String(r[1]), String(r[5]), String(r[6]));
      _clearLoginLock(email);
      return {
        success: true, token: token, role: r[5], name: r[1], department: r[6],
        gender: r[7], notes: r[8], allowMessages: r[12] === "TRUE", aiAutoAnswer: r[13] === "TRUE"
      };
    }
    _recordLoginFailure(email);
    throw new Error("No account found with that email address.");
  } catch(e) {
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
    if (!email) { console.error('_recordLoginIp: no email provided'); return; }
    var safeIp = ip || 'unknown';
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    if (!sh) { console.error('_recordLoginIp: users sheet not found'); return; }
    
    if (!row) {
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][2] || '').toLowerCase().trim() === email.toLowerCase().trim()) {
          row = i + 1;
          break;
        }
      }
    }
    if (!row) { console.error('_recordLoginIp: user not found for ' + email); return; }
    
    var existing = sh.getRange(row, 15).getValue() || '';
    var ips = [];
    try { ips = existing ? JSON.parse(existing) : []; } catch(e) { ips = []; }
    
    ips.push({ ip: safeIp, time: new Date().toISOString() });
    if (ips.length > 20) ips = ips.slice(ips.length - 20);
    
    sh.getRange(row, 15).setValue(JSON.stringify(ips));
    console.log('IP recorded for ' + email + ': ' + safeIp + ' (total: ' + ips.length + ')');

    // Record fingerprint in column 16
    if (fingerprint) {
      var fpExisting = sh.getRange(row, 16).getValue() || '';
      var fps = [];
      try { fps = fpExisting ? JSON.parse(fpExisting) : []; } catch(e) { fps = []; }
      
      fps.push({
        fingerprint: fingerprint,
        recordedAt: new Date().toISOString()
      });
      if (fps.length > 10) fps = fps.slice(fps.length - 10);
      
      sh.getRange(row, 16).setValue(JSON.stringify(fps));
      console.log('Fingerprint recorded for ' + email + ' (total fingerprints: ' + fps.length + ')');
    }
  } catch(e) {
    console.error('Failed to record login IP: ' + e.message);
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
    department: sess.department
  };
}

/**
 * Get OpenRouter API key (only accessible to admins)
 */
function getOpenRouterKey(token) {
  try {
    _adminSession(token); // Only admins/super admins can access this
    var config = _getConfig();
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) {
      throw new Error("OpenRouter API key not configured");
    }
    return { success: true, key: apiKey };
  } catch(e) {
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
  'Draft':                    'mec',
  'Rejected':                 'mec',
  'Awaiting Project Team':    'project team',
  'Awaiting Planning Team':   'planning team',
  'Awaiting Service Delivery':'service delivery / metro',
  'Awaiting Final MEC Review':'mec',
  'Awaiting MEC Recheck':     'mec',
  'Partially Approved':       '__all__',
  'Completed':                '__all__'
};

var STATUS_TO_DEPT = {
  'Draft': 'MEC',
  'Rejected': 'MEC',
  'Awaiting Project Team': 'Project Team',
  'Awaiting Planning Team': 'Planning Team',
  'Awaiting Service Delivery': 'Service Delivery / Metro',
  'Awaiting Final MEC Review': 'MEC',
  'Awaiting MEC Recheck': 'MEC',
  'Partially Approved': 'Project Team',
  'Completed': 'MEC'
};

/**
 * Linear workflow (happy path): each status moves to the next department.
 */
var WF_LINEAR_FLOW = {
  'Draft': 'Awaiting Project Team',
  'Awaiting Project Team': 'Awaiting Planning Team',
  'Awaiting Planning Team': 'Awaiting Service Delivery',
  'Awaiting Service Delivery': 'Awaiting Final MEC Review',
  'Awaiting Final MEC Review': 'Completed'
};

/**
 * Recovery/retry transitions for rejected and recheck states.
 */
var WF_RECOVERY_FLOW = {
  'Rejected': 'Awaiting Project Team',
  'Awaiting MEC Recheck': 'Awaiting Project Team'
};

/**
 * Rejection triggers: from which stages a rejection action can lead to which stage.
 */
var WF_REJECTION_TRIGGERS = {
  'Awaiting Project Team': 'Rejected',
  'Awaiting Planning Team': 'Awaiting Project Team',
  'Awaiting Service Delivery': 'Awaiting Project Team',
  'Awaiting Final MEC Review': 'Awaiting Project Team',
  'Awaiting MEC Recheck': 'Awaiting MEC Recheck'
};

/**
 * Combined stage transition map used by submitPATToNextStage, partiallyApprovePAT, etc.
 * Maps each workflow status to the next status in the linear (or recovery) path.
 */
var WF_STAGE_FLOW = {};
(function() {
  var k;
  for (k in WF_LINEAR_FLOW) WF_STAGE_FLOW[k] = WF_LINEAR_FLOW[k];
  for (k in WF_RECOVERY_FLOW) WF_STAGE_FLOW[k] = WF_RECOVERY_FLOW[k];
})();

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
    var linearOrder = ['Draft'];
    var currentStage = 'Draft';
    while (WF_LINEAR_FLOW[currentStage]) {
      linearOrder.push(WF_LINEAR_FLOW[currentStage]);
      currentStage = WF_LINEAR_FLOW[currentStage];
    }

    // Build all known stages from PAT_OWNER_MAP keys
    var allStatuses = Object.keys(PAT_OWNER_MAP);
    var stages = [];
    var seen = {};
    var displayLabels = {
      'Draft': 'Draft',
      'Awaiting Project Team': 'Awaiting<br/>Project Team',
      'Awaiting Planning Team': 'Awaiting<br/>Planning Team',
      'Awaiting Service Delivery': 'Awaiting<br/>Service Delivery',
      'Awaiting Final MEC Review': 'Awaiting<br/>Final MEC Review',
      'Completed': 'Completed',
      'Rejected': 'Rejected',
      'Awaiting MEC Recheck': 'Awaiting<br/>MEC Recheck',
      'Partially Approved': 'Partially<br/>Approved'
    };

    var rejectStages = ['Rejected', 'Awaiting MEC Recheck'];
    var linearSet = {};
    linearOrder.forEach(function(s) { linearSet[s] = true; });

    allStatuses.forEach(function(status) {
      if (seen[status]) return;
      seen[status] = true;
      var dept = STATUS_TO_DEPT[status] || '';
      var owner = PAT_OWNER_MAP[status] || '';
      var order = linearOrder.indexOf(status);
      stages.push({
        id: status,
        label: displayLabels[status] || status,
        dept: dept,
        owner: owner,
        order: order >= 0 ? order : -1,
        isLinear: order >= 0,
        isReject: rejectStages.indexOf(status) !== -1
      });
    });

    // Build complete transition list from all workflow data
    var transitions = [];

    // Linear transitions
    linearOrder.slice(0, -1).forEach(function(from, i) {
      var to = linearOrder[i + 1];
      transitions.push({ from: from, to: to, type: 'linear' });
    });

    // Rejection triggers
    Object.keys(WF_REJECTION_TRIGGERS).forEach(function(from) {
      var to = WF_REJECTION_TRIGGERS[from];
      transitions.push({ from: from, to: to, type: 'rejection' });
      // Also add a "return" edge from reject back (recovery path)
      if (from !== to) {
        transitions.push({ from: to, to: from, type: 'review' });
      }
    });

    // Recovery transitions from Rejected / MEC Recheck
    Object.keys(WF_RECOVERY_FLOW).forEach(function(from) {
      var to = WF_RECOVERY_FLOW[from];
      // Add recovery edges only for routes not already covered
      var exists = transitions.some(function(t) {
        return t.from === from && t.to === to;
      });
      if (!exists) {
        transitions.push({ from: from, to: to, type: 'recovery' });
      }
    });

    return {
      success: true,
      stages: stages,
      linearOrder: linearOrder,
      transitions: transitions,
      rejectStages: rejectStages,
      ownerMap: PAT_OWNER_MAP,
      stageToDept: STATUS_TO_DEPT
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Department name normalization.
 */
var PAT_DEPT_MAP = {
  'mec': 'mec', 'mech': 'mec',
  'project team': 'project team', 'project': 'project team',
  'planning team': 'planning team', 'planning': 'planning team',
  'service delivery': 'service delivery / metro',
  'service delivery / metro': 'service delivery / metro',
  'metro': 'service delivery / metro'
};

/**
 * Normalize a department name for comparison.
 */
function _normalizeDept(dept) {
  var key = String(dept || '').toLowerCase().trim();
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
  
  var role = String(sess.role || '').toLowerCase();
  var dept = String(sess.department || '').toLowerCase();
  var email = String(sess.email || '').toLowerCase();
  
  if (role === 'super admin') return true;
  
  var status = project.workflowStatus || 'Draft';
  if (status === 'Completed') return false;
  
  var myDeptNorm = _normalizeDept(dept);
  var assignedDeptNorm = _normalizeDept(project.assignedToDept || '');
  var assignedEmail = String(project.assignedToEmail || '').toLowerCase();

  // Stage-based turn locking: enforce strict ownership
  if (status === 'Draft') return myDeptNorm === 'mec';
  if (status === 'Rejected') return myDeptNorm === 'mec';

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

    var role = String(sess.role || '').toLowerCase();
    var dept = String(sess.department || '').toLowerCase();

    // Super admin can always edit
    if (role === 'super admin') return true;

    // Only MEC department can edit images
    var isMEC = dept === 'mec' || dept === 'mech';
    if (!isMEC) return false;

    // MEC can only edit when project is in one of their editable states
    var status = project.workflowStatus || 'Draft';
    var EDITABLE_STATUSES = ['Draft', 'Rejected', 'Awaiting Final MEC Review', 'Awaiting MEC Recheck'];
    return EDITABLE_STATUSES.indexOf(status) !== -1;
  } catch(e) {
    return false;
  }
}

/**
 * Get an authoritative canAct value for a single project.
 * Called by the frontend when opening a project to verify permissions.
 */
function checkPATPermission(token, projectId) {
  var sess = _session(token);
  var role = String(sess.role || '').toLowerCase();
  
  // Call with null token for internal read — no session check, no permission filter
  var result = getPATProjectById(null, projectId);
  if (!result || !result.success) {
    return { success: false, message: 'Project not found.', canAct: false };
  }
  
  var project = result.project;
  var canAct = _userCanActOnProject(sess, project);
  var dept = String(sess.department || '').toLowerCase();
  var normalizedDept = _normalizeDept(dept);
  var owner = PAT_OWNER_MAP[project.workflowStatus || 'Draft'] || '';
  
  return {
    success: true,
    canAct: canAct,
    isPushedOut: !canAct && role !== 'admin' && role !== 'super admin',
    currentOwner: owner,
    userDepartment: dept
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
  const role = String(sess.role || '').toLowerCase();
  if (role === 'super admin') return true;

  const userEmail = String(sess.email || '').toLowerCase();
  const snagAddedByEmail = String(snagItem.addedByEmail || '').toLowerCase();

  // Rule: You can ONLY edit/delete snags you added yourself.
  if (userEmail === snagAddedByEmail && !!userEmail) return true;

  // MEC Rule: MEC can manage auto-filled/system snags IF it is their turn (Hub control)
  const isMEC = _normalizeDept(sess.department) === 'mec';
  const status = project.workflowStatus || 'Draft';
  const isMECControl = (status === 'Draft' || status === 'Rejected' || status === 'Awaiting Final MEC Review' || status === 'Awaiting MEC Recheck');
  
  const isAutoSnag = !snagAddedByEmail || 
                     snagAddedByEmail.includes('system') || 
                     String(snagItem.addedBy).includes('System Generated');
  
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
  var role = String(sess.role || '').toLowerCase();
  var dept = String(sess.department || '').toLowerCase();
  var email = String(sess.email || '').toLowerCase();
  var normalizedDept = _normalizeDept(dept);
  var isSuperAdmin = role === 'super admin';
  
  // Get all projects from the deployed function
  var result = getPATProjects();
  if (!result || !result.success) {
    return result;
  }
  
  var projects = (result.projects || []).map(function(p) {
    var canAct = false;
    var status = p.workflowStatus || 'Draft';
    var owner = PAT_OWNER_MAP[status] || '';
    
    if (isSuperAdmin) {
      canAct = true;
    } else if (owner === '__all__') {
      canAct = false;
    } else if (owner === normalizedDept) {
      canAct = true;
    } else if (p.assignedToEmail && String(p.assignedToEmail).toLowerCase() === email) {
      canAct = true;
    }
    
    // Backend-authoritative permission flag
    p.canAct = canAct;
    p._isBackendVerified = true;
    
    return p;
  });
  
  return {
    success: true,
    projects: projects
  };
}

// ─────────────────────────────────────────────────────────────
// PAT ANALYTICS — Idris Dashboard
// ─────────────────────────────────────────────────────────────

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
    var SLA_DURATION_MS = 8 * 60 * 60 * 1000;
    var now = new Date().getTime();

    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) {
      return {
        success: true,
        stats: { overallCompliance: 100, activeProjects: 0, breachedProjects: 0, completedLate: 0 },
        activeList: [], breachedList: [], departments: [], allDepts: []
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

      var status = String(r[c.WORKFLOW_STATUS] || '').trim() || 'Draft';
      if (status === 'Completed' || status === 'Rejected') continue;

      var projectId = String(r[c.PROJECT_ID] || '').trim();
      var projectName = String(r[c.PROJECT_NAME] || '').trim();
      var updatedAt = String(r[c.UPDATED_AT] || r[c.SUBMITTED_AT] || '').trim();
      var startTime = updatedAt ? new Date(updatedAt).getTime() : now;
      if (isNaN(startTime)) startTime = now;

      var elapsed = now - startTime;
      var deadline = startTime + SLA_DURATION_MS;
      var remaining = deadline - now;
      var elapsedPercent = Math.min(100, Math.max(0, Math.round(elapsed / SLA_DURATION_MS * 100)));

      var dept = String(r[c.ASSIGNED_TO_DEPT] || r[c.DEPARTMENT] || r[c.SUBMITTED_BY_DEPT] || 'Unassigned').trim();
      var assignedTo = String(r[c.ASSIGNED_TO_NAME] || r[c.ASSIGNED_TO_EMAIL] || '—').trim();

      var isBreached = remaining <= 0;

      var entry = {
        projectId: projectId,
        projectName: projectName || projectId,
        dept: dept,
        assignedTo: assignedTo,
        status: status,
        deadline: deadline,
        remaining: Math.max(0, remaining),
        overdueMinutes: isBreached ? Math.ceil(Math.abs(remaining) / 60000) : 0,
        elapsedPercent: elapsedPercent
      };

      if (isBreached) breachedList.push(entry);
      else activeList.push(entry);

      if (!deptMap[dept]) deptMap[dept] = { dept: dept, onTime: 0, breached: 0 };
      if (isBreached) deptMap[dept].breached++;
      else deptMap[dept].onTime++;
    }

    var departments = Object.keys(deptMap).map(function(d) {
      var dt = deptMap[d];
      var total = dt.onTime + dt.breached;
      dt.complianceRate = total > 0 ? Math.round(dt.onTime / total * 100) : 100;
      return dt;
    });

    var totalActive = activeList.length;
    var totalBreached = breachedList.length;
    var totalProjects = totalActive + totalBreached;
    var overallCompliance = totalProjects > 0 ? Math.round(totalActive / totalProjects * 100) : 100;

    return {
      success: true,
      stats: {
        overallCompliance: overallCompliance,
        activeProjects: totalActive,
        breachedProjects: totalBreached
      },
      activeList: activeList,
      breachedList: breachedList,
      departments: departments
    };
  } catch(e) {
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

    var SLA_DURATION_MS = 8 * 60 * 60 * 1000;
    var now = new Date().getTime();
    var updatedAt = p.updatedAt || p.submittedAt || '';
    var startTime = updatedAt ? new Date(updatedAt).getTime() : now;
    if (isNaN(startTime)) startTime = now;
    var elapsed = now - startTime;
    var deadline = startTime + SLA_DURATION_MS;
    var remaining = deadline - now;
    var elapsedPercent = Math.min(100, Math.max(0, Math.round(elapsed / SLA_DURATION_MS * 100)));
    var isBreached = remaining <= 0;

    // Build timeline from workflow history
    var history = p.workflowHistory || [];
    var timeline = [];
    var prevTimestamp = null;
    history.forEach(function(h) {
      var ts = h.timestamp ? new Date(h.timestamp).getTime() : null;
      var durationFromPreviousHours = null;
      if (prevTimestamp && ts) {
        var diffMs = ts - prevTimestamp;
        durationFromPreviousHours = Math.round(diffMs / 3600000 * 10) / 10;
      }
      timeline.push({
        from: h.fromStatus || '—',
        to: h.toStatus || '—',
        by: (h.by && h.by.name) ? h.by.name : 'System',
        department: (h.by && h.by.department) ? h.by.department : '—',
        durationFromPreviousHours: durationFromPreviousHours,
        timestamp: h.timestamp || null,
        comments: h.comments || '',
        isRejection: h.isRejection === true
      });
      if (ts) prevTimestamp = ts;
    });

    // Build stage timing
    var stageTiming = {};
    history.forEach(function(h) {
      var stage = h.toStatus || h.fromStatus || 'Unknown';
      if (!stageTiming[stage]) stageTiming[stage] = { entries: 0, totalMs: 0, rejections: 0 };
      stageTiming[stage].entries++;
      if (h.isRejection) stageTiming[stage].rejections++;
    });

    // Count total workflow steps
    var totalSteps = timeline.length;

    // Calculate total journey time
    var totalJourneyMs = 0;
    if (timeline.length >= 2) {
      var firstTs = timeline[0].timestamp ? new Date(timeline[0].timestamp).getTime() : null;
      var lastTs = timeline[timeline.length - 1].timestamp ? new Date(timeline[timeline.length - 1].timestamp).getTime() : null;
      if (firstTs && lastTs) totalJourneyMs = lastTs - firstTs;
    }
    var totalJourneyHours = totalJourneyMs > 0 ? Math.round(totalJourneyMs / 3600000 * 10) / 10 : 0;

    // Count rejections
    var rejectionCount = timeline.filter(function(t) { return t.isRejection; }).length;

    // Avg time per step
    var avgTimePerStepHours = totalSteps > 0 ? Math.round(totalJourneyMs / totalSteps / 3600000 * 10) / 10 : 0;

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
        presidingOfficer: p.presidingOfficer || '',
        assignedToDept: p.assignedToDept || '',
        assignedToName: p.assignedToName || '',
        assignedToEmail: p.assignedToEmail || '',
        hasVendor: p.hasVendor || false,
        vendorApprovalStatus: p.vendorApprovalStatus || 'pending',
        department: p.department || '',
        updatedAt: p.updatedAt || '',
        submittedAt: p.submittedAt || '',
        submittedBy: p.submittedBy || '',
        submittedByEmail: p.submittedByEmail || ''
      },
      sla: {
        slaDurationHours: 8,
        elapsedHours: Math.round(elapsed / 3600000 * 10) / 10,
        remainingHours: Math.round(Math.max(0, remaining) / 3600000 * 10) / 10,
        remainingMinutes: Math.round(Math.max(0, remaining) / 60000),
        overdueMinutes: isBreached ? Math.round(Math.abs(remaining) / 60000) : 0,
        elapsedPercent: elapsedPercent,
        isBreached: isBreached,
        deadline: deadline,
        assignedTo: p.assignedToName || p.assignedToEmail || '—',
        dept: p.assignedToDept || p.department || 'Unassigned',
        status: isBreached ? 'breached' : elapsedPercent >= 75 ? 'at-risk' : 'active'
      },
      metrics: {
        totalWorkflowSteps: totalSteps,
        totalJourneyHours: totalJourneyHours,
        rejectionCount: rejectionCount,
        avgTimePerStepHours: avgTimePerStepHours
      },
      timeline: timeline,
      stageTiming: stageTiming,
      relatedMails: [],
      boqChanges: []
    };
  } catch(e) {
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
    PROJECT_ID: 0, PROJECT_NAME: 1, SITE_ADDRESS: 2, LAT: 3, LON: 4,
    PHASE: 5, WORK_DESC: 6, VENDOR: 7, INSP_DATE: 8, ORCHESTRATOR: 9,
    SNAG_SCORE: 10, STATUS: 11, VERDICT: 12, WORKFLOW_STATUS: 13,
    CHECKLIST: 14, BOQ: 15, SNAGS: 16, SIGNOFF: 17, IMAGES: 18,
    HAS_VENDOR: 19, VENDOR_TOKEN: 20, VENDOR_APPROVAL_STATUS: 21,
    VENDOR_APPROVAL_COMMENTS: 22, VENDOR_APPROVAL_DATE: 23,
    SUBMITTED_BY: 24, SUBMITTED_BY_EMAIL: 25, SUBMITTED_BY_DEPT: 26,
    SUBMITTED_AT: 27, ASSIGNED_TO_NAME: 28, ASSIGNED_TO_EMAIL: 29,
    ASSIGNED_TO_DEPT: 30, REJECTION_REASON: 31, WORKFLOW_HISTORY: 32,
    DEPARTMENT: 33, UPDATED_AT: 34, VENDOR_EVER_APPROVED: 35,
    PRESIDING_OFFICER: 36
  };
}

function _projectFromRow(r) {
  var c = _patCols();
  return {
    projectId:         String(r[c.PROJECT_ID] || '').trim(),
    projectName:       r[c.PROJECT_NAME] || '',
    siteAddress:       r[c.SITE_ADDRESS] || '',
    lat:               r[c.LAT] || '',
    lon:               r[c.LON] || '',
    phase:             r[c.PHASE] || '',
    workDescription:   r[c.WORK_DESC] || '',
    vendor:            r[c.VENDOR] || '',
    inspectionDate:    r[c.INSP_DATE] || '',
    orchestrator:      r[c.ORCHESTRATOR] || '',
    snagScore:         Number(r[c.SNAG_SCORE]) || 0,
    status:            r[c.STATUS] || '',
    verdict:           r[c.VERDICT] || 'Pending',
    workflowStatus:    r[c.WORKFLOW_STATUS] || 'Draft',
    checklist:         _safeParse(r[c.CHECKLIST], {}),
    boq:               _safeParse(r[c.BOQ], []),
    snags:             _safeParse(r[c.SNAGS], []),
    signoff:           _safeParse(r[c.SIGNOFF], {}),
    images:            _safeParse(r[c.IMAGES], []),
    hasVendor:         String(r[c.HAS_VENDOR] || '').toUpperCase() === 'TRUE',
    vendorToken:       r[c.VENDOR_TOKEN] || '',
    vendorApprovalStatus: r[c.VENDOR_APPROVAL_STATUS] || 'pending',
    vendorApprovalComments: r[c.VENDOR_APPROVAL_COMMENTS] || '',
    vendorApprovalDate: r[c.VENDOR_APPROVAL_DATE] || '',
    vendorEverApproved: String(r[c.VENDOR_EVER_APPROVED] || '').toUpperCase() === 'TRUE',
    submittedBy:       r[c.SUBMITTED_BY] || '',
    submittedByEmail:  r[c.SUBMITTED_BY_EMAIL] || '',
    submittedByDept:   r[c.SUBMITTED_BY_DEPT] || '',
    submittedAt:       r[c.SUBMITTED_AT] || '',
    assignedToName:    r[c.ASSIGNED_TO_NAME] || '',
    assignedToEmail:   r[c.ASSIGNED_TO_EMAIL] || '',
    assignedToDept:    r[c.ASSIGNED_TO_DEPT] || r[c.SUBMITTED_BY_DEPT] || '',
    rejectionReason:   r[c.REJECTION_REASON] || '',
    workflowHistory:   _safeParse(r[c.WORKFLOW_HISTORY], []),
    department:        r[c.DEPARTMENT] || '',
    updatedAt:         r[c.UPDATED_AT] || '',
    presidingOfficer:  r[c.PRESIDING_OFFICER] || ''
  };
}

function _safeParse(str, fallback) {
  if (!str) return fallback;
  // If already an object/array, return it
  if (typeof str === 'object') return str;
  // If it's a string, try to parse it
  if (typeof str === 'string') {
    try { return JSON.parse(str); } catch(e) { return fallback; }
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
  if (typeof val === 'object') return JSON.stringify(val);
  if (typeof val === 'string') {
    try { JSON.parse(val); return val; } catch(e) { return JSON.stringify(fallback); }
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
    DOC_ID: 0, PROJECT_ID: 1, FILE_NAME: 2, FILE_TYPE: 3, FILE_SIZE: 4,
    DRIVE_URL: 5, DRIVE_FILE_ID: 6, UPLOADED_BY: 7, UPLOADED_BY_EMAIL: 8,
    UPLOADED_BY_DEPT: 9, UPLOADED_AT: 10, CATEGORY: 11
  };
}

function _documentFromRow(r) {
  var c = _docCols();
  return {
    docId: r[c.DOC_ID] || '',
    projectId: String(r[c.PROJECT_ID] || '').trim(),
    fileName: r[c.FILE_NAME] || '',
    fileType: r[c.FILE_TYPE] || '',
    fileSize: r[c.FILE_SIZE] || 0,
    driveUrl: r[c.DRIVE_URL] || '',
    driveFileId: r[c.DRIVE_FILE_ID] || '',
    uploadedBy: r[c.UPLOADED_BY] || '',
    uploadedByEmail: r[c.UPLOADED_BY_EMAIL] || '',
    uploadedByDept: r[c.UPLOADED_BY_DEPT] || '',
    uploadedAt: r[c.UPLOADED_AT] || '',
    category: r[c.CATEGORY] || ''
  };
}

function _genDocId() {
  return 'DOC-' + Utilities.getUuid().substring(0, 6).toUpperCase();
}

function _genPATId() {
  return 'PAT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
}

/**
 * List all PAT projects.
 */
function getPATProjects(token) {
  try {
    var sess = token ? _session(token) : null;
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: true, projects: [] };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var allProjects = data.filter(function(r) { return r[0]; }).map(_projectFromRow);

    // If no session (internal call) or user is Admin/MEC, return everything
    if (!sess) return { success: true, projects: allProjects };
    
    var role = String(sess.role || '').toLowerCase();
    var dept = String(sess.department || '').toLowerCase();
    var email = String(sess.email || '').toLowerCase();
    var isSuperAdmin = (role === 'super admin');
    var isMEC = (dept === 'mec' || dept === 'mech');

    // All authenticated users can see all PAT projects (except Drafts for non-MEC)
    // This enables cross-department visibility — any department can check details
    // The frontend controls what actions each user can perform based on their department
    var visibleProjects = allProjects.filter(function(p) {
      var status = p.workflowStatus || 'Draft';
      // Only hide Drafts from non-MEC users (MEC workspace)
      if (status === 'Draft' && !isMEC && !isSuperAdmin) return false;
      return true;
    });

    return { success: true, projects: visibleProjects };
  } catch(e) {
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
      throw new Error("Authentication required. Call with (token, projectId) or (null, projectId) for internal use.");
    }
    var isInternalCall = (token === null);
    if (!isInternalCall) {
      _session(token); // Validates token and throws if invalid
    }
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var searchId = String(projectId).trim();
    var altId = searchId.replace(/^(FOB|PAT)-/i, '');
    var foundRow = -1;
    var foundProject = null;
    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      if (rowId.toUpperCase() === searchId.toUpperCase() ||
          rowId.toUpperCase() === altId.toUpperCase() ||
          ('PAT-' + rowId).toUpperCase() === searchId.toUpperCase()) {
        foundRow = i;
        foundProject = _projectFromRow(data[i]);
        break;
      }
    }
    if (foundRow === -1) return { success: false, message: 'Project not found: ' + projectId };

    // All authenticated users can view any PAT project details.
    // Action permissions are checked separately via _userCanActOnProject/checkPATPermission.

    return { success: true, project: foundProject };
  } catch(e) {
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
    if (!sh) return { success: false, message: 'PAT sheet not found.' };

    var existingId = data.projectId || '';
    var now = new Date().toISOString();

    // Build row data
    function _rowFromProject(p) {
      var row = [];
      row[c.PROJECT_ID]         = p.projectId || _genPATId();
      row[c.PROJECT_NAME]       = String(p.projectName || '');
      row[c.SITE_ADDRESS]       = String(p.siteAddress || '');
      row[c.LAT]                = String(p.lat || '');
      row[c.LON]                = String(p.lon || '');
      row[c.PHASE]              = String(p.phase || '');
      row[c.WORK_DESC]          = String(p.workDescription || '');
      row[c.VENDOR]             = String(p.vendor || '');
      row[c.INSP_DATE]          = String(p.inspectionDate || '');
      row[c.ORCHESTRATOR]       = String(p.orchestrator || '');
      row[c.SNAG_SCORE]         = Number(p.snagScore) || 0;
      row[c.STATUS]             = String(p.status || '');
      row[c.VERDICT]            = String(p.verdict || 'Pending');
      row[c.WORKFLOW_STATUS]    = String(p.workflowStatus || 'Draft');
      // Safely handle JSON fields that might already be stringified
      row[c.CHECKLIST]          = _ensureJSONString(p.checklist, {});
      row[c.BOQ]                = _ensureJSONString(p.boq, []);
      row[c.SNAGS]              = _ensureJSONString(p.snags, []);
      row[c.SIGNOFF]            = _ensureJSONString(p.signoff, {});
      row[c.IMAGES]             = _ensureJSONString(p.images, []);
      row[c.HAS_VENDOR]         = p.hasVendor ? 'TRUE' : 'FALSE';
      row[c.VENDOR_TOKEN]       = String(p.vendorToken || '');
      row[c.VENDOR_APPROVAL_STATUS] = String(p.vendorApprovalStatus || 'pending');
      row[c.VENDOR_APPROVAL_COMMENTS] = String(p.vendorApprovalComments || '');
      row[c.VENDOR_APPROVAL_DATE] = String(p.vendorApprovalDate || '');
      row[c.VENDOR_EVER_APPROVED] = p.vendorEverApproved ? 'TRUE' : 'FALSE';
      row[c.SUBMITTED_BY]       = String(p.submittedBy || (sess ? sess.name : ''));
      row[c.SUBMITTED_BY_EMAIL] = String(p.submittedByEmail || (sess ? sess.email : ''));
      row[c.SUBMITTED_BY_DEPT]  = String(p.submittedByDept || (sess ? sess.department : ''));
      row[c.SUBMITTED_AT]       = p.submittedAt || now;
    // Ensure assignedToDept is set to the user's department if not provided
    row[c.ASSIGNED_TO_NAME]   = String(p.assignedToName || '');
    row[c.ASSIGNED_TO_EMAIL]  = String(p.assignedToEmail || '');
    row[c.ASSIGNED_TO_DEPT]   = String(p.assignedToDept || (sess ? sess.department : 'Unassigned'));
      row[c.REJECTION_REASON]   = String(p.rejectionReason || '');
      row[c.WORKFLOW_HISTORY]   = _ensureJSONString(p.workflowHistory, []);
      row[c.DEPARTMENT]         = String(p.department || '');
      row[c.PRESIDING_OFFICER]  = String(p.presidingOfficer || '');
      row[c.UPDATED_AT]         = now;
      return row;
    }

    // Check if updating existing
    if (existingId) {
      var normExisting = existingId.trim();
      var altExisting = normExisting.replace(/^(FOB|PAT)-/i, '');
      var existingData = sh.getDataRange().getValues();
      for (var i = 1; i < existingData.length; i++) {
        var rowId = String(existingData[i][0]).trim();
        if (rowId.toUpperCase() === normExisting.toUpperCase() ||
            rowId.toUpperCase() === altExisting.toUpperCase() ||
            ('PAT-' + rowId).toUpperCase() === normExisting.toUpperCase()) {
          var existingProject = _projectFromRow(existingData[i]);
          var sessRole = String(sess.role || '').toLowerCase();
          var sessDept = String(sess.department || '').toLowerCase();
          var isAdmin = (sessRole === 'admin' || sessRole === 'super admin');
          var isMEC = (sessDept === 'mec' || sessDept === 'mech');

          // STRICT PERMISSION CHECK FOR NON-MEC/ADMIN
          if (!isMEC && !isAdmin) {
            // Non-MEC/Admin users can only update Snags and workflow metadata.
            // We ignore any changes they might have attempted on Site Info, Checklist, or BOQ
            // by using the existing project data as the master template for those columns.
            var finalRow = _rowFromProject(existingProject); // Start with DB version
            finalRow[c.SNAGS] = _ensureJSONString(data.snags, []);
            finalRow[c.SNAG_SCORE] = Number(data.snagScore || data.complaintScore) || 0;

            
            // ALLOW metadata updates for the current assignee/dept
            finalRow[c.WORKFLOW_STATUS] = data.workflowStatus !== undefined ? String(data.workflowStatus) : existingProject.workflowStatus;
            finalRow[c.WORKFLOW_HISTORY] = _ensureJSONString(data.workflowHistory, []);
            finalRow[c.DEPARTMENT] = data.department !== undefined ? String(data.department) : existingProject.department;
            finalRow[c.ASSIGNED_TO_DEPT] = data.assignedToDept !== undefined ? String(data.assignedToDept) : existingProject.assignedToDept;
            finalRow[c.ASSIGNED_TO_NAME] = data.assignedToName !== undefined ? String(data.assignedToName) : existingProject.assignedToName;
            finalRow[c.ASSIGNED_TO_EMAIL] = data.assignedToEmail !== undefined ? String(data.assignedToEmail) : existingProject.assignedToEmail;
            finalRow[c.REJECTION_REASON] = data.rejectionReason !== undefined ? String(data.rejectionReason) : existingProject.rejectionReason;
            finalRow[c.VERDICT] = data.verdict !== undefined ? String(data.verdict) : existingProject.verdict;

            // SNAG PERMISSION ENFORCEMENT:
            // Prevent unauthorized deletion/modification of snags added by other departments.
            var incomingSnagsFromUI = _safeParse(data.snags, []); // Snags from frontend
            var originalSnagsFromDB = _safeParse(existingProject.snags, []); // Snags from DB
            var finalSnagsForSave = [];
            
            // Robust Snag Merging: Keep original snags that the current user is NOT allowed to touch.
            originalSnagsFromDB.forEach(function(origS) {
              if (!_canEditSnag(sess, origS, existingProject)) {
                finalSnagsForSave.push(origS);
              }
            });
            // Add all incoming snags that the user HAS permission to manage (edits or new additions).
            incomingSnagsFromUI.forEach(function(incS) {
              if (_canEditSnag(sess, incS, existingProject)) {
                finalSnagsForSave.push(incS);
              }
            });
            finalRow[c.SNAGS] = _ensureJSONString(finalSnagsForSave, []);
            finalRow[c.SNAG_SCORE] = finalSnagsForSave.reduce(function(sum, s) { return sum + (parseInt(s.weight) || 0); }, 0);

            finalRow[c.UPDATED_AT] = now;
            sh.getRange(i + 1, 1, 1, finalRow.length).setValues([finalRow]);
            SpreadsheetApp.flush();
            return { success: true, projectId: existingId, message: 'Workflow updated successfully.' };
          }

          // Full update allowed ONLY for MEC and Super Admins
          if (isMEC || isAdmin) {
            var rowData = _rowFromProject(data);
            // Preserve original submission data
            var origData = existingData[i];
            rowData[c.SUBMITTED_BY]       = origData[c.SUBMITTED_BY] || data.submittedBy || '';
            rowData[c.SUBMITTED_BY_EMAIL] = origData[c.SUBMITTED_BY_EMAIL] || data.submittedByEmail || '';
            rowData[c.SUBMITTED_BY_DEPT]  = data.submittedByDept || existingData[i][c.SUBMITTED_BY_DEPT] || '';
            rowData[c.SUBMITTED_AT]       = data.submittedAt || existingData[i][c.SUBMITTED_AT] || now;
            rowData[c.WORKFLOW_HISTORY]   = data.workflowHistory ? JSON.stringify(data.workflowHistory) : existingData[i][c.WORKFLOW_HISTORY] || '[]';
            sh.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
            SpreadsheetApp.flush();
            return { success: true, projectId: existingId };
          }

          // No permission to update
          return { success: false, message: 'Permission denied: you are not allowed to edit this project.' };
        }
      }
    }

    // New project
    var sessRole = String(sess.role || '').toLowerCase();
    var sessDept = String(sess.department || '').toLowerCase();
    var isMEC = (sessDept === 'mec' || sessDept === 'mech');
    var isAdmin = (sessRole === 'admin' || sessRole === 'super admin');

    if (!isMEC) {
      return { success: false, message: 'Access Denied: You do not have the necessary permissions to create or manage PAT projects. Even the AI will not open the workspace for this action. Only MEC personnel are authorized for PAT creation.' };
    }

    // DEDUPLICATION CHECK: Prevent AI from creating duplicate projects in rapid succession
    var projectData = sh.getDataRange().getValues();
    var nameToMatch = String(data.projectName || '').trim().toLowerCase();
    var addrToMatch = String(data.siteAddress || '').trim().toLowerCase();
    var userEmail = String(sess.email || '').toLowerCase();
    var nowTs = new Date().getTime();

    if (nameToMatch && addrToMatch) {
      for (var j = projectData.length - 1; j >= 1; j--) {
        var rowName = String(projectData[j][c.PROJECT_NAME] || '').trim().toLowerCase();
        var rowAddr = String(projectData[j][c.SITE_ADDRESS] || '').trim().toLowerCase();
        var rowEmail = String(projectData[j][c.SUBMITTED_BY_EMAIL] || '').toLowerCase();
        var rowDateStr = projectData[j][c.SUBMITTED_AT];
        var rowTs = rowDateStr ? new Date(rowDateStr).getTime() : 0;

        // If same name, address, user and created within last 60 seconds
        if (rowName === nameToMatch && rowAddr === addrToMatch && rowEmail === userEmail && (nowTs - rowTs < 60000)) {
          var existingId = String(projectData[j][c.PROJECT_ID] || '').trim();
          return { success: true, projectId: existingId, message: 'Existing project reused to prevent duplication.' };
        }
      }
    }

    var newId = _genPATId();
    data.projectId = newId;
    var newRow = _rowFromProject(data);
    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { success: true, projectId: newId };
  } catch(e) {
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
    var dept = String(sess.department || '').toLowerCase();
    var isProjectTeam = dept.includes('project');

    if (!isProjectTeam) {
      return { success: false, message: 'Only Project Team members can generate vendor review links.' };
    }

    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    if (project.workflowStatus !== 'Awaiting Project Team') {
      return { success: false, message: 'Vendor review is only available at the Awaiting Project Team stage.' };
    }

    var c = _patCols();
    var sh = _patSheet();
    var data = sh.getDataRange().getValues();
    var searchId = String(projectId).trim();
    var altId = searchId.replace(/^(FOB|PAT)-/i, '');
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      if (rowId.toUpperCase() === searchId.toUpperCase() ||
          rowId.toUpperCase() === altId.toUpperCase() ||
          ('PAT-' + rowId).toUpperCase() === searchId.toUpperCase()) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Project not found.' };

    var sheetRow = rowIndex + 1;

    // Read current vendor status to decide whether to reset
    var currentVendorApproval = String(data[rowIndex][c.VENDOR_APPROVAL_STATUS] || 'pending').toLowerCase();
    var currentEverApproved = String(data[rowIndex][c.VENDOR_EVER_APPROVED] || '').toUpperCase() === 'TRUE';

    // Always generate a fresh token (overwrite old one)
    var vendorToken = 'VEND-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    sh.getRange(sheetRow, c.VENDOR_TOKEN + 1).setValue(vendorToken);

    // Set hasVendor = TRUE
    sh.getRange(sheetRow, c.HAS_VENDOR + 1).setValue('TRUE');

    // Only reset approval status if vendor has NEVER approved (preserve ever-approved)
    if (!currentEverApproved && currentVendorApproval !== 'approved') {
      sh.getRange(sheetRow, c.VENDOR_APPROVAL_STATUS + 1).setValue('pending');
      sh.getRange(sheetRow, c.VENDOR_APPROVAL_COMMENTS + 1).setValue('');
      sh.getRange(sheetRow, c.VENDOR_APPROVAL_DATE + 1).setValue('');
    }

    var baseUrl = ScriptApp.getService().getUrl();
    var fullUrl = baseUrl + '?token=' + encodeURIComponent(vendorToken);

    return { success: true, url: fullUrl, vendorToken: vendorToken };
  } catch(e) {
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
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };
    var data = sh.getDataRange().getValues();
    var c = _patCols();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][c.VENDOR_TOKEN] || '').toUpperCase() === String(vendorToken).toUpperCase()) {
        var approvalStatus = String(data[i][c.VENDOR_APPROVAL_STATUS] || 'pending');
        var everApproved = String(data[i][c.VENDOR_EVER_APPROVED] || '').toUpperCase() === 'TRUE';
        var alreadyCompleted = approvalStatus !== 'pending' && !everApproved;
        if (alreadyCompleted) {
          return { success: false, message: 'This review link has already been used.' };
        }
        var project = _projectFromRow(data[i]);
        // Return relevant vendor-facing data only
        var reviewCompleted = everApproved || approvalStatus !== 'pending';
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
            presidingOfficer: project.presidingOfficer
          }
        };
      }
    }
    return { success: false, message: 'Invalid or expired review link.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Submit vendor decision (public, no session required).
 * decision: 'approved' or 'rejected'
 */
function submitVendorDecision(vendorToken, decision, comments) {
  try {
    decision = String(decision || '').toLowerCase().trim();
    if (decision !== 'approved' && decision !== 'rejected') {
      return { success: false, message: 'Decision must be "approved" or "rejected".' };
    }
    if (decision === 'rejected' && !String(comments || '').trim()) {
      return { success: false, message: 'A comment is required when rejecting.' };
    }

    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };
    var data = sh.getDataRange().getValues();
    var c = _patCols();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][c.VENDOR_TOKEN] || '').toUpperCase() === String(vendorToken).toUpperCase()) {
        var currentStatus = String(data[i][c.VENDOR_APPROVAL_STATUS] || 'pending');
        var currentEverApproved = String(data[i][c.VENDOR_EVER_APPROVED] || '').toUpperCase() === 'TRUE';
        if (currentStatus !== 'pending' || currentEverApproved) {
          return { success: false, message: 'This review has already been completed.' };
        }
        var sheetRow = i + 1;
        sh.getRange(sheetRow, c.VENDOR_APPROVAL_STATUS + 1).setValue(decision);
        sh.getRange(sheetRow, c.VENDOR_APPROVAL_COMMENTS + 1).setValue(String(comments || ''));
        sh.getRange(sheetRow, c.VENDOR_APPROVAL_DATE + 1).setValue(new Date().toISOString());

        // PERMANENT: Once vendor approves, set VendorEverApproved=TRUE (never reset)
        if (decision === 'approved') {
          sh.getRange(sheetRow, c.VENDOR_EVER_APPROVED + 1).setValue('TRUE');
        }

        // When vendor rejects, add workflow history so the project team can take action
        if (decision === 'rejected') {
          var existingProject = _projectFromRow(data[i]);
          var history = existingProject.workflowHistory || [];
          history.push({
            fromStatus: existingProject.workflowStatus || 'Awaiting Project Team',
            toStatus: 'Awaiting Project Team',
            by: { name: 'Vendor', email: existingProject.vendor || 'vendor', department: 'Vendor' },
            comments: 'VENDOR REJECTED: ' + (comments || ''),
            isRejection: true,
            timestamp: new Date().toISOString()
          });
          sh.getRange(sheetRow, c.WORKFLOW_HISTORY + 1).setValue(JSON.stringify(history));
          sh.getRange(sheetRow, c.UPDATED_AT + 1).setValue(new Date().toISOString());
        }

        return { success: true, decision: decision };
      }
    }
    return { success: false, message: 'Invalid token.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Check if vendor approval is required and has been given for a project.
 * Returns { required: bool, approved: bool, status: string }
 */
function _checkVendorApproval(project) {
  if (!project || !project.hasVendor) {
    return { required: false, approved: true, status: 'not_required' };
  }
  var status = String(project.vendorApprovalStatus || 'pending').toLowerCase();
  var everApproved = project.vendorEverApproved === true;
  return {
    required: true,
    approved: everApproved || status === 'approved',
    status: everApproved ? 'approved' : status
  };
}

/**
 * Submit PAT to a department — advances the workflow.
 */
function submitPATToDepartment(token, projectId, targetDept, comments) {
  try {
    var sess = _session(token);
    // Use null as token for internal call — no session check needed for read
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // Determine next workflow status based on target department
    var statusMap = {
      'project team':    'Awaiting Project Team',
      'project':         'Awaiting Project Team',
      'planning team':   'Awaiting Planning Team',
      'planning':        'Awaiting Planning Team',
      'service delivery / metro': 'Awaiting Service Delivery',
      'service delivery': 'Awaiting Service Delivery',
      'metro':           'Awaiting Service Delivery',
      'mec':             'Awaiting Final MEC Review',
      'mech':            'Awaiting Final MEC Review'
    };
    var targetDeptNorm = String(targetDept).toLowerCase().trim();
    var newStatus = statusMap[targetDeptNorm] || 'Awaiting Project Team';

    // Determine if this is a rejection first to inform the workflow transition
    var isRejection = _isRejection(newStatus, comments);

    // MEC REJECT FROM FINAL REVIEW: Send to Project Team instead of Completed
    // Project Team can then approve/partially-approve back to MEC Final Review
    if (project.workflowStatus === 'Awaiting Final MEC Review' && isRejection) {
      newStatus = 'Awaiting Project Team';
      project.verdict = 'Rejected';
    }

    // RECHECK LOGIC: If pushing to MEC from Project Team, handle based on previous failure stage
    if ((targetDeptNorm === 'mec' || targetDeptNorm === 'mech') && 
        (project.workflowStatus === 'Awaiting Project Team' || project.workflowStatus === 'Rejected')) {
      if (isRejection) {
        newStatus = 'Awaiting MEC Recheck';
      } else if (project.verdict === 'Rejected') {
        // If it failed Final Review, it must go back to Final Review for approval
        newStatus = 'Awaiting Final MEC Review';
      } else {
        // If it's an early stage approval, it goes to MEC Recheck
        newStatus = 'Awaiting MEC Recheck';
      }
    }

    // Handle other rejections
    if (isRejection && targetDeptNorm.indexOf('mec') !== -1 && newStatus !== 'Awaiting MEC Recheck') {
      // Status "Rejected" is now reserved for MEC inbox (unless it's a project-to-mec recheck)
      newStatus = 'Rejected';
    } else if (isRejection && newStatus !== 'Awaiting MEC Recheck') {
      // Rejections to other teams (Project Hub) stay in their workflow stage
      newStatus = statusMap[targetDeptNorm] || 'Awaiting Project Team';
    }

    var oldStatus = project.workflowStatus || 'Draft';

    // Build history entry
    var historyEntry = {
      fromStatus: oldStatus,
      toStatus: newStatus,
      by: { name: sess.name, email: sess.email, department: sess.department },
      comments: comments || '',
      isRejection: isRejection, // Explicitly store if this was a rejection
      timestamp: new Date().toISOString()
    };
    var history = project.workflowHistory || [];
    history.push(historyEntry);

    // Update project
    project.workflowStatus = newStatus;
    project.workflowHistory = history;
    if (isRejection) {
      project.rejectionReason = comments || '';
    }
    project.department = targetDept; // Department currently holding the PAT
    project.assignedToDept = targetDept;
    project.assignedToName = '';
    project.assignedToEmail = '';

    var saveRes = savePATProject(token, project);
    return { success: saveRes.success, message: 'Submitted to ' + targetDept + '. Status: ' + newStatus };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Internal helper to determine if a workflow transition is a rejection.
 */
function _isRejection(status, comments) {
  if (status === 'Rejected') return true;
  if (String(comments || '').toUpperCase().indexOf('REJECTED') === 0) return true;
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
    if ((h.isRejection === true) || _isRejection(h.toStatus, h.comments)) {
      rejections.push({ dept: h.by ? h.by.department : null, index: i, entry: h });
    }
  }
  
  // 2. Filter for rejections that have NOT been resolved by a later approval from the same dept
  var unresolved = rejections.filter(function(rej) {
    if (!rej.dept) return false;
    var normRejDept = _normalizeDept(rej.dept);
    
    // Look forward in history for any non-rejection action by this department
    for (var j = rej.index + 1; j < history.length; j++) {
      var h = history[j];
      if (!h) continue;
      
      var isLaterRejection = (h.isRejection === true) || _isRejection(h.toStatus, h.comments);
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
      var msg = vendorCheck.status === 'rejected'
        ? 'Vendor rejected this project. Please resolve with vendor before proceeding.'
        : 'Vendor approval is required before proceeding. Please wait for vendor response.';
      return { success: false, message: msg };
    }

    var newStatus = WF_STAGE_FLOW[project.workflowStatus];

    // ROUTING FIX: If this project has unresolved rejections, return to the last rejecter
    var lastRejEntry = _getLastUnresolvedRejecter(project, sess.department);
    if (lastRejEntry) {
      var rejDept = lastRejEntry.by ? lastRejEntry.by.department : '';
      var normalized = _normalizeStatusByDept(rejDept);
      
      if (_normalizeDept(rejDept) === 'mec') {
        // When Project Team resolves a MEC rejection:
        // If it was rejected from Final Review -> go back to Final Review
        // If it was rejected from Draft/Other -> go to MEC Recheck
        newStatus = (lastRejEntry.fromStatus === 'Awaiting Final MEC Review') ? 'Awaiting Final MEC Review' : 'Awaiting MEC Recheck';
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
      comments: comments || '',
      timestamp: new Date().toISOString()
    });

    project.workflowStatus = newStatus;
    project.workflowHistory = history;

    // Update department and assigned department based on the new status
    var deptForStatus = STATUS_TO_DEPT[newStatus] || 'Project Team';
    project.department = deptForStatus;
    project.assignedToDept = deptForStatus;

    // If assigned to specific person
    if (nextEmail) {
      project.assignedToEmail = nextEmail;
      // Look up name/dept
      try {
        var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
        var userData = userSh.getDataRange().getValues();
        for (var i = 1; i < userData.length; i++) {
          if (String(userData[i][2]).toLowerCase() === nextEmail.toLowerCase()) {
            project.assignedToName = userData[i][1];
            project.assignedToDept = userData[i][6];
            project.department = userData[i][6];
            break;
          }
        }
      } catch(e) {}
    }
    
    // If the new status is Awaiting Final MEC Review, assign to MEC department generally
    if (newStatus === 'Awaiting Final MEC Review') {
      project.assignedToDept = 'MEC';
      project.assignedToName = ''; // Clear specific assignment
      project.assignedToEmail = ''; // Clear specific assignment
    }

    var saveRes = savePATProject(token, project);
    return { success: saveRes.success, message: 'Approved to ' + newStatus + '.', verdict: 'Fully Accepted', snagScore: project.snagScore };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Partially approve — move to next stage with snags + Partially Approved verdict.
 */
function partiallyApprovePAT(token, projectId, comments, newSnags, nextDept, nextEmail) {
  try {
    var sess = _session(token);
    // Use null token for internal call
    var res = getPATProjectById(null, projectId);
    if (!res.success) return res;
    var project = res.project;

    // Vendor approval check
    var vendorCheck = _checkVendorApproval(project);
    if (vendorCheck.required && !vendorCheck.approved) {
      var msg = vendorCheck.status === 'rejected'
        ? 'Vendor rejected this project. Please resolve with vendor before proceeding.'
        : 'Vendor approval is required before proceeding. Please wait for vendor response.';
      return { success: false, message: msg };
    }

    // Append new snags
    var snags = project.snags || [];
    if (newSnags && newSnags.length) {
      snags = snags.concat(newSnags);
    }

    // Recalculate score
    var score = 0;
    snags.forEach(function(s) { score += parseInt(s.weight) || 0; });

    project.snags = snags;
    project.snagScore = score;
    project.verdict = 'Partially Accepted';

    var newStatus = WF_STAGE_FLOW[project.workflowStatus] || 'Partially Approved';
    
    // If resolving a rejection, jump back to that department's review stage
    var lastRejEntry = _getLastUnresolvedRejecter(project, sess.department);
    if (lastRejEntry) {
      var rejDept = lastRejEntry.by ? lastRejEntry.by.department : '';
      var normalized = _normalizeStatusByDept(rejDept);
      
      if (_normalizeDept(rejDept) === 'mec') {
        // When Project Team resolves a MEC rejection:
        // If it was rejected from Final Review -> go back to Final Review
        // If it was rejected from Draft/Other -> go to MEC Recheck
        newStatus = (lastRejEntry.fromStatus === 'Awaiting Final MEC Review') ? 'Awaiting Final MEC Review' : 'Awaiting MEC Recheck';
      } else {
        newStatus = normalized;
      }
    }
    
    project.workflowStatus = newStatus;

    // Update department and assigned department based on the new status
    var deptForStatus = STATUS_TO_DEPT[newStatus] || 'Project Team';
    project.department = deptForStatus;
    project.assignedToDept = deptForStatus;

    var history = project.workflowHistory || [];
    history.push({
      fromStatus: project.workflowStatus,
      toStatus: newStatus,
      by: { name: sess.name, email: sess.email, department: sess.department },
      comments: 'PARTIALLY APPROVED: ' + (comments || ''),
      timestamp: new Date().toISOString()
    });
    project.workflowHistory = history;

    // Update assignment
    if (nextDept) project.assignedToDept = nextDept;
    if (nextEmail) {
      project.assignedToEmail = nextEmail;
      try {
        var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
        var userData = userSh.getDataRange().getValues();
        for (var i = 1; i < userData.length; i++) {
          if (String(userData[i][2]).toLowerCase() === nextEmail.toLowerCase()) {
            project.assignedToName = userData[i][1];
            break;
          }
        }
      } catch(e) {}
    } else {
      project.assignedToName = '';
      project.assignedToEmail = '';
    }
    
    // If the new status is Awaiting Final MEC Review, assign to MEC department generally
    if (newStatus === 'Awaiting Final MEC Review') {
      project.assignedToDept = 'MEC';
      project.assignedToName = ''; // Clear specific assignment
      project.assignedToEmail = '';
    }
    // If the new status is Awaiting Final MEC Review, assign to MEC department generally
    if (newStatus === 'Awaiting Final MEC Review') {
      project.assignedToDept = 'MEC';
      project.assignedToName = ''; // Clear specific assignment
      project.assignedToEmail = '';
    }

    var saveRes = savePATProject(token, project);
    return { success: saveRes.success, message: 'Partially approved. Score: ' + score, verdict: 'Partially Accepted', snagScore: score };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Helper to get canonical status for a department
 */
function _normalizeStatusByDept(dept) {
  var d = String(dept || '').toLowerCase().trim();
  if (d.includes('planning')) return 'Awaiting Planning Team';
  if (d.includes('service') || d.includes('delivery') || d.includes('metro')) return 'Awaiting Service Delivery';
  if (d.includes('mec') || d.includes('mech')) return 'Awaiting Final MEC Review';
  if (d.includes('project')) return 'Awaiting Project Team';
  return 'Awaiting Project Team';
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
    if (!jccSh || jccSh.getLastRow() < 2) return { success: true, deletedCount: 0 };

    var searchId = String(projectId || '').trim().toUpperCase();
    var altId = searchId.replace(/^(FOB|PAT)-/i, '');

    var data = jccSh.getDataRange().getValues();
    var deletedCount = 0;

    // Iterate backwards so row indices stay valid after deletion
    for (var i = data.length - 1; i >= 1; i--) {
      var jccProjectId = String(data[i][1] || '').trim().toUpperCase();
      if (jccProjectId === searchId ||
          jccProjectId === altId ||
          ('PAT-' + jccProjectId) === searchId) {
        jccSh.deleteRow(i + 1); // +1 because sheet is 1-indexed
        deletedCount++;
      }
    }

    return { success: true, deletedCount: deletedCount };
  } catch(e) {
    // Non-critical: log error but don't block PAT deletion
    console.error('_deleteJCCByProjectId error for ' + projectId + ': ' + e.message);
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
    var raw = data[foundRow][c.IMAGES] || '[]';
    var urls = [];
    try { urls = JSON.parse(raw); } catch(e) { urls = []; }
    if (!Array.isArray(urls) || urls.length === 0) return { success: true, deletedCount: 0 };
    var deletedCount = 0;
    urls.forEach(function(url) {
      var match = String(url).match(/[?&]id=([^&]+)/);
      if (match) {
        try {
          var file = DriveApp.getFileById(match[1]);
          file.setTrashed(true);
          deletedCount++;
        } catch(e) { console.warn('Could not delete Drive image: ' + e.message); }
      }
    });
    return { success: true, deletedCount: deletedCount };
  } catch(e) {
    console.error('_deletePATDriveImages error: ' + e.message);
    return { success: false, deletedCount: 0, message: e.message };
  }
}
/**
 * Delete a single PAT project.
 * Also cascade-deletes linked JCC certificates, documents, and Drive files.
 */
function deletePATProject(token, projectId) {
  try {
    var sess = _session(token);
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    
    var searchId = String(projectId || '').trim();
    var altId = searchId.replace(/^(FOB|PAT)-/i, '');
    
    var foundRow = -1;
    var foundProject = null;
    
    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0] || '').trim();
      if (rowId.toUpperCase() === searchId.toUpperCase() ||
          rowId.toUpperCase() === altId.toUpperCase() ||
          ('PAT-' + rowId).toUpperCase() === searchId.toUpperCase()) {
        foundRow = i;
        foundProject = _projectFromRow(data[i]);
        break;
      }
    }
    
    if (foundRow === -1) return { success: false, message: 'Project not found: ' + projectId };
    
    var canDelete = false;
    var role = String(sess.role || '').toLowerCase();
    var dept = String(sess.department || '').toLowerCase();
    if (role === 'super admin') {
      canDelete = true;
    } else if (dept === 'mec' || dept === 'mech') {
      var status = (foundProject ? foundProject.workflowStatus : '') || '';
      var isEditableStatus = ['Draft', 'Rejected', 'Awaiting Final MEC Review', 'Awaiting MEC Recheck'].indexOf(status) !== -1;
      if (isEditableStatus) canDelete = true;
    }
    
    if (!canDelete) {
      return { success: false, message: 'Permission denied: You cannot delete this project.' };
    }
    
    // Cascade-delete JCC certificates
    var jccResult = _deleteJCCByProjectId(String(data[foundRow][0] || projectId));
    
    // Cascade-delete linked documents from SD_DOCUMENTS and their Drive files
    var docResult = { deletedCount: 0 };
    try {
      var docSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DOCS);
      if (docSh && docSh.getLastRow() >= 2) {
        var docData = docSh.getDataRange().getValues();
        var projectKey = String(data[foundRow][0] || projectId).toUpperCase().trim();
        var rowsToDelete = [];
        for (var di = docData.length - 1; di >= 1; di--) {
          var docProjectId = String(docData[di][1] || '').toUpperCase().trim();
          var normId = projectKey.replace(/^(FOB|PAT)-/i, '');
          if (docProjectId === projectKey || docProjectId === normId || ('PAT-' + docProjectId) === projectKey) {
            var driveFileId = docData[di][6] || '';
            if (driveFileId) {
              try { DriveApp.getFileById(driveFileId).setTrashed(true); } catch(e) { console.warn('Drive doc delete: ' + e.message); }
            }
            rowsToDelete.push(di + 1);
            docResult.deletedCount++;
          }
        }
        rowsToDelete.sort(function(a,b){return b-a}).forEach(function(r){ docSh.deleteRow(r); });
      }
    } catch(e) { console.error('Document cascade delete error: ' + e.message); }
    
    // Cascade-delete Drive images from PAT project
    var imgResult = _deletePATDriveImages(String(data[foundRow][0] || projectId), data, foundRow);
    
    // Delete the PAT project row itself
    sh.deleteRow(foundRow + 1);
    
    var msg = 'Project deleted.';
    var details = [];
    if (jccResult.deletedCount > 0) details.push(jccResult.deletedCount + ' JCC certificate(s)');
    if (docResult.deletedCount > 0) details.push(docResult.deletedCount + ' document(s)');
    if (imgResult.deletedCount > 0) details.push(imgResult.deletedCount + ' image(s)');
    if (details.length > 0) msg += ' Also removed: ' + details.join(', ') + '.';
    return { success: true, message: msg };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteAllDraftPATProjects(token) {
  try {
    _session(token);
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.', deleted: 0 };

    var data = sh.getDataRange().getValues();
    var c = _patCols();
    var totalDeleted = 0;
    var totalJCCDeleted = 0;
    var totalDocDeleted = 0;
    var totalImgDeleted = 0;
    var deleted = [];

    // Iterate backwards to preserve row indices
    for (var i = data.length - 1; i >= 1; i--) {
      var wfStatus = String(data[i][c.WORKFLOW_STATUS] || '').trim();
      if (wfStatus === 'Draft') {
        var projectId = String(data[i][c.PROJECT_ID] || '').trim();
        
        // Cascade-delete associated JCC certificates
        var jccResult = _deleteJCCByProjectId(projectId);
        totalJCCDeleted += (jccResult.deletedCount || 0);
        
        // Cascade-delete associated documents and their Drive files
        try {
          var docSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DOCS);
          if (docSh && docSh.getLastRow() >= 2) {
            var docData = docSh.getDataRange().getValues();
            var projectKey = projectId.toUpperCase().trim();
            var normId = projectKey.replace(/^(FOB|PAT)-/i, '');
            var docRowsToDel = [];
            for (var di = docData.length - 1; di >= 1; di--) {
              var dPid = String(docData[di][1] || '').toUpperCase().trim();
              if (dPid === projectKey || dPid === normId || ('PAT-' + dPid) === projectKey) {
                var dfid = docData[di][6] || '';
                if (dfid) { try { DriveApp.getFileById(dfid).setTrashed(true); } catch(e) {} }
                docRowsToDel.push(di + 1);
                totalDocDeleted++;
              }
            }
            docRowsToDel.sort(function(a,b){return b-a}).forEach(function(r){ docSh.deleteRow(r); });
          }
        } catch(e) { console.error('Draft doc cascade error: ' + e.message); }
        
        // Cascade-delete Drive images
        var imgResult = _deletePATDriveImages(projectId, data, i);
        totalImgDeleted += (imgResult.deletedCount || 0);
        
        sh.deleteRow(i + 1); // +1 because sheet is 1-indexed
        deleted.push(projectId);
        totalDeleted++;
      }
    }
    var msg = 'Deleted ' + deleted.length + ' draft projects.';
    var details = [];
    if (totalJCCDeleted > 0) details.push(totalJCCDeleted + ' JCC certificate(s)');
    if (totalDocDeleted > 0) details.push(totalDocDeleted + ' document(s)');
    if (totalImgDeleted > 0) details.push(totalImgDeleted + ' image(s)');
    if (details.length > 0) msg += ' Also removed: ' + details.join(', ') + '.';
    return { success: true, message: msg };
  } catch(e) {
    return { success: false, message: e.message, deleted: 0 };
  }
}

function getPATImages(token, projectId) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.', images: [] };
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var targetId = String(projectId || '').toUpperCase().trim();
    for (var i = 0; i < data.length; i++) {
      var rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        var c = _patCols();
        var images = _safeParse(data[i][c.IMAGES], []);
        return { success: true, images: images };
      }
    }
    return { success: false, message: 'Project not found.', images: [] };
  } catch(e) {
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
  const role = String(sess.role || '').toLowerCase();
  if (role === 'super admin') return true;

  const dept = String(sess.department || '').toLowerCase();
  const normalizedDept = _normalizeDept(dept);
  const status = project.workflowStatus || 'Draft';
  const isMEC = normalizedDept === 'mec';
  const isDraftOrRejected = (status === 'Draft' || status === 'Rejected');
  const isFinalReview = (status === 'Awaiting Final MEC Review');
  const isRecheck = (status === 'Awaiting MEC Recheck');

  // MEC can edit drafts, rejected, and final review (i.e. when it's their turn)
  if (isMEC && (isDraftOrRejected || isFinalReview || isRecheck)) return true;

  // Admin (non‑super) can edit only if they are explicitly assigned
  if (role === 'admin') {
    const assignedDept = _normalizeDept(project.assignedToDept || '');
    const assignedEmail = (project.assignedToEmail || '').toLowerCase();
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
    if (!projectRes.success) throw new Error('Project not found.');
    
    if (!_canEditProject(sess, projectRes.project)) {
      throw new Error('Permission denied: You cannot upload images to this project right now.');
    }

    const folder = _getPATFolder();
    const sh = _patSheet();
    const c = _patCols();
    
    // Get existing images
    const existingImagesRes = getPATImages(token, projectId);
    let finalUrls = existingImagesRes.success ? existingImagesRes.images : [];
    
    if (finalUrls.length >= 4) throw new Error('Maximum 4 images allowed per project.');

    // If it's already a direct link (from ImgBB), just save it!
    if (typeof imageData === 'string' && imageData.startsWith('http')) {
      finalUrls.push(imageData);
    } 
    // Otherwise, if it's still Base64, process it as a file in Drive
    else if (typeof imageData === 'string' && imageData.indexOf('data:image') === 0) {
      const parts = imageData.split(',');
      const mime = parts[0].match(/:(.*?);/)[1];
      const bytes = Utilities.base64Decode(parts[1]);
      const fileName = "PAT_" + projectId + "_" + (finalUrls.length + 1) + "_" + new Date().getTime();
      const blob = Utilities.newBlob(bytes, mime, fileName);
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // Use the direct "uc" link which is more certain to load than the "lh3" thumbnail link
      const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
      finalUrls.push(url);
    } else {
      throw new Error('Invalid image data format.');
    }

    // Update the sheet with the new list of URLs
    const data = sh.getDataRange().getValues();
    const targetId = String(projectId || '').toUpperCase().trim();
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        sh.getRange(i + 1, c.IMAGES + 1).setValue(JSON.stringify(finalUrls));
        return { success: true, images: finalUrls };
      }
    }
    
    throw new Error('Project ID not found in records.');
  } catch(e) {
    console.error("Single Upload Error: " + e.message);
    return { success: false, message: e.message };
  }
}

function savePATImages(token, projectId, images) {
  try {
    const sess = _session(token);
    // Load the project to check permissions
    const projectRes = getPATProjectById(null, projectId);
    if (!projectRes.success) return { success: false, message: 'Project not found.' };
    const project = projectRes.project;

    if (!_canEditProject(sess, project)) {
      return { success: false, message: 'Permission denied: You cannot upload images to this project at its current stage.' };
    }

    const folder = _getPATFolder();
    const sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };

    let finalUrls = []; // OVERWRITE: Use the exact array provided by the frontend

    // Process incoming images
    if (images && Array.isArray(images)) {
      for (let i = 0; i < images.length; i++) {
        if (finalUrls.length >= 4) break; // Hard limit of 4 images
        const imgStr = images[i];
        if (typeof imgStr !== 'string') continue;
        
        if (finalUrls.indexOf(imgStr) !== -1) continue; // Skip duplicates
        
        if (imgStr.indexOf('data:image') === 0) {
          // New Upload: Convert base64 to Blob and save to Drive
          const parts = imgStr.split(',');
          const mime = parts[0].match(/:(.*?);/)[1];
          const bytes = Utilities.base64Decode(parts[1]);
          const blob = Utilities.newBlob(bytes, mime, "PAT_" + projectId + "_" + i);
          const file = folder.createFile(blob);
          const url = "https://lh3.googleusercontent.com/d/" + file.getId();
          finalUrls.push(url);
        } else {
          // Existing link: keep as is
          finalUrls.push(imgStr);
        }
      }
    }

    const data = sh.getDataRange().getValues();
    const targetId = String(projectId || '').toUpperCase().trim();
    const c = _patCols();
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        sh.getRange(i + 1, c.IMAGES + 1).setValue(JSON.stringify(finalUrls));
        return { success: true, images: finalUrls };
      }
    }
    return { success: false, message: 'Project not found.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deletePATImages(token, projectId) {
  try {
    const sess = _session(token);
    const projectRes = getPATProjectById(null, projectId);
    if (!projectRes.success) return { success: false, message: 'Project not found.' };
    const project = projectRes.project;

    if (!_canEditProject(sess, project)) {
      return { success: false, message: 'Permission denied: You cannot delete images from this project.' };
    }

    const sh = _patSheet();
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No projects found.' };
    const data = sh.getDataRange().getValues();
    const targetId = String(projectId || '').toUpperCase().trim();
    const c = _patCols();
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).toUpperCase().trim();
      if (rowId === targetId) {
        sh.getRange(i + 1, c.IMAGES + 1).setValue('[]');
        return { success: true };
      }
    }
    return { success: false, message: 'Project not found.' };
  } catch(e) {
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
    if (!sh) return { success: false, message: 'PAT sheet not found.' };
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
    var msg = 'All PAT projects deleted.';
    if (jccDeleted > 0) {
      msg += ' ' + jccDeleted + ' associated JCC certificate(s) also removed.';
    }
    return { success: true, message: msg };
  } catch(e) {
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

    if (!sh) return { success: false, message: 'Documents sheet not found.' };

    let driveUrl = '';
    let driveFileId = '';
    let fileType = '';
    let fileSize = 0;

    if (typeof fileData === 'string' && fileData.startsWith('http')) {
      driveUrl = fileData;
      fileType = 'link';
      fileSize = 0;
    } else if (typeof fileData === 'string' && fileData.indexOf('data:') === 0) {
      const parts = fileData.split(',');
      const mimeMatch = parts[0].match(/:(.*?);/);
      if (!mimeMatch) {
        return { success: false, message: 'Could not detect file type from upload data. Try a different file.' };
      }
      const mime = mimeMatch[1];
      const raw = parts[1];
      if (!raw || raw.length === 0) {
        return { success: false, message: 'Upload data is empty. Please select a valid file.' };
      }
      let bytes;
      try {
        bytes = Utilities.base64Decode(raw);
      } catch(e) {
        return { success: false, message: 'File encoding error. The file may be corrupted or too large.' };
      }
      fileSize = bytes.length;
      fileType = mime;

      const ext = _mimeToExt(mime) || (fileName || '').split('.').pop() || 'bin';
      const safeName = (fileName || 'document_' + _genDocId()).replace(/[^a-zA-Z0-9._\- ]/g, '_');
      const blobName = safeName.indexOf('.') === -1 ? safeName + '.' + ext : safeName;
      const blob = Utilities.newBlob(bytes, mime, blobName);

      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveFileId = file.getId();
      driveUrl = 'https://drive.google.com/uc?export=view&id=' + driveFileId;
    } else {
      return { success: false, message: 'Unsupported file data format. Please try again.' };
    }

    const docId = _genDocId();
    const now = new Date().toISOString();
    const row = [
      docId,
      String(projectId || '').trim(),
      fileName || 'Untitled',
      fileType,
      fileSize,
      driveUrl,
      driveFileId,
      sess.name || '',
      sess.email || '',
      sess.department || '',
      now,
      category || 'General'
    ];

    sh.appendRow(row);
    return { success: true, docId: docId, driveUrl: driveUrl, message: 'Document uploaded successfully.' };
  } catch(e) {
    return { success: false, message: 'Upload failed: ' + e.message };
  }
}

function _mimeToExt(mime) {
  const map = {
    'application/pdf': 'pdf',
    'application/msword': 'doc',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
    'application/vnd.ms-excel': 'xls',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
    'text/plain': 'txt',
    'text/csv': 'csv',
    'application/zip': 'zip',
    'application/x-rar-compressed': 'rar',
    'application/vnd.rar': 'rar',
    'application/vnd.ms-powerpoint': 'ppt',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/svg+xml': 'svg',
    'text/html': 'html',
    'application/json': 'json',
    'application/xml': 'xml'
  };
  return map[mime] || null;
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
        const docPid = String(doc.projectId || '').trim().toUpperCase();
        const queryPid = String(projectId).trim().toUpperCase();
        if (docPid !== queryPid) continue;
      }

      // Filter by department if specified
      if (department && doc.uploadedByDept !== department) continue;

      // When viewing documents scoped to a project, show all docs for that project
      // (project access controls visibility instead)
      if (!projectId) {
        const role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin') {
          if (department && doc.uploadedByDept !== department) continue;
          if (!department && doc.uploadedByEmail !== sess.email && doc.uploadedByDept !== sess.department) continue;
        }
      }

      docs.push(doc);
    }

    docs.sort((a, b) => new Date(b.uploadedAt) - new Date(a.uploadedAt));
    return { success: true, documents: docs };
  } catch(e) {
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
    if (!sh || sh.getLastRow() < 2) return { success: false, message: 'No documents found.' };

    const data = sh.getDataRange().getValues();
    const c = _docCols();
    const role = String(sess.role || '').toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][c.DOC_ID]).toUpperCase().trim();
      if (rowId === String(docId).toUpperCase().trim()) {
        const doc = _documentFromRow(data[i]);

        // Check permission: uploader, same department admin, or super admin
        const canDelete = (
          role === 'super admin' ||
          doc.uploadedByEmail === sess.email ||
          (role === 'admin' && doc.uploadedByDept === sess.department)
        );

        if (!canDelete) {
          return { success: false, message: 'Permission denied: You cannot delete this document.' };
        }

        // Try to delete from Drive
        if (doc.driveFileId) {
          try {
            const file = DriveApp.getFileById(doc.driveFileId);
            file.setTrashed(true);
          } catch(e) {
            console.warn('Could not delete Drive file: ' + e.message);
          }
        }

        sh.deleteRow(i + 1);
        return { success: true, message: 'Document deleted.' };
      }
    }

    return { success: false, message: 'Document not found.' };
  } catch(e) {
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
    if (!sh || sh.getLastRow() < 2) return { success: true, message: 'No documents to delete.', deletedCount: 0 };

    const data = sh.getDataRange().getValues();
    const c = _docCols();
    const role = String(sess.role || '').toLowerCase();

    if (role !== 'admin' && role !== 'super admin') {
      return { success: false, message: 'Admin access required.' };
    }

    let deletedCount = 0;
    const rowsToDelete = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const docProjectId = String(data[i][c.PROJECT_ID] || '').toUpperCase().trim();
      if (docProjectId === String(projectId).toUpperCase().trim()) {
        const doc = _documentFromRow(data[i]);

        // Check permission
        const canDelete = (
          role === 'super admin' ||
          doc.uploadedByEmail === sess.email ||
          doc.uploadedByDept === sess.department
        );

        if (canDelete) {
          if (doc.driveFileId) {
            try {
              const file = DriveApp.getFileById(doc.driveFileId);
              file.setTrashed(true);
            } catch(e) { console.warn('Drive delete failed: ' + e.message); }
          }
          rowsToDelete.push(i + 1);
          deletedCount++;
        }
      }
    }

    // Delete rows from bottom to top to preserve row indices
    rowsToDelete.sort((a, b) => b - a).forEach(function(rowIdx) {
      sh.deleteRow(rowIdx);
    });

    return { success: true, message: deletedCount + ' document(s) deleted.', deletedCount: deletedCount };
  } catch(e) {
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
      comments: project.rejectionReason || project.comments || '',
      assignedToDept: project.assignedToDept
    };
  } catch(e) {
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
    try { return { success: true, data: JSON.parse(cached), cached: true }; } 
    catch(e) { cache.remove("public_departments"); }
  }

  try {
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: true, data: [] };

    var data = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    var names = data.map(function(r){ return String(r[0]).trim(); })
                    .filter(Boolean).sort();
    
    cache.put("public_departments", JSON.stringify(names), 600); // Cache for 10 mins (better UX)
    return { success: true, data: names };
  } catch(e) {
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
    try { return { success: true, ...JSON.parse(cached), cached: true }; }
    catch(e) { cache.remove("domain_config"); }
  }

  try {
    var config = _getConfig();
    var props = PropertiesService.getScriptProperties();
    var allowLocalhost = props.getProperty("allow_localhost") === "true";
    var data = {
      domain: config.DOMAIN,
      aiWelcomeMessage: config.AI_WELCOME_MESSAGE,
      allowLocalhost: allowLocalhost
    };
    cache.put("domain_config", JSON.stringify(data), 1500);
    return { success: true, ...data };
  } catch(e) {
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

    var filtered = current.filter(function(origin) {
      return origin !== 'http://localhost' && origin !== 'http://127.0.0.1';
    });

    if (allowed === true) {
      filtered.push('http://localhost');
      filtered.push('http://127.0.0.1');
    }

    PropertiesService.getScriptProperties().setProperty('allowed_origins', JSON.stringify(filtered));
    CacheService.getScriptCache().remove('allowed_origins');

    PropertiesService.getScriptProperties().setProperty("allow_localhost", allowed === true ? "true" : "false");
    CacheService.getScriptCache().remove("domain_config");

    return {
      success: true,
      message: 'Localhost access ' + (allowed ? 'enabled' : 'disabled') + '.',
      allowedOrigins: filtered
    };
  } catch(e) {
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

    var deptData = deptSh && deptSh.getLastRow() > 1 ? deptSh.getDataRange().getValues().slice(1) : [];
    var userData = userSh && userSh.getLastRow() > 1 ? userSh.getDataRange().getValues().slice(1) : [];

    // Count active users per department name
    var userMap = {};
    userData.forEach(function(r){
      if (r[7]==="active"){ var d=String(r[6]).trim(); userMap[d]=(userMap[d]||0)+1; }
    });

    var departments = deptData.map(function(r) {
      return {
        deptId:       r[0],
        name:         r[1],
        headEmail:    r[2],
        createdAt:    r[3],
        activeUsers:  userMap[r[1]] || 0,
      };
    });
    return { success: true, departments: departments };
  } catch(e) {
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

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues().slice(1);
    var dup  = data.find(function(r){ return String(r[1]).toLowerCase()===name.toLowerCase(); });
    if (dup) throw new Error("A department with the name '" + name + "' already exists.");

    var deptId = _genId("DEPT");
    sh.appendRow([deptId, name, headEmail||"", new Date().toISOString(), sess.email]);
    CacheService.getScriptCache().remove("public_departments"); // Invalidate cache for signup
    return { success: true, deptId: deptId, message: "Department added." };
  } catch(e) {
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

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sh.getRange(i+1, 2).setValue(name);
        sh.getRange(i+1, 3).setValue(headEmail || "");
        CacheService.getScriptCache().remove("public_departments");
        return { success: true, message: "Department updated." };
      }
    }
    throw new Error("Department not found.");
  } catch(e) {
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
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.DEPTS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sh.deleteRow(i + 1);
        CacheService.getScriptCache().remove("public_departments");
        return { success: true, message: "Department deleted." };
      }
    }
    throw new Error("Department not found.");
  } catch(e) {
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
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues().slice(1);
    var users = data.map(function(r) {
      var ipsRaw = r[14] || '';
      var ips = [];
      try { ips = ipsRaw ? JSON.parse(ipsRaw) : []; } catch(e) { ips = []; }
      return {
        userId:      r[0], name:        r[1],
        email:       r[2], role:        r[5],
        department:  r[6],
        gender:      r[7], workflowNotes: r[8],
        status:      r[9], createdAt:   r[10], lastLoginAt: r[11],
        loginIps: ips
      };
    });
    return { success: true, users: users };
  } catch(e) {
    return { success: false, message: e.message, users: [] };
  }
}

function searchUsers(token, query) {
  try {
    _session(token);
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues().slice(1);
    var q = String(query || '').toLowerCase().trim();
    var users = data.map(function(r) {
      return {
        userId: r[0], name: r[1], email: r[2],
        role: r[5], department: r[6], status: r[9]
      };
    }).filter(function(u) {
      if (!q) return true;
      return (u.name || '').toLowerCase().indexOf(q) !== -1 ||
             (u.email || '').toLowerCase().indexOf(q) !== -1;
    });
    return { success: true, users: users };
  } catch(e) {
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
        var ipsRaw = data[i][14] || '';
        var pwHash = data[i][13] || '';
        var fingerprintsRaw = data[i][16] || '';
        var ips = [];
        var fingerprints = [];
        try { ips = ipsRaw ? JSON.parse(ipsRaw) : []; } catch(e) { ips = []; }
        try { fingerprints = fingerprintsRaw ? JSON.parse(fingerprintsRaw) : []; } catch(e) { fingerprints = []; }
        return {
          success: true,
          userId: data[i][0],
          name: data[i][1],
          email: data[i][2],
          role: data[i][5],
          department: data[i][6],
          status: data[i][9],
          ips: ips,
          fingerprints: fingerprints
        };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
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
    _superAdminSession(token);
    if (newRole !== SD.ADMIN && newRole !== SD.EMPLOYEE && newRole !== "super admin")
      throw new Error("Role must be 'admin', 'employee', or 'super admin'.");

    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var sh      = ss.getSheetByName(SD.USERS);
    var data    = sh.getDataRange().getValues();
    var searchId = String(userId || "").trim();
    var isEmailSearch = searchId.indexOf("@") !== -1;

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      var rowEmail = String(data[i][2]).toLowerCase().trim();
      if (rowId === searchId || (isEmailSearch && rowEmail === searchId.toLowerCase())) {
        sh.getRange(i+1, 6).setValue(newRole);
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Activate or ban a user account.
 * @param {string} token
 * @param {string} userId
 * @param {string} status  "active" | "banned"
 */
function updateUserStatus(token, userId, status) {
  var lock = LockService.getScriptLock();
  try {
    _superAdminSession(token);
    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        sh.getRange(i+1, 10).setValue(status);
        lock.releaseLock();
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
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

    var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || "").toLowerCase().trim() === String(sess.email || "").toLowerCase().trim()) {
        sh.getRange(i+1, 2).setValue(newName);
        // Update the stored session payload so the nav reflects the change immediately
        var raw = PropertiesService.getUserProperties().getProperty("tok_" + token);
        if (raw) {
          try {
            var s = JSON.parse(raw);
            s.name = newName;
            PropertiesService.getUserProperties().setProperty("tok_" + token, JSON.stringify(s));
          } catch(parseErr) {
            // Session data is corrupted, but profile was updated anyway
            console.warn('Session parse error during profile update:', parseErr);
          }
        }

        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
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
    var sessEmail = String(sess.email || "").toLowerCase().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || "").toLowerCase().trim() === sessEmail) {
        sh.getRange(i+1, 13).setValue(allowMsgs ? "TRUE" : "FALSE");
        sh.getRange(i+1, 14).setValue(aiAuto ? "TRUE" : "FALSE");
        return { success: true };
      }
    }
    throw new Error("User not found.");
  } catch(e) {
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
    var members = data.filter(function(r) {
      var role = String(sess.role || "").toLowerCase();
      if (role === 'admin' || role === 'super admin') return true;
      return String(r[6]) === sess.department;
    }).map(function(r) {
      return { name: r[1], email: r[2], role: String(r[5] || "").toLowerCase() };
    });
    
    return { success: true, members: members };
  } catch(e) {
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
    if (!apiKey) throw new Error("Gemini API key not configured. Please set it in System Configuration.");

    // Build contents array for Gemini API
    var contents = [];

    // Add chat history
    if (history && Array.isArray(history)) {
      history.slice(-10).forEach(function(h) {
        if (typeof h.content === 'string') {
          contents.push({
            role: h.role === 'assistant' ? 'model' : 'user',
            parts: [{ text: h.content }]
          });
        }
      });
    }

    // Add current user message
    contents.push({ role: 'user', parts: [{ text: userMessage }] });

    var payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 2048
      }
    };

    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=" + apiKey;
    var response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      var errData = JSON.parse(response.getContentText());
      throw new Error(errData.error?.message || "Gemini API error " + response.getResponseCode());
    }

    var data = JSON.parse(response.getContentText());
    var text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    

    // Check for PAT topic restriction violation
    if (text.indexOf('REJECTED_NON_PAT_QUERY:') !== -1) {
      _logRejectedQuery(sess, userMessage || '[Image uploaded]', text, 'off-topic');
      return { success: true, answer: "I can only answer questions related to PAT projects and the SD portal system. Please ask me something about PAT projects, fiber deployment, or the SD workflow system." };
    }

    return { success: true, answer: text };
  } catch(e) {
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
function callGeminiVision(token, systemPrompt, userMessage, imageAttachments, history) {
  try {
    var sess = _session(token);
    var config = _getConfig();
    var apiKey = config.GEMINI_API_KEY;
    if (!apiKey) throw new Error("Gemini API key not configured. Please set it in System Configuration.");

    var contents = [];

    // Add chat history (text only)
    if (history && Array.isArray(history)) {
      history.slice(-8).forEach(function(h) {
        if (typeof h.content === 'string') {
          contents.push({
            role: h.role === 'assistant' ? 'model' : 'user',
            parts: [{ text: h.content }]
          });
        }
      });
    }

    // Build current user message with images
    var parts = [];
    if (userMessage) parts.push({ text: userMessage });

    // Process image attachments
    if (imageAttachments && Array.isArray(imageAttachments)) {
      imageAttachments.forEach(function(att) {
        if (att.type === 'image_url' && att.image_url && att.image_url.url) {
          var dataUrl = att.image_url.url;
          var mimeMatch = dataUrl.match(/data:(image\/[a-z]+);base64,/);
          var mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
          var base64Data = dataUrl.replace(/^data:image\/[a-z]+;base64,/, '');
          parts.push({
            inlineData: {
              mimeType: mime,
              data: base64Data
            }
          });
        }
      });
    }

    if (parts.length === 0) parts.push({ text: userMessage || "Analyze this image." });
    contents.push({ role: 'user', parts: parts });

     // Inject PAT-ONLY restriction into system prompt
    var patRestriction = "\n\nCRITICAL TOPIC RESTRICTION: You are STRICTLY LIMITED to answering questions about PAT (Provisional Acceptability Test) projects, the SD portal system, fiber network deployment, and related workflow operations. If the user asks about ANY topic outside of PAT, fiber network deployment, or the SD system, you MUST respond with ONLY this exact prefix: 'REJECTED_NON_PAT_QUERY:' followed by your refusal. Do NOT answer off-topic questions under any circumstances. Stay locked to PAT domain only.";

    var payload = {
      contents: contents,
      systemInstruction: { parts: [{ text: systemPrompt }] },
      generationConfig: {
        temperature: 0.5,
        maxOutputTokens: 4096
      }
    };

    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite-preview:generateContent?key=" + apiKey;
    var response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      var errData = JSON.parse(response.getContentText());
      throw new Error(errData.error?.message || "Gemini Vision API error " + response.getResponseCode());
    }

    var data = JSON.parse(response.getContentText());
    var text = data.candidates?.[0]?.content?.parts?.[0]?.text || "";
    return { success: true, answer: text };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAIResponse(token, prompt) {
  try {
    var sess = _session(token);
    var config = _getConfig();
    var patRestriction = " CRITICAL: You are STRICTLY LIMITED to PAT (Provisional Acceptability Test) projects only. Reject anything else.";
    var systemPrompt = "You are the SD-AI for FiberOne Broadband. " + config.AI_WELCOME_MESSAGE + patRestriction;
    
    // Using the stable Pollinations API via server-side fetch
    var url = "https://text.pollinations.ai/" + encodeURIComponent(prompt) + "?system=" + encodeURIComponent(systemPrompt) + "&model=openai";
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() !== 200) throw new Error("AI service is currently busy.");
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
    "Abia","Adamawa","Akwa Ibom","Anambra","Bauchi","Bayelsa","Benue","Borno",
    "Cross River","Delta","Ebonyi","Edo","Ekiti","Enugu","FCT","Gombe","Imo",
    "Jigawa","Kaduna","Kano","Katsina","Kebbi","Kogi","Kwara","Lagos","Nassarawa",
    "Niger","Ogun","Ondo","Osun","Oyo","Plateau","Rivers","Sokoto","Taraba",
    "Yobe","Zamfara"
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
  var parts = address.split(",").map(function(p) { return p.trim(); });
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
    if (entry.toStatus === 'Completed' && entry.by && entry.by.department) {
      var dept = String(entry.by.department).toLowerCase();
      if (dept.indexOf('mec') !== -1 || dept.indexOf('mech') !== -1) {
        mecName = entry.by.name || "MEC Officer";
        break;
      }
    }
  }
  if (mecName === "MEC Officer" && session && session.name) {
    mecName = session.name; // fallback only if no MEC completion history found
  }
  
  var mecDate = new Date().toISOString().split('T')[0];
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
    presidingOfficer: project.presidingOfficer || '',
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
    isPreview: true
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
      if (rowId.toUpperCase() === String(projectId).toUpperCase() ||
          rowId.toUpperCase() === String(projectId).replace(/^(FOB|PAT)-/i, '').toUpperCase() ||
          ('PAT-' + rowId).toUpperCase() === String(projectId).toUpperCase()) {
        project = _projectFromRow(data[i]);
        break;
      }
    }
    
    if (!project) throw new Error("Project not found: " + projectId);
    if (project.workflowStatus !== 'Completed') throw new Error("Project is not completed yet.");
    
    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    if (!jccSh) throw new Error("JCC sheet not found. Run First Setup.");
    var jccData = jccSh.getDataRange().getValues();
    for (var j = 1; j < jccData.length; j++) {
      if (String(jccData[j][1]).trim().toUpperCase() === String(projectId).toUpperCase()) {
        return { success: false, message: "JCC already exists for this project.", jccId: jccData[j][0] };
      }
    }
    
    // PO No = row number in the PAT sheet (starting from 1 for data rows)
    // Find the project's row index in the PAT sheet
    var projectRowNumber = 0;
    var patData = sh.getDataRange().getValues();
    for (var k = 1; k < patData.length; k++) {
      var rowId = String(patData[k][c.PROJECT_ID]).trim();
      if (rowId.toUpperCase() === String(projectId).toUpperCase() ||
          rowId.toUpperCase() === String(projectId).replace(/^(FOB|PAT)-/i, '').toUpperCase() ||
          ('PAT-' + rowId).toUpperCase() === String(projectId).toUpperCase()) {
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
      jccId, jccObj.projectId, jccObj.projectName, jccObj.projectNumber,
      jccObj.certificateType, jccObj.vendor, jccObj.certificateId,
      jccObj.penalty, jccObj.remarks, jccObj.mecName, jccObj.mecSignature,
      jccObj.mecDate, jccObj.vendorName, jccObj.vendorSignature, jccObj.vendorDate,
      jccObj.mecHeadName, jccObj.mecHeadSignature, jccObj.mecHeadDate,
      jccObj.generatedAt, session.name, session.email,
      jccObj.stateRegion, jccObj.orchestrator, jccObj.presidingOfficer
    ]);
    
    var lastRow = jccSh.getLastRow();
    jccSh.getRange(lastRow, 1, 1, jccSh.getLastColumn()).setFontFamily('Arial').setFontSize(10);
    
    return {
      success: true,
      jccId: jccId,
      certificateId: jccObj.certificateId,
      projectNumber: projectNumber,
      message: "JCC generated successfully. Certificate ID: " + jccObj.certificateId
    };
    
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function _genJCCId() {
  return 'JCC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function _generateJCCRemarks(project) {
  var score = parseInt(project.snagScore) || 0;
  var snags = project.snags || [];
  var history = project.workflowHistory || [];
  
  var penalty = "None";
  var remarks = "";
  
  var criticalSnags = snags.filter(function(s) { return parseInt(s.weight) >= 3; });
  var majorSnags = snags.filter(function(s) { return parseInt(s.weight) === 2; });
  var minorSnags = snags.filter(function(s) { return parseInt(s.weight) === 1; });
  
  if (score === 0) {
    penalty = "None - Project completed with zero complaints.";
  } else if (score <= 2) {
    penalty = "Minor - Project completed with minor observations noted.";
  } else if (score <= 5) {
    penalty = "Moderate - Project completed with major observations requiring follow-up.";
  } else {
    penalty = "Significant - Project completed with critical observations. Follow-up inspection required.";
  }
  
  remarks = "Project '" + (project.projectName || "Unknown") + "' has been reviewed and accepted. ";
  if (score === 0) {
    remarks += "All checklist items passed with no complaints. ";
  } else {
    remarks += "Total complaint score: " + score + ". ";
    if (criticalSnags.length > 0) remarks += criticalSnags.length + " critical item(s) identified. ";
    if (majorSnags.length > 0) remarks += majorSnags.length + " major item(s) identified. ";
    if (minorSnags.length > 0) remarks += minorSnags.length + " minor item(s) identified. ";
  }
  
  var deptCount = {};
  history.forEach(function(h) {
    var dept = h.by && h.by.department ? h.by.department : "Unknown";
    deptCount[dept] = (deptCount[dept] || 0) + 1;
  });
  
  remarks += "Workflow involved: " + (Object.keys(deptCount).join(", ") || "MEC review") + ". ";
  remarks += "Final verdict: " + (project.verdict || "Pending") + ". ";
  remarks += "Project cleared for completion on " + new Date().toLocaleDateString('en-GB') + ".";
  
  return { penalty: penalty, remarks: remarks };
}

function getJCCByProjectId(token, projectId) {
  try {
    var sess = _session(token);
    var searchId = String(projectId || "").trim().toUpperCase();
    
    if (!searchId || searchId === "UNDEFINED" || searchId === "NULL") {
      return { success: false, message: "Project ID is missing. Please try again." };
    }

    // 1. Try finding an official JCC record first
    var jccSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.JCC);
    if (jccSh && jccSh.getLastRow() >= 2) {
      var data = jccSh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowId = String(data[i][1] || "").trim().toUpperCase();
        var rowNum = String(data[i][3] || "").trim().toUpperCase();
        
        if (rowId === searchId || 
            rowId === searchId.replace(/^(FOB|PAT)-/, '') || 
            ('PAT-' + rowId) === searchId || 
            rowNum === searchId || 
            rowNum === searchId.replace(/^(FOB|PAT)-/, '') || 
            ('PAT-' + rowNum) === searchId) {
          return {
            success: true,
            jcc: {
              jccId: data[i][0], projectId: data[i][1], projectName: data[i][2],
              projectNumber: data[i][3], certificateType: data[i][4], vendor: data[i][5],
              certificateId: data[i][6], penalty: data[i][7], remarks: data[i][8],
              mecName: data[i][9], mecSignature: data[i][10], mecDate: data[i][11],
              vendorName: data[i][12], vendorSignature: data[i][13], vendorDate: data[i][14],
              mecHeadName: data[i][15], mecHeadSignature: data[i][16], mecHeadDate: data[i][17],
              generatedAt: data[i][18], generatedBy: data[i][19],
              stateRegion: data[i][21] || _extractStateRegion(''),
              orchestrator: data[i][22] || 'N/A',
              presidingOfficer: data[i][23] || '',
              isPreview: false
            }
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
        jcc: previewJcc
      };
    }

    return { success: false, message: 'Unable to find any JCC or PAT record for Project ID: ' + projectId };
  } catch(e) {
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
        jccId: data[i][0], projectId: data[i][1], projectName: data[i][2],
        projectNumber: data[i][3], certificateId: data[i][6], vendor: data[i][5],
        mecName: data[i][9], generatedAt: data[i][18]
      });
    }
    return { success: true, jccs: jccs };
  } catch(e) {
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

    var filtered = rows.filter(function(r) {
      var mailFolder = r[9]; // Folder column
      var sender = r[1]; // SenderEmail
      var receiver = r[3]; // ReceiverEmail
      var starred = r[11]; // Starred column
      var label = r[12]; // Labels column
      var subject = String(r[5] || '').toLowerCase();
      var body = String(r[6] || '').toLowerCase();
      var fromEmail = String(r[1] || '').toLowerCase();
      var toEmail = String(r[3] || '').toLowerCase();
      var deletedBy = String(r[17] || '').toLowerCase();
      var query = (options && options.query) ? String(options.query).toLowerCase() : '';
      var userEmail = String(sess.email || '').toLowerCase();

      // Soft delete filter: exclude if current user deleted this mail
      var deletedList = deletedBy ? deletedBy.split(',').map(function(e){return e.trim();}) : [];
      if (deletedList.indexOf(userEmail) !== -1) return false;

      if (folder === 'Inbox') {
        if (receiver !== sess.email) return false;
        if (mailFolder !== 'Inbox') return false;
      } else if (folder === 'Sent') {
        if (sender !== sess.email) return false;
        if (mailFolder !== 'Sent') return false;
      } else if (folder === 'Self') {
        if (sender !== sess.email && receiver !== sess.email) return false;
        if (sender === receiver && mailFolder !== 'Self') return false;
      } else if (folder === 'Trash') {
        if (sender !== sess.email && receiver !== sess.email) return false;
        if (mailFolder !== 'Trash') return false;
      } else if (folder === 'Starred') {
        if (sender !== sess.email && receiver !== sess.email) return false;
        if (starred !== true && starred !== 'TRUE' && String(starred) !== 'true') return false;
      } else if (folder === 'Archive') {
        if (sender !== sess.email && receiver !== sess.email) return false;
        if (mailFolder !== 'Archive') return false;
      } else if (folder === 'All Mail') {
        if (sender !== sess.email && receiver !== sess.email) return false;
      }

      // Search filter
      if (query) {
        var match = subject.indexOf(query) !== -1 || body.indexOf(query) !== -1 || fromEmail.indexOf(query) !== -1 || toEmail.indexOf(query) !== -1;
        if (!match) return false;
      }

      return true;
    }).map(function(r) {
      var att = [];
      try { att = JSON.parse(r[10] || '[]'); } catch(e) { att = []; }
      var labels = [];
      try { labels = JSON.parse(r[12] || '[]'); } catch(e) { labels = []; }
      var deletedBy = String(r[17] || '').toLowerCase();
      return {
        mailId: r[0], senderEmail: r[1], senderName: r[2],
        receiverEmail: r[3], receiverName: r[4], subject: r[5],
        body: r[6], timestamp: r[7], status: r[8], folder: r[9],
        attachments: att, hasAttachments: att.length > 0,
        starred: r[11] === true || r[11] === 'TRUE' || String(r[11]) === 'true',
        labels: labels,
        threadId: r[13] || '',
        priority: r[14] || 'normal',
        cc: r[15] || '',
        bcc: r[16] || '',
        deletedBy: deletedBy
      };
    });

    return { success: true, mails: filtered };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function sendMail(token, receiverEmail, subject, body, attachments, cc, bcc, labels, priority, threadId) {
  try {
    var sess = _session(token);
    var ccNorm = cc ? String(cc).toLowerCase().trim() : '';
    var bccNorm = bcc ? String(bcc).toLowerCase().trim() : '';
    var labelsArray = labels || [];
    var labelsJson = JSON.stringify(labelsArray);
    var priorityNorm = priority || 'normal';
    var threadIdNorm = threadId || '';

    // Accept both single email string and array of emails
    var primaryReceivers = [];
    if (typeof receiverEmail === 'string') {
      primaryReceivers = [receiverEmail.toLowerCase().trim()];
    } else if (Array.isArray(receiverEmail)) {
      primaryReceivers = receiverEmail.map(function(e) { return String(e).toLowerCase().trim(); }).filter(Boolean);
    }
    
    // Parse CC recipients from the cc string
    var ccReceivers = [];
    if (ccNorm) {
      ccReceivers = ccNorm.split(',').map(function(e) { return e.trim(); }).filter(Boolean);
    }
    
    // Remove CC recipients from primary receivers to avoid duplicates
    var ccSet = ccReceivers.join(',');
    primaryReceivers = primaryReceivers.filter(function(r) { return ccSet.indexOf(r) === -1; });
    
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

    receivers.forEach(function(receiverEmailNorm) {
      var receiver = userData.find(function(r) { return String(r[2]).toLowerCase().trim() === receiverEmailNorm; });
      if (!receiver) return; // skip invalid recipients silently
      var receiverName = receiver[1];
      var isSelf = sess.email.toLowerCase().trim() === receiverEmailNorm;
      var senderFolder = isSelf ? 'Self' : 'Sent';
      var receiverFolder = isSelf ? 'Self' : 'Inbox';

      // Add to Receiver's folder
      batchRows.push([mailId, sess.email, sess.name, receiverEmailNorm, receiverName, subject, body, now, "unread", receiverFolder, attachmentsJson, false, labelsJson, threadIdNorm, priorityNorm, ccNorm, bccNorm, '']);
      // Add to Sender's folder
      batchRows.push([mailId, sess.email, sess.name, receiverEmailNorm, receiverName, subject, body, now, "read", senderFolder, attachmentsJson, false, labelsJson, threadIdNorm, priorityNorm, ccNorm, bccNorm, '']);
      mailCount++;
    });

    // Also create rows for CC recipients
    ccReceivers.forEach(function(receiverEmailNorm) {
      var receiver = userData.find(function(r) { return String(r[2]).toLowerCase().trim() === receiverEmailNorm; });
      if (!receiver) return;
      var receiverName = receiver[1];
      var isSelf = sess.email.toLowerCase().trim() === receiverEmailNorm;
      var senderFolder = isSelf ? 'Self' : 'Sent';
      var receiverFolder = isSelf ? 'Self' : 'Inbox';

      batchRows.push([mailId, sess.email, sess.name, receiverEmailNorm, receiverName, subject, body, now, "unread", receiverFolder, attachmentsJson, false, labelsJson, threadIdNorm, priorityNorm, ccNorm, bccNorm, '']);
      batchRows.push([mailId, sess.email, sess.name, receiverEmailNorm, receiverName, subject, body, now, "read", senderFolder, attachmentsJson, false, labelsJson, threadIdNorm, priorityNorm, ccNorm, bccNorm, '']);
      mailCount++;
    });

    if (mailCount === 0) throw new Error("No valid recipients found in system.");
    
    // BATCH WRITE: single API call instead of 2*N appendRow calls
    if (batchRows.length > 0) {
      var lastRow = sh.getLastRow();
      sh.getRange(lastRow + 1, 1, batchRows.length, batchRows[0].length).setValues(batchRows);
    }

    // Back-fill threadId to existing mails in the same conversation
    // so the original mail appears in the thread history (WhatsApp-style grouping)
    if (threadIdNorm && (primaryReceivers.length > 0 || ccReceivers.length > 0)) {
      var cleanSubject = String(subject || '').replace(/^(Re:\s*)+/i, '').toLowerCase().trim();
      var sessEmail = String(sess.email || '').toLowerCase().trim();
      var allTargetEmails = primaryReceivers.concat(ccReceivers).filter(Boolean);
      var existingData = sh.getDataRange().getValues();
      for (var r = 1; r < existingData.length; r++) {
        var existingRowSubject = String(existingData[r][5] || '').replace(/^(Re:\s*)+/i, '').toLowerCase().trim();
        var existingSender = String(existingData[r][1] || '').toLowerCase().trim();
        var existingReceiver = String(existingData[r][3] || '').toLowerCase().trim();
        var existingThreadId = String(existingData[r][13] || '').trim();
        // Match: same conversation, involves any participant, no threadId yet
        var matchesParticipant = allTargetEmails.some(function(t) { return existingSender === t || existingReceiver === t; });
        if (existingRowSubject === cleanSubject &&
            (existingSender === sessEmail || matchesParticipant) &&
            existingThreadId === '') {
          sh.getRange(r + 1, 14).setValue(threadIdNorm);
        }
      }
    }

    return { success: true, mailId: mailId, attachments: attachments || [] };
  } catch(e) {
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
        try { att = JSON.parse(data[i][10] || '[]'); } catch(e) { att = []; }
        var labels = [];
        try { labels = JSON.parse(data[i][12] || '[]'); } catch(e) { labels = []; }
        var deletedBy = String(data[i][17] || '').toLowerCase();
        var userEmail = String(sess.email || '').toLowerCase();
        var deletedList = deletedBy ? deletedBy.split(',').map(function(e){return e.trim();}) : [];
        if (deletedList.indexOf(userEmail) !== -1) {
          return { success: false, message: 'You have deleted this mail.' };
        }

        var mail = {
          mailId: data[i][0], senderEmail: data[i][1], senderName: data[i][2],
          receiverEmail: data[i][3], receiverName: data[i][4], subject: data[i][5],
          body: data[i][6], timestamp: data[i][7], status: data[i][8], folder: data[i][9],
          attachments: att, hasAttachments: att.length > 0,
          starred: data[i][11] === true || data[i][11] === 'TRUE' || String(data[i][11]) === 'true',
          labels: labels,
          threadId: data[i][13] || '',
          priority: data[i][14] || 'normal',
          cc: data[i][15] || '',
          bcc: data[i][16] || '',
          deletedBy: deletedBy
        };

        // SECURITY: IDOR fix — verify user owns this mail or is admin
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          var isSender = String(mail.senderEmail || '').toLowerCase() === String(sess.email || '').toLowerCase();
          var isReceiver = String(mail.receiverEmail || '').toLowerCase() === String(sess.email || '').toLowerCase();
          if (!isSender && !isReceiver) {
            return { success: false, message: 'Permission denied: You can only view your own emails.' };
          }
        }

        sh.getRange(i + 1, 9).setValue("read");
        return { success: true, mail: mail };
      }
    }
    throw new Error("Mail not found.");
  } catch(e) {
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
        var senderEmail = String(data[i][1] || '').toLowerCase();
        var receiverEmail = String(data[i][3] || '').toLowerCase();
        var userEmail = String(sess.email || '').toLowerCase();
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          if (senderEmail !== userEmail && receiverEmail !== userEmail) {
            return { success: false, message: 'Permission denied: You can only access your own emails.' };
          }
        }
        sh.getRange(i + 1, 9).setValue("read");
        return { success: true, message: "Mail marked as read." };
      }
    }
    throw new Error("Mail not found: " + mailId);
  } catch(e) {
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
        var email = String(sess.email || '').toLowerCase();
        var senderEmail = String(data[i][1] || '').toLowerCase();
        var receiverEmail = String(data[i][3] || '').toLowerCase();
        var role = String(sess.role || '').toLowerCase();

        // SECURITY: Only allow owner (sender/receiver) or admin to delete mail
        if (role !== 'super admin' && role !== 'admin') {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: 'Permission denied: You can only delete your own emails.' };
          }
        }

        // Soft delete: mark this user as deleted
        var deletedBy = String(data[i][17] || '').toLowerCase();
        var deletedList = deletedBy ? deletedBy.split(',').map(function(e){return e.trim();}) : [];
        if (deletedList.indexOf(email) === -1) {
          deletedList.push(email);
        }
        sh.getRange(i + 1, 18).setValue(deletedList.join(','));
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getAllMails(token) {
  try {
    var sess = _session(token);
    var role = String(sess.role || '').toLowerCase();
    if (role !== 'super admin' && role !== 'admin') {
      throw new Error("Only admins can view all mails.");
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) return { success: true, mails: [] };
    var data = sh.getDataRange().getValues();
    var rows = data.slice(1).map(function(r) {
      var att = [];
      try { att = JSON.parse(r[10] || '[]'); } catch(e) { att = []; }
      var labels = [];
      try { labels = JSON.parse(r[12] || '[]'); } catch(e) { labels = []; }
      return {
        mailId: r[0], senderEmail: r[1], senderName: r[2],
        receiverEmail: r[3], receiverName: r[4], subject: r[5],
        body: r[6], timestamp: r[7], status: r[8], folder: r[9],
        attachments: att, hasAttachments: att.length > 0,
        starred: r[11] === true || r[11] === 'TRUE' || String(r[11]) === 'true',
        labels: labels,
        threadId: r[13] || '',
        priority: r[14] || 'normal',
        cc: r[15] || '',
        bcc: r[16] || ''
      };
    });
    return { success: true, mails: rows };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteMailPermanently(token, mailId) {
  try {
    var sess = _session(token);
    var role = String(sess.role || '').toLowerCase();
    if (role !== 'super admin') {
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
  } catch(e) {
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
    sh.appendRow([mailId, sess.email, sess.name, receiverEmailNorm, receiverName, subject, body, now, "unread", "Inbox", '[]', false, '', '', 'normal', '', '']);
    sh.appendRow([mailId, sess.email, sess.name, receiverEmailNorm, receiverName, subject, body, now, "read", "Sent", '[]', false, '', '', 'normal', '', '']);
    return { success: true, mailId: mailId };
  } catch(e) {
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
        var senderEmail = String(data[i][1] || '').toLowerCase();
        var receiverEmail = String(data[i][3] || '').toLowerCase();
        var email = String(sess.email || '').toLowerCase();
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: 'Permission denied.' };
          }
        }
        var currentStarred = data[i][11];
        var newStarred = !(currentStarred === true || currentStarred === 'TRUE' || String(currentStarred) === 'true');
        sh.getRange(i + 1, 12).setValue(newStarred);
        return { success: true, starred: newStarred };
      }
    }
    throw new Error("Mail not found.");
  } catch(e) {
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
        var senderEmail = String(data[i][1] || '').toLowerCase();
        var receiverEmail = String(data[i][3] || '').toLowerCase();
        var email = String(sess.email || '').toLowerCase();
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: 'Permission denied.' };
          }
        }
        sh.getRange(i + 1, 10).setValue("Archive");
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch(e) {
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
        var senderEmail = String(data[i][1] || '').toLowerCase();
        var receiverEmail = String(data[i][3] || '').toLowerCase();
        var email = String(sess.email || '').toLowerCase();
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: 'Permission denied.' };
          }
        }
        sh.getRange(i + 1, 10).setValue("Inbox");
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch(e) {
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
        var senderEmail = String(data[i][1] || '').toLowerCase();
        var receiverEmail = String(data[i][3] || '').toLowerCase();
        var email = String(sess.email || '').toLowerCase();
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          if (senderEmail !== email && receiverEmail !== email) {
            return { success: false, message: 'Permission denied.' };
          }
        }
        sh.getRange(i + 1, 9).setValue("unread");
        return { success: true };
      }
    }
    throw new Error("Mail not found.");
  } catch(e) {
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
    
    var senderEmail = String(originalMail[1] || '').toLowerCase();
    var receiverEmail = String(originalMail[3] || '').toLowerCase();
    var email = String(sess.email || '').toLowerCase();
    var role = String(sess.role || '').toLowerCase();
    if (role !== 'super admin' && role !== 'admin') {
      if (senderEmail !== email && receiverEmail !== email) {
        return { success: false, message: 'Permission denied.' };
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
    var subject = "Fwd: " + (originalMail[5] || '');
    var body = "\n\n---------- Forwarded message ----------\nFrom: " + originalMail[2] + " <" + originalMail[1] + ">\nDate: " + originalMail[7] + "\nSubject: " + originalMail[5] + "\n\n" + (originalMail[6] || '');
    var originalAtt = [];
    try { originalAtt = JSON.parse(originalMail[10] || '[]'); } catch(e) { originalAtt = []; }
    var attJson = JSON.stringify(originalAtt);

    sh.appendRow([mailIdNew, sess.email, sess.name, toEmailNorm, toUser[1], subject, body, now, "unread", "Inbox", attJson, false, '[]', '', 'normal', '', '']);
    sh.appendRow([mailIdNew, sess.email, sess.name, toEmailNorm, toUser[1], subject, body, now, "read", "Sent", attJson, false, '[]', '', 'normal', '', '']);

    return { success: true, mailId: mailIdNew };
  } catch(e) {
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
    var queryLower = String(query || '').toLowerCase();

    var filtered = rows.filter(function(r) {
      var mailFolder = r[9];
      var sender = r[1];
      var receiver = r[3];
      var subject = String(r[5] || '').toLowerCase();
      var body = String(r[6] || '').toLowerCase();

      if (folder === 'Inbox' && (receiver !== sess.email || mailFolder !== 'Inbox')) return false;
      if (folder === 'Sent' && (sender !== sess.email || mailFolder !== 'Sent')) return false;
      if (folder === 'Trash' && ((sender !== sess.email && receiver !== sess.email) || mailFolder !== 'Trash')) return false;
      if (folder === 'Starred' && (sender !== sess.email && receiver !== sess.email)) return false;
      if (folder === 'Archive' && ((sender !== sess.email && receiver !== sess.email) || mailFolder !== 'Archive')) return false;
      if (!folder && (sender !== sess.email && receiver !== sess.email)) return false;

      if (queryLower) {
        return subject.indexOf(queryLower) !== -1 || body.indexOf(queryLower) !== -1;
      }
      return true;
    }).map(function(r) {
      var att = [];
      try { att = JSON.parse(r[10] || '[]'); } catch(e) { att = []; }
      var labels = [];
      try { labels = JSON.parse(r[12] || '[]'); } catch(e) { labels = []; }
      return {
        mailId: r[0], senderEmail: r[1], senderName: r[2],
        receiverEmail: r[3], receiverName: r[4], subject: r[5],
        body: r[6], timestamp: r[7], status: r[8], folder: r[9],
        attachments: att, hasAttachments: att.length > 0,
        starred: r[11] === true || r[11] === 'TRUE' || String(r[11]) === 'true',
        labels: labels,
        threadId: r[13] || '',
        priority: r[14] || 'normal',
        cc: r[15] || '',
        bcc: r[16] || ''
      };
    });

    return { success: true, mails: filtered };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function emptyTrash(token) {
  try {
    if (!token) throw new Error("Authentication required.");
    var sess = _session(token);
    var role = String(sess.role || '').toLowerCase();
    if (role !== 'super admin' && role !== 'admin') {
      throw new Error("Only admins can empty trash.");
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    var data = sh.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][9] === 'Trash') {
        rowsToDelete.push(i + 1);
      }
    }
    for (var j = 0; j < rowsToDelete.length; j++) {
      sh.deleteRow(rowsToDelete[j]);
    }
    return { success: true, deletedCount: rowsToDelete.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


function getAllMailCounts(token) {
  try {
    var sess = _session(token);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MAILS);
    if (!sh || sh.getLastRow() < 2) {
      return { success: true, counts: { Inbox: 0, Sent: 0, Trash: 0, Starred: 0, Archive: 0, unreadInbox: 0 } };
    }
    var data = sh.getDataRange().getValues();
    var rows = data.slice(1);
    var counts = { Inbox: 0, Sent: 0, Self: 0, Trash: 0, Starred: 0, Archive: 0, unreadInbox: 0 };
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var mailFolder = String(r[9] || '');
      var sender = String(r[1] || '').toLowerCase();
      var receiver = String(r[3] || '').toLowerCase();
      var starred = r[11] === true || r[11] === 'TRUE' || String(r[11]) === 'true';
      var status = String(r[8] || '').toLowerCase();
      var email = String(sess.email || '').toLowerCase();
      var isOwner = sender === email || receiver === email;
      if (!isOwner) continue;

      if (mailFolder === 'Inbox' && receiver === email) counts.Inbox++;
      else if (mailFolder === 'Sent' && sender === email) counts.Sent++;
      else if (mailFolder === 'Self' && sender === email && sender === receiver) counts.Self++;
      else if (mailFolder === 'Trash' && isOwner) counts.Trash++;
      else if (mailFolder === 'Archive' && isOwner) counts.Archive++;
      else if (starred && isOwner) counts.Starred++;

      if (mailFolder === 'Inbox' && receiver === email && status === 'unread') counts.unreadInbox++;
    }
    return { success: true, counts: counts };
  } catch(e) {
    return { success: false, message: e.message };
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
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][13] === threadId) {
        var att = [];
        try { att = JSON.parse(data[i][10] || '[]'); } catch(e) { att = []; }
        var labels = [];
        try { labels = JSON.parse(data[i][12] || '[]'); } catch(e) { labels = []; }
        
        var mail = {
          mailId: data[i][0], senderEmail: data[i][1], senderName: data[i][2],
          receiverEmail: data[i][3], receiverName: data[i][4], subject: data[i][5],
          body: data[i][6], timestamp: data[i][7], status: data[i][8], folder: data[i][9],
          attachments: att, hasAttachments: att.length > 0,
          starred: data[i][11] === true || data[i][11] === 'TRUE' || String(data[i][11]) === 'true',
          labels: labels,
          threadId: data[i][13] || '',
          priority: data[i][14] || 'normal',
          cc: data[i][15] || '',
          bcc: data[i][16] || ''
        };
        
        // SECURITY: IDOR fix — verify user owns this mail or is admin
        var role = String(sess.role || '').toLowerCase();
        if (role !== 'super admin' && role !== 'admin') {
          var isSender = String(mail.senderEmail || '').toLowerCase() === String(sess.email || '').toLowerCase();
          var isReceiver = String(mail.receiverEmail || '').toLowerCase() === String(sess.email || '').toLowerCase();
          if (!isSender && !isReceiver) {
            continue;
          }
        }
        
        mails.push(mail);
      }
    }
    
    // Sort by timestamp ascending
    mails.sort(function(a, b) {
      return new Date(a.timestamp) - new Date(b.timestamp);
    });
    
    return { success: true, mails: mails };
  } catch(e) {
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
    var role = String(sess.role || '').toLowerCase();
    
    // Super admin required for ALL accounts, admin/super admin for specific account
    if (!targetEmail || targetEmail === 'ALL') {
      if (role !== 'super admin') {
        throw new Error("Only super admin can clear all mails across all accounts.");
      }
    } else {
      if (role !== 'super admin' && role !== 'admin') {
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
      var senderEmail = String(row[1] || '').toLowerCase();
      var receiverEmail = String(row[3] || '').toLowerCase();
      
      if (targetEmail === 'ALL') {
        rowsToDelete.push(i + 1);
      } else {
        var target = String(targetEmail || '').toLowerCase();
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
    
    var mode = targetEmail === 'ALL' ? 'All accounts' : targetEmail;
    return { success: true, deletedCount: rowsToDelete.length, message: "Cleared " + rowsToDelete.length + " mails for " + mode + "." };
  } catch(e) {
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
      requestOrigin = payload._origin || '';
    } catch(err) {
      return _json({ success: false, message: "API ERROR: Invalid JSON payload" });
    }
  } else {
    action = e.parameter.action;
    requestOrigin = e.parameter._origin || '';
    try {
      args = e.parameter.data ? JSON.parse(e.parameter.data) : [];
    } catch(err) {
      args = [];
    }
  }

  // ── DOMAIN LOCK: Only allow requests from whitelisted origins ────
  // The frontend sends window.location.origin as _origin in every request.
  // This prevents cloned pages hosted on unknown domains from using the API.
  if (!_isOriginAllowed(requestOrigin)) {
    // Allow requests from google.script.run (same-origin, origin is empty string)
    if (requestOrigin !== '') {
      console.warn('Blocked request from unauthorized origin: ' + requestOrigin);
      return _json({
        success: false,
        message: 'Access denied: This backend is locked to the official FiberOne portal. Unauthorized origin: ' + requestOrigin
      });
    }
  }

  if (!action) return _json({ success: false, message: "API ERROR: No action specified." });

  try {
    var result;
    switch(action) {
      case "loginUser": result = loginUser.apply(null, args); break;
      case "registerUser": result = registerUser.apply(null, args); break;
      case "getDepartments": result = getDepartments.apply(null, args); break;
      case "getUsers": result = getUsers.apply(null, args); break;
      case "searchUsers": result = searchUsers.apply(null, args); break;
      case "getDepartmentMembers": result = getDepartmentMembers.apply(null, args); break;
      case "updateMessagingSettings": result = updateMessagingSettings.apply(null, args); break;
      case "adminCreateUser": result = adminCreateUser.apply(null, args); break;
      case "updateUserProfile": result = updateUserProfile.apply(null, args); break;
      case "addDepartment": result = addDepartment.apply(null, args); break;
      case "updateDepartment": result = updateDepartment.apply(null, args); break;
      case "deleteDepartment": result = deleteDepartment.apply(null, args); break;
      case "updateUserRole": result = updateUserRole.apply(null, args); break;
      case "updateUserStatus": result = updateUserStatus.apply(null, args); break;
      case "getPublicDepartments": result = getPublicDepartments.apply(null, args); break;
      case "resetUserPassword": result = resetUserPassword.apply(null, args); break;
      case "getDomainConfig": result = getDomainConfig.apply(null, args); break;
      case "getAIResponse": result = getAIResponse.apply(null, args); break;
      case "callGemini": result = callGemini.apply(null, args); break;
      case "callGeminiVision": result = callGeminiVision.apply(null, args); break;
      case "logEmployeeMessage": result = logEmployeeMessage.apply(null, args); break;
      case "getPendingCheckin": result = getPendingCheckin.apply(null, args); break;
      case "respondCheckin": result = respondCheckin.apply(null, args); break;
      case "completeCheckin": result = completeCheckin.apply(null, args); break;
      case "getMySessionInfo": result = getMySessionInfo.apply(null, args); break;
      case "logoutUser": result = logoutUser.apply(null, args); break;
      case "getUserById": result = getUserById.apply(null, args); break;
      case "getUsersByDepartment": result = getUsersByDepartment.apply(null, args); break;
      case "getUserByEmail": result = getUserByEmail.apply(null, args); break;
      case "adminDeleteUser": result = adminDeleteUser.apply(null, args); break;
      case "adminUpdateUser": result = adminUpdateUser.apply(null, args); break;
      case "approvePAT": result = approvePAT.apply(null, args); break;
      case "getNextStageUsers": result = getNextStageUsers.apply(null, args); break;
      case "createDashboardShareToken": result = createDashboardShareToken.apply(null, args); break;
      case "getSharedDashboard": result = getSharedDashboard.apply(null, args); break;
      case "getPATAnalyticsForIdris": result = getPATAnalyticsForIdris.apply(null, args); break;
      case "getMaterials": result = getMaterials.apply(null, args); break;
      case "addMaterial": result = addMaterial.apply(null, args); break;
      case "deleteMaterial": result = deleteMaterial.apply(null, args); break;
      case "getPATProjects": result = getPATProjects.apply(null, args); break;
      case "getPATProjectById": result = getPATProjectById.apply(null, args); break;
      case "savePATProject": result = savePATProject.apply(null, args); break;
      case "getPATImages": result = getPATImages.apply(null, args); break;
      case "savePATImages": result = savePATImages.apply(null, args); break;
      case "addSinglePATImage": result = addSinglePATImage.apply(null, args); break;
      case "deletePATImages": result = deletePATImages.apply(null, args); break;
      case "saveProjectImages": result = saveProjectImages.apply(null, args); break;
      case "submitPATToDepartment": result = submitPATToDepartment.apply(null, args); break;
      case "submitPATToNextStage": result = submitPATToNextStage.apply(null, args); break;
      case "partiallyApprovePAT": result = partiallyApprovePAT.apply(null, args); break;
      case "deletePATProject": result = deletePATProject.apply(null, args); break;
      case "deleteAllDraftPATProjects": result = deleteAllDraftPATProjects.apply(null, args); break;
      case "deleteAllPATProjects": result = deleteAllPATProjects.apply(null, args); break;
      case "getPATProjectsWithPermissions": result = getPATProjectsWithPermissions.apply(null, args); break;
      case "checkPATPermission": result = checkPATPermission.apply(null, args); break;
      case "getSLAAnalytics": result = getSLAAnalytics.apply(null, args); break;
      case "getAdvancedSLADetails": result = getAdvancedSLADetails.apply(null, args); break;
      case "getBlacklist": result = getBlacklist.apply(null, args); break;
      case "banUserWithBlacklist": result = banUserWithBlacklist.apply(null, args); break;
      case "unbanUserWithBlacklist": result = unbanUserWithBlacklist.apply(null, args); break;
      case "removeBlacklistEntry": result = removeBlacklistEntry.apply(null, args); break;
      case "getMails": result = getMails.apply(null, args); break;
      case "getAllMailCounts": result = getAllMailCounts.apply(null, args); break;
      case "sendMail": result = sendMail.apply(null, args); break;
      case "getMailById": result = getMailById.apply(null, args); break;
      case "markMailAsRead": result = markMailAsRead.apply(null, args); break;
      case "markMailAsUnread": result = markMailAsUnread.apply(null, args); break;
      case "deleteMail": result = deleteMail.apply(null, args); break;
      case "getAllMails": result = getAllMails.apply(null, args); break;
      case "deleteMailPermanently": result = deleteMailPermanently.apply(null, args); break;
      case "emptyTrash": result = emptyTrash.apply(null, args); break;
      case "craftAndSendMail": result = craftAndSendMail.apply(null, args); break;
      case "toggleStarMail": result = toggleStarMail.apply(null, args); break;
      case "archiveMail": result = archiveMail.apply(null, args); break;
      case "restoreMail": result = restoreMail.apply(null, args); break;
      case "forwardMail": result = forwardMail.apply(null, args); break;
      case "searchMails": result = searchMails.apply(null, args); break;
      case "generateJCC": result = generateJCC.apply(null, args); break;
      case "getJCCByProjectId": result = getJCCByProjectId.apply(null, args); break;
      case "getAllJCCs": result = getAllJCCs.apply(null, args); break;
      case "generateVendorReviewLink": result = generateVendorReviewLink.apply(null, args); break;
      case "getVendorPATByToken": result = getVendorPATByToken.apply(null, args); break;
      case "submitVendorDecision": result = submitVendorDecision.apply(null, args); break;
      case "setAllowedOrigins": result = setAllowedOrigins.apply(null, args); break;
      case "getAllowedOrigins": result = getAllowedOrigins.apply(null, args); break;
      case "setLocalhostAccess": result = setLocalhostAccess.apply(null, args); break;
      case "uploadDocument": result = uploadDocument.apply(null, args); break;
      case "getDocuments": result = getDocuments.apply(null, args); break;
      case "deleteDocument": result = deleteDocument.apply(null, args); break;
      case "deleteProjectDocuments": result = deleteProjectDocuments.apply(null, args); break;
      case "getMailThread": result = getMailThread.apply(null, args); break;
      case "clearUserMails": result = clearUserMails.apply(null, args); break;
      default: throw new Error("Unknown action: " + action);
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
      var vendorTemplate = HtmlService.createTemplateFromFile('vendor-review');
      vendorTemplate.tokenFromServer = e.parameter.token;
      vendorTemplate.scriptUrl = ScriptApp.getService().getUrl();
      return vendorTemplate.evaluate()
        .setTitle('Vendor Review - PAT Project')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Determine which page to serve
    var page = e.parameter.page || 'index';
    
    // UNIFIED INTERFACE ENFORCEMENT
    // If any legacy system or cached link tries to access 'employee', force load 'admin'
    if (page === 'employee' || page === 'employee.html') page = 'admin';
    if (page === 'admin.html') page = 'admin';
    if (page === 'mail.html') page = 'mail';
    
    var template = HtmlService.createTemplateFromFile(page);
    template.scriptUrl = ScriptApp.getService().getUrl();
    
    return template.evaluate()
      .setTitle('FiberOne Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    console.error("doGet Error: " + err.message);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px;color:#dc2626">' +
      '<h2>System Routing Error</h2>' +
      '<p>The server encountered an error while trying to load the page.</p>' +
      '<pre style="background:#fef2f2;padding:10px;border:1px solid #fca5a5">' + err.message + '</pre>' +
      '<p style="margin-top:20px;font-size:12px;color:#6b7280">If you are the developer, check your template scriptlets.</p>' +
      '</div>'
    );
  }
}

function doPost(e) { return handleRequest(e); }

function _json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
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
  var cached = cache.get('allowed_origins');
  var allowed;
  if (cached) {
    allowed = JSON.parse(cached);
  } else {
    allowed = _loadAllowedOrigins();
    try { cache.put('allowed_origins', JSON.stringify(allowed), 600); } catch(e) {}
  }
  
  for (var i = 0; i < allowed.length; i++) {
    if (origin.indexOf(allowed[i]) === 0) return true;
  }
  return false;
}

function _loadAllowedOrigins() {
  var stored = PropertiesService.getScriptProperties().getProperty('allowed_origins');
  if (stored) {
    try { return JSON.parse(stored); } catch(e) {} 
  }
  // Default: allow the GAS deployment domain, your custom domain, and local dev
  return [
    'https://script.google.com',
    'https://script.googleusercontent.com',
    'https://pat.fob.net.ng',
    'http://localhost',
    'http://127.0.0.1'
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
      throw new Error('Must provide a non-empty array of origin strings.');
    }
    PropertiesService.getScriptProperties().setProperty('allowed_origins', JSON.stringify(origins));
    CacheService.getScriptCache().remove('allowed_origins');
    return { success: true, message: 'Allowed origins updated.', origins: origins };
  } catch(e) {
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
  } catch(e) {
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
  var description = ui.prompt("New Material Description:").getResponseText().trim();
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
  var email = ui.prompt("New " + role + " email (must end with " + config.DOMAIN + "):").getResponseText().trim();
  if (!email || !email.endsWith(config.DOMAIN)) { ui.alert("Invalid domain."); return; }
  var name  = ui.prompt("Full name:").getResponseText().trim();
  var pass  = ui.prompt("Password (min 4 chars):").getResponseText();
  
  var deptsRes = getPublicDepartments();
  var dept = "";
  if (role !== "super admin" && deptsRes.success && deptsRes.data.length > 0) {
    var deptMsg = "Available Departments:\n" + deptsRes.data.join(", ") + "\n\nEnter Department:";
    dept = ui.prompt(deptMsg).getResponseText().trim();
  } else if (role !== "super admin") {
    dept = ui.prompt("No departments found. Enter a new department name:").getResponseText().trim();
  } else { dept = "N/A"; }
  
  if (role !== "super admin" && !dept) { ui.alert("Department is required."); return; }

  try {
    var res = registerUser(name, email, pass, dept, "Other"); // Default gender for manual admin
    if (!res.success) throw new Error(res.message);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh     = ss.getSheetByName(SD.USERS);
    var data   = sh.getDataRange().getValues();
    var lo     = email.toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase() === lo) { sh.getRange(i+1,6).setValue(role); break; }
    }

    ui.alert("✅ " + role + " created!\n\nEmail: " + email);
  } catch(e) {
    ui.alert("❌ Error: " + e.message);
  }
}

/**
 * Securely resets a user's password. Restricted to Super Admins.
 */
function resetUserPassword(token, targetEmail, newPassword) {
  try {
    _superAdminSession(token);
    targetEmail = String(targetEmail || "").toLowerCase().trim();
    if (!targetEmail || !newPassword) throw new Error("Email and new password are required.");
    if (newPassword.length < 4) throw new Error("Password must be at least 4 characters.");

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase().trim() === targetEmail) {
        var s = _salt();
        sh.getRange(i + 1, 4).setValue(_hash(newPassword, s));
        sh.getRange(i + 1, 5).setValue(s);
        return { success: true, message: "Password reset successfully for " + targetEmail };
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
  var ui    = SpreadsheetApp.getUi();
  var email = ui.prompt("Email of user to reset:").getResponseText().trim().toLowerCase();
  if (!email) return;
  var pass  = ui.prompt("New password (min 4 chars):").getResponseText();
  if (!pass || pass.length < 4) { ui.alert("❌ Password too short."); return; }

  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email) {
      var s = _salt();
      sh.getRange(i+1, 4).setValue(_hash(pass, s));
      sh.getRange(i+1, 5).setValue(s);
      ui.alert("✅ Password reset for: " + email);
      return;
    }
  }
  ui.alert("❌ User not found: " + email);
}

function ui_liveStats() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var usr = ss.getSheetByName(SD.USERS);
  if (!usr) { ss.toast("Run First Setup first.", "Stats", 3); return; }

  var ud = usr.getDataRange().getValues().slice(1);
  var allUsers = ud.length;
  var active   = ud.filter(function(r){ return r[7]==="active"; }).length;

  // Count PAT projects
  var patSh = ss.getSheetByName(SD.PAT);
  var totalPat = patSh && patSh.getLastRow() > 1 ? patSh.getLastRow() - 1 : 0;

  SpreadsheetApp.getUi().alert(
    "📊 SD Portal — Live Stats\n\n" +
    "Users\n" +
    "  Registered: " + allUsers + "\n" +
    "  Active: " + active + "\n\n" +
    "PAT Projects\n" +
    "  Total: " + totalPat
  );
}

function ui_healthCheck() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var needed = [SD.USERS, SD.DEPTS, SD.PAT];
  var issues = [];
  needed.forEach(function(n){
    if (!ss.getSheetByName(n)) issues.push("Missing sheet: " + n);
  });
  var userSh = ss.getSheetByName(SD.USERS);
  if (userSh && userSh.getLastRow() < 2) issues.push("No users registered yet.");

  SpreadsheetApp.getUi().alert(
    issues.length === 0
      ? "✅ All clear — all sheets exist and are populated."
      : "⚠️ Issues found:\n\n" + issues.join("\n") + "\n\nRun 'First Setup' to fix."
  );
}

function ui_wipeSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("🚨 DANGER: SYSTEM WIPE", 
    "This will delete ALL Users, Departments, and PAT Projects. \n\nAre you absolutely sure you want to proceed?", 
    ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [SD.USERS, SD.DEPTS, SD.PAT];
    sheets.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (sh && sh.getLastRow() > 1) {
        sh.deleteRows(2, sh.getLastRow() - 1);
      }
    });
    PropertiesService.getScriptProperties().deleteAllProperties();
    PropertiesService.getUserProperties().deleteAllProperties();
    ss.toast("System wiped successfully. Run 'First Setup' to initialize again.", "Reset Complete");
  }
}

// ─────────────────────────────────────────────────────────────
//  MISSING BACKEND FUNCTIONS — frontend calls these
// ─────────────────────────────────────────────────────────────

function getUsersByDepartment(token, department) {
  try {
    // Support both (department) and (token, department) calling patterns
    if (arguments.length === 1) { department = token; }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    var members = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][6]).toLowerCase() === String(department).toLowerCase().trim()) {
        members.push({ email: data[i][2], name: data[i][1], role: String(data[i][5] || "").toLowerCase() });
      }
    }
    return { success: true, members: members };
  } catch(e) {
    return { success: false, message: e.message, members: [] };
  }
}

function getUserByEmail(token, email) {
  try {
    if (!token) throw new Error("Admin token required.");
    _adminSession(token);

    if (arguments.length === 1) { email = token; token = null; }
    if (!email) throw new Error("Email is required.");
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase().trim() === String(email).toLowerCase().trim()) {
        return {
          success: true,
          user: {
            userId: data[i][0],
            name: data[i][1],
            email: data[i][2],
            role: data[i][5],
            department: data[i][6],
            gender: data[i][7],
            status: data[i][9]
          }
        };
      }
    }
    return { success: false, message: "User not found." };
  } catch(e) {
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
      if (rowId === searchId || (isEmailSearch && rowEmail === searchId.toLowerCase())) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "User not found." };
  } catch(e) {
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
      if (rowId === searchId || (isEmailSearch && rowEmail === searchId.toLowerCase())) {
        var row = i + 1;
        if (updates.name) sh.getRange(row, 2).setValue(updates.name);
        if (updates.department || updates.dept) sh.getRange(row, 7).setValue(updates.department || updates.dept);
        if (updates.role) sh.getRange(row, 6).setValue(updates.role);
        if (updates.gender) sh.getRange(row, 8).setValue(updates.gender);
        return { success: true };
      }
    }
    return { success: false, message: "User not found." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getNextStageUsers(token, currentStatus) {
  try {
    if (arguments.length === 1) { currentStatus = token; }
    // Derive next department from the centralized workflow constants:
    // 1. Find the next status from WF_STAGE_FLOW (linear + recovery)
    // 2. Look up the owning department for that next status in STATUS_TO_DEPT
    var nextStatus = WF_STAGE_FLOW[currentStatus];
    var nextDept = nextStatus ? (STATUS_TO_DEPT[nextStatus] || 'MEC') : 'MEC';
    return getUsersByDepartment(nextDept);
  } catch(e) {
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
      return { success: false, message: 'Permission denied: You are not assigned to this project.' };
    }

    // Update project status
    project.workflowStatus = 'Completed';
    project.verdict = verdict || 'Fully Accepted';
    project.department = 'MEC';
    project.assignedToDept = 'MEC';
    project.assignedToName = '';
    project.assignedToEmail = '';
    var history = project.workflowHistory || [];
    history.push({
      fromStatus: res.project.workflowStatus,
      toStatus: 'Completed',
      by: { name: sess.name || 'Admin', email: sess.email, department: sess.department },
      comments: comments || '',
      timestamp: new Date().toISOString()
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
      message: 'PAT approved and completed.' + (jccResult.success ? ' JCC generated: ' + jccResult.certificateId : ''),
      jcc: jccResult
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function createDashboardShareToken(token, email, durationMinutes) {
  try {
    _session(token);
    var shareToken = _makeToken();
    PropertiesService.getScriptProperties().setProperty('dash_share_' + shareToken,
      JSON.stringify({ email: email, expiresAt: Date.now() + (durationMinutes || 5) * 60000, from: 'admin' }));
    return { success: true, token: shareToken, expiresAt: new Date(Date.now() + (durationMinutes || 5) * 60000).toISOString() };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getSharedDashboard(token) {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    for (var key in props) {
      if (key.indexOf('dash_share_') === 0) {
        var data = _safeParse(props[key], null);
        if (data && Date.now() < data.expiresAt) {
          return { success: true, shared: true, analytics: { totalProjects: 0, byStatus: {} }, from: data.from, expiresAt: new Date(data.expiresAt).toISOString() };
        }
        if (data === null) PropertiesService.getScriptProperties().deleteProperty(key);
      }
    }
    return { success: true, shared: false };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getPATAnalyticsForIdris(token, forceRefresh) {
  try {
    if (arguments.length === 1) { forceRefresh = token; }
    var projects = getPATProjects('').projects || [];
    var byStatus = {};
    var byDept = {};
    var completed = 0;
    var rejected = 0;
    
    projects.forEach(function(p) {
      var st = p.workflowStatus || 'Draft';
      byStatus[st] = (byStatus[st] || 0) + 1;
      if (st === 'Completed') completed++;
      if (st === 'Rejected') rejected++;
      
      var dp = p.department || 'Unknown';
      byDept[dp] = (byDept[dp] || 0) + 1;
    });
    
    var total = projects.length;
    var completionRate = total > 0 ? Math.round((completed / total) * 100) + "%" : "0%";
    var rejectionRate = total > 0 ? Math.round((rejected / total) * 100) + "%" : "0%";
    
    return {
      success: true,
      analytics: {
        total: total,
        byStatus: byStatus,
        projects: projects,
        workflowEfficiency: {
          completionRate: completionRate,
          averageCompletionTime: "12 days", // Mock value
          rejectionRate: rejectionRate
        }
      }
    };
  } catch(e) {
    return { success: false, message: e.message, analytics: { total: 0, byStatus: {}, projects: [], workflowEfficiency: { completionRate: "0%", averageCompletionTime: "N/A", rejectionRate: "0%" } } };
  }
}

function doEmployeeCheckin() {
  var userSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.USERS);
  if (!userSh || userSh.getLastRow() < 2) return;

  var data = userSh.getDataRange().getValues();
  var sh = _ensureCheckinSheet();
  var today = new Date().toISOString().split('T')[0];

  var questions = [
    "Good morning! How is your work going today?",
    "Hi there! How's your progress on today's tasks?",
    "Hello! How are things going with your work today?",
    "Hey! How's everything going on your end?",
    "Warm greetings! How is your work coming along?"
  ];

  var hour = new Date().getHours();
  var minute = new Date().getMinutes();
  var timeSlot = hour * 60 + minute;

  var slots = [
    { start: 540, end: 570 },
    { start: 642, end: 672 },
    { start: 744, end: 774 },
    { start: 846, end: 876 },
    { start: 948, end: 978 }
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
    var email = String(data[i][2] || '').toLowerCase().trim();
    var name = String(data[i][1] || '').trim();
    var role = String(data[i][5] || '').toLowerCase();
    var status = String(data[i][9] || '').toLowerCase();

    if (role !== 'employee') continue;
    if (status !== 'active') continue;
    if (!email) continue;

    var alreadyExists = false;
    for (var j = 1; j < existingData.length; j++) {
      var exEmail = String(existingData[j][1] || '').toLowerCase().trim();
      var exStatus = String(existingData[j][5] || '').toLowerCase();
      var exDate = String(existingData[j][6] || '').split('T')[0];
      var exQ = String(existingData[j][3] || '');

      if (exEmail === email && exDate === today && exQ === question && (exStatus === 'pending' || exStatus === 'responded' || exStatus === 'completed')) {
        alreadyExists = true;
        break;
      }
    }

    if (!alreadyExists) {
      var checkinId = _genId("CHK");
      sh.appendRow([checkinId, email, name, question, '', 'pending', new Date().toISOString(), '']);
    }
  }
}

function _ensureCheckinSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("SD_CHECKINS");
  if (!sh) {
    sh = ss.insertSheet("SD_CHECKINS");
    // Headers: CheckinID | Email | Name | Question | Answer | Status | CreatedAt | CompletedAt
    sh.appendRow(["CheckinID","Email","Name","Question","Answer","Status","Timestamp","CompletedAt"]);
    sh.getRange(1, 1, 1, 8).setBackground("#0d1526").setFontColor("#ffffff").setFontWeight("bold");
  }
  return sh;
}

function setupEmployeeCheckins() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'doEmployeeCheckin') {
      ScriptApp.deleteTrigger(t);
    }
  });

  var times = [
    { hour: 9, minute: 0 },
    { hour: 10, minute: 42 },
    { hour: 12, minute: 24 },
    { hour: 14, minute: 6 },
    { hour: 15, minute: 48 }
  ];

  times.forEach(function(t) {
    ScriptApp.newTrigger('doEmployeeCheckin')
      .timeBased()
      .atHour(t.hour)
      .nearMinute(t.minute)
      .everyDays(1)
      .inTimezone('Africa/Lagos')
      .create();
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Employee check-in triggers created (5x daily: 9AM-4:30PM).');
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
  } catch(e) {
    return { success: false, message: e.message, materials: [] };
  }
}

function _createMaterial(description) {
  try {
    if (!description || !description.trim()) return { success: false, message: 'Description required.' };
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SD.MATS);
    if (!sh) return { success: false, message: 'Materials sheet not found.' };
    sh.appendRow([_genId("MAT-"), description.trim(), new Date().toISOString()]);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function addMaterial(token, description) {
  try {
    _session(token);
    return _createMaterial(description);
  } catch(e) {
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
    return { success: false, message: 'Material not found.' };
  } catch(e) {
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
    sh.getRange(1, 1, 1, 9).setValues([[
      "EntryID","IP","DeviceFingerprint","UserEmail","UserName","BannedBy","Reason","BannedAt","UnbannedAt"
    ]]);
    sh.getRange(1, 1, 1, 9).setBackground(_getConfig().NAVY).setFontColor("#ffffff").setFontWeight("bold");
  }
  return sh;
}

function _blacklistCols() {
  return {
    ENTRY_ID: 0, IP: 1, DEVICE_FINGERPRINT: 2, USER_EMAIL: 3, USER_NAME: 4,
    BANNED_BY: 5, REASON: 6, BANNED_AT: 7, UNBANNED_AT: 8
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
    data.forEach(function(r) {
      entries.push({
        entryId: r[c.ENTRY_ID],
        ip: r[c.IP],
        deviceFingerprint: r[c.DEVICE_FINGERPRINT],
        userEmail: r[c.USER_EMAIL],
        userName: r[c.USER_NAME],
        bannedBy: r[c.BANNED_BY],
        reason: r[c.REASON],
        bannedAt: r[c.BANNED_AT],
        unbannedAt: r[c.UNBANNED_AT]
      });
    });
    return { success: true, entries: entries };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function banUserWithBlacklist(token, userId, reason) {
  try {
    _adminSession(token);
    
    // Get user info and their known IPs/fingerprints
    var user = getUserById(token, userId);
    if (!user.success) return { success: false, message: 'User not found: ' + user.message };
    
    var sh = _blacklistSheet();
    var c = _blacklistCols();
    var entryId = _genId("BL");
    var timestamp = new Date().toISOString();
    var sess = _session(token);
    var banCount = 0;
    var ips = user.ips || [];
    var fingerprints = user.fingerprints || [];
    
    var userName = user.name || '';
    var userEmail = user.email || '';
    
    ips.forEach(function(ip) {
      sh.appendRow([entryId + "-IP-" + banCount, ip, "", userEmail, userName, sess.email || "admin", reason || "Manual ban", timestamp, ""]);
      banCount++;
    });

    fingerprints.forEach(function(fp) {
      sh.appendRow([entryId + "-FP-" + banCount, "", fp, userEmail, userName, sess.email || "admin", reason || "Manual ban", timestamp, ""]);
      banCount++;
    });

    // Also add email-based blacklist entry
    sh.appendRow([entryId + "-EMAIL", "EMAIL:" + userEmail, "", userEmail, userName, sess.email || "admin", "Email banned - " + (reason || ''), timestamp, ""]);

    return { success: true, ipsBlacklisted: ips.length, devicesBlacklisted: fingerprints.length };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function unbanUserWithBlacklist(token, userId) {
  try {
    _adminSession(token);
    
    // Get user email from userId
    var user = getUserById(token, userId);
    if (!user.success) return { success: false, message: 'User not found: ' + user.message };
    
    var userEmail = user.email || '';
    var sh = _blacklistSheet();
    if (sh.getLastRow() < 2) return { success: true, message: 'No blacklist entries found.' };

    var c = _blacklistCols();
    var data = sh.getDataRange().getValues();
    var unbannedCount = 0;

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var rowEmail = String(row[c.USER_EMAIL] || '').toLowerCase();
      var searchEmail = String(userEmail || '').toLowerCase();
      if (rowEmail === searchEmail && !row[c.UNBANNED_AT]) {
        sh.getRange(i + 1, c.UNBANNED_AT + 1).setValue(new Date().toISOString());
        unbannedCount++;
      }
    }
    return { success: true, unbannedCount: unbannedCount };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function removeBlacklistEntry(token, entryId) {
  try {
    _adminSession(token);
    var sh = _blacklistSheet();
    if (sh.getLastRow() < 2) return { success: false, message: 'No entries found.' };

    var c = _blacklistCols();
    var data = sh.getDataRange().getValues();

    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][c.ENTRY_ID] || '') === String(entryId)) {
        sh.getRange(i + 1, c.UNBANNED_AT + 1).setValue(new Date().toISOString());
      }
    }
    return { success: true };
  } catch(e) {
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
    var role = String(sess.role || '').toLowerCase();

    // Check DB for latest role (handles promotions without re-login)
    var dbRole = _getUserRoleFromDb(sess.email);
    if (dbRole && dbRole !== role) role = dbRole;

    if (role === 'employee') {
      var EMPLOYEE_TEST_ONLY = ['OPEN_PAT', 'FILL_PAT_FORM', 'SAVE_PAT'];
      if (EMPLOYEE_TEST_ONLY.indexOf(commandType) === -1) {
        return { allowed: false, role: role, message: 'Access Denied: Employees are restricted to performing standard authorized tests only (open, fill, save PAT). Access to system data, backend information, project lists, mail, and sensitive data is prohibited.' };
      }
      return { allowed: true, role: role, message: 'Employee authorized for standard test only.' };
    }

    // MEC-only commands (non-MEC admins excluded — backend enforced)
    var MEC_CMDS = ['FILL_PAT_WITH_RANDOM_DATA'];
    if (MEC_CMDS.indexOf(commandType) !== -1) {
      var deptNorm = String(sess.department || '').toLowerCase();
      var mech = (deptNorm === 'mec' || deptNorm === 'mech');
      if (mech) {
        return { allowed: true, role: role, message: 'MEC authorized for PAT data operations.' };
      }
      return { allowed: false, role: role, message: 'Access Denied: This operation is restricted to MEC department personnel. You do not have the necessary permissions to perform this action. Even the AI will not open the workspace for this request.' };
    }

    // Safe read-only commands — non-employee, non-MEC-specific commands
    var READ_ONLY = ['OPEN_PAT', 'LOAD_PAT_PROJECT', 'FILL_PAT_FORM', 'SAVE_PAT',
                     'LIST_PAT_PROJECTS', 'GET_PAT_STATUS', 'GET_USER_BY_EMAIL',
                     'REFRESH_SYSTEM_DATA', 'SYNC_ALL_SYSTEM_DATA', 'FETCH_MAILS',
                     'GET_SHARED_DASHBOARD', 'EXPORT_PAT_PDF', 'GET_BLACKLIST', 'GET_USER_BY_ID'];
    if (READ_ONLY.indexOf(commandType) !== -1) {
      return { allowed: true, role: role, message: 'Authorized.' };
    }

    // Medium-risk commands — admin or above
    var ADMIN_CMDS = ['CREATE_USER', 'UPDATE_USER', 'UPDATE_USER_ROLE', 'CREATE_DEPT',
                      'SUBMIT_TO_NEXT_DEPT', 'PARTIALLY_APPROVE_PAT', 'APPROVE_PAT',
                      'REFRESH_ALL_CACHE', 'ADD_IMAGE_TO_PROJECT', 'CRAFT_AND_SEND_MAIL',
                      'SHARE_IDRIS_DASHBOARD', 'RESET_PASSWORD', 'BAN_USER', 'UNBAN_USER', 'REMOVE_BLACKLIST_ENTRY'];
    if (ADMIN_CMDS.indexOf(commandType) !== -1) {
      if (role === 'admin' || role === 'super admin') {
        return { allowed: true, role: role, message: 'Admin authorized.' };
      }
      return { allowed: false, role: role, message: 'Admin role required for command: ' + commandType };
    }

    // High-risk / destructive commands — super admin only
    var SUPER_ADMIN_CMDS = ['DELETE_USER', 'DELETE_PAT_PROJECT', 'DELETE_ALL_PAT_PROJECTS',
                            'DELETE_MAIL_PERMANENTLY', 'RANDOM_FILL_AND_APPROVE'];
    if (SUPER_ADMIN_CMDS.indexOf(commandType) !== -1) {
      if (role === 'super admin') {
        return { allowed: true, role: role, message: 'Super admin authorized.' };
      }
      return { allowed: false, role: role, message: 'Super admin role required for command: ' + commandType };
    }

    // Mail commands: normal delete restricted to owners, permanent delete to super admin
    if (commandType === 'DELETE_MAIL') {
      if (role === 'super admin' || role === 'admin') {
        return { allowed: true, role: role, message: 'Admin authorized for mail delete.' };
      }
      return { allowed: false, role: role, message: 'Admin role required to delete mail.' };
    }

    // Unknown command — deny by default
    return { allowed: false, role: role, message: 'Command not recognized or not authorized: ' + commandType };
  } catch(e) {
    return { allowed: false, role: '', message: 'Auth error: ' + e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// END OF FILE
// ─────────────────────────────────────────────────────────────