'use strict';
// Offline regression test for the PAT "clears for no reason" bug:
// 1. savePATProject (MEC/Admin) with empty boq/images must PRESERVE DB values.
// 2. savePATImages must RECONCILE: removed URL -> trashed, others kept.
const fs = require('fs');
const vm = require('vm');

function makeSheetWith(data) {
  const rows = data.map(r => r.slice());
  return {
    _rows: rows,
    getAll: () => rows,
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    getLastRow: () => rows.length,
    getLastColumn: () => (rows[0] ? rows[0].length : 0),
    getRange: (a, b, c, d) => {
      if (c !== undefined) return { getValues: () => rows.slice(1).map(r => r.slice()) };
      return {
        setValue: (v) => { if (rows[a - 1]) rows[a - 1][b - 1] = v; },
        getValue: () => (rows[a - 1] ? rows[a - 1][b - 1] : ''),
      };
    },
    deleteRow: (r) => { if (rows[r - 1]) rows.splice(r - 1, 1); },
    appendRow: (r) => rows.push(r.slice()),
  };
}

const PAT_HEADERS = ['ProjectID','ProjectName','SiteAddress','Lat','Lon','Phase','WorkDescription','Vendor','InspectionDate','Orchestrator','SnagScore','Status','Verdict','WorkflowStatus','Checklist','BOQ','Snags','Signoff','Images','HasVendor','VendorToken','VendorApprovalStatus','VendorApprovalComments','VendorApprovalDate','SubmittedBy','SubmittedByEmail','SubmittedByDept','SubmittedAt','AssignedToName','AssignedToEmail','AssignedToDept','RejectionReason','WorkflowHistory','Department','UpdatedAt','VendorEverApproved','PresidingOfficer','EditLog'];

let trashed = [];   // Drive file ids trashed
let docRows = [];   // SD_DOCUMENTS rows

let patSheetInst = null;
let docSheetInst = null;
function getPatSheet() {
  if (!patSheetInst) patSheetInst = makeSheetWith([PAT_HEADERS.slice(), patRow()]);
  return patSheetInst;
}
function getDocSheet() {
  if (!docSheetInst) docSheetInst = {
    getLastRow: () => docRows.length + 1,
    getDataRange: () => ({ getValues: () => [['DocID','ProjectID','Name','Mime','Size','DriveURL','DriveFileID','a','b','c','d','e']].concat(docRows) }),
    deleteRow: (r) => { if (docRows[r - 2]) docRows.splice(r - 2, 1); },
    appendRow: (r) => docRows.push(r.slice()),
  };
  return docSheetInst;
}

function patRow(over) {
  const base = new Array(PAT_HEADERS.length).fill('');
  base[0] = 'PAT-5D8DD0';
  base[1] = 'Pole Tracker';
  base[2] = 'Lagos';
  base[13] = 'Draft';
  base[15] = JSON.stringify([{desc:'Fiber Cables',scope:'5',installed:'5',variance:'0',remark:'—'},{desc:'PCC',scope:'3',installed:'2',variance:'-1',remark:'insufficient material 1 was added'}]);
  base[18] = JSON.stringify(['https://lh3.googleusercontent.com/d/AAA','https://lh3.googleusercontent.com/d/BBB']);
  return Object.assign(base, over || {});
}

const GAS = {
  console, Date, Math, JSON,
  Utilities: { getUuid: () => 'UUID-' + Math.random().toString(36).slice(2, 10) },
  ScriptApp: { getService: () => ({ getUrl: () => 'U' }) },
  DriveApp: {
    Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
    Permission: { VIEW: 'VIEW' },
    getFoldersByName: () => ({ hasNext: () => false, next: () => ({}) }),
    createFolder: () => ({ setSharing: () => {} }),
    getFileById: (id) => ({ setTrashed: () => { trashed.push(id); } }),
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: (name) => {
        if (name === 'SD_PAT_PROJECTS') return getPatSheet();
        if (name === 'SD_DOCUMENTS') return getDocSheet();
        return null;
      },
      toast: () => {},
    }),
  },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} }), getUserProperties: () => ({ getProperty: () => null, setProperty: () => {}, getProperties: () => ({}) }), getDocumentProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
  Session: { getActiveUser: () => ({ getEmail: () => 'x@x.com' }) },
  LockService: { getUserLock: () => ({ waitLock: () => true, releaseLock: () => {} }), getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
};

const src = fs.readFileSync('Code.js', 'utf8');
const ctx = Object.assign({ module: {}, exports: {} }, GAS);
vm.createContext(ctx);
vm.runInContext(src + '\n;this.__x = { savePATProject, savePATImages, deletePATImages, _patSheet, _patCols, _projectFromRow };', ctx);
const api = ctx.__x;

// Fake MEC admin session
const TOKEN = 'FAKE-TOKEN';
const SESS = { email: 'mec@fob.ng', name: 'MEC Admin', role: 'admin', department: 'MEC' };
// Monkeypatch _session by injecting via the save functions' own _session call.
// Easiest: stub PropertiesService.getUserProperties to return our SESS for any tok_ key.
GAS.PropertiesService.getUserProperties = () => ({
  getProperty: (k) => { if (k === 'tok_' + TOKEN) return JSON.stringify(SESS); return null; },
  setProperty: () => {}, getProperties: () => ({}),
});

let pass = 0, fail = 0;
function check(name, cond) { if (cond) { pass++; console.log('  PASS  ' + name); } else { fail++; console.log('  FAIL  ' + name); } }

console.log('\n[TEST 1] savePATProject with empty boq/images must NOT wipe DB');
api.savePATProject(TOKEN, {
  projectId: 'PAT-5D8DD0', projectName: 'Pole Tracker', siteAddress: 'Lagos',
  // deliberately omit boq & images
  boq: [], images: [], checklist: {}, snags: [], signoff: {},
  workflowStatus: 'Awaiting Project Team', verdict: 'Pending',
});
const sh = api._patSheet();
const after = api._projectFromRow(sh.getAll()[1]);
check('BOQ preserved (2 rows)', Array.isArray(after.boq) && after.boq.length === 2);
check('IMAGES preserved (2 urls)', Array.isArray(after.images) && after.images.length === 2);

console.log('\n[TEST 2] savePATImages reconcile: drop one URL, keep other');
trashed = [];
const r2 = api.savePATImages(TOKEN, 'PAT-5D8DD0', ['https://lh3.googleusercontent.com/d/BBB']);
console.log('  savePATImages ->', JSON.stringify(r2).slice(0, 160));
const sh2 = api._patSheet();
const after2 = api._projectFromRow(sh2.getAll()[1]);
check('Remaining image = BBB only', JSON.stringify(after2.images) === JSON.stringify(['https://lh3.googleusercontent.com/d/BBB']));
check('AAA Drive file trashed', trashed.indexOf('AAA') !== -1);

console.log('\n[TEST 3] savePATImages with empty incoming + existing must NOT wipe');
trashed = [];
const r3 = api.savePATImages(TOKEN, 'PAT-5D8DD0', []);
const sh3 = api._patSheet();
const after3 = api._projectFromRow(sh3.getAll()[1]);
check('Images kept (guard against spurious clear)', after3.images.length === 1 && trashed.length === 0);

console.log('\n[TEST 4] deletePATImages explicitly wipes + trashes all');
trashed = [];
api.deletePATImages(TOKEN, 'PAT-5D8DD0');
const sh4 = api._patSheet();
const after4 = api._projectFromRow(sh4.getAll()[1]);
check('Images cleared', after4.images.length === 0);
check('BBB Drive file trashed', trashed.indexOf('BBB') !== -1);

console.log('\n==== RESULT: ' + pass + ' passed, ' + fail + ' failed ====');
process.exit(fail ? 1 : 0);
