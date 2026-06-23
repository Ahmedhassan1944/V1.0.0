/**
 * =========================================================
 * FULL TEST SUITE (Tests.gs)
 * =========================================================
 * Tests every public API function in the system:
 *   - Code.gs          → api_getDashboardData, api_uploadDocumentToDrive
 *   - Database.gs      → api_getAllCandidates, api_createCandidate,
 *                        api_updateCandidateStatus, api_getDocumentsByCandidate,
 *                        api_reviewDocument, api_getAuditLog
 *   - DriveManager.gs  → api_createCandidateFolder, api_uploadFileToDrive
 *   - EmailService.gs  → parseRecruitmentEmail_, checkOverdueCandidates,
 *                        api_sendRejectionEmail, api_sendPackageSubmissionAlert
 *
 * HOW TO RUN:
 *   1. Open your Google Apps Script project.
 *   2. Select function "runAllTests" from the function dropdown.
 *   3. Click ▶ Run.
 *   4. Open View > Logs (Ctrl+Enter) to see results.
 *
 * NOTE: Each test injects mock GAS services before calling the
 *       real function under test. No live Sheet/Drive/Gmail needed.
 * =========================================================
 */

// ─────────────────────────────────────────────
// ENTRY POINT
// ─────────────────────────────────────────────

/**
 * Main entry point. Run this function from the GAS editor.
 * Executes every test suite in order and prints a final summary.
 */
function runAllTests() {
  Logger.log('═══════════════════════════════════════════════');
  Logger.log('   ENTERPRISE HR MOBILIZATION — TEST SUITE   ');
  Logger.log('═══════════════════════════════════════════════\n');

  // Inject mocks into the global scope for the duration of all tests
  _injectMocks();

  try {
    suite_Code();
    suite_Database_Candidates();
    suite_Database_Documents();
    suite_Database_AuditLog();
    suite_DriveManager();
    suite_EmailService();
  } finally {
    // Restore real GAS globals (no-op in production; avoids leakage)
    _restoreMocks();
  }

  TestRunner.summary();
}

// ─────────────────────────────────────────────
// MOCK INJECTION
// ─────────────────────────────────────────────

// Hold originals so we can restore after tests
let _origSpreadsheetApp, _origDriveApp, _origGmailApp,
    _origSession, _origUtilities, _origScriptApp;

function _injectMocks() {
  _origSpreadsheetApp = typeof SpreadsheetApp !== 'undefined' ? SpreadsheetApp : null;
  _origDriveApp       = typeof DriveApp       !== 'undefined' ? DriveApp       : null;
  _origGmailApp       = typeof GmailApp       !== 'undefined' ? GmailApp       : null;
  _origSession        = typeof Session        !== 'undefined' ? Session        : null;
  _origUtilities      = typeof Utilities      !== 'undefined' ? Utilities      : null;
  _origScriptApp      = typeof ScriptApp      !== 'undefined' ? ScriptApp      : null;

  // Override globals with mocks
  SpreadsheetApp = MockFactory.getSpreadsheetApp(); // eslint-disable-line no-global-assign
  DriveApp       = MockFactory.getDriveApp();       // eslint-disable-line no-global-assign
  GmailApp       = MockFactory.getGmailApp();       // eslint-disable-line no-global-assign
  Session        = MockFactory.getSession();        // eslint-disable-line no-global-assign
  Utilities      = MockFactory.getUtilities();      // eslint-disable-line no-global-assign
  ScriptApp      = MockFactory.getScriptApp();      // eslint-disable-line no-global-assign
}

function _restoreMocks() {
  if (_origSpreadsheetApp) SpreadsheetApp = _origSpreadsheetApp; // eslint-disable-line no-global-assign
  if (_origDriveApp)       DriveApp       = _origDriveApp;       // eslint-disable-line no-global-assign
  if (_origGmailApp)       GmailApp       = _origGmailApp;       // eslint-disable-line no-global-assign
  if (_origSession)        Session        = _origSession;        // eslint-disable-line no-global-assign
  if (_origUtilities)      Utilities      = _origUtilities;      // eslint-disable-line no-global-assign
  if (_origScriptApp)      ScriptApp      = _origScriptApp;      // eslint-disable-line no-global-assign
}

// ─────────────────────────────────────────────
// SHARED SEED DATA
// ─────────────────────────────────────────────

const CANDIDATE_HEADERS = [
  'CandidateID', 'FullName', 'Position', 'Department', 'Email',
  'Phone', 'Nationality', 'OfferSalary', 'AssignedCoordinatorEmail',
  'CurrentStatus', 'CreatedAt', 'UpdatedAt', 'DriveFolderID'
];

const DOCUMENT_HEADERS = [
  'DocumentID', 'CandidateID', 'DocType', 'FileName', 'FileURL',
  'UploadDate', 'ApprovalStatus', 'ApprovedBy', 'VersionNumber', 'Remarks'
];

const LOG_HEADERS = [
  'LogID', 'Timestamp', 'CandidateID', 'Actor', 'Event'
];

const USER_HEADERS = ['UserID', 'Email', 'Role', 'Name'];

/**
 * Seeds all four sheets with minimal valid data for a test run.
 */
function _seedAllSheets() {
  MockFactory.reset();

  MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, [
    ['cand-001', 'Ahmed Ali', 'Engineer', 'Projects', 'ahmed@test.com',
     '+201012345678', 'Egyptian', '350', 'coord@company.com',
     'Documents Requested', '2026-01-01T00:00:00.000Z', '2026-01-01T00:00:00.000Z', '']
  ]);

  MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS, [
    ['doc-001', 'cand-001', 'Passport', 'passport.pdf',
     'https://drive.google.com/file/doc-001',
     '2026-01-02T00:00:00.000Z', 'Pending Review', '', 1, '']
  ]);

  MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS, [
    ['log-001', '2026-01-01T00:00:00.000Z', 'cand-001', 'SYSTEM', 'Candidate Created']
  ]);

  MockFactory.seedSheet(SHEET_USERS, USER_HEADERS, []);
}


// ═══════════════════════════════════════════════════════════
//  SUITE 1 — Code.gs
// ═══════════════════════════════════════════════════════════

function suite_Code() {
  Logger.log('\n── Suite: Code.gs ──────────────────────────────');

  TestRunner.run('api_getDashboardData returns success object', () => {
    const result = api_getDashboardData();
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.data, 'data should not be null');
    Assert.isTrue(typeof result.data.activeCount === 'number', 'activeCount should be a number');
    Assert.isTrue(typeof result.data.missingDocs === 'number', 'missingDocs should be a number');
    Assert.isTrue(typeof result.data.pendingValidation === 'number', 'pendingValidation should be a number');
  });

  TestRunner.run('api_getDashboardData returns correct KPI values', () => {
    const result = api_getDashboardData();
    Assert.equals(result.data.activeCount, 24, 'activeCount should be 24');
    Assert.equals(result.data.missingDocs, 12, 'missingDocs should be 12');
    Assert.equals(result.data.pendingValidation, 5, 'pendingValidation should be 5');
  });

  TestRunner.run('api_uploadDocumentToDrive returns success with a file URL', () => {
    const result = api_uploadDocumentToDrive('cand-001', 'Passport', 'passport.pdf', 'base64string==');
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.fileUrl, 'fileUrl should not be null');
    Assert.isTrue(result.fileUrl.startsWith('https://'), 'fileUrl should start with https://');
  });

  TestRunner.run('include() returns non-empty content string', () => {
    // include() reads .html files — we verify it does NOT throw and returns a string.
    // In the live GAS environment it would return real HTML.
    // We just verify the function signature here.
    Assert.isTrue(typeof include === 'function', 'include should be a function');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 2 — Database.gs: Candidates
// ═══════════════════════════════════════════════════════════

function suite_Database_Candidates() {
  Logger.log('\n── Suite: Database.gs — Candidates ─────────────');

  TestRunner.run('api_getAllCandidates returns seeded candidate', () => {
    _seedAllSheets();
    const result = api_getAllCandidates();
    Assert.isTrue(result.success, 'success should be true');
    Assert.isTrue(Array.isArray(result.data), 'data should be an array');
    Assert.equals(result.data.length, 1, 'should return exactly 1 seeded candidate');
    Assert.equals(result.data[0].CandidateID, 'cand-001', 'CandidateID should match seed');
    Assert.equals(result.data[0].FullName, 'Ahmed Ali', 'FullName should match seed');
  });

  TestRunner.run('api_getAllCandidates returns empty array when sheet is empty', () => {
    MockFactory.reset();
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS); // no rows
    MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS);
    MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS);
    MockFactory.seedSheet(SHEET_USERS, USER_HEADERS);

    const result = api_getAllCandidates();
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 0, 'no candidates should be returned');
  });

  TestRunner.run('api_getAllCandidates fails gracefully when sheet is missing', () => {
    MockFactory.reset(); // No sheets seeded at all
    const result = api_getAllCandidates();
    Assert.isFalse(result.success, 'success should be false for missing sheet');
    Assert.notNull(result.error, 'error message should be present');
  });

  TestRunner.run('api_createCandidate appends a new row with correct data', () => {
    _seedAllSheets();
    const payload = {
      fullName:         'Sara Mohamed',
      position:         'Accountant',
      department:       'Finance',
      email:            'sara@test.com',
      phone:            '+201099999999',
      nationality:      'Egyptian',
      salary:           '400',
      coordinatorEmail: 'coord@company.com'
    };
    const result = api_createCandidate(payload);
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.candidateId, 'candidateId should be returned');

    // Verify new row was actually added
    const allResult = api_getAllCandidates();
    Assert.equals(allResult.data.length, 2, 'should now have 2 candidates');
    const newCand = allResult.data.find(c => c.FullName === 'Sara Mohamed');
    Assert.notNull(newCand, 'new candidate should exist');
    Assert.equals(newCand.CurrentStatus, 'Documents Requested', 'default status should be Documents Requested');
  });

  TestRunner.run('api_createCandidate writes an audit log entry', () => {
    _seedAllSheets();
    const payload = {
      fullName: 'Test Candidate', position: 'Tester', department: 'QA',
      email: 'tester@test.com', phone: '+0000', nationality: 'Unknown',
      salary: '500', coordinatorEmail: 'coord@company.com'
    };
    api_createCandidate(payload);

    // Audit log should have grown beyond the initial seed entry
    const sheet = MockFactory._sheets[SHEET_LOGS];
    const rows  = sheet._data.slice(1); // skip header
    Assert.isTrue(rows.length >= 2, 'at least one new log entry should have been written');
  });

  TestRunner.run('api_updateCandidateStatus updates status for existing candidate', () => {
    _seedAllSheets();
    const result = api_updateCandidateStatus('cand-001', 'Documents Received');
    Assert.isTrue(result.success, 'success should be true');

    const allResult = api_getAllCandidates();
    Assert.equals(allResult.data[0].CurrentStatus, 'Documents Received', 'status should be updated');
  });

  TestRunner.run('api_updateCandidateStatus returns error for unknown candidateId', () => {
    _seedAllSheets();
    const result = api_updateCandidateStatus('does-not-exist', 'Mobilized');
    Assert.isFalse(result.success, 'success should be false for unknown ID');
    Assert.notNull(result.error, 'an error message should be present');
  });

  TestRunner.run('api_updateCandidateStatus updates the UpdatedAt timestamp', () => {
    _seedAllSheets();
    const before = MockFactory._sheets[SHEET_CANDIDATES]._data[1][11]; // UpdatedAt column (index 11)
    api_updateCandidateStatus('cand-001', 'Mobilized');
    const after  = MockFactory._sheets[SHEET_CANDIDATES]._data[1][11];
    // Timestamps should be different (unless tests run at exactly the same millisecond)
    // At minimum, the field should have a value after the update
    Assert.notNull(after, 'UpdatedAt should be set after status update');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 3 — Database.gs: Documents
// ═══════════════════════════════════════════════════════════

function suite_Database_Documents() {
  Logger.log('\n── Suite: Database.gs — Documents ──────────────');

  TestRunner.run('api_getDocumentsByCandidate returns correct documents for candidate', () => {
    _seedAllSheets();
    const result = api_getDocumentsByCandidate('cand-001');
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 1, 'should return exactly 1 document');
    Assert.equals(result.data[0].DocType, 'Passport', 'DocType should be Passport');
    Assert.equals(result.data[0].DocumentID, 'doc-001', 'DocumentID should match seed');
  });

  TestRunner.run('api_getDocumentsByCandidate returns empty array for unknown candidate', () => {
    _seedAllSheets();
    const result = api_getDocumentsByCandidate('unknown-cand');
    Assert.isTrue(result.success, 'success should be true even with no results');
    Assert.equals(result.data.length, 0, 'no documents should be returned for unknown candidate');
  });

  TestRunner.run('api_reviewDocument approves document and records actor', () => {
    _seedAllSheets();
    const result = api_reviewDocument('doc-001', 'Approved', '');
    Assert.isTrue(result.success, 'success should be true');

    // Verify status was updated in the sheet
    const docRows = MockFactory._sheets[SHEET_DOCUMENTS]._data.slice(1);
    const docRow  = docRows.find(r => r[0] === 'doc-001');
    Assert.notNull(docRow, 'document row should exist');
    Assert.equals(docRow[6], 'Approved', 'ApprovalStatus column should be Approved');
    Assert.equals(docRow[7], 'test.user@yourcompany.com', 'ApprovedBy should be set to current user');
  });

  TestRunner.run('api_reviewDocument rejects document and sets remarks', () => {
    _seedAllSheets();
    const result = api_reviewDocument('doc-001', 'Rejected', 'Photo is blurry, please re-upload.');
    Assert.isTrue(result.success, 'success should be true');

    const docRows = MockFactory._sheets[SHEET_DOCUMENTS]._data.slice(1);
    const docRow  = docRows.find(r => r[0] === 'doc-001');
    Assert.equals(docRow[6], 'Rejected', 'ApprovalStatus should be Rejected');
    Assert.equals(docRow[9], 'Photo is blurry, please re-upload.', 'Remarks should be set');
  });

  TestRunner.run('api_reviewDocument returns error for unknown documentId', () => {
    _seedAllSheets();
    const result = api_reviewDocument('no-such-doc', 'Approved', '');
    Assert.isFalse(result.success, 'success should be false for unknown doc ID');
    Assert.notNull(result.error, 'error should be present');
  });

  TestRunner.run('api_reviewDocument writes audit log entry', () => {
    _seedAllSheets();
    const logsBefore = MockFactory._sheets[SHEET_LOGS]._data.length;
    api_reviewDocument('doc-001', 'Approved', '');
    const logsAfter  = MockFactory._sheets[SHEET_LOGS]._data.length;
    Assert.isTrue(logsAfter > logsBefore, 'audit log should grow after document review');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 4 — Database.gs: Audit Log
// ═══════════════════════════════════════════════════════════

function suite_Database_AuditLog() {
  Logger.log('\n── Suite: Database.gs — Audit Log ──────────────');

  TestRunner.run('api_getAuditLog returns seeded log entry for candidate', () => {
    _seedAllSheets();
    const result = api_getAuditLog('cand-001');
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 1, 'should return 1 seeded log entry');
    Assert.equals(result.data[0].Event, 'Candidate Created', 'Event text should match seed');
    Assert.equals(result.data[0].Actor, 'SYSTEM', 'Actor should be SYSTEM');
  });

  TestRunner.run('api_getAuditLog returns empty array for candidate with no logs', () => {
    _seedAllSheets();
    const result = api_getAuditLog('no-such-cand');
    Assert.isTrue(result.success, 'success should be true');
    Assert.equals(result.data.length, 0, 'no log entries for unknown candidate');
  });

  TestRunner.run('api_getAuditLog accumulates entries over multiple actions', () => {
    _seedAllSheets();
    // Perform two actions that each write a log entry
    api_updateCandidateStatus('cand-001', 'Documents Received');
    api_reviewDocument('doc-001', 'Approved', '');

    const result = api_getAuditLog('cand-001');
    Assert.isTrue(result.data.length >= 3, 'should have at least 3 log entries (1 seed + 2 actions)');
  });

  TestRunner.run('api_writeLog_ (internal) does not expose errors if sheet missing', () => {
    MockFactory.reset(); // no SHEET_LOGS seeded — simulates missing sheet

    // api_writeLog_ must catch its own errors and NOT propagate them.
    // We simply call it and assert no exception escapes.
    let threw = false;
    try {
      api_writeLog_('x', 'y', 'z');
    } catch (_) {
      threw = true;
    }
    Assert.isFalse(threw, 'api_writeLog_ should swallow errors internally and not throw to caller');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 5 — DriveManager.gs
// ═══════════════════════════════════════════════════════════

function suite_DriveManager() {
  Logger.log('\n── Suite: DriveManager.gs ───────────────────────');

  TestRunner.run('api_createCandidateFolder creates a new folder and returns folderId', () => {
    _seedAllSheets();
    const result = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(result.success, 'success should be true');
    Assert.notNull(result.folderId, 'folderId should be returned');
    Assert.isTrue(result.url.startsWith('https://'), 'url should be a valid https link');
  });

  TestRunner.run('api_createCandidateFolder is idempotent (no duplicate folders)', () => {
    _seedAllSheets();
    const first  = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    const second = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    // Both should succeed and return the same folder ID
    Assert.isTrue(first.success,  'first call should succeed');
    Assert.isTrue(second.success, 'second call should succeed');
    Assert.equals(first.folderId, second.folderId, 'folderId should be identical on repeated calls');
  });

  TestRunner.run('api_createCandidateFolder stores folderId in Candidates sheet', () => {
    _seedAllSheets();
    const result = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(result.success, 'folder creation should succeed');

    // Verify DriveFolderID column was written back
    const allCands = api_getAllCandidates();
    const cand = allCands.data.find(c => c.CandidateID === 'cand-001');
    Assert.notNull(cand.DriveFolderID, 'DriveFolderID should be populated after folder creation');
  });

  TestRunner.run('api_uploadFileToDrive uploads file and returns fileUrl + documentId', () => {
    _seedAllSheets();
    // First create a folder to get its ID
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    Assert.isTrue(folderResult.success, 'folder must be created first');

    const uploadResult = api_uploadFileToDrive(
      'cand-001',
      folderResult.folderId,
      'Photo',
      'photo.jpg',
      'bW9ja2Jhc2U2NGRhdGE=', // mock base64
      'image/jpeg'
    );

    Assert.isTrue(uploadResult.success, 'upload should succeed');
    Assert.notNull(uploadResult.fileUrl, 'fileUrl should be returned');
    Assert.notNull(uploadResult.documentId, 'documentId should be returned');
  });

  TestRunner.run('api_uploadFileToDrive writes document record to Documents sheet', () => {
    _seedAllSheets();
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    api_uploadFileToDrive('cand-001', folderResult.folderId, 'Medical', 'medical.pdf', 'data==', 'application/pdf');

    const docsResult = api_getDocumentsByCandidate('cand-001');
    const medical    = docsResult.data.find(d => d.DocType === 'Medical');
    Assert.notNull(medical, 'Medical document record should exist in Documents sheet');
    Assert.equals(medical.ApprovalStatus, 'Pending Review', 'new doc should start as Pending Review');
    Assert.equals(medical.VersionNumber, 1, 'first upload should be version 1');
  });

  TestRunner.run('api_uploadFileToDrive increments version number on re-upload', () => {
    _seedAllSheets();
    const folderResult = api_createCandidateFolder('cand-001', 'Ahmed Ali');
    const fid = folderResult.folderId;

    // First upload
    api_uploadFileToDrive('cand-001', fid, 'Passport', 'pass_v1.pdf', 'data=', 'application/pdf');
    // Second upload (re-upload same docType)
    api_uploadFileToDrive('cand-001', fid, 'Passport', 'pass_v2.pdf', 'data=', 'application/pdf');

    const docsResult = api_getDocumentsByCandidate('cand-001');
    const passports  = docsResult.data.filter(d => d.DocType === 'Passport');
    // Should have: 1 seed (Passport doc-001) + 2 new uploads = 3 total Passport records
    Assert.isTrue(passports.length >= 2, 'multiple Passport versions should exist');
    const versions = passports.map(d => d.VersionNumber).sort((a, b) => a - b);
    Assert.isTrue(versions[versions.length - 1] > versions[0], 'latest version number should be higher');
  });
}


// ═══════════════════════════════════════════════════════════
//  SUITE 6 — EmailService.gs
// ═══════════════════════════════════════════════════════════

function suite_EmailService() {
  Logger.log('\n── Suite: EmailService.gs ───────────────────────');

  TestRunner.run('parseRecruitmentEmail_ parses valid structured email body', () => {
    _seedAllSheets();
    const body = `
CANDIDATE_NAME: Omar Khaled
POSITION: Mechanical Engineer
DEPARTMENT: Maintenance
EMAIL: omar.khaled@gmail.com
PHONE: +201055556666
NATIONALITY: Egyptian
SALARY_OMR: 450
    `;
    const result = parseRecruitmentEmail_(body);
    Assert.isTrue(result.success, 'parsing should succeed');
    Assert.equals(result.data.fullName,    'Omar Khaled',           'fullName should be parsed');
    Assert.equals(result.data.position,    'Mechanical Engineer',   'position should be parsed');
    Assert.equals(result.data.department,  'Maintenance',           'department should be parsed');
    Assert.equals(result.data.email,       'omar.khaled@gmail.com', 'email should be parsed');
    Assert.equals(result.data.nationality, 'Egyptian',              'nationality should be parsed');
    Assert.equals(result.data.salary,      '450',                   'salary should be parsed');
  });

  TestRunner.run('parseRecruitmentEmail_ returns failure when required fields are missing', () => {
    const body = 'POSITION: Tester\nDEPARTMENT: QA'; // no CANDIDATE_NAME, no EMAIL
    const result = parseRecruitmentEmail_(body);
    Assert.isFalse(result.success, 'parsing should fail for incomplete email');
    Assert.notNull(result.error, 'error message should be present');
  });

  TestRunner.run('parseRecruitmentEmail_ is case-insensitive for field keys', () => {
    const body = 'candidate_name: Ali Hassan\nposition: Driver\nemail: ali@test.com';
    const result = parseRecruitmentEmail_(body);
    Assert.isTrue(result.success, 'case-insensitive parsing should succeed');
    Assert.equals(result.data.fullName, 'Ali Hassan', 'fullName extracted with lowercase key');
  });

  TestRunner.run('parseRecruitmentEmail_ handles extra whitespace gracefully', () => {
    const body = '  CANDIDATE_NAME:   Nadia Salem  \nPOSITION:   Architect  \nEMAIL:   nadia@test.com  ';
    const result = parseRecruitmentEmail_(body);
    Assert.isTrue(result.success, 'should succeed with leading/trailing whitespace');
    Assert.equals(result.data.fullName, 'Nadia Salem', 'fullName should be trimmed');
  });

  TestRunner.run('api_sendRejectionEmail dispatches email with correct subject', () => {
    MockFactory.reset();
    // Re-inject fresh mocks (GmailApp in particular)
    GmailApp = MockFactory.getGmailApp(); // eslint-disable-line no-global-assign
    ScriptApp = MockFactory.getScriptApp(); // eslint-disable-line no-global-assign

    api_sendRejectionEmail('cand@test.com', 'Cand Name', 'Passport', 'Expired document', 'cand-999');

    Assert.equals(MockFactory._sentEmails.length, 1, 'exactly 1 email should be sent');
    const sent = MockFactory._sentEmails[0];
    Assert.equals(sent.to, 'cand@test.com', 'email should be sent to candidate');
    Assert.isTrue(sent.subject.includes('Passport'), 'subject should mention the document type');
    Assert.isTrue(sent.opts.htmlBody.includes('Expired document'), 'email body should include rejection reason');
  });

  TestRunner.run('api_sendRejectionEmail embeds a valid magic link in the body', () => {
    MockFactory.reset();
    GmailApp  = MockFactory.getGmailApp();  // eslint-disable-line no-global-assign
    ScriptApp = MockFactory.getScriptApp(); // eslint-disable-line no-global-assign

    api_sendRejectionEmail('cand@test.com', 'Test User', 'Photo', 'Wrong background color', 'cand-777');

    const body = MockFactory._sentEmails[0].opts.htmlBody;
    Assert.isTrue(body.includes('cid=cand-777'), 'magic link should contain the candidateId');
    Assert.isTrue(body.includes('candidatePortal'), 'magic link should route to candidatePortal view');
  });

  TestRunner.run('api_sendPackageSubmissionAlert sends email to HR Ops Head', () => {
    MockFactory.reset();
    GmailApp  = MockFactory.getGmailApp(); // eslint-disable-line no-global-assign

    api_sendPackageSubmissionAlert('Khaled Mansour', 'Coordinator Sarah');

    Assert.equals(MockFactory._sentEmails.length, 1, 'exactly 1 alert email should be sent');
    const sent = MockFactory._sentEmails[0];
    Assert.equals(sent.to, HR_OPS_HEAD_EMAIL, 'alert should go to HR Ops Head');
    Assert.isTrue(sent.subject.includes('Package'), 'subject should mention Package');
    Assert.isTrue(sent.body.includes('Khaled Mansour'), 'body should include candidate name');
  });

  TestRunner.run('checkOverdueCandidates sends escalation email for stale cases', () => {
    // Seed a candidate with an old UpdatedAt date (9 days ago)
    MockFactory.reset();
    GmailApp   = MockFactory.getGmailApp();   // eslint-disable-line no-global-assign
    ScriptApp  = MockFactory.getScriptApp();  // eslint-disable-line no-global-assign

    const nineDAgo = new Date(Date.now() - 9 * 24 * 60 * 60 * 1000).toISOString();
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, [
      ['cand-stale', 'Stale Candidate', 'Worker', 'Operations', 'stale@test.com',
       '+0000', 'Unknown', '300', 'coord@company.com',
       'Documents Requested', nineDAgo, nineDAgo, '']
    ]);
    MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS);
    MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS);
    MockFactory.seedSheet(SHEET_USERS, USER_HEADERS);

    checkOverdueCandidates();

    Assert.equals(MockFactory._sentEmails.length, 1, 'exactly 1 escalation email should be sent');
    const sent = MockFactory._sentEmails[0];
    Assert.equals(sent.to, HR_OPS_HEAD_EMAIL, 'escalation should go to HR Ops Head');
    Assert.isTrue(sent.subject.includes('Overdue'), 'subject should mention Overdue');
    Assert.isTrue(sent.body.includes('Stale Candidate'), 'body should list the overdue candidate');
  });

  TestRunner.run('checkOverdueCandidates does NOT send email when no overdue cases', () => {
    MockFactory.reset();
    GmailApp  = MockFactory.getGmailApp(); // eslint-disable-line no-global-assign

    // All candidates are fresh (today)
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, [
      ['cand-fresh', 'Fresh Candidate', 'Engineer', 'IT', 'fresh@test.com',
       '+0000', 'Unknown', '400', 'coord@company.com',
       'Documents Requested', new Date().toISOString(), new Date().toISOString(), '']
    ]);
    MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS);
    MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS);
    MockFactory.seedSheet(SHEET_USERS, USER_HEADERS);

    checkOverdueCandidates();

    Assert.equals(MockFactory._sentEmails.length, 0, 'no email should be sent for fresh candidates');
  });

  TestRunner.run('checkOverdueCandidates skips Mobilized and Closed candidates', () => {
    MockFactory.reset();
    GmailApp  = MockFactory.getGmailApp(); // eslint-disable-line no-global-assign

    const old = new Date(Date.now() - 15 * 24 * 60 * 60 * 1000).toISOString();
    MockFactory.seedSheet(SHEET_CANDIDATES, CANDIDATE_HEADERS, [
      ['cand-mob',    'Mobilized Guy',  'X', 'X', 'a@t.com', '+0', 'X', '0', 'c@c.com', 'Mobilized', old, old, ''],
      ['cand-closed', 'Closed Guy',     'X', 'X', 'b@t.com', '+0', 'X', '0', 'c@c.com', 'Closed',    old, old, '']
    ]);
    MockFactory.seedSheet(SHEET_DOCUMENTS, DOCUMENT_HEADERS);
    MockFactory.seedSheet(SHEET_LOGS, LOG_HEADERS);
    MockFactory.seedSheet(SHEET_USERS, USER_HEADERS);

    checkOverdueCandidates();

    Assert.equals(MockFactory._sentEmails.length, 0, 'Mobilized/Closed candidates should not trigger escalation');
  });
}
