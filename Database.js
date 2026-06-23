/**
 * =========================================================
 * DATABASE MODULE (Database.gs)
 * =========================================================
 * All Google Sheets read/write operations happen here.
 * The Spreadsheet must have these sheets (tabs):
 *  - tbl_Candidates
 *  - tbl_Documents
 *  - tbl_Users
 *  - tbl_SystemLogs
 *
 * To link: Open your Google Sheet, copy its ID from the URL
 * (the long string between /d/ and /edit), and paste below.
 * =========================================================
 */

const SPREADSHEET_ID = '1yvmt1D5ag5ALuGd40jVauQSB8C7fqY9tIdIe3UHADZ0'; // <-- INSERT YOUR SHEET ID

// Sheet name constants
const SHEET_CANDIDATES   = 'tbl_Candidates';
const SHEET_DOCUMENTS    = 'tbl_Documents';
const SHEET_USERS        = 'tbl_Users';
const SHEET_LOGS         = 'tbl_SystemLogs';

/**
 * Helper: Returns a specific sheet by name.
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);
  return sheet;
}

/**
 * Helper: Generates a unique UUID for primary keys.
 */
function generateUUID_() {
  return Utilities.getUuid();
}

// ─────────────────────────────────────────────
// CANDIDATES
// ─────────────────────────────────────────────

/**
 * Returns all candidate rows as an array of objects.
 * Called from the frontend via: google.script.run.api_getAllCandidates()
 */
function api_getAllCandidates() {
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const candidates = rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return { success: true, data: candidates };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Creates a new candidate record in the Candidates sheet.
 * @param {object} candidateData - { fullName, position, department, email, phone, nationality, salary, coordinatorEmail }
 */
function api_createCandidate(candidateData) {
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const id = generateUUID_();
    const now = new Date().toISOString();
    
    sheet.appendRow([
      id,                                   // CandidateID
      candidateData.fullName,               // FullName
      candidateData.position,               // Position
      candidateData.department,             // Department
      candidateData.email,                  // Email
      candidateData.phone,                  // Phone
      candidateData.nationality,            // Nationality
      candidateData.salary,                 // OfferSalary
      candidateData.coordinatorEmail,       // AssignedCoordinatorEmail
      'Documents Requested',                // CurrentStatus
      now,                                  // CreatedAt
      now,                                  // UpdatedAt
      ''                                    // DriveFolderID (filled by DriveManager)
    ]);
    
    api_writeLog_(id, 'SYSTEM', 'Candidate Created: ' + candidateData.fullName);
    return { success: true, candidateId: id };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Updates a candidate's profile details.
 * @param {string} candidateId
 * @param {object} updates - { phone, notes }
 */
function api_updateCandidateDetails(candidateId, updates) {
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('CandidateID');
    const phoneCol = headers.indexOf('Phone');
    const notesCol = headers.indexOf('Notes'); // Column N
    const updatedCol = headers.indexOf('UpdatedAt');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === candidateId) {
        if (updates.phone !== undefined && phoneCol >= 0) {
          sheet.getRange(i + 1, phoneCol + 1).setValue(updates.phone);
        }
        if (updates.notes !== undefined) {
          // If Notes column missing, gracefully append it as Column N
          if (notesCol === -1) {
            sheet.getRange(1, headers.length + 1).setValue('Notes');
            sheet.getRange(i + 1, headers.length + 1).setValue(updates.notes);
          } else {
            sheet.getRange(i + 1, notesCol + 1).setValue(updates.notes);
          }
        }
        sheet.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
        api_writeLog_(candidateId, Session.getActiveUser().getEmail(), 'Profile Updated');
        return { success: true };
      }
    }
    return { success: false, error: 'Candidate not found.' };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Updates a candidate's status.
 * @param {string} candidateId
 * @param {string} newStatus
 */
function api_updateCandidateStatus(candidateId, newStatus) {
  try {
    const sheet = getSheet_(SHEET_CANDIDATES);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('CandidateID');
    const statusCol = headers.indexOf('CurrentStatus');
    const updatedCol = headers.indexOf('UpdatedAt');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === candidateId) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        sheet.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
        api_writeLog_(candidateId, Session.getActiveUser().getEmail(), 'Status Changed: ' + newStatus);
        return { success: true };
      }
    }
    return { success: false, error: 'Candidate not found.' };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────
// DOCUMENTS
// ─────────────────────────────────────────────

/**
 * Gets all documents for a specific candidate.
 * @param {string} candidateId
 */
function api_getDocumentsByCandidate(candidateId) {
  try {
    const sheet = getSheet_(SHEET_DOCUMENTS);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const documents = rows
      .filter(row => row[headers.indexOf('CandidateID')] === candidateId)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    return { success: true, data: documents };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Approves or rejects a specific document.
 * @param {string} documentId
 * @param {string} action - 'Approved' or 'Rejected'
 * @param {string} remarks - Required if rejecting
 */
function api_reviewDocument(documentId, action, remarks) {
  try {
    const sheet = getSheet_(SHEET_DOCUMENTS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const docIdCol    = headers.indexOf('DocumentID');
    const statusCol   = headers.indexOf('ApprovalStatus');
    const remarkCol   = headers.indexOf('Remarks');
    const reviewCol   = headers.indexOf('ReviewedAt');   // Optional — may not exist in older sheets
    const approverCol = headers.indexOf('ApprovedBy');
    const actor       = Session.getActiveUser().getEmail();

    // Guard: validate required columns exist
    if (docIdCol < 0 || statusCol < 0 || remarkCol < 0 || approverCol < 0) {
      return { success: false, error: 'tbl_Documents is missing required columns. Check SheetSchema.md.' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][docIdCol] === documentId) {
        sheet.getRange(i + 1, statusCol   + 1).setValue(action);
        sheet.getRange(i + 1, remarkCol   + 1).setValue(remarks || '');
        sheet.getRange(i + 1, approverCol + 1).setValue(actor);

        // Only write ReviewedAt if the column exists (column K in the updated schema)
        if (reviewCol >= 0) {
          sheet.getRange(i + 1, reviewCol + 1).setValue(new Date().toISOString());
        }

        api_writeLog_(
          data[i][headers.indexOf('CandidateID')],
          actor,
          'Document ' + action + ': ' + data[i][headers.indexOf('DocType')]
        );
        return { success: true };
      }
    }
    return { success: false, error: 'Document not found.' };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

// ─────────────────────────────────────────────
// AUDIT LOG
// ─────────────────────────────────────────────

/**
 * Internal-only: Writes an immutable audit log entry.
 */
function api_writeLog_(candidateId, actor, event) {
  try {
    const sheet = getSheet_(SHEET_LOGS);
    sheet.appendRow([
      generateUUID_(),              // LogID
      new Date().toISOString(),     // Timestamp
      candidateId,                  // CandidateID
      actor,                        // Actor
      event                         // Event description
    ]);
  } catch(e) {
    Logger.log('AUDIT LOG ERROR: ' + e.message);
  }
}

/**
 * Public: Retrieves all audit log entries for a candidate.
 * @param {string} candidateId
 */
function api_getAuditLog(candidateId) {
  try {
    const sheet = getSheet_(SHEET_LOGS);
    const [headers, ...rows] = sheet.getDataRange().getValues();
    const logs = rows
      .filter(row => row[headers.indexOf('CandidateID')] === candidateId)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    return { success: true, data: logs };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}
