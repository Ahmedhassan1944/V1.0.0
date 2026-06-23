/**
 * =========================================================
 * DRIVE MANAGER MODULE (DriveManager.gs)
 * =========================================================
 * Manages all Google Drive operations:
 * - Creating candidate folders automatically
 * - Uploading binary files from frontend
 * - Preventing duplicate folder creation
 * =========================================================
 */

// Root folder name in Google Drive for all candidate files
const ROOT_FOLDER_NAME = 'Oman Mobilization';

/**
 * Helper: Gets or creates the root "Oman Mobilization" folder in Drive.
 */
function getRootFolder_() {
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(ROOT_FOLDER_NAME);
}

/**
 * Creates a candidate's personal folder inside the root folder.
 * Naming convention: [UUID] - [FullName]
 * Prevents duplicates by searching first.
 *
 * @param {string} candidateId - The UUID of the candidate
 * @param {string} fullName    - The candidate's full name
 * @returns {string} The folder ID to be stored in the Candidates sheet
 */
function api_createCandidateFolder(candidateId, fullName) {
  try {
    const root = getRootFolder_();
    const folderName = candidateId + ' - ' + fullName;
    
    // Check if folder already exists (prevents duplicates)
    const existing = root.getFoldersByName(folderName);
    if (existing.hasNext()) {
      const existingFolder = existing.next();
      Logger.log('Folder already exists: ' + existingFolder.getId());
      return { success: true, folderId: existingFolder.getId(), url: existingFolder.getUrl() };
    }

    // Create the new folder
    const newFolder = root.createFolder(folderName);
    
    // Optional: Restrict sharing — no public access
    newFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    
    // Store the folder ID back into the Candidates sheet
    api_storeFolderIdForCandidate_(candidateId, newFolder.getId());
    
    api_writeLog_(candidateId, 'SYSTEM', 'Drive Folder Created: ' + folderName);
    return { success: true, folderId: newFolder.getId(), url: newFolder.getUrl() };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Uploads a document file into the candidate's Drive folder.
 * Receives base64-encoded binary data from the frontend.
 *
 * @param {string} candidateId  - Links to the Candidates sheet row
 * @param {string} folderId     - The candidate's Google Drive folder ID
 * @param {string} docType      - E.g., 'Passport', 'Photo', 'Medical'
 * @param {string} fileName     - Original file name
 * @param {string} base64Data   - The file contents encoded in base64
 * @param {string} mimeType     - 'application/pdf', 'image/jpeg', or 'image/png'
 */
function api_uploadFileToDrive(candidateId, folderId, docType, fileName, base64Data, mimeType) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    
    // Decode and build the file blob
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, docType + '_' + fileName);
    
    // Remove any previous version of the same DocType (for re-uploads)
    const existingFiles = folder.getFilesByName(docType + '_' + fileName);
    while (existingFiles.hasNext()) {
      const oldFile = existingFiles.next();
      oldFile.setName('[ARCHIVE] ' + oldFile.getName()); // Soft-rename instead of delete
    }
    
    // Create the new file
    const uploadedFile = folder.createFile(blob);
    const fileUrl = uploadedFile.getUrl();
    
    // Write metadata to the Documents sheet
    const docId = api_writeDocumentRecord_(candidateId, docType, fileName, fileUrl, mimeType);
    
    api_writeLog_(candidateId, Session.getActiveUser().getEmail(), 'Document Uploaded: ' + docType);
    return { success: true, fileUrl: fileUrl, documentId: docId };
  } catch(e) {
    Logger.log(e);
    return { success: false, error: e.message };
  }
}

/**
 * Internal: Writes a new document metadata row into the Documents sheet.
 */
function api_writeDocumentRecord_(candidateId, docType, fileName, fileUrl, mimeType) {
  const sheet = getSheet_(SHEET_DOCUMENTS);
  const docId = generateUUID_();
  const now = new Date().toISOString();
  
  // Determine the current version number for this DocType
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const existingVersions = data
    .filter(row => row[headers.indexOf('CandidateID')] === candidateId &&
                   row[headers.indexOf('DocType')] === docType)
    .length;
  const versionNumber = existingVersions + 1;
  
  sheet.appendRow([
    docId,              // DocumentID
    candidateId,        // CandidateID
    docType,            // DocType
    fileName,           // FileName
    fileUrl,            // FileURL
    now,                // UploadDate
    'Pending Review',   // ApprovalStatus
    '',                 // ApprovedBy
    versionNumber,      // VersionNumber
    ''                  // Remarks
  ]);
  
  return docId;
}

/**
 * Internal: Writes Back the Drive Folder ID into the Candidates sheet.
 */
function api_storeFolderIdForCandidate_(candidateId, folderId) {
  const sheet = getSheet_(SHEET_CANDIDATES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('CandidateID');
  const folderCol = headers.indexOf('DriveFolderID');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === candidateId) {
      sheet.getRange(i + 1, folderCol + 1).setValue(folderId);
      return;
    }
  }
}
