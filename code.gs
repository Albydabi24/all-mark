/**
 * ========================================
 * ALL-MARK DASHBOARD V2.0 - GOOGLE APPS SCRIPT
 * ========================================
 * Complete Rebuild with Enhanced Features
 * - ID-based Authentication
 * - Selective Sync System
 * - Bulk Operations Support
 * - Month Filtering
 * - Progress Tracking
 * - Delete Functionality
 * 
 * Spreadsheet ID: 1doN1q9h_AS68u1mH3NncDKC0ed9SMvtHV5UVV8wP4zA
 * 
 * @author CMO JOY - Muhammad Nurul Qolbi
 * @version 2.0.0
 */

// ========================================
// GLOBAL CONFIGURATION
// ========================================

const SPREADSHEET_ID = '1doN1q9h_AS68u1mH3NncDKC0ed9SMvtHV5UVV8wP4zA';

const SHEET_NAMES = {
  MATRIX_REGULER: 'Matrix Reguler',
  MATRIX_CC: 'Matrix CC',
  MATRIX_CW: 'Matrix CW',
  MATRIX_GD: 'Matrix GD',
  USERS: 'Users'
};

// Role definitions
const ROLES = {
  ADMIN: ['Founder', 'Co-Founder', 'Chief Marketing Officer (CMO)', 'Social Media Specialist (SMS)'],
  DIVISIONS: {
    CC: 'Content Creator (CC)',
    CW: 'Content Writer (CW)',
    GD: 'Graphic Designer (GD)'
  }
};

// User credentials mapping (ID to Role)
const USER_CREDENTIALS = {
  'founder-joy': {
    password: 'founder-joy26',
    role: 'Founder',
    fullName: 'Founder JOY'
  },
  'co-founder-joy': {
    password: 'co-founder-joy26',
    role: 'Co-Founder',
    fullName: 'Co-Founder JOY'
  },
  'cmo-joy': {
    password: 'cmo-joy26',
    role: 'Chief Marketing Officer (CMO)',
    fullName: 'CMO JOY'
  },
  'sms-joy': {
    password: 'sms-joy26',
    role: 'Social Media Specialist (SMS)',
    fullName: 'SMS JOY'
  },
  'cc-joy': {
    password: 'cc-joy26',
    role: 'Content Creator (CC)',
    fullName: 'Content Creator JOY'
  },
  'cw-joy': {
    password: 'cw-joy26',
    role: 'Content Writer (CW)',
    fullName: 'Content Writer JOY'
  },
  'gd-joy': {
    password: 'gd-joy26',
    role: 'Graphic Designer (GD)',
    fullName: 'Graphic Designer JOY'
  }
};

// PIC lists per division
const PIC_LISTS = {
  CC: ['Obi', 'Refan', 'Desy', 'Caitlin', 'Mia', 'Falen', 'Qonita'],
  CW: ['Obi', 'Astri', 'Fifa', 'Nadiyah', 'Klosse', 'Afra', 'Danis', 'Asha'],
  GD: ['Obi', 'Gopal', 'Shafni', 'Nuri', 'Diana', 'Nopal', 'Shelby', 'Bayu'],
  SMS: ['Obi', 'Zahra', 'Marsha', 'Juju', 'Nichell', 'Sauma']
};

// Review status options
const REVIEW_STATUS = ['Reviewed', 'Unreviewed yet', 'On hold'];

// Content types
const CONTENT_TYPES = ['IGS-CW', 'IGS-SMS', 'IGR', 'IGF', 'Linkedin', 'Tiktok'];

// Upload times
const UPLOAD_TIMES = ['12.00 WIB', '17.00 WIB'];

// Month filter options (NOV-25 to DES-26)
const MONTH_FILTERS = [
  'NOV-25', 'DES-25', 'JAN-26', 'FEB-26', 'MAR-26', 'APR-26',
  'MAY-26', 'JUN-26', 'JUL-26', 'AUG-26', 'SEP-26', 'OCT-26',
  'NOV-26', 'DES-26'
];

// ========================================
// WEB APP ENTRY POINT
// ========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('All-Mark Dashboard V2.0')
    .setFaviconUrl('https://www.google.com/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ========================================
// AUTHENTICATION FUNCTIONS
// ========================================

/**
 * Validates user login with ID-based authentication
 * Auto-detects role from username
 */
function validateLogin(username, password) {
  try {
    if (!username || !password) {
      return {
        success: false,
        message: 'Username and password are required'
      };
    }

    const inputUser = username.toString().trim().toLowerCase();
    const inputPass = password.toString().trim();

    // Check if user exists in credentials
    if (!USER_CREDENTIALS[inputUser]) {
      return {
        success: false,
        message: 'Invalid username or password'
      };
    }
    
    const userCred = USER_CREDENTIALS[inputUser];
    
    // Validate password
    if (userCred.password !== inputPass) {
      return {
        success: false,
        message: 'Invalid username or password'
      };
    }
    
    // Successful login
    return {
      success: true,
      user: {
        username: inputUser,
        role: userCred.role,
        fullName: userCred.fullName,
        isAdmin: ROLES.ADMIN.includes(userCred.role)
      }
    };
    
  } catch (error) {
    Logger.log('Login error: ' + error.toString());
    return {
      success: false,
      message: 'Login error: ' + error.toString()
    };
  }
}

// ========================================
// MATRIX REGULER FUNCTIONS
// ========================================

/**
 * Gets all data from Matrix Reguler sheet
 */
function getMatrixRegulerData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAMES.MATRIX_REGULER);
    
    if (!sheet) {
      sheet = createMatrixRegulerSheet(ss);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    const rows = data.slice(1).map((row, index) => {
      return {
        no: index + 1,
        uploadDeadline: formatDate(row[1]),
        uploadDay: row[2] || '',
        uploadTime: row[3] || '',
        contentIdeas: row[4] || '',
        references: row[5] || '',
        smsDirection: row[6] || '',
        contentType: row[7] || '',
        picSMS: row[8] || '',
        syncToCC: row[9] || false,
        syncToCW: row[10] || false,
        ccResult: row[11] || '',
        cwResult: row[12] || '',
        gdResult: row[13] || ''
      };
    });
    
    return { success: true, data: rows };
    
  } catch (error) {
    Logger.log('Error getting Matrix Reguler data: ' + error.toString());
    return { success: false, message: error.toString(), data: [] };
  }
}

/**
 * Saves Matrix Reguler data with selective sync
 */
function saveMatrixRegulerData(rowsData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAMES.MATRIX_REGULER);
    
    if (!sheet) {
      sheet = createMatrixRegulerSheet(ss);
    }
    
    // Clear existing data (except headers)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // Prepare data for insertion
    const dataToInsert = rowsData.map((row, index) => {
      const uploadDate = parseDateString(row.uploadDeadline);
      const dayName = uploadDate ? getDayName(uploadDate) : '';
      
      return [
        index + 1,
        row.uploadDeadline || '',
        dayName,
        row.uploadTime || '',
        row.contentIdeas || '',
        row.references || '',
        row.smsDirection || '',
        row.contentType || '',
        row.picSMS || '',
        row.syncToCC || false,
        row.syncToCW || false,
        row.ccResult || '',
        row.cwResult || '',
        row.gdResult || ''
      ];
    });
    
    if (dataToInsert.length > 0) {
      sheet.getRange(2, 1, dataToInsert.length, dataToInsert[0].length).setValues(dataToInsert);
    }
    
    // Selective sync to division matrices
    syncMatrixRegulerToDivisions();
    
    return { success: true, message: 'Data saved successfully' };
    
  } catch (error) {
    Logger.log('Error saving Matrix Reguler data: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Deletes a row from Matrix Reguler and syncs deletion across all matrices
 */
function deleteMatrixRegulerRow(rowNo) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const regulerSheet = ss.getSheetByName(SHEET_NAMES.MATRIX_REGULER);
    
    if (!regulerSheet) {
      return { success: false, message: 'Matrix Reguler not found' };
    }
    
    const data = regulerSheet.getDataRange().getValues();
    
    // Find and delete the row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowNo) {
        regulerSheet.deleteRow(i + 1);
        
        // Also delete from division matrices
        deleteDivisionRow(ss, SHEET_NAMES.MATRIX_CC, rowNo);
        deleteDivisionRow(ss, SHEET_NAMES.MATRIX_CW, rowNo);
        deleteDivisionRow(ss, SHEET_NAMES.MATRIX_GD, rowNo);
        
        return { success: true, message: 'Row deleted successfully' };
      }
    }
    
    return { success: false, message: 'Row not found' };
    
  } catch (error) {
    Logger.log('Error deleting row: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Creates Matrix Reguler sheet with new structure
 */
function createMatrixRegulerSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAMES.MATRIX_REGULER);
  
  const headers = [
    'No',
    'Upload Deadline',
    'Upload Day',
    'Upload Time',
    'Content Ideas',
    'References',
    'SMS Direction',
    'Content Type',
    'PIC SMS',
    'Sync to CC',
    'Sync to CW',
    'CC Result',
    'CW Result',
    'GD Result'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header with new color scheme
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#30678e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set column widths (optimized for responsiveness)
  sheet.setColumnWidth(1, 50);   // No
  sheet.setColumnWidth(2, 110);  // Upload Deadline
  sheet.setColumnWidth(3, 90);   // Upload Day
  sheet.setColumnWidth(4, 90);   // Upload Time
  sheet.setColumnWidth(5, 180);  // Content Ideas
  sheet.setColumnWidth(6, 130);  // References
  sheet.setColumnWidth(7, 150);  // SMS Direction
  sheet.setColumnWidth(8, 110);  // Content Type
  sheet.setColumnWidth(9, 100);  // PIC SMS
  sheet.setColumnWidth(10, 80);  // Sync to CC
  sheet.setColumnWidth(11, 80);  // Sync to CW
  sheet.setColumnWidth(12, 130); // CC Result
  sheet.setColumnWidth(13, 130); // CW Result
  sheet.setColumnWidth(14, 130); // GD Result
  
  return sheet;
}

/**
 * Syncs Matrix Reguler to division matrices with selective logic
 */
function syncMatrixRegulerToDivisions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const regulerSheet = ss.getSheetByName(SHEET_NAMES.MATRIX_REGULER);
    
    if (!regulerSheet) return;
    
    const data = regulerSheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    // Sync to each division based on sync flags
    syncToDivisionMatrix(ss, SHEET_NAMES.MATRIX_CC, data, 9);  // Column 9 = syncToCC
    syncToDivisionMatrix(ss, SHEET_NAMES.MATRIX_CW, data, 10); // Column 10 = syncToCW
    syncToDivisionMatrix(ss, SHEET_NAMES.MATRIX_GD, data, -1); // Always sync GD
    
  } catch (error) {
    Logger.log('Error syncing to divisions: ' + error.toString());
  }
}

/**
 * Syncs to specific division matrix with selective logic
 */
function syncToDivisionMatrix(ss, sheetName, regulerData, syncColumnIndex) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = createDivisionMatrixSheet(ss, sheetName);
  }
  
  // Detect header format
  const format = detectHeaderFormat(sheet);
  
  const existingData = sheet.getDataRange().getValues();
  const existingRows = existingData.slice(1);
  
  for (let i = 1; i < regulerData.length; i++) {
    const regulerRow = regulerData[i];
    const rowNo = regulerRow[0];
    
    // Check if should sync (for CC/CW) or always sync (for GD)
    const shouldSync = syncColumnIndex === -1 || regulerRow[syncColumnIndex] === true || regulerRow[syncColumnIndex] === 'TRUE';
    
    if (!shouldSync) continue;
    
    // Check if row exists
    let existingRowIndex = -1;
    for (let j = 0; j < existingRows.length; j++) {
      if (existingRows[j][0] === rowNo) {
        existingRowIndex = j;
        break;
      }
    }
    
    if (existingRowIndex === -1) {
      // New row - create with correct format
      let newRow;
      if (format.format === 'database') {
        // Database format: No | Content Idea | Reference Link | Content Type | Upload Date | Upload Day | Upload Time | Content Creating Deadline | Content Result | PJ | Checked by REVIEWER | Status
        newRow = [
          rowNo,
          regulerRow[4] || '',  // Content Idea
          regulerRow[5] || '',  // Reference Link
          regulerRow[7] || '',  // Content Type
          regulerRow[1] || '',  // Upload Date
          regulerRow[2] || '',  // Upload Day
          regulerRow[3] || '',  // Upload Time
          '',                   // Content Creating Deadline
          '',                   // Content Result
          '',                   // PJ
          'Unreviewed yet',     // Checked by REVIEWER
          'Unreviewed yet'      // Status
        ];
      } else if (format.format === 'new') {
        // New format: NO | IDEA | SMS DIRECTION | REFERENCE | TYPE | UPLOAD DEADLINE | DAY | TIME | [DEADLINE] | RESULT LINK | PIC | [REVIEWER] | REVIEW
        const isCW = sheetName === SHEET_NAMES.MATRIX_CW;
        const deadlineCol = isCW ? 'BRIEF DEADLINE' : (sheetName === SHEET_NAMES.MATRIX_CC ? 'VIDEO DEADLINE' : 'DESIGN DEADLINE');
        
        if (isCW) {
          // CW format with Reviewer column
          newRow = [
            rowNo,              // NO
            regulerRow[4] || '',  // IDEA (Content Ideas)
            regulerRow[6] || '',  // SMS DIRECTION
            regulerRow[5] || '',  // REFERENCE (References)
            regulerRow[7] || '',  // TYPE (Content Type)
            regulerRow[1] || '',  // UPLOAD DEADLINE (Upload Deadline)
            regulerRow[2] || '',  // DAY (Upload Day)
            regulerRow[3] || '',  // TIME (Upload Time)
            '',                   // BRIEF DEADLINE
            '',                   // RESULT LINK
            '',                   // PIC
            '',                   // REVIEWER
            'Unreviewed yet'      // REVIEW
          ];
        } else {
          // CC/GD format without Reviewer
          newRow = [
            rowNo,              // NO
            regulerRow[4] || '',  // IDEA (Content Ideas)
            regulerRow[6] || '',  // SMS DIRECTION
            regulerRow[5] || '',  // REFERENCE (References)
            regulerRow[7] || '',  // TYPE (Content Type)
            regulerRow[1] || '',  // UPLOAD DEADLINE (Upload Deadline)
            regulerRow[2] || '',  // DAY (Upload Day)
            regulerRow[3] || '',  // TIME (Upload Time)
            '',                   // [DEADLINE]
            '',                   // RESULT LINK
            '',                   // PIC
            'Unreviewed yet'      // REVIEW
          ];
        }
      } else {
        // Old format without SMS DIRECTION
        newRow = [
          rowNo,
          regulerRow[4] || '',  // Content Ideas
          regulerRow[5] || '',  // References
          regulerRow[7] || '',  // Content Type
          regulerRow[1] || '',  // Upload Date
          regulerRow[2] || '',  // Upload Day
          regulerRow[3] || '',  // Upload Time
          '',                   // Deadline
          '',                   // Result
          '',                   // PIC
          'Unreviewed yet'      // Review
        ];
      }
      
      sheet.appendRow(newRow);
    } else {
      // Update existing row (only auto-populated fields)
      const targetRow = existingRowIndex + 2;
      if (format.format === 'database') {
        sheet.getRange(targetRow, 2).setValue(regulerRow[4] || ''); // Content Idea
        sheet.getRange(targetRow, 3).setValue(regulerRow[5] || ''); // Reference Link
        sheet.getRange(targetRow, 4).setValue(regulerRow[7] || ''); // Content Type
        sheet.getRange(targetRow, 5).setValue(regulerRow[1] || ''); // Upload Date
        sheet.getRange(targetRow, 6).setValue(regulerRow[2] || ''); // Upload Day
        sheet.getRange(targetRow, 7).setValue(regulerRow[3] || ''); // Upload Time
      } else if (format.format === 'new') {
        // New format: NO | IDEA | SMS DIRECTION | REFERENCE | TYPE | UPLOAD DEADLINE | DAY | TIME | [DEADLINE] | RESULT LINK | PIC | [REVIEWER] | REVIEW
        sheet.getRange(targetRow, 2).setValue(regulerRow[4] || ''); // IDEA
        sheet.getRange(targetRow, 3).setValue(regulerRow[6] || ''); // SMS DIRECTION
        sheet.getRange(targetRow, 4).setValue(regulerRow[5] || ''); // REFERENCE
        sheet.getRange(targetRow, 5).setValue(regulerRow[7] || ''); // TYPE
        sheet.getRange(targetRow, 6).setValue(regulerRow[1] || ''); // UPLOAD DEADLINE
        sheet.getRange(targetRow, 7).setValue(regulerRow[2] || ''); // DAY
        sheet.getRange(targetRow, 8).setValue(regulerRow[3] || ''); // TIME
      } else {
        sheet.getRange(targetRow, 2).setValue(regulerRow[4] || ''); // Ideas
        sheet.getRange(targetRow, 3).setValue(regulerRow[5] || ''); // Reference
        sheet.getRange(targetRow, 4).setValue(regulerRow[7] || ''); // Type
        sheet.getRange(targetRow, 5).setValue(regulerRow[1] || ''); // Date
        sheet.getRange(targetRow, 6).setValue(regulerRow[2] || ''); // Day
        sheet.getRange(targetRow, 7).setValue(regulerRow[3] || ''); // Time
      }
    }
  }
}

// ========================================
// DIVISION MATRIX FUNCTIONS
// ========================================

/**
 * Detects which header format a sheet is using
 * Returns format type: 'new' (current app format), 'database' (old DB format), or 'old' (legacy)
 */
function detectHeaderFormat(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return { format: 'new', headers: [] };
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Check for new format markers (current app format with IDEA/TYPE/UPLOAD DEADLINE)
    const hasNewFormat = headers.includes('IDEA') || headers.includes('TYPE') || headers.includes('SMS DIRECTION') || headers.includes('UPLOAD DEADLINE');
    // Check for database format (old format used in sheets)
    const hasDatabaseFormat = headers.includes('Content Idea') || headers.includes('Reference Link') || headers.includes('Upload Date');
    // Check for old format markers (legacy)
    const hasOldFormat = headers.includes('Content Type') && !hasDatabaseFormat && !hasNewFormat;
    
    if (hasNewFormat) {
      return { format: 'new', isNew: true, isOld: false, headers: headers };
    } else if (hasDatabaseFormat) {
      return { format: 'database', isNew: false, isOld: false, headers: headers };
    } else {
      return { format: 'old', isNew: false, isOld: true, headers: headers };
    }
  } catch (error) {
    Logger.log('Error detecting header format: ' + error.toString());
    return { format: 'new', isNew: true, isOld: false, headers: [] };
  }
}

/**
 * Gets division matrix data with backwards compatibility for old/new headers
 */
function getDivisionMatrixData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = createDivisionMatrixSheet(ss, sheetName);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    // Detect which header format is being used
    const format = detectHeaderFormat(sheet);
    
    // Map column indices based on detected format
    let cols;
    if (format.format === 'new') {
      // New format (current app format): NO | IDEA | SMS DIRECTION | REFERENCE | TYPE | UPLOAD DEADLINE | DAY | TIME | [DEADLINE] | RESULT LINK | PIC | [REVIEWER] | REVIEW
      // For CW: deadline=8, result=9, pic=10, reviewer=11, review=12
      // For CC/GD: deadline=8, result=9, pic=10, review=11
      const hasReviewer = sheetName === SHEET_NAMES.MATRIX_CW;
      cols = {
        no: 0, idea: 1, smsDir: 2, ref: 3, type: 4,
        date: 5, day: 6, time: 7, deadline: 8,
        result: 9, pic: 10, review: hasReviewer ? 12 : 11
      };
      if (hasReviewer) {
        cols.reviewer = 11;
      }
    } else if (format.format === 'database') {
      // Database format (old): No | Content Idea | Reference Link | Content Type | Upload Date | Upload Day | Upload Time | Content Creating Deadline | Content Result | PJ | Checked by REVIEWER | Status
      cols = {
        no: 0, idea: 1, smsDir: -1, ref: 2, type: 3,
        date: 4, day: 5, time: 6, deadline: 7,
        result: 8, pic: 9, review: 10, status: 11
      };
    } else {
      // Old format (legacy): NO | Content Idea | Reference Link | Content Type | Upload Date | Upload Day | Upload Hour | Deadline | Result | PIC | Review
      cols = {
        no: 0, idea: 1, smsDir: -1, ref: 2, type: 3,
        date: 4, day: 5, time: 6, deadline: 7,
        result: 8, pic: 9, review: 10
      };
    }
    
    const rows = data.slice(1).map(row => {
      const rowData = {
        no: row[cols.no] || '',
        idea: row[cols.idea] || '',
        smsDirection: cols.smsDir >= 0 ? (row[cols.smsDir] || '') : '',
        reference: row[cols.ref] || '',
        contentType: row[cols.type] || '',
        uploadDate: formatDate(row[cols.date]),
        uploadDay: row[cols.day] || '',
        uploadTime: row[cols.time] || '',
        deadline: formatDate(row[cols.deadline]),
        resultLink: row[cols.result] || '',
        pic: row[cols.pic] || '',
        review: row[cols.review] || 'Unreviewed yet'
      };
      
      // Add reviewer field for CW matrix
      if (sheetName === SHEET_NAMES.MATRIX_CW) {
        if (format.format === 'new' && cols.reviewer !== undefined) {
          rowData.reviewer = row[cols.reviewer] || '';
        } else if (row.length > cols.pic + 1) {
          rowData.reviewer = row[cols.pic + 1] || '';
        }
      }
      
      return rowData;
    });
    
    return { success: true, data: rows };
    
  } catch (error) {
    Logger.log('Error getting division matrix data: ' + error.toString());
    return { success: false, message: error.toString(), data: [] };
  }
}

function getMatrixCCData() { return getDivisionMatrixData(SHEET_NAMES.MATRIX_CC); }
function getMatrixCWData() { return getDivisionMatrixData(SHEET_NAMES.MATRIX_CW); }

/**
 * Gets GD matrix data with CC/CW result links from regular matrix
 * Also filters out "No Idea" entries
 */
function getMatrixGDData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get regular matrix data for CC/CW result links
    const regulerSheet = ss.getSheetByName(SHEET_NAMES.MATRIX_REGULER);
    const regulerData = regulerSheet ? regulerSheet.getDataRange().getValues() : [];
    const regulerMap = {};
    for (let i = 1; i < regulerData.length; i++) {
      const rowNo = regulerData[i][0];
      regulerMap[rowNo] = {
        ccResult: regulerData[i][11] || '',
        cwResult: regulerData[i][12] || ''
      };
    }
    
    // Get GD matrix data
    let gdSheet = ss.getSheetByName(SHEET_NAMES.MATRIX_GD);
    if (!gdSheet) {
      gdSheet = createDivisionMatrixSheet(ss, SHEET_NAMES.MATRIX_GD);
      return { success: true, data: [] };
    }
    
    const data = gdSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    const format = detectHeaderFormat(gdSheet);
    
    let cols;
    if (format.format === 'database') {
      cols = {
        no: 0, idea: 1, smsDir: -1, ref: 2, type: 3,
        date: 4, day: 5, time: 6, deadline: 7,
        result: 8, pic: 9, review: 10, status: 11
      };
    } else if (format.isNew) {
      cols = {
        no: 0, idea: 1, smsDir: 2, ref: 3, type: 4,
        date: 5, day: 6, time: 7, deadline: 8,
        result: 9, pic: 10, review: 11
      };
    } else {
      cols = {
        no: 0, idea: 1, smsDir: -1, ref: 2, type: 3,
        date: 4, day: 5, time: 6, deadline: 7,
        result: 8, pic: 9, review: 10
      };
    }
    
    const rows = data.slice(1)
      .map(row => {
        const idea = (row[cols.idea] || '').toString().trim();
        // Filter out "No Idea" entries and rows with only dashes
        if (idea.toLowerCase().includes('no idea') && idea.replace(/[^a-z0-9]/gi, '').length < 10) {
          return null;
        }
        // Filter out rows that are mostly dashes (like "---")
        const allCells = row.join('').replace(/\s/g, '');
        if (allCells.replace(/[-]/g, '').length < 3) {
          return null;
        }
        
        const rowNo = row[cols.no];
        // Skip rows with invalid row numbers
        if (!rowNo || rowNo === '' || rowNo === '-') {
          return null;
        }
        
        const regulerInfo = regulerMap[rowNo] || { ccResult: '', cwResult: '' };
        
        // Clean resultLink - remove any HTML tags that might have been accidentally stored
        let resultLink = (row[cols.result] || '').toString().trim();
        // Remove HTML tags if present
        resultLink = resultLink.replace(/<[^>]*>/g, '').trim();
        // Remove any remaining HTML entities or fragments
        resultLink = resultLink.replace(/&[^;]+;/g, '').trim();
        
        return {
          no: rowNo || '',
          idea: idea || '',
          smsDirection: cols.smsDir >= 0 ? (row[cols.smsDir] || '').toString().trim() : '',
          reference: (row[cols.ref] || '').toString().trim(),
          contentType: (row[cols.type] || '').toString().trim(),
          uploadDate: formatDate(row[cols.date]),
          uploadDay: (row[cols.day] || '').toString().trim(),
          uploadTime: (row[cols.time] || '').toString().trim(),
          deadline: formatDate(row[cols.deadline]),
          resultLink: resultLink,
          pic: (row[cols.pic] || '').toString().trim(),
          review: (row[cols.review] || 'Unreviewed yet').toString().trim(),
          ccResult: (regulerInfo.ccResult || '').toString().trim(),
          cwResult: (regulerInfo.cwResult || '').toString().trim()
        };
      })
      .filter(row => row !== null);
    
    return { success: true, data: rows };
    
  } catch (error) {
    Logger.log('Error getting GD matrix data: ' + error.toString());
    return { success: false, message: error.toString(), data: [] };
  }
}

/**
 * Saves single row from division matrix
 */
function saveDivisionRow(sheetName, rowData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    // Detect header format
    const format = detectHeaderFormat(sheet);
    
    // Map column indices for editable fields
    let cols;
    if (format.format === 'new') {
      // New format: NO | IDEA | SMS DIRECTION | REFERENCE | TYPE | UPLOAD DEADLINE | DAY | TIME | [DEADLINE] | RESULT LINK | PIC | [REVIEWER] | REVIEW
      // For CW: deadline=8, result=9, pic=10, reviewer=11, review=12
      // For CC/GD: deadline=8, result=9, pic=10, review=11
      const isCW = sheetName === SHEET_NAMES.MATRIX_CW;
      cols = { deadline: 8, result: 9, pic: 10, review: isCW ? 12 : 11 };
      if (isCW) {
        cols.reviewer = 11;
      }
    } else if (format.format === 'database') {
      // Database format: Content Creating Deadline=7, Content Result=8, PJ=9, Checked by REVIEWER=10, Status=11
      cols = { deadline: 7, result: 8, pic: 9, review: 10, status: 11 };
    } else {
      // Old format: Deadline=7, Result=8, PIC=9, Review=10
      cols = { deadline: 7, result: 8, pic: 9, review: 10 };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find the row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowData.no) {
        // Update the row (columns are 1-indexed in setRange)
        const targetRow = i + 1;
        sheet.getRange(targetRow, cols.deadline + 1).setValue(rowData.deadline || '');
        sheet.getRange(targetRow, cols.result + 1).setValue(rowData.resultLink || '');
        sheet.getRange(targetRow, cols.pic + 1).setValue(rowData.pic || '');
        
        // Save reviewer field for CW matrix
        if (sheetName === SHEET_NAMES.MATRIX_CW && rowData.reviewer !== undefined) {
          if (format.format === 'new' && cols.reviewer !== undefined) {
            sheet.getRange(targetRow, cols.reviewer + 1).setValue(rowData.reviewer || '');
          } else if (cols.pic !== undefined) {
            sheet.getRange(targetRow, cols.pic + 2).setValue(rowData.reviewer || '');
          }
        }
        
        const reviewValue = rowData.review || 'Unreviewed yet';
        sheet.getRange(targetRow, cols.review + 1).setValue(reviewValue);
        // Update Status column if it exists (database format)
        if (cols.status !== undefined) {
          sheet.getRange(targetRow, cols.status + 1).setValue(reviewValue);
        }
        
        return { success: true, message: 'Row saved successfully' };
      }
    }
    
    return { success: false, message: 'Row not found' };
    
  } catch (error) {
    Logger.log('Error saving division row: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function saveCCRow(rowData) { return saveDivisionRow(SHEET_NAMES.MATRIX_CC, rowData); }
function saveCWRow(rowData) { return saveDivisionRow(SHEET_NAMES.MATRIX_CW, rowData); }
function saveGDRow(rowData) { return saveDivisionRow(SHEET_NAMES.MATRIX_GD, rowData); }

/**
 * Sends division content result back to Matrix Reguler
 * Also triggers auto-sync to GD when CC or CW completes work
 */
function sendToMatrixReguler(division, rowNo, contentResult) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.MATRIX_REGULER);
    
    if (!sheet) {
      return { success: false, message: 'Matrix Reguler not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const columnIndex = division === 'CC' ? 12 : division === 'CW' ? 13 : 14; // CC=12, CW=13, GD=14
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowNo) {
        sheet.getRange(i + 1, columnIndex).setValue(contentResult);
        
        // Auto-sync to GD matrix when CC or CW completes work
        if (division === 'CC' || division === 'CW') {
          autoSyncCompletedToGD(division, rowNo);
        }
        
        return { success: true, message: 'Sent to Matrix Reguler successfully' };
      }
    }
    
    return { success: false, message: 'Row not found' };
    
  } catch (error) {
    Logger.log('Error sending to Matrix Reguler: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Auto-syncs completed CC/CW work to GD matrix
 * This allows GD team to see what content needs design work
 */
function autoSyncCompletedToGD(division, rowNo) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const divisionSheet = ss.getSheetByName(
      division === 'CC' ? SHEET_NAMES.MATRIX_CC : SHEET_NAMES.MATRIX_CW
    );
    
    if (!divisionSheet) return;
    
    // Get the completed row data from CC/CW matrix
    const data = divisionSheet.getDataRange().getValues();
    let rowData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowNo) {
        rowData = data[i];
        break;
      }
    }
    
    if (!rowData) return;
    
    // Sync to GD matrix
    const gdSheet = ss.getSheetByName(SHEET_NAMES.MATRIX_GD);
    if (!gdSheet) return;
    
    const existingData = gdSheet.getDataRange().getValues();
    
    // Check if row exists in GD
    let existingRowIndex = -1;
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === rowNo) {
        existingRowIndex = i;
        break;
      }
    }
    
    if (existingRowIndex === -1) {
      // Detect format of GD sheet
      const gdFormat = detectHeaderFormat(gdSheet);
      
      let newRow;
      if (gdFormat.format === 'database') {
        // Database format: No | Content Idea | Reference Link | Content Type | Upload Date | Upload Day | Upload Time | Content Creating Deadline | Content Result | PJ | Checked by REVIEWER | Status
        newRow = [
          rowNo,              // No
          rowData[1] || '',   // Content Idea
          rowData[3] || '',   // Reference Link
          rowData[4] || '',   // Content Type
          rowData[5] || '',   // Upload Date
          rowData[6] || '',   // Upload Day
          rowData[7] || '',   // Upload Time
          '',                 // Content Creating Deadline (empty, GD fills)
          '',                 // Content Result (empty, GD fills)
          '',                 // PJ (empty, GD fills)
          'Unreviewed yet',   // Checked by REVIEWER
          'Unreviewed yet'    // Status
        ];
      } else if (gdFormat.format === 'new') {
        // New format: NO | IDEA | SMS DIRECTION | REFERENCE | TYPE | UPLOAD DEADLINE | DAY | TIME | DESIGN DEADLINE | RESULT LINK | PIC | REVIEW
        newRow = [
          rowNo,              // NO
          rowData[1] || '',   // IDEA
          rowData[2] || '',   // SMS DIRECTION
          rowData[3] || '',   // REFERENCE
          rowData[4] || '',   // TYPE
          rowData[5] || '',   // UPLOAD DEADLINE
          rowData[6] || '',   // DAY
          rowData[7] || '',   // TIME
          '',                 // DESIGN DEADLINE (empty, GD fills)
          '',                 // RESULT LINK (empty, GD fills)
          '',                 // PIC (empty, GD fills)
          'Unreviewed yet'    // REVIEW
        ];
      } else {
        // Old format
        newRow = [
          rowNo,              // NO
          rowData[1] || '',   // Content Idea
          rowData[3] || '',   // Reference Link
          rowData[4] || '',   // Content Type
          rowData[5] || '',   // Upload Date
          rowData[6] || '',   // Upload Day
          rowData[7] || '',   // Upload Time
          '',                 // Deadline
          '',                 // Result
          '',                 // PIC
          'Unreviewed yet'    // Review
        ];
      }
      gdSheet.appendRow(newRow);
      Logger.log(`Auto-synced ${division} row ${rowNo} to GD matrix`);
    }
    // If row already exists in GD, don't overwrite their work
    
  } catch (error) {
    Logger.log('Error auto-syncing to GD: ' + error.toString());
    // Don't fail the main operation if sync fails
  }
}

/**
 * Deletes row from division matrix
 */
function deleteDivisionRow(ss, sheetName, rowNo) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rowNo) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
  } catch (error) {
    Logger.log('Error deleting division row: ' + error.toString());
  }
}

function deleteCCRow(rowNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  deleteDivisionRow(ss, SHEET_NAMES.MATRIX_CC, rowNo);
  return { success: true };
}

function deleteCWRow(rowNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  deleteDivisionRow(ss, SHEET_NAMES.MATRIX_CW, rowNo);
  return { success: true };
}

function deleteGDRow(rowNo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  deleteDivisionRow(ss, SHEET_NAMES.MATRIX_GD, rowNo);
  return { success: true };
}

/**
 * Creates division matrix sheet with headers matching application display
 */
function createDivisionMatrixSheet(ss, sheetName) {
  const sheet = ss.insertSheet(sheetName);
  
  // Headers matching application display format
  // Format: NO | IDEA | SMS DIRECTION | REFERENCE | TYPE | UPLOAD DEADLINE | DAY | TIME | [DEADLINE] | RESULT LINK | PIC | [REVIEWER] | REVIEW
  let headers;
  
  if (sheetName.includes('CW')) {
    // CW matrix includes Reviewer column
    headers = [
      'NO',
      'IDEA',
      'SMS DIRECTION',
      'REFERENCE',
      'TYPE',
      'UPLOAD DEADLINE',
      'DAY',
      'TIME',
      'BRIEF DEADLINE',
      'RESULT LINK',
      'PIC',
      'REVIEWER',
      'REVIEW'
    ];
  } else if (sheetName.includes('CC')) {
    // CC matrix
    headers = [
      'NO',
      'IDEA',
      'SMS DIRECTION',
      'REFERENCE',
      'TYPE',
      'UPLOAD DEADLINE',
      'DAY',
      'TIME',
      'VIDEO DEADLINE',
      'RESULT LINK',
      'PIC',
      'REVIEW'
    ];
  } else {
    // GD matrix
    headers = [
      'NO',
      'IDEA',
      'SMS DIRECTION',
      'REFERENCE',
      'TYPE',
      'UPLOAD DEADLINE',
      'DAY',
      'TIME',
      'DESIGN DEADLINE',
      'RESULT LINK',
      'PIC',
      'REVIEW'
    ];
  }
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Color based on division
  let headerColor = '#30678e';
  if (sheetName.includes('CC')) headerColor = '#ef6426';
  if (sheetName.includes('CW')) headerColor = '#30678e';
  if (sheetName.includes('GD')) headerColor = '#ef6426';
  
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(headerColor)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set optimized column widths
  sheet.setColumnWidth(1, 40);   // NO
  sheet.setColumnWidth(2, 200);  // IDEA
  sheet.setColumnWidth(3, 150);  // SMS DIRECTION
  sheet.setColumnWidth(4, 150);  // REFERENCE
  sheet.setColumnWidth(5, 110);  // TYPE
  sheet.setColumnWidth(6, 110);  // UPLOAD DEADLINE
  sheet.setColumnWidth(7, 90);   // DAY
  sheet.setColumnWidth(8, 90);   // TIME
  sheet.setColumnWidth(9, 140);  // [DEADLINE]
  sheet.setColumnWidth(10, 150); // RESULT LINK
  sheet.setColumnWidth(11, 100); // PIC
  if (sheetName.includes('CW')) {
    sheet.setColumnWidth(12, 100); // REVIEWER
    sheet.setColumnWidth(13, 120); // REVIEW
  } else {
    sheet.setColumnWidth(12, 120); // REVIEW
  }
  
  return sheet;
}

/**
 * Calculates progress for a division
 * Progress is based on DATE column in top table and REVIEW status in bottom table
 */
function calculateProgress(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: true, progress: 0 };
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, progress: 0 };
    }
    
    const format = detectHeaderFormat(sheet);
    
    // Map column indices
    let dateColIndex, reviewColIndex;
    if (format.format === 'new') {
      dateColIndex = 5; // UPLOAD DEADLINE column (index 5, 0-based)
      const isCW = sheetName === SHEET_NAMES.MATRIX_CW;
      reviewColIndex = isCW ? 12 : 11; // REVIEW column (12 for CW, 11 for CC/GD)
    } else if (format.format === 'database') {
      dateColIndex = 4; // Upload Date column
      reviewColIndex = 10; // Checked by REVIEWER column
    } else {
      dateColIndex = 4; // Upload Date column
      reviewColIndex = 10; // Review column
    }
    
    let total = 0;
    let completed = 0;
    
    for (let i = 1; i < data.length; i++) {
      // Check if row has a valid date (not empty, not just dashes)
      const dateValue = data[i][dateColIndex];
      const hasValidDate = dateValue && dateValue.toString().trim() !== '' && !dateValue.toString().trim().match(/^[-]+$/);
      
      if (hasValidDate) {
        total++;
        // Check if review status is "Reviewed"
        const reviewValue = (data[i][reviewColIndex] || '').toString().trim();
        if (reviewValue === 'Reviewed') {
          completed++;
        }
      }
    }
    
    const progress = total > 0 ? Math.round((completed / total) * 100) : 0;
    
    return { success: true, progress: progress };
    
  } catch (error) {
    Logger.log('Error calculating progress: ' + error.toString());
    return { success: false, progress: 0 };
  }
}

function getCCProgress() { return calculateProgress(SHEET_NAMES.MATRIX_CC); }
function getCWProgress() { return calculateProgress(SHEET_NAMES.MATRIX_CW); }
function getGDProgress() { return calculateProgress(SHEET_NAMES.MATRIX_GD); }

/**
 * Cleans up "No Idea" entries from GD matrix
 */
function cleanupGDNoIdeaEntries() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const gdSheet = ss.getSheetByName(SHEET_NAMES.MATRIX_GD);
    
    if (!gdSheet) {
      return { success: false, message: 'GD sheet not found' };
    }
    
    const data = gdSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, deleted: 0 };
    }
    
    const format = detectHeaderFormat(gdSheet);
    let ideaColIndex;
    if (format.format === 'database') {
      ideaColIndex = 1; // Content Idea column
    } else if (format.isNew) {
      ideaColIndex = 1; // IDEA column
    } else {
      ideaColIndex = 1; // Content Idea column
    }
    
    let deleted = 0;
    // Delete from bottom to top to avoid index shifting issues
    for (let i = data.length - 1; i >= 1; i--) {
      const idea = (data[i][ideaColIndex] || '').toString().trim().toLowerCase();
      if (idea.includes('no idea') && idea.replace(/[^a-z0-9]/gi, '').length < 10) {
        gdSheet.deleteRow(i + 1);
        deleted++;
      }
    }
    
    return { success: true, deleted: deleted };
    
  } catch (error) {
    Logger.log('Error cleaning up GD entries: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ========================================
// UTILITY FUNCTIONS
// ========================================

function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      // Check if it's already in dd/mm/yyyy format
      if (dateValue.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        return dateValue;
      }
      date = new Date(dateValue);
    } else {
      return '';
    }
    
    if (isNaN(date.getTime())) return '';
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    
    return `${day}/${month}/${year}`;
    
  } catch (error) {
    return '';
  }
}

function parseDateString(dateStr) {
  if (!dateStr) return null;
  
  try {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      // Expecting dd/mm/yyyy format
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const year = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }
    
    return null;
  } catch (error) {
    return null;
  }
}

function getDayName(date) {
  if (!date || !(date instanceof Date)) return '';
  
  const indonesianDays = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  return indonesianDays[date.getDay()];
}

function getConfigData() {
  return {
    success: true,
    data: {
      picLists: PIC_LISTS,
      reviewStatus: REVIEW_STATUS,
      contentTypes: CONTENT_TYPES,
      uploadTimes: UPLOAD_TIMES,
      monthFilters: MONTH_FILTERS,
      roles: {
        admin: ROLES.ADMIN,
        divisions: Object.values(ROLES.DIVISIONS)
      }
    }
  };
}

/**
 * Updates existing sheet headers to match application format
 */
function updateSheetHeaders(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const format = detectHeaderFormat(sheet);
    
    // Only update if not already in new format
    if (format.format === 'new') {
      return { success: true, message: 'Headers already in correct format' };
    }
    
    // Get new headers based on sheet type
    let newHeaders;
    if (sheetName === SHEET_NAMES.MATRIX_REGULER) {
      // Matrix Reguler headers are already correct
      return { success: true, message: 'Matrix Reguler headers are correct' };
    } else if (sheetName.includes('CW')) {
      newHeaders = [
        'NO', 'IDEA', 'SMS DIRECTION', 'REFERENCE', 'TYPE', 'UPLOAD DEADLINE', 'DAY', 'TIME',
        'BRIEF DEADLINE', 'RESULT LINK', 'PIC', 'REVIEWER', 'REVIEW'
      ];
    } else if (sheetName.includes('CC')) {
      newHeaders = [
        'NO', 'IDEA', 'SMS DIRECTION', 'REFERENCE', 'TYPE', 'UPLOAD DEADLINE', 'DAY', 'TIME',
        'VIDEO DEADLINE', 'RESULT LINK', 'PIC', 'REVIEW'
      ];
    } else if (sheetName.includes('GD')) {
      newHeaders = [
        'NO', 'IDEA', 'SMS DIRECTION', 'REFERENCE', 'TYPE', 'UPLOAD DEADLINE', 'DAY', 'TIME',
        'DESIGN DEADLINE', 'RESULT LINK', 'PIC', 'REVIEW'
      ];
    } else {
      return { success: false, message: 'Unknown sheet type' };
    }
    
    // Update header row
    const lastCol = sheet.getLastColumn();
    const headerRange = sheet.getRange(1, 1, 1, Math.max(lastCol, newHeaders.length));
    
    // Clear existing headers and set new ones
    headerRange.clear();
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    
    // Format header
    let headerColor = '#30678e';
    if (sheetName.includes('CC')) headerColor = '#ef6426';
    if (sheetName.includes('CW')) headerColor = '#30678e';
    if (sheetName.includes('GD')) headerColor = '#ef6426';
    
    sheet.getRange(1, 1, 1, newHeaders.length)
      .setBackground(headerColor)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    Logger.log('✓ Updated headers for: ' + sheetName);
    return { success: true, message: 'Headers updated successfully' };
    
  } catch (error) {
    Logger.log('Error updating headers: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Initialize all sheets and update headers if needed
 */
function initializeSheets() {
  Logger.log('Initializing All-Mark Dashboard V2.0...');
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('✓ Connected to spreadsheet');
    
    const sheets = [
      SHEET_NAMES.MATRIX_REGULER,
      SHEET_NAMES.MATRIX_CC,
      SHEET_NAMES.MATRIX_CW,
      SHEET_NAMES.MATRIX_GD
    ];
    
    sheets.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        if (sheetName === SHEET_NAMES.MATRIX_REGULER) {
          createMatrixRegulerSheet(ss);
        } else {
          createDivisionMatrixSheet(ss, sheetName);
        }
        Logger.log('✓ Created sheet: ' + sheetName);
      } else {
        // Update headers if needed
        updateSheetHeaders(sheetName);
        Logger.log('✓ Sheet exists: ' + sheetName);
      }
    });
    
    Logger.log('✓ Initialization complete!');
    return 'All sheets initialized successfully!';
    
  } catch (error) {
    Logger.log('✗ Error: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}
