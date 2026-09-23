// IMPROVED GOOGLE APPS SCRIPT CODE WITH STUDENT ID COLUMN

// 1. CONFIGURATION CONSTANTS
const CONFIG = {
  LOGO_FILE_ID: '1cDQOCGL74Y56XvKsX3cSkUrIL6DbE7dJ',
  STUDENT_DATA_SHEET: 'Student Data',
  TIMEZONE: 'America/New_York', // Update to your timezone
  DATE_FORMAT: 'MM/dd/yyyy',
  TIME_FORMAT: 'hh:mm a',
  HEADERS: ['Date', 'Time', 'Status', 'Student Name', 'Parent Name', 'Student ID', 'Returning', 'Logged', 'Bus'],
  STUDENT_DATA_HEADERS: {
    HOME_ROOM: 'Homeroom Section',
    FIRST: 'First',
    NICKNAME: 'Nickname',
    LAST: 'Last',
    FULL_NAME: 'First Last',
    STUDENT_ID: 'Student ID',
    CARPOOL_ID: 'Carpool ID',
    BUS_ROUTE: 'Bus Route',
    ACTIVE: 'Active'
  },
  MAX_BATCH_SIZE: 1000,
  CACHE_DURATION: 300 // 5 minutes
};

// 2. ERROR HANDLING
class AppError extends Error {
  constructor(message, code = 'UNKNOWN_ERROR', details = {}) {
    super(message);
    this.name = 'AppError';
    this.code = code;
    this.details = details;
    this.timestamp = new Date().toISOString();
  }
}

const ErrorCodes = {
  SHEET_NOT_FOUND: 'SHEET_NOT_FOUND',
  INVALID_DATA: 'INVALID_DATA',
  PERMISSION_DENIED: 'PERMISSION_DENIED',
  VALIDATION_ERROR: 'VALIDATION_ERROR'
};

// 3. LOGGING UTILITY
const Logger = {
  info(message, data = {}) {
    console.log(`[INFO] ${message}`, data);
  },
  
  error(message, error = {}) {
    console.error(`[ERROR] ${message}`, error);
  },
  
  warn(message, data = {}) {
    console.warn(`[WARN] ${message}`, data);
  }
};

// 4. VALIDATION UTILITIES
const Validator = {
  sanitize(value) {
    if (value === null || value === undefined) return '';
    // Preserve Unicode names/diacritics. Only remove control characters and trim.
    return String(value).replace(/[\u0000-\u001F\u007F]/g, '').trim();
  },

  safeCellText(value) {
    const cleaned = this.sanitize(value);
    // Prevent user-entered text from being interpreted as a spreadsheet formula.
    return /^[=+@]/.test(cleaned) ? "'" + cleaned : cleaned;
  },
  
  validateStudentName(name) {
    const sanitized = this.sanitize(name);
    if (!sanitized || sanitized.length < 2) {
      throw new AppError('Student name must be at least 2 characters', ErrorCodes.VALIDATION_ERROR);
    }
    if (sanitized.length > 100) {
      throw new AppError('Student name too long', ErrorCodes.VALIDATION_ERROR);
    }
    return sanitized;
  },
  
  validateParentName(name) {
    const sanitized = this.sanitize(name);
    if (!sanitized || sanitized.length < 2) {
      throw new AppError('Parent name must be at least 2 characters', ErrorCodes.VALIDATION_ERROR);
    }
    return sanitized;
  },
  
  validateStatus(status) {
    const validStatuses = ['In', 'Out'];
    if (!validStatuses.includes(status)) {
      throw new AppError('Invalid status', ErrorCodes.VALIDATION_ERROR);
    }
    return status;
  },
  
  validateSheetName(name) {
    const sanitized = this.sanitize(name);
    if (!/^\d{2}-\d{4}$/.test(sanitized)) {
      throw new AppError('Invalid sheet name format', ErrorCodes.VALIDATION_ERROR);
    }
    return sanitized;
  }
};

// 5. CACHING LAYER
const CacheManager = {
  cache: CacheService.getScriptCache(),
  
  get(key) {
    try {
      const cached = this.cache.get(key);
      return cached ? JSON.parse(cached) : null;
    } catch (error) {
      Logger.warn('Cache get error', { key, error: error.message });
      return null;
    }
  },
  
  set(key, value, expirationInSeconds = CONFIG.CACHE_DURATION) {
    try {
      this.cache.put(key, JSON.stringify(value), expirationInSeconds);
    } catch (error) {
      Logger.warn('Cache set error', { key, error: error.message });
    }
  },
  
  remove(key) {
    try {
      this.cache.remove(key);
    } catch (error) {
      Logger.warn('Cache remove error', { key, error: error.message });
    }
  },
  
  clear() {
    try {
      this.cache.removeAll([
        'all_students',
        `logo_${CONFIG.LOGO_FILE_ID}`
      ]);
    } catch (error) {
      Logger.warn('Cache clear error', { error: error.message });
      throw error;
    }
  }
};

// 6. SPREADSHEET UTILITIES
const SpreadsheetManager = {
  getActiveSpreadsheet() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        throw new AppError('No active spreadsheet found', ErrorCodes.PERMISSION_DENIED);
      }
      return ss;
    } catch (error) {
      Logger.error('Failed to get active spreadsheet', error);
      throw new AppError('Unable to access spreadsheet', ErrorCodes.PERMISSION_DENIED);
    }
  },
  
  getOrCreateSheet(name) {
  try {
    const ss = this.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      Logger.info('Creating new sheet', { name });
      sheet = ss.insertSheet(name);

      // Main headers A:I
      const headers = [
        'Date',
        'Time',
        'Status',
        'Student Name',
        'Parent Name',
        'Student ID',
        'Returning',
        'Logged',
        'Bus'
      ];

      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);

      // Format headers
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
      headerRange.setHorizontalAlignment('center');

      // Freeze header row
      sheet.setFrozenRows(1);

      // Set column widths
      sheet.setColumnWidth(1, 100); // Date
      sheet.setColumnWidth(2, 100); // Time
      sheet.setColumnWidth(3, 80);  // Status
      sheet.setColumnWidth(4, 200); // Student Name
      sheet.setColumnWidth(5, 200); // Parent Name
      sheet.setColumnWidth(6, 120); // Student ID
      sheet.setColumnWidth(7, 100); // Returning
      sheet.setColumnWidth(8, 100); // Logged
      sheet.setColumnWidth(9, 100); // Bus
    }

    return sheet;

  } catch (error) {
    Logger.error('Failed to get or create sheet', { name, error });
    throw new AppError(
      `Unable to access sheet: ${name}`,
      ErrorCodes.SHEET_NOT_FOUND
    );
  }
},
  
  batchWrite(sheet, data) {
    if (!data || !data.length) return;

    if (data.length > CONFIG.MAX_BATCH_SIZE) {
      throw new AppError('Batch size too large', ErrorCodes.INVALID_DATA);
    }

    const lock = LockService.getScriptLock();

    try {
      lock.waitLock(10000);

      const startRow = sheet.getLastRow() + 1;
      const numRows = data.length;
      const numCols = data[0].length;

      // Rows are written as A:F, including the authoritative Student ID.
      sheet.getRange(startRow, 1, numRows, numCols).setValues(data);

      const returningFormulas = [];
      const busFormulas = [];

      for (let i = 0; i < numRows; i++) {
        const rowNum = startRow + i;

        // G — Returning
        returningFormulas.push([
          `=IF(AND(C${rowNum}="In",COUNTIFS($C$1:C${rowNum - 1},"Out",$A$1:A${rowNum - 1},A${rowNum},$F$1:F${rowNum - 1},F${rowNum})>0),"x","")`
        ]);

        // I — Bus
        busFormulas.push([
          `=IF(C${rowNum}="Out",IFERROR(INDEX('${CONFIG.STUDENT_DATA_SHEET}'!H:H,MATCH(F${rowNum},'${CONFIG.STUDENT_DATA_SHEET}'!F:F,0)),""),"")`
        ]);
      }

      sheet.getRange(startRow, 7, numRows, 1).setFormulas(returningFormulas);
      sheet.getRange(startRow, 9, numRows, 1).setFormulas(busFormulas);

      Logger.info('Batch write successful', {
        startRow,
        numRows,
        numCols,
        sheetName: sheet.getName()
      });

    } catch (error) {
      Logger.error('Batch write failed', {
        error,
        dataLength: data.length
      });

      throw new AppError(
        'Failed to write data to sheet',
        ErrorCodes.UNKNOWN_ERROR,
        { cause: error.message }
      );
    } finally {
      try {
        lock.releaseLock();
      } catch (releaseError) {
        Logger.warn('Failed to release write lock', { error: releaseError.message });
      }
    }
  }
};

// 7. MAIN FUNCTIONS
function doGet(e) {
  try {
    const isManual = e && e.parameter && String(e.parameter.view || '').toLowerCase() === 'manual';
    const page = isManual ? 'manual-entry' : 'index';
    const title = isManual ? 'Manual Student Time Entry' : 'Check In/Out Students';

    Logger.info('Serving HTML UI', { page: page });

    const template = HtmlService.createTemplateFromFile(page);
    template.logoBytes = loadImageBytes();

    return template.evaluate()
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
      
  } catch (error) {
    Logger.error('doGet failed', error);
    // Return error page
    const errorTemplate = HtmlService.createTemplate(`
      <html>
        <body style="font-family: Arial, sans-serif; padding: 2rem; text-align: center;">
          <h1 style="color: #f44336;">Service Unavailable</h1>
          <p>The application is temporarily unavailable. Please try again later.</p>
          <p style="color: #666; font-size: 0.9rem;">Error: ${error.message}</p>
        </body>
      </html>
    `);
    return errorTemplate.evaluate();
  }
}

function loadImageBytes() {
  const cacheKey = `logo_${CONFIG.LOGO_FILE_ID}`;
  
  try {
    // Try cache first
    let logoBytes = CacheManager.get(cacheKey);
    if (logoBytes) {
      Logger.info('Logo loaded from cache');
      return logoBytes;
    }
    
    // Load from Drive
    Logger.info('Loading logo from Drive', { fileId: CONFIG.LOGO_FILE_ID });
    const blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    logoBytes = Utilities.base64Encode(blob.getBytes());
    
    // Cache for future use (1 hour)
    CacheManager.set(cacheKey, logoBytes, 3600);
    
    return logoBytes;
    
  } catch (error) {
    Logger.error('Failed to load logo', error);
    // Return empty string - app will work without logo
    return '';
  }
}

function getHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[String(header || '').trim().toLowerCase()] = index;
  });
  return map;
}

function getStudentRoster_(includeInactive) {
  const sheet = SpreadsheetManager.getActiveSpreadsheet().getSheetByName(CONFIG.STUDENT_DATA_SHEET);

  if (!sheet) {
    throw new AppError('Student Data sheet not found', ErrorCodes.SHEET_NOT_FOUND);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headerMap = getHeaderMap_(data[0]);
  const column = (name) => headerMap[String(name).toLowerCase()];

  [CONFIG.STUDENT_DATA_HEADERS.FULL_NAME, CONFIG.STUDENT_DATA_HEADERS.STUDENT_ID].forEach((header) => {
    if (column(header) === undefined) {
      throw new AppError(
        `Student Data is missing required column: ${header}`,
        ErrorCodes.INVALID_DATA
      );
    }
  });

  const activeCol = column(CONFIG.STUDENT_DATA_HEADERS.ACTIVE);
  const isActiveValue = (value) => {
    if (activeCol === undefined) return true;
    if (value === false || value === 0) return false;
    const normalized = String(value === undefined ? '' : value).trim().toLowerCase();
    return !['false', 'no', 'inactive', '0'].includes(normalized);
  };

  return data.slice(1)
    .filter(row => row[column(CONFIG.STUDENT_DATA_HEADERS.FULL_NAME)])
    .filter(row => includeInactive || isActiveValue(row[activeCol]))
    .map(row => ({
      homeRoom: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.HOME_ROOM)] || ''),
      firstName: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.FIRST)] || ''),
      nickname: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.NICKNAME)] || ''),
      lastName: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.LAST)] || ''),
      fullName: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.FULL_NAME)] || ''),
      studentId: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.STUDENT_ID)] || ''),
      carpoolId: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.CARPOOL_ID)] || ''),
      busRoute: Validator.sanitize(row[column(CONFIG.STUDENT_DATA_HEADERS.BUS_ROUTE)] || ''),
      active: activeCol === undefined ? true : isActiveValue(row[activeCol])
    }))
    .filter(student => student.fullName && student.studentId);
}

function resolveStudentSelections_(selections) {
  if (!Array.isArray(selections) || selections.length === 0) {
    throw new AppError('No students provided', ErrorCodes.VALIDATION_ERROR);
  }

  const roster = getStudentRoster_(true);
  const byId = {};
  const byName = {};

  roster.forEach(student => {
    byId[student.studentId] = student;
    byName[student.fullName.toLowerCase()] = student;
  });

  return selections.map(selection => {
    const rawSelection = Validator.sanitize(selection);
    const suppliedId = selection && typeof selection === 'object'
      ? Validator.sanitize(selection.studentId || '')
      : (/^\d+$/.test(rawSelection) ? rawSelection : '');

    const suppliedName = selection && typeof selection === 'object'
      ? Validator.sanitize(selection.fullName || '')
      : (suppliedId ? '' : rawSelection);

    if (suppliedId && byId[suppliedId]) {
      return byId[suppliedId];
    }

    if (suppliedName && byName[suppliedName.toLowerCase()]) {
      return byName[suppliedName.toLowerCase()];
    }

    // Preserve the manual-add fallback for students not yet present in Student Data.
    if (suppliedName) {
      return {
        fullName: Validator.validateStudentName(suppliedName),
        studentId: 'Not Found',
        active: true
      };
    }

    throw new AppError('Student selection could not be resolved', ErrorCodes.VALIDATION_ERROR);
  });
}

function getAllStudents() {
  const cacheKey = 'all_students';

  try {
    const cached = CacheManager.get(cacheKey);
    if (cached) {
      Logger.info('Students loaded from cache', { count: cached.length });
      return cached;
    }

    const students = getStudentRoster_(false)
      .sort((a, b) => a.fullName.localeCompare(b.fullName));

    Logger.info('Students loaded successfully', { count: students.length });
    CacheManager.set(cacheKey, students);
    return students;

  } catch (error) {
    Logger.error('getAllStudents failed', error);
    if (error instanceof AppError) throw error;
    throw new AppError('Unable to load student list', ErrorCodes.UNKNOWN_ERROR);
  }
}

function logCheckInOutToSheet(monthYear, status, parentName, studentSelections) {
  try {
    const validatedSheetName = Validator.validateSheetName(monthYear);
    const validatedStatus = Validator.validateStatus(status);
    const validatedParent = Validator.safeCellText(Validator.validateParentName(parentName));
    const validatedStudents = resolveStudentSelections_(studentSelections);

    Logger.info('Processing check-in/out', {
      sheet: validatedSheetName,
      status: validatedStatus,
      parent: validatedParent,
      studentCount: validatedStudents.length
    });

    const sheet = SpreadsheetManager.getOrCreateSheet(validatedSheetName);
    const now = new Date();
    const date = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
    const time = Utilities.formatDate(now, CONFIG.TIMEZONE, CONFIG.TIME_FORMAT);

    const rows = validatedStudents.map(student => [
      date,
      time,
      validatedStatus,
      student.fullName,
      validatedParent,
      student.studentId
    ]);

    SpreadsheetManager.batchWrite(sheet, rows);

    Logger.info('Check-in/out completed successfully', {
      studentsProcessed: validatedStudents.length,
      timestamp: now.toISOString()
    });

    return {
      success: true,
      count: validatedStudents.length
    };

  } catch (error) {
    Logger.error('logCheckInOutToSheet failed', error);
    if (error instanceof AppError) throw error;
    throw new AppError('Failed to record check-in/out', ErrorCodes.UNKNOWN_ERROR);
  }
}


/**
 * Records an administrator-selected group of students at a manually supplied
 * date and time. The HTML page sends date as YYYY-MM-DD and time as HH:MM.
 * Large groups are automatically written in chunks so the UI has no practical
 * student-count limit.
 */
function logManualCheckInOut(status, manualDate, manualTime, studentSelections) {
  try {
    const validatedStatus = Validator.validateStatus(status);

    if (!Array.isArray(studentSelections) || studentSelections.length === 0) {
      throw new AppError('No students provided', ErrorCodes.VALIDATION_ERROR);
    }

    const dateMatch = String(manualDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!dateMatch) {
      throw new AppError('Invalid date. Expected YYYY-MM-DD.', ErrorCodes.VALIDATION_ERROR);
    }

    const timeMatch = String(manualTime || '').match(/^(\d{2}):(\d{2})$/);
    if (!timeMatch) {
      throw new AppError('Invalid time. Expected HH:MM.', ErrorCodes.VALIDATION_ERROR);
    }

    const year = Number(dateMatch[1]);
    const month = Number(dateMatch[2]);
    const day = Number(dateMatch[3]);
    const hour24 = Number(timeMatch[1]);
    const minute = Number(timeMatch[2]);

    if (month < 1 || month > 12 || day < 1 || day > 31 ||
        hour24 < 0 || hour24 > 23 || minute < 0 || minute > 59) {
      throw new AppError('Invalid manual date or time.', ErrorCodes.VALIDATION_ERROR);
    }

    // Reject impossible calendar dates such as 2026-02-31.
    const dateCheck = new Date(Date.UTC(year, month - 1, day));
    if (dateCheck.getUTCFullYear() !== year ||
        dateCheck.getUTCMonth() !== month - 1 ||
        dateCheck.getUTCDate() !== day) {
      throw new AppError('Invalid calendar date.', ErrorCodes.VALIDATION_ERROR);
    }

    const validatedStudents = resolveStudentSelections_(studentSelections);
    const sheetName = String(month).padStart(2, '0') + '-' + year;
    const validatedSheetName = Validator.validateSheetName(sheetName);

    const formattedDate = String(month).padStart(2, '0') + '/' +
      String(day).padStart(2, '0') + '/' + year;

    const hour12 = hour24 % 12 || 12;
    const amPm = hour24 >= 12 ? 'PM' : 'AM';
    const formattedTime = String(hour12).padStart(2, '0') + ':' +
      String(minute).padStart(2, '0') + ' ' + amPm;

    const sheet = SpreadsheetManager.getOrCreateSheet(validatedSheetName);
    const manualSource = 'Manual Entry';

    const rows = validatedStudents.map(student => [
      formattedDate,
      formattedTime,
      validatedStatus,
      student.fullName,
      manualSource,
      student.studentId
    ]);

    // SpreadsheetManager protects individual writes at MAX_BATCH_SIZE.
    // Chunk here so the manual page can submit any size group.
    for (let start = 0; start < rows.length; start += CONFIG.MAX_BATCH_SIZE) {
      SpreadsheetManager.batchWrite(
        sheet,
        rows.slice(start, start + CONFIG.MAX_BATCH_SIZE)
      );
    }

    Logger.info('Manual check-in/out completed successfully', {
      sheet: validatedSheetName,
      status: validatedStatus,
      date: formattedDate,
      time: formattedTime,
      studentCount: validatedStudents.length
    });

    return {
      success: true,
      count: validatedStudents.length,
      status: validatedStatus,
      date: formattedDate,
      time: formattedTime
    };

  } catch (error) {
    Logger.error('logManualCheckInOut failed', error);

    if (error instanceof AppError) {
      throw error;
    }

    throw new AppError('Failed to record manual check-in/out', ErrorCodes.UNKNOWN_ERROR);
  }
}

// 8. UTILITY FUNCTIONS FOR MAINTENANCE
function clearCache_() {
  try {
    CacheManager.clear();
    Logger.info('Cache cleared successfully');
    return 'Cache cleared';
  } catch (error) {
    Logger.error('Failed to clear cache', error);
    throw new AppError('Failed to clear cache', ErrorCodes.UNKNOWN_ERROR);
  }
}

function getSystemInfo_() {
  try {
    const ss = SpreadsheetManager.getActiveSpreadsheet();
    return {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      timezone: CONFIG.TIMEZONE,
      userEmail: Session.getActiveUser().getEmail(),
      timestamp: new Date().toISOString(),
      version: '2.1.0' // Updated version
    };
  } catch (error) {
    Logger.error('Failed to get system info', error);
    throw new AppError('Failed to get system information', ErrorCodes.UNKNOWN_ERROR);
  }
}

// 9. PERFORMANCE MONITORING FUNCTION
function getPerformanceMetrics_() {
  try {
    const ss = SpreadsheetManager.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    const metrics = {
      totalSheets: sheets.length,
      totalStudents: 0,
      recentActivity: {},
      cacheStatus: 'enabled'
    };
    
    // Get student count
    try {
      const students = getStudentRoster_(false);
      metrics.totalStudents = students.length;
    } catch (error) {
      metrics.totalStudents = 'error';
    }
    
    // Get recent check-ins/outs
    const currentMonth = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'MM-yyyy');
    const currentSheet = ss.getSheetByName(currentMonth);
    
    if (currentSheet) {
      const data = currentSheet.getDataRange().getValues();
      metrics.recentActivity = {
        totalRecords: data.length - 1, // Exclude header
        sheetName: currentMonth
      };
    }
    
    return metrics;
  } catch (error) {
    Logger.error('Failed to get performance metrics', error);
    return { error: error.message };
  }
}

// 10. HELPER FUNCTION TO TEST STUDENT ID LOOKUP
function testStudentIdLookup_() {
  try {
    const ss = SpreadsheetManager.getActiveSpreadsheet();
    const studentDataSheet = ss.getSheetByName(CONFIG.STUDENT_DATA_SHEET);
    
    if (!studentDataSheet) {
      Logger.error('Student Data sheet not found');
      return 'Student Data sheet not found';
    }
    
    const data = studentDataSheet.getDataRange().getValues();
    Logger.info('Student Data structure preview:', {
      headers: data[0],
      sampleRow: data[1],
      totalRows: data.length
    });
    
    return {
      message: 'Student Data sheet structure logged to console',
      headers: data[0],
      totalStudents: data.length - 1
    };
    
  } catch (error) {
    Logger.error('Failed to test student ID lookup', error);
    return { error: error.message };
  }
}