// IMPROVED GOOGLE APPS SCRIPT CODE WITH STUDENT ID COLUMN

// 1. CONFIGURATION CONSTANTS
const CONFIG = {
  LOGO_FILE_ID: '1cDQOCGL74Y56XvKsX3cSkUrIL6DbE7dJ',
  STUDENT_DATA_SHEET: 'Student Data',
  TIMEZONE: 'America/New_York', // Update to your timezone
  DATE_FORMAT: 'MM/dd/yyyy',
  TIME_FORMAT: 'hh:mm a',
  HEADERS: ['Date', 'Time', 'Status', 'Student Name', 'Parent Name', 'Student ID'], // Added Student ID
  SANITIZE_REGEX: /[^\w\s'-]/g,
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
  sanitize(text) {
    if (typeof text !== 'string') return '';
    return text.replace(CONFIG.SANITIZE_REGEX, '').trim();
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
      this.cache.removeAll();
    } catch (error) {
      Logger.warn('Cache clear error', { error: error.message });
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
        
        // Add headers
        const headerRange = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
        headerRange.setValues([CONFIG.HEADERS]);
        
        // Format headers
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f0f0f0');
        
        // Freeze header row
        sheet.setFrozenRows(1);
        
        // Auto-resize columns
        sheet.autoResizeColumns(1, CONFIG.HEADERS.length);
        
        // Set column widths for better display
        sheet.setColumnWidth(1, 100); // Date
        sheet.setColumnWidth(2, 100); // Time
        sheet.setColumnWidth(3, 80);  // Status
        sheet.setColumnWidth(4, 200); // Student Name
        sheet.setColumnWidth(5, 200); // Parent Name
        sheet.setColumnWidth(6, 120); // Student ID
      }
      
      return sheet;
    } catch (error) {
      Logger.error('Failed to get or create sheet', { name, error });
      throw new AppError(`Unable to access sheet: ${name}`, ErrorCodes.SHEET_NOT_FOUND);
    }
  },
  
  batchWrite(sheet, data) {
    if (!data || !data.length) return;
    
    try {
      // Validate batch size
      if (data.length > CONFIG.MAX_BATCH_SIZE) {
        throw new AppError('Batch size too large', ErrorCodes.INVALID_DATA);
      }
      
      const startRow = sheet.getLastRow() + 1;
      const numRows = data.length;
      const numCols = data[0].length;
      
      // Write the main data (Date, Time, Status, Student Name, Parent Name)
      const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
      dataRange.setValues(data);
      
      // Add Student ID formulas in column F (6th column)
      if (numCols >= 5) { // Make sure we have at least 5 columns before adding formulas
        const formulaRange = sheet.getRange(startRow, 6, numRows, 1); // Column F for Student ID
        const formulas = [];
        
        for (let i = 0; i < numRows; i++) {
          const rowNum = startRow + i;
          // Formula to lookup Student ID from Student Data sheet
          // Looks up the student name in column D (Student Name) against column E in Student Data sheet
          // and returns the corresponding value from column F in Student Data sheet
          const formula = `=IFERROR(INDEX('${CONFIG.STUDENT_DATA_SHEET}'!F:F,MATCH(D${rowNum},'${CONFIG.STUDENT_DATA_SHEET}'!E:E,0)),"Not Found")`;
          formulas.push([formula]);
        }
        
        formulaRange.setFormulas(formulas);
      }
      
      Logger.info('Batch write successful', { 
        startRow, 
        numRows, 
        numCols,
        sheetName: sheet.getName()
      });
      
    } catch (error) {
      Logger.error('Batch write failed', { error, dataLength: data.length });
      throw new AppError('Failed to write data to sheet', ErrorCodes.UNKNOWN_ERROR);
    }
  }
};

// 7. MAIN FUNCTIONS
function doGet() {
  try {
    Logger.info('Serving HTML UI');
    
    const template = HtmlService.createTemplateFromFile('index');
    template.logoBytes = loadImageBytes();
    
    return template.evaluate()
      .setTitle('Check In/Out Students')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
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

function getAllStudents() {
  const cacheKey = 'all_students';
  
  try {
    // Try cache first
    let students = CacheManager.get(cacheKey);
    if (students) {
      Logger.info('Students loaded from cache', { count: students.length });
      return students;
    }
    
    // Load from sheet
    Logger.info('Loading students from sheet');
    const sheet = SpreadsheetManager.getActiveSpreadsheet().getSheetByName(CONFIG.STUDENT_DATA_SHEET);
    
    if (!sheet) {
      throw new AppError('Student Data sheet not found', ErrorCodes.SHEET_NOT_FOUND);
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.warn('No student data found');
      return [];
    }
    
    // Skip header row and process data
    students = data.slice(1)
      .filter(row => row[4]) // Must have full name
      .map(row => ({
        homeRoom: Validator.sanitize(row[0] || ''),
        firstName: Validator.sanitize(row[1] || ''),
        nickname: Validator.sanitize(row[2] || ''), // ADDED: Column C (index 2)
        lastName: Validator.sanitize(row[3] || ''),
        fullName: Validator.sanitize(row[4] || ''),
        studentId: Validator.sanitize(row[5] || '')
      }))
      .filter(student => student.fullName.length > 0)
      .sort((a, b) => a.fullName.localeCompare(b.fullName));
    
    Logger.info('Students loaded successfully', { count: students.length });
    
    // Cache results
    CacheManager.set(cacheKey, students);
    
    return students;
    
  } catch (error) {
    Logger.error('getAllStudents failed', error);
    
    if (error instanceof AppError) {
      throw error;
    }
    
    throw new AppError('Unable to load student list', ErrorCodes.UNKNOWN_ERROR);
  }
}

function logCheckInOutToSheet(monthYear, status, parentName, studentNames) {
  try {
    // Validation
    const validatedSheetName = Validator.validateSheetName(monthYear);
    const validatedStatus = Validator.validateStatus(status);
    const validatedParent = Validator.validateParentName(parentName);
    
    if (!Array.isArray(studentNames) || studentNames.length === 0) {
      throw new AppError('No students provided', ErrorCodes.VALIDATION_ERROR);
    }
    
    const validatedStudents = studentNames.map(name => Validator.validateStudentName(name));
    
    Logger.info('Processing check-in/out', {
      sheet: validatedSheetName,
      status: validatedStatus,
      parent: validatedParent,
      studentCount: validatedStudents.length
    });
    
    // Get or create sheet
    const sheet = SpreadsheetManager.getOrCreateSheet(validatedSheetName);
    
    // Prepare data
    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const date = Utilities.formatDate(now, tz, CONFIG.DATE_FORMAT);
    const time = Utilities.formatDate(now, tz, CONFIG.TIME_FORMAT);
    
    // Prepare rows with 5 columns (Student ID will be added by formula)
    const rows = validatedStudents.map(studentName => [
      date,
      time,
      validatedStatus,
      studentName,
      validatedParent
    ]);
    
    // Write to sheet (the batchWrite function will handle adding the Student ID formulas)
    SpreadsheetManager.batchWrite(sheet, rows);
    
    // Clear relevant caches
    CacheManager.remove('all_students');
    
    Logger.info('Check-in/out completed successfully', {
      studentsProcessed: validatedStudents.length,
      timestamp: now.toISOString()
    });
    
  } catch (error) {
    Logger.error('logCheckInOutToSheet failed', error);
    
    if (error instanceof AppError) {
      throw error;
    }
    
    throw new AppError('Failed to record check-in/out', ErrorCodes.UNKNOWN_ERROR);
  }
}

// 8. UTILITY FUNCTIONS FOR MAINTENANCE
function clearCache() {
  try {
    CacheManager.clear();
    Logger.info('Cache cleared successfully');
    return 'Cache cleared';
  } catch (error) {
    Logger.error('Failed to clear cache', error);
    throw new AppError('Failed to clear cache', ErrorCodes.UNKNOWN_ERROR);
  }
}

function getSystemInfo() {
  try {
    const ss = SpreadsheetManager.getActiveSpreadsheet();
    return {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      timezone: Session.getScriptTimeZone(),
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
function getPerformanceMetrics() {
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
      const students = getAllStudents();
      metrics.totalStudents = students.length;
    } catch (error) {
      metrics.totalStudents = 'error';
    }
    
    // Get recent check-ins/outs
    const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-yyyy');
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
function testStudentIdLookup() {
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
