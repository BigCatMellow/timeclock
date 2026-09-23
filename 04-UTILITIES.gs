// ============================================================================
// 04-UTILITIES.GS
// Shared utility functions (date/time, student data, formatting)
// ============================================================================

// DATE FORMATTING
function formatDate(date) {
  if (!date) return '';
  
  // If it's already a string in MM/DD/YYYY format, return it
  if (typeof date === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(date)) {
    return date;
  }
  
  // If it's a Date object, format it
  if (date instanceof Date && !isNaN(date.getTime())) {
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    return `${month}/${day}/${year}`;
  }
  
  // Try to parse the date if it's a string
  if (typeof date === 'string') {
    const parsedDate = new Date(date);
    if (!isNaN(parsedDate.getTime())) {
      const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
      const day = String(parsedDate.getDate()).padStart(2, '0');
      const year = parsedDate.getFullYear();
      return `${month}/${day}/${year}`;
    }
  }
  
  // If we can't format it, return it as a string
  return String(date);
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  let date;
  if (typeof dateStr === 'string') {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
      const month = parseInt(parts[0]) - 1;
      const day = parseInt(parts[1]);
      const year = parseInt(parts[2]);
      date = new Date(year, month, day);
    } else {
      date = new Date(dateStr);
    }
  } else if (dateStr instanceof Date) {
    date = new Date(dateStr);
  } else {
    return null;
  }
  return isNaN(date.getTime()) ? null : date;
}

function getMonthsBetweenDates(startDate, endDate) {
  const months = [];
  const current = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

  while (current <= end) {
    const month = String(current.getMonth() + 1).padStart(2, '0');
    const year = current.getFullYear();
    months.push(`${month}-${year}`);
    current.setMonth(current.getMonth() + 1);
  }

  return months;
}

function getRelevantMonthSheets(allSheets, startDate, endDate) {
  const monthsToCheck = getMonthsBetweenDates(startDate, endDate);
  return allSheets.filter(sheet => {
    const sheetName = sheet.getName();
    return /^\d{2}-\d{4}$/.test(sheetName) && monthsToCheck.includes(sheetName);
  });
}

// TIME FORMATTING AND CONVERSION
function formatTimeForDisplay(timeStr) {
  if (!timeStr) return '';
  
  if (typeof timeStr === 'string' && /^\d{1,2}:\d{2}\s*(AM|PM|am|pm)?$/i.test(timeStr.trim())) {
    return timeStr.trim();
  }
  
  if (timeStr instanceof Date) {
    return timeStr.toLocaleTimeString('en-US', { 
      hour: 'numeric', 
      minute: '2-digit',
      hour12: true 
    });
  }
  
  if (typeof timeStr === 'string' && timeStr.includes('/')) {
    try {
      const date = new Date(timeStr);
      if (!isNaN(date.getTime())) {
        return date.toLocaleTimeString('en-US', { 
          hour: 'numeric', 
          minute: '2-digit',
          hour12: true 
        });
      }
    } catch (e) {
      // Continue to manual extraction
    }
  }
  
  try {
    let cleanTime = String(timeStr).trim();
    
    if (cleanTime.includes(' ')) {
      const parts = cleanTime.split(' ');
      for (const part of parts) {
        if (/^\d{1,2}:\d{2}(:\d{2})?$/i.test(part)) {
          cleanTime = part;
          break;
        }
      }
    }
    
    cleanTime = cleanTime.replace(/:\d{2}$/, '');
    
    const [hours, minutes] = cleanTime.split(':').map(num => parseInt(num) || 0);
    const date = new Date();
    date.setHours(hours, minutes, 0, 0);
    
    return date.toLocaleTimeString('en-US', { 
      hour: 'numeric', 
      minute: '2-digit',
      hour12: true 
    });
  } catch (error) {
    return String(timeStr);
  }
}

function convertTimeForSort(timeStr) {
  if (!timeStr) return '00:00';

  try {
    if (timeStr instanceof Date) {
      return `${String(timeStr.getHours()).padStart(2, '0')}:${String(timeStr.getMinutes()).padStart(2, '0')}`;
    }

    let time = String(timeStr).toLowerCase().trim();
    const isPM = time.includes('pm');
    time = time.replace(/[ap]m/gi, '').trim();
    
    const [hours, minutes] = time.split(':').map(part => parseInt(part) || 0);
    let hour24 = hours;
    
    if (isPM && hours !== 12) {
      hour24 += 12;
    } else if (!isPM && hours === 12) {
      hour24 = 0;
    }
    
    return `${String(hour24).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  } catch (error) {
    return '00:00';
  }
}

function convertTimeToMinutes(timeStr) {
  try {
    if (!timeStr) return null;
    
    let cleanTime = String(timeStr).trim();
    
    // Handle Date objects directly.
    if (timeStr instanceof Date) {
      return timeStr.getHours() * 60 + timeStr.getMinutes();
    }
    
    // Handle 24-hour format (HH:MM)
    if (/^\d{1,2}:\d{2}$/.test(cleanTime)) {
      const [hours, minutes] = cleanTime.split(':').map(num => parseInt(num));
      return hours * 60 + minutes;
    }
    
    // Handle 12-hour format (H:MM AM/PM)
    const match = cleanTime.match(/^(\d{1,2}):(\d{2})\s*(AM|PM|am|pm)?$/i);
    if (match) {
      let hours = parseInt(match[1]);
      const minutes = parseInt(match[2]);
      const ampm = match[3] ? match[3].toUpperCase() : '';
      
      if (ampm === 'PM' && hours !== 12) {
        hours += 12;
      } else if (ampm === 'AM' && hours === 12) {
        hours = 0;
      }
      
      return hours * 60 + minutes;
    }
    
    // Extract time from longer strings
    const timeMatch = cleanTime.match(/(\d{1,2}):(\d{2})/);
    if (timeMatch) {
      const hours = parseInt(timeMatch[1]);
      const minutes = parseInt(timeMatch[2]);
      return hours * 60 + minutes;
    }
    
    return null;
    
  } catch (error) {
    console.log('Error converting time to minutes:', timeStr, error);
    return null;
  }
}

// STUDENT DATA ACCESS
function getStudentGrades_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dataSheet = ss.getSheetByName('Student Data');
    
    if (!dataSheet) {
      console.log('Student Data sheet not found - grades will show as Unknown');
      return {};
    }
    
    const data = dataSheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('Student Data sheet is empty - grades will show as Unknown');
      return {};
    }
    
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const studentIdCol = headers.findIndex(h => h.includes('student') && h.includes('id') || h === 'id');
    const homeroomCol = headers.findIndex(h => h.includes('homeroom') && h.includes('section'));
    
    if (studentIdCol === -1) {
      console.log('Student ID column not found in Student Data sheet - grades will show as Unknown');
      return {};
    }
    
    if (homeroomCol === -1) {
      console.log('Homeroom Section column not found in Student Data sheet - grades will show as Unknown');
      return {};
    }
    
    const grades = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const studentId = String(row[studentIdCol]).trim();
      const homeroom = String(row[homeroomCol]).trim();
      
      if (studentId && homeroom) {
        grades[studentId] = homeroom;
      }
    }
    
    console.log(`Successfully loaded grades for ${Object.keys(grades).length} students`);
    return grades;
    
  } catch (error) {
    console.log('Error loading student grades:', error.message);
    console.log('Grades will show as Unknown');
    return {};
  }
}

function getStudentDirectory_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const directorySheet = ss.getSheetByName('parent db');

  if (!directorySheet) {
    console.log('parent db sheet not found');
    return {};
  }

  const data = directorySheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  const headers = data[0].map(h => String(h || '').toLowerCase().trim());
  const findCol = (...names) => {
    const wanted = names.map(name => String(name).toLowerCase().trim());
    return headers.findIndex(header => wanted.includes(header));
  };

  const cols = {
    studentId: findCol('Student ID'),
    homeRoom: findCol('Home Room Class', 'Homeroom Section'),
    firstName: findCol('First'),
    lastName: findCol('Last'),
    parentId: findCol('Parent ID'),
    combinedSalutations: findCol('Combined Salutations'),
    primaryContactName: findCol('Primary Contact: Name'),
    primaryContactEmail: findCol('Primary Contact: Email'),
    primaryContactCell: findCol('Primary Contact: Cell Phone'),
    secondaryContactName: findCol('Secondary Contact: Name'),
    secondaryContactEmail: findCol('Secondary Contact: Email'),
    secondaryContactCell: findCol('Secondary Contact: Cell Phone')
  };

  const requiredColumns = [
    ['Student ID', cols.studentId],
    ['Parent ID', cols.parentId],
    ['First', cols.firstName],
    ['Last', cols.lastName],
    ['Home Room Class', cols.homeRoom],
    ['Combined Salutations', cols.combinedSalutations],
    ['Primary Contact: Email', cols.primaryContactEmail],
    ['Secondary Contact: Email', cols.secondaryContactEmail]
  ];
  const missingColumns = requiredColumns.filter(([, index]) => index === -1).map(([name]) => name);
  if (missingColumns.length > 0) {
    throw new Error('parent db is missing required columns: ' + missingColumns.join(', '));
  }

  const valueAt = (row, index) => index === -1 ? '' : String(row[index] || '').trim();
  const directory = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const studentId = valueAt(row, cols.studentId);
    if (!studentId) continue;

    const firstName = valueAt(row, cols.firstName);
    const lastName = valueAt(row, cols.lastName);

    directory[studentId] = {
      studentId: studentId,
      parentId: valueAt(row, cols.parentId),
      firstName: firstName,
      lastName: lastName,
      fullName: [firstName, lastName].filter(Boolean).join(' '),
      grade: valueAt(row, cols.homeRoom),
      combinedSalutations: valueAt(row, cols.combinedSalutations),
      primaryContactName: valueAt(row, cols.primaryContactName),
      primaryContactEmail: valueAt(row, cols.primaryContactEmail),
      primaryContactCell: valueAt(row, cols.primaryContactCell),
      secondaryContactName: valueAt(row, cols.secondaryContactName),
      secondaryContactCell: valueAt(row, cols.secondaryContactCell),
      secondaryContactEmail: valueAt(row, cols.secondaryContactEmail)
    };
  }

  return directory;
}
