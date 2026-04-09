// ============================================================================
// 02-REPORT-GENERATION.GS
// All report generation logic (tardy reports, student history, parsing)
// ============================================================================

// CALENDAR FORM PROCESSOR - WITH TIME FILTERING
function generateTardyReportFromCalendar(formData) {
  try {
    const startDate = new Date(formData.startDate + 'T00:00:00');
    const endDate = new Date(formData.endDate + 'T23:59:59');
    
    if (!startDate || !endDate || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return { success: false, error: 'Invalid date format received from calendar' };
    }
    
    if (startDate > endDate) {
      return { success: false, error: 'Start date cannot be after end date' };
    }
    
    // IMPORTANT: Always use the time-filtered version
    const result = generateTardyReportWithTimeFilter(
      startDate, 
      endDate, 
      formData.studentId || '', 
      formData.minOccurrences || 1,
      formData.startTime,
      formData.endTime
    );
    
    return result;
    
  } catch (error) {
    return { success: false, error: 'Failed to process calendar selection: ' + error.message };
  }
}

// MAIN TARDY REPORT GENERATION WITH TIME FILTERING
function generateTardyReportWithTimeFilter(startDate, endDate, searchId, minOccurrences, startTime, endTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const tardyInstances = [];

  const relevantSheets = getRelevantMonthSheets(allSheets, startDate, endDate);
  if (relevantSheets.length === 0) {
    return { success: false, error: 'No relevant month sheets found for the specified date range.' };
  }

  relevantSheets.forEach(sheet => {
    const tardies = processSheetForTardiesWithTimeFilter(
      sheet, 
      startDate, 
      endDate, 
      searchId, 
      startTime, 
      endTime
    );
    tardyInstances.push(...tardies);
  });

  const groupedById = groupByIdAndFilter(tardyInstances, minOccurrences);
  const studentCount = Object.keys(groupedById).length;

  const sheetName = createGroupedTardyReportSheetWithTimeFilter(
    groupedById, 
    startDate, 
    endDate, 
    startTime, 
    endTime
  );

  return {
    success: true,
    sheetName: sheetName,
    tardyCount: tardyInstances.length,
    studentCount: studentCount
  };
}

// PROCESS SHEETS WITH TIME FILTERING
function processSheetForTardiesWithTimeFilter(sheet, startDate, endDate, searchId, startTime, endTime) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const studentGrades = getStudentGrades();
  const dailyEntries = {};
  const tardies = [];

  for (let i = 1; i < data.length; i++) {
    const [dateStr, timeStr, status, studentName, parentName, studentId] = data[i];

    if (!dateStr || !status || !studentName || !['In', 'Out'].includes(status)) {
      continue;
    }

    const rowDate = parseDate(dateStr);
    if (!rowDate || rowDate < startDate || rowDate > endDate) {
      continue;
    }

    if (searchId && studentId != searchId) continue;

    // Time filtering
    if (!isTimeInRange(timeStr, startTime, endTime)) {
      continue;
    }

    const dateKey = formatDate(rowDate);
    const studentKey = `${studentName}_${dateKey}`;

    if (!dailyEntries[studentKey]) {
      dailyEntries[studentKey] = [];
    }

    dailyEntries[studentKey].push({
      time: timeStr,
      status: status,
      parentName: parentName,
      studentName: studentName,
      studentId: studentId || 'Not Found',
      grade: studentGrades[studentId] || 'Unknown',
      date: rowDate
    });
  }

  for (const studentKey in dailyEntries) {
    const entries = dailyEntries[studentKey];
    
    entries.sort((a, b) => {
      return convertTimeForSort(a.time).localeCompare(convertTimeForSort(b.time));
    });

    if (entries[0].status === 'In') {
      tardies.push(entries[0]);
    }
  }

  return tardies;
}

// TIME RANGE VALIDATION FUNCTION
function isTimeInRange(timeStr, startTime, endTime) {
  try {
    const recordTime = convertTimeToMinutes(timeStr);
    const rangeStart = convertTimeToMinutes(startTime);
    const rangeEnd = convertTimeToMinutes(endTime);
    
    // Debug logging
    console.log(`Checking time: ${timeStr} (${recordTime} mins) against range ${startTime} (${rangeStart} mins) to ${endTime} (${rangeEnd} mins)`);
    
    if (recordTime === null || rangeStart === null || rangeEnd === null) {
      console.log(`  -> EXCLUDED (null values detected)`);
      return false;
    }
    
    const inRange = recordTime >= rangeStart && recordTime <= rangeEnd;
    console.log(`  -> ${inRange ? 'INCLUDED' : 'EXCLUDED'}`);
    
    return inRange;
    
  } catch (error) {
    console.log('Error checking time range for:', timeStr, error);
    return false;
  }
}

// CREATE REPORT SHEET WITH TIME FILTER INFO
function createGroupedTardyReportSheetWithTimeFilter(groupedById, startDate, endDate, startTime, endTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateRangeStr = `${formatDate(startDate)} to ${formatDate(endDate)}`;
  const timeRangeStr = `${startTime}-${endTime}`;
  const sheetName = `Late Arrivals Report - ${dateRangeStr} (${timeRangeStr})`;

  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) ss.deleteSheet(existingSheet);

  const reportSheet = ss.insertSheet(sheetName);

  let row = 1;
  
  reportSheet.getRange(row, 1).setValue(`Late Arrivals Report: ${dateRangeStr}`);
  reportSheet.getRange(row, 1).setFontWeight('bold').setFontSize(14);
  row++;
  
  reportSheet.getRange(row, 1).setValue(`Time Frame: ${startTime} to ${endTime}`);
  reportSheet.getRange(row, 1).setFontWeight('bold').setFontColor('#1a73e8');
  row += 2;

  const totalStudents = Object.keys(groupedById).length;
  const totalInstances = Object.values(groupedById).reduce((sum, group) => sum + group.entries.length, 0);
  
  reportSheet.getRange(row, 1).setValue(`Students with tardies: ${totalStudents}`);
  reportSheet.getRange(row + 1, 1).setValue(`Total Late Arrivals instances: ${totalInstances}`);
  reportSheet.getRange(row, 1, 2, 1).setFontWeight('bold');
  row += 3;

  const headers = ['Student ID', 'Student Name', 'Grade', 'Date', 'Time', 'Parent Name'];
  reportSheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  reportSheet.getRange(row, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  row++;

  let studentIndex = 0;
  for (const studentId in groupedById) {
    const group = groupedById[studentId];
    const isEvenStudent = studentIndex % 2 === 0;
    const backgroundColor = isEvenStudent ? '#ffffff' : '#f8f9fa';

    if (group.entries.length > 0) {
      const firstEntry = group.entries[0];
      reportSheet.getRange(row, 1, 1, headers.length).setValues([[
        studentId,
        group.studentName,
        group.grade || 'Unknown',
        formatDate(firstEntry.date),
        formatTimeForDisplay(firstEntry.time),
        firstEntry.parentName
      ]]);
      reportSheet.getRange(row, 1, 1, headers.length).setBackground(backgroundColor);
      row++;

      for (let i = 1; i < group.entries.length; i++) {
        const entry = group.entries[i];
        reportSheet.getRange(row, 1, 1, headers.length).setValues([[
          '',
          '',
          '',
          formatDate(entry.date),
          formatTimeForDisplay(entry.time),
          entry.parentName
        ]]);
        reportSheet.getRange(row, 1, 1, headers.length).setBackground(backgroundColor);
        row++;
      }
    }

    if (studentIndex < Object.keys(groupedById).length - 1) {
      row++;
    }

    studentIndex++;
  }

  reportSheet.setFrozenRows(7);
  reportSheet.autoResizeColumns(1, headers.length);
  
  reportSheet.setColumnWidth(1, 100);
  reportSheet.setColumnWidth(2, 200);
  reportSheet.setColumnWidth(3, 80);
  reportSheet.setColumnWidth(4, 100);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 200);

  const timeColumn = reportSheet.getRange(8, 4, row - 7, 1);
  timeColumn.setNumberFormat('@');

  return sheetName;
}

// GROUP AND FILTER FUNCTIONS
function groupByIdAndFilter(tardyInstances, minOccurrences) {
  minOccurrences = parseInt(minOccurrences, 10) || 1;
  const grouped = {};

  tardyInstances.forEach(tardy => {
    if (!grouped[tardy.studentId]) {
      grouped[tardy.studentId] = {
        studentName: tardy.studentName,
        grade: tardy.grade || 'Unknown',
        entries: []
      };
    }
    grouped[tardy.studentId].entries.push({
      date: tardy.date,
      time: tardy.time,
      parentName: tardy.parentName
    });
  });

  for (const id in grouped) {
    if (grouped[id].entries.length < minOccurrences) {
      delete grouped[id];
    } else {
      grouped[id].entries.sort((a, b) => a.date - b.date);
    }
  }

  return grouped;
}

// STUDENT HISTORY REPORT
function generateStudentHistoryReport(formData) {
  try {
    const studentIdentifier = String(formData.studentIdentifier).trim();
    
    if (!studentIdentifier) {
      return { success: false, error: 'Please provide a student name or ID' };
    }
    
    const startDate = formData.startDate ? new Date(formData.startDate + 'T00:00:00') : null;
    const endDate = formData.endDate ? new Date(formData.endDate + 'T23:59:59') : null;
    
    let allowedStatuses = [];
    if (formData.filterBoth) {
      allowedStatuses = ['In', 'Out'];
    } else {
      if (formData.filterIn) allowedStatuses.push('In');
      if (formData.filterOut) allowedStatuses.push('Out');
    }
    
    if (allowedStatuses.length === 0) {
      return { success: false, error: 'Please select at least one entry type' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const allEntries = [];
    let studentName = 'Unknown Student';
    let matchedStudentId = null;
    
    const isNumericId = /^\d+$/.test(studentIdentifier);
    const searchLower = studentIdentifier.toLowerCase();
    
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      if (!/^\d{2}-\d{4}$/.test(sheetName)) {
        return;
      }
      
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowStudentId = row[5] ? String(row[5]).trim() : '';
        const rowStudentName = row[3] ? String(row[3]).trim() : '';
        
        let isMatch = false;
        
        if (isNumericId && rowStudentId === studentIdentifier) {
          isMatch = true;
          matchedStudentId = rowStudentId;
        }
        else if (!isNumericId && rowStudentName.toLowerCase().includes(searchLower)) {
          isMatch = true;
          matchedStudentId = rowStudentId;
        }
        
        if (isMatch) {
          const rowDate = parseDate(row[0]);
          const rowStatus = row[2];
          
          if (startDate && rowDate && rowDate < startDate) continue;
          if (endDate && rowDate && rowDate > endDate) continue;
          
          if (!allowedStatuses.includes(rowStatus)) continue;
          
          if (studentName === 'Unknown Student' && rowStudentName) {
            studentName = rowStudentName;
          }
          
          allEntries.push({
            date: row[0],
            time: row[1],
            status: row[2],
            studentName: rowStudentName,
            parentName: row[4],
            studentId: rowStudentId,
            sheetName: sheetName
          });
        }
      }
    });
    
    if (allEntries.length === 0) {
      let errorMsg = `No entries found for: ${studentIdentifier}`;
      if (startDate || endDate) {
        errorMsg += ` within the specified date range`;
      }
      if (!formData.filterBoth) {
        errorMsg += ` for the selected entry types`;
      }
      return { success: false, error: errorMsg };
    }
    
    // Sort chronologically (oldest first)
    allEntries.sort((a, b) => {
      const dateA = parseDate(a.date);
      const dateB = parseDate(b.date);
      if (!dateA || !dateB) return 0;
      return dateA - dateB;
    });
    
    const dateRangeStr = startDate && endDate ? 
      ` (${formatDate(startDate)} to ${formatDate(endDate)})` : 
      (startDate ? ` (from ${formatDate(startDate)})` : 
      (endDate ? ` (to ${formatDate(endDate)})` : ''));
    
    const filterStr = formData.filterBoth ? 'All Entries' : 
      (formData.filterIn && formData.filterOut ? 'Check-Ins & Check-Outs' :
      (formData.filterIn ? 'Check-Ins Only' : 'Check-Outs Only'));
    
    const sheetName = createStudentHistorySheet(
      matchedStudentId || studentIdentifier, 
      studentName, 
      allEntries, 
      dateRangeStr, 
      filterStr
    );
    
    return {
      success: true,
      sheetName: sheetName,
      entryCount: allEntries.length,
      studentName: studentName
    };
    
  } catch (error) {
    console.error('Error in generateStudentHistoryReport:', error);
    return { 
      success: false, 
      error: 'Failed to generate report: ' + error.message 
    };
  }
}

// CREATE STUDENT HISTORY SHEET
function createStudentHistorySheet(studentId, studentName, entries, dateRangeStr, filterStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `Student History - ${studentId}`;
  
  const existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) ss.deleteSheet(existingSheet);
  
  const reportSheet = ss.insertSheet(sheetName);
  
  let row = 1;
  
  reportSheet.getRange(row, 1).setValue(`Student History Report${dateRangeStr}`);
  reportSheet.getRange(row, 1).setFontWeight('bold').setFontSize(14);
  row++;
  
  reportSheet.getRange(row, 1).setValue(`Student: ${studentName}`);
  reportSheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  reportSheet.getRange(row, 1).setValue(`Student ID: ${studentId}`);
  reportSheet.getRange(row, 1).setFontWeight('bold');
  row++;
  
  reportSheet.getRange(row, 1).setValue(`Filter: ${filterStr}`);
  reportSheet.getRange(row, 1).setFontWeight('bold').setFontColor('#1a73e8');
  row++;
  
  reportSheet.getRange(row, 1).setValue(`Total Entries: ${entries.length}`);
  reportSheet.getRange(row, 1).setFontWeight('bold');
  row += 2;
  
  const headers = ['Date', 'Time', 'Status', 'Parent Name', 'Month'];
  reportSheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  reportSheet.getRange(row, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  row++;
  
  entries.forEach(entry => {
    reportSheet.getRange(row, 1, 1, 5).setValues([[
      formatDate(entry.date),
      formatTimeForDisplay(entry.time),
      entry.status,
      entry.parentName,
      entry.sheetName
    ]]);
    
    if (entry.status === 'In') {
      reportSheet.getRange(row, 3).setBackground('#d4edda');
    } else if (entry.status === 'Out') {
      reportSheet.getRange(row, 3).setBackground('#f8d7da');
    }
    
    row++;
  });
  
  reportSheet.setFrozenRows(7);
  reportSheet.autoResizeColumns(1, headers.length);
  
  reportSheet.setColumnWidth(1, 100);
  reportSheet.setColumnWidth(2, 100);
  reportSheet.setColumnWidth(3, 80);
  reportSheet.setColumnWidth(4, 200);
  reportSheet.setColumnWidth(5, 100);
  
  return sheetName;
}

// PARSE REPORT DATA FOR EMAIL GENERATION
function parseReportData(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    const students = [];
    let currentStudent = null;
    
    let headerRow = -1;
    let dataStartRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'Student ID' && data[i][1] === 'Student Name') {
        headerRow = i;
        dataStartRow = i + 1;
        break;
      }
    }
    
    if (dataStartRow === -1) {
      console.log('Could not find header row, attempting to find data directly');
      for (let i = 0; i < data.length; i++) {
        const [col0, col1, col2] = data[i];
        
        if (col0 && col1 && 
            !isNaN(Number(col0)) && 
            String(col0).length >= 4 && 
            String(col1).trim().length > 0 &&
            !String(col1).toLowerCase().includes('student')) {
          
          dataStartRow = i;
          break;
        }
      }
      
      if (dataStartRow === -1) {
        console.log('No valid data found in report');
        return { students: [] };
      }
    }
    
    console.log(`Starting to parse data from row ${dataStartRow}`);
    
    for (let i = dataStartRow; i < data.length; i++) {
      const [studentId, studentName, grade, date, time, parentName] = data[i];
      
      if (!studentId && !studentName && !grade && !date && !time && !parentName) {
        if (currentStudent) {
          console.log(`Empty row found after student ${currentStudent.studentName}`);
        }
        continue;
      }
      
      if (studentId && studentName && String(studentId).trim() && String(studentName).trim()) {
        const idStr = String(studentId).trim();
        
        if (isNaN(Number(idStr)) || idStr.length < 3) {
          console.log(`Skipping invalid student ID: ${idStr}`);
          continue;
        }
        
        currentStudent = {
          studentId: idStr,
          studentName: String(studentName).trim(),
          grade: grade ? String(grade).trim() : 'Unknown',
          parentName: parentName ? String(parentName).trim() : '',
          tardies: []
        };
        
        students.push(currentStudent);
        console.log(`Added student: ${currentStudent.studentName} (${currentStudent.studentId}) - Grade: ${currentStudent.grade}`);
        
        if (date && time) {
          const tardyEntry = {
            date: date,
            time: time,
            parentName: parentName ? String(parentName).trim() : ''
          };
          currentStudent.tardies.push(tardyEntry);
          console.log(`  Added tardy: ${formatDate(date)} at ${time}`);
        }
      }
      else if (currentStudent && date && time) {
        const tardyEntry = {
          date: date,
          time: time,
          parentName: parentName ? String(parentName).trim() : ''
        };
        currentStudent.tardies.push(tardyEntry);
        console.log(`  Added additional Late Arrivals for ${currentStudent.studentName}: ${formatDate(date)} at ${time}`);
      }
    }
    
    students.forEach(student => {
      if (student.tardies.length > 0 && !student.parentName) {
        const parentNames = student.tardies
          .map(t => t.parentName)
          .filter(name => name && name.trim());
        
        if (parentNames.length > 0) {
          student.parentName = parentNames[0];
        }
      }
    });
    
    console.log(`Finished parsing. Found ${students.length} students with tardies`);
    
    students.forEach(student => {
      console.log(`Student: ${student.studentName} (${student.studentId}) - Grade: ${student.grade} - Tardies: ${student.tardies.length}`);
    });
    
    return { students: students };
    
  } catch (error) {
    console.error('Error in parseReportData:', error);
    console.error('Error stack:', error.stack);
    return { students: [] };
  }
}
