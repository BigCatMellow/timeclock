/**
 * Parent Contact Consistency Checker
 * This script identifies and helps fix inconsistent parent contact information
 * for students sharing the same Parent ID
 */

// Configuration - Update these if your columns are different
var PARENT_CONTACT_CONFIG = {
  dataSheet: 'parent db', // Change to your sheet name
  studentIdCol: 1,
  parentIdCol: 6,
  homeRoomCol: 2,
  firstNameCol: 3,
  nicknameCol: 4,
  lastNameCol: 5,
  primaryRelCol: 7,
  primaryNameCol: 8,
  primarySalutationCol: 9,
  primaryPhoneCol: 10,
  primaryEmailCol: 11,
  secondaryRelCol: 12,
  secondaryNameCol: 13,
  secondarySalutationCol: 14,
  secondaryPhoneCol: 15,
  secondaryEmailCol: 16
};

// Relationship standardization mapping
var PARENT_RELATIONSHIP_MAP = {
  'father': 'Father',
  'dad': 'Father',
  'daddy': 'Father',
  'mother': 'Mother',
  'mom': 'Mother',
  'mommy': 'Mother',
  'mum': 'Mother',
  'guardian': 'Guardian',
  'grandparent': 'Grandparent',
  'grandmother': 'Grandmother',
  'grandfather': 'Grandfather',
  'grandma': 'Grandmother',
  'grandpa': 'Grandfather',
  'aunt': 'Aunt',
  'uncle': 'Uncle',
  'stepfather': 'Stepfather',
  'stepmother': 'Stepmother',
  'stepdad': 'Stepfather',
  'stepmom': 'Stepmother',
  'other': 'Other',
  'unknown': 'Unknown',
  '': 'Unknown'
};

/**
 * Creates custom menu when spreadsheet opens
 */


/**
 * Step 1: Find all inconsistencies and create a report
 */
function findParentContactInconsistencies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(PARENT_CONTACT_CONFIG.dataSheet);
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "' + PARENT_CONTACT_CONFIG.dataSheet + '" not found. Please update PARENT_CONTACT_CONFIG.dataSheet in the script.');
    return;
  }
  
  var data = dataSheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  // Group students by Parent ID
  var parentGroups = {};
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var parentId = row[PARENT_CONTACT_CONFIG.parentIdCol - 1];
    
    if (parentId && parentId !== '') {
      if (!parentGroups[parentId]) {
        parentGroups[parentId] = [];
      }
      
      parentGroups[parentId].push({
        rowNum: i + 2,
        studentId: row[PARENT_CONTACT_CONFIG.studentIdCol - 1],
        firstName: row[PARENT_CONTACT_CONFIG.firstNameCol - 1],
        lastName: row[PARENT_CONTACT_CONFIG.lastNameCol - 1],
        homeRoom: row[PARENT_CONTACT_CONFIG.homeRoomCol - 1],
        primaryRel: row[PARENT_CONTACT_CONFIG.primaryRelCol - 1],
        primaryName: row[PARENT_CONTACT_CONFIG.primaryNameCol - 1],
        primaryPhone: row[PARENT_CONTACT_CONFIG.primaryPhoneCol - 1],
        primaryEmail: row[PARENT_CONTACT_CONFIG.primaryEmailCol - 1],
        secondaryRel: row[PARENT_CONTACT_CONFIG.secondaryRelCol - 1],
        secondaryName: row[PARENT_CONTACT_CONFIG.secondaryNameCol - 1],
        secondaryPhone: row[PARENT_CONTACT_CONFIG.secondaryPhoneCol - 1],
        secondaryEmail: row[PARENT_CONTACT_CONFIG.secondaryEmailCol - 1]
      });
    }
  }
  
  // Find inconsistencies
  var issues = [];
  
  for (var parentId in parentGroups) {
    var students = parentGroups[parentId];
    
    if (students.length > 1) {
      var inconsistencies = checkParentContactConsistency(students);
      
      if (inconsistencies.length > 0) {
        issues.push({
          parentId: parentId,
          students: students,
          issues: inconsistencies
        });
      }
    }
  }
  
  // Create report sheet
  createParentContactInconsistencyReport(ss, issues);
  
  SpreadsheetApp.getUi().alert(
    'Inconsistency Check Complete',
    'Found ' + issues.length + ' Parent IDs with inconsistencies.\n\nCheck the "Inconsistency Report" sheet for details.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Check consistency within a group of students
 */
function checkParentContactConsistency(students) {
  var issues = [];
  
  // Get unique values for each field
  var primaryRels = getUniqueValues(students.map(function(s) { return normalizeParentString(s.primaryRel); }));
  var primaryNames = getUniqueValues(students.map(function(s) { return normalizeParentString(s.primaryName); }));
  var primaryPhones = getUniqueValues(students.map(function(s) { return normalizeParentString(s.primaryPhone); }));
  var primaryEmails = getUniqueValues(students.map(function(s) { return normalizeParentString(s.primaryEmail); }));
  
  var secondaryRels = getUniqueValues(students.map(function(s) { return normalizeParentString(s.secondaryRel); }));
  var secondaryNames = getUniqueValues(students.map(function(s) { return normalizeParentString(s.secondaryName); }));
  var secondaryPhones = getUniqueValues(students.map(function(s) { return normalizeParentString(s.secondaryPhone); }));
  var secondaryEmails = getUniqueValues(students.map(function(s) { return normalizeParentString(s.secondaryEmail); }));
  
  if (primaryRels.length > 1) {
    issues.push('Primary relationship varies: ' + primaryRels.join(', '));
  }
  
  if (primaryNames.length > 1) {
    issues.push('Primary contact name varies: ' + primaryNames.join(', '));
  }
  
  if (primaryPhones.length > 1) {
    issues.push('Primary phone varies: ' + primaryPhones.join(', '));
  }
  
  if (primaryEmails.length > 1) {
    issues.push('Primary email varies: ' + primaryEmails.join(', '));
  }
  
  if (secondaryRels.length > 1) {
    issues.push('Secondary relationship varies: ' + secondaryRels.join(', '));
  }
  
  if (secondaryNames.length > 1) {
    issues.push('Secondary contact name varies: ' + secondaryNames.join(', '));
  }
  
  if (secondaryPhones.length > 1) {
    issues.push('Secondary phone varies: ' + secondaryPhones.join(', '));
  }
  
  if (secondaryEmails.length > 1) {
    issues.push('Secondary email varies: ' + secondaryEmails.join(', '));
  }
  
  return issues;
}

/**
 * Helper to get unique values from array
 */
function getUniqueValues(arr) {
  var unique = [];
  for (var i = 0; i < arr.length; i++) {
    if (unique.indexOf(arr[i]) === -1) {
      unique.push(arr[i]);
    }
  }
  return unique;
}

/**
 * Create a detailed inconsistency report
 */
function createParentContactInconsistencyReport(ss, issues) {
  // Delete existing report if it exists
  var existingSheet = ss.getSheetByName('Inconsistency Report');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  var reportSheet = ss.insertSheet('Inconsistency Report');
  
  // Headers
  var headers = [
    'Parent ID',
    'Issue Count',
    'Student Names',
    'Row Numbers',
    'Issues Found',
    'Primary Rel Values',
    'Secondary Rel Values',
    'Recommendations'
  ];
  
  reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  reportSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Data
  var reportData = [];
  for (var i = 0; i < issues.length; i++) {
    var issue = issues[i];
    var studentNames = issue.students.map(function(s) { return s.firstName + ' ' + s.lastName; }).join(', ');
    var rowNumbers = issue.students.map(function(s) { return s.rowNum; }).join(', ');
    var issuesText = issue.issues.join('\n');
    
    var primaryRels = getUniqueValues(issue.students.map(function(s) { return s.primaryRel; })).join(', ');
    var secondaryRels = getUniqueValues(issue.students.map(function(s) { return s.secondaryRel; })).join(', ');
    
    var recommendation = generateParentContactRecommendation(issue);
    
    reportData.push([
      issue.parentId,
      issue.issues.length,
      studentNames,
      rowNumbers,
      issuesText,
      primaryRels,
      secondaryRels,
      recommendation
    ]);
  }
  
  if (reportData.length > 0) {
    reportSheet.getRange(2, 1, reportData.length, headers.length).setValues(reportData);
  }
  
  // Formatting
  reportSheet.setFrozenRows(1);
  reportSheet.autoResizeColumns(1, headers.length);
  reportSheet.getRange(2, 5, reportData.length, 1).setWrap(true);
  reportSheet.getRange(2, 8, reportData.length, 1).setWrap(true);
  
  // Alternating row colors
  for (var i = 2; i <= reportData.length + 1; i++) {
    if (i % 2 === 0) {
      reportSheet.getRange(i, 1, 1, headers.length).setBackground('#f3f3f3');
    }
  }
}

/**
 * Generate recommendations for fixing issues
 */
function generateParentContactRecommendation(issue) {
  var recommendations = [];
  
  // Find most common values
  var primaryRels = [];
  var secondaryRels = [];
  
  for (var i = 0; i < issue.students.length; i++) {
    if (issue.students[i].primaryRel && issue.students[i].primaryRel !== '') {
      primaryRels.push(issue.students[i].primaryRel);
    }
    if (issue.students[i].secondaryRel && issue.students[i].secondaryRel !== '') {
      secondaryRels.push(issue.students[i].secondaryRel);
    }
  }
  
  if (primaryRels.length > 0) {
    var mostCommonPrimary = findMostCommonParentValue(primaryRels);
    recommendations.push('Standardize primary to: "' + mostCommonPrimary + '"');
  }
  
  if (secondaryRels.length > 0) {
    var mostCommonSecondary = findMostCommonParentValue(secondaryRels);
    recommendations.push('Standardize secondary to: "' + mostCommonSecondary + '"');
  }
  
  return recommendations.join('\n');
}

/**
 * Step 2: Standardize relationship names
 */
function standardizeParentRelationships() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(PARENT_CONTACT_CONFIG.dataSheet);
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet not found');
    return;
  }
  
  var data = dataSheet.getDataRange().getValues();
  var rows = data.slice(1);
  var changesCount = 0;
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rowNum = i + 2;
    
    // Standardize primary relationship
    var primaryRel = row[PARENT_CONTACT_CONFIG.primaryRelCol - 1];
    if (primaryRel) {
      var standardized = standardizeParentRelationship(primaryRel);
      if (standardized !== primaryRel) {
        dataSheet.getRange(rowNum, PARENT_CONTACT_CONFIG.primaryRelCol).setValue(standardized);
        changesCount++;
      }
    }
    
    // Standardize secondary relationship
    var secondaryRel = row[PARENT_CONTACT_CONFIG.secondaryRelCol - 1];
    if (secondaryRel) {
      var standardized = standardizeParentRelationship(secondaryRel);
      if (standardized !== secondaryRel) {
        dataSheet.getRange(rowNum, PARENT_CONTACT_CONFIG.secondaryRelCol).setValue(standardized);
        changesCount++;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(
    'Standardization Complete',
    'Standardized ' + changesCount + ' relationship fields.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Standardize a relationship string
 */
function standardizeParentRelationship(rel) {
  if (!rel) return 'Unknown';
  
  var normalized = rel.toString().toLowerCase().trim();
  return PARENT_RELATIONSHIP_MAP[normalized] || rel;
}

/**
 * Step 3: Generate comprehensive report
 */
function generateParentContactFullReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(PARENT_CONTACT_CONFIG.dataSheet);
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet not found');
    return;
  }
  
  var data = dataSheet.getDataRange().getValues();
  var rows = data.slice(1);
  
  // Create summary sheet
  var existingSheet = ss.getSheetByName('Parent ID Summary');
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  
  var summarySheet = ss.insertSheet('Parent ID Summary');
  
  // Group by Parent ID
  var parentGroups = {};
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var parentId = row[PARENT_CONTACT_CONFIG.parentIdCol - 1];
    
    if (parentId && parentId !== '') {
      if (!parentGroups[parentId]) {
        parentGroups[parentId] = [];
      }
      
      parentGroups[parentId].push({
        rowNum: i + 2,
        studentId: row[PARENT_CONTACT_CONFIG.studentIdCol - 1],
        firstName: row[PARENT_CONTACT_CONFIG.firstNameCol - 1],
        lastName: row[PARENT_CONTACT_CONFIG.lastNameCol - 1],
        homeRoom: row[PARENT_CONTACT_CONFIG.homeRoomCol - 1],
        primaryRel: row[PARENT_CONTACT_CONFIG.primaryRelCol - 1],
        primaryName: row[PARENT_CONTACT_CONFIG.primaryNameCol - 1],
        secondaryRel: row[PARENT_CONTACT_CONFIG.secondaryRelCol - 1],
        secondaryName: row[PARENT_CONTACT_CONFIG.secondaryNameCol - 1]
      });
    }
  }
  
  // Create summary
  var headers = [
    'Parent ID',
    'Student Count',
    'Students',
    'Home Rooms',
    'Primary Contact',
    'Secondary Contact',
    'Status',
    'Action Needed'
  ];
  
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  summarySheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  var summaryData = [];
  for (var parentId in parentGroups) {
    var students = parentGroups[parentId];
    var studentNames = students.map(function(s) { return s.firstName + ' ' + s.lastName; }).join(', ');
    var homeRooms = getUniqueValues(students.map(function(s) { return s.homeRoom; })).join(', ');
    
    var primaryContacts = getUniqueValues(students.map(function(s) { return s.primaryRel + ': ' + s.primaryName; }));
    var secondaryContacts = getUniqueValues(students.map(function(s) { return s.secondaryRel + ': ' + s.secondaryName; }));
    
    var isConsistent = primaryContacts.length === 1 && secondaryContacts.length === 1;
    var status = isConsistent ? '✅ Consistent' : '⚠️ Needs Review';
    var action = isConsistent ? 'None' : 'Review and standardize contacts';
    
    summaryData.push([
      parentId,
      students.length,
      studentNames,
      homeRooms,
      primaryContacts.join('\n'),
      secondaryContacts.join('\n'),
      status,
      action
    ]);
  }
  
  summarySheet.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
  
  // Formatting
  summarySheet.setFrozenRows(1);
  summarySheet.autoResizeColumns(1, headers.length);
  
  var parentCount = 0;
  for (var key in parentGroups) {
    if (parentGroups.hasOwnProperty(key)) parentCount++;
  }
  
  SpreadsheetApp.getUi().alert(
    'Full Report Generated',
    'Created summary for ' + parentCount + ' unique Parent IDs.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Run all three steps in sequence
 */
function runAllParentContactSteps() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Run All Steps',
    'This will:\n1. Standardize all relationship names\n2. Find inconsistencies\n3. Generate a full report\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    standardizeParentRelationships();
    findParentContactInconsistencies();
    generateParentContactFullReport();
    
    ui.alert(
      'All Steps Complete!',
      'Check the "Inconsistency Report" and "Parent ID Summary" sheets for results.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Helper function to normalize strings for comparison
 */
function normalizeParentString(str) {
  if (!str) return '';
  return str.toString().trim().toLowerCase();
}

/**
 * Helper function to find most common item in array
 */
function findMostCommonParentValue(arr) {
  var frequency = {};
  var maxCount = 0;
  var mostCommon = arr[0];
  
  for (var i = 0; i < arr.length; i++) {
    var item = arr[i];
    frequency[item] = (frequency[item] || 0) + 1;
    if (frequency[item] > maxCount) {
      maxCount = frequency[item];
      mostCommon = item;
    }
  }
  
  return mostCommon;
}
