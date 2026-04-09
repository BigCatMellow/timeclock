// ============================================================================
// 03-EMAIL-SYSTEM.GS
// Email generation, tracker management, content creation, document generation
// ============================================================================

// EMAIL STATUS TRACKER SETUP
function getEmailStatusTracker() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let trackerSheet = ss.getSheetByName('Email Status Tracker');
    
    if (!trackerSheet) {
      trackerSheet = ss.insertSheet('Email Status Tracker');
      
      const headers = [
        'Student ID',
        'Student Name', 
        'Report Date Range',
        'Date Generated',
        'Email Sent',
        'Date Sent',
        'Override Next Time',
        'Never Send',
        'Notes'
      ];
      
      trackerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      trackerSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#34a853')
        .setFontColor('#ffffff');
      
      trackerSheet.setColumnWidth(1, 100);
      trackerSheet.setColumnWidth(2, 200);
      trackerSheet.setColumnWidth(3, 200);
      trackerSheet.setColumnWidth(4, 120);
      trackerSheet.setColumnWidth(5, 90);
      trackerSheet.setColumnWidth(6, 120);
      trackerSheet.setColumnWidth(7, 120);
      trackerSheet.setColumnWidth(8, 100);
      trackerSheet.setColumnWidth(9, 300);
      
      SpreadsheetApp.flush();
    } else {
      const headers = trackerSheet.getRange(1, 1, 1, trackerSheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Never Send')) {
        const lastCol = trackerSheet.getLastColumn();
        trackerSheet.insertColumnAfter(lastCol - 1);
        trackerSheet.getRange(1, lastCol).setValue('Never Send');
        trackerSheet.getRange(1, lastCol)
          .setFontWeight('bold')
          .setBackground('#34a853')
          .setFontColor('#ffffff');
        trackerSheet.setColumnWidth(lastCol, 100);
        
        if (trackerSheet.getLastRow() > 1) {
          const checkboxValidation = SpreadsheetApp.newDataValidation()
            .requireCheckbox()
            .build();
          trackerSheet.getRange(2, lastCol, trackerSheet.getLastRow() - 1, 1)
            .setDataValidation(checkboxValidation);
        }
      }
    }
    
    return trackerSheet;
    
  } catch (error) {
    throw new Error(`Failed to create/get Email Status Tracker: ${error.message}`);
  }
}

function logStudentToTracker(studentId, studentName, reportDateRange, notes = '') {
  try {
    const cleanStudentId = studentId ? String(studentId).trim() : 'MISSING_ID';
    const cleanStudentName = studentName ? String(studentName).trim() : 'MISSING_NAME';
    const cleanReportDateRange = reportDateRange ? String(reportDateRange).trim() : 'MISSING_RANGE';
    const cleanNotes = notes ? String(notes).trim() : 'No notes';
    
    const trackerSheet = getEmailStatusTracker();
    const currentDate = new Date();
    
    const row = [
      cleanStudentId,
      cleanStudentName,
      cleanReportDateRange,
      currentDate,
      false,
      '',
      false,
      false,
      cleanNotes
    ];
    
    trackerSheet.appendRow(row);
    
    const lastRow = trackerSheet.getLastRow();
    const emailSentCell = trackerSheet.getRange(lastRow, 5);
    const overrideCell = trackerSheet.getRange(lastRow, 7);
    const neverSendCell = trackerSheet.getRange(lastRow, 8);
    
    const checkboxValidation = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    
    emailSentCell.setDataValidation(checkboxValidation);
    overrideCell.setDataValidation(checkboxValidation);
    neverSendCell.setDataValidation(checkboxValidation);
    
    SpreadsheetApp.flush();
    
  } catch (error) {
    throw error;
  }
}

function wasStudentRecentlyEmailed(studentId, dayWindow = 14) {
  const trackerSheet = getEmailStatusTracker();
  const data = trackerSheet.getDataRange().getValues();
  
  if (data.length <= 1) return { wasEmailed: false };
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - dayWindow);
  
  for (let i = data.length - 1; i >= 1; i--) {
    const [recordStudentId, studentName, reportRange, dateGenerated, emailSent, dateSent, override, notes] = data[i];
    
    if (recordStudentId == studentId && emailSent === true) {
      const sentDate = dateSent ? new Date(dateSent) : new Date(dateGenerated);
      
      if (sentDate >= cutoffDate) {
        const daysSince = Math.floor((new Date() - sentDate) / (1000 * 60 * 60 * 24));
        return {
          wasEmailed: true,
          lastEmailDate: sentDate,
          daysSince: daysSince,
          reportRange: reportRange
        };
      }
    }
  }
  
  return { wasEmailed: false };
}

function hasOverrideFlag(studentId) {
  const trackerSheet = getEmailStatusTracker();
  const data = trackerSheet.getDataRange().getValues();
  
  if (data.length <= 1) return false;
  
  for (let i = data.length - 1; i >= 1; i--) {
    const [recordStudentId, studentName, reportRange, dateGenerated, emailSent, dateSent, override, notes] = data[i];
    
    if (recordStudentId == studentId && override === true) {
      trackerSheet.getRange(i + 1, 7).setValue(false);
      trackerSheet.getRange(i + 1, 8).setValue((notes || '') + ' [Override used]');
      return true;
    }
  }
  
  return false;
}

function isStudentNeverSend(studentId) {
  const trackerSheet = getEmailStatusTracker();
  const data = trackerSheet.getDataRange().getValues();
  
  if (data.length <= 1) return false;
  
  for (let i = data.length - 1; i >= 1; i--) {
    const [recordStudentId, studentName, reportRange, dateGenerated, emailSent, dateSent, override, neverSend, notes] = data[i];
    
    if (recordStudentId == studentId && neverSend === true) {
      return true;
    }
  }
  
  return false;
}

function logEmailGenerationToTracker(emailData, skippedStudents, reportDateRange) {
  try {
    if (emailData && emailData.length > 0) {
      emailData.forEach((email) => {
        email.students.forEach((student) => {
          logStudentToTracker(
            student.studentId,
            student.studentName,
            reportDateRange,
            `Email generated - ${email.emailType} (${student.tardies.length} tardies)`
          );
        });
      });
    }
    
    if (skippedStudents && skippedStudents.length > 0) {
      skippedStudents.forEach((item) => {
        logStudentToTracker(
          item.student.studentId,
          item.student.studentName,
          reportDateRange,
          `SKIPPED: ${item.reason}`
        );
      });
    }
    
  } catch (error) {
    throw error;
  }
}

// EMAIL GENERATION FROM REPORT
function generateEmailsFromReportWithNewTracker(formData) {
  const defaultResult = {
    success: false,
    error: 'Unknown error occurred',
    emailCount: 0,
    skippedCount: 0,
    documentName: '',
    documentUrl: '',
    todayFilterUsed: false,
    todayOnlyCount: 0
  };
  
  try {
    console.log('Starting email generation with enhanced error handling');
    console.log('Form data received:', JSON.stringify(formData));
    
    if (!formData) {
      console.error('No form data received');
      return {
        ...defaultResult,
        error: 'No form data received'
      };
    }
    
    if (!formData.reportSheetName) {
      console.error('No report sheet name specified');
      return {
        ...defaultResult,
        error: 'No report sheet name specified'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = ss.getSheetByName(formData.reportSheetName);
    
    if (!reportSheet) {
      console.error('Report sheet not found:', formData.reportSheetName);
      return { 
        ...defaultResult,
        error: 'Selected report sheet not found: ' + formData.reportSheetName
      };
    }
    
    console.log('Found report sheet, parsing data...');
    const reportData = parseReportData(reportSheet);
    
    if (!reportData) {
      console.error('parseReportData returned null or undefined');
      return {
        ...defaultResult,
        error: 'Failed to parse report data'
      };
    }
    
    if (!reportData.students || reportData.students.length === 0) {
      console.log('No student data found in report');
      return { 
        ...defaultResult,
        error: 'No student data found in the selected report'
      };
    }
    
    console.log(`Parsed ${reportData.students.length} students from report`);
    
    const reportDateRange = formData.reportSheetName.replace('Late Arrivals Report - ', '');
    
    console.log('Generating email data...');
    const trackingResult = generateEmailDataWithNewTracker(
      reportData, 
      formData.useThreshold || false, 
      formData.thresholdValue || 10,
      14,
      formData.todayOnly || false
    );
    
    if (!trackingResult) {
      console.error('generateEmailDataWithNewTracker returned null or undefined');
      return {
        ...defaultResult,
        error: 'Failed to generate email data'
      };
    }
    
    trackingResult.emailData = trackingResult.emailData || [];
    trackingResult.skippedStudents = trackingResult.skippedStudents || [];
    
    console.log(`Email data generated: ${trackingResult.emailData.length} emails, ${trackingResult.skippedStudents.length} skipped`);
    
    try {
      console.log('Logging to tracker...');
      logEmailGenerationToTracker(
        trackingResult.emailData,
        trackingResult.skippedStudents,
        reportDateRange
      );
      console.log('Successfully logged to tracker');
    } catch (trackerError) {
      console.error('Error logging to tracker:', trackerError);
    }
    
    let documentResult = null;
    
    if (trackingResult.emailData.length > 0 || trackingResult.skippedStudents.length > 0) {
      try {
        console.log('Creating email document...');
        
        let documentTitle = formData.reportSheetName;
        if (formData.todayOnly) {
          documentTitle += ' (Today Only)';
        }
        
        documentResult = createEmailDocumentWithSkipped(
          trackingResult.emailData,
          trackingResult.skippedStudents,
          documentTitle
        );
        
        if (!documentResult) {
          console.error('createEmailDocumentWithSkipped returned null');
          documentResult = {
            documentName: 'Document creation failed',
            documentUrl: '',
            documentId: ''
          };
        } else {
          console.log('Email document created successfully');
        }
      } catch (docError) {
        console.error('Error creating document:', docError);
        documentResult = {
          documentName: 'Document creation failed: ' + docError.message,
          documentUrl: '',
          documentId: ''
        };
      }
    } else {
      console.log('No emails or skipped students to document');
      documentResult = {
        documentName: 'No emails generated',
        documentUrl: '',
        documentId: ''
      };
    }
    
    const successResult = {
      success: true,
      documentName: documentResult ? documentResult.documentName : 'No document created',
      documentUrl: documentResult ? documentResult.documentUrl : '',
      emailCount: trackingResult.emailData.length,
      skippedCount: trackingResult.skippedStudents.length,
      todayFilterUsed: formData.todayOnly || false,
      todayOnlyCount: trackingResult.todayOnlyCount || 0,
      summary: trackingResult.summary || {
        totalStudents: reportData.students.length,
        emailsToGenerate: trackingResult.emailData.length,
        skippedCount: trackingResult.skippedStudents.length
      }
    };
    
    console.log('Returning success result:', JSON.stringify(successResult));
    return successResult;
    
  } catch (error) {
    console.error('ERROR in generateEmailsFromReportWithNewTracker:', error);
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    
    return {
      ...defaultResult,
      error: 'Failed to generate emails: ' + (error.message || 'Unknown error')
    };
  }
}

// EMAIL DATA GENERATION WITH TODAY FILTER
function generateEmailDataWithNewTracker(reportData, useThreshold, thresholdValue, dayWindow = 14, todayOnly = false) {
  try {
    if (!reportData || !reportData.students || reportData.students.length === 0) {
      return {
        emailData: [],
        skippedStudents: [],
        processedStudents: [],
        todayOnlyCount: 0,
        summary: { totalStudents: 0, emailsToGenerate: 0, skippedCount: 0 }
      };
    }
    
    const studentDirectory = getStudentDirectory();
    const salutationGroups = {};
    const skippedStudents = [];
    const processedStudents = [];
    let todayOnlyCount = 0;
    
    const today = new Date();
    const todayStr = formatDate(today);
    
    const earlyGrades = ['Beg. A', 'Beg. B', 'PreK A', 'PreK B', 'PreK C'];
    
    const familyGroups = buildFamilyGroups(reportData.students, studentDirectory);
    
    const familiesWithMultipleTardies = new Set();
    
    if (todayOnly) {
      Object.entries(familyGroups).forEach(([salutation, students]) => {
        const studentsWithTardiesToday = students.filter(student => {
          return student.tardies.some(tardy => {
            const tardyDateStr = tardy.date instanceof Date ? formatDate(tardy.date) : String(tardy.date);
            return tardyDateStr === todayStr;
          });
        });
        
        if (studentsWithTardiesToday.length > 1) {
          familiesWithMultipleTardies.add(salutation);
        }
      });
    } else {
      Object.entries(familyGroups).forEach(([salutation, students]) => {
        if (students.length > 1) {
          familiesWithMultipleTardies.add(salutation);
        }
      });
    }
    
    reportData.students.forEach((student) => {
      if (isStudentNeverSend(student.studentId)) {
        skippedStudents.push({
          student: student,
          reason: 'Never Send - marked in tracker',
          lastEmailDate: null,
          lastReportRange: null
        });
        return;
      }
      
      const directoryInfo = studentDirectory[student.studentId] || {};
      const salutationKey = directoryInfo.combinedSalutations || student.parentName || 'Unknown Parent';
      
      const studentGrade = student.grade || directoryInfo.grade || 'Unknown';
      const isEarlyGrade = earlyGrades.includes(studentGrade);
      
      if (isEarlyGrade && !familiesWithMultipleTardies.has(salutationKey)) {
        skippedStudents.push({
          student: student,
          reason: `Early grade (${studentGrade}) with no siblings late`,
          lastEmailDate: null,
          lastReportRange: null
        });
        return;
      }
      
      if (todayOnly) {
        const hasTardyToday = student.tardies.some(tardy => {
          const tardyDateStr = tardy.date instanceof Date ? formatDate(tardy.date) : String(tardy.date);
          return tardyDateStr === todayStr;
        });
        
        if (!hasTardyToday) {
          skippedStudents.push({
            student: student,
            reason: `No tardies today (${todayStr})`,
            lastEmailDate: null,
            lastReportRange: null
          });
          return;
        }
        
        todayOnlyCount++;
      }
      
      const recentEmailCheck = wasStudentRecentlyEmailed(student.studentId, dayWindow);
      const hasOverride = hasOverrideFlag(student.studentId);
      
      if (recentEmailCheck.wasEmailed && !hasOverride) {
        skippedStudents.push({
          student: student,
          reason: `Recently emailed ${recentEmailCheck.daysSince} days ago`,
          lastEmailDate: recentEmailCheck.lastEmailDate,
          lastReportRange: recentEmailCheck.reportRange
        });
        return;
      }
      
      const enhancedStudent = {
        ...student,
        firstName: directoryInfo.firstName || student.studentName.split(' ')[0] || '',
        lastName: directoryInfo.lastName || student.studentName.split(' ').slice(1).join(' ') || '',
        combinedSalutations: directoryInfo.combinedSalutations || '',
        primaryContactName: directoryInfo.primaryContactName || '',
        primaryContactEmail: directoryInfo.primaryContactEmail || '',
        primaryContactCell: directoryInfo.primaryContactCell || '',
        secondaryContactName: directoryInfo.secondaryContactName || '',
        secondaryContactCell: directoryInfo.secondaryContactCell || '',
        secondaryContactEmail: directoryInfo.secondaryContactEmail || ''
      };
      
      if (!salutationGroups[salutationKey]) {
        salutationGroups[salutationKey] = [];
      }
      salutationGroups[salutationKey].push(enhancedStudent);
      processedStudents.push(student);
    });
    
    const emailData = [];
    
    for (const salutation in salutationGroups) {
      const students = salutationGroups[salutation];
      
      let maxTardies = Math.max(...students.map(s => s.tardies.length));
      let emailType = 'gentle';
      
      if (useThreshold && maxTardies >= thresholdValue) {
        emailType = 'warning';
      }
      
      const isPlural = students.length > 1;
      const emailContent = generateEmailContent(students, salutation, emailType, isPlural);
      
      const contactInfo = {
        primaryContactName: students[0].primaryContactName || '',
        primaryContactEmail: students[0].primaryContactEmail || '',
        primaryContactCell: students[0].primaryContactCell || '',
        secondaryContactName: students[0].secondaryContactName || '',
        secondaryContactCell: students[0].secondaryContactCell || '',
        secondaryContactEmail: students[0].secondaryContactEmail || ''
      };
      
      emailData.push({
        parentName: students[0].parentName || 'Unknown Parent',
        combinedSalutations: salutation,
        students: students,
        emailType: emailType,
        isPlural: isPlural,
        subject: emailContent.subject,
        body: emailContent.body,
        totalTardies: students.reduce((sum, s) => sum + s.tardies.length, 0),
        contactInfo: contactInfo
      });
    }
    
    return {
      emailData: emailData,
      skippedStudents: skippedStudents,
      processedStudents: processedStudents,
      todayOnlyCount: todayOnlyCount,
      summary: {
        totalStudents: reportData.students.length,
        emailsToGenerate: emailData.length,
        skippedCount: skippedStudents.length
      }
    };
    
  } catch (error) {
    return {
      emailData: [],
      skippedStudents: [],
      processedStudents: [],
      todayOnlyCount: 0,
      summary: { totalStudents: 0, emailsToGenerate: 0, skippedCount: 0 }
    };
  }
}

function buildFamilyGroups(students, studentDirectory) {
  const familyGroups = {};
  
  students.forEach(student => {
    const directoryInfo = studentDirectory[student.studentId] || {};
    const salutationKey = directoryInfo.combinedSalutations || student.parentName || 'Unknown Parent';
    
    if (!familyGroups[salutationKey]) {
      familyGroups[salutationKey] = [];
    }
    
    familyGroups[salutationKey].push(student);
  });
  
  return familyGroups;
}

// EMAIL CONTENT GENERATION
function generateEmailContent(students, combinedSalutations, emailType, isPlural) {
  const subject = 'Late Arrivals';
  
  const salutationToUse = combinedSalutations && combinedSalutations !== 'Unknown Parent' 
    ? combinedSalutations 
    : `${students[0].parentName || 'Guardian'}`;
  
  const studentFirstNames = students.map(s => s.firstName || s.studentName.split(' ')[0]);
  const studentFirstNamesText = isPlural ? 
    (studentFirstNames.length === 2 ? 
      `${studentFirstNames[0]} and ${studentFirstNames[1]}` : 
      `${studentFirstNames.slice(0, -1).join(', ')}, and ${studentFirstNames[studentFirstNames.length - 1]}`
    ) : studentFirstNames[0];
  
  let body = '';
  
  if (emailType === 'gentle') {
    if (isPlural) {
      body = `Dear ${salutationToUse},

I hope you're doing well! I wanted to reach out because we've noticed that ${studentFirstNamesText} have arrived at school after 8:20 a.m. several times recently. We know mornings can be busy, and we want to be sure that the children start the day feeling calm and ready to learn.

Arriving between 8:00 and 8:20 a.m. allows students to settle in, connect with friends, and transition smoothly into their day. When they arrive later, it can sometimes feel a bit rushed, which we want to help avoid so that they feel happy and at ease.

If there are any extenuating circumstances making morning arrivals challenging, please let us know. We're happy to work with you to support a smoother start. I appreciate your help in making each morning a positive experience and setting your children up for the best day possible.`;
    } else {
      body = `Dear ${salutationToUse},

I hope you're doing well! I wanted to reach out because we've noticed that ${studentFirstNamesText} has arrived at school after 8:20 a.m. several times recently. We know mornings can be busy, and we want to be sure that each student starts the day feeling calm and ready to learn.

Arriving between 8:00 and 8:20 a.m. allows them to settle in, connect with friends, and transition smoothly into their day. When arriving later, it can sometimes feel a bit rushed, which we want to help avoid so that students can feel happy and at ease.

If there are any extenuating circumstances making morning arrivals challenging, please let us know. We're happy to work with you to support a smoother start. I appreciate your help in making each morning a positive experience and setting ${studentFirstNamesText} up for the best day possible.`;
    }
  } else {
    if (isPlural) {
      body = `Dear ${salutationToUse},

Despite my previous email in which I mentioned your children ${studentFirstNamesText} had been arriving late to school, it appears that this continues with unusual frequency. If there are extenuating circumstances, please let me know. I have no desire to contribute to an already stressful situation, but I would like to reiterate just how detrimental it is for children to be tardy.

I frequently tell new families that the simplest way to get the most out of the Nysmith experience is to have their child at school at 8:00 each morning. Frequently dropping off late sets students up for a different experience.

In addition to having to rush to class, which adds preventable stress, late-arriving students lose out on the valuable socialization that takes place before class begins–and most importantly, they miss hearing the teacher's instructions. Also, with older children, not receiving that key information can translate into poorer academic performance and lower grades.

We are not penalizing your children for arriving late, but tardiness can create a multitude of unfortunate circumstances. I strongly encourage you to prioritize dropping them off at school on time every day.`;
    } else {
      body = `Dear ${salutationToUse},

Despite my previous email in which I mentioned ${studentFirstNamesText} had been arriving late to school, it appears that this continues with unusual frequency. If there are extenuating circumstances, please let me know. I have no desire to contribute to an already stressful situation, but I would like to reiterate just how detrimental it is for students to be tardy.

I frequently tell new families that the simplest way to get the most out of the Nysmith experience is to have their child at school at 8:00 each morning. Frequently dropping off late sets them up for a different experience.

In addition to having to rush to class, which adds preventable stress, late-arriving students lose out on the valuable socialization that takes place before class begins–and most importantly, they miss hearing the teacher's instructions. Also, with older children, not receiving that key information can translate into poorer academic performance and lower grades.

We are not penalizing ${studentFirstNamesText} for arriving late, but tardiness can create a multitude of unfortunate circumstances. I strongly encourage you to prioritize dropping them off at school on time every day.`;
    }
  }
  
  return { subject: subject, body: body };
}

// CREATE EMAIL DOCUMENT WITH SKIPPED STUDENTS
function createEmailDocumentWithSkipped(emailData, skippedStudents, originalSheetName) {
  try {
    const safeName = originalSheetName || 'Unknown Report';
    const documentName = `Late Arrivals Emails - ${safeName.replace('Late Arrivals Report - ', '')}`;
    
    const doc = DocumentApp.create(documentName);
    const body = doc.getBody();
    body.clear();
    
    const title = body.appendParagraph(`Generated Emails - ${safeName.replace('Late Arrivals Report - ', '')}`);
    title.setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    body.appendParagraph('');
    body.appendParagraph(`Emails to send: ${emailData.length}`);
    body.appendParagraph(`Students skipped (recently emailed): ${skippedStudents.length}`);
    
    if (emailData.length > 0) {
      body.appendParagraph('');
      const tardyReportHeader = body.appendParagraph('Late Arrivals REPORT SUMMARY');
      tardyReportHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      tardyReportHeader.editAsText().setForegroundColor('#1a73e8');
      
      const allStudentsForEmails = [];
      emailData.forEach(email => {
        email.students.forEach(student => {
          allStudentsForEmails.push(student);
        });
      });
      
      allStudentsForEmails.sort((a, b) => a.studentId.localeCompare(b.studentId));
      
      allStudentsForEmails.forEach(student => {
        const studentHeader = body.appendParagraph(`${student.studentName} (${student.studentId}) - Grade: ${student.grade || 'Unknown'}`);
        studentHeader.editAsText().setBold(true);
        studentHeader.setIndentStart(20);
        
        if (student.tardies && student.tardies.length > 0) {
          student.tardies.forEach(tardy => {
            const tardyDate = tardy.date instanceof Date ? formatDate(tardy.date) : String(tardy.date);
            const tardyTime = formatTimeForDisplay(tardy.time);
            const tardyPara = body.appendParagraph(`• ${tardyDate} at ${tardyTime}`);
            tardyPara.setIndentStart(40);
            tardyPara.editAsText().setForegroundColor('#5f6368');
          });
        } else {
          const noTardiesPara = body.appendParagraph('• No Late Arrivals details available');
          noTardiesPara.setIndentStart(40);
          noTardiesPara.editAsText().setForegroundColor('#ea4335');
        }
        
        body.appendParagraph('');
      });
    }
    
    if (skippedStudents.length > 0) {
      body.appendParagraph('');
      const skippedHeader = body.appendParagraph('STUDENTS SKIPPED (Recently Emailed)');
      skippedHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      skippedHeader.editAsText().setForegroundColor('#ea4335');
      
      skippedStudents.forEach(item => {
        const skippedPara = body.appendParagraph(
          `• ${item.student.studentName} (${item.student.studentId}) - ${item.reason}`
        );
        skippedPara.setIndentStart(20);
        skippedPara.editAsText().setForegroundColor('#5f6368');
      });
      
      body.appendParagraph('');
      const notesPara = body.appendParagraph(
        'NOTE: To override the 2-week rule for urgent cases, check the "Override Next Time" box in the Email Status Tracker sheet.'
      );
      notesPara.editAsText().setItalic(true).setForegroundColor('#1a73e8');
    }
    
    if (emailData.length === 0) {
      body.appendParagraph('');
      body.appendParagraph('All students were recently emailed. No new emails generated.');
      
      doc.saveAndClose();
      return {
        documentName: documentName,
        documentUrl: doc.getUrl(),
        documentId: doc.getId()
      };
    }
    
    body.appendParagraph('');
    body.appendHorizontalRule();
    body.appendPageBreak();
    
    emailData.forEach((email, index) => {
      const studentNames = email.students.map(s => s.studentName).join(' & ');
      const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
      
      const headerText = `${studentNames} - ${currentDate}`;
      const emailHeader = body.appendParagraph(headerText);
      emailHeader.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      emailHeader.editAsText().setBold(true).setFontSize(14);
      
      const nameEndIndex = studentNames.length;
      emailHeader.editAsText().setBackgroundColor(0, nameEndIndex - 1, '#ffff00');
      
      const grades = [...new Set(email.students.map(s => s.grade || 'Unknown'))].join(', ');
      const gradePara = body.appendParagraph(grades);
      gradePara.editAsText().setBold(true);
      
      const subjectPara = body.appendParagraph(email.subject || 'Late Arrivals');
      subjectPara.editAsText().setBold(true);
      
      const emails = [];
      if (email.contactInfo && email.contactInfo.primaryContactEmail) {
        emails.push(email.contactInfo.primaryContactEmail);
      }
      if (email.contactInfo && email.contactInfo.secondaryContactEmail) {
        emails.push(email.contactInfo.secondaryContactEmail);
      }
      
      if (emails.length > 0) {
        const emailPara = body.appendParagraph(emails.join(' ; '));
        emailPara.editAsText().setBold(true);
      } else {
        const emailPara = body.appendParagraph('No email addresses available');
        emailPara.editAsText().setBold(true).setForegroundColor('#ea4335');
      }
      
      body.appendParagraph('');
      
      const emailBodyText = email.body || 'No email body generated';
      const bodyLines = emailBodyText.split('\n');
      bodyLines.forEach((line) => {
        if (line.trim() === '') {
          body.appendParagraph('');
        } else {
          const para = body.appendParagraph(line);
        }
      });
      
      if (index < emailData.length - 1) {
        body.appendPageBreak();
      }
    });
    
    doc.saveAndClose();
    
    try {
      const foldername = "Late Arrivals Email Reports";
      let folder;
      const folders = DriveApp.getFoldersByName(foldername);
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(foldername);
      }
      
      const file = DriveApp.getFileById(doc.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (folderError) {
      // Folder error doesn't break the main function
    }
    
    return {
      documentName: documentName,
      documentUrl: doc.getUrl(),
      documentId: doc.getId()
    };
    
  } catch (error) {
    throw new Error('Failed to create email document: ' + error.message);
  }
}

// TRACKER MANAGEMENT FUNCTIONS
function markEmailsAsSentInTracker(reportDateRange) {
  const trackerSheet = getEmailStatusTracker();
  const data = trackerSheet.getDataRange().getValues();
  const currentDate = new Date();
  
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const [studentId, studentName, recordReportRange, dateGenerated, emailSent, dateSent, override, notes] = data[i];
    
    if (recordReportRange === reportDateRange && emailSent !== true && !notes.includes('SKIPPED:')) {
      trackerSheet.getRange(i + 1, 5).setValue(true);
      trackerSheet.getRange(i + 1, 6).setValue(currentDate);
      updatedCount++;
    }
  }
  
  return updatedCount;
}

function generateEmailActivitySummary(days = 30) {
  const trackerSheet = getEmailStatusTracker();
  const data = trackerSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return {
      totalEntries: 0,
      emailsSent: 0,
      skippedStudents: 0,
      pendingEmails: 0,
      recentActivity: []
    };
  }
  
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - days);
  
  let totalEntries = 0;
  let emailsSent = 0;
  let skippedStudents = 0;
  let pendingEmails = 0;
  const recentActivity = [];
  
  for (let i = 1; i < data.length; i++) {
    const [studentId, studentName, reportRange, dateGenerated, emailSent, dateSent, override, notes] = data[i];
    const genDate = new Date(dateGenerated);
    
    if (genDate >= cutoffDate) {
      totalEntries++;
      
      if (notes && notes.includes('SKIPPED:')) {
        skippedStudents++;
      } else if (emailSent === true) {
        emailsSent++;
      } else {
        pendingEmails++;
      }
      
      recentActivity.push({
        studentId: studentId,
        studentName: studentName,
        reportRange: reportRange,
        dateGenerated: genDate,
        status: notes && notes.includes('SKIPPED:') ? 'Skipped' : (emailSent ? 'Sent' : 'Pending'),
        notes: notes
      });
    }
  }
  
  recentActivity.sort((a, b) => b.dateGenerated - a.dateGenerated);
  
  return {
    totalEntries: totalEntries,
    emailsSent: emailsSent,
    skippedStudents: skippedStudents,
    pendingEmails: pendingEmails,
    recentActivity: recentActivity.slice(0, 50)
  };
}

function showEmailActivitySummary() {
  const summary = generateEmailActivitySummary(30);
  
  let summaryText = `EMAIL ACTIVITY SUMMARY (Last 30 Days)\n\n`;
  summaryText += `Total Entries: ${summary.totalEntries}\n`;
  summaryText += `Emails Sent: ${summary.emailsSent}\n`;
  summaryText += `Students Skipped: ${summary.skippedStudents}\n`;
  summaryText += `Pending Emails: ${summary.pendingEmails}\n\n`;
  
  if (summary.recentActivity.length > 0) {
    summaryText += `RECENT ACTIVITY (showing ${Math.min(10, summary.recentActivity.length)} most recent):\n\n`;
    summary.recentActivity.slice(0, 10).forEach(activity => {
      summaryText += `${formatDate(activity.dateGenerated)} - ${activity.studentName} (${activity.studentId})\n`;
      summaryText += `  Status: ${activity.status} | Report: ${activity.reportRange}\n\n`;
    });
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Email Activity Summary', summaryText, ui.ButtonSet.OK);
}

function openEmailStatusTracker() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackerSheet = getEmailStatusTracker();
    
    SpreadsheetApp.flush();
    
    const dataRange = trackerSheet.getDataRange();
    const rowCount = dataRange.getNumRows();
    const dataCount = rowCount > 1 ? rowCount - 1 : 0;
    
    const url = ss.getUrl() + '#gid=' + trackerSheet.getSheetId();
    
    const ui = SpreadsheetApp.getUi();
    
    const summary = `Email Status Tracker Sheet Info:

Sheet Name: ${trackerSheet.getName()}
Total Records: ${dataCount} (plus header row)
Last Updated: ${new Date().toLocaleString()}

Click OK to get the direct link to open the sheet.`;
    
    ui.alert('Email Status Tracker Status', summary, ui.ButtonSet.OK);
    
    ui.alert(
      'Email Status Tracker Sheet Link', 
      `Open this URL in a new tab to view the tracker:\n\n${url}\n\nIf the sheet appears empty, try refreshing the page or reopening the spreadsheet.`, 
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', `Could not open Email Status Tracker: ${error.message}`, ui.ButtonSet.OK);
  }
}

function addReportToTracker(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportSheet = ss.getSheetByName(formData.reportSheetName);
    
    if (!reportSheet) {
      return { success: false, error: 'Selected report sheet not found: ' + formData.reportSheetName };
    }
    
    const reportData = parseReportData(reportSheet);
    
    if (!reportData || !reportData.students || reportData.students.length === 0) {
      return { success: false, error: 'No student data found in the selected report' };
    }
    
    const reportDateRange = formData.reportSheetName.replace('Late Arrivals Report - ', '');
    let statusNotes = '';
    let emailSentStatus = false;
    let dateSent = '';
    
    switch (formData.status) {
      case 'sent':
        statusNotes = 'Manually added - emails already sent';
        emailSentStatus = true;
        dateSent = new Date();
        break;
      case 'skipped':
        statusNotes = 'Manually added - emails skipped';
        emailSentStatus = false;
        break;
      case 'pending':
      default:
        statusNotes = 'Manually added - emails pending';
        emailSentStatus = false;
        break;
    }
    
    if (formData.customNotes) {
      statusNotes += ` | ${formData.customNotes}`;
    }
    
    let studentsAdded = 0;
    const trackerSheet = getEmailStatusTracker();
    
    reportData.students.forEach((student) => {
      try {
        const row = [
          student.studentId || 'MISSING_ID',
          student.studentName || 'MISSING_NAME',
          reportDateRange,
          new Date(),
          emailSentStatus,
          dateSent,
          false,
          statusNotes
        ];
        
        trackerSheet.appendRow(row);
        
        const lastRow = trackerSheet.getLastRow();
        const emailSentCell = trackerSheet.getRange(lastRow, 5);
        const overrideCell = trackerSheet.getRange(lastRow, 7);
        
        const checkboxValidation = SpreadsheetApp.newDataValidation()
          .requireCheckbox()
          .build();
        
        emailSentCell.setDataValidation(checkboxValidation);
        overrideCell.setDataValidation(checkboxValidation);
        
        studentsAdded++;
        
      } catch (studentError) {
        // Continue with other students if one fails
      }
    });
    
    SpreadsheetApp.flush();
    
    return {
      success: true,
      reportName: formData.reportSheetName,
      studentsAdded: studentsAdded,
      status: formData.status.charAt(0).toUpperCase() + formData.status.slice(1),
      notes: formData.customNotes || null
    };
    
  } catch (error) {
    return { success: false, error: 'Failed to add students to tracker: ' + error.message };
  }
}
