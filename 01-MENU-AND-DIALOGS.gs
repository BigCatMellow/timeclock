// ============================================================================
// 01-MENU-AND-DIALOGS.GS
// Menu setup, triggers, and all UI dialogs
// ============================================================================

// MAIN MENU SETUP
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Reports')
    .addItem('Generate Late Arrivals Report', 'showCalendarDatePicker')
    .addSeparator()
    .addItem('Student History Report', 'showStudentHistoryDialog')
    .addSeparator()
    .addItem('Delete Old Late Arrivals Reports', 'deleteOldTardyReports')
    .addSeparator()
    .addItem('Generate Emails from Report', 'showEmailGenerationDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('Email Status Tracker')
      .addItem('Open Email Status Tracker', 'openEmailStatusTracker')
      .addItem('Email Activity Summary', 'showEmailActivitySummary')
      .addItem('Mark Emails as Sent', 'showMarkEmailsSentDialog')
      .addSeparator()
      .addItem('Manage Never Send List', 'showNeverSendManager')
      .addSeparator()
      .addItem('Manually Add Report to Tracker', 'showManualTrackerDialog')
      .addSeparator()
      .addItem('Tracker Manager', 'showEmailTrackerManager'))
    .addToUi();
}

// AUTO-POPULATE DATE SENT WHEN EMAIL SENT IS CHECKED
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() !== 'Email Status Tracker') return;
  if (range.getColumn() !== 5) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
  if (range.getRow() === 1) return;
  
  const emailSentValue = range.getValue();
  const row = range.getRow();
  
  if (emailSentValue === true) {
    const dateSentCell = sheet.getRange(row, 6);
    if (!dateSentCell.getValue()) {
      dateSentCell.setValue(new Date());
    }
  } else if (emailSentValue === false) {
    const dateSentCell = sheet.getRange(row, 6);
    dateSentCell.clearContent();
  }
}

// CALENDAR DATE PICKER DIALOG
function showCalendarDatePicker() {
  const htmlTemplate = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Google Sans', Roboto, Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f8f9fa;
    }
    
    .container {
      max-width: 900px;
      background: white;
      border-radius: 12px;
      padding: 24px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.1);
    }
    
    .form-row {
      display: flex;
      gap: 20px;
      margin-bottom: 20px;
    }
    
    .form-row .form-group {
      flex: 1;
      margin-bottom: 0;
    }
    
    .filters-row {
      display: flex;
      gap: 20px;
      align-items: end;
    }
    
    .filters-row .form-group {
      flex: 1;
    }
    
    .header {
      text-align: center;
      margin-bottom: 24px;
    }
    
    .title {
      color: #1a73e8;
      font-size: 24px;
      font-weight: 500;
      margin: 0;
    }
    
    .subtitle {
      color: #5f6368;
      font-size: 14px;
      margin: 8px 0 0 0;
    }
    
    .form-group {
      margin-bottom: 20px;
    }
    
    .form-label {
      display: block;
      color: #202124;
      font-size: 14px;
      font-weight: 500;
      margin-bottom: 8px;
    }
    
    .form-input {
      width: 100%;
      padding: 12px;
      border: 2px solid #dadce0;
      border-radius: 8px;
      font-size: 16px;
      box-sizing: border-box;
      transition: border-color 0.2s;
    }
    
    .form-input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 3px rgba(26,115,232,0.1);
    }
    
    .time-group {
      display: flex;
      gap: 12px;
    }
    
    .time-group .form-input {
      flex: 1;
    }
    
    .options-section {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 8px;
      margin: 20px 0;
    }
    
    .options-title {
      color: #202124;
      font-size: 16px;
      font-weight: 500;
      margin-bottom: 12px;
    }
    
    .checkbox-group {
      display: flex;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .checkbox-group input[type="checkbox"] {
      margin-right: 8px;
      transform: scale(1.2);
    }
    
    .checkbox-group label {
      color: #202124;
      font-size: 14px;
      cursor: pointer;
    }
    
    .input-group {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 8px;
    }
    
    .input-group input {
      flex: 1;
      padding: 8px 12px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      font-size: 14px;
    }
    
    .button-group {
      display: flex;
      gap: 12px;
      margin-top: 24px;
    }
    
    .btn {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
    }
    
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    
    .btn-primary:hover {
      background: #1557b2;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(26,115,232,0.3);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #5f6368;
      border: 1px solid #dadce0;
    }
    
    .btn-secondary:hover {
      background: #e8eaed;
    }
    
    .btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    
    .error {
      color: #d93025;
      font-size: 14px;
      margin-top: 8px;
    }
    
    .date-hint {
      color: #5f6368;
      font-size: 12px;
      margin-top: 4px;
    }
    
    .time-hint {
      color: #5f6368;
      font-size: 12px;
      margin-top: 4px;
      text-align: center;
    }
    
    .loading {
      display: none;
      text-align: center;
      color: #5f6368;
      font-size: 14px;
      margin-top: 16px;
    }
    
    .loading::before {
      content: '';
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid #dadce0;
      border-top-color: #1a73e8;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-right: 8px;
      vertical-align: middle;
    }
    
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 class="title">Select Date and Time Range</h1>
      <p class="subtitle">Choose date range and time frame for your Late Arrivals report</p>
    </div>
    
    <form id="dateForm">
      <!-- Date Selection Row -->
      <div class="form-row">
        <div class="form-group">
          <label class="form-label" for="startDate">Start Date</label>
          <input type="date" id="startDate" class="form-input" required>
          <div class="date-hint">First day to include</div>
        </div>
        
        <div class="form-group">
          <label class="form-label" for="endDate">End Date</label>
          <input type="date" id="endDate" class="form-input" required>
          <div class="date-hint">Last day to include</div>
        </div>
      </div>
      
      <!-- Time Frame Row -->
      <div class="form-row">
        <div class="form-group">
          <label class="form-label" for="startTime">Start Time</label>
          <input type="time" id="startTime" class="form-input" value="08:00" required>
          <div class="time-hint">Earliest time to include</div>
        </div>
        
        <div class="form-group">
          <label class="form-label" for="endTime">End Time</label>
          <input type="time" id="endTime" class="form-input" value="09:00" required>
          <div class="time-hint">Latest time to include</div>
        </div>
      </div>
      
      <!-- Filter Options Row -->
      <div class="options-section">
        <div class="options-title">Filter Options (Optional)</div>
        <div class="filters-row">
          <div class="form-group">
            <label class="form-label" for="studentId">Student ID Filter</label>
            <input type="text" id="studentId" placeholder="Leave blank for all students" class="form-input">
            <div class="date-hint">Specific Student ID or leave blank</div>
          </div>
          
          <div class="form-group">
            <label class="form-label" for="minOccurrences">Minimum Tardies</label>
            <input type="number" id="minOccurrences" placeholder="Leave blank for all" class="form-input" min="1" max="50" value="5">
            <div class="date-hint">Minimum number of tardies or leave blank</div>
          </div>
        </div>
      </div>
      
      <div id="errorMessage" class="error"></div>
      <div id="loadingMessage" class="loading">Generating report...</div>
      
      <div class="button-group">
        <button type="button" class="btn btn-secondary" onclick="closeDialog()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="generateBtn">Generate Report</button>
      </div>
    </form>
  </div>

  <script>
    window.onload = function() {
      const today = new Date();
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(today.getDate() - 14);
      
      document.getElementById('endDate').value = formatDateForInput(today);
      document.getElementById('startDate').value = formatDateForInput(thirtyDaysAgo);
    };
    
    function formatDateForInput(date) {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return year + '-' + month + '-' + day;
    }
    
    function showError(message) {
      document.getElementById('errorMessage').textContent = message;
    }
    
    function clearError() {
      document.getElementById('errorMessage').textContent = '';
    }
    
    function setLoading(loading) {
      document.getElementById('loadingMessage').style.display = loading ? 'block' : 'none';
      document.getElementById('generateBtn').disabled = loading;
    }
    
    function closeDialog() {
      google.script.host.close();
    }
    
    document.getElementById('dateForm').addEventListener('submit', function(e) {
      e.preventDefault();
      clearError();
      
      const startDate = document.getElementById('startDate').value;
      const endDate = document.getElementById('endDate').value;
      const startTime = document.getElementById('startTime').value;
      const endTime = document.getElementById('endTime').value;
      
      if (!startDate || !endDate) {
        showError('Please select both start and end dates');
        return;
      }
      
      if (!startTime || !endTime) {
        showError('Please select both start and end times');
        return;
      }
      
      if (new Date(startDate) > new Date(endDate)) {
        showError('Start date cannot be after end date');
        return;
      }
      
      if (startTime >= endTime) {
        showError('Start time must be before end time');
        return;
      }
      
      const formData = {
        startDate: startDate,
        endDate: endDate,
        startTime: startTime,
        endTime: endTime,
        studentId: document.getElementById('studentId').value.trim() || '',
        minOccurrences: parseInt(document.getElementById('minOccurrences').value) || 1
      };
      
      setLoading(true);
      
      google.script.run
        .withSuccessHandler(function(result) {
          setLoading(false);
          if (result.success) {
            alert('Report generated successfully!\\n\\n' +
                  'Sheet: "' + result.sheetName + '"\\n' +
                  'Total Late Arrivals instances: ' + result.tardyCount + '\\n' +
                  'Students shown: ' + result.studentCount);
            closeDialog();
          } else {
            showError(result.error);
          }
        })
        .withFailureHandler(function(error) {
          setLoading(false);
          showError('Error generating report: ' + error.message);
        })
        .generateTardyReportFromCalendar(formData);
    });
  </script>
</body>
</html>
  `);
  
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(950)
    .setHeight(630)
    .setTitle('Select Date and Time Range for Late Arrivals Report');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Late Arrivals Report Date and Time Selection');
}

// STUDENT HISTORY DIALOG
function showStudentHistoryDialog() {
  const htmlTemplate = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Google Sans', Roboto, Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f8f9fa;
    }
    
    .container {
      max-width: 700px;
      background: white;
      border-radius: 12px;
      padding: 24px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.1);
    }
    
    .header {
      text-align: center;
      margin-bottom: 24px;
    }
    
    .title {
      color: #1a73e8;
      font-size: 24px;
      font-weight: 500;
      margin: 0;
    }
    
    .subtitle {
      color: #5f6368;
      font-size: 14px;
      margin: 8px 0 0 0;
    }
    
    .form-row {
      display: flex;
      gap: 20px;
      margin-bottom: 20px;
    }
    
    .form-row .form-group {
      flex: 1;
      margin-bottom: 0;
    }
    
    .form-group {
      margin-bottom: 20px;
    }
    
    .form-label {
      display: block;
      color: #202124;
      font-size: 14px;
      font-weight: 500;
      margin-bottom: 8px;
    }
    
    .form-input {
      width: 100%;
      padding: 12px;
      border: 2px solid #dadce0;
      border-radius: 8px;
      font-size: 16px;
      box-sizing: border-box;
      transition: border-color 0.2s;
    }
    
    .form-input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 3px rgba(26,115,232,0.1);
    }
    
    .date-hint {
      color: #5f6368;
      font-size: 12px;
      margin-top: 4px;
    }
    
    .options-section {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 8px;
      margin: 20px 0;
    }
    
    .options-title {
      color: #202124;
      font-size: 16px;
      font-weight: 500;
      margin-bottom: 12px;
    }
    
    .checkbox-group {
      display: flex;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .checkbox-group input[type="checkbox"] {
      margin-right: 8px;
      transform: scale(1.2);
    }
    
    .checkbox-group label {
      color: #202124;
      font-size: 14px;
      cursor: pointer;
    }
    
    .button-group {
      display: flex;
      gap: 12px;
      margin-top: 24px;
    }
    
    .btn {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
    }
    
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    
    .btn-primary:hover {
      background: #1557b2;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(26,115,232,0.3);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #5f6368;
      border: 1px solid #dadce0;
    }
    
    .btn-secondary:hover {
      background: #e8eaed;
    }
    
    .btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    
    .error {
      color: #d93025;
      font-size: 14px;
      margin-top: 8px;
    }
    
    .loading {
      display: none;
      text-align: center;
      color: #5f6368;
      font-size: 14px;
      margin-top: 16px;
    }
    
    .loading::before {
      content: '';
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid #dadce0;
      border-top-color: #1a73e8;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-right: 8px;
      vertical-align: middle;
    }
    
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 class="title">Student History Report</h1>
      <p class="subtitle">View all check-in/out entries for a specific student</p>
    </div>
    
    <form id="historyForm">
      <div class="form-group">
        <label class="form-label" for="studentIdentifier">Student Name or ID</label>
        <input 
          type="text" 
          id="studentIdentifier" 
          class="form-input" 
          placeholder="Enter student's name or ID..."
          required>
        <div class="date-hint">Enter the student's full name or ID number</div>
      </div>
      
      <!-- Date Range Selection -->
      <div class="form-row">
        <div class="form-group">
          <label class="form-label" for="startDate">Start Date (Optional)</label>
          <input type="date" id="startDate" class="form-input">
          <div class="date-hint">Leave blank for all dates</div>
        </div>
        
        <div class="form-group">
          <label class="form-label" for="endDate">End Date (Optional)</label>
          <input type="date" id="endDate" class="form-input">
          <div class="date-hint">Leave blank for all dates</div>
        </div>
      </div>
      
      <!-- Status Filter -->
      <div class="options-section">
        <div class="options-title">Entry Type Filter</div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="filterIn">
          <label for="filterIn">Include Check-Ins</label>
        </div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="filterOut">
          <label for="filterOut">Include Check-Outs</label>
        </div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="filterBoth" checked>
          <label for="filterBoth">Include Both (All Entries)</label>
        </div>
      </div>
      
      <div id="errorMessage" class="error"></div>
      <div id="loadingMessage" class="loading">Generating report...</div>
      
      <div class="button-group">
        <button type="button" class="btn btn-secondary" onclick="closeDialog()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="generateBtn">Generate Report</button>
      </div>
    </form>
  </div>

  <script>
    // Filter checkbox logic
    document.getElementById('filterBoth').addEventListener('change', function() {
      if (this.checked) {
        document.getElementById('filterIn').checked = false;
        document.getElementById('filterOut').checked = false;
      }
    });
    
    document.getElementById('filterIn').addEventListener('change', function() {
      if (this.checked || document.getElementById('filterOut').checked) {
        document.getElementById('filterBoth').checked = false;
      }
    });
    
    document.getElementById('filterOut').addEventListener('change', function() {
      if (this.checked || document.getElementById('filterIn').checked) {
        document.getElementById('filterBoth').checked = false;
      }
    });
    
    function showError(message) {
      document.getElementById('errorMessage').textContent = message;
    }
    
    function clearError() {
      document.getElementById('errorMessage').textContent = '';
    }
    
    function setLoading(loading) {
      document.getElementById('loadingMessage').style.display = loading ? 'block' : 'none';
      document.getElementById('generateBtn').disabled = loading;
    }
    
    function closeDialog() {
      google.script.host.close();
    }
    
    document.getElementById('historyForm').addEventListener('submit', function(e) {
      e.preventDefault();
      clearError();
      
      const studentIdentifier = document.getElementById('studentIdentifier').value.trim();
      const startDate = document.getElementById('startDate').value;
      const endDate = document.getElementById('endDate').value;
      const filterBoth = document.getElementById('filterBoth').checked;
      const filterIn = document.getElementById('filterIn').checked;
      const filterOut = document.getElementById('filterOut').checked;
      
      if (!studentIdentifier) {
        showError('Please enter a student name or ID');
        return;
      }
      
      if (!filterBoth && !filterIn && !filterOut) {
        showError('Please select at least one entry type filter');
        return;
      }
      
      if (startDate && endDate && new Date(startDate) > new Date(endDate)) {
        showError('Start date cannot be after end date');
        return;
      }
      
      const formData = {
        studentIdentifier: studentIdentifier,
        startDate: startDate || null,
        endDate: endDate || null,
        filterBoth: filterBoth,
        filterIn: filterIn,
        filterOut: filterOut
      };
      
      setLoading(true);
      
      google.script.run
        .withSuccessHandler(function(result) {
          setLoading(false);
          if (result.success) {
            alert('Report generated successfully!\\n\\n' +
                  'Sheet: "' + result.sheetName + '"\\n' +
                  'Total entries: ' + result.entryCount + '\\n' +
                  'Student: ' + result.studentName);
            closeDialog();
          } else {
            showError(result.error);
          }
        })
        .withFailureHandler(function(error) {
          setLoading(false);
          showError('Error generating report: ' + error.message);
        })
        .generateStudentHistoryReport(formData);
    });
  </script>
</body>
</html>
  `);
  
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(750)
    .setHeight(600)
    .setTitle('Student History Report');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Student History Report');
}

// DELETE OLD REPORTS
function deleteOldTardyReports() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Delete Old Reports',
    'This will delete ALL sheets that start with "Late Arrivals Report -". Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let deleted = 0;

  sheets.forEach(sheet => {
    if (sheet.getName().startsWith('Late Arrivals Report -')) {
      ss.deleteSheet(sheet);
      deleted++;
    }
  });

  ui.alert('Cleanup Complete', `Deleted ${deleted} old Late Arrivals report sheet(s).`, ui.ButtonSet.OK);
}

// EMAIL GENERATION DIALOG
function showEmailGenerationDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tardySheets = ss.getSheets().filter(sheet => 
    sheet.getName().startsWith('Late Arrivals Report -')
  );
  
  if (tardySheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Reports Found',
      'No Late Arrivals reports found. Please generate a Late Arrivals report first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const htmlTemplate = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Google Sans', Roboto, Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f8f9fa;
    }
    
    .container {
      max-width: 600px;
      background: white;
      border-radius: 12px;
      padding: 24px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.1);
    }
    
    .header {
      text-align: center;
      margin-bottom: 24px;
    }
    
    .title {
      color: #1a73e8;
      font-size: 24px;
      font-weight: 500;
      margin: 0;
    }
    
    .subtitle {
      color: #5f6368;
      font-size: 14px;
      margin: 8px 0 0 0;
    }
    
    .form-group {
      margin-bottom: 20px;
    }
    
    .form-label {
      display: block;
      color: #202124;
      font-size: 14px;
      font-weight: 500;
      margin-bottom: 8px;
    }
    
    .form-select, .form-input {
      width: 100%;
      padding: 12px;
      border: 2px solid #dadce0;
      border-radius: 8px;
      font-size: 16px;
      box-sizing: border-box;
      transition: border-color 0.2s;
    }
    
    .form-select:focus, .form-input:focus {
      outline: none;
      border-color: #1a73e8;
      box-shadow: 0 0 0 3px rgba(26,115,232,0.1);
    }
    
    .options-section {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 8px;
      margin: 20px 0;
    }
    
    .options-title {
      color: #202124;
      font-size: 16px;
      font-weight: 500;
      margin-bottom: 12px;
    }
    
    .checkbox-group {
      display: flex;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .checkbox-group input[type="checkbox"] {
      margin-right: 8px;
      transform: scale(1.2);
    }
    
    .checkbox-group label {
      color: #202124;
      font-size: 14px;
      cursor: pointer;
    }
    
    .input-group {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 8px;
    }
    
    .input-group input {
      flex: 1;
      padding: 8px 12px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      font-size: 14px;
    }
    
    .button-group {
      display: flex;
      gap: 12px;
      margin-top: 24px;
    }
    
    .btn {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
    }
    
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    
    .btn-primary:hover {
      background: #1557b2;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(26,115,232,0.3);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #5f6368;
      border: 1px solid #dadce0;
    }
    
    .btn-secondary:hover {
      background: #e8eaed;
    }
    
    .btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    
    .error {
      color: #d93025;
      font-size: 14px;
      margin-top: 8px;
    }
    
    .loading {
      display: none;
      text-align: center;
      color: #5f6368;
      font-size: 14px;
      margin-top: 16px;
    }
    
    .loading::before {
      content: '';
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid #dadce0;
      border-top-color: #1a73e8;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-right: 8px;
      vertical-align: middle;
    }
    
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
    
    .today-filter-section {
      background: #fff3cd;
      border: 1px solid #ffeaa7;
      padding: 16px;
      border-radius: 8px;
      margin: 20px 0;
    }
    
    .today-filter-section .options-title {
      color: #856404;
      margin-bottom: 8px;
    }
    
    .today-filter-section .checkbox-group label {
      color: #856404;
    }
    
    .today-filter-hint {
      color: #856404;
      font-size: 12px;
      margin-top: 4px;
      font-style: italic;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 class="title">Generate Late Arrivals Emails</h1>
      <p class="subtitle">Select a report and configure email settings</p>
    </div>
    
    <form id="emailForm">
      <div class="form-group">
        <label class="form-label" for="reportSheet">Select Late Arrivals Report</label>
        <select id="reportSheet" class="form-select" required>
          <option value="">-- Select a Report --</option>
          ${tardySheets.map(sheet => 
            `<option value="${sheet.getName()}">${sheet.getName()}</option>`
          ).join('')}
        </select>
      </div>
      
      <!-- TODAY ONLY FILTER SECTION -->
      <div class="today-filter-section">
        <div class="options-title">Date Filter</div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="todayOnly" checked>
          <label for="todayOnly">Only send emails for students with tardies TODAY</label>
        </div>
        <div class="today-filter-hint">
          When checked, only students who were Late Arrivals today will receive emails, even if they had other tardies in the report period.
        </div>
      </div>
      
      <div class="options-section">
        <div class="options-title">Email Settings</div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="useThreshold" onchange="toggleThresholdInput()" checked>
          <label for="useThreshold">Use Late Arrivals count threshold for email type</label>
        </div>
        <div class="input-group" id="thresholdGroup">
          <label>Threshold for strong warning email:</label>
          <input type="number" id="thresholdValue" value="10" min="2" max="50" class="form-input">
        </div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="previewMode" checked>
          <label for="previewMode">Preview mode (generate document only)</label>
        </div>
        
        <div class="checkbox-group">
          <input type="checkbox" id="sendEmails">
          <label for="sendEmails">Actually send emails via Gmail</label>
        </div>
      </div>
      
      <div id="errorMessage" class="error"></div>
      <div id="loadingMessage" class="loading">Generating emails...</div>
      
      <div class="button-group">
        <button type="button" class="btn btn-secondary" onclick="closeDialog()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="generateBtn">Generate Emails</button>
      </div>
    </form>
  </div>

  <script>
    function showError(message) {
      document.getElementById('errorMessage').textContent = message;
    }
    
    function clearError() {
      document.getElementById('errorMessage').textContent = '';
    }
    
    function setLoading(loading) {
      document.getElementById('loadingMessage').style.display = loading ? 'block' : 'none';
      document.getElementById('generateBtn').disabled = loading;
    }
    
    function closeDialog() {
      google.script.host.close();
    }
    
    function toggleThresholdInput() {
      const checkbox = document.getElementById('useThreshold');
      const group = document.getElementById('thresholdGroup');
      group.style.display = checkbox.checked ? 'flex' : 'none';
    }
    
    document.getElementById('emailForm').addEventListener('submit', function(e) {
      e.preventDefault();
      clearError();
      
      const reportSheet = document.getElementById('reportSheet').value;
      
      if (!reportSheet) {
        showError('Please select a Late Arrivals report');
        return;
      }
      
      const formData = {
        reportSheetName: reportSheet,
        useThreshold: document.getElementById('useThreshold').checked,
        thresholdValue: parseInt(document.getElementById('thresholdValue').value) || 5,
        previewMode: document.getElementById('previewMode').checked,
        sendEmails: document.getElementById('sendEmails').checked,
        todayOnly: document.getElementById('todayOnly').checked
      };
      
      setLoading(true);
      
      google.script.run
        .withSuccessHandler(function(result) {
          setLoading(false);
          if (result.success) {
            document.getElementById('emailForm').style.display = 'none';
            
            const container = document.querySelector('.container');
            const successDiv = document.createElement('div');
            successDiv.style.textAlign = 'center';
            successDiv.style.padding = '20px';
            
            const title = document.createElement('h2');
            title.textContent = 'Emails Generated Successfully!';
            title.style.color = '#1a73e8';
            title.style.marginBottom = '20px';
            successDiv.appendChild(title);
            
            const infoBox = document.createElement('div');
            infoBox.style.background = '#f8f9fa';
            infoBox.style.padding = '20px';
            infoBox.style.borderRadius = '8px';
            infoBox.style.marginBottom = '20px';
            
            const docInfo = document.createElement('p');
            docInfo.innerHTML = '<strong>Document:</strong> ' + result.documentName;
            infoBox.appendChild(docInfo);
            
            const emailInfo = document.createElement('p');
            emailInfo.innerHTML = '<strong>Emails to send:</strong> ' + result.emailCount;
            infoBox.appendChild(emailInfo);
            
            const skippedInfo = document.createElement('p');
            skippedInfo.innerHTML = '<strong>Students skipped (recently emailed):</strong> ' + result.skippedCount;
            infoBox.appendChild(skippedInfo);
            
            if (result.todayFilterUsed) {
              const todayInfo = document.createElement('p');
              todayInfo.innerHTML = '<strong>Today-only filter:</strong> ' + (result.todayOnlyCount || 0) + ' students had tardies today';
              todayInfo.style.color = '#856404';
              infoBox.appendChild(todayInfo);
            }
            
            if (result.sentCount) {
              const sentInfo = document.createElement('p');
              sentInfo.innerHTML = '<strong>Emails sent:</strong> ' + result.sentCount;
              infoBox.appendChild(sentInfo);
            }
            
            successDiv.appendChild(infoBox);
            
            const linkContainer = document.createElement('div');
            linkContainer.style.margin = '20px 0';
            
            const docLink = document.createElement('a');
            docLink.href = result.documentUrl;
            docLink.target = '_blank';
            docLink.textContent = 'Open Email Document';
            docLink.style.display = 'inline-block';
            docLink.style.background = '#1a73e8';
            docLink.style.color = 'white';
            docLink.style.padding = '12px 24px';
            docLink.style.textDecoration = 'none';
            docLink.style.borderRadius = '8px';
            docLink.style.fontWeight = '500';
            docLink.style.fontSize = '16px';
            docLink.style.margin = '10px';
            
            linkContainer.appendChild(docLink);
            successDiv.appendChild(linkContainer);
            
            const instructions = document.createElement('p');
            instructions.innerHTML = 'The document includes both emails to send and skipped students.<br>Use the Email Status Tracker to mark emails as sent after sending them.';
            instructions.style.color = '#5f6368';
            instructions.style.fontSize = '14px';
            instructions.style.marginTop = '20px';
            successDiv.appendChild(instructions);
            
            const closeBtn = document.createElement('button');
            closeBtn.textContent = 'Close';
            closeBtn.onclick = closeDialog;
            closeBtn.style.background = '#f8f9fa';
            closeBtn.style.color = '#5f6368';
            closeBtn.style.border = '1px solid #dadce0';
            closeBtn.style.padding = '10px 20px';
            closeBtn.style.borderRadius = '6px';
            closeBtn.style.cursor = 'pointer';
            closeBtn.style.marginTop = '20px';
            successDiv.appendChild(closeBtn);
            
            container.appendChild(successDiv);
          } else {
            showError(result.error);
          }
        })
        .withFailureHandler(function(error) {
          setLoading(false);
          showError('Error generating emails: ' + error.message);
        })
        .generateEmailsFromReportWithNewTracker(formData);
    });
  </script>
</body>
</html>
  `);
  
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(650)
    .setHeight(750)
    .setTitle('Generate Late Arrivals Emails');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email Generation');
}

// MANUAL TRACKER DIALOG
function showManualTrackerDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tardySheets = ss.getSheets().filter(sheet => 
    sheet.getName().startsWith('Late Arrivals Report -')
  );
  
  if (tardySheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Reports Found',
      'No Late Arrivals reports found. Please generate a Late Arrivals report first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const htmlTemplate = HtmlService.createTemplate(`
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: 'Google Sans', Roboto, Arial, sans-serif;
      padding: 20px;
      margin: 0;
      background: #f8f9fa;
    }
    
    .container {
      max-width: 500px;
      background: white;
      border-radius: 12px;
      padding: 24px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.1);
    }
    
    .header {
      text-align: center;
      margin-bottom: 24px;
    }
    
    .title {
      color: #34a853;
      font-size: 24px;
      font-weight: 500;
      margin: 0;
    }
    
    .subtitle {
      color: #5f6368;
      font-size: 14px;
      margin: 8px 0 0 0;
    }
    
    .form-group {
      margin-bottom: 20px;
    }
    
    .form-label {
      display: block;
      color: #202124;
      font-size: 14px;
      font-weight: 500;
      margin-bottom: 8px;
    }
    
    .form-select, .form-input, .form-textarea {
      width: 100%;
      padding: 12px;
      border: 2px solid #dadce0;
      border-radius: 8px;
      font-size: 16px;
      box-sizing: border-box;
      transition: border-color 0.2s;
      font-family: 'Google Sans', Roboto, Arial, sans-serif;
    }
    
    .form-textarea {
      min-height: 80px;
      resize: vertical;
    }
    
    .form-select:focus, .form-input:focus, .form-textarea:focus {
      outline: none;
      border-color: #34a853;
      box-shadow: 0 0 0 3px rgba(52,168,83,0.1);
    }
    
    .options-section {
      background: #f8f9fa;
      padding: 16px;
      border-radius: 8px;
      margin: 20px 0;
    }
    
    .options-title {
      color: #202124;
      font-size: 16px;
      font-weight: 500;
      margin-bottom: 12px;
    }
    
    .radio-group {
      display: flex;
      align-items: center;
      margin-bottom: 12px;
    }
    
    .radio-group input[type="radio"] {
      margin-right: 8px;
      transform: scale(1.2);
    }
    
    .radio-group label {
      color: #202124;
      font-size: 14px;
      cursor: pointer;
    }
    
    .button-group {
      display: flex;
      gap: 12px;
      margin-top: 24px;
    }
    
    .btn {
      flex: 1;
      padding: 12px 24px;
      border: none;
      border-radius: 8px;
      font-size: 16px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s;
    }
    
    .btn-primary {
      background: #34a853;
      color: white;
    }
    
    .btn-primary:hover {
      background: #2d8f47;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(52,168,83,0.3);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #5f6368;
      border: 1px solid #dadce0;
    }
    
    .btn-secondary:hover {
      background: #e8eaed;
    }
    
    .btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    
    .error {
      color: #d93025;
      font-size: 14px;
      margin-top: 8px;
    }
    
    .hint {
      color: #5f6368;
      font-size: 12px;
      margin-top: 4px;
    }
    
    .loading {
      display: none;
      text-align: center;
      color: #5f6368;
      font-size: 14px;
      margin-top: 16px;
    }
    
    .loading::before {
      content: '';
      display: inline-block;
      width: 16px;
      height: 16px;
      border: 2px solid #dadce0;
      border-top-color: #34a853;
      border-radius: 50%;
      animation: spin 1s linear infinite;
      margin-right: 8px;
      vertical-align: middle;
    }
    
    @keyframes spin {
      to { transform: rotate(360deg); }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 class="title">Add Report to Tracker</h1>
      <p class="subtitle">Manually add students from a Late Arrivals report to the Email Status Tracker</p>
    </div>
    
    <form id="trackerForm">
      <div class="form-group">
        <label class="form-label" for="reportSheet">Select Late Arrivals Report</label>
        <select id="reportSheet" class="form-select" required>
          <option value="">-- Select a Report --</option>
          ${tardySheets.map(sheet => 
            `<option value="${sheet.getName()}">${sheet.getName()}</option>`
          ).join('')}
        </select>
        <div class="hint">Choose the report containing students you want to add to the tracker</div>
      </div>
      
      <div class="options-section">
        <div class="options-title">Entry Status</div>
        
        <div class="radio-group">
          <input type="radio" id="statusPending" name="status" value="pending" checked>
          <label for="statusPending">Pending (emails not yet sent)</label>
        </div>
        
        <div class="radio-group">
          <input type="radio" id="statusSent" name="status" value="sent">
          <label for="statusSent">Sent (emails already sent)</label>
        </div>
        
        <div class="radio-group">
          <input type="radio" id="statusSkipped" name="status" value="skipped">
          <label for="statusSkipped">Skipped (no emails sent)</label>
        </div>
      </div>
      
      <div class="form-group">
        <label class="form-label" for="customNotes">Notes (Optional)</label>
        <textarea id="customNotes" class="form-textarea" placeholder="Add any additional notes about this entry..."></textarea>
        <div class="hint">These notes will appear in the tracker for reference</div>
      </div>
      
      <div id="errorMessage" class="error"></div>
      <div id="loadingMessage" class="loading">Adding students to tracker...</div>
      
      <div class="button-group">
        <button type="button" class="btn btn-secondary" onclick="closeDialog()">Cancel</button>
        <button type="submit" class="btn btn-primary" id="addBtn">Add to Tracker</button>
      </div>
    </form>
  </div>

  <script>
    function showError(message) {
      document.getElementById('errorMessage').textContent = message;
    }
    
    function clearError() {
      document.getElementById('errorMessage').textContent = '';
    }
    
    function setLoading(loading) {
      document.getElementById('loadingMessage').style.display = loading ? 'block' : 'none';
      document.getElementById('addBtn').disabled = loading;
    }
    
    function closeDialog() {
      google.script.host.close();
    }
    
    document.getElementById('trackerForm').addEventListener('submit', function(e) {
      e.preventDefault();
      clearError();
      
      const reportSheet = document.getElementById('reportSheet').value;
      const statusRadios = document.getElementsByName('status');
      let selectedStatus = '';
      
      for (const radio of statusRadios) {
        if (radio.checked) {
          selectedStatus = radio.value;
          break;
        }
      }
      
      const customNotes = document.getElementById('customNotes').value.trim();
      
      if (!reportSheet) {
        showError('Please select a Late Arrivals report');
        return;
      }
      
      const formData = {
        reportSheetName: reportSheet,
        status: selectedStatus,
        customNotes: customNotes
      };
      
      setLoading(true);
      
      google.script.run
        .withSuccessHandler(function(result) {
          setLoading(false);
          if (result.success) {
            document.getElementById('trackerForm').style.display = 'none';
            
            const container = document.querySelector('.container');
            const successDiv = document.createElement('div');
            successDiv.style.textAlign = 'center';
            successDiv.style.padding = '20px';
            
            const title = document.createElement('h2');
            title.textContent = 'Students Added Successfully!';
            title.style.color = '#34a853';
            title.style.marginBottom = '20px';
            successDiv.appendChild(title);
            
            const infoBox = document.createElement('div');
            infoBox.style.background = '#f8f9fa';
            infoBox.style.padding = '20px';
            infoBox.style.borderRadius = '8px';
            infoBox.style.marginBottom = '20px';
            
            const reportInfo = document.createElement('p');
            reportInfo.innerHTML = '<strong>Report:</strong> ' + result.reportName;
            infoBox.appendChild(reportInfo);
            
            const studentInfo = document.createElement('p');
            studentInfo.innerHTML = '<strong>Students Added:</strong> ' + result.studentsAdded;
            infoBox.appendChild(studentInfo);
            
            const statusInfo = document.createElement('p');
            statusInfo.innerHTML = '<strong>Status:</strong> ' + result.status;
            infoBox.appendChild(statusInfo);
            
            if (result.notes) {
              const notesInfo = document.createElement('p');
              notesInfo.innerHTML = '<strong>Notes:</strong> ' + result.notes;
              infoBox.appendChild(notesInfo);
            }
            
            successDiv.appendChild(infoBox);
            
            const instructions = document.createElement('p');
            instructions.textContent = 'Students have been added to the Email Status Tracker. You can view them using "Open Email Status Tracker" from the menu.';
            instructions.style.color = '#5f6368';
            instructions.style.fontSize = '14px';
            instructions.style.marginTop = '20px';
            successDiv.appendChild(instructions);
            
            const closeBtn = document.createElement('button');
            closeBtn.textContent = 'Close';
            closeBtn.onclick = closeDialog;
            closeBtn.style.background = '#f8f9fa';
            closeBtn.style.color = '#5f6368';
            closeBtn.style.border = '1px solid #dadce0';
            closeBtn.style.padding = '10px 20px';
            closeBtn.style.borderRadius = '6px';
            closeBtn.style.cursor = 'pointer';
            closeBtn.style.marginTop = '20px';
            successDiv.appendChild(closeBtn);
            
            container.appendChild(successDiv);
          } else {
            showError(result.error);
          }
        })
        .withFailureHandler(function(error) {
          setLoading(false);
          showError('Error adding students to tracker: ' + error.message);
        })
        .addReportToTracker(formData);
    });
  </script>
</body>
</html>
  `);
  
  const htmlOutput = htmlTemplate.evaluate()
    .setWidth(550)
    .setHeight(650)
    .setTitle('Add Report to Email Status Tracker');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Manual Tracker Entry');
}

// TRACKER MANAGEMENT DIALOGS
function showMarkEmailsSentDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Mark Emails as Sent',
    'Enter the report date range to mark as sent\n(e.g., "12/01/2024 to 12/15/2024"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const reportRange = response.getResponseText().trim();
    const count = markEmailsAsSentInTracker(reportRange);
    ui.alert('Success', `Marked ${count} emails as sent for: ${reportRange}`, ui.ButtonSet.OK);
  }
}

function showEmailTrackerManager() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Email Tracker Manager',
    'What would you like to do?\n\n' +
    'YES - View Activity Summary\n' +
    'NO - Mark Recent Emails as Sent\n' +
    'CANCEL - Open Tracker Sheet',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    showEmailActivitySummary();
  } else if (response === ui.Button.NO) {
    showMarkEmailsSentDialog();
  } else if (response === ui.Button.CANCEL) {
    openEmailStatusTracker();
  }
}

function showNeverSendManager() {
  const trackerSheet = getEmailStatusTracker();
  const data = trackerSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert(
      'No Data',
      'The Email Status Tracker is empty. Generate some reports first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const neverSendStudents = new Map();
  
  for (let i = 1; i < data.length; i++) {
    const [studentId, studentName, reportRange, dateGenerated, emailSent, dateSent, override, neverSend, notes] = data[i];
    
    if (neverSend === true) {
      neverSendStudents.set(studentId, {
        id: studentId,
        name: studentName
      });
    }
  }
  
  let message = 'STUDENTS MARKED "NEVER SEND"\n\n';
  
  if (neverSendStudents.size === 0) {
    message += 'No students are currently marked as "Never Send".\n\n';
  } else {
    message += `Total: ${neverSendStudents.size} student(s)\n\n`;
    Array.from(neverSendStudents.values()).forEach(student => {
      message += `• ${student.name} (${student.id})\n`;
    });
    message += '\n';
  }
  
  message += 'To add or remove students from the "Never Send" list:\n';
  message += '1. Open the Email Status Tracker sheet\n';
  message += '2. Find the student\'s row\n';
  message += '3. Check the "Never Send" checkbox to prevent emails\n';
  message += '4. Uncheck to allow emails again\n\n';
  message += 'Note: The most recent "Never Send" status for each student is used.';
  
  SpreadsheetApp.getUi().alert('Never Send List', message, SpreadsheetApp.getUi().ButtonSet.OK);
}
