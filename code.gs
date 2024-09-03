function onFormSubmit(e) {
  // Get the submitted responses from the event object
  var responses = e.values;

  // Extract the responses into variables
  var employeeName = responses[1];
  var employeeEmail = responses[2];
  var leaveType = responses[3];
  var leaveStartDate = new Date(responses[4]);
  var leaveEndDate = new Date(responses[5]);
  var reason = responses[6];

  // Calculate leave days
  var leaveDays = calculateLeaveDays(leaveStartDate, leaveEndDate);

  // Default approval status (can be updated later)
  var approvalStatus = 'Pending';

  // Generate a random Leave ID
  var leaveId = generateRandomLeaveId();

  // Log the data (for debugging purposes)
  Logger.log('Employee Name: ' + employeeName);
  Logger.log('Employee Email: ' + employeeEmail);
  Logger.log('Leave Type: ' + leaveType);
  Logger.log('Leave Start Date: ' + leaveStartDate);
  Logger.log('Leave End Date: ' + leaveEndDate);
  Logger.log('Reason: ' + reason);
  Logger.log('Leave Days: ' + leaveDays);
  Logger.log('Approval Status: ' + approvalStatus);
  Logger.log('Leave ID: ' + leaveId);

  // Example: Send email notification (adjust as needed)
  var subject = 'New Leave Request from ' + employeeName;
  var body = 'Leave Type: ' + leaveType + '\n' +
             'Start Date: ' + leaveStartDate.toDateString() + '\n' +
             'End Date: ' + leaveEndDate.toDateString() + '\n' +
             'Reason: ' + reason + '\n' +
             'Leave Days: ' + leaveDays + '\n' +
             'Leave ID: ' + leaveId;

  MailApp.sendEmail('bibatib646@avashost.com', subject, body);

  // Optionally, you can update the Google Sheet with the Leave ID and Approval Status
  updateSheet(e.range, leaveId, approvalStatus,leaveDays);
}

// Function to calculate leave days
function calculateLeaveDays(startDate, endDate) {
  var oneDay = 24 * 60 * 60 * 1000; // milliseconds in a day
  return Math.round(Math.abs((endDate - startDate) / oneDay)) + 1; // inclusive of both start and end date
}

// Function to generate a random Leave ID
function generateRandomLeaveId() {
  return Math.floor(Math.random() * 1000000); // Generates a random number between 0 and 999999
}

// Function to update the Google Sheet with Leave ID and Approval Status
function updateSheet(range, leaveId, approvalStatus,leaveDays) {
  var sheet = range.getSheet();
  var row = range.getRow();
  
  // Assume Leave ID is in the 9th column and Approval Status is in the 10th column
  sheet.getRange(row, 8).setValue(leaveDays)
  sheet.getRange(row, 9).setValue(approvalStatus);
  sheet.getRange(row, 10).setValue(leaveId);
  
}


