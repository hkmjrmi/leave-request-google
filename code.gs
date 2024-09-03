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

  // Update Sheet
  updateSheet(e.range, leaveId, approvalStatus, leaveDays);

  var sheeturl = 'https://docs.google.com/spreadsheets/d/1WVxHbl3p_hzn_77IasYFAyIJiwV0g6DAKU7DxVbxPOw/edit?usp=sharing';

  // Example: Send email notification to HR
  var hremail = 'terip99864@janfab.com';
  var hrsubject = 'New Leave Request from ' + employeeName;
  var hrbody = '<h2>New Leave Request</h2>' +
               '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
               '<p><strong>Start Date:</strong> ' + leaveStartDate.toDateString() + '</p>' +
               '<p><strong>End Date:</strong> ' + leaveEndDate.toDateString() + '</p>' +
               '<p><strong>Reason:</strong> ' + reason + '</p>' +
               '<p><strong>Leave Days:</strong> ' + leaveDays + '</p>' +
               '<p><strong>Leave ID:</strong> ' + leaveId + '</p>' +
               '<p><a href="' + sheeturl + '">Approve or Reject the Leave Request</a></p>';

  MailApp.sendEmail({
    to: hremail,
    subject: hrsubject,
    htmlBody: hrbody
  });

  // Example: Send email notification to employee
  var subject = 'New Leave Request Submitted';
  var body = '<p>Dear ' + employeeName + ',</p>' +
             '<p>Your leave request has been submitted. Please wait for approval.</p>' +
             '<p>Thank you,</p>' +
             '<p>HR Team</p>';

  MailApp.sendEmail({
    to: employeeEmail,
    subject: subject,
    htmlBody: body
  });
}

// Function to calculate leave days
function calculateLeaveDays(startDate, endDate) {
  var oneDay = 24 * 60 * 60 * 1000; // milliseconds in a day
  return Math.round(Math.abs((endDate - startDate) / oneDay)) + 1; // inclusive of both start and end date
}

// Function to generate a random Leave ID
function generateRandomLeaveId() {
  return Math.floor(Math.random() * 100); // Generates a random number between 0 and 999999
}

// Function to update the Google Sheet
function updateSheet(range, leaveId, approvalStatus, leaveDays) {
  var sheet = range.getSheet();
  var row = range.getRow();
  
  // Assume Leave ID is in the 9th column and Approval Status is in the 10th column
  sheet.getRange(row, 8).setValue(leaveDays);
  sheet.getRange(row, 9).setValue(approvalStatus);
  sheet.getRange(row, 10).setValue(leaveId);
}

function onEdit(e) {
  // Get the sheet, range, and values related to the edit event
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var editedRow = range.getRow();
  var editedColumn = range.getColumn();
  
  // Check if the edited column is the Approval Status column (column 9)
  var approvalStatusColumn = 9; // Adjust the index if Approval Status is in a different column
  if (editedColumn === approvalStatusColumn) {
    var approvalStatus = range.getValue();
    var data = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Extract employee email and other details from the row
    var employeeName = data[1]; // Assuming Employee Name is in column 2
    var employeeEmail = data[2]; // Assuming Employee Email is in column 3
    var leaveType = data[3]; // Assuming Leave Type is in column 4
    var leaveStartDate = data[4]; // Assuming Leave Start Date is in column 5
    var leaveEndDate = data[5]; // Assuming Leave End Date is in column 6
    var reason = data[6]; // Assuming Reason is in column 7
    var leaveDays = data[7]; // Assuming Leave Days is in column 8
    var leaveId = data[9]; // Assuming Leave ID is in column 10

    // Send email notification to the employee
    sendApprovalStatusEmail(employeeEmail, employeeName, approvalStatus, leaveType, leaveStartDate, leaveEndDate, reason, leaveDays, leaveId);
  }
}

function sendApprovalStatusEmail(employeeEmail, employeeName, approvalStatus, leaveType, leaveStartDate, leaveEndDate, reason, leaveDays, leaveId) {
  var subject = 'Leave Request Status Updated';
  var body = '<p>Dear ' + employeeName + ',</p>' +
             '<p>Your leave request has been updated. Please find the details below:</p>' +
             '<ul>' +
             '<li><strong>Leave Type:</strong> ' + leaveType + '</li>' +
             '<li><strong>Start Date:</strong> ' + new Date(leaveStartDate).toDateString() + '</li>' +
             '<li><strong>End Date:</strong> ' + new Date(leaveEndDate).toDateString() + '</li>' +
             '<li><strong>Reason:</strong> ' + reason + '</li>' +
             '<li><strong>Leave Days:</strong> ' + leaveDays + '</li>' +
             '<li><strong>Leave ID:</strong> ' + leaveId + '</li>' +
             '<li><strong>Approval Status:</strong> ' + approvalStatus + '</li>' +
             '</ul>' +
             '<p>Thank you,</p>' +
             '<p>HR Team</p>';

  MailApp.sendEmail({
    to: employeeEmail,
    subject: subject,
    htmlBody: body
  });
}
