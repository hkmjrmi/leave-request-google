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

  // Update Sheet
  updateSheet(e.range, leaveId, approvalStatus, leaveDays);

  updateLeaveBalanceSheet(employeeName, employeeEmail, leaveType, leaveDays);

  var sheeturl = 'https://docs.google.com/spreadsheets/d/1WVxHbl3p_hzn_77IasYFAyIJiwV0g6DAKU7DxVbxPOw/edit?usp=sharing';

  // Send email notification to HR
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

  // Send email notification to employee
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
  //  Get the sheet, range, and values related
  var sheet = range.getSheet();
  var row = range.getRow();
  
  // Set value for sheet and column
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
    var employeeName = data[1]; // Employee Name is in column 2
    var employeeEmail = data[2]; // Employee Email is in column 3
    var leaveType = data[3]; // Leave Type is in column 4
    var leaveStartDate = data[4]; // Leave Start Date is in column 5
    var leaveEndDate = data[5]; // Leave End Date is in column 6
    var reason = data[6]; // Reason is in column 7
    var leaveDays = data[7]; // Leave Days is in column 8
    var leaveId = data[9]; // Leave ID is in column 10

    // Send email notification to the employee
    sendApprovalStatusEmail(employeeEmail, employeeName, approvalStatus, leaveType, leaveStartDate, leaveEndDate, reason, leaveDays, leaveId);

    var userEmail = Session.getActiveUser().getEmail();
  
    // Assume "Last Modified By" is in the 11th column
    var lastModifiedColumn = 11; 
    sheet.getRange(editedRow, lastModifiedColumn).setValue(userEmail);
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

function updateLeaveBalanceSheet(employeeName, employeeEmail, leaveType, leaveDays) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leaveBalanceSheet = ss.getSheetByName('Leave Balances'); // Change to your sheet name
  var data = leaveBalanceSheet.getDataRange().getValues();
  
  // Find the row with the matching Employee Name and Email
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === employeeName && data[i][1] === employeeEmail) {
      var leaveTypeIndex = -1;
      
      // Find the column for the leave type
      for (var j = 2; j < data[i].length; j++) {
        if (data[0][j] === leaveType) { // Assuming header contains Leave Types
          leaveTypeIndex = j;
          break;
        }
      }
      
      if (leaveTypeIndex !== -1) {
        var leaveTotal = data[i][leaveTypeIndex];
        var leaveUsed = data[i][leaveTypeIndex + 1] || 0; // Assuming Leave Used is the next column
        var leaveBalance = leaveTotal - (leaveUsed + leaveDays);

        // Update Leave Used and Leave Balance
        leaveBalanceSheet.getRange(i + 1, leaveTypeIndex + 1).setValue(leaveUsed + leaveDays);
        leaveBalanceSheet.getRange(i + 1, leaveTypeIndex + 2).setValue(leaveBalance); // Assuming Leave Balance is the next column
      } else {
        Logger.log('Leave Type column not found for: ' + leaveType);
      }
      
      return; // Exit after updating the row
    }
  }

  Logger.log('Employee not found: ' + employeeName + ', ' + employeeEmail);
}

