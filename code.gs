function onFormSubmit(e) {
  // Get the submitted responses from the event object
  var responses = e.values;

  // Extract and validate responses
  var employeeName = responses[1] || 'N/A'; // Employee Name
  var employeeEmail = responses[2] || 'N/A'; // Employee Email
  var leaveType = responses[3] || 'N/A'; // Leave Type
  var leaveStartDate = parseDate(responses[4]); // Leave Start Date
  var leaveEndDate = parseDate(responses[5]); // Leave End Date
  var reason = responses[6] || 'N/A'; // Reason

  // Calculate leave days
  var leaveDays = calculateLeaveDays(leaveStartDate, leaveEndDate);

  // Default approval status (can be updated later)
  var approvalStatus = 'Pending';

  // Generate a random Leave ID
  var leaveId = generateRandomLeaveId();

  // Update Sheet
  updateSheet(e.range, leaveId, approvalStatus, leaveDays);

  var sheeturl = 'https://docs.google.com/spreadsheets/d/1WVxHbl3p_hzn_77IasYFAyIJiwV0g6DAKU7DxVbxPOw/edit?usp=sharing';

  // Send email notification to HR
  var hremail = 'terip99864@janfab.com';
  var hrsubject = 'New Leave Request from ' + employeeName;
  var hrbody = '<h2>New Leave Request</h2>' +
               '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
               '<p><strong>Start Date:</strong> ' + formatDate(leaveStartDate) + '</p>' +
               '<p><strong>End Date:</strong> ' + formatDate(leaveEndDate) + '</p>' +
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

// Function to parse date strings
function parseDate(dateString) {
  var date = new Date(dateString);
  return isNaN(date.getTime()) ? 'Invalid Date' : date;
}

// Function to format date as a readable string
function formatDate(date) {
  if (date instanceof Date && !isNaN(date.getTime())) {
    return date.toDateString(); // Format as readable string
  }
  return 'Invalid Date';
}

// Function to calculate leave days
function calculateLeaveDays(startDate, endDate) {
  if (!(startDate instanceof Date) || isNaN(startDate.getTime()) ||
      !(endDate instanceof Date) || isNaN(endDate.getTime())) {
    return 'Invalid Dates';
  }

  var oneDay = 24 * 60 * 60 * 1000; // milliseconds in a day
  var diffDays = Math.round(Math.abs((endDate - startDate) / oneDay)) + 1; // inclusive of both start and end date
  
  return diffDays;
}

// Function to generate a random Leave ID
function generateRandomLeaveId() {
  return Math.floor(Math.random() * 1000); // Generates a random number between 0 and 999
}

// Function to update the Google Sheet
function updateSheet(range, leaveId, approvalStatus, leaveDays) {
  var sheet = range.getSheet();
  var row = range.getRow();
  
  // Update Leave ID, Approval Status, Leave Days
  sheet.getRange(row, 2).setValue(leaveId); // Leave ID
  sheet.getRange(row, 9).setValue(leaveDays); // Leave Days
  sheet.getRange(row, 10).setValue('Pending'); // HR Approval
  sheet.getRange(row, 11).setValue(''); // HR Justification
  sheet.getRange(row, 12).setValue(approvalStatus); // Approval Status
  sheet.getRange(row, 13).setValue(''); // Last Modified (will be set in onEdit)
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var editedRow = range.getRow();
  var editedColumn = range.getColumn();
  
  // Check if the edited column is the HR Approval column (column 10)
  var hrApprovalColumn = 10;
  if (editedColumn === hrApprovalColumn) {
    var hrApproval = range.getValue();
    var data = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Extract employee email and other details from the row
    var employeeName = data[2]; // Employee Name
    var employeeEmail = data[3]; // Employee Email
    var leaveType = data[4]; // Leave Type
    var leaveStartDate = data[5]; // Leave Start Date
    var leaveEndDate = data[6]; // Leave End Date
    var reason = data[7]; // Reason
    var leaveDays = data[8]; // Leave Days
    var leaveId = data[1]; // Leave ID
    var hrJustification = data[10]; // HR Justification

    // Send email notification to the supervisor if HR Approval is 'Approved'
    if (hrApproval === 'Approved') {
      var supervisorEmail = 'vidiv39862@ploncy.com'; // Replace with actual supervisor email
      var subject = 'Leave Request Approved';
      var body = '<p>Dear Supervisor,</p>' +
                 '<p>The following leave request has been approved by HR:</p>' +
                 '<ul>' +
                 '<li><strong>Leave ID:</strong> ' + leaveId + '</li>' +
                 '<li><strong>Employee Name:</strong> ' + employeeName + '</li>' +
                 '<li><strong>Employee Email:</strong> ' + employeeEmail + '</li>' +
                 '<li><strong>Leave Type:</strong> ' + leaveType + '</li>' +
                 '<li><strong>Start Date:</strong> ' + new Date(leaveStartDate).toDateString() + '</li>' +
                 '<li><strong>End Date:</strong> ' + new Date(leaveEndDate).toDateString() + '</li>' +
                 '<li><strong>Reason:</strong> ' + reason + '</li>' +
                 '<li><strong>Leave Days:</strong> ' + leaveDays + '</li>' +
                 '<li><strong>HR Justification:</strong> ' + hrJustification + '</li>' +
                 '</ul>' +
                 '<p>Thank you,</p>' +
                 '<p>HR Team</p>';
      MailApp.sendEmail({
        to: supervisorEmail,
        subject: subject,
        htmlBody: body
      });
    }

    // Update "Last Modified By" column (column 13)
    var lastModifiedColumn = 13;
    var userEmail = Session.getActiveUser().getEmail();
    sheet.getRange(editedRow, lastModifiedColumn).setValue(userEmail);
  }
}
