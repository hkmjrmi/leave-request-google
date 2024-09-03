function onFormSubmit(e) {
  // Get the submitted responses from the event object
  var responses = e.values;

  // Extract the responses into variables
  var employeeName = responses[1]; // Employee Name
  var employeeEmail = responses[2]; // Employee Email
  var leaveType = responses[3]; // Leave Type
  var leaveStartDate = new Date(responses[4]); // Leave Start Date
  var leaveEndDate = new Date(responses[5]); // Leave End Date
  var reason = responses[6]; // Reason

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
  return Math.floor(Math.random() * 1000); // Generates a random number between 0 and 999
}

// Function to update the Google Sheet
function updateSheet(range, leaveId, approvalStatus, leaveDays) {
  var sheet = range.getSheet();
  var row = range.getRow();
  
  // Update Leave ID, Approval Status, Leave Days
  sheet.getRange(row, 8).setValue(leaveDays);
  sheet.getRange(row, 9).setValue(approvalStatus);
  sheet.getRange(row, 10).setValue(leaveId);
  sheet.getRange(row, 12).setValue('Pending'); // HR Approval
  sheet.getRange(row, 13).setValue(''); // HR Justification
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var editedRow = range.getRow();
  var editedColumn = range.getColumn();
  
  // Check if the edited column is the HR Approval column (column 12)
  var hrApprovalColumn = 12;
  if (editedColumn === hrApprovalColumn) {
    var hrApproval = range.getValue();
    var data = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Extract employee email and other details from the row
    var employeeName = data[1]; // Employee Name
    var employeeEmail = data[2]; // Employee Email
    var leaveType = data[3]; // Leave Type
    var leaveStartDate = data[4]; // Leave Start Date
    var leaveEndDate = data[5]; // Leave End Date
    var reason = data[6]; // Reason
    var leaveDays = data[7]; // Leave Days
    var leaveId = data[9]; // Leave ID
    var hrJustification = data[12]; // HR Justification

    // Send email notification to the supervisor if HR Approval is 'Approved'
    if (hrApproval === 'Approved') {
      var supervisorEmail = 'vidiv39862@ploncy.com'; // Replace with actual supervisor email
      var subject = 'Leave Request Approved';
      var body = '<p>Dear Supervisor,</p>' +
                 '<p>The following leave request has been approved by HR:</p>' +
                 '<ul>' +
                 '<li><strong>Employee Name:</strong> ' + employeeName + '</li>' +
                 '<li><strong>Leave Type:</strong> ' + leaveType + '</li>' +
                 '<li><strong>Start Date:</strong> ' + new Date(leaveStartDate).toDateString() + '</li>' +
                 '<li><strong>End Date:</strong> ' + new Date(leaveEndDate).toDateString() + '</li>' +
                 '<li><strong>Reason:</strong> ' + reason + '</li>' +
                 '<li><strong>Leave Days:</strong> ' + leaveDays + '</li>' +
                 '<li><strong>Leave ID:</strong> ' + leaveId + '</li>' +
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

    // Assume "Last Modified By" is in the 11th column
    var lastModifiedColumn = 11;
    var userEmail = Session.getActiveUser().getEmail();
    sheet.getRange(editedRow, lastModifiedColumn).setValue(userEmail);
  }
}
