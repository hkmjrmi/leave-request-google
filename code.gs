function onFormSubmit(e) {
  var sheetName = "Form Responses 1";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var activeRow = e.range.getRow();

  // Fetch values from columns A to Z
  var values = sheet.getRange(activeRow, 1, 1, 26).getValues()[0];
  var [
    columnA, columnB, columnC, columnD, columnE, columnF, columnG, columnH, columnI, columnJ, 
    columnK, columnL, columnM, columnN, columnO, columnP, columnQ, columnR, columnS, 
    columnT, columnU, columnV, columnW, columnX, columnY, columnZ
  ] = values;

  // Auto-increment ID by finding the last used ID and adding 1
  var lastRow = sheet.getLastRow();
  var lastId = sheet.getRange(lastRow, 2).getValue(); // Assuming Column B is for IDs
  var id = (lastId && !isNaN(lastId)) ? lastId + 1 : 1; // Start at 1 if no valid ID found
  sheet.getRange(activeRow, 2).setValue(id); // Set the auto-increment ID in Column B

  // Initialize variables based on columns
  var name = columnC;
  var email = columnD;
  var leaveType = columnE;
  var leaveStart = columnF;
  var leaveEnd = columnG;
  var reason = columnH;

  // Set the unique ID in a new column (Column AA, which is the 27th column)
  sheet.getRange(activeRow, 27).setValue(id);

  // Set the default value of Status Column (Column J, which is the 10th column) to "Pending"
  sheet.getRange(activeRow, 10).setValue("Pending");

  // Calculate the duration between "From Date" (Column F) and "To Date" (Column G)
  var duration = calculateDateDifference(leaveStart, leaveEnd);
  sheet.getRange(activeRow, 9).setValue(duration); // Set the calculated duration in Column I (index 9)

  // Email content
  var emailSubject = "New Leave Request Submitted";
  var emailBody = `
    <p>A new leave request has been submitted.</p>
    <p><strong>Employee Name:</strong> ${name}</p>
    <p><strong>Leave Type:</strong> ${leaveType}</p>
    <p><strong>From Date:</strong> ${leaveStart}</p>
    <p><strong>To Date:</strong> ${leaveEnd}</p>
    <p><strong>Reason:</strong> ${reason}</p>
    <p><strong>Duration:</strong> ${duration} days</p>
    <p><strong>Current Status:</strong> Pending</p>

    <p>Approve or Reject in <a href="https://docs.google.com/spreadsheets/d/1Pt-hz5bmZ2U_BqTxIp0bacoFybKn0dj05iNywM6LYDs/edit?usp=sharing">this link</a>.</p>
  `;

  // Send email to HR
  var hrEmail = "m.hakimjurimi@gmail.com"; // Replace with HR's email address
  sendMail(hrEmail, emailSubject, emailBody);

  // Log or utilize the Reason from Column G as needed
  Logger.log("Reason for leave: " + reason);
}


function calculateDateDifference(startDate, endDate) {
  var startTimestamp = new Date(startDate).getTime();
  var endTimestamp = new Date(endDate).getTime();
  
  // Check for valid dates and ensure start date is before or equal to end date
  if (isNaN(startTimestamp) || isNaN(endTimestamp) || startTimestamp > endTimestamp) {
    return "Invalid Date";
  }
  
  var millisecondsInADay = 1000 * 60 * 60 * 24;
  var differenceInDays = Math.floor((endTimestamp - startTimestamp) / millisecondsInADay) + 1; // Include the start day
  return differenceInDays;
}

// function getLeaveBalance(email, leaveType) {
//   var leaveBalanceSheetName = "Leave Balance"; // Name of the "Leave Balance" sheet
//   var leaveBalanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(leaveBalanceSheetName);
//   var leaveBalanceData = leaveBalanceSheet.getDataRange().getValues();

//   // Find the row that matches the email and leave type
//   for (var i = 1; i < leaveBalanceData.length; i++) { // Assuming row 1 contains headers
//     var sheetEmail = leaveBalanceData[i][1]; // Email is in the second column (index 1)
//     var sheetLeaveType = leaveBalanceData[i][2]; // Leave type is in the third column (index 2)
//     var sheetLeaveBalance = leaveBalanceData[i][5]; // Leave balance is in the sixth column (index 5)

//     if (sheetEmail === email && sheetLeaveType === leaveType) {
//       return sheetLeaveBalance;
//     }
//   }
  
//   // If no matching record is found, return a default value
//   return 0; // Default balance if no match is found
// }


// function sendInsufficientBalanceEmail(sendTo, leaveType, duration) {
//   // Log the sendTo variable to verify it's a correct email string
//   Logger.log("Email to send: " + sendTo);
  
//   if (typeof sendTo !== 'string' || !sendTo.includes('@')) {
//     Logger.log("Invalid email address detected: " + sendTo);
//     return; // Stop the function if sendTo is not a valid email
//   }
  
//   var mailSubject = "Insufficient Leave Balance - " + leaveType;
//   var mailBody =
//     "Dear Employee,<br>" +
//     "Your request for " +
//     leaveType +
//     " leave has been received, but the requested duration (" +
//     duration +
//     " days) exceeds your available leave balance.<br>" +
//     "Please review your leave balance and consider adjusting your request accordingly.<br><br>" +
//     "Thank you,<br>" +
//     "HR Department";

//   sendMail(sendTo, mailSubject, mailBody);
// }


// function sendSufficientBalanceEmail(sendTo, leaveType, duration) {
//   var mailSubject = "Leave Request Submitted - " + leaveType;
//   var mailBody = 
//     "Dear Employee,<br>" +
//     "Your request for " + leaveType + " leave has been submitted. The requested duration is " +
//     duration + " days.<br>" +
//     "After approved, you will get email notification.<br>" +
//     "Please ensure to manage your workload accordingly during your absence.<br><br>" +
//     "Thank you,<br>" +
//     "HR Department";

//   sendMail(sendTo, mailSubject, mailBody);
// }

function onColumnChangeApprovedReject(e) {
    if (!e || !e.range) {
        Logger.log("Invalid event object");
        return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
    var activeRange = e.range;

    if (activeRange.getColumn() === 10) { // Column J is 10 (Status column)
        var row = activeRange.getRow();
        
        // Fetch values from columns A to Z of the active row
        var values = sheet.getRange(row, 1, 1, 26).getValues()[0];
        var [
            columnA, columnB, columnC, columnD, columnE, columnF, columnG, columnH, columnI, columnJ, 
            columnK, columnL, columnM, columnN, columnO, columnP, columnQ, columnR, columnS, 
            columnT, columnU, columnV, columnW, columnX, columnY, columnZ
        ] = values;

        // Initialize variables based on columns
        var name = columnC;
        var email = columnD;
        var leaveType = columnE;
        var leaveStart = columnF;
        var leaveEnd = columnG;
        var reason = columnH;
        var status = columnJ; // New status value in Column J

        // HR email address
        var hrEmail = 'm.hakimjurimi@gmail.com'; 

        // Check the status change and send the appropriate email
        if (status === "Approved") {
            var mailSubject = "Leave Application Approved - " + leaveType;
            var mailBody =
                "Dear " + name + ",<br>" +
                "Your leave application has been approved for the period from " + leaveStart + " to " + leaveEnd + ".<br>" +
                "Enjoy your well-deserved break!<br><br>" +
                "Thank you,<br>" +
                "HR Department";

            sendMail(email, mailSubject, mailBody); // Sending email to the employee
        } else if (status === "Reject") {
            var mailSubject = "Leave Application Rejected - " + leaveType;
            var mailBody =
                "Dear " + name + ",<br>" +
                "Unfortunately, your leave application for the period from " + leaveStart + " to " + leaveEnd + " has been rejected.<br>" +
                "Please review and adjust your leave request as needed.<br><br>" +
                "Thank you,<br>" +
                "HR Department";

            sendMail(email, mailSubject, mailBody); // Sending email to the employee
        }
    }
}



function sendMail(sendTo, mailSubject, mailBody) {
  // Implement your email sending code here (e.g., using MailApp)
  // This function should send the email to the specified recipient(s).
  MailApp.sendEmail({
    to: sendTo,
    subject: mailSubject,
    htmlBody: mailBody,
  });
}
