  // Define the Admin email addresses
  const adminEmails = [
    'yousuf@zahurmostafiz.com',
    'sattar@zahurmostafiz.com'
  ];


  // Define the email addresses of the users and the corresponding column ranges
  var emailColumnPermissions = [
    { name: 'cloud server inc', email: 'cloudserverinc@gmail.com', startcolumn: 4,endcolumn: 390, startRow: 3, endRow: 8 },  // Columns A and B for razuvkt.cpa@gmail.com
    { name: 'xp style edition', email: 'xpstyleedition@gmail.com', startcolumn: 4,endcolumn: 390, startRow: 10, endRow: 14 }  // Columns C and D for pinkykathun460@gmail.com
  ];

  


  // Get Active User Email addresses
  var activeUserEmail = Session.getActiveUser().getEmail();

  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get Active Sheet
  var activeSheet = spreadsheet.getActiveSheet();

  // Get the sheet name
  var sheetName = activeSheet.getName();

  // Get the owner's email address
  var owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
  var ownerEmail = owner.getEmail();

  // Log the sheet name and owner email address
  Logger.log('Google sheet name : ' + sheetName);
  Logger.log('Sheet : ' + activeSheet);
  Logger.log('Sheet Name: ' + sheetName);
  Logger.log('Owner : ' + owner);
  Logger.log('Owner Email: ' + ownerEmail);
    
  // ####### go to specific cell and write somthing######
  // Define the row (6) and column (50)  
  var row = 6;
  var column = 390;
  
  // Set the cursor to the specified cell
  var activeRange = activeSheet.getRange(row, column);
  //activeSheet.setActiveRange(activeRange);
  
  // Optionally, you can write something to the cell
  //activeRange.setValue("⏬ Enter here ⏬"); // Replace "Your text here" with the value you want to set

  // Add a comment to the cell
  //activeRange.setComment("1Enter your Todays Data here");
  activeRange.setNote(activeUserEmail + " Enter your Todays Data here");

  // ####### End go to specific cell ######


function protectColumns() {

  
  // Define the column (390th column is the same as column number 390)
  var columnNumber = 390;
  
  // Write "end" in the first row of the 390th column
  //activeSheet.getRange(1, columnNumber).setValue("end");

  
  // Step 1: Remove all existing protection rules on the sheet
  var protections = activeSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(protection) {
    protection.remove(); // Remove each protection
  });

  // Step 2: Loop through each permission and set column range protection
  emailColumnPermissions.forEach(function(permission) {
    var name = permission.name;
    var email = permission.email;
    var startColumn = permission.startcolumn;  // Array of columns the user is allowed to edit
    var startRow = permission.startRow;
    var endRow = permission.endRow;
    

      var range = activeSheet.getRange(startRow, startColumn, endRow - startRow + 1, 200); // Protect the specific range for this column
      // Log data for this column
      Logger.log('Name : ' + name + ', email: ' + email + ', Column: ' + startColumn +', start row: '+ startRow + ', end row: '+endRow);
      //  columnData.forEach(function(row, index) {
      //   Logger.log('Row ' + (startRow + index) + ': ' + row[0]); // row[0] is the value in the single column
      // });

      // Create a protection object
      var protection = range.protect().setDescription(name+' ' + startColumn + ' Range ' + startRow + '-' + endRow);
      


     // Allow the owner to edit this range
    protection.addEditor(ownerEmail);

    // Allow Admin's to edit this range - Use addEditors to add all admins at once
    protection.addEditors(adminEmails);  // Adding all admins in one go

    // Log the admin emails that have been added
    adminEmails.forEach(function(adminemail) {
      Logger.log('Admin email added: ' + adminemail);
    });

    // Add the specified email (the user) to the editors
    protection.addEditor(email);

    // Remove all other editors except the owner, admin, and the specified user
    var editors = protection.getEditors();
    editors.forEach(function(editor) {
      if (editor.getEmail() !== ownerEmail && editor.getEmail() !== email && !adminEmails.includes(editor.getEmail())) {
        protection.removeEditor(editor);
      }
    });

  });
}


function markToday() {
  // Get the active spreadsheet and the first sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the range D1:H1 (dates)
  var dateRange = sheet.getRange("D1:CW1");
  var dates = dateRange.getValues()[0];  // This gives us an array of dates
  
  // Get today's date (without time)
  var today = new Date();
  today.setHours(0, 0, 0, 0);  // Set time to 00:00 to ignore the time part

  // Loop through the dates and check for a match with today's date
  for (var i = 0; i < dates.length; i++) {
    var cellDate = new Date(dates[i]);
    cellDate.setHours(0, 0, 0, 0);  // Ensure the time part is ignored when comparing
    
    // Check if the date matches today's date
    if (cellDate.getTime() === today.getTime()) {
      // Mark the entire column (D to H)
      var selectCells = sheet.getRange(9, i + 4, 5, 1); //.setBackground('#9fc5e8');  // sheet.getRange(9, i + 4, sheet.getMaxRows(), 1).setBackground('yellow');
      sheet.setActiveRange(selectCells);
      Logger.log("i = " + i);
      Logger.log("getMaxRows = " + sheet.getMaxRows());
    } else {
      // Optionally reset the background for cells that don't match
      sheet.getRange(1, i + 4, sheet.getMaxRows(), 1).setBackground('');  // Clear previous highlights
    }
  }
}

function selectTodayRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the value from F1 (start)
  var startDateColumn = "H";
  var columnNumber = sheet.getRange(startDateColumn+"1").getColumn();
  
  
  // Get today's date
  var today = new Date();
  
  // Get the start of the year (January 1st of the current year)
  var startOfYear = new Date(today.getFullYear(), 0, 1);
  
  // Calculate the day number of the year
  var dayOfYear = Math.ceil((today - startOfYear) / (1000 * 60 * 60 * 24)) + 0;
  
  // Calculate the result (start + day of the year)
  var result = columnNumber + dayOfYear-1;
  
  // Output the result to a specific cell (e.g., in cell G1)
  //sheet.getRange("G1").setValue(result);
  Logger.log("Todays day = " + dayOfYear);
  Logger.log("column name = " + startDateColumn);
  Logger.log("column Number = " + columnNumber);  // This will log the column number in the log
  Logger.log("result = " + result);
  var columnName = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getRange(1, result).getA1Notation().replace(/[0-9]/g, '');
  
 Logger.log("Today column name"+ columnName);  // This will log "G"
  var cellsView = (result)+6;
    // First select the range
  var selectCellsView = sheet.getRange(4, cellsView, 5, 1);
  sheet.setActiveRange(selectCellsView);
  
  // Force the changes to take effect immediately
  SpreadsheetApp.flush();
  // Pause for 5 seconds (5000 milliseconds)
  //Utilities.sleep(5000);
  
  // Then select the second range
  var selectCells = sheet.getRange(4, result, 5, 1).setBackground('yellow');
  sheet.setActiveRange(selectCells);

}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedRange = e.range;
  const userEmail = Session.getActiveUser().getEmail(); // Get the email of the user editing

  // Specify the row range you want to lock (3 to 7 in this example)
  const lockedRows = { start: 3, end: 7 };

  // Specify the column that should remain editable (column L, which is column 12)
  const editableColumn = 12; 

  // Check if the edit is within rows 3-7 and columns other than L
  if (editedRange.getRow() >= lockedRows.start && editedRange.getRow() <= lockedRows.end) {
    if (editedRange.getColumn() !== editableColumn) {
      // If the edited cell is outside column L (column 12), prevent the edit
      e.range.setValue(e.oldValue); // Revert to the old value, effectively locking it
      SpreadsheetApp.getUi().alert("You are not allowed Today to edit this cell.");
    }
  }
}

function resetCellColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the range of all cells in the sheet (you can specify a different range if needed)
  const range = sheet.getDataRange();
  
  // Reset the background color to white (no color)
  range.setBackground(null); // Or you can use setBackground("#FFFFFF") for white color
}

function insertComments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Define the range of cells (M4 to M8)
  const cells = ['M4', 'M5', 'M6', 'M7', 'M8'];
  
  // Define the comment for each cell
  const comments = [
    'This is a comment for M4',
    'This is a comment for M5',
    'This is a comment for M6',
    'This is a comment for M7',
    'This is a comment for M8'
  ];
  
  // Loop through the cells and add comments
  for (let i = 0; i < cells.length; i++) {
    const cell = sheet.getRange(cells[i]);
    cell.setComment(comments[i]);
  }
}
function checkUserEmail() {
  // List of allowed admin emails
  var adminEmail = [ 
    { email: 'admin1@gmail.com' },
    { email: 'admin2@gmail.com' }
  ];

  // Get the email of the current user
  const currentUserEmail = Session.getActiveUser().getEmail();
  
  // Check if the current user's email matches any of the admin emails
  const isAdmin = adminEmail.some(function(admin) {
    return admin.email === currentUserEmail;
  });
  
  // If the current user is not an admin, log "ok"
  if (!isAdmin) {
    console.log('ok');
  }
}


function onOpen() {
  selectTodayRow();  // Call the markToday function when the sheet is opened
}
resetCellColors();
checkUserEmail();
//insertComments();
//selectTodayRow();

