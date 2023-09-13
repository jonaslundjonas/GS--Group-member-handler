// Function to create custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Group Member Handler')
      .addItem('Run Batch', 'manageGroupMembers')
      .addToUi();
}

function manageGroupMembers() {
  // Display initial message
  var ui = SpreadsheetApp.getUi();
  ui.alert('If you have more than 450 entries, this script will work for a while. Go and drink coffee! The script will inform you when all is done.');
  
  // Open the Google Sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  // Freeze the first row and make it bold
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
  sheet.setFrozenRows(1);
  
  // Read data from columns A to C, starting from row 2
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, 3).getValues();
  
  for (var i = 0; i < data.length; i++) {
    // Pause for 30 seconds after every 450 entries
    if (i % 450 === 0 && i !== 0) {
      Utilities.sleep(30000); // Pause for 30 seconds
    }
    
    var action = data[i][0]; // "add" or "remove"
    var groupEmail = data[i][1]; // Group's email address
    var memberEmail = data[i][2]; // Member's email address
    var status = ""; // Status message to be logged in Column D
    
    try {
      if (action === "add") {
        // Add member to the group
        var member = {
          email: memberEmail,
          role: "MEMBER"
        };
        AdminDirectory.Members.insert(member, groupEmail);
        status = "Success";
      } else if (action === "remove") {
        // Remove member from the group
        AdminDirectory.Members.remove(groupEmail, memberEmail);
        status = "Success";
      } else {
        status = "Invalid action";
      }
    } catch (e) {
      // Check the type of error and set the status message accordingly
      if (e.message.includes("Resource Not Found: groupKey")) {
        status = "Group does not exist";
      } else if (e.message.includes("Resource Not Found: memberKey")) {
        status = "Member does not exist in the group";
      } else {
        status = "Error: " + e;
      }
    }
    
    // Log the status in Column D
    sheet.getRange(i + 2, 4).setValue(status);
    Logger.log("Row " + (i + 2) + ": " + status);
  }
  
  // Display completion message
  ui.alert('Operation done! You can see the results in the Log column D.');
}
