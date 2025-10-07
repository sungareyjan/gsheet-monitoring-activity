function onEdit(e) {
  var sheet = e.range.getSheet();
  
  // Only run on "Activity" sheet and when column G (Office) is edited
  if (sheet.getName() === "Activity" && e.range.getColumn() == 8) {
    var office = e.range.getValue();
    var officesSheet = e.source.getSheetByName("Division");
    
    // Get office + color mapping
    var data = officesSheet.getRange("B2:C" + officesSheet.getLastRow()).getValues();
    
    // Find the color for the selected office
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == office) {
        // Column O is the 7th column → same row as the edited dropdown
        var targetCell = sheet.getRange(e.range.getRow(), 7);
        targetCell.setBackground(data[i][1]);
        break;
      }
    }
  }
}
function whenhaveValue(e) {
  // Safety checks
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();

  // Only run on "testSearch" sheet and when column I (9) is edited
  if (sheet.getName() !== "testSearch" || e.range.getColumn() !== 9) return;

  var office = e.range.getValue();
  if (!office) return; // do nothing if blank

  var ss = e.source;
  var officesSheet = ss.getSheetByName("Division");
  if (!officesSheet) return; // safety: "Division" sheet must exist

  // Get Division → Color mapping (columns B and C)
  var data = officesSheet.getRange("B2:C" + officesSheet.getLastRow()).getValues();

  // Find matching Division
  var color = null;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === office.toString().trim().toLowerCase()) {
      color = data[i][1]; // color code (e.g., #4FC3F7)
      break;
    }
  }

  // Apply color if found
  var targetCell = sheet.getRange(e.range.getRow(), 8); // Column H
  if (color) {
    targetCell.setBackground(color);
  } else {
    // Optional: clear if Division not found
    targetCell.clearBackground();
  }
}



/** @OnlyCurrentDoc */
function saveToActivity() {
  var ss = SpreadsheetApp.getActive();
  var form = ss.getSheetByName("Form Activity");
  var activity = ss.getSheetByName("Activity");
  var targetRow = activity.getLastRow() + 1;

  var mappings = [
    { from: "C2", name: "Date Start" },       
    { from: "C3", name: "Time Start" },       
    { from: "C5", name: "Date End" },         
    { from: "C6", name: "Time End" },         
    { from: "C8", name: "Activity Title" },   
    { from: "B13", name: "Target Participants" }, 
    { from: "C10", name: "Office Facilitator" },  
    { from: "B27", name: "Remarks" },        
    { from: "C33", name: "Link" }            
  ];

  var errors = [];
  var rowData = [];

  mappings.forEach(function(m) {
    var value = form.getRange(m.from).getValue();

    // Handle empty cells
    if (value === "" || value === null) {
      errors.push(m.name + " is required.");
    }

    // Detect date/time and format properly
    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
      // If the mapping is a date cell (C2, C5), format as date
      if (m.from === "C2" || m.from === "C5") {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MMM d, yyyy");
      }
      // If it's a time cell (C3, C6), format as time
      else if (m.from === "C3" || m.from === "C6") {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "h:mm a");
      }
    } else if (typeof value === "string") {
      value = value.trim();
    }

    rowData.push(value);
  });

  if (errors.length > 0) {
    throw new Error("Validation failed:\n" + errors.join("\n"));
  }

  // Prepare row for Activity sheet
  var newRow = new Array(11).fill("");
  newRow[0] = rowData[0]; // Date Start
  newRow[1] = rowData[1]; // Time Start
  newRow[2] = rowData[2]; // Date End
  newRow[3] = rowData[3]; // Time End
  newRow[4] = rowData[4]; // Title
  newRow[5] = rowData[5]; // Participants
  newRow[7] = rowData[6]; // Office Facilitator
  newRow[9] = rowData[7]; // Remarks
  newRow[10] = rowData[8]; // Link

  activity.getRange(targetRow, 1, 1, 11).setValues([newRow]);

  // Clear form
  mappings.forEach(function(m) {
    form.getRange(m.from).clearContent();
  });

  return "Data saved successfully in row " + targetRow;
}

/** @OnlyCurrentDoc */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index");
}
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Activity Form");
  SpreadsheetApp.getUi().showSidebar(html);
}

// Save form data into Activity sheet
function saveToActivityFromHtml(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activity = ss.getSheetByName("Activity");
  var targetRow = activity.getLastRow() + 1;

  // Write values in proper columns
  activity.getRange(targetRow, 1).setValue(data.dateStart);          // A: Date Start
  activity.getRange(targetRow, 2).setValue(data.timeStart);          // B: Time Start
  activity.getRange(targetRow, 3).setValue(data.dateEnd);            // C: Date End
  activity.getRange(targetRow, 4).setValue(data.timeEnd);            // D: Time End
  activity.getRange(targetRow, 5).setValue(data.title);              // E: Title
  activity.getRange(targetRow, 6).setValue(data.targetParticipants); // F: Participants
  activity.getRange(targetRow, 8).setValue(data.officeFacilitator);  // H: Office Facilitator
  activity.getRange(targetRow, 10).setValue(data.remarks);           // J: Remarks
  activity.getRange(targetRow, 11).setValue(data.link);              // K: Link
  activity.getRange(targetRow, 12).setValue(data.color);              // L: Color Code
    // ✅ Set background color in Office Facilitator cell
  if (data.color) {
    activity.getRange(targetRow, 7).setBackground(data.color);
  }

  return { row: targetRow };
}



function getOffices() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Division");
  var data = sheet.getRange("B2:C" + sheet.getLastRow()).getValues();

  // Filter out empty rows and map to objects
  var result = data
    .filter(function(row) {
      return row[0] && row[1]; // ensure both name and color exist
    })
    .map(function(row) {
      return {
        name: row[0],
        color: row[1]
      };
    });

  return result;
}

function movePastActivitiesToArchive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activitySheet = ss.getSheetByName("Activity");
  var archiveSheet = ss.getSheetByName("ActivityArchive");

  if (!activitySheet || !archiveSheet) {
    Logger.log("One or both sheets are missing.");
    return;
  }

  var data = activitySheet.getDataRange().getValues();
  if (data.length < 2) return; // only header

  var header = data[0];
  var today = new Date();
  today.setHours(0, 0, 0, 0); // ignore time

  var toKeep = [header];
  var toArchive = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateEnd = row[2]; // column C = Date End

    // check if valid date and already ended
    if (dateEnd instanceof Date && dateEnd < today) {
      toArchive.push(row);
    } else {
      toKeep.push(row);
    }
  }

  if (toArchive.length > 0) {
    // append to ActivityArchive
    var archiveLast = archiveSheet.getLastRow();
    archiveSheet.getRange(archiveLast + 1, 1, toArchive.length, toArchive[0].length)
      .setValues(toArchive);

    // rewrite Activity with remaining rows
    activitySheet.clearContents();
    activitySheet.getRange(1, 1, toKeep.length, toKeep[0].length)
      .setValues(toKeep);
  }

  Logger.log("Moved " + toArchive.length + " rows to ActivityArchive.");
}



