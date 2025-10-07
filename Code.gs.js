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


/** @OnlyCurrentDoc */
function saveToActivity() {
  var ss = SpreadsheetApp.getActive();
  var form = ss.getSheetByName("Form Activity");
  var activity = ss.getSheetByName("Activity");
  var divisionSheet = ss.getSheetByName("Division"); // <-- Sheet with Division & Hex color
  var targetRow = activity.getLastRow() + 1;

  var mappings = [
    { from: "C2", name: "Date Start" },       
    { from: "C3", name: "Time Start" },       
    { from: "C5", name: "Date End" },         
    { from: "C6", name: "Time End" },         
    { from: "C8", name: "Activity Title" },   
    { from: "B13", name: "Target Participants" }, 
    { from: "C10", name: "DivisionFacilitator" },   
    { from: "B27", name: "Remarks" },        
    { from: "C33", name: "Link" }            
  ];

  var errors = [];
  var rowData = [];

  // --- Validate & Format ---
  mappings.forEach(function(m) {
    var value = form.getRange(m.from).getValue();
    if (value === "" || value === null) errors.push(m.name + " is required.");

    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
      if (m.from === "C2" || m.from === "C5") {
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MMM d, yyyy");
      } else if (m.from === "C3" || m.from === "C6") {
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

  // --- Prepare data for Activity sheet ---
  var newRow = new Array(11).fill("");
  newRow[0] = rowData[0]; // A Date Start
  newRow[1] = rowData[1]; // B Time Start
  newRow[2] = rowData[2]; // C Date End
  newRow[3] = rowData[3]; // D Time End
  newRow[4] = rowData[4]; // E Title
  newRow[5] = rowData[5]; // F Participants
  // G (color) skipped for now
  newRow[7] = rowData[6]; // H Division Facilitator
  newRow[9] = rowData[7]; // J Remarks
  newRow[10] = rowData[8]; // K Link

  // --- Get hex color based on Division Facilitator ---
  var facilitator = rowData[6];
  var divData = divisionSheet.getRange("B2:C" + divisionSheet.getLastRow()).getValues();
  var hexColor = "#ffffff"; // default white

  for (var i = 0; i < divData.length; i++) {
    if (String(divData[i][0]).trim().toLowerCase() === facilitator.toLowerCase()) {
      hexColor = divData[i][1];
      break;
    }
  }

  // --- Write main data (A–K) ---
  activity.getRange(targetRow, 1, 1, 11).setValues([newRow]);

  // --- Apply color (G) and save hex (L) ---
  activity.getRange(targetRow, 7).setBackground(hexColor);
  activity.getRange(targetRow, 12).setValue(hexColor);

  // --- Clear form ---
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

  var dataRange = activitySheet.getDataRange();
  var data = dataRange.getValues();
  var bg = dataRange.getBackgrounds(); // all background colors

  if (data.length < 2) return; // only header row

  var header = data[0];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var toKeep = [header];
  var toArchive = [];
  var gColorsToArchive = []; // G col colors for archive
  var gColorsToKeep = [[bg[0][6]]]; // header G color

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateEnd = row[2]; // column C = Date End

    if (dateEnd instanceof Date && dateEnd < today) {
      // --- Move to archive ---
      toArchive.push(row);
      gColorsToArchive.push([bg[i][6]]); // keep G color
    } else {
      // --- Keep in Activity ---
      toKeep.push(row);

      // if G has value → keep its color
      // else → use color from column L (index 11)
      if (String(row[6]).trim() !== "") {
        gColorsToKeep.push([bg[i][6]]);
      } else {
        gColorsToKeep.push([bg[i][6]]); // take color from column L
        console.log([bg[i][6]])
      }
    }
  }

  // --- Move to archive if needed ---
  if (toArchive.length > 0) {
    var archiveLast = archiveSheet.getLastRow();
    var startRow = archiveLast + 1;

    archiveSheet
      .getRange(startRow, 1, toArchive.length, toArchive[0].length)
      .setValues(toArchive);

    archiveSheet
      .getRange(startRow, 7, gColorsToArchive.length, 1)
      .setBackgrounds(gColorsToArchive);
  }

  // --- Rewrite Activity with kept rows ---
 var lastRow = activitySheet.getMaxRows();
  var lastCol = activitySheet.getMaxColumns();

  // Clear all contents
  activitySheet.clearContents();

  // Reset background of used area (or entire sheet)
  activitySheet.getRange(3, 1, lastRow, lastCol).setBackground('#ffffff');
  activitySheet
    .getRange(1, 1, toKeep.length, toKeep[0].length)  
    .setValues(toKeep);

  // --- Restore G column colors ---
  activitySheet
    .getRange(1, 7, gColorsToKeep.length, 1)
    .setBackgrounds(gColorsToKeep);

  Logger.log("Moved " + toArchive.length + " rows to ActivityArchive.");
}

function searchWithColor(startDate, endDate, keyword) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity");
  if (!sheet) return [["Activity sheet not found"]];

  // get all data and colors
  var dataRange = sheet.getRange("A3:M");
  var data = dataRange.getValues();
  var colors = dataRange.getBackgrounds(); // for cell colors

  var results = [];
  var today = new Date();

  // convert parameters safely
  var start = startDate instanceof Date ? startDate : null;
  var end = endDate instanceof Date ? endDate : null;
  var searchTerm = keyword ? keyword.toString().toLowerCase().trim() : "";

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateStart = row[0];
    var dateEnd = row[2];

    // skip empty rows
    if (!dateStart && !dateEnd && row.join("") === "") continue;

    // match filters
    var match = true;

    if (start && dateEnd instanceof Date && dateEnd < start) match = false;
    if (end && dateStart instanceof Date && dateStart > end) match = false;

    // keyword search in entire row (case-insensitive)
    if (searchTerm) {
      var rowText = row.join(" ").toLowerCase();
      if (!rowText.includes(searchTerm)) match = false;
    }

    if (match) {
      var gColor = colors[i][6]; // G column (7th col)
      var lColor = colors[i][11]; // L column (12th col), if you need it too
      var resultRow = row.concat([gColor, lColor]); // append colors
      results.push(resultRow);
    }
  }

  if (results.length === 0) return [["Search Not Found"]];
  return results;
}

function searchAndDisplayWithColor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("Form Activity"); // where you search
  var activitySheet = ss.getSheetByName("Activity");

  // read filters
  var startDate = formSheet.getRange("C3").getValue();
  var endDate = formSheet.getRange("C4").getValue();
  var keyword = formSheet.getRange("C6").getValue().toString().toLowerCase().trim();

  // get data + background
  var dataRange = activitySheet.getRange("A3:M");
  var data = dataRange.getValues();
  var colors = dataRange.getBackgrounds();

  // clear old results
  var outputStartRow = 9; // adjust where your table starts
  formSheet.getRange(outputStartRow, 1, 200, 13).clearContent().clearFormat();

  var results = [];
  var colorResults = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dateStart = row[0];
    var dateEnd = row[2];

    // skip empty
    if (row.join("").trim() === "") continue;

    var match = true;
    if (startDate && dateEnd instanceof Date && dateEnd < startDate) match = false;
    if (endDate && dateStart instanceof Date && dateStart > endDate) match = false;
    if (keyword && !row.join(" ").toLowerCase().includes(keyword)) match = false;

    if (match) {
      results.push(row);
      colorResults.push(colors[i]);
    }
  }

  if (results.length === 0) {
    formSheet.getRange(outputStartRow, 1).setValue("Search Not Found");
    return;
  }

  // write data
  formSheet.getRange(outputStartRow, 1, results.length, results[0].length).setValues(results);

  // apply G column color
  var gBg = colorResults.map(r => [r[6]]);
  formSheet.getRange(outputStartRow, 7, gBg.length, 1).setBackgrounds(gBg);
}






