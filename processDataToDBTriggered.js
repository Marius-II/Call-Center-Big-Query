// this is the new process data ti db function that is using the sheet name not active sheet so the trigger can do its job. The trigger can't activate sheets so we need to use sheet name
function processRowsIfInTimeRangeFollowUpCC() {
  var currentTime = new Date();
  var currentHour = currentTime.getHours();

  // Check if the current time is between 19:00 and 23:00
  if (currentHour >= 21 && currentHour < 22) {
    processRowsFollowUpCC();  // Only run this if the time is within the desired range
  } else {
    Logger.log("Current time is outside the specified range (19:00-23:00). No processing done.");
  }
}

function processRowsIfInTimeRangePeriodic() {
  var currentTime = new Date();
  var currentHour = currentTime.getHours();

  // Check if the current time is between 19:00 and 23:00
  if (currentHour >= 19 && currentHour < 21) {
    processRowsPeriodic();  // Only run this if the time is within the desired range
  } else {
    Logger.log("Current time is outside the specified range (19:00-23:00). No processing done.");
  }
}

function processRowsIfInTimeRangeFollowUpPeriodic() {
  var currentTime = new Date();
  var currentHour = currentTime.getHours();

  // Check if the current time is between 19:00 and 23:00
  if (currentHour >= 23 && currentHour < 24) {
    processRowsFollowUpPeriodic();  // Only run this if the time is within the desired range
  } else {
    Logger.log("Current time is outside the specified range (19:00-23:00). No processing done.");
  }
}


function processRowsPeriodic() {
  // Setup the sheet you want the trigger to run
  var triggerSheetName = "Periodic";

  var followUpSheetsNamesArray = ["Periodic", "Follow Up - P", "Follow Up - CC", "Cautare 1", "Cautare 2", "Istoric"];
  var dateProcessButtonPressed = new Date();
  var dataArrayFromCallCenterSheet = extractDataFromPeriodicSheet(5, triggerSheetName);

  var activeSheetName = triggerSheetName;

  // Set a time limit of 5 minutes and 30 seconds (330,000 milliseconds)
  var maxExecutionTime = 330000; // 5 minutes and 30 seconds
  var startTime = Date.now();

  var rowsProcessed = 0; // Track how many rows were processed
  for (var i = 0; i < dataArrayFromCallCenterSheet.length; i++) {
    var index = 0;
    if (followUpSheetsNamesArray.includes(activeSheetName)) {
      index = 1;
    }
    var phoneNumber = dataArrayFromCallCenterSheet[i][0];
    var name = dataArrayFromCallCenterSheet[i][1];
    var location = dataArrayFromCallCenterSheet[i][2];
    var oldCallResult = dataArrayFromCallCenterSheet[i][5];
    var callResult = dataArrayFromCallCenterSheet[i][5 + index];
    var reInteractTime = dataArrayFromCallCenterSheet[i][6 + index];

    var conversationDetails;
    if (activeSheetName === "Periodic" || activeSheetName === "Follow Up - P") {
      conversationDetails = dataArrayFromCallCenterSheet[i][7 + index] + " - Periodic";
    } else {
      conversationDetails = dataArrayFromCallCenterSheet[i][7 + index];
    }

    var statusDB = dataArrayFromCallCenterSheet[i][8 + index];
    var codeFromCallCenter = dataArrayFromCallCenterSheet[i][9 + index];

    if (callResult !== '' && reInteractTime !== '' && statusDB !== "DB processed") {
      // Search by code in BigQuery
      var rowToBeModified = searchByCodeInBigQuery(codeFromCallCenter);

      if (rowToBeModified) {
        updateRow(rowToBeModified, codeFromCallCenter, name, phoneNumber, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheetTriggered(i, triggerSheetName);
      } else {
        insertNotFoundedRows(codeFromCallCenter, phoneNumber, name, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheetTriggered(i, triggerSheetName);
      }
      countResultTransition(oldCallResult, callResult, location)
      rowsProcessed++;  // Increment the processed row count

      // Check if the time limit has been exceeded
      var currentTime = Date.now();
      if (currentTime - startTime >= maxExecutionTime) {
        Logger.log("Execution time limit reached. Stopping process.");
        break; // Stop the loop if the time limit has been exceeded
      }
    }
  }

  Logger.log("Rows processed in this run: " + rowsProcessed);


  if (followUpSheetsNamesArray.includes(activeSheetName)) {
    updateFollowUpLogWithCounts(dateProcessButtonPressed);
    logResultTransitions(triggerSheetName)
  } else if (activeSheetName === "Call Center") {
    updateCallCenterLogWithCounts(dateProcessButtonPressed);
  }
}

function processRowsFollowUpPeriodic() {
  // Setup the sheet you want the trigger to run
  var triggerSheetName = "Follow Up - P";

  var followUpSheetsNamesArray = ["Periodic", "Follow Up - P", "Follow Up - CC", "Cautare 1", "Cautare 2", "Istoric"];
  var dateProcessButtonPressed = new Date();
  var dataArrayFromCallCenterSheet = extractDataFromPeriodicSheet(5, triggerSheetName);

  var activeSheetName = triggerSheetName;

  // Set a time limit of 5 minutes and 30 seconds (330,000 milliseconds)
  var maxExecutionTime = 330000; // 5 minutes and 30 seconds
  var startTime = Date.now();

  var rowsProcessed = 0; // Track how many rows were processed
  for (var i = 0; i < dataArrayFromCallCenterSheet.length; i++) {
    var index = 0;
    if (followUpSheetsNamesArray.includes(activeSheetName)) {
      index = 1;
    }
    var phoneNumber = dataArrayFromCallCenterSheet[i][0];
    var name = dataArrayFromCallCenterSheet[i][1];
    var location = dataArrayFromCallCenterSheet[i][2];
    var oldCallResult = dataArrayFromCallCenterSheet[i][5];
    var callResult = dataArrayFromCallCenterSheet[i][5 + index];
    var reInteractTime = dataArrayFromCallCenterSheet[i][6 + index];

    var conversationDetails;
    if (activeSheetName === "Periodic" || activeSheetName === "Follow Up - P") {
      conversationDetails = dataArrayFromCallCenterSheet[i][7 + index] + " - Periodic";
    } else {
      conversationDetails = dataArrayFromCallCenterSheet[i][7 + index];
    }

    var statusDB = dataArrayFromCallCenterSheet[i][8 + index];
    var codeFromCallCenter = dataArrayFromCallCenterSheet[i][9 + index];

    if (callResult !== '' && reInteractTime !== '' && statusDB !== "DB processed") {
      // Search by code in BigQuery
      var rowToBeModified = searchByCodeInBigQuery(codeFromCallCenter);

      if (rowToBeModified) {
        updateRow(rowToBeModified, codeFromCallCenter, name, phoneNumber, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheetTriggered(i, triggerSheetName);
      } else {
        insertNotFoundedRows(codeFromCallCenter, phoneNumber, name, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheetTriggered(i, triggerSheetName);
      }
      rowsProcessed++;  // Increment the processed row count
      countResultTransition(oldCallResult, callResult, location)
      // Check if the time limit has been exceeded
      var currentTime = Date.now();
      if (currentTime - startTime >= maxExecutionTime) {
        Logger.log("Execution time limit reached. Stopping process.");
        break; // Stop the loop if the time limit has been exceeded
      }
    }
  }

  Logger.log("Rows processed in this run: " + rowsProcessed);


  if (followUpSheetsNamesArray.includes(activeSheetName)) {
    updateFollowUpLogWithCounts(dateProcessButtonPressed);
    logResultTransitions(triggerSheetName)
  } else if (activeSheetName === "Call Center") {
    updateCallCenterLogWithCounts(dateProcessButtonPressed);
  }
}

function processRowsFollowUpCC() {
  // Setup the sheet you want the trigger to run
  var triggerSheetName = "Follow Up - CC";

  var followUpSheetsNamesArray = ["Periodic", "Follow Up - P", "Follow Up - CC", "Cautare 1", "Cautare 2", "Istoric"];
  var dateProcessButtonPressed = new Date();
  var dataArrayFromCallCenterSheet = extractDataFromPeriodicSheet(5, triggerSheetName);

  var activeSheetName = triggerSheetName;

  // Set a time limit of 5 minutes and 30 seconds (330,000 milliseconds)
  var maxExecutionTime = 330000; // 5 minutes and 30 seconds
  var startTime = Date.now();

  var rowsProcessed = 0; // Track how many rows were processed
  for (var i = 0; i < dataArrayFromCallCenterSheet.length; i++) {
    var index = 0;
    if (followUpSheetsNamesArray.includes(activeSheetName)) {
      index = 1;
    }
    var phoneNumber = dataArrayFromCallCenterSheet[i][0];
    var name = dataArrayFromCallCenterSheet[i][1];
    var location = dataArrayFromCallCenterSheet[i][2];
    var oldCallResult = dataArrayFromCallCenterSheet[i][5];
    var callResult = dataArrayFromCallCenterSheet[i][5 + index];
    var reInteractTime = dataArrayFromCallCenterSheet[i][6 + index];

    var conversationDetails;
    if (activeSheetName === "Periodic" || activeSheetName === "Follow Up - P") {
      conversationDetails = dataArrayFromCallCenterSheet[i][7 + index] + " - Periodic";
    } else {
      conversationDetails = dataArrayFromCallCenterSheet[i][7 + index];
    }

    var statusDB = dataArrayFromCallCenterSheet[i][8 + index];
    var codeFromCallCenter = dataArrayFromCallCenterSheet[i][9 + index];

    if (callResult !== '' && reInteractTime !== '' && statusDB !== "DB processed") {
      // Search by code in BigQuery
      var rowToBeModified = searchByCodeInBigQuery(codeFromCallCenter);

      if (rowToBeModified) {
        updateRow(rowToBeModified, codeFromCallCenter, name, phoneNumber, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheetTriggered(i, triggerSheetName);
      } else {
        insertNotFoundedRows(codeFromCallCenter, phoneNumber, name, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheetTriggered(i, triggerSheetName);
      }
      rowsProcessed++;  // Increment the processed row count
      countResultTransition(oldCallResult, callResult, location)
      
      // Check if the time limit has been exceeded
      var currentTime = Date.now();
      if (currentTime - startTime >= maxExecutionTime) {
        Logger.log("Execution time limit reached. Stopping process.");
        break; // Stop the loop if the time limit has been exceeded
      }
    }
  }

  Logger.log("Rows processed in this run: " + rowsProcessed);


  if (followUpSheetsNamesArray.includes(activeSheetName)) {
    updateFollowUpLogWithCounts(dateProcessButtonPressed);
    logResultTransitions(triggerSheetName)
  } else if (activeSheetName === "Call Center") {
    updateCallCenterLogWithCounts(dateProcessButtonPressed);
  }
}

// Initialize the transition count object
const locationResultTransitions = {};

function countResultTransition(oldCallResult, callResult, location) {
  location = transformLocation(location);

  // Initialize the location in the transitions object if it doesn't exist
  if (!locationResultTransitions[location]) {
    locationResultTransitions[location] = {};
  }

  // Initialize the old result in the location if it doesn't exist
  if (!locationResultTransitions[location][oldCallResult]) {
    locationResultTransitions[location][oldCallResult] = {
      "Programat": 0,
      "De revenit": 0,
      "Nu raspunde": 0,
      "Refuza": 0,
    };
  }

  // Increment the transition count for the specific oldCallResult to callResult
  locationResultTransitions[location][oldCallResult][callResult]++;
}

function logResultTransitions(triggerSheetName) {
  const spreadsheet = SpreadsheetApp.openById("1yc5s-keQtGyrBEqeZlHcNsjB-c-E63OBKUue1befRnc");
  var sheetNameForLogResults = ""
  switch (triggerSheetName) {
    case "Periodic":
      sheetNameForLogResults = "CountByLocationPeriodic"
      break;
    case "Follow Up - CC":
      sheetNameForLogResults = "CountByLocationFollowUpCC"
      break;
    case "Follow Up - P":
      sheetNameForLogResults = "CountByLocationFollowUpP"
      break;
    default:
      sheetNameForLogResults = ""
  }
  const sheet = spreadsheet.getSheetByName(sheetNameForLogResults);

  // Get the current date
  const currentDate = new Date();

  // Starting row for logging (you can adjust this if needed)
  let startRow = sheet.getLastRow() + 1;

  // Loop through locationResultTransitions to log data
  for (let location in locationResultTransitions) {
    for (let oldResult in locationResultTransitions[location]) {
      for (let newResult in locationResultTransitions[location][oldResult]) {
        let transitionCount = locationResultTransitions[location][oldResult][newResult];

        // Only log if transition count is greater than 0
        if (transitionCount > 0) {
          // Write data to the sheet
          sheet.getRange(startRow, 1).setValue(currentDate);             // Column A: Date
          sheet.getRange(startRow, 2).setValue(location);                // Column B: Location
          sheet.getRange(startRow, 3).setValue(oldResult);               // Column C: Old Result
          sheet.getRange(startRow, 4).setValue(newResult);               // Column D: New Result
          sheet.getRange(startRow, 5).setValue(transitionCount);         // Column E: Transition Count

          // Move to the next row for each new transition entry
          startRow++;
        }
      }
    }
  }
}


function extractDataFromPeriodicSheet(headerRow, triggerSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var callCenterSheet = spreadsheet.getSheetByName(triggerSheetName);
  
  if (callCenterSheet) {
    var lastRow = callCenterSheet.getLastRow();
    var numRowsToExtract = lastRow - headerRow;
    if (numRowsToExtract > 0) {
      var lastColumn = callCenterSheet.getLastColumn();
      var dataArrayFromCallCenterSheet = callCenterSheet.getRange(headerRow + 1, 4, lastRow - headerRow, lastColumn - 3).getValues();
    }
  } else {
    Logger.log('Call Center sheet not found.');
  }

  return dataArrayFromCallCenterSheet;
}

function updateRowInCallCenterSheetTriggered(rowNumberInDataSet, triggerSheetName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(triggerSheetName);
  var sheetName = sheet.getName();
  var range;

  switch (sheetName) {
    case "Call Center":
      range = sheet.getRange(6 + rowNumberInDataSet, 12);
      break;
    case "Periodic":
    case "Follow Up - CC":
    case "Follow Up - P":
    case "Cautare 1":
    case "Cautare 2":
    case "Istoric":
      range = sheet.getRange(6 + rowNumberInDataSet, 13);
      break;
    default:
      showUserMessage("Numele unuia sau mai multor sheeturi s-a modificat, prin urmare programul nu poate scrie 'DB Processed' in coloana corespunzatoare, conditie imperios necesara pentru functionare corespunzatoare. Contacteaza developer: 0752484554. ---sau mergi la functia updateRowInCallCenterSheet pentru a updata corespunzator");
  }

  Logger.log(range.getA1Notation());
  range.setValue("DB processed");
}