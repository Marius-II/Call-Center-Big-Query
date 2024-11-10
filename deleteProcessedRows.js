function clearAllRowsFromHeaderOnwards(headerRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataArrayFromCallCenterSheet = extractDataFromCallCenterSheet(5);
  var rowsToCopy = []
  var unprocessedRowExist;

  for (i = 0; i < dataArrayFromCallCenterSheet.length; i++){
    var phoneNumber = dataArrayFromCallCenterSheet[i][0]
    var name = dataArrayFromCallCenterSheet[i][1]
    var location = dataArrayFromCallCenterSheet[i][2]
    var callResult = dataArrayFromCallCenterSheet[i][6];
    var reInteractTime = dataArrayFromCallCenterSheet[i][7];
    var conversationDetails = dataArrayFromCallCenterSheet[i][8];
    var statusDB = dataArrayFromCallCenterSheet[i][9];
    var codeFromCallCenter = dataArrayFromCallCenterSheet[i][10];
    var rowUpdated = false;

    if 
    (
      (phoneNumber && (callResult !== '' || reInteractTime !== '' || conversationDetails !== '') && statusDB === '') // this are the rows that have content noted and are not processed
    
    ){
      
      // Push the entire row to the 'rowsToCopy' array
      rowsToCopy.push(dataArrayFromCallCenterSheet[i]);      

    }
  }
  
  // Get the last row number in the sheet
  var lastRow = sheet.getLastRow();

  // Calculate the number of rows to clear (from headerRow to lastRow)
  var numRowsToClear = Math.max(0, lastRow - headerRow); // Exclude the header row

  // Clear the content of the range starting from the row after the headerRow
  if (numRowsToClear > 0) {
    var rangeToClear = sheet.getRange(headerRow + 1, 1, numRowsToClear, sheet.getLastColumn());
    rangeToClear.clearContent();
  }

  if (rowsToCopy.length > 0) {
    unprocessedRowExist = true;
    var startRow = headerRow + 1;  // Start from the row after the header
    var numRowsToInsert = rowsToCopy.length;
    var numColumns = rowsToCopy[0].length;
    var insertRange = sheet.getRange(startRow, 4, numRowsToInsert, numColumns);

    // Set the values from rowsToCopy into the sheet
    insertRange.setValues(rowsToCopy);
  }

  return unprocessedRowExist;
}


// this I don't use
function deleteFURowsBasedOnConditions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var headerRow = 5
  // Start processing rows from the header onwards (row 6)
  for (var i = headerRow+1; i <= lastRow; i++) {
    var cellD = sheet.getRange(i, 4).getValue();
    var cellJ = sheet.getRange(i, 10).getValue();
    var cellK = sheet.getRange(i, 11).getValue();
    var cellL = sheet.getRange(i, 12).getValue();
    var cellM = sheet.getRange(i, 13).getValue();
    

    if ((cellD && (cellJ === '' && cellK === '' && cellL === '')) || cellM === 'DB processed') {
      
      // Check if there's only one row below the header, clear it and exit
      if (i === lastRow && lastRow === 6) {
      sheet.getRange(i, 1, 1, sheet.getLastColumn()).clear();
      return;
      }

      sheet.deleteRow(i);
      // Since we deleted a row, we need to adjust the loop counter
      lastRow--;
      i--;
    }
  }
}



function deleteRowsBasedOnConditions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var callCenterSheet = ss.getSheetByName("Call Center");

  if (!callCenterSheet) {
    Logger.log("Call Center sheet not found.");
    return;
  }

  var dataRange = callCenterSheet.getDataRange();
  var values = dataRange.getValues();
  var rowsToDelete = [];

  for (var i = values.length - 1; i >= 0; i--) {
    var callResult = values[i][8];  // column I is callResult
    var colLValue = values[i][11];  // column L is "DB processed"
    var colJValue = values[i][9];  // column J is the date/time column

    // Check conditions for deletion
    if (colLValue === "DB processed") {

      if ((callResult === "Programat" || callResult === "De revenit") && isValidDateTime(colJValue)) {
      rowsToDelete.push(i + 1);  // Adding 1 to convert to 1-indexed row number
      }
      if (callResult === "Nu raspunde" || callResult === "Refuza") {
      rowsToDelete.push(i + 1);  // Adding 1 to convert to 1-indexed row number
      }
    }
  }
  // Delete rows
  for (var j = 0; j < rowsToDelete.length; j++) {
    callCenterSheet.deleteRow(rowsToDelete[j]);
  }
}

function isValidDateTime(value) {
  return !isNaN(Date.parse(value));
  
}