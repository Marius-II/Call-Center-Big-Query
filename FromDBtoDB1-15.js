function copyRowsToAnotherSpreadsheet() {
  // Source spreadsheet and sheet name
  var sourceSpreadsheet = SpreadsheetApp.openById("1C8h1F08Me8a1oSHAMAxW5X_-zj-skiyliXNHBKxBvJU");
  var sourceSheetName = "Sheet1";
  
  // Destination spreadsheet and sheet name prefix
  var destinationSpreadsheet = SpreadsheetApp.openById("1T74mRoyG5YsNqyT1Q2rFPKnC2yEczrM6iHyS5DW_Eig");
  var destinationSheetNamePrefix = "DB";

  // Number of rows to copy at once
  var batchSize = 3500;

  // Get the data range in the source sheet
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  var lastRow = sourceSheet.getLastRow();
  var numRowsCopied = 0;

  // Loop through the rows in batches and copy to destination sheets
  for (var startRow = 2; startRow <= lastRow; startRow += batchSize) {
    var endRow = Math.min(startRow + batchSize - 1, lastRow);
    var numRowsInBatch = endRow - startRow + 1;
    
    // Get the data in the batch
    var dataRange = sourceSheet.getRange(startRow, 1, numRowsInBatch, sourceSheet.getLastColumn());
    var batchValues = dataRange.getValues();
    
    // Determine the destination sheet name
    var destinationSheetName = destinationSheetNamePrefix + Math.ceil(startRow / batchSize);

    // Get the destination sheet or create it if it doesn't exist
    var destinationSheet = destinationSpreadsheet.getSheetByName(destinationSheetName);
    if (!destinationSheet) {
      destinationSheet = destinationSpreadsheet.insertSheet(destinationSheetName);
    }
    
    // Copy the batch to the destination sheet
    destinationSheet.getRange(destinationSheet.getLastRow() + 1, 1, numRowsInBatch, dataRange.getLastColumn()).setValues(batchValues);
    
    numRowsCopied += numRowsInBatch;
  }

  // Log the number of rows copied
  Logger.log("Copied " + numRowsCopied + " rows in total.");
}
