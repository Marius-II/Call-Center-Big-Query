function mainDelete() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var callCenterSheet = spreadsheet.getActiveSheet()
  var sheetName = callCenterSheet.getName();
  var cell = callCenterSheet.getActiveCell()


  //Check if there is a row selected
  var isActiveCell = checkIfRowSelectedIsValid(callCenterSheet, cell);
  if (!isActiveCell){
    showUserMessage("Nu ai selectat un rand valid")
   return;
  }

  //Extract details
  var rowData = extractRowDetails (callCenterSheet, cell);
  var codeValue = extractCodeValue(rowData,sheetName).toString()
  var clientName = rowData[1]
  
  var deleteReason = colectDeleteReason(clientName)

  
  if (deleteReason){
    if (deleteReason === "error"){
      return;
    }else{

      appendRowToRanduriSterse (spreadsheet, sheetName, rowData, cell, deleteReason);
      
      deleteFromBigQuery (codeValue)
      
      //deleteRow from call center sheet
      callCenterSheet.deleteRow(cell.getRow())

      //Show message to user
      showUserMessage("Pacientul a fost sters cu succes. Pentru a vizualiza randurile sterse mergeti la sheetul - Randuri Sterse")

    }
  }
}

function checkIfRowSelectedIsValid(callCenterSheet, cell) {
  var startingCell = callCenterSheet.getRange("D6");

  // Calculate the end cell based on the sheet's dimensions
  var lastRow = callCenterSheet.getLastRow();
  var lastColumn = callCenterSheet.getLastColumn();
  
  // Get the row and column of the cell
  var cellRow = cell.getRow();
  var cellColumn = cell.getColumn();

  // Check if the cell is within the desired range
  return cellRow >= startingCell.getRow() && cellColumn >= startingCell.getColumn() &&
    cellRow <= lastRow && cellColumn <= lastColumn;
}




function extractRowDetails(callCenterSheet, cell) {

  // Extract details from the row
  var columns = callCenterSheet.getLastColumn() - 3
  var rowData = callCenterSheet.getRange(cell.getRow(), 4, 1, columns).getValues()[0];

  return rowData;
}

function extractCodeValue(rowData,sheetName){
  var codeValue;

  if (sheetName === "Call center"){
    codeValue = rowData[9];
  }else{
    codeValue = rowData[10];
  }
  return codeValue;
}


function appendRowToRanduriSterse(spreadsheet, sheetName, rowData, cell, deleteReason) {

  var sourceSheet = spreadsheet.getSheetByName(sheetName);
  var destinationSheet = spreadsheet.getSheetByName('Randuri Sterse');
  var userName = Session.getActiveUser().getEmail();
  var rowNumber = cell.getRow();

  var code = rowData[10]
  var name = rowData[1]
  var phone = rowData[0]
  var location = rowData[2]
  var lastInteractionDateTime = rowData[3]
  var lastResult = rowData[5]
  var historyDetails = rowData[11]

  var dataArray = [[code,name,phone,location,lastInteractionDateTime,lastResult,historyDetails]]


  if (sourceSheet && destinationSheet) {

    var lastDestinationRow = destinationSheet.getLastRow() + 1;
    var currentDate = new Date();

    // Get the range in the last empty row of the destination sheet
    var destinationRange = destinationSheet.getRange(lastDestinationRow, 1, 1, 7);
    destinationRange.setValues(dataArray)

    var deleteReasonRange = destinationSheet.getRange(lastDestinationRow, 8);
    var userNameRange = destinationSheet.getRange(lastDestinationRow, 9);
    var deleteDateRange = destinationSheet.getRange(lastDestinationRow,10);

    // Copy the row to the last empty row in the "Randuri Sterse" sheet

    deleteReasonRange.setValue(deleteReason);
    userNameRange.setValue(userName);
    deleteDateRange.setValue(formatDateToCustomFormat(currentDate))


    console.log('Row ' + rowNumber + ' appended from ' + sheetName + ' to Randuri Sterse');
  } else {
    console.log('Source or destination sheet not found.');
  }
}

function colectDeleteReason(numeClient) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Clientul: " + numeClient + " va fi sters din baza de date. Introdu motivul stergerii");
  
  //Get the button that the user pressed.
  var button = result.getSelectedButton();
  
  if (button === ui.Button.OK) {
    Logger.log("The user clicked the [OK] button.");
    Logger.log(result.getResponseText());
    var deleteReason = result.getResponseText()
    return deleteReason;
    
  } else if (button === ui.Button.CLOSE) {
    Logger.log("The user clicked the [X] button and closed the prompt dialog.");
    return "error"
  }
    
}





