function importDataSearchModule() {
  var startTime = Date.now();
  var unprocessedRowExist = clearAllRowsFromHeaderOnwards(5);
  var message = `Exista randuri partial completate, adica randuri in care au fost introduse date in coloana pentru 'Rezultatul nou', 'Detalii' sau 'Data programarii', iar acestea nu au fost procesate.

  \n1. Completeaza 'Rezultatul nou' si 'Data/Ora programarii sau pentru revenire' pentru toate randurile ramase.
  2. Proceseaza datele din nou`;

  if (unprocessedRowExist){
    Browser.msgBox('Mesaj pentru utilizator', message, Browser.Buttons.OK);
    return;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var headerRow = 5

  var activeSheet = spreadsheet.getActiveSheet();

  var searcheableCodeRange = activeSheet.getRange("D3")
  var searcheableNameRange = activeSheet.getRange("E3")

  var searcheableCode = searcheableCodeRange.getValue()
  var searcheableName = searcheableNameRange.getValue()

  if(searcheableCode || searcheableName){
    if (searcheableCode){
      var results = searchByCodeInBigQuery(searcheableCode)
    }else{
      var results = searchByNameInBigQuery(searcheableName)
    }
  }
  //to be compatible with the insert result function the array must be an array of arrays
  if (results){
    results = rearangeForInsertion(results)
    
    var resultsArray = [results]
    insertResults(resultsArray, activeSheet)
  }else{
    showUserMessage(`Nu au fost gasite rezultate conform cerintelor cautarii`)
  }



  try{
    copyFormatFromRow5ToRow6Onwards(headerRow)
  } catch(error){
    Logger.log('Error: ' + error);
  }

  var endTime = Date.now();
  var elapsedTime = endTime - startTime;
  Logger.log('Elapsed Time: ' + elapsedTime + ' ms');
}
