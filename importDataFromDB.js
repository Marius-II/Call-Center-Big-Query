function importDataFromDB() {
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
  var callCenterFollowUpSheet = spreadsheet.getSheetByName("Follow Up - CC")
  var callCenterFollowUpSheetName = callCenterFollowUpSheet.getName()
  if (spreadsheet.getActiveSheet().getName() === "Periodic"){
    var periodicSheet = spreadsheet.getSheetByName("Periodic")
    var periodicSheetName = periodicSheet.getName()
  } else if (spreadsheet.getActiveSheet().getName() === "Istoric"){
    var periodicSheet = spreadsheet.getSheetByName("Istoric")
    var periodicSheetName = periodicSheet.getName()    
  }


  var periodicFollowUpSheet = spreadsheet.getSheetByName("Follow Up - P")
  var periodicFollowUpSheetName = periodicFollowUpSheet.getName()

  var activeSheet = spreadsheet.getActiveSheet()
  var activeSheetName = activeSheet.getName();
  
  //var dataBaseArray = createDataBaseArray();

  var sheets = [callCenterFollowUpSheetName,periodicSheetName,periodicFollowUpSheetName];
  var resultsArray = []
  if (sheets.includes(activeSheetName )) {
    switch (activeSheetName ){

      case callCenterFollowUpSheetName:

        var deRevenitCallCenter = deRevenit(activeSheetName , callCenterFollowUpSheetName, periodicFollowUpSheetName)
        var nuRaspundeArray = nuRaspunde ()
        var nuRaspunde2Array = nuRaspunde2 ()

        resultsArray = [deRevenitCallCenter,nuRaspundeArray,nuRaspunde2Array]
        insertResults(resultsArray, activeSheet)
        break;
      case periodicSheetName:

        var oldInteractionArray = getOldInteraction(periodicSheet)
        resultsArray = [oldInteractionArray]
        insertResults(resultsArray, activeSheet)
      break;
      case periodicFollowUpSheetName:

        var deRevenitCallCenter = deRevenit(activeSheetName , callCenterFollowUpSheetName, periodicFollowUpSheetName)
        var nuRaspunde4Array = nuRaspunde4 ()
        var nuRaspunde5Array = nuRaspunde5 ()

        resultsArray = [deRevenitCallCenter, nuRaspunde4Array,nuRaspunde5Array]
        insertResults(resultsArray, activeSheet)
      break;
      
      default:
      break;

    }
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

function getOldInteraction(followUpSheet){
  
  var resultsArray = [];
  
  var howManyDaysAgoMin = getSheetNumber(followUpSheet,"h3");
  var howManyDaysAgoMax = getSheetNumber(followUpSheet,"g3");
  
  var resultsArray = fetchOldResultsBigQuery (howManyDaysAgoMax, howManyDaysAgoMin)
  resultsArray = rearangeForInsertion (resultsArray)

  // Sort the array based on date
  resultsArray.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });
  
  return resultsArray;
}

function getSheetNumber(sheet,cell){
  if (!sheet) {
  Logger.log("Sheet '" + sheet.getName() + "' not found.");
  return;
  }
  var daysNumber = sheet.getRange(cell).getValue();
  return daysNumber;
}

function insertResults(resultsArray, activeSheet){
  var headerRow = 5
  var insertionRow = getInsertionRow(headerRow);

  for (var i = 0; i < resultsArray.length; i++) {
    var array = resultsArray[i]
    var insertionRow = getInsertionRow(headerRow);
    
    if (array.length > 0) {
      var numRowsToInsert = array.length;
      var numColumns = array[0].length;
      var insertRange = activeSheet.getRange(insertionRow, 4, numRowsToInsert, numColumns);
      
      insertRange.setValues(array);  
    }

  }

}

function deRevenit(activeSheet, callCenterFollowUpSheet, periodicFollowUpSheet) {
  var dataBaseArray = fetchFromBigQuery("De revenit")
  var resultsArrayCallCenter = []
  var checkFutureDeRevenit = []
  var resultsArrayPeriodic = []
  var today = new Date;

  for (var i = 0; i < dataBaseArray.length; i++) {

    var historyDetails = dataBaseArray[i][6]
    var typeOfPacient = returnTypeOfPacient (historyDetails)
    
    var resultSlice = dataBaseArray[i];
    var programedDate = new Date (dataBaseArray[i][4])

    if(programedDate <= today){
      if (typeOfPacient === "Call Center"){
        resultsArrayCallCenter.push(resultSlice)
      }else{
        resultsArrayPeriodic.push(resultSlice)
      }

    }else if (programedDate > today){
      checkFutureDeRevenit.push(resultSlice)
    }else if (!programedDate){
      Logger.log(dataBaseArray[i][4])
    }
  

  }

  resultsArrayCallCenter = rearangeForInsertion(resultsArrayCallCenter)
  resultsArrayPeriodic = rearangeForInsertion(resultsArrayPeriodic)
  checkFutureDeRevenit = rearangeForInsertion(checkFutureDeRevenit)

  // Sort the array based on date
  resultsArrayCallCenter.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });

  resultsArrayPeriodic.sort(function(a, b) {
    return new Date(a[3]) - new Date(b[3]);
  });

  if (activeSheet === callCenterFollowUpSheet){
    return resultsArrayCallCenter
  } else if (activeSheet === periodicFollowUpSheet){
    return resultsArrayPeriodic
  }
}

function nuRaspunde() {
  var lastResult = "Nu raspunde";
  var resultsArray = fetchFromBigQuery(lastResult);
  resultsArray = rearangeForInsertion (resultsArray)

  // Sort the array based on the 6th element (index 5)
  resultsArray.sort(function(a, b) {
    if (a[5] < b[5]) {
      return -1;
    }
    if (a[5] > b[5]) {
      return 1;
    }
    return 0;
  });

  return resultsArray;
}

function nuRaspunde2 () {
  var lastResult = "Nu raspunde 2";
  var resultsArray = fetchFromBigQuery(lastResult);
  if (resultsArray && Array.isArray(resultsArray)){
    resultsArray = rearangeForInsertion (resultsArray)
    
    // Sort the array based on the 6th element (index 5)
    resultsArray.sort(function(a, b) {
      if (a[5] < b[5]) {
        return -1;
      }
      if (a[5] > b[5]) {
        return 1;
      }
      return 0;
    });

    return resultsArray;
  }else{
    return []
  }



}

function nuRaspunde4 () {
  var lastResult = "Nu raspunde 4";
  var resultsArray = fetchFromBigQuery(lastResult);

  // Ensure resultsArray is defined and is an array; initialize as empty array if not
  if (!Array.isArray(resultsArray)) {
    resultsArray = [];
  }
  resultsArray = rearangeForInsertion (resultsArray)
  
  if (resultsArray.length === 0) {
    // Return the empty array without attempting to sort
    return resultsArray;
  }
  // Sort the array based on the 6th element (index 5)
  resultsArray.sort(function(a, b) {
    if (a[5] < b[5]) {
      return -1;
    }
    if (a[5] > b[5]) {
      return 1;
    }
    return 0;
  });

  return resultsArray;
}

function nuRaspunde5 () {
  var lastResult = "Nu raspunde 5";
  var resultsArray = fetchFromBigQuery(lastResult);
  
  // Ensure resultsArray is defined and is an array; initialize as empty array if not
  if (!Array.isArray(resultsArray)) {
    resultsArray = [];
  }
  resultsArray = rearangeForInsertion (resultsArray)
  
  if (resultsArray.length === 0) {
    // Return the empty array without attempting to sort
    return resultsArray;
  }
  // Sort the array based on the 6th element (index 5)
  resultsArray.sort(function(a, b) {
    if (a[5] < b[5]) {
      return -1;
    }
    if (a[5] > b[5]) {
      return 1;
    }
    return 0;
  });

  return resultsArray;
}

function rearangeForInsertion(array){
  var timeElapsed;
  var resultsArray = [];

  for (var i = 0; i < array.length; i++){

    var codeFromDB = array[i][0];
    var name = array[i][1];
    var phoneNumber = array[i][2];
    var location = array[i][3];
    var lastInteractionTime = array[i][4];
    var callResultDB = array[i][5];
    var conversationDetailsDB = array[i][6];
    var whatsappMessage = ""

    var message = `👋 Bună ziua!\n\n🎉 Vrem să-ți reamintim că au trecut 12 luni de la ultimul tău control al vederii la Terra Optic. Este momentul perfect pentru o verificare gratuită! 🕒\n\n🔍 Pentru a-ți menține sănătatea vederii în cea mai bună formă, te invităm să profiți de controlul optometric gratuit oferit în cabinetele noastre. Tot ce trebuie să faci este să apeși pe linkul de mai jos și să te programezi în cel mai apropiat cabinet Terra Optic:\n\n👉 Programează-te aici: https://terraoptic.ro/programari_online\n\nTe așteptăm cu drag și nerăbdare să te vedem! 💙`;
    
    // Creating the WhatsApp message link
    var whatsappMessage = `=HYPERLINK("https://web.whatsapp.com/send?phone=+40${codeFromDB}&text=${encodeURIComponent(message)}", "Send WApp Mes")`;

    // Format the last interaction time using the provided function
    var formattedInteractionTime = formatDateToCustomFormat(lastInteractionTime);
    
    // Calculate the elapsed time using the provided function
    timeElapsed = calculateElipsedTime(lastInteractionTime);
    
    // Create the results array slice
    var resultsArraySlice = [
      phoneNumber,
      name,
      location,
      formattedInteractionTime,
      timeElapsed,
      callResultDB,
      '', // empty for newResult
      '', // empty for data/ora programarii
      '', // empty for insertion of new detail
      '', // empty status DB
      codeFromDB,
      conversationDetailsDB,
      whatsappMessage
    ];
    resultsArray.push(resultsArraySlice)
  }

  return resultsArray;
}

function returnTypeOfPacient (inputString) {

  var typeOfPacient = "Call Center"
  
  if (inputString != "" & inputString.includes("Periodic")){
    typeOfPacient = "Periodic"
  }

  return typeOfPacient;
}

function calculateElipsedTime(initialDate){
  var currentTime = new Date();
  initialDate = new Date(initialDate)

  var timePassed = Math.floor((currentTime - initialDate) / (1000 * 60 * 60 * 24)) + ' zile și ' + Math.floor((currentTime - initialDate) % (1000 * 60 * 60 * 24) / (1000 * 60 * 60)) + ' ore';
  return timePassed;
}

function deleteFURowsBasedOnConditions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var headerRow = 5
  // Start processing rows from the header onwards (row 6)
  for (var i = headerRow+1; i <= lastRow; i++) {
    var cellD = sheet.getRange(i, 4).getValue(); //0
    var cellJ = sheet.getRange(i, 10).getValue(); //6
    var cellK = sheet.getRange(i, 11).getValue(); //7
    var cellL = sheet.getRange(i, 12).getValue(); //8
    var cellM = sheet.getRange(i, 13).getValue(); //9
    

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

function getInsertionRow (headerRow){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = spreadsheet.getActiveSheet();
  var lastRow = activeSheet.getLastRow();

  if (lastRow > headerRow){
    var insertionRow = lastRow + 2
  } else{
    var insertionRow = lastRow + 1
  }
  return insertionRow;
}
