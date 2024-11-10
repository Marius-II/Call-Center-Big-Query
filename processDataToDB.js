let answerCountTowns = {};
let locationResultCounts = {};
let locationSourceResultCounts = {};
let answerCount = {};
var counterProcessedProgramat = 0;
var counterProcessedDeRevenit = 0;
var counterProcessedNuRaspunde = 0;
var counterProcessedRefuza = 0; 
var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
Logger.log(activeSheet.getName());

function processRows() {
  var followUpSheetsNamesArray = ["Periodic", "Follow Up - P", "Follow Up - CC", "Cautare 1", "Cautare 2", "Istoric"];
  var dateProcessButtonPressed = new Date();
  var dataArrayFromCallCenterSheet = extractDataFromCallCenterSheet(5);

  var activeSheetName = activeSheet.getName();

  for (var i = 0; i < dataArrayFromCallCenterSheet.length; i++) {
    var index = 0;
    if (followUpSheetsNamesArray.includes(activeSheetName)) {
      index = 1;
    }
    var phoneNumber = dataArrayFromCallCenterSheet[i][0];
    var name = dataArrayFromCallCenterSheet[i][1];
    var location = dataArrayFromCallCenterSheet[i][2];
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
    var source = dataArrayFromCallCenterSheet[i][11 + index];
    var rowUpdated = false;

    if (callResult !== '' && reInteractTime !== '' && statusDB !== "DB processed") {
      // Search by code in BigQuery
      var rowToBeModified = searchByCodeInBigQuery(codeFromCallCenter);

      if (rowToBeModified) {
        updateRow(rowToBeModified, codeFromCallCenter, name, phoneNumber, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheet(i);
      } else {
        insertNotFoundedRows(codeFromCallCenter, phoneNumber, name, location, reInteractTime, callResult, conversationDetails);
        updateRowInCallCenterSheet(i);
      }
      var startTime = Date.now();
      countCallCenterResults(activeSheetName, callResult, location, source);
      var endTime = Date.now();
      var timeTaken = endTime - startTime;
      Logger.log("Time taken: count result " + timeTaken + " milliseconds");
    }
  }
  console.log(answerCount);


  if (followUpSheetsNamesArray.includes(activeSheetName)) {
    updateFollowUpLogWithCounts(dateProcessButtonPressed);
  } else if (activeSheetName === "Call Center") {
    updateCallCenterLogWithCounts(dateProcessButtonPressed);
    logCountByLocationCallCenter();
  }
}

function countCallCenterResults(activeSheetName, callResult, location, source) {
  location = transformLocation(location);

  if (activeSheetName === "Call Center") {
    // Initialize the location object if it doesn't exist
    if (!locationSourceResultCounts[location]) {
      locationSourceResultCounts[location] = {};
    }

    // Initialize the source object for the location if it doesn't exist
    if (!locationSourceResultCounts[location][source]) {
      locationSourceResultCounts[location][source] = {
        "Programat": 0,
        "De revenit": 0,
        "Nu raspunde": 0,
        "Refuza": 0,
      };
    }

    // Count overall results and increment the location-source-specific counts
    switch (callResult) {
      case "Programat":
        counterProcessedProgramat++;
        locationSourceResultCounts[location][source]["Programat"]++;
        break;
      case "De revenit":
        counterProcessedDeRevenit++;
        locationSourceResultCounts[location][source]["De revenit"]++;
        break;
      case "Nu raspunde":
        counterProcessedNuRaspunde++;
        locationSourceResultCounts[location][source]["Nu raspunde"]++;
        break;
      case "Refuza":
        counterProcessedRefuza++;
        locationSourceResultCounts[location][source]["Refuza"]++;
        break;
      default:
        break;
    }
  }
}



function logCountByLocationCallCenter() {
  // Open the sheet by name

  const spreadsheet = SpreadsheetApp.openById("1yc5s-keQtGyrBEqeZlHcNsjB-c-E63OBKUue1befRnc");
  const sheet = spreadsheet.getSheetByName('CountByLocationCallCenter');
  
  // Get the current date
  const currentDate = new Date();
  
  // Starting row for logging (you can adjust this if needed)
  let startRow = sheet.getLastRow() + 1;
  
  // Loop through locationResultCounts and write data to the sheet
  // Loop through locationSourceResultCounts and log data
  for (let location in locationSourceResultCounts) {
    for (let source in locationSourceResultCounts[location]) {
      let result = locationSourceResultCounts[location][source];

      // Write data to the sheet
      sheet.getRange(startRow, 1).setValue(currentDate);            // Column A: Date
      sheet.getRange(startRow, 2).setValue(location);               // Column B: Location
      sheet.getRange(startRow, 3).setValue(source);                 // Column C: Source
      sheet.getRange(startRow, 4).setValue(result["Programat"]);       // Column D: Programat
      sheet.getRange(startRow, 5).setValue(result["De revenit"]);   // Column E: De revenit
      sheet.getRange(startRow, 6).setValue(result["Nu raspunde"]);  // Column F: Nu raspunde
      sheet.getRange(startRow, 7).setValue(result["Refuza"]);          // Column G: Refuza

      // Move to the next row for each new location-source entry
      startRow++;
    }
  }
}

// // Example function to count the results and then log them
// function countAndLogCallCenterResults(activeSheetName, callResult, location) {
//   // Count the results first
//   countCallCenterResults(activeSheetName, callResult, location);
  
//   // Log results to the sheet
//   logResultsToSheet();
// }


function updateRow(rowToBeModified, codeFromCallCenter, name, phoneNumber, location, reInteractTime, callResult, conversationDetails) {
  var callResultDB = rowToBeModified[0][5];
  var conversationDetailsDB = rowToBeModified[0][6];
  var newCallResult = setNewCallResult(callResult, callResultDB);
  var newConversationDetails = setNewConversationDetails(newCallResult, conversationDetails, conversationDetailsDB);
  var reInteractTime = convertTimestampToDateTime(reInteractTime);

  updateRowInBigQuery(codeFromCallCenter, name, phoneNumber, location, reInteractTime, newCallResult, newConversationDetails);
}

function insertNotFoundedRows(codeFromCallCenter, phoneNumber, name, location, lastInteractionTime, callResult, conversationDetails) {
  var newConvDetails = setNewConversationDetails(callResult, conversationDetails, "");
  var lastInteractionDateTime = convertTimestampToDateTime(lastInteractionTime);

  insertRowToBigQuery({
    code: codeFromCallCenter,
    fullName: name,
    phoneNumber: phoneNumber,
    source: location,
    lastInteractionDate: lastInteractionDateTime,
    lastResult: callResult,
    conversationHistory: newConvDetails
  });
}

function convertTimestampToDateTime(timestamp) {
  var date = new Date(timestamp);
  var formattedDateTime = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  return formattedDateTime;
}

function updateRowInCallCenterSheet(rowNumberInDataSet) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
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

function setNewCallResult(callResult, callResultDB) {
  var newCallResult;
  switch (callResult) {
    case 'Nu raspunde':
      newCallResult = handleNuRaspunde(callResultDB);
      countAnswerPairs(callResultDB, callResult);
      break;
    case 'De revenit':
    case 'Refuza':
    case 'Programat':
      newCallResult = callResult;
      countAnswerPairs(callResultDB, callResult);
      break;
    default:
      newCallResult = callResult;
  }
  return newCallResult;
}

function setNewConversationDetails(newCallResult, conversationDetails, conversationDetailsDB) {
  var currentDateTime = new Date();
  var formattedDateTime = formatDateToCustomFormat(currentDateTime);
  var newConversationDetailsTitle = formattedDateTime + " - Raspuns: " + newCallResult;
  var newConversationDetails = newConversationDetailsTitle + "\n" + conversationDetails + "\n\n" + conversationDetailsDB;
  return newConversationDetails;
}

function formatDateToCustomFormat(inputDate) {
  var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var inputDateTime = new Date(inputDate);

  var day = inputDateTime.getDate();
  var month = months[inputDateTime.getMonth()];
  var year = inputDateTime.getFullYear();
  var hours = inputDateTime.getHours();
  var minutes = inputDateTime.getMinutes();
  var seconds = inputDateTime.getSeconds();

  day = day < 10 ? "0" + day : day;
  hours = hours < 10 ? "0" + hours : hours;
  minutes = minutes < 10 ? "0" + minutes : minutes;
  seconds = seconds < 10 ? "0" + seconds : seconds;

  return day + "-" + month + "-" + year + " " + hours + ":" + minutes + ":" + seconds;
}

// pentru aceasta functie "callResultDB" este rezultatul vechi din baza de date, iar rezultatul nou este "Nu raspunde" (acesta fiind motivul pentru care se activeaza functia)
// Daca rezultatul vechi incepe cu Nu raspunde, e simplu, se adauga la rezultatul vechi un numar
// Daca rezultatul vechi este de revenit rezultatul nou va ramane de revenit si se va afisa pe viitor in seciunea de revenit
// Daca rezultatul vechi este "Programat" sau "Refuza" intervine problema. Dar daca rezultatul vechi este "Programat" sau "Refuza" atunci e clar ca el are un istoric de cel putin un an, asa cum functioneaza programul acum, in Noiembrie 2024. Deci daca ii alocam un rezultat de "Nu raspunde 4", acesta se va duce in Folow Up - P si nu in Follow up CC. In acest moment, toate persoanele care au Nu raspunde, Nu raspunde 2 sunt importate in follow up CC fara a se tine cont de istoric. Practic daca o persoana i se aloca dupa un an Nu raspunde, iar ea a fost programata ultima data, aceasta va avea rezultatul nou "Nu raspunde" si ea in loc sa fie importata in "Follow Up periodic, va fi importata in Follow Up CC". Rezolvarea este ca la aceste persoane, pentru ca au deja istoric, daca nu raspund, sa le alocam un nou rezultat de "Nu raspunde 4", astel acestea vor fi importate in Follow up periodic
function handleNuRaspunde(callResultDB) {
  var callResultsArray = ["Programat", "De revenit", "Nu raspunde", "Nu raspunde 2", "Nu raspunde 3", "Nu raspunde 4", "Nu raspunde 5", "Nu raspunde 6", "Nu raspunde 7", "Nu raspunde 8", "Refuza"];
  if (callResultDB.toString().startsWith("Nu raspunde")) {
    var increment = parseInt(callResultDB.split(' ')[2]) || 1;
    increment++;
    callResultDB = "Nu raspunde " + increment;
  } else if (callResultsArray.includes(callResultDB)) {
    callResultDB = "Nu raspunde 4"; // aici am modificat
  } else {
    callResultDB = "Nu raspunde 4";
  }
  return callResultDB;
}

function extractDataFromCallCenterSheet(headerRow) {
  var dataArrayFromCallCenterSheet = [];
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var callCenterSheet = spreadsheet.getActiveSheet();
  
  if (callCenterSheet) {
    var lastRow = callCenterSheet.getLastRow();
    var numRowsToExtract = lastRow - headerRow;
    if (numRowsToExtract > 0) {
      var lastColumn = callCenterSheet.getLastColumn();
      var data = callCenterSheet.getRange(headerRow + 1, 4, lastRow - headerRow, lastColumn - 3).getValues();
      for (var i = 0; i < data.length; i++) {
        dataArrayFromCallCenterSheet.push(data[i]);
      }
    }
  } else {
    Logger.log('Call Center sheet not found.');
  }

  return dataArrayFromCallCenterSheet;
}

function updateCallCenterLogWithCounts(dateProcessButtonPressed) {
  const spreadsheet = SpreadsheetApp.openById("1yc5s-keQtGyrBEqeZlHcNsjB-c-E63OBKUue1befRnc");
  const sheet = spreadsheet.getSheetByName("Log Procesari CallCenter");
  const lastRow = sheet.getLastRow() + 1;

  let date = formatDateToCustomFormat(dateProcessButtonPressed);
  let total = counterProcessedProgramat + counterProcessedDeRevenit + counterProcessedNuRaspunde + counterProcessedRefuza;

  sheet.getRange(lastRow, 5).setValue(total);
  sheet.getRange(lastRow, 4).setValue(date);
  sheet.getRange(lastRow, 6).setValue(counterProcessedProgramat);
  sheet.getRange(lastRow, 7).setValue(counterProcessedDeRevenit);
  sheet.getRange(lastRow, 8).setValue(counterProcessedNuRaspunde);
  sheet.getRange(lastRow, 9).setValue(counterProcessedRefuza);
}

function updateFollowUpLogWithCounts(dateProcessButtonPressed) {
  const spreadsheet = SpreadsheetApp.openById("1yc5s-keQtGyrBEqeZlHcNsjB-c-E63OBKUue1befRnc");
  const sheet = spreadsheet.getSheetByName("Log Procesari FollowUp");
  const lastRow = sheet.getLastRow() + 1;

  let date = formatDateToCustomFormat(dateProcessButtonPressed);
  let total = 0;

  for (const pairKey in answerCount) {
    const count = answerCount[pairKey];
    const [oldAnswer, newAnswer] = pairKey.split("to");
    var columnIndex = getColumnIndex(sheet, `${oldAnswer}to${newAnswer}`);

    if (columnIndex !== -1) {
      // Column exists, add count to the existing column
      sheet.getRange(lastRow, columnIndex).setValue(count);
    } else {
      // Column doesn't exist, so append a new column
      const newColumnIndex = sheet.getLastColumn() + 1;
      // Set the header in row 5 for this transition
      sheet.getRange(5, newColumnIndex).setValue(`${oldAnswer}to${newAnswer}`);
      // Set the count in the current row
      sheet.getRange(lastRow, newColumnIndex).setValue(count);
    }
    
    total += count;
  }

  // Add date and total to the appropriate columns in lastRow
  sheet.getRange(lastRow, 4).setValue(date);
  sheet.getRange(lastRow, 5).setValue(total);
}


function getColumnIndex(sheet, title) {
  var headers = sheet.getRange(5, 6, 1, sheet.getLastColumn() - 5).getValues()[0];
  headers = headers.map(element => element.replace(/\n/g, " "));
  const columnIndex = headers.indexOf(title) + 6;
  return columnIndex;
}

function countAnswerPairs(oldAnswer, newAnswer) {
  const possibleOldAnswers = ["Programat", "De revenit", "Nu raspunde", "Refuza", "Nu raspunde 2", "Nu raspunde 3", "Nu raspunde 4", "Nu raspunde 5", "Nu raspunde 6", "Nu raspunde 7", "Nu raspunde 8", "Neinregistrat"];
  const possibleNewAnswers = ["Programat", "De revenit", "Nu raspunde", "Refuza"];

  if (!possibleOldAnswers.includes(oldAnswer)) {
    oldAnswer = "Neinregistrat";
  } else if (!possibleNewAnswers.includes(newAnswer)) {
    showUserMessage(`Ati introdus un rezultat (${newAnswer}) care este invalid`);
  }

  const pairKey = `${oldAnswer} to ${newAnswer}`;
  answerCount[pairKey] = (answerCount[pairKey] || 0) + 1;

  console.log(`Count for ${pairKey}: ${answerCount[pairKey]}`);
}
function transformLocation(location) {
  let lowerLocation = location.toLowerCase();
  if (lowerLocation === "brăila" || lowerLocation.includes("br mess") || lowerLocation === "braila") {
    location = "Braila";
  } else if (lowerLocation === "bv leads" || lowerLocation === "bv mess" || lowerLocation === "brasov" || lowerLocation === "brașov") {
    location = "Brasov";
  } else if (lowerLocation === "constanța" || lowerLocation === "ct mess" || lowerLocation === "constanta") {
    location = "Constanta";
  } else if (lowerLocation === "foc leads" || lowerLocation === "foc mess" || lowerLocation === "focsani" || lowerLocation === "focșani") {
    location = "Focsani";
  } else if (lowerLocation === "tecuci" || lowerLocation === "tec mess") {
    location = "Tecuci";
  } else if (lowerLocation === "galați" || lowerLocation === "gl mess" || lowerLocation === "galati") {
    location = "Galati";
  } else if (lowerLocation === "bucurești" || lowerLocation === "buc leads" || lowerLocation === "bucuresti" || lowerLocation === "buc mess") {
    location = "Bucuresti";
  } else if (lowerLocation === "selectati"){
    location = `z.${location}`
  }
  return location;
}
function transformSheetNameintoLocationName(location) {
  let lowerSheetName = location.toLowerCase();
  if (lowerSheetName === "brăila" || lowerSheetName.includes("br mess")) {
    location = "Braila";
  } else if (lowerSheetName === "bv leads" || lowerSheetName === "bv mess") {
    location = "Brasov";
  } else if (lowerSheetName === "constanța" || lowerSheetName === "ct mess") {
    location = "Constanta";
  } else if (lowerSheetName === "foc leads" || lowerSheetName === "foc mess") {
    location = "Focsani";
  } else if (lowerSheetName === "tecuci" || lowerSheetName === "tec mess") {
    location = "Tecuci";
  } else if (lowerSheetName === "galați" || lowerSheetName === "gl mess") {
    location = "Galati";
  }
  return location;
}

function showUserMessage(message) {
  Browser.msgBox('Message', message, Browser.Buttons.OK);
}
