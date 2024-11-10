let date = new Date
let countPerLocationSiteKeyValuePair = {};
let countPerLocationChatBootKeyValuePair = {};

function importData() {
  //get the spreadshhet source
  var sourceSpreadsheet = SpreadsheetApp.openById("1mLhQZGkz7MZsnmZ5Efcrzd9Ne_F1bpDbXmOLq4qrNJY")

  // Array of source sheet names
  var sourceSheetNames = ["Site", "ChatBoot Leads", "Constanța", "CT Mess", "Brăila", "Br Mess", "Tecuci", "Tec Mess", "Galați", "Gl Mess", "BV Leads", "BV Mess", "Foc Leads", "Foc Mess","Buc Leads","Buc Mess"];

  // Destination sheet name
  var destinationSheetName = "Call Center";
  // Get the destination sheet
  var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(destinationSheetName);

  noOfRowsToImport

  // Array to store processed data
  var processedDataArray = [];



  // Iterate through each source sheet
  for (var i = 0; i < sourceSheetNames.length; i++) {
    Logger.log("import from" + sourceSheetNames[i])
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetNames[i]);
    if(sourceSheet){
      var noOfRowsToImport = 200
      var lastRow = sourceSheet.getLastRow();
      var startRow = Math.max(2, lastRow - noOfRowsToImport + 1); // Start from the last 100 rows or the second row (exclude header)
      
      // Get the data range for the last 100 rows
      var dataRange = sourceSheet.getRange(startRow, 1, noOfRowsToImport, 12); // Include header row
      var data = dataRange.getValues();
      
      rearangeDataArrayBasedOnSheet(sourceSheetNames[i],data);
      var counter = 0
      
      if (sourceSheetNames[i] === "ChatBoot Leads"){
        Logger.log ("test here chgat boot")
      }
      
      for (var j = 0; j < data.length; j++) {
        
        var rezultat = data[j][10];
        var dateValue = data[j][0];

        // Process the date value
        if (typeof dateValue === 'string' && /\d{2}\/\d{2}\/\d{5}\s\d{2}:\d{2}:\d{2}/.test(dateValue)) {
          var parts = dateValue.match(/(\d{2})\/(\d{2})\/(\d{5})\s(\d{2}:\d{2}:\d{2})/);
          if (parts) {
            var year = parseInt(parts[3]) - 1;
            var month = parseInt(parts[2]) - 1;
            var day = parseInt(parts[1]);
            var time = parts[4];
            var formattedDate = new Date(Date.UTC(year, month, day, parseInt(time.split(':')[0]), parseInt(time.split(':')[1]), parseInt(time.split(':')[2])));
            formattedDate.setHours(formattedDate.getHours() + (formattedDate.getTimezoneOffset() / 60));
          } else {
            Logger.log("skipped invalid date at import" + sourceSheetNames[i])
            continue; // Skip invalid date format
          }
        } else {
          formattedDate = new Date(dateValue);
        }

        // Skip if date is not valid
        if (isNaN(formattedDate.getTime())) {
          continue;
        }

        // Process data only if Rezultat is empty and date is after September 19, 2023
        if (rezultat === '') {
          // Format the date as a string
          var formattedDateString = Utilities.formatDate(formattedDate, 'Europe/Bucharest', 'yyyy/MM/dd HH:mm:ss');
          var currentTime = new Date();
          var timePassed = Math.floor((currentTime - formattedDate) / (1000 * 60 * 60 * 24)) + ' zile și ' + Math.floor((currentTime - formattedDate) % (1000 * 60 * 60 * 24) / (1000 * 60 * 60)) + ' ore';
          var conversationDetails = data[j][4] + " - " + data[j][5];
          var location;

          if (sourceSheetNames[i] === "Site"){
            location = data[j][5];
            let noDiacriticsLocation = removeDiacritics(location)
            if(countPerLocationSiteKeyValuePair[noDiacriticsLocation]){
              countPerLocationSiteKeyValuePair[noDiacriticsLocation]++
            }else{
              countPerLocationSiteKeyValuePair[noDiacriticsLocation]=1
            }
          }else if (sourceSheetNames[i] === "ChatBoot Leads"){
            location = data[j][5];
            let noDiacriticsLocation = removeDiacritics(location)
            if(countPerLocationChatBootKeyValuePair[noDiacriticsLocation]){
              countPerLocationChatBootKeyValuePair[noDiacriticsLocation]++
            }else{
              countPerLocationChatBootKeyValuePair[noDiacriticsLocation]=1
            }
          }else{
            location = sourceSheetNames[i];
          }

          //process phone numbers like this "(0752) 487 445" to this "0752487445"
          var telString = data[j][3];
          
          try{
            var telStringFinal = telString.toString().replace(/[^0-9]/g, '');
          } catch(error){
            Logger.log("telString error for "+ telString+ " error is: "+ error);
          }

          var telStringFinal = telString.toString().replace(/[^0-9]/g, '');
          

          // Create an array with processed data
          var processedData = [
            telStringFinal , // Telefon
            data[j][2], // Nume
            location, // Locatie (sheet name)
            formattedDateString, // DataOraInteractiuni
            timePassed, // TimpScurs
            data[j][10], // RezultatulApelului
            data[j][11], // DataOraProgramare
            conversationDetails, // DetaliiConversatie
            '' // Verifica
          ];
          
          var telefonSlice = telStringFinal.slice(-9);

          // Check if telefonSlice is not already in the processedDataArray
          var isTelefonSliceAlreadyPresent = processedDataArray.some(function(processedData) {
            return processedData[9] === telefonSlice;
          });
          
          // If telefonSlice is not already present, push the processedData to the array
          if (!isTelefonSliceAlreadyPresent) {
            processedData.push(telefonSlice);
            processedData.push("") // left free for history
            processedData.push(sourceSheetNames[i])
            processedDataArray.push(processedData);
            counter++;
          }

          

          // Remove data validation from the specified range in each sheet
          sourceSheet.getRange(startRow + j, 11).clearDataValidations();

          // Process data and update column K ("Rezultat")
          sourceSheet.getRange(startRow + j, 11).setValue('Processed');
        }
      }
    }
    let sheetName = sourceSheetNames[i]
    countNewCustomersBySheet (counter,sheetName, location,countPerLocationSiteKeyValuePair,countPerLocationChatBootKeyValuePair);
    counter = 0
  }

  // Insert the processed data into Call Center sheet below header
  var headerRow = 5; // Header is in row 5
  var startRow = destinationSheet.getLastRow()+1; // Start inserting data from row 7

  // Batch update data in the sheet without inserting new rows
  if (processedDataArray.length > 0) {
    var numRowsToInsert = processedDataArray.length;
    var numColumns = processedDataArray[0].length;
    var insertRange = destinationSheet.getRange(startRow, 4, numRowsToInsert, numColumns);
    
    // Overwrite the existing values in the range with the processed data
    insertRange.setValues(processedDataArray);
    
  }

  try{
    copyFormatFromRow5ToRow6Onwards(headerRow)
  } catch(error){
    Logger.log('Error: ' + error);
  }
  
  deleteRowsBasedOnConditions();
  //insertHistoryLink();
}

function countNewCustomersBySheet (numberOfResults,sheetName,location,countPerLocationSiteKeyValuePair, countPerLocationChatBootKeyValuePair){

  //set date;
  date = formatDateToCustomFormat(date);

  //set location
  let lowerSheetName = sheetName.toLowerCase();
  if(lowerSheetName === "brăila" || lowerSheetName.includes("br mess")){
    location = "Braila"
  }else if(lowerSheetName === "bv leads" || lowerSheetName === "bv mess"){
    location = "Brasov"
  }else if(lowerSheetName === "constanța" || lowerSheetName === "ct mess"){
    location = "Constanta"
  }else if(lowerSheetName === "foc leads" || lowerSheetName === "foc mess"){
    location = "Focsani"
  }else if(lowerSheetName === "tecuci" || lowerSheetName === "tec mess"){
    location = "Tecuci"
  }else if(lowerSheetName === "galați" || lowerSheetName === "gl mess"){
    location = "Galati"
  }else if(lowerSheetName === "site" || lowerSheetName === "chatboot leads"){
    location = "" // left free because there is other function counting that
  }else if(lowerSheetName === "buc leads" || lowerSheetName === "buc mess"){
    location = "Bucuresti"
  }
    
  //count leads
  let leadsSheetArray = ["Brăila", "Tecuci", "Constanța", "Galați", "BV Leads", "Foc Leads", "Buc Leads"];
  
  let counterLeads = 0;
  let counterMess = 0;
  let counterChatBoot = 0;
  let counterSite = 0;
  let totalCountIn = 0;


  if (sheetName.endsWith("Mess")){
    counterLeads = 0
    counterMess = numberOfResults
  }else if(leadsSheetArray.includes(sheetName)){
    counterLeads = numberOfResults
    counterMess = 0
    if (countPerLocationSiteKeyValuePair[location]){
      counterSite = countPerLocationSiteKeyValuePair[location];
    }
    if (countPerLocationChatBootKeyValuePair[location]){
      counterChatBoot = countPerLocationChatBootKeyValuePair[location];;
    }
  }


  totalCountIn = counterLeads + counterMess + counterChatBoot + counterSite;


  let dataArray = [
    date,
    location,
    counterLeads,
    counterMess,
    counterChatBoot,
    counterSite,
    totalCountIn,
  ]

  //set sheet and range to insert
  if(lowerSheetName != "site"){

    let kpiLogSpreadSheet = SpreadsheetApp.openById("1yc5s-keQtGyrBEqeZlHcNsjB-c-E63OBKUue1befRnc")

    let kpiLogSheet = kpiLogSpreadSheet.getSheetByName("Log Intrari Call Center");

    let kpiLogCustomerRange = kpiLogSheet.getRange(kpiLogSheet.getLastRow()+1,4,1,dataArray.length);
    kpiLogCustomerRange.setValues([dataArray]);
  }

}

function insertHistoryLink() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var lastRow = sheet.getLastRow();

  for (var row = 6; row <= lastRow; row++) {
    var code = sheet.getRange(row, 13).getValue();
    var rowsResults = searchByCodeInBigQuery(code);

    if (rowsResults){
      var firstRow = rowsResults[0];
      var details = firstRow[6]
    }

    var cell = sheet.getRange(row, 14); // Column N, change if needed
    cell.setValue(details)
  }
}

function rearangeDataArrayBasedOnSheet (sourceSheetName, data) {
  Logger.log (sourceSheetName)
  switch (sourceSheetName) {
    case "Site":
      for (var i = 0; i < data.length; i++) {
        var extractedString = data[i][2];

        // Extract variables using regular expressions
        var locationVariable = extractVariable(extractedString, "Locatie");
        var nameVariable = extractVariable(extractedString, "Nume si Prenume");
        var emailVariable = extractVariable(extractedString, "Email");
        var phoneVariable = extractVariable(extractedString, "Telefon");
        var messageVariable = extractVariable(extractedString, "Mesaj");

        // Update the data array with extracted variables
        data[i][2] = nameVariable// Nume
        data[i][3] = phoneVariable // Telefon
        data[i][4] = messageVariable + emailVariable// detalii 1
        
        data[i][5] = removeDiacritics(locationVariable) //detalii 2

      }
      break;

    default:
      return data;
  }

  return data;
}

function removeDiacritics(text) {
  var diacriticsMap = {
    'ă': 'a', 'Ă': 'A', 'â': 'a', 'Â': 'A', 'î': 'i', 'Î': 'I', 'ș': 's', 'Ș': 'S', 'ț': 't', 'Ț': 'T', 'ţ': 't', 'Ţ': 'T'
  };

  return text.replace(/[ăâîșțţĂÂÎȘȚŢ]/g, function(matched) {
    return diacriticsMap[matched];
  });
}

function extractVariable(string, variableName) {
  Logger.log(variableName)
  if (variableName === "Nume si Prenume") {
    var regex = new RegExp(variableName + ':\\s*([^\\n]+?)(?=\\s*Email|$)');
  }
  else if (variableName === "Telefon") {
    var regex = new RegExp(variableName + ':\\s*([^\\n]+?)(?=\\s*Mesaj|$)');
  }
  else if(variableName === "Mesaj"){
    var regex = new RegExp(variableName + ':\\s*([^\\n]+?)(?=\\s*$)');
  }
  else{
    var regex = new RegExp(variableName + ':\\s*([^\\s]+)');
  }
  try{
    var match = string.match(regex);
  } catch(error){
    console.error('An error occurred while extracting the variable:', error);
  }

  
  return match ? match[1] : '';
}

function showDetails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var sheetName = sheet.getName();

  var followUpSheetsNamesArray = ["Periodic","Follow Up - P","Follow Up - CC","Cautare 1","Cautare 2"]

  // Check if the edited cell is in column N or O, dependes on sheet, and row 6 or later
  var detailsColumnNumber;
  if (sheetName == "Call Center"){
    detailsColumnNumber = 14
  }else if(followUpSheetsNamesArray.includes(sheetName)){
    detailsColumnNumber = 15
  }
  if (range.getColumn() === detailsColumnNumber && range.getRow() >= 6) {
    // Retrieve the rich text value
    var richTextValue = range.getRichTextValue();

    // Split the content into lines
    var lines = richTextValue.getText().split('\n');

    // Get the formatted text as HTML with line breaks and bold for first lines
    var formattedText = '';
    var prevLineEmpty = true;  // Start with the first line as bold
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];

      if (line === '') {
        // Empty line, indicating a break
        prevLineEmpty = true;
        formattedText += '<br>';
      } else {
        // Non-empty line
        if (prevLineEmpty) {
          // First non-empty line after a break, make it bold
          formattedText += '<span style="font-weight: bold;">' + line + '</span><br>';
          prevLineEmpty = false;
        } else {
          // Normal line
          formattedText += line + '<br>';
        }
      }
    }

    // Open a custom modal dialog with specified title and content
    var title = 'History Details';
    openModalDialog(title, formattedText);
  }
}

function openModalDialog(title, content) {
  var htmlOutput = HtmlService.createHtmlOutput('<div><p>' + content + '</p></div>')
      .setWidth(400)
      .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

function getHistoryDetailsFromDb(dataBaseArray, code) {
  console.time('getHistoryDetailsFromDb');
  var details = binarySearchAndGetDetails(dataBaseArray,code);
  console.timeEnd('getHistoryDetailsFromDb');
  return details;
}

function binarySearchAndGetDetails(dataBaseArray, codeNumber) {
  console.time('binarySearchAndGetDetails')
  // Assuming dataBaseArray is sorted based on code numbers
  var left = 0;
  var right = dataBaseArray.length - 1;

  while (left <= right) {
    var mid = Math.floor((left + right) / 2);
    var currentCode = dataBaseArray[mid][2];

    if (currentCode === codeNumber) {
      // Code number found, return the conversationDetailsDB
      return dataBaseArray[mid][8];
    } else if (currentCode < codeNumber) {
      left = mid + 1;
    } else {
      right = mid - 1;
    }
  }

  // Code number not found
  console.timeEnd('binarySearchAndGetDetails')
  return null;
}

function sortArray(arrayName, columnNumberforSorting) {
  console.time('sortArray')
  arrayName.sort(function(a, b) {
    return a[columnNumberforSorting] - b[columnNumberforSorting];
  });
  console.timeEnd('sortArray')
}

function copyFormatFromRow5ToRow6Onwards(headerRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the source range (row 5) and target range (from row 6 onwards)
  var sourceRange = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn());
  var targetRange = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn());

  // Copy the format (excluding background color) from the source range to the target range
  sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // Clear the background color for the target range
  targetRange.setBackground(null);
  targetRange.setFontColor("black");
}
