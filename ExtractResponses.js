function EXTRACT_RESPONSES(input) {
  // Define the regular expression to match "Raspuns: " followed by any text until an empty line or the end of the text
  var regex = /Raspuns:\s*([^\n]+(?:\n(?!\n)[^\n]+)*)/g;
  var matches;
  var output = [];

  // Loop through all matches in the input text
  while ((matches = regex.exec(input)) !== null) {
    // Add the matched text (excluding "Raspuns: ") to the output array, replacing new lines with spaces
    output.push(matches[1].trim().replace(/\n/g, ' '));
  }

  // Join the extracted values with a comma and return the result
  return output.join(", ");
}


function writeToColumnQ() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Gets the data range of the active sheet
  var values = range.getValues(); // Gets the values of the cells in the data range
  
  for (var i = 5; i < values.length; i++) {
    var cellValue = values[i][14]; // Assuming column O is the 15th column, index starting from 0
    if (cellValue) { // Check if the cell is not empty
      var response = EXTRACT_RESPONSES(cellValue);
      sheet.getRange(i + 1, 17).setValue(response); // Write the result to the same row in column Q, assuming column Q is the 17th column
    }
  }
}
