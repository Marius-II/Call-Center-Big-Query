function initiateDataBaseArray() {
  var dataBaseArray = [];
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  // Collect data from sheets
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    if (/^DB[1-9]\d?$|^DB100$/.test(sheetName)) { // Check sheet name pattern
      var data = sheet.getDataRange().getValues();
      
      for (var j = 0; j < data.length; j++) {
        var rowData = data[j];
        if (rowData[0] !== '') { // Check if 'Cod' is not empty
          dataBaseArray.push(rowData);
        }
      }
    }
    Logger.log(sheetName);
  }

  var batchData = [];

  // Transform and collect data for batch write
  for (var i = 0; i < dataBaseArray.length; i++) {
    var code = dataBaseArray[i][0];
    var name = dataBaseArray[i][1];
    var phoneNumber = dataBaseArray[i][2];
    var location = dataBaseArray[i][3];
    var reInteractTime = convertTimestampToDateTime(dataBaseArray[i][4]);
    var callResult = dataBaseArray[i][5];
    var conversationDetails = dataBaseArray[i][6];

    batchData.push({
      code: code,
      fullName: name,
      phoneNumber: phoneNumber,
      source: location,
      lastInteractionDate: reInteractTime,
      lastResult: callResult,
      conversationHistory: conversationDetails
    });
  }

  // Write the batch data to BigQuery
  writeToBigQueryBatch(batchData);
  console.log("Insertion completed succesfuly to BIGQUERY")
}

function writeToBigQueryBatch(batchData) {
  var projectId = 'teraopticcontaproject';
  var datasetId = 'call_center';
  var tableId = 'customers';
  var batchSize = 500; // Define a suitable batch size

  for (var i = 0; i < batchData.length; i += batchSize) {
    var batch = batchData.slice(i, i + batchSize);
    var rows = batch.map(function(data) {
      return { json: data };
    });

    var insertAllRequest = {
      kind: 'bigquery#tableDataInsertAllRequest',
      rows: rows
    };

    try {
      BigQuery.Tabledata.insertAll(insertAllRequest, projectId, datasetId, tableId);
    } catch (error) {
      console.error('Error in batch starting at row ' + i, error.message, error.stack);
    }
  }
}

function writeToBigQuery(data) {
  var projectId = 'teraopticcontaproject';
  var datasetId = 'call_center';
  var tableId = 'customers';

  var tableReference = {
    projectId: projectId,
    datasetId: datasetId,
    tableId: tableId
  };

  var rows = [{
    json: data
  }];

  var insertAllRequest = {
    kind: 'bigquery#tableDataInsertAllRequest',
    rows: rows
  };

  try {
    BigQuery.Tabledata.insertAll(insertAllRequest, projectId, datasetId, tableId);
    console.log("Insertion completed succesfuly to BIGQUERY")
  } catch (error) {
    console.error('Error while inserting into BigQuery:', error);
  }
}

function searchByCodeInBigQuery (code) {
  var startTime = Date.now();
  code = code.toString().trim()
  var projectId = 'teraopticcontaproject'; // Replace with your project ID
  var datasetId = 'call_center'; // Your dataset ID
  var tableId = 'customers'; // Your table ID
  
  code = String(code)

  // Construct the search query
  var query = `SELECT * FROM \`${projectId}.${datasetId}.${tableId}\` WHERE LOWER(code) LIKE '%' || LOWER(@code) || '%'`;
  
  var queryParameters = [{name: 'code', parameterType: {type: 'STRING'}, parameterValue: {value: code}}];

  // Execute the query
  var request = {
    query: query,
    useLegacySql: false,
    parameterMode: 'NAMED',
    queryParameters: queryParameters
  };

  var queryResults = BigQuery.Jobs.query(request, projectId);
  var rowsArray = [];

  // Process the query results
  if (queryResults.rows && queryResults.rows.length > 0) {
    var numFields = queryResults.schema.fields.length; // Get the number of fields

    for (var i = 0; i < queryResults.rows.length; i++) {
      var row = queryResults.rows[i];
      var rowData = [];

      for (var j = 0; j < numFields; j++) {
        rowData.push(row.f[j].v); // Add each field value to rowData
      }
    
      rowsArray.push(rowData);
    }
    var endTime = Date.now();
    var timeTaken = endTime - startTime;
    Logger.log("Time taken: search by code in BigQuery " + timeTaken + " milliseconds");
    return rowsArray

  } else {
    console.log('No row found for code:', code);
    var activeSheetName = SpreadsheetApp.getActiveSheet().getName()
    if (activeSheetName === "Cautare 1" || activeSheetName === "Cautare 2"){
      showUserMessage(`Niciun rezultat cu codul ${code} nu a fost gasit`)
    }

  }
  var endTime = Date.now();
  var timeTaken = endTime - startTime;
  Logger.log("Time taken: search by code in BigQuery " + timeTaken + " milliseconds");
}

function searchByNameInBigQuery (name) {
  name = name.trim()
  var projectId = 'teraopticcontaproject'; // Replace with your project ID
  var datasetId = 'call_center'; // Your dataset ID
  var tableId = 'customers'; // Your table ID

  // Construct the search query
  
  // Adjust the query for case-insensitive, partial match
  var query = `SELECT * FROM \`${projectId}.${datasetId}.${tableId}\` WHERE LOWER(fullName) LIKE '%' || LOWER(@name) || '%'`;

  var queryParameters = [{name: 'name', parameterType: {type: 'STRING'}, parameterValue: {value: name}}];

  // Execute the query
  var request = {
    query: query,
    useLegacySql: false,
    parameterMode: 'NAMED',
    queryParameters: queryParameters
  };

  var queryResults = BigQuery.Jobs.query(request, projectId);
  var rowsArray = [];

  // Process the query results
  if (queryResults.rows && queryResults.rows.length > 0) {
    var numFields = queryResults.schema.fields.length; // Get the number of fields

    for (var i = 0; i < queryResults.rows.length; i++) {
      var row = queryResults.rows[i];
      var rowData = [];

      for (var j = 0; j < numFields; j++) {
        rowData.push(row.f[j].v); // Add each field value to rowData
      }
    
      rowsArray.push(rowData);
    }
    return rowsArray

  } else {
    console.log('No row found for code:', name);
    showUserMessage(`Niciun rezultat cu codul ${name} nu a fost gasit`)
  }
}

function updateRowInBigQuery(code, fullName, phoneNumber, source, lastInteractionDate, lastResult, conversationHistory) {
  var startTime = Date.now();
  var projectId = 'teraopticcontaproject'; // Replace with your project ID
  var datasetId = 'call_center'; // Your dataset ID
  var tableId = 'customers'; // Your table ID
  
  // Construct the update query
  var query = 'UPDATE `' + projectId + '.' + datasetId + '.' + tableId + '` ' +
              'SET fullName = @fullName, phoneNumber = @phoneNumber, source = @source, ' +
              'lastInteractionDate = @lastInteractionDate, lastResult = @lastResult, ' +
              'conversationHistory = @conversationHistory ' +
              'WHERE code = @code';

  // Set the parameters for the query
  var queryParameters = [
    {name: 'code', parameterType: {type: 'STRING'}, parameterValue: {value: code}},
    {name: 'fullName', parameterType: {type: 'STRING'}, parameterValue: {value: fullName}},
    {name: 'phoneNumber', parameterType: {type: 'STRING'}, parameterValue: {value: phoneNumber}},
    {name: 'source', parameterType: {type: 'STRING'}, parameterValue: {value: source}},
    {name: 'lastInteractionDate', parameterType: {type: 'DATETIME'}, parameterValue: {value: lastInteractionDate}},
    {name: 'lastResult', parameterType: {type: 'STRING'}, parameterValue: {value: lastResult}},
    {name: 'conversationHistory', parameterType: {type: 'STRING'}, parameterValue: {value: conversationHistory}}
  ];

  // Execute the query
  var request = {
    query: query,
    useLegacySql: false,
    parameterMode: 'NAMED',
    queryParameters: queryParameters
  };

  // Using BigQuery instead of BigQueryApp
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var endTime = Date.now();
  var timeTaken = endTime - startTime;
  Logger.log('Row updated successfully in BigQuery table');
  Logger.log("Time taken updateRow in BigQuery: " + timeTaken + " milliseconds");
}

function insertRowToBigQuery(data) {
  var projectId = 'teraopticcontaproject';
  var datasetId = 'call_center';
  var tableId = 'customers';

  // Construct SQL INSERT query
  var sqlQuery = `INSERT INTO \`${projectId}.${datasetId}.${tableId}\` ` +
                 `(code, fullName, phoneNumber, source, lastInteractionDate, lastResult, conversationHistory) ` +
                 `VALUES (@code, @fullName, @phoneNumber, @source, @lastInteractionDate, @lastResult, @conversationHistory)`;

  Logger.log(sqlQuery)

  // Set the parameters for the query
  var queryParameters = Object.keys(data).map(function(key) {
    return {
      name: key,
      parameterType: { type: getTypeOf(data[key]) }, // Ensure getTypeOf function is defined and returns the correct BigQuery data type
      parameterValue: { value: data[key] }
    };
  });

  Logger.log(queryParameters)

  // Configuration for the query
  var jobConfigurationQuery = {
    query: sqlQuery,
    useLegacySql: false,
    parameterMode: 'NAMED',
    queryParameters: queryParameters
  };

  var jobData = {
    configuration: {
      query: jobConfigurationQuery
    }
  };

  // Execute the query
  try {
    BigQuery.Jobs.insert(jobData, projectId);
    console.log("Insertion completed successfully to BIGQUERY");
  } catch (error) {
    console.error('Error while inserting into BigQuery:', error);
  }
}

function fetchFromBigQuery(lastResult) {
  
  var projectId = 'teraopticcontaproject';
  var datasetId = 'call_center';
  var tableId = 'customers';

  // Define the SQL query
  var sqlQuery = 'SELECT * FROM `' + projectId + '.' + datasetId + '.' + tableId + '` ' +
                 'WHERE lastResult = "' + lastResult + '" '

  // Set up the BigQuery request
  var request = {
    query: sqlQuery,
    useLegacySql: false
  };

  // Run the query
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // Get rows of results
  var rows = queryResults.rows;
  if (!rows) {
    Logger.log('No rows returned.');
    return;
  }

  // Create an array to store the data
  var data = [];

  // Iterate through query results and push rows to array
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rowData = [];
    for (var j = 0; j < row.f.length; j++) {
      rowData.push(row.f[j].v);
    }
    data.push(rowData);
  }
  
  Logger.log(data);
  return data;
}

function fetchOldResultsBigQuery (howManyDaysAgoMax, howManyDaysAgoMin) {
  
  var today = new Date();
  var startDate = new Date();
  startDate.setDate(today.getDate() - howManyDaysAgoMax);
  var endDate = new Date();
  endDate.setDate(today.getDate() - howManyDaysAgoMin);

  // Arrow function for formatting dates
  const formatDateForBigQuery = (date) => date.toISOString().split('T')[0];

  var startDateString = formatDateForBigQuery(startDate);
  var endDateString = formatDateForBigQuery(endDate);

  // Replace with your project ID, dataset ID, and table ID
  var projectId = 'teraopticcontaproject';
  var datasetId = 'call_center';
  var tableId = 'customers';

  var sqlQuery = `SELECT *
  FROM \`${projectId}.${datasetId}.${tableId}\`
  WHERE lastInteractionDate BETWEEN '${startDateString}' AND '${endDateString}'`;

  // Set up the BigQuery request
  var request = {
    query: sqlQuery,
    useLegacySql: false
  };

  // Run the query
  var queryResults = BigQuery.Jobs.query(request, projectId);
  var jobId = queryResults.jobReference.jobId;

  // Check on status of the Query Job
  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  // Get rows of results
  var rows = queryResults.rows;
  if (!rows) {
    Logger.log('No rows returned.');
    return;
  }

  // Create an array to store the data
  var data = [];

  // Iterate through query results and push rows to array
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var rowData = [];
    for (var j = 0; j < row.f.length; j++) {
      rowData.push(row.f[j].v);
    }
    data.push(rowData);
  }
  return data;
}

function deleteFromBigQuery(codeValue) {
  var projectId = 'teraopticcontaproject';
  var datasetId = 'call_center';
  var tableId = 'customers';

  // Define the SQL query for deletion
  var sqlQuery = 'DELETE FROM `' + projectId + '.' + datasetId + '.' + tableId + '` ' +
                 'WHERE code = "' + codeValue + '" ';

  // Set up the BigQuery request
  var request = {
    query: sqlQuery,
    useLegacySql: false
  };

  // Run the query
  BigQuery.Jobs.query(request, projectId);

  //DELETE doesn't return rows, no need to process the results
  Logger.log('Deletion query executed for code: ' + codeValue);
}


// Helper function to get the BigQuery data type of a JavaScript value
function getTypeOf(value) {
  if (value === null) {
    return 'STRING'; // Adjust as needed
  } else if (typeof value === 'boolean') {
    return 'BOOL';
  } else if (typeof value === 'number') {
    return 'STRING'; // or 'INTEGER', depending on the context
  } else if (value instanceof Date) {
    return 'DATETIME'; // or 'DATETIME', based on your table's schema
  } else {
    return 'STRING';
  }
}
