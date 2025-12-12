const ESTIMATE_REF = "00000000-0000-0000-0000-000000000000";
const DEFAULT_BATCH_SIZE = 100;

type TSpreadsheetValues = Number | Boolean | Date | String

function getSpreadSheetData<T>(spreadsheetName: string): T[] {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheetName);
  if(!sheet) throw new Error(`Could not find spreadsheet: "${spreadsheetName}"`)
  const dataRange = sheet.getDataRange(); // Get data
  const data = dataRange.getValues(); // create 2D array
  
  // Process data (e.g., converting to JSON format for API)
  const headers = data[0]; 
  const jsonData = [];

  for(let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row: Record<string, TSpreadsheetValues> = {}
    for(let colIndex = 0; colIndex < headers.length; colIndex++) {
      let value = data[rowIndex][colIndex] as TSpreadsheetValues;
      // Trim whitespace if the value is a string
      if(typeof value === 'string') {
        value = value.trim()
      }
      row[headers[colIndex]] = value;
    }
    jsonData.push(row);
  }
  return jsonData as T[];
}

function createHeaders(token: string, additionalHeaders?: Record<string, string>) {
    const baseUrl = PropertiesService.getUserProperties().getProperty('baseUrl')
    const userName = PropertiesService.getUserProperties().getProperty('userName')
    const serverName = PropertiesService.getUserProperties().getProperty('serverName')
    const clientID = PropertiesService.getUserProperties().getProperty('clientID')
    const clientSecret = PropertiesService.getUserProperties().getProperty('clientSecret')
    const dbName = PropertiesService.getUserProperties().getProperty('dbName')
    if(!baseUrl || !userName || !serverName || !dbName || !clientID || !clientSecret) {
      throw new Error('Missing required user properties')
    }
    const connectionString = `Server=${serverName};Database=${dbName};MultipleActiveResultSets=true;Integrated Security=SSPI;`
    
    return {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'ConnectionString': connectionString,
        'ClientID': clientID,
        'ClientSecret': clientSecret,
        ...additionalHeaders
    }
}
function batchFetch(batchOptions: (string | GoogleAppsScript.URL_Fetch.URLFetchRequest)[], retryCount: number = 0) {
  Utilities.sleep(retryCount * retryCount * 1000); // Exponential Backoff

  const sliceCount = Math.ceil(batchOptions.length / DEFAULT_BATCH_SIZE)
  const responses: GoogleAppsScript.URL_Fetch.HTTPResponse[] = []
  
  for(let i = 0; i < sliceCount; i++) {
    if(retryCount === 0) {
      SpreadsheetApp.getUi().alert(`Posting batch ${i + 1} of ${sliceCount}`)
    }
    responses.push(...UrlFetchApp.fetchAll(batchOptions.slice(i * DEFAULT_BATCH_SIZE, (i + 1) * DEFAULT_BATCH_SIZE))) // passing a value greater than the length of the array will include all values to the end of the array.
    // if only one call is being made or on the last call, don't sleep
    if(sliceCount > 1 && i < sliceCount - 1) {
      Utilities.sleep(1000)
    }
  }
  const retries: (string | GoogleAppsScript.URL_Fetch.URLFetchRequest)[] = [];
  const responseIndices: number[] = []; 
  responses.forEach((response, index) => {
    const responseCode = response.getResponseCode()
    const responseMessage = response.getContentText();
    if(responseCode === 500 && responseMessage.includes("Connection Timeout Expired.")) {
      retries.push(batchOptions[index])
      responseIndices.push(index);
    }
  })
  if(retryCount < 5 && retries.length > 0) {
    Logger.log(`${retries.length} entries failed due to connection timeout, retrying...`)
    SpreadsheetApp.getUi().alert(`${retries.length} entries failed due to connection timeout, retrying...`)
    const retryResponses = batchFetch(retries, retryCount + 1);
    retryResponses.forEach((response, index) => {
      responses[responseIndices[index]] = response;
    })
  }
  return responses
}

function highlightRows(rowIndices: number[], color: string) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const lastColumn = sheet.getLastColumn()
    rowIndices.forEach((row) => {
      sheet.getRange(row, 1,1, lastColumn).setBackground(color)
    })
}