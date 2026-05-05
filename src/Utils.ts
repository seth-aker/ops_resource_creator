
function getSpreadSheetData<T>(spreadsheet: string) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet);
  if(!sheet) throw new Error(`Could not find spreadsheet: "${spreadsheet}"`)
  const dataRange = sheet.getDataRange(); // Get data
  const data = dataRange.getValues(); // create 2D array
  
  // Process data (e.g., converting to JSON format for API)
  const headers = data[0];
  if(!headers) throw new Error(`No headers found in row`)
  const jsonData = [];

  for(let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row: Record<string, any> = {}
    for(let colIndex = 0; colIndex < headers.length; colIndex++) {
      const rowData = data[rowIndex];
      if(!rowData) continue;
      let value = rowData[colIndex];
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

function createHeaders(token: string) {
  return {
    "Authorization": `Bearer ${token}`,
    'Content-Type': 'application/json'
  }
}