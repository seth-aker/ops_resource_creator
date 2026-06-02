function writeLogToSpreadsheet(message: string | string[]) {
  Logger.log(message);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let logsSheet = spreadsheet.getSheetByName('Logs')
  if(!logsSheet) {
    logsSheet = spreadsheet.insertSheet('Logs');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    logsSheet.appendRow(["Script", "Timestamp", "Message"]).getRange(1,1,1,3).setTextStyle(bold)
  }
  const scriptName = PropertiesService.getUserProperties().getProperty('currentScript') ?? "Unknown"
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  if(typeof message === 'string') {
    logsSheet.appendRow([scriptName, timestamp, message])
  } else {
    const startRow = logsSheet.getLastRow() + 1
    const rows = message.map(each => [scriptName, timestamp, each])
    logsSheet.getRange(startRow, 1, rows.length, 3).setValues(rows)
  }
}

/**
 * Logs one or many events to 3 places
 * 
 * 1. Calls Logger.log(message)
 * 2. Adds the message to the scriptEvents user cache
 * 3. Passes message to the function printLog(), which appends data in the Logs spreadsheet
 */ 
function logEvent(message: string | string[]) {
  const service = CacheService.getUserCache();
  const raw = service.get('scriptEvents');
  const events: string[] = raw ? JSON.parse(raw) : [];
  if(typeof message === 'string') {
    writeLogToSpreadsheet(message);
    events.push(message);
  } else if (typeof message === 'object' && message.length > 0) {
    writeLogToSpreadsheet(message)
    events.push(...message)
  }


  let jsonString = JSON.stringify(events);
  let byteLength = Utilities.newBlob(jsonString).getBytes().length;

  while (byteLength > MAX_CACHE_SIZE && events.length > 0) {
    events.shift();
    jsonString = JSON.stringify(events);
    byteLength = Utilities.newBlob(jsonString).getBytes().length;
  }
  service.put('scriptEvents', jsonString)
}