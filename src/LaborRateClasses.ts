interface IRawLaborRateClass {
  "Business Unit": string,
  "Rate Class Name": string,
  "Rate Class ID": string,
  "Integration Key"?: string | null,
  "Notes"?: string | null,
  "Is Inactive?": boolean,
  "ObjectID"?: string
}
interface LaborRateClassDTO {
  BusinessUnitUniqueName: string,
  Name: string,
  RateClassID: string,
  IntegrationKey?: string | null,
  Notes?: string | null,
  IsInactive: boolean ,
  ObjectID?: string
}
const LABOR_RATE_CLASS_KEYS: Array<keyof IRawLaborRateClass> = [
  "Business Unit",
  "Rate Class Name",
  "Rate Class ID",
  "Integration Key",
  "Notes",
  "Is Inactive?",
  "ObjectID"
];
const LABOR_RATE_CLASS_FILTER_OPTIONS: IFilterByOptions[] = [
  { value: 'BusinessUnitUniqueName', type: 'string' },
  { value: 'Name', type: 'string' },
  { value: 'RateClassID', type: 'string' },
  { value: 'IntegrationKey', type: 'string' },
  { value: 'Notes', type: 'string' },
  { value: 'IsInactive', type: 'boolean' },
  { value: 'ObjectID', type: 'string' }
];

function DisplayLaborRateClassFilterBuilder() {
  const template = HtmlService.createTemplateFromFile('BuildFilter') as IBuildFilterTemplate
  template.filterByOptions = LABOR_RATE_CLASS_FILTER_OPTIONS;
  template.serverFunctionName = "GetLaborRateClasses"
  const html = template.evaluate()
    .setHeight(900)
    .setWidth(1100)
  const ui = SpreadsheetApp.getUi()
  ui.showModalDialog(html, "Labor Rate Classes Filter")
}

function GetLaborRateClasses(options: GetOptions) {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("GetUsers")
  openProgressSidebar("Getting Users");

  logEvent("Starting get users script")
  const token = getOpsToken();
  const baseUrl = getBaseURL();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let laborRateClassSheet = spreadsheet.getSheetByName('Labor Rate Classes')
  if(!laborRateClassSheet) {
    laborRateClassSheet = spreadsheet.insertSheet('Labor Rate Classes');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    laborRateClassSheet.appendRow(LABOR_RATE_CLASS_KEYS).getRange(1,1,1, LABOR_RATE_CLASS_KEYS.length).setTextStyle(bold)
  }
  const currentSpreadSheetData = getSpreadSheetData<IRawLaborRateClass>("Labor Rate Classes")
  const ui = SpreadsheetApp.getUi();
  if(currentSpreadSheetData.length > 0) {
    const response = ui.alert("WARNING: any existing data will be overwritten. Do you want to continue?",
      ui.ButtonSet.YES_NO
    )
    if(response === ui.Button.NO) {
      logEvent("Get Labor Rate Class Script Canceled")
      setIsScriptFinished(true);
      return;
    } else {
      laborRateClassSheet.getRange(2, 1, laborRateClassSheet.getLastRow(), laborRateClassSheet.getLastColumn()).clearContent()
    }
  }
  const headers = createHeaders(token)

  const rateClasses = getDatabaseItems<LaborRateClassDTO>(`${baseUrl}/LaborRateClass${options.filterQuery}`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })
  logEvent(`${rateClasses.length} rate classes recieved.`)
  const headerValues = laborRateClassSheet.getDataRange().getValues()[0] as typeof LABOR_RATE_CLASS_KEYS
  LABOR_RATE_CLASS_KEYS.forEach((key) => {
    if(!headerValues.includes(key)) {
      headerValues.push(key)
      laborRateClassSheet.getRange(1, headerValues.length, 1,1).setValue(key)
    }
  })

  // arranges each row to match the order of the headers in the spreadsheet.
  const rowValues = rateClasses.map(e => {
    const values = mapLaborRateClassToRaw(e)
    return headerValues.map(key => values[key] ?? "")
  })
  const startRow = 2;

  laborRateClassSheet.getRange(startRow, 1, rowValues.length, headerValues.length).setValues(rowValues)
  
  logEvent("Script Complete!")
  ui.alert("Script Complete!")
  setIsScriptFinished(true)
}
function CreateLaborRateClasses() {
  try { 
    setIsScriptFinished(false);
    clearScriptProgress()
    setCurrentScript("CreateLaborRateClasses")
    openProgressSidebar("Creating Labor Rate Classes");

    logEvent("Starting Create Labor Rate Classes Script")
    const token = getOpsToken()
    const baseUrl = getBaseURL()
    
    const rateClassData = getSpreadSheetData<IRawLaborRateClass>('Labor Rate Classes')
    if(!rateClassData || rateClassData.length === 0) {
      SpreadsheetApp.getUi().alert("No data found to send!")
      setIsScriptFinished(false);
      return;
    }

    const url = baseUrl + "/LaborRateClass"
    const headers = createHeaders(token);
    const userDTOs = rateClassData.map((row) => {
        return mapRawToRateClassDTO(row);
      })
    const batchOptions = userDTOs.map(row => {
      const options = {
        url,
        method: 'post' as const,
        headers,
        payload: JSON.stringify(row),
        muteHttpExceptions: true
      }
      return options;
    })
    logEvent(`Uploading ${batchOptions.length} rate classes...`)
    const results = batchFetch(batchOptions);
    
    const failedRows: number[] = []
    const failureMessages: string[] = [];
    results.forEach((result, idx) => {
      const code = result.getResponseCode();
      if(code >= 400) {
        failedRows.push(idx);
        failureMessages.push(`Row ${idx + 2}: [${code} Error]: ${result.getContentText()}`)
      }
    })
    
    if(failedRows.length > 0) {
      logEvent([`Some rows failed!:`, ...failureMessages])
      highlightRows(failedRows.map(each => each + 2), 'red');
    } else {
      logEvent("All rate classes created successfully!")
    }
    logEvent("Script Complete!")
    SpreadsheetApp.getUi().alert("Script Complete!")
    setIsScriptFinished(true);
  } catch (err) {
    setIsScriptFinished(true)
    throw err
  }
}

function UpdateLaborRateClasses() {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("UpdateLaborRateClasses")
  openProgressSidebar("Updating Labor Rate Classes");

  logEvent("Starting Update Labor Rate Classes Script")

  const token = getOpsToken()
  const baseUrl = getBaseURL()
  
  const rateClassData = getSpreadSheetData<IRawLaborRateClass>('Labor Rate Classes')
  if(!rateClassData || rateClassData.length === 0) {
    SpreadsheetApp.getUi().alert("No data found to send!")
    setIsScriptFinished(false);
    return;
  }

  const url = baseUrl + "/LaborRateClass"
  const headers = createHeaders(token);
  const rateClassDtos = rateClassData.map(mapRawToRateClassDTO)
  let payloads: LaborRateClassDTO[] = []
  if(!rateClassDtos.every(entry => entry.ObjectID && entry.ObjectID.length > 0)) {
    const getOptions = {
      headers,
      method: 'get' as const,
      muteHttpExceptions: true
    }
    const existingRateClasses = getDatabaseItems<LaborRateClassDTO>(url, getOptions)
    payloads = rateClassDtos.map((each, idx) => {
      const rcId = each.RateClassID
      const intKey = each.IntegrationKey
      const existing = existingRateClasses.find(rc => {
        const idMatches = rcId === rc.RateClassID
        if(rc.IntegrationKey) {
          return idMatches && rc.IntegrationKey === intKey
        } else {
          return idMatches;
        }
      })
      if(!existing) {
        const errorMessage = `Error: Could not find existing rate class with Rate Class ID: ${rcId}${intKey ? ` and Integration Key: ${intKey}`: "."}`
        logEvent(errorMessage)
        highlightRows([idx + 2], 'red')
        throw new Error(errorMessage)
      }
      return {
        ...existing,
        ...each
      }
    })
  } else { // all rows have objectIds and can be updated directly
    payloads = rateClassDtos
  }
  const batchOptions = payloads.map(row => {
    const options = {
      url,
      method: 'put' as const,
      headers,
      payload: JSON.stringify(row),
      muteHttpExceptions: true
    }
    return options;
  })
  logEvent("Updating rate classes...")
  const results = batchFetch(batchOptions);
  
  const failedRows: number[] = []
  const failureMessages: string[] = [];
  results.forEach((result, idx) => {
    const code = result.getResponseCode();
    if(code >= 400) {
      failedRows.push(idx);
      failureMessages.push(`Row ${idx + 2}: [${code} Error]: ${result.getContentText()}`)
    }
  })
  
  if(failedRows.length > 0) {
    logEvent([`Some rows failed!:`, ...failureMessages])
    highlightRows(failedRows.map(f => f + 2), 'red')
  } else {
    logEvent("All rate classes updated successfully!")
  }
  logEvent("Script Complete")
  SpreadsheetApp.getUi().alert("Script Complete!")
  setIsScriptFinished(true)
}

function mapRawToRateClassDTO(row: IRawLaborRateClass): LaborRateClassDTO {
  return {
    BusinessUnitUniqueName: row["Business Unit"],
    RateClassID: row["Rate Class ID"],
    Name: row["Rate Class Name"],
    Notes: row.Notes,
    IntegrationKey: row["Integration Key"],
    IsInactive: row["Is Inactive?"],
    ObjectID: row.ObjectID,
  }
}
function mapLaborRateClassToRaw(dto: LaborRateClassDTO): IRawLaborRateClass {
  return {
    "Business Unit": dto.BusinessUnitUniqueName,
    "Rate Class ID": dto.RateClassID,
    "Rate Class Name": dto.Name,
    "Notes": dto.Notes,
    "Integration Key": dto.IntegrationKey,
    "Is Inactive?": dto.IsInactive,
    "ObjectID": dto.ObjectID,
  }
}