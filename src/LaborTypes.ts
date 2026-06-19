interface IRawLaborType {
  "Business Unit": string,
  "Labor Type ID": string,
  "Labor Type Name": string,
  "Is Inactive?": boolean,
  "Notes"?: string | null
  "Labor Type Integration Key"?: string | null
  "Labor Type ObjectID"?: string,
  // Labor Type Rate
  "Labor Rate Class ID"?: string,
  "Labor Rate Class Integration Key"?: string | null,
  "Alternate Labor Type ID"?: string | null,
  "Regular Cost"?: number,
  "Overtime Cost"?: number,
  "Double Time Cost"?: number,
  "Labor Type Rate Object ID"?: string
}
const LABOR_TYPE_KEYS: Array<keyof IRawLaborType> = [
  "Business Unit",
  "Labor Type Name",
  "Labor Type ID",
  "Labor Type Integration Key",
  "Notes",
  "Is Inactive?",
  "Labor Rate Class ID",
  "Labor Rate Class Integration Key",
  "Alternate Labor Type ID",
  "Regular Cost",
  "Overtime Cost",
  "Double Time Cost",
  "Labor Type ObjectID",
  "Labor Type Rate Object ID"
];
interface LaborTypeDTO {
  BusinessUnitUniqueName: string,
  LaborTypeID: string,
  Name: string,
  IsInactive: boolean,
  Notes?: string | null,
  IntegrationKey?: string | null,
  ObjectID?: string
}
interface LaborTypeRateDTO {
  LaborTypeID: string,
  LaborTypeIntegrationKey?: string | null,
  LaborRateClassID: string,
  LaborRateClassIntegrationKey?: string | null
  AlternateLaborTypeID?: string | null
  UnitRegularCost: number,
  UnitOvertimeCost: number,
  UnitDoubleTimeCost: number,
  ObjectID?: string
}
const LABOR_TYPE_FILTER_OPTIONS: IFilterByOptions[] = [
  { value: 'BusinessUnitUniqueName', type: 'string' },
  { value: 'LaborTypeID', type: 'string' },
  { value: 'Name', type: 'string' },
  { value: 'IsInactive', type: 'boolean' },
  { value: 'Notes', type: 'string' },
  { value: 'IntegrationKey', type: 'string' },
  { value: 'ObjectID', type: 'string' }
];

const LABOR_TYPE_RATE_FILTER_OPTIONS: IFilterByOptions[] = [
  { value: 'LaborTypeID', type: 'string' },
  { value: 'LaborTypeIntegrationKey', type: 'string' },
  { value: 'LaborRateClassID', type: 'string' },
  { value: 'LaborRateClassIntegrationKey', type: 'string' },
  { value: 'AlternateLaborTypeID', type: 'string' },
  { value: 'UnitRegularCost', type: 'number' },
  { value: 'UnitOvertimeCost', type: 'number' },
  { value: 'UnitDoubleTimeCost', type: 'number' },
  { value: 'ObjectID', type: 'string' }
];

function DisplayLaborTypeFilterBuilder() {
  const template = HtmlService.createTemplateFromFile('BuildFilter') as IBuildFilterTemplate
  template.filterByOptions = LABOR_TYPE_FILTER_OPTIONS;
  template.serverFunctionName = "GetLaborTypes"
  const html = template.evaluate()
    .setHeight(900)
    .setWidth(1100)
  const ui = SpreadsheetApp.getUi()
  ui.showModalDialog(html, "Labor Types Filter")
}
function GetLaborTypes(options: GetOptions) {
  setIsScriptFinished(false)
  clearScriptProgress()
  setCurrentScript("GetLaborTypes")
  openProgressSidebar("Getting Labor Types")

  logEvent("Starting Get Labor Types Script")
  const token = getOpsToken()
  const baseUrl = getBaseURL()

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let laborTypeSheet = spreadsheet.getSheetByName('Labor Types')
  if(!laborTypeSheet) {
    laborTypeSheet = spreadsheet.insertSheet('Labor Types');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    laborTypeSheet.appendRow(LABOR_TYPE_KEYS).getRange(1,1,1, LABOR_TYPE_KEYS.length).setTextStyle(bold)
  }
  const currentSpreadSheetData = getSpreadSheetData<IRawLaborType>("Labor Types")
  const ui = SpreadsheetApp.getUi();
  if(currentSpreadSheetData.length > 0) {
    const response = ui.alert("WARNING: any existing data will be overwritten. Do you want to continue?",
      ui.ButtonSet.YES_NO
    )
    if(response === ui.Button.NO) {
      logEvent("Get Labor Types Script Canceled")
      setIsScriptFinished(true);
      return;
    } else {
      laborTypeSheet.getRange(2, 1, laborTypeSheet.getLastRow(), laborTypeSheet.getLastColumn()).clearContent()
    }
  }
  const headers = createHeaders(token)
  logEvent("Retrieving Labor Types...")
  const laborTypes = getDatabaseItems<LaborTypeDTO>(`${baseUrl}/LaborType${options.filterQuery}`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })
  logEvent(`${laborTypes.length} labor types recieved`)
  logEvent(`Retrieving labor type rates...`)
  const laborTypeRates = getDatabaseItems<LaborTypeRateDTO>(`${baseUrl}/LaborTypeRate`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })
  logEvent(`${laborTypeRates.length} labor type rates recieved`)
  const rowData: IRawLaborType[] = []
  laborTypeRates.forEach(rateDto => {
    const laborTypeDto = laborTypes.find(t => {
      if(t.IntegrationKey) {
        return t.LaborTypeID === rateDto.LaborTypeID && t.IntegrationKey === rateDto.LaborTypeIntegrationKey
      } else {
        return t.LaborTypeID === rateDto.LaborTypeID
      }
    });
    if(laborTypeDto) {
      rowData.push(mapDtosToLaborTypeRow(laborTypeDto, rateDto))
    }
  })

  const headerValues = laborTypeSheet.getDataRange().getValues()[0] as typeof LABOR_TYPE_KEYS
  LABOR_TYPE_KEYS.forEach((key) => {
    if(!headerValues.includes(key)) {
      headerValues.push(key)
      laborTypeSheet.getRange(1, headerValues.length, 1,1).setValue(key)
    }
  })

  const orderedRowValues = rowData.map(r => {
    return headerValues.map(key => r[key] ?? '')
  })
  const startRow = 2;
  laborTypeSheet.getRange(startRow, 1, orderedRowValues.length, headerValues.length).setValues(orderedRowValues)
  logEvent("Script Complete!")
  ui.alert("Script Complete!")
  setIsScriptFinished(true)
}

function CreateLaborTypes() {
  setIsScriptFinished(false)
  clearScriptProgress()
  setCurrentScript("CreateLaborTypes")
  openProgressSidebar("Creating Labor Types")

  logEvent("Starting Create Labor Types Script")
  const token = getOpsToken()
  const baseUrl = getBaseURL()
  
  const laborTypesData = getSpreadSheetData<IRawLaborType>('Labor Types')
  if(!laborTypesData || laborTypesData.length === 0) {
    SpreadsheetApp.getUi().alert("No data found to send!")
    setIsScriptFinished(false);
    return;
  }
  const laborTypeURL = baseUrl + "/LaborType"
  const headers = createHeaders(token);
  
  const laborTypeDtos = laborTypesData.map(mapLaborTypeRowToDTO)

  const uniqueLaborTypesMap = new Map<string, LaborTypeDTO>();
  laborTypeDtos.forEach(dto => {
    const intKey = dto.IntegrationKey
    const id = `${dto.LaborTypeID}${intKey ? intKey : ""}`
    if(!uniqueLaborTypesMap.has(id)) {
      uniqueLaborTypesMap.set(id, dto)
    }
  })
  const uniqueLaborTypes = Array.from(uniqueLaborTypesMap.values())
  const batchOptions = uniqueLaborTypes.map(row => {
    const options = {
      url: laborTypeURL,
      method: 'post' as const,
      headers,
      payload: JSON.stringify(row),
      muteHttpExceptions: true
    }
    return options
  })

  logEvent(`Uploading ${batchOptions.length} labor types...`)
  const laborTypeResults = batchFetch(batchOptions);
  const failedIdxs: number[] = [];
  const failureMessages: string[] = [];
  laborTypeResults.forEach((result, idx) => {
    const code = result.getResponseCode();
    if(code >= 400) {
      failedIdxs.push(idx);
      failureMessages.push(`[${code} Error]: ${result.getContentText()}`)
    }
  })

  if(failedIdxs.length > 0) {
    logEvent([`Some rows failed!: `, ...failureMessages])
    logEvent("Canceling script")
    setIsScriptFinished(true);
    SpreadsheetApp.getUi().alert("Some rows failed, script canceled.")
    return;
  } else {
    logEvent(`${laborTypeResults.length} labor types successfully created.`)
  }
  
  const rowsWithCost = laborTypesData.filter(row => row["Labor Type ID"] && row["Labor Rate Class ID"] && row["Regular Cost"] && row["Overtime Cost"] && row["Double Time Cost"])
  if(rowsWithCost.length > 0) {
    const laborTypeRateURL = baseUrl + "/LaborTypeRate"
    const laborTypeRateDtos = rowsWithCost.map(mapLaborTypeRowToRateDTO)
    const batchOptions = laborTypeRateDtos.map(dto => {
      return {
        url: laborTypeRateURL,
        method: 'post' as const,
        headers,
        payload: JSON.stringify(dto),
        muteHttpExceptions: true
      }
    })

    const responses = batchFetch(batchOptions);
    const failedIdxs: number[] = [];
    const failureMessages: string[] = [];
    responses.forEach((result, idx) => {
      const code = result.getResponseCode();
      if(code >= 400) {
        failedIdxs.push(idx);
        failureMessages.push(`[${code} Error]: ${result.getContentText()}`)
      }
    })

    if(failedIdxs.length > 0) {
      logEvent([`Some labor rates failed!:`, ...failureMessages])

    } else {
      logEvent("All labor rates successfully created.")
    }
  } else {
    logEvent("No rows have valid cost data, skipping cost upload!")
  }

  logEvent("Script Complete")
  SpreadsheetApp.getUi().alert("Script Complete!")
  setIsScriptFinished(true)    
}

function mapLaborTypeRowToDTO(row: IRawLaborType): LaborTypeDTO {
  return {
    BusinessUnitUniqueName: row["Business Unit"],
    LaborTypeID: row["Labor Type ID"],
    Name: row["Labor Type Name"],
    IsInactive: row["Is Inactive?"],
    Notes: row.Notes,
    IntegrationKey: row["Labor Type Integration Key"],
    ObjectID: row["Labor Type ObjectID"]
  }
}

function mapLaborTypeRowToRateDTO(row: IRawLaborType): LaborTypeRateDTO {
  return {
    LaborTypeID: row["Labor Type ID"],
    LaborTypeIntegrationKey: row["Labor Type Integration Key"],
    AlternateLaborTypeID: row["Alternate Labor Type ID"],
    LaborRateClassID: row["Labor Rate Class ID"] ?? '',
    LaborRateClassIntegrationKey: row["Labor Rate Class Integration Key"],
    UnitRegularCost: row["Regular Cost"] ?? 0,
    UnitOvertimeCost: row["Overtime Cost"] ?? 0,
    UnitDoubleTimeCost: row["Double Time Cost"] ?? 0
  }
}
function mapDtosToLaborTypeRow(laborType: LaborTypeDTO, laborTypeRate: LaborTypeRateDTO): IRawLaborType {
  return {
    "Business Unit": laborType.BusinessUnitUniqueName,
    "Labor Type ID": laborType.LaborTypeID,
    "Labor Type Name": laborType.Name,
    "Is Inactive?": laborType.IsInactive,
    "Notes": laborType.Notes,
    "Labor Type Integration Key": laborType.IntegrationKey,
    "Labor Type ObjectID": laborType.ObjectID,
    "Labor Rate Class ID": laborTypeRate.LaborRateClassID,
    "Labor Rate Class Integration Key": laborTypeRate.LaborRateClassIntegrationKey,
    "Alternate Labor Type ID": laborTypeRate.AlternateLaborTypeID,
    "Regular Cost": laborTypeRate.UnitRegularCost,
    "Overtime Cost": laborTypeRate.UnitOvertimeCost,
    "Double Time Cost": laborTypeRate.UnitDoubleTimeCost,
    "Labor Type Rate Object ID": laborTypeRate.ObjectID
  }
}