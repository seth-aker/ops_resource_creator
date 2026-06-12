interface IMaterialDTO {
  ObjectID?: string;
  MaterialID: string;
  AlternateMaterialID?: string | null;
  IntegrationKey?: string | null;
  IsInactive?: boolean;
  BusinessUnitUniqueName: string;
  Description: string;
  Notes?: string | null;
  Category?: string | null;
  Subcategory?: string | null;
  UnitOfMeasure: string;
  NumberOf?: number | null;
  UnitCost?: number | null;
  IsTemporaryMaterial?: boolean | null;
  IsTrackableMaterial?: boolean | null;
  UnitWeight?: number | null;
  WeightUnitOfMeasure?: string | null;
  UnitVolume?: number | null;
  VolumeUnitOfMeasure?: string | null;
  LossPercent?: number | null;
}

interface IRawMaterials {
  "Material ID": string;
  "Alternate Material ID"?: string | null;
  "Integration Key"?: string | null;
  "Is Inactive?"?: boolean;
  "Business Unit": string;
  "Description": string;
  "Notes"?: string | null;
  "Category"?: string | null;
  "Subcategory"?: string | null;
  "Unit Of Measure": string;
  "Number Of"?: number | null;
  "Unit Cost"?: number | null;
  "Is Temporary Material?"?: boolean | null;
  "Is Trackable Material?"?: boolean | null;
  "Unit Weight"?: number | null;
  "Weight Unit Of Measure"?: string | null;
  "Unit Volume"?: number | null;
  "Volume Unit Of Measure"?: string | null;
  "Loss Percent"?: number | null;
  "ObjectID"?: string
}
const MATERIAL_SPREADSHEET_KEYS: Array<keyof IRawMaterials> = [
  "Business Unit",
  "Description",
  "Material ID",
  "Alternate Material ID",
  "Integration Key",
  "Category",
  "Subcategory",
  "Unit Of Measure",
  "Unit Cost",
  "Notes",
  "Number Of",
  "Is Temporary Material?",
  "Is Trackable Material?",
  "Unit Weight",
  "Weight Unit Of Measure",
  "Unit Volume",
  "Volume Unit Of Measure",
  "Loss Percent",
  "Is Inactive?",
  "ObjectID"
];

const MATERIAL_FILTER_OPTIONS: IFilterByOptions[] = [
  { value: 'ObjectID', type: 'string' },
  { value: 'MaterialID', type: 'string' },
  { value: 'AlternateMaterialID', type: 'string' },
  { value: 'IntegrationKey', type: 'string' },
  { value: 'IsInactive', type: 'boolean' },
  { value: 'BusinessUnitUniqueName', type: 'string' },
  { value: 'Description', type: 'string' },
  { value: 'Notes', type: 'string' },
  { value: 'Category', type: 'string' },
  { value: 'Subcategory', type: 'string' },
  { value: 'UnitOfMeasure', type: 'string' },
  { value: 'NumberOf', type: 'number' },
  { value: 'UnitCost', type: 'number' },
  { value: 'IsTemporaryMaterial', type: 'boolean' },
  { value: 'IsTrackableMaterial', type: 'boolean' },
  { value: 'UnitWeight', type: 'number' },
  { value: 'WeightUnitOfMeasure', type: 'string' },
  { value: 'UnitVolume', type: 'number' },
  { value: 'VolumeUnitOfMeasure', type: 'string' },
  { value: 'LossPercent', type: 'number' },
];

function DisplayMaterialsFilterBuilder() {
  const template = HtmlService.createTemplateFromFile('BuildFilter') as IBuildFilterTemplate
  template.filterByOptions = MATERIAL_FILTER_OPTIONS
  template.serverFunctionName = 'GetMaterials'
  const html = template.evaluate()
    .setHeight(900)
    .setWidth(1100)
  const ui = SpreadsheetApp.getUi()
  ui.showModalDialog(html, "Materials Filter")
}

function GetMaterials(options: GetOptions) {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("GetMaterials")
  openProgressSidebar("Getting Materials");

  logEvent("Starting get materials script")
  const token = getOpsToken();
  const baseUrl = getBaseURL();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let materialsSheet = spreadsheet.getSheetByName('Materials')
  if(!materialsSheet) {
    materialsSheet = spreadsheet.insertSheet('Materials');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    materialsSheet.appendRow(MATERIAL_SPREADSHEET_KEYS).getRange(1,1,1, MATERIAL_SPREADSHEET_KEYS.length).setTextStyle(bold)
  }
  const currentSpreadSheetData = getSpreadSheetData<IRawUserData>("Materials")
  const ui = SpreadsheetApp.getUi();
  if(currentSpreadSheetData.length > 0) {
    const response = ui.alert("WARNING: any existing data will be overwritten. Do you want to continue?",
      ui.ButtonSet.YES_NO
    )
    if(response === ui.Button.NO) {
      logEvent("Get Materials Script Canceled")
      setIsScriptFinished(true);
      return;
    } else {
      materialsSheet.getRange(2, 1, materialsSheet.getLastRow(), materialsSheet.getLastColumn()).clearContent()
    }
  }
  const headers = createHeaders(token)

  const users = getDatabaseItems<IMaterialDTO>(`${baseUrl}/Material${options.filterQuery}`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })
  logEvent(`${users.length} materials recieved.`)
  const headerValues = materialsSheet.getDataRange().getValues()[0] as typeof MATERIAL_SPREADSHEET_KEYS
  MATERIAL_SPREADSHEET_KEYS.forEach((key) => {
    if(!headerValues.includes(key)) {
      headerValues.push(key)
      materialsSheet.getRange(1, headerValues.length, 1,1).setValue(key)
    }
  })

  // arranges each row to match the order of the headers in the spreadsheet.
  const rowValues = users.map(e => {
    const values = mapMaterialDTOToRaw(e)
    return headerValues.map(key => values[key] ?? "")
  })
  const startRow = 2;

  materialsSheet.getRange(startRow, 1, rowValues.length, headerValues.length).setValues(rowValues)
  
  logEvent("Script Complete!")
  ui.alert("Script Complete!")
  setIsScriptFinished(true)
}
function CreateMaterials() {
  try{
    setIsScriptFinished(false);
    clearScriptProgress()
    setCurrentScript("CreateMaterials")
    openProgressSidebar("Creating Materials");
    
    logEvent("Starting Create Materials Script")
    const token = getOpsToken()
    const baseUrl = getBaseURL()
    const materialData = getSpreadSheetData<IRawMaterials>('Materials')
    if(!materialData || materialData.length === 0) {
      SpreadsheetApp.getUi().alert("No data found to send!")
      setIsScriptFinished(false);
      return
    }

    const url = baseUrl + "/material"
    const headers = createHeaders(token);
    const materialDTO = materialData.map((row) => {
        return mapRawToMaterialDTO(row);
      })
    const batchOptions = materialDTO.map(row => {
      const options = {
        url,
        method: 'post' as const,
        headers,
        payload: JSON.stringify(row),
        muteHttpExceptions: true
      }
      return options;
    })
    logEvent(`Uploading ${batchOptions.length} materials...`)
    const results = batchFetch(batchOptions);
    
    const failedRows: number[] = []
    results.forEach((result, idx) => {
      const code = result.getResponseCode();
      if(code >= 400) {
        failedRows.push(idx);
        writeLogToSpreadsheet(`${code} Error: ${result.getContentText()}`)
      }
    })
    
    if(failedRows.length > 0) {
      const errorMessages = failedRows.map(idx => `${results[idx].getResponseCode()} Error: ${results[idx].getContentText()}`)
      const failedResults = errorMessages.map((message, idx) => `Row ${failedRows[idx] + 2}: ${message}`) 
      logEvent([`Some rows failed!:`, ...failedResults])
      highlightRows(failedRows.map(each => each + 2), 'red');
    } else {
      logEvent("All materials created successfully!")
    }
    logEvent("Script Complete!")
    SpreadsheetApp.getUi().alert("Script Complete!")
    setIsScriptFinished(true);
  } catch (err) {
    setIsScriptFinished(true)
    throw err
  }
}

function UpdateMaterials() {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("UpdateMaterials")
  openProgressSidebar("Updating Materials");

  logEvent("Starting Update Materials Script")

  const token = getOpsToken()
  const baseUrl = getBaseURL()
  
  const materialsData = getSpreadSheetData<IRawMaterials>('Materials')
  if(!materialsData || materialsData.length === 0) {
    SpreadsheetApp.getUi().alert("No data found to send!")
    setIsScriptFinished(false);
    return;
  }

  const url = baseUrl + "/material"
  const headers = createHeaders(token);
  const materialDTOS = materialsData.map(mapRawToMaterialDTO)
  let payloads: IMaterialDTO[] = []
  if(!materialDTOS.every(entry => entry.ObjectID && entry.ObjectID.length > 0)) {
    const getOptions = {
      headers,
      method: 'get' as const,
      muteHttpExceptions: true
    }
    const existingMaterials = getDatabaseItems<IMaterialDTO>(url, getOptions)
    payloads = materialDTOS.map((each, idx) => {
      const matID = each.MaterialID
      const intKey = each.IntegrationKey
      const existing = existingMaterials.find((m) => {
        const idMatches = matID === m.MaterialID
        if(m.IntegrationKey) {
          return idMatches && m.IntegrationKey === intKey
        } else {
          return idMatches
        }
      })
      if(!existing) {
        const errorMessage = `Error: Could not find existing material with id: ${matID}${each.IntegrationKey ? ` and Integration Key: ${each.IntegrationKey}`: ""}`
        logEvent(errorMessage)
        highlightRows([idx + 2], 'red')
        throw new Error(errorMessage)
      }
      return {
        ...existing,
        ...each
      }
    })
  } else {
    payloads = materialDTOS
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
  logEvent(`Updating ${batchOptions.length} materials...`)
  const results = batchFetch(batchOptions);
  const failed = [] as number[]
  
  results.forEach((res, idx) => {
    const code = res.getResponseCode()
    if(code > 299) {
      writeLogToSpreadsheet(`Error Code: ${code}, Message: ${res.getContentText()}`)
      failed.push(idx)
    }
  })
  if(failed.length > 0) {
    const failureMessages = failed.map(idx => `Row ${idx + 2}: ${results[idx].getContentText()}`)
    logEvent(["Some rows failed", ...failureMessages])
    highlightRows(failed.map(f => f + 2), 'red')
  } else {
    logEvent("All materials updated successfully!")
  }
  logEvent("Script Complete")
  SpreadsheetApp.getUi().alert("Script Complete!")
  setIsScriptFinished(true)
}

function DeleteMaterials() {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("DeleteMaterials")
  openProgressSidebar("Deleting Materials");

  logEvent("Starting Delete Materials Script")

  const token = getOpsToken()
  const baseUrl = getBaseURL()
  
  const materialsData = getSpreadSheetData<IRawMaterials>('Materials')
  if(!materialsData || materialsData.length === 0) {
    SpreadsheetApp.getUi().alert("No data found to send!")
    setIsScriptFinished(false);
    return;
  }

  const url = baseUrl + "/material"
  const headers = createHeaders(token);
  if(!materialsData.every(row => row.ObjectID)) {
    logEvent("Error, not all of the material rows have an ObjectID. This is required to delete them.")
    logEvent("Canceling Script")
    setIsScriptFinished(true)
    return
  }
  const batchOptions = materialsData.map(row => {
    const options = {
      url: `${url}?ObjectID=${row.ObjectID}`,
      method: 'delete' as const,
      headers,
      payload: JSON.stringify(row),
      muteHttpExceptions: true
    }
    return options;
  })
  logEvent(`Deleting ${batchOptions.length} materials...`)
  const results = batchFetch(batchOptions);
  const failed = [] as number[]
  
  results.forEach((res, idx) => {
    const code = res.getResponseCode()
    if(code > 299) {
      writeLogToSpreadsheet(`Error Code: ${code}, Message: ${res.getContentText()}`)
      failed.push(idx)
    }
  })
  if(failed.length > 0) {
    const failureMessages = failed.map(idx => `Row ${idx + 2}: ${results[idx].getContentText()}`)
    logEvent(["Some rows failed", ...failureMessages])
    highlightRows(failed.map(f => f + 2), 'red')
  } else {
    logEvent("All materials deleted successfully!")
  }
  logEvent("Script Complete")
  SpreadsheetApp.getUi().alert("Script Complete!")
  setIsScriptFinished(true)
}

function mapRawToMaterialDTO(raw: IRawMaterials): IMaterialDTO {
  return {
    MaterialID: raw["Material ID"],
    AlternateMaterialID: raw["Alternate Material ID"],
    IntegrationKey: raw["Integration Key"],
    IsInactive: raw["Is Inactive?"], 
    BusinessUnitUniqueName: raw["Business Unit"], 
    Description: raw["Description"],
    Notes: raw["Notes"],
    Category: raw["Category"],
    Subcategory: raw["Subcategory"],
    UnitOfMeasure: raw["Unit Of Measure"],
    NumberOf: raw["Number Of"],
    UnitCost: raw["Unit Cost"],
    IsTemporaryMaterial: raw["Is Temporary Material?"], 
    IsTrackableMaterial: raw["Is Trackable Material?"], 
    UnitWeight: raw["Unit Weight"],
    WeightUnitOfMeasure: raw["Weight Unit Of Measure"],
    UnitVolume: raw["Unit Volume"],
    VolumeUnitOfMeasure: raw["Volume Unit Of Measure"],
    LossPercent: raw["Loss Percent"],
    ObjectID: raw.ObjectID
  };
}
function mapMaterialDTOToRaw(dto: IMaterialDTO): IRawMaterials {
  return {
    "Material ID": dto.MaterialID,
    "Alternate Material ID": dto.AlternateMaterialID,
    "Integration Key": dto.IntegrationKey,
    "Is Inactive?": dto.IsInactive,
    "Business Unit": dto.BusinessUnitUniqueName,
    "Description": dto.Description,
    "Notes": dto.Notes,
    "Category": dto.Category,
    "Subcategory": dto.Subcategory,
    "Unit Of Measure": dto.UnitOfMeasure,
    "Number Of": dto.NumberOf,
    "Unit Cost": dto.UnitCost,
    "Is Temporary Material?": dto.IsTemporaryMaterial,
    "Is Trackable Material?": dto.IsTrackableMaterial,
    "Unit Weight": dto.UnitWeight,
    "Weight Unit Of Measure": dto.WeightUnitOfMeasure,
    "Unit Volume": dto.UnitVolume,
    "Volume Unit Of Measure": dto.VolumeUnitOfMeasure,
    "Loss Percent": dto.LossPercent,
    "ObjectID": dto.ObjectID
  };
}