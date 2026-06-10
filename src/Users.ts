interface IRawUserData {
  "Business Unit"?: string,
  "First Name"?: string,
  "Last Name"?: string,
  "Employee ID"?: string,
  "Title"?: string,
  "TID / Mobile Email Address"?: string,
  "Notes"?: string,
  "Track License"?: TRawLicenseType,
  "Field Employee License"?: TRawFieldEmpLicenseType
  "Schedule License"?: TRawLicenseType
  "Maintain Manager License"?: TRawLicenseType,
  "Maintain Mechanic License"?: TRawLicenseType,
  "Employee Integration Key"?: string,
  "Integration Mapping"?: string,
  "Is Inactive?": boolean,
  "ObjectID"?: string
}

interface IUserDTO {
  ObjectID?: string,
  BusinessUnitUniqueName: string,
  FirstName?: string,
  LastName?: string,
  IsInactive: boolean,
  EmployeeID?: string,
  Title?: string,
  MobileEmailAddress?: string
  Notes: string,
  TrackLicense?: TLicenseTypeDTO,
  FieldEmployeeLicense?: TFieldEmployeeLicenseDTO,
  ScheduleLicense?: TLicenseTypeDTO
  MaintainMechanicLicense?: TLicenseTypeDTO,
  MaintainManagerLicense?: TLicenseTypeDTO,
  EmployeeIntegrationKey?: string,
  IntegrationMapping?: string
}
type TRawLicenseType = "None" | "Read Only" | "Full Access"
type TRawFieldEmpLicenseType = "None" | "Full Access"
type TLicenseTypeDTO = "None" | "ReadOnly" | "Full" 
type TFieldEmployeeLicenseDTO = "None" | "Full"

interface GetUserOptions {
  filterQuery: string
}

const USER_SPREADSHEET_KEYS: Array<keyof IRawUserData> = [
  "Business Unit",
  "First Name",
  "Last Name",
  "Employee ID",
  "Title",
  "TID / Mobile Email Address",
  "Notes",
  "Track License",
  "Field Employee License",
  "Schedule License",
  "Maintain Manager License",
  "Maintain Mechanic License",
  "Employee Integration Key",
  "Integration Mapping",
  "Is Inactive?",
  "ObjectID"
]
const LICENSE_MAP = new Map<TRawLicenseType, TLicenseTypeDTO>()
  .set("None", "None")
  .set("Read Only", "ReadOnly")
  .set("Full Access", "Full")
function GetUsers(options: GetUserOptions) {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("GetUsers")
  openProgressSidebar("Getting Users");

  logEvent("Starting get users script")
  const token = getOpsToken();
  const baseUrl = getBaseURL();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let usersSheet = spreadsheet.getSheetByName('Users')
  if(!usersSheet) {
    usersSheet = spreadsheet.insertSheet('Users');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    usersSheet.appendRow(USER_SPREADSHEET_KEYS).getRange(1,1,1, USER_SPREADSHEET_KEYS.length).setTextStyle(bold)
  }
  const currentSpreadSheetData = getSpreadSheetData<IRawUserData>("Users")
  const ui = SpreadsheetApp.getUi();
  if(currentSpreadSheetData.length > 0) {
    const response = ui.alert("The User spreadsheet already has data. This will be overwritten. Do you want to contiune?",
      ui.ButtonSet.YES_NO
    )
    if(response === ui.Button.NO) {
      logEvent("Get Users Script Canceled")
      setIsScriptFinished(true);
      return;
    }
  }
  const headers = createHeaders(token)

  const users = getDatabaseItems<IUserDTO>(`${baseUrl}/Users${options.filterQuery}`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })

  const rowValues = users.map(e => {
    const values = createRawUsers(e)
    const headerValues = usersSheet.getDataRange().getValues()[0] as typeof USER_SPREADSHEET_KEYS
    USER_SPREADSHEET_KEYS.forEach((key) => {
      if(!headerValues.includes(key)) {
        headerValues.push(key)
        usersSheet.getRange(1, headerValues.length, 1,1).setValue(key)
      }
    })
    return headerValues.map(key => values[key] ?? "")
  })
  const startRow = usersSheet.getLastRow() + 1;

  usersSheet.getRange(startRow, 1, rowValues.length, USER_SPREADSHEET_KEYS.length).setValues(rowValues)
  
  logEvent("Script Complete!")
  ui.alert("Script Complete!")
  setIsScriptFinished(true)
}
function CreateUsers() {
  try { 
    setIsScriptFinished(false);
    clearScriptProgress()
    setCurrentScript("CreateUsers")
    openProgressSidebar("Creating Users");

    logEvent("Starting Create Users Script")
    const token = getOpsToken()
    const baseUrl = getBaseURL()
    
    const userData = getSpreadSheetData<IRawUserData>('Users')
    if(!userData || userData.length === 0) {
      SpreadsheetApp.getUi().alert("No data found to send!")
      setIsScriptFinished(false);
    }

    const url = baseUrl + "/user"
    const headers = createHeaders(token);
    const userDTOs = userData.map((row) => {
        return createUserDTO(row);
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
    logEvent("Uploading users...")
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
    }
    logEvent("Script Complete!")
    SpreadsheetApp.getUi().alert("Script Complete!")
    setIsScriptFinished(true);
  } catch (err) {
    setIsScriptFinished(true)
    throw err
  }
}
function UpdateUsers() {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("UpdateUsers")
  openProgressSidebar("Updating Users");

  logEvent("Starting Update Users Script")

  const token = getOpsToken()
  const baseUrl = getBaseURL()
  
  const userData = getSpreadSheetData<IRawUserData>('Users')
  if(!userData || userData.length === 0) {
    SpreadsheetApp.getUi().alert("No data found to send!")
    setIsScriptFinished(false);
  }

  const url = baseUrl + "/user"
  const headers = createHeaders(token);
  const userDTOs = userData.map(createUserDTO)
  let payloads: IUserDTO[] = []
  if(!userDTOs.every(entry => entry.ObjectID && entry.ObjectID.length > 0)) {
    const getOptions = {
      headers,
      method: 'get' as const,
      muteHttpExceptions: true
    }
    const existingUsers = getDatabaseItems<IUserDTO>(url, getOptions)
    payloads = userDTOs.map((each, idx) => {
      const mobileEmailAddress = each.MobileEmailAddress
      const existing = existingUsers.find(u => {
        return mobileEmailAddress === u.MobileEmailAddress
      })
      if(!existing) {
        const errorMessage = `Error: Could not find existing user whose tid matches: ${mobileEmailAddress}`
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
    payloads = userDTOs
  }
  const batchOptions = userDTOs.map(row => {
    const options = {
      url,
      method: 'put' as const,
      headers,
      payload: JSON.stringify(row),
      muteHttpExceptions: true
    }
    return options;
  })
  logEvent("Updating users...")
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
  }
  logEvent("Script Complete")
  SpreadsheetApp.getUi().alert("Script Complete!")
  setIsScriptFinished(true)
}

function createUserDTO(row: IRawUserData) {
  return {
    BusinessUnitUniqueName: row["Business Unit"],
    FirstName: row["First Name"],
    LastName: row["Last Name"],
    IsInactive: false,
    EmployeeID: row["Employee ID"],
    Title: row.Title,
    MobileEmailAddress: row["TID / Mobile Email Address"],
    Notes: row.Notes,
    TrackLicense: row["Track License"] ? LICENSE_MAP.get(row["Track License"]) : undefined,
    FieldEmployeeLicense: row["Field Employee License"] ? LICENSE_MAP.get(row["Field Employee License"]) : undefined,
    ScheduleLicense: row["Schedule License"] ? LICENSE_MAP.get(row["Schedule License"]) : undefined,
    MaintainMechanicLicense: row["Maintain Mechanic License"] ? LICENSE_MAP.get(row["Maintain Mechanic License"]) : undefined,
    MaintainManagerLicense: row["Maintain Manager License"] ? LICENSE_MAP.get(row["Maintain Manager License"]) : undefined,
    EmployeeIntegrationKey: row["Employee Integration Key"],
    IntegrationMapping: row["Integration Mapping"]
  } as IUserDTO
}
function createRawUsers(u: IUserDTO): IRawUserData {
  const licenseOptions = Array.from(LICENSE_MAP.entries())
  return {
    "Business Unit": u.BusinessUnitUniqueName,
    "First Name": u.FirstName,
    "Last Name": u.LastName,
    "Is Inactive?": u.IsInactive,
    "Employee ID": u.EmployeeID,
    "Title": u.Title,
    "TID / Mobile Email Address": u.MobileEmailAddress,
    "Notes": u.Notes,
    "Track License": licenseOptions.find(([_key, val]) => u.TrackLicense === val)?.[0],
    "Schedule License": licenseOptions.find(([_key, val]) => u.ScheduleLicense === val)?.[0],
    "Maintain Manager License": licenseOptions.find(([_key, val]) => u.MaintainManagerLicense === val)?.[0],
    "Maintain Mechanic License": licenseOptions.find(([_key, val]) => u.MaintainMechanicLicense === val)?.[0],
    "Employee Integration Key": u.EmployeeIntegrationKey,
    "Integration Mapping": u.IntegrationMapping,
    "ObjectID": u.ObjectID
  }
}
