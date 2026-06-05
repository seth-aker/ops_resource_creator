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
  "Integration Mapping"?: string
}

interface IUserDTO {
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

const LICENSE_MAP: Record<TRawLicenseType, TLicenseTypeDTO> = {
  "None": "None",
  "Read Only": "ReadOnly",
  "Full Access": "Full"
}

function CreateUsers() {
  try { 
    setIsScriptFinished(false);
    clearScriptProgress()
    setCurrentScript("CreateUsers")
    openProgressSidebar("Creating Users");

    logEvent("Starting Create Users Script")
    const token = getOpsToken()
    const baseUrl = PropertiesService.getUserProperties().getProperty('opsBaseUrl');
    
    const userData = getSpreadSheetData<IRawUserData>('Users')
    if(!userData || userData.length === 0) {
     
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
    TrackLicense: row["Track License"] ? LICENSE_MAP[row["Track License"]] : undefined,
    FieldEmployeeLicense: row["Field Employee License"] ? LICENSE_MAP[row["Field Employee License"]] : undefined,
    ScheduleLicense: row["Schedule License"] ? LICENSE_MAP[row["Schedule License"]] : undefined,
    MaintainMechanicLicense: row["Maintain Mechanic License"] ? LICENSE_MAP[row["Maintain Mechanic License"]] : undefined,
    MaintainManagerLicense: row["Maintain Manager License"] ? LICENSE_MAP[row["Maintain Manager License"]] : undefined,
    EmployeeIntegrationKey: row["Employee Integration Key"],
    IntegrationMapping: row["Integration Mapping"]
  } as IUserDTO
}