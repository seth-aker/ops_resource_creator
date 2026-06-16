interface IRawLaborType {
  "Business Unit": string,
  "Labor Type ID": string,
  "Labor Type Name": string,
  "Is Inactive?": boolean,
  "Notes"?: string | null
  "Integration Key"?: string | null
  "ObjectID"?: string
}
interface LaborTypeDTO {
  BusinessUnitUniqueName: string,
  LaborTypeID: string,
  Name: string,
  IsInactive: boolean,
  Notes?: string | null,
  IntegrationKey?: string | null,
  ObjectID?: string
}

function CreateLaborTypes() {
  setIsScriptFinished(false)
  clearScriptProgress()
  setCurrentScript("CreateLaborTypes")
  openProgressSidebar("Creating Labor Types")

  logEvent("Starting Create Labor Types Script")
  const token = getOpsToken()
  const baseUrl = getBaseURL()
  
  const laborTypes = getSpreadSheetData<IRawLaborType>('Labor Types')
  if(!laborTypes || laborTypes.length === 0) {
    SpreadsheetApp.getUi().alert("No data found to send!")
    setIsScriptFinished(false);
    return;
  }

  
}