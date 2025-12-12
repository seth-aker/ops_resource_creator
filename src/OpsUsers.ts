interface IOpsUser {
  BusinessUnitUniqueName: string,
  WindowsAccountName?: string,
  FirstName?: string,
  LastName?: string,
  IsInactive?: boolean,
  EmployeeID?: string,
  EmployeeIntegrationKey?: string,
  Title?: string,
  EmailAddress?: string,
  MobileDevicePin?: string,
  TrackLicense?: string,
  FieldEmployeeLicense?: string,
  MobileEmailAddress?: string,
  ScheduleLicense?: string,
  MaintainMechanicLicense?: string,
  MaintainManagerLicense?: string,
  Notes?: string,
  IntegrationMapping?: string
  ObjectID?: string // Only exists on created Users
}

function CreateUsers() {
  const {token, baseUrl} = authenticate();
  const data = getSpreadSheetData<IOpsUser>("Ops Users");

  if(!data || data.length === 0) {
    Logger.log("No data to send!")
    SpreadsheetApp.getUi().alert("No data to send!");
  }

  const headers = createHeaders(token);
  const url = baseUrl + '/User'
  const failedRows: number[] = [];
  const existingRows: number[] = [];
  const batchOptions = data.map(row => {
    const options = {
      url, 
      method: 'post' as const,
      headers, 
      payload: JSON.stringify(row),
      muteHttpExceptions: true
    }
    return options
  })

  try {
    const responses = batchFetch(batchOptions);
    responses.forEach((response, index) => {
      const responseCode = response.getResponseCode()
      if(responseCode === 409 || responseCode === 200) {
        Logger.log(`Row ${index + 2}: Already exists in the database.`)
        existingRows.push(index + 2)
      } else if(responseCode >= 400) {
        Logger.log(`User at row ${index + 2} failed with status code: ${responseCode}. Error Message: ${response.getContentText()}`)
        failedRows.push(index + 2)
      } else {
        Logger.log(`Row ${index + 2}: Successfully created.`);
      }
    })
  } catch (err) {
    Logger.log(`An unexpected error occured: ${err}`)
    throw new Error(`An unexpected error occurred created materials. Please check the logs for more details`)
  }

  if(failedRows.length === 0 && existingRows.length === 0) {
    SpreadsheetApp.getUi().alert("All users created successfully!")
    return;
  }
  if(failedRows.length > 0) {
    highlightRows(failedRows, 'red')
    SpreadsheetApp.getUi().alert(`${failedRows.length} users failed to be created at rows: ${failedRows.join(", ")}`)
  }
  if(existingRows.length > 0) {
    highlightRows(existingRows, 'yellow');
    SpreadsheetApp.getUi().alert(`${existingRows.length} users already existed in the database. Rows: ${existingRows.join(', ')}` )
  }
}