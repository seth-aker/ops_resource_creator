interface ILaborTypeLightweightDTO {
  LaborTypeID: string;
  IntegrationKey?: string | null;
}

interface IEmployeeDTO {
  BusinessUnitUniqueName: string;
  FirstName?: string | null;
  LastName: string;
  MiddleInitial?: string | null;
  Nickname?: string | null;
  EmailAddress?: string | null;
  EmployeeID: string;
  IntegrationKey?: string | null;
  JobTitle?: string | null;
  DefaultLaborTypeID?: string | null;
  DefaultLaborTypeIntegrationKey?: string | null;
  LaborTypes?: ILaborTypeLightweightDTO[] | null;
  HomePhone?: string | null;
  CellPhone?: string | null;
  WorkPhone?: string | null;
  MobileEmailAddress?: string | null;
  ReviewerPIN?: string | null;
  MobileDevicePin?: string | null;
  NotificationType?: string | null;
  Notes?: string | null;
  IsFieldEmployee?: boolean | null;
  IsForeman?: boolean;
  IsProjectManager?: boolean;
  IsSupervisor?: boolean;
  IsPurchaseOrderApprover?: boolean;
  IsBuyer?: boolean;
  IsFieldLogReviewer?: boolean;
  IsDriver?: boolean;
  IsMechanic?: boolean;
  IsTruckDriver?: boolean;
  IsInactive?: boolean;
  IsOptedInForSMS?: boolean;
  InactiveReason?: string | null;
  ObjectID?: string;
}

interface IRawEmployee {
  "Business Unit": string;
  "First Name"?: string | null;
  "Last Name": string;
  "Middle Initial"?: string | null;
  "Nickname"?: string | null;
  "Email Address"?: string | null;
  "Employee ID": string;
  "Integration Key"?: string | null;
  "Job Title"?: string | null;
  "Default Labor Type ID"?: string | null;
  "Default Labor Type Integration Key"?: string | null;
  "Labor Types"?: string | null;
  "Home Phone"?: string | null;
  "Cell Phone"?: string | null;
  "Work Phone"?: string | null;
  "Mobile Email Address"?: string | null;
  "Reviewer PIN"?: string | null;
  "Mobile Device Pin"?: string | null;
  "Notification Type"?: string | null;
  "Notes"?: string | null;
  "Is Field Employee?"?: boolean | null;
  "Is Foreman?"?: boolean;
  "Is Project Manager?"?: boolean;
  "Is Supervisor?"?: boolean;
  "Is Purchase Order Approver?"?: boolean;
  "Is Buyer?"?: boolean;
  "Is Field Log Reviewer?"?: boolean;
  "Is Driver?"?: boolean;
  "Is Mechanic?"?: boolean;
  "Is Truck Driver?"?: boolean;
  "Is Inactive?"?: boolean;
  "Is Opted In For SMS?"?: boolean;
  "Inactive Reason"?: string | null;
  "Object ID"?: string;
}
const EMPLOYEE_FILTER_OPTIONS: IFilterByOptions[] = [
  { value: 'BusinessUnitUniqueName', type: 'string' },
  { value: 'FirstName', type: 'string' },
  { value: 'LastName', type: 'string' },
  { value: 'MiddleInitial', type: 'string' },
  { value: 'Nickname', type: 'string' },
  { value: 'EmailAddress', type: 'string' },
  { value: 'EmployeeID', type: 'string' },
  { value: 'IntegrationKey', type: 'string' },
  { value: 'JobTitle', type: 'string' },
  { value: 'DefaultLaborTypeID', type: 'string' },
  { value: 'DefaultLaborTypeIntegrationKey', type: 'string' },
  { value: 'HomePhone', type: 'string' },
  { value: 'CellPhone', type: 'string' },
  { value: 'WorkPhone', type: 'string' },
  { value: 'MobileEmailAddress', type: 'string' },
  { value: 'ReviewerPIN', type: 'string' },
  { value: 'MobileDevicePin', type: 'string' },
  { value: 'NotificationType', type: 'string' },
  { value: 'Notes', type: 'string' },
  { value: 'IsFieldEmployee', type: 'boolean' },
  { value: 'IsForeman', type: 'boolean' },
  { value: 'IsProjectManager', type: 'boolean' },
  { value: 'IsSupervisor', type: 'boolean' },
  { value: 'IsPurchaseOrderApprover', type: 'boolean' },
  { value: 'IsBuyer', type: 'boolean' },
  { value: 'IsFieldLogReviewer', type: 'boolean' },
  { value: 'IsDriver', type: 'boolean' },
  { value: 'IsMechanic', type: 'boolean' },
  { value: 'IsTruckDriver', type: 'boolean' },
  { value: 'IsInactive', type: 'boolean' },
  { value: 'IsOptedInForSMS', type: 'boolean' },
  { value: 'InactiveReason', type: 'string' },
  { value: 'ObjectID', type: 'string' }
];
const RAW_EMPLOYEE_KEYS: Array<keyof IRawEmployee> = [
  "Business Unit",
  "First Name",
  "Last Name",
  "Middle Initial",
  "Nickname",
  "Email Address",
  "Employee ID",
  "Integration Key",
  "Job Title",
  "Default Labor Type ID",
  "Default Labor Type Integration Key",
  "Labor Types",
  "Home Phone",
  "Cell Phone",
  "Work Phone",
  "Mobile Email Address",
  "Reviewer PIN",
  "Mobile Device Pin",
  "Notification Type",
  "Notes",
  "Is Field Employee?",
  "Is Foreman?",
  "Is Project Manager?",
  "Is Supervisor?",
  "Is Purchase Order Approver?",
  "Is Buyer?",
  "Is Field Log Reviewer?",
  "Is Driver?",
  "Is Mechanic?",
  "Is Truck Driver?",
  "Is Inactive?",
  "Is Opted In For SMS?",
  "Inactive Reason",
  "Object ID"
];
function DisplayEmployeeFilterBuilder() {
  const template = HtmlService.createTemplateFromFile('BuildFilter') as IBuildFilterTemplate
  template.filterByOptions = EMPLOYEE_FILTER_OPTIONS;
  template.serverFunctionName = "GetEmployees"
  const html = template.evaluate()
    .setHeight(900)
    .setWidth(1100)
  const ui = SpreadsheetApp.getUi()
  ui.showModalDialog(html, "Employee Filter")
}
function GetEmployees(options: GetOptions) {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("GetEmployees")
  openProgressSidebar("Getting Employees");

  logEvent("Starting get employees script")
  const token = getOpsToken();
  const baseUrl = getBaseURL();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let employeesSheet = spreadsheet.getSheetByName('Employees')
  if(!employeesSheet) {
    employeesSheet = spreadsheet.insertSheet('Employees');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    employeesSheet.appendRow(RAW_EMPLOYEE_KEYS).getRange(1,1,1, RAW_EMPLOYEE_KEYS.length).setTextStyle(bold)
  }
  const currentSpreadSheetData = getSpreadSheetData<IRawEmployee>("Employees")
  const ui = SpreadsheetApp.getUi();
  if(currentSpreadSheetData.length > 0) {
    const response = ui.alert("The Employees spreadsheet already has data. This will be overwritten. Do you want to contiune?",
      ui.ButtonSet.YES_NO
    )
    if(response === ui.Button.NO) {
      logEvent("Get Employees Script Canceled")
      setIsScriptFinished(true);
      return;
    } else {
      employeesSheet.getRange(2, 1, employeesSheet.getLastRow(), employeesSheet.getLastColumn()).clearContent()
    }
  }
  const headers = createHeaders(token)

  const employees = getDatabaseItems<IEmployeeDTO>(`${baseUrl}/Employee${options.filterQuery}`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })
  logEvent(`${employees.length} employees recieved.`)
  const headerValues = employeesSheet.getDataRange().getValues()[0] as typeof RAW_EMPLOYEE_KEYS
  RAW_EMPLOYEE_KEYS.forEach((key) => {
    if(!headerValues.includes(key)) {
      headerValues.push(key)
      employeesSheet.getRange(1, headerValues.length, 1,1).setValue(key)
    }
  })

  const rowValues = employees.map(e => {
    const values = createRawEmployee(e)
    return headerValues.map(key => values[key] ?? "")
  })
  const startRow = 2;

  employeesSheet.getRange(startRow, 1, rowValues.length, headerValues.length).setValues(rowValues)
  
  logEvent("Script Complete!")
  ui.alert("Script Complete!")
  setIsScriptFinished(true)
}

function createRawEmployee(employee: IEmployeeDTO): IRawEmployee {
  let laborTypesString: string | null = null;
  
  if (employee.LaborTypes && employee.LaborTypes.length > 0) {
    laborTypesString = employee.LaborTypes.map((lt) => {
      return lt.IntegrationKey 
        ? `${lt.LaborTypeID}|${lt.IntegrationKey}` 
        : lt.LaborTypeID;
    }).join(',');
  }

  return {
    "Business Unit": employee.BusinessUnitUniqueName,
    "First Name": employee.FirstName,
    "Last Name": employee.LastName,
    "Middle Initial": employee.MiddleInitial,
    "Nickname": employee.Nickname,
    "Email Address": employee.EmailAddress,
    "Employee ID": employee.EmployeeID,
    "Integration Key": employee.IntegrationKey,
    "Job Title": employee.JobTitle,
    "Default Labor Type ID": employee.DefaultLaborTypeID,
    "Default Labor Type Integration Key": employee.DefaultLaborTypeIntegrationKey,
    "Labor Types": laborTypesString,
    "Home Phone": employee.HomePhone,
    "Cell Phone": employee.CellPhone,
    "Work Phone": employee.WorkPhone,
    "Mobile Email Address": employee.MobileEmailAddress,
    "Reviewer PIN": employee.ReviewerPIN,
    "Mobile Device Pin": employee.MobileDevicePin,
    "Notification Type": employee.NotificationType,
    "Notes": employee.Notes,
    "Is Field Employee?": employee.IsFieldEmployee,
    "Is Foreman?": employee.IsForeman,
    "Is Project Manager?": employee.IsProjectManager,
    "Is Supervisor?": employee.IsSupervisor,
    "Is Purchase Order Approver?": employee.IsPurchaseOrderApprover,
    "Is Buyer?": employee.IsBuyer,
    "Is Field Log Reviewer?": employee.IsFieldLogReviewer,
    "Is Driver?": employee.IsDriver,
    "Is Mechanic?": employee.IsMechanic,
    "Is Truck Driver?": employee.IsTruckDriver,
    "Is Inactive?": employee.IsInactive,
    "Is Opted In For SMS?": employee.IsOptedInForSMS,
    "Inactive Reason": employee.InactiveReason,
    "Object ID": employee.ObjectID
  };
}

function createEmployeeDTO(raw: IRawEmployee): IEmployeeDTO {
  let laborTypesArray: ILaborTypeLightweightDTO[] | null = null;
  
  if (raw["Labor Types"] && raw["Labor Types"].trim() !== "") {
    laborTypesArray = raw["Labor Types"].split(',').map((lt) => {
      const parts = lt.split('|');
      
      const laborType: ILaborTypeLightweightDTO = {
        LaborTypeID: parts[0].trim()
      };
      
      if (parts.length > 1 && parts[1].trim() !== "") {
        laborType.IntegrationKey = parts[1].trim();
      } else {
        laborType.IntegrationKey = null;
      }
      
      return laborType;
    });
  }

  return {
    BusinessUnitUniqueName: raw["Business Unit"],
    FirstName: raw["First Name"],
    LastName: raw["Last Name"],
    MiddleInitial: raw["Middle Initial"],
    Nickname: raw["Nickname"],
    EmailAddress: raw["Email Address"],
    EmployeeID: raw["Employee ID"],
    IntegrationKey: raw["Integration Key"],
    JobTitle: raw["Job Title"],
    DefaultLaborTypeID: raw["Default Labor Type ID"],
    DefaultLaborTypeIntegrationKey: raw["Default Labor Type Integration Key"],
    LaborTypes: laborTypesArray,
    HomePhone: raw["Home Phone"],
    CellPhone: raw["Cell Phone"],
    WorkPhone: raw["Work Phone"],
    MobileEmailAddress: raw["Mobile Email Address"],
    ReviewerPIN: raw["Reviewer PIN"],
    MobileDevicePin: raw["Mobile Device Pin"],
    NotificationType: raw["Notification Type"],
    Notes: raw["Notes"],
    IsFieldEmployee: raw["Is Field Employee?"],
    IsForeman: raw["Is Foreman?"],
    IsProjectManager: raw["Is Project Manager?"],
    IsSupervisor: raw["Is Supervisor?"],
    IsPurchaseOrderApprover: raw["Is Purchase Order Approver?"],
    IsBuyer: raw["Is Buyer?"],
    IsFieldLogReviewer: raw["Is Field Log Reviewer?"],
    IsDriver: raw["Is Driver?"],
    IsMechanic: raw["Is Mechanic?"],
    IsTruckDriver: raw["Is Truck Driver?"],
    IsInactive: raw["Is Inactive?"],
    IsOptedInForSMS: raw["Is Opted In For SMS?"],
    InactiveReason: raw["Inactive Reason"],
    ObjectID: raw["Object ID"]
  };
}