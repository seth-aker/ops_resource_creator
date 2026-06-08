interface EquipmentPart {
  PartID: string;
  IntegrationKey?: string | null;
  Quantity?: number | null;
  RowVersion?: number;
}

interface EquipmentTag {
  Category: string;
  Notes?: string | null;
  RowVersion?: number;
}

interface IEquipmentDTO {
  ObjectID?: string;
  EquipmentID: string;
  IntegrationKey?: string | null;
  IsInactive?: boolean;
  BusinessUnitUniqueName: string;
  Description: string;
  EquipmentTypeID?: string | null;
  EquipmentTypeIntegrationKey?: string | null;
  MobilityType: string;
  OwnershipType: string;
  OrganizationID?: string | null;
  OrganizationIntegrationKey?: string | null;
  OperatorContactName?: string | null;
  OperatorEmployeeID?: string | null;
  OperatorEmployeeIntegrationKey?: string | null;
  SerialNumber?: string | null;
  Manufacturer?: string | null;
  Model?: string | null;
  Year?: number | null;
  Notes?: string | null;
  LocationName?: string | null;
  ExcludeFromFieldLogs?: boolean;
  Length?: number | null;
  LengthUnitOfMeasure?: string | null;
  Width?: number | null;
  WidthUnitOfMeasure?: string | null;
  Height?: number | null;
  HeightUnitOfMeasure?: string | null;
  MaxWeight?: number | null;
  MaxWeightUnitOfMeasure?: string | null;
  GroundPressure?: number | null;
  GroundPressureUnitOfMeasure?: string | null;
  CombinedWeight?: number | null;
  CombinedWeightUnitOfMeasure?: string | null;
  TareWeight?: number | null;
  TareWeightUnitOfMeasure?: string | null;
  LicensePlate?: string | null;
  Color?: string | null;
  Lojack?: string | null;
  HUTSticker?: string | null;
  EZPass?: string | null;
  ProductionDate?: string | null;
  Engine?: string | null;
  EngineArrangement?: string | null;
  EngineSerialNumber?: string | null;
  FuelType?: string | null;
  FuelTankCapacity?: number | null;
  InitialFuelReading?: number | null;
  InitialFuelCost?: number | null;
  FuelUnitOfMeasure?: string | null;
  InitialFuelDate?: string | null;
  HorsePower?: string | null;
  TransmissionModel?: string | null;
  TransmissionSerialNumber?: string | null;
  TireSize?: string | null;
  WheelType?: string | null;
  TrackType?: string | null;
  BrakeType?: string | null;
  CuttingEdge?: string | null;
  G_E_T?: string | null;
  HydraulicPumpType?: string | null;
  HydraulicFlowRate?: string | null;
  PurchasedFrom?: string | null;
  PurchasedDate?: string | null;
  PurchasedPrice?: number | null;
  TitleHolder?: string | null;
  SoldTo?: string | null;
  DispositionDate?: string | null;
  SalePrice?: number | null;
  InsuranceValue?: number | null;
  CCAClass?: string | null;
  MarketValue?: number | null;
  RentalNumber?: string | null;
  StartDate?: string | null;
  ReturnDate?: string | null;
  DailyRental?: number | null;
  WeeklyRental?: number | null;
  MonthlyRental?: number | null;
  EquipmentUpc?: string | null;
}

interface IRawEquipmentPart {
  "Part ID": string;
  "Integration Key"?: string | null;
  "Quantity"?: number | null;
}

interface IRawEquipmentTag {
  "Category": string;
  "Notes"?: string | null;
}

interface IRawEquipment {
  "ObjectID"?: string;
  "Equipment ID": string;
  "Integration Key"?: string | null;
  "Is Inactive"?: boolean;
  "Business Unit Unique Name": string;
  "Description": string;
  "Equipment Type ID"?: string | null;
  "Equipment Type Integration Key"?: string | null;
  "Mobility Type": string;
  "Ownership Type": string;
  "Organization ID"?: string | null;
  "Organization Integration Key"?: string | null;
  "Operator Contact Name"?: string | null;
  "Operator Employee ID"?: string | null;
  "Operator Employee Integration Key"?: string | null;
  "Serial Number"?: string | null;
  "Manufacturer"?: string | null;
  "Model"?: string | null;
  "Year"?: number | null;
  "Notes"?: string | null;
  "Location Name"?: string | null;
  "Exclude From Field Logs"?: boolean;
  "Equipment Parts"?: IRawEquipmentPart[] | null;
  "Equipment Tags"?: IRawEquipmentTag[] | null;
  "Length"?: number | null;
  "Length Unit Of Measure"?: string | null;
  "Width"?: number | null;
  "Width Unit Of Measure"?: string | null;
  "Height"?: number | null;
  "Height Unit Of Measure"?: string | null;
  "Max Weight"?: number | null;
  "Max Weight Unit Of Measure"?: string | null;
  "Ground Pressure"?: number | null;
  "Ground Pressure Unit Of Measure"?: string | null;
  "Combined Weight"?: number | null;
  "Combined Weight Unit Of Measure"?: string | null;
  "Tare Weight"?: number | null;
  "Tare Weight Unit Of Measure"?: string | null;
  "License Plate"?: string | null;
  "Color"?: string | null;
  "Lojack"?: string | null;
  "HUT Sticker"?: string | null;
  "EZ Pass"?: string | null;
  "Production Date"?: string | null;
  "Engine"?: string | null;
  "Engine Arrangement"?: string | null;
  "Engine Serial Number"?: string | null;
  "Fuel Type"?: string | null;
  "Fuel Tank Capacity"?: number | null;
  "Initial Fuel Reading"?: number | null;
  "Initial Fuel Cost"?: number | null;
  "Fuel Unit Of Measure"?: string | null;
  "Initial Fuel Date"?: string | null;
  "Horse Power"?: string | null;
  "Transmission Model"?: string | null;
  "Transmission Serial Number"?: string | null;
  "Tire Size"?: string | null;
  "Wheel Type"?: string | null;
  "Track Type"?: string | null;
  "Brake Type"?: string | null;
  "Cutting Edge"?: string | null;
  "G.E.T."?: string | null;
  "Hydraulic Pump Type"?: string | null;
  "Hydraulic Flow Rate"?: string | null;
  "Purchased From"?: string | null;
  "Purchased Date"?: string | null;
  "Purchased Price"?: number | null;
  "Title Holder"?: string | null;
  "Sold To"?: string | null;
  "Disposition Date"?: string | null;
  "Sale Price"?: number | null;
  "Insurance Value"?: number | null;
  "CCA Class"?: string | null;
  "Market Value"?: number | null;
  "Rental Number"?: string | null;
  "Start Date"?: string | null;
  "Return Date"?: string | null;
  "Daily Rental"?: number | null;
  "Weekly Rental"?: number | null;
  "Monthly Rental"?: number | null;
  "Equipment Upc"?: string | null;
}
const equiptmentFilterByOptions: IFilterByOptions[] = [
  { "value": "ObjectID", "type": "string" },
  { "value": "EquipmentID", "type": "string" },
  { "value": "IntegrationKey", "type": "string" },
  { "value": "IsInactive", "type": "boolean" },
  { "value": "BusinessUnitUniqueName", "type": "string" },
  { "value": "Description", "type": "string" },
  { "value": "EquipmentTypeID", "type": "string" },
  { "value": "EquipmentTypeIntegrationKey", "type": "string" },
  { "value": "MobilityType", "type": "string" },
  { "value": "OwnershipType", "type": "string" },
  { "value": "OrganizationID", "type": "string" },
  { "value": "OrganizationIntegrationKey", "type": "string" },
  { "value": "OperatorContactName", "type": "string" },
  { "value": "OperatorEmployeeID", "type": "string" },
  { "value": "OperatorEmployeeIntegrationKey", "type": "string" },
  { "value": "SerialNumber", "type": "string" },
  { "value": "Manufacturer", "type": "string" },
  { "value": "Model", "type": "string" },
  { "value": "Year", "type": "number" },
  { "value": "Notes", "type": "string" },
  { "value": "LocationName", "type": "string" },
  { "value": "ExcludeFromFieldLogs", "type": "boolean" },
  { "value": "Length", "type": "number" },
  { "value": "LengthUnitOfMeasure", "type": "string" },
  { "value": "Width", "type": "number" },
  { "value": "WidthUnitOfMeasure", "type": "string" },
  { "value": "Height", "type": "number" },
  { "value": "HeightUnitOfMeasure", "type": "string" },
  { "value": "MaxWeight", "type": "number" },
  { "value": "MaxWeightUnitOfMeasure", "type": "string" },
  { "value": "GroundPressure", "type": "number" },
  { "value": "GroundPressureUnitOfMeasure", "type": "string" },
  { "value": "CombinedWeight", "type": "number" },
  { "value": "CombinedWeightUnitOfMeasure", "type": "string" },
  { "value": "TareWeight", "type": "number" },
  { "value": "TareWeightUnitOfMeasure", "type": "string" },
  { "value": "LicensePlate", "type": "string" },
  { "value": "Color", "type": "string" },
  { "value": "Lojack", "type": "string" },
  { "value": "HUTSticker", "type": "string" },
  { "value": "EZPass", "type": "string" },
  { "value": "ProductionDate", "type": "string" },
  { "value": "Engine", "type": "string" },
  { "value": "EngineArrangement", "type": "string" },
  { "value": "EngineSerialNumber", "type": "string" },
  { "value": "FuelType", "type": "string" },
  { "value": "FuelTankCapacity", "type": "number" },
  { "value": "InitialFuelReading", "type": "number" },
  { "value": "InitialFuelCost", "type": "number" },
  { "value": "FuelUnitOfMeasure", "type": "string" },
  { "value": "InitialFuelDate", "type": "string" },
  { "value": "HorsePower", "type": "string" },
  { "value": "TransmissionModel", "type": "string" },
  { "value": "TransmissionSerialNumber", "type": "string" },
  { "value": "TireSize", "type": "string" },
  { "value": "WheelType", "type": "string" },
  { "value": "TrackType", "type": "string" },
  { "value": "BrakeType", "type": "string" },
  { "value": "CuttingEdge", "type": "string" },
  { "value": "G_E_T", "type": "string" },
  { "value": "HydraulicPumpType", "type": "string" },
  { "value": "HydraulicFlowRate", "type": "string" },
  { "value": "PurchasedFrom", "type": "string" },
  { "value": "PurchasedDate", "type": "string" },
  { "value": "PurchasedPrice", "type": "number" },
  { "value": "TitleHolder", "type": "string" },
  { "value": "SoldTo", "type": "string" },
  { "value": "DispositionDate", "type": "string" },
  { "value": "SalePrice", "type": "number" },
  { "value": "InsuranceValue", "type": "number" },
  { "value": "CCAClass", "type": "string" },
  { "value": "MarketValue", "type": "number" },
  { "value": "RentalNumber", "type": "string" },
  { "value": "StartDate", "type": "string" },
  { "value": "ReturnDate", "type": "string" },
  { "value": "DailyRental", "type": "number" },
  { "value": "WeeklyRental", "type": "number" },
  { "value": "MonthlyRental", "type": "number" },
  { "value": "EquipmentUpc", "type": "string" }
]
const SPREADSHEET_EQUIPMENT_KEYS: Array<keyof IRawEquipment> = [
  "ObjectID",
  "Equipment ID",
  "Integration Key",
  "Is Inactive",
  "Business Unit Unique Name",
  "Description",
  "Equipment Type ID",
  "Equipment Type Integration Key",
  "Mobility Type",
  "Ownership Type",
  "Organization ID",
  "Organization Integration Key",
  "Operator Contact Name",
  "Operator Employee ID",
  "Operator Employee Integration Key",
  "Serial Number",
  "Manufacturer",
  "Model",
  "Year",
  "Notes",
  "Location Name",
  "Exclude From Field Logs",
  "Equipment Parts",
  "Equipment Tags",
  "Length",
  "Length Unit Of Measure",
  "Width",
  "Width Unit Of Measure",
  "Height",
  "Height Unit Of Measure",
  "Max Weight",
  "Max Weight Unit Of Measure",
  "Ground Pressure",
  "Ground Pressure Unit Of Measure",
  "Combined Weight",
  "Combined Weight Unit Of Measure",
  "Tare Weight",
  "Tare Weight Unit Of Measure",
  "License Plate",
  "Color",
  "Lojack",
  "HUT Sticker",
  "EZ Pass",
  "Production Date",
  "Engine",
  "Engine Arrangement",
  "Engine Serial Number",
  "Fuel Type",
  "Fuel Tank Capacity",
  "Initial Fuel Reading",
  "Initial Fuel Cost",
  "Fuel Unit Of Measure",
  "Initial Fuel Date",
  "Horse Power",
  "Transmission Model",
  "Transmission Serial Number",
  "Tire Size",
  "Wheel Type",
  "Track Type",
  "Brake Type",
  "Cutting Edge",
  "G.E.T.",
  "Hydraulic Pump Type",
  "Hydraulic Flow Rate",
  "Purchased From",
  "Purchased Date",
  "Purchased Price",
  "Title Holder",
  "Sold To",
  "Disposition Date",
  "Sale Price",
  "Insurance Value",
  "CCA Class",
  "Market Value",
  "Rental Number",
  "Start Date",
  "Return Date",
  "Daily Rental",
  "Weekly Rental",
  "Monthly Rental",
  "Equipment Upc"
]

interface GetOptions {
  filterQuery: string
}

interface UpdateEquipmentOptions {
  filters?: IFilter[],
  updateEquipmentIDs?: boolean,
  updateIntegrationKeys?: boolean
}

interface IBuildFilterTemplate extends GoogleAppsScript.HTML.HtmlTemplate {
  filterByOptions: IFilterByOptions[]
}

function DisplayEquipmentFilterBuilder() {
  const template = HtmlService.createTemplateFromFile('BuildFilter') as IBuildFilterTemplate
  template.filterByOptions = equiptmentFilterByOptions;
  const html = template.evaluate()
    .setHeight(900)
    .setWidth(1100)
  const ui = SpreadsheetApp.getUi()
  ui.showModalDialog(html, "Build Equipment Filter")
}

function GetEquipment(options: GetOptions) {
  setIsScriptFinished(false);
  clearScriptProgress();
  setCurrentScript("Getting Equipment");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let equipmentSheet = spreadsheet.getSheetByName('Equipment')
  if(!equipmentSheet) {
    equipmentSheet = spreadsheet.insertSheet('Equipment');
    const bold = SpreadsheetApp.newTextStyle().setBold(true).build()
    equipmentSheet.appendRow(SPREADSHEET_EQUIPMENT_KEYS).getRange(1,1,1, SPREADSHEET_EQUIPMENT_KEYS.length).setTextStyle(bold)
  }
  const currentSpreadSheetData = getSpreadSheetData<IRawEquipment>("Equipment")
  const ui = SpreadsheetApp.getUi();
  if(currentSpreadSheetData.length > 0) {
    const response = ui.prompt("The Equipment spreadsheet already has data. This will be overwritten. Do you want to contiune?",
      ui.ButtonSet.YES_NO
    )
    if(response.getSelectedButton() === ui.Button.NO) {
      setIsScriptFinished(true);
      logEvent("Get Equipment Script Canceled")
      return;
    }
  }
  const baseUrl = getBaseURL()
  const token = getOpsToken();
  const headers = createHeaders(token)

  const equiment = getDatabaseItems<IEquipmentDTO>(`${baseUrl}${options.filterQuery}`, {
    method: 'get',
    headers,
    muteHttpExceptions: true
  })
  logEvent(`${equiment.length} pieces of equipment recieved.`)
  
  const rowValues = equiment.map(e => {
    const values = createRawEquipment(e)
    return SPREADSHEET_EQUIPMENT_KEYS.map(key => values[key] ?? "")
  })
  const startRow = equipmentSheet.getLastRow() + 1;

  equipmentSheet.getRange(startRow, 1, rowValues.length, SPREADSHEET_EQUIPMENT_KEYS.length).setValues(rowValues);

  logEvent("Script Complete!")
  ui.alert("Script Complete!")
  setIsScriptFinished(true)
}

function UpdateEquipment(_options: UpdateEquipmentOptions) {
  setIsScriptFinished(false);
  clearScriptProgress()
  setCurrentScript("UpdateEquipment")
  openProgressSidebar("Updating Equipment");

  logEvent("Starting Update Equipment Script")

  const token = getOpsToken();
  const baseUrl = getBaseURL();

  const equipmentData = getSpreadSheetData<IRawEquipment>('Equipment')
  if(!equipmentData || equipmentData.length === 0) {
      SpreadsheetApp.getUi().alert("No data found to send!")
    clearScriptProgress();
    return;
  }
  const headers = createHeaders(token);
  
  const equipmentDtos = equipmentData.map(createEquipmentDTO)
  let payloads: IEquipmentDTO[] = [];
  
  if(!equipmentDtos.every(entry => entry.ObjectID && entry.ObjectID.length > 0)) {
    const getAllURL = baseUrl + "/Equipment"
    const getOptions = {
      headers,
      method: 'get' as const,
      muteHttpExceptions: true
    }
    const existingEquipment = getDatabaseItems<IEquipmentDTO>(getAllURL, getOptions)
    
    payloads = equipmentDtos.map((each, idx) => {
      const equipId = each.EquipmentID
      const intKey = each.IntegrationKey
      
      const existing = existingEquipment.find((e) => {
        const idMatches = equipId === e.EquipmentID
        
        if(e.IntegrationKey) {
          return idMatches && e.IntegrationKey === intKey
        } else {
          return idMatches;
        }
      })
      if(!existing) {
        const errorMessage = `Error: Could not find and existing piece of equiment that matches Equipment ID: ${equipId}${intKey ? ` and Integration Key: ${intKey}`: "."}`
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
    payloads = equipmentDtos
  }
  const batchOptions = payloads.map(body => ({
    url: baseUrl + '/Equipment',
    method: 'put' as const,
    headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  }))
  logEvent(`Updating ${batchOptions.length} pieces of equipment`)
  const failed = [] as number[]
  const responses = batchFetch(batchOptions);
  responses.forEach((res, idx) => {
    const code = res.getResponseCode();
    if(code > 299) {
      writeLogToSpreadsheet(`Error Code: ${res.getResponseCode()}, Message: ${res.getContentText()}`)
      failed.push(idx)
    }
  })
  if(failed.length > 0) {
    const failureMessages = failed.map(idx => `Row ${idx + 2}: ${responses[idx].getContentText()}`)
    logEvent(["Some rows failed!", ...failureMessages])
    highlightRows(failed.map(f => f+2), 'red')
  }
  logEvent("Script Complete!")
  SpreadsheetApp.getUi().alert("Script Complete!")
  setIsScriptFinished(true);
}

function createEquipmentDTO(raw: IRawEquipment): IEquipmentDTO {
  return {
    ObjectID: raw.ObjectID,
    EquipmentID: raw["Equipment ID"],
    IntegrationKey: raw["Integration Key"],
    IsInactive: raw["Is Inactive"],
    BusinessUnitUniqueName: raw["Business Unit Unique Name"],
    Description: raw["Description"],
    EquipmentTypeID: raw["Equipment Type ID"],
    EquipmentTypeIntegrationKey: raw["Equipment Type Integration Key"],
    MobilityType: raw["Mobility Type"],
    OwnershipType: raw["Ownership Type"],
    OrganizationID: raw["Organization ID"],
    OrganizationIntegrationKey: raw["Organization Integration Key"],
    OperatorContactName: raw["Operator Contact Name"],
    OperatorEmployeeID: raw["Operator Employee ID"],
    OperatorEmployeeIntegrationKey: raw["Operator Employee Integration Key"],
    SerialNumber: raw["Serial Number"],
    Manufacturer: raw["Manufacturer"],
    Model: raw["Model"],
    Year: raw["Year"],
    Notes: raw["Notes"],
    LocationName: raw["Location Name"],
    ExcludeFromFieldLogs: raw["Exclude From Field Logs"],
    Length: raw["Length"],
    LengthUnitOfMeasure: raw["Length Unit Of Measure"],
    Width: raw["Width"],
    WidthUnitOfMeasure: raw["Width Unit Of Measure"],
    Height: raw["Height"],
    HeightUnitOfMeasure: raw["Height Unit Of Measure"],
    MaxWeight: raw["Max Weight"],
    MaxWeightUnitOfMeasure: raw["Max Weight Unit Of Measure"],
    GroundPressure: raw["Ground Pressure"],
    GroundPressureUnitOfMeasure: raw["Ground Pressure Unit Of Measure"],
    CombinedWeight: raw["Combined Weight"],
    CombinedWeightUnitOfMeasure: raw["Combined Weight Unit Of Measure"],
    TareWeight: raw["Tare Weight"],
    TareWeightUnitOfMeasure: raw["Tare Weight Unit Of Measure"],
    LicensePlate: raw["License Plate"],
    Color: raw["Color"],
    Lojack: raw["Lojack"],
    HUTSticker: raw["HUT Sticker"],
    EZPass: raw["EZ Pass"],
    ProductionDate: raw["Production Date"],
    Engine: raw["Engine"],
    EngineArrangement: raw["Engine Arrangement"],
    EngineSerialNumber: raw["Engine Serial Number"],
    FuelType: raw["Fuel Type"],
    FuelTankCapacity: raw["Fuel Tank Capacity"],
    InitialFuelReading: raw["Initial Fuel Reading"],
    InitialFuelCost: raw["Initial Fuel Cost"],
    FuelUnitOfMeasure: raw["Fuel Unit Of Measure"],
    InitialFuelDate: raw["Initial Fuel Date"],
    HorsePower: raw["Horse Power"],
    TransmissionModel: raw["Transmission Model"],
    TransmissionSerialNumber: raw["Transmission Serial Number"],
    TireSize: raw["Tire Size"],
    WheelType: raw["Wheel Type"],
    TrackType: raw["Track Type"],
    BrakeType: raw["Brake Type"],
    CuttingEdge: raw["Cutting Edge"],
    G_E_T: raw["G.E.T."],
    HydraulicPumpType: raw["Hydraulic Pump Type"],
    HydraulicFlowRate: raw["Hydraulic Flow Rate"],
    PurchasedFrom: raw["Purchased From"],
    PurchasedDate: raw["Purchased Date"],
    PurchasedPrice: raw["Purchased Price"],
    TitleHolder: raw["Title Holder"],
    SoldTo: raw["Sold To"],
    DispositionDate: raw["Disposition Date"],
    SalePrice: raw["Sale Price"],
    InsuranceValue: raw["Insurance Value"],
    CCAClass: raw["CCA Class"],
    MarketValue: raw["Market Value"],
    RentalNumber: raw["Rental Number"],
    StartDate: raw["Start Date"],
    ReturnDate: raw["Return Date"],
    DailyRental: raw["Daily Rental"],
    WeeklyRental: raw["Weekly Rental"],
    MonthlyRental: raw["Monthly Rental"],
    EquipmentUpc: raw["Equipment Upc"]
  };
}
function createRawEquipment(dto: IEquipmentDTO): IRawEquipment {
  return {
    "ObjectID": dto.ObjectID,
    "Equipment ID": dto.EquipmentID,
    "Integration Key": dto.IntegrationKey,
    "Is Inactive": dto.IsInactive,
    "Business Unit Unique Name": dto.BusinessUnitUniqueName,
    "Description": dto.Description,
    "Equipment Type ID": dto.EquipmentTypeID,
    "Equipment Type Integration Key": dto.EquipmentTypeIntegrationKey,
    "Mobility Type": dto.MobilityType,
    "Ownership Type": dto.OwnershipType,
    "Organization ID": dto.OrganizationID,
    "Organization Integration Key": dto.OrganizationIntegrationKey,
    "Operator Contact Name": dto.OperatorContactName,
    "Operator Employee ID": dto.OperatorEmployeeID,
    "Operator Employee Integration Key": dto.OperatorEmployeeIntegrationKey,
    "Serial Number": dto.SerialNumber,
    "Manufacturer": dto.Manufacturer,
    "Model": dto.Model,
    "Year": dto.Year,
    "Notes": dto.Notes,
    "Location Name": dto.LocationName,
    "Exclude From Field Logs": dto.ExcludeFromFieldLogs,
    "Length": dto.Length,
    "Length Unit Of Measure": dto.LengthUnitOfMeasure,
    "Width": dto.Width,
    "Width Unit Of Measure": dto.WidthUnitOfMeasure,
    "Height": dto.Height,
    "Height Unit Of Measure": dto.HeightUnitOfMeasure,
    "Max Weight": dto.MaxWeight,
    "Max Weight Unit Of Measure": dto.MaxWeightUnitOfMeasure,
    "Ground Pressure": dto.GroundPressure,
    "Ground Pressure Unit Of Measure": dto.GroundPressureUnitOfMeasure,
    "Combined Weight": dto.CombinedWeight,
    "Combined Weight Unit Of Measure": dto.CombinedWeightUnitOfMeasure,
    "Tare Weight": dto.TareWeight,
    "Tare Weight Unit Of Measure": dto.TareWeightUnitOfMeasure,
    "License Plate": dto.LicensePlate,
    "Color": dto.Color,
    "Lojack": dto.Lojack,
    "HUT Sticker": dto.HUTSticker,
    "EZ Pass": dto.EZPass,
    "Production Date": dto.ProductionDate,
    "Engine": dto.Engine,
    "Engine Arrangement": dto.EngineArrangement,
    "Engine Serial Number": dto.EngineSerialNumber,
    "Fuel Type": dto.FuelType,
    "Fuel Tank Capacity": dto.FuelTankCapacity,
    "Initial Fuel Reading": dto.InitialFuelReading,
    "Initial Fuel Cost": dto.InitialFuelCost,
    "Fuel Unit Of Measure": dto.FuelUnitOfMeasure,
    "Initial Fuel Date": dto.InitialFuelDate,
    "Horse Power": dto.HorsePower,
    "Transmission Model": dto.TransmissionModel,
    "Transmission Serial Number": dto.TransmissionSerialNumber,
    "Tire Size": dto.TireSize,
    "Wheel Type": dto.WheelType,
    "Track Type": dto.TrackType,
    "Brake Type": dto.BrakeType,
    "Cutting Edge": dto.CuttingEdge,
    "G.E.T.": dto.G_E_T,
    "Hydraulic Pump Type": dto.HydraulicPumpType,
    "Hydraulic Flow Rate": dto.HydraulicFlowRate,
    "Purchased From": dto.PurchasedFrom,
    "Purchased Date": dto.PurchasedDate,
    "Purchased Price": dto.PurchasedPrice,
    "Title Holder": dto.TitleHolder,
    "Sold To": dto.SoldTo,
    "Disposition Date": dto.DispositionDate,
    "Sale Price": dto.SalePrice,
    "Insurance Value": dto.InsuranceValue,
    "CCA Class": dto.CCAClass,
    "Market Value": dto.MarketValue,
    "Rental Number": dto.RentalNumber,
    "Start Date": dto.StartDate,
    "Return Date": dto.ReturnDate,
    "Daily Rental": dto.DailyRental,
    "Weekly Rental": dto.WeeklyRental,
    "Monthly Rental": dto.MonthlyRental,
    "Equipment Upc": dto.EquipmentUpc
  };
}
