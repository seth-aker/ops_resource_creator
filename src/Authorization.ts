interface IOpsVars extends Record<string,string> {
  opsBaseUrl: string,
  clientId: string,
  opsClientSecret: string
}
interface ISetOpsTemplate extends GoogleAppsScript.HTML.HtmlTemplate {
  baseUrl: string,
  clientID: string,
  hasClientSecret: boolean
}
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OPS Properties')
      .addItem('Set OPS Properties', 'setOpsProperties')
      .addItem('View Current OPS Properties', 'viewOpsProperties')
      .addItem('Clear OPS Properties', 'clearOpsProperties')
      .addToUi()

  SpreadsheetApp.getUi()
    .createMenu('Run Scripts')
      .addItem("Create Users", "CreateUsers")
      .addToUi()

}
function setOpsProperties() {
  const template = HtmlService.createTemplateFromFile('SetOpsProperties') as ISetOpsTemplate
  const userProperties = PropertiesService.getUserProperties();
  const props = userProperties.getProperties() as IOpsVars
  template.baseUrl = props.opsBaseUrl
  template.clientID = props.clientId
  template.hasClientSecret = props.opsClientSecret ? true : false
  SpreadsheetApp.getUi().showModalDialog(template.evaluate(), "Set Ops Environment Variables")
}
function clearOpsProperties() {
  PropertiesService.getUserProperties().deleteAllProperties()
  SpreadsheetApp.getUi().alert("Database properties successfully deleted")
}
function viewOpsProperties() {
  const props = PropertiesService.getUserProperties().getProperties() as IOpsVars
  SpreadsheetApp.getUi().alert(`Current API Properties: \nBase URL: ${props.opsBaseUrl}\nClientID: ${props.clientId}`)
}
function getOpsToken() {
  const baseUrl = PropertiesService.getUserProperties().getProperty('opsBaseUrl')
  const clientId = PropertiesService.getUserProperties().getProperty('clientId')
  const clientSecret = PropertiesService.getUserProperties().getProperty('opsClientSecret')

  if(!baseUrl) {
    throw new Error("Missing Ops API Url!")
  }
  if(!clientId) {
    throw new Error("Missing Ops Client Id!")
  }
  if(!clientSecret) {
    throw new Error("Missing Client Secret!")
  }
  // this will throw an error, we want this to crash the program and send the message to the user.
  const res = UrlFetchApp.fetch(`${baseUrl}/login`, {
    headers: {
      clientId,
      clientSecret
    },
  })
  const payload = JSON.parse(res.getContentText()) as {AccessToken: string, RefreshToken: string}
  return payload.AccessToken;
}
function setOpsVariables(vars: {baseUrl: string, clientSecret: string, clientID: string} & Record<string,string>) {
  for(const [key, val] of Object.entries(vars)) {
    vars[key] = val.trim()
  }
  const userProperties = PropertiesService.getUserProperties();
  if(vars.clientSecret === "********") {
    vars.clientSecret = userProperties.getProperty('opsClientSecret') ?? ""
  }
  userProperties.setProperties({
    opsBaseUrl: vars.baseUrl,
    opsClientSecret: vars.clientSecret,
    clientId: vars.clientID
  })
}

// Simple check to make sure that an access token is successfully granted. 
// Will throw and get caught by the handleError in SetUserProperties.html
function validateOpsAuthentication() {
  getOpsToken();
}