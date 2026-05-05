interface IUserVars extends Record<string,string> {
  baseUrl: string,
  clientId: string,
  clientSecret: string
}
function getToken() {
  const baseUrl = PropertiesService.getUserProperties().getProperty('baseUrl')
  const clientId = PropertiesService.getUserProperties().getProperty('clientId')
  const clientSecret = PropertiesService.getUserProperties().getProperty('clientSecret')

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
function setUserVariables(vars: IUserVars) {
  for(const [key, val] of Object.entries(vars)) {
    vars[key] = val.trim()
  }
  const userProperties = PropertiesService.getUserProperties();
  if(vars.clientSecret === "********") {
    vars.clientSecret = userProperties.getProperty('clientSecret') ?? ""
  }
  userProperties.setProperties(vars)
}

// Simple check to make sure that an access token is successfully granted. Will throw and get caught by the handleError in SetUserProperties.html
function validateAuthentication() {
  getToken();
}