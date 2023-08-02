//OAuth 2 lib = 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
const DOMAIN = "-api.toolsquare.io";

const OauthURL = "/o/authorize/";
const OauthTokenURL = "/o/token/";
const USERS_GET_URL = '/open-api/v1/users/user/';
const USERS_CREATE_URL = '/api/user_core/user/bulkcreate/';
const USER_GROUPS_URL = '/open-api/v1/groups/usergroup/';
const USERS_UPDATE_URL = '/open-api/v1/users/user/bulkupdate/';
const USERS_DESTROY_URL = '/open-api/v1/users/user/bulkdestroy/';

let theProps = PropertiesService.getDocumentProperties();

let NrOfNewUsers = 0;
let NrOfCalls = 0;
const PageSize = 100;

//UI
function onOpen() {
  //Add functions to menu
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Toolsquare')
    //.addItem('Create User Template', 'CreateUserRegistrationSheet')
    .addItem('Create User Template', 'createUserTemplate')
    .addItem('Create Users', 'createUsers')
    .addItem('Delete all credentials', 'deleteAll')
    .addToUi()
}

function getUrl(slug) {
  let tenant = theProps.getProperty('CLIENT_TENANT');
  return "https://" + tenant + DOMAIN + slug;
}

function getUsers() {
  let pageIndex = 1;
  let nextResults = true;
  var users = [];

  while (nextResults) {

    Logger.log("# retrieving users page " + pageIndex);
    let results = getTSData(USERS_GET_URL, pageIndex).results;

    if (results != null) {
      pageIndex++;
      for (var n = 0; n < results.length; n++) {
        users.push(results[n]);
      }
    } else {
      Logger.log("# users = " + users.length);
      nextResults = false;
    }
  }
  return users;
}

function getUserGroups() {
  let usergroups = getTSData(USER_GROUPS_URL, 1);
  Logger.log(JSON.stringify(usergroups, null, 2));
  return usergroups.results;
}

function createUsers() {

  let sheet = null;
  let users = [];
  let usersWithgroups = [];

  doc = SpreadsheetApp.getActiveSheet();
  sheet = SpreadsheetApp.getActiveSheet();

  let lastRow = sheet.getLastRow();
  let userRange = "A2:D" + lastRow;
  let userData = sheet.getRange(userRange).getValues();

  for (let i = 0; i < userData.length; ++i) {
    let u_object = {};
    let ug_object = {};
    let fName = userData[i][0];
    let lName = userData[i][1];
    let email = userData[i][2];
    let userGroup = userData[i][3];

    var re = /\S+@\S+\.\S+/;
    if (re.test(email)) {
      u_object.first_name = fName;
      u_object.last_name = lName;
      u_object.email = email;

      ug_object.first_name = fName;
      ug_object.last_name = lName;
      ug_object.email = email;
      ug_object.usergroup = userGroup;

      users.push(u_object);
      usersWithgroups.push(ug_object)
    }
  }

  Logger.log(JSON.stringify(users));
  let result = apiTSData(USERS_CREATE_URL, users, 'post');
  let responsecode = result.getResponseCode().toString();
  Logger.log("Users created response: " + responsecode);
  let responsetext = JSON.parse(result.getContentText());

  NrOfNewUsers = responsetext.data.length;
  Logger.log(NrOfNewUsers);

  updateUserData(usersWithgroups);

}

function updateUserData(theUpdateList, allUsers = null) {
  if (allUsers == null) {
    allUsers = getUsers();
  }
  let usergroups = getUserGroups();
  let changelist = [];

  for (e in theUpdateList) {
    let updateemail = theUpdateList[e].email;
    updateemail = updateemail.toLowerCase();
    Logger.log(updateemail);
    for (i in allUsers) {
      if (updateemail == allUsers[i].email) {

        let object = {};
        object.first_name = theUpdateList[e].first_name;
        object.last_name = theUpdateList[e].last_name;
        object.id = allUsers[i].id;
        object.usergroup = getUserGroupID(usergroups, theUpdateList[e].usergroup);
        changelist.push(object);

      }
    }
  }
  Logger.log(changelist);
  let result = apiTSData(USERS_UPDATE_URL, changelist, "PATCH");
  let responsecode = result.getResponseCode().toString();
  Logger.log("Users updated response: " + responsecode);
  let responsetext = JSON.parse(result.getContentText());

  let NrOfUsersUpdated = responsetext.data.length;
  Logger.log(NrOfUsersUpdated);

  let ui = SpreadsheetApp.getUi();
  ui.alert("There are " + NrOfNewUsers + " created and " + NrOfUsersUpdated + " updated");
}

function getUserGroupID(theusergroups, thegroupname) {
  for (e in theusergroups) {
    if (theusergroups[e].name === thegroupname) {
      return (theusergroups[e].id);
    }
  }
}

function getService() {
  //get check credentials
  let C_TENANT = theProps.getProperty('CLIENT_TENANT');
  if (C_TENANT === null) {
    setKey_('CLIENT_TENANT', 'tenant');
    C_TENANT = theProps.getProperty('CLIENT_TENANT');
  }

  let C_ID = theProps.getProperty('CLIENT_KEY');
  if (C_ID === null) {
    setKey_('CLIENT_KEY', 'client id');
    C_ID = theProps.getProperty('CLIENT_KEY');
  }

  let C_SECRET = theProps.getProperty('CLIENT_SECRET');
  if (C_SECRET === null) {
    setKey_('CLIENT_SECRET', 'client secret');
    C_SECRET = theProps.getProperty('CLIENT_SECRET');
  }

  let authURL = getUrl(OauthURL);
  let tokenURL = getUrl(OauthTokenURL);

  return OAuth2.createService('TSUsers')
    .setAuthorizationBaseUrl(authURL)
    .setTokenUrl(tokenURL)
    .setClientId(C_ID)
    .setClientSecret(C_SECRET)
    .setCallbackFunction('oauthCallback')
    .setPropertyStore(theProps)
    .setScope('read')
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(C_ID + ':' + C_SECRET),
      'Content-Type': 'application/x-www-form-urlencoded'
    })
    .setCache(CacheService.getUserCache());
}

function getTSData(theURL, pageNr = 1) {
  NrOfCalls++;
  let service = getService();
  let isConnected = checkAccess(service);

  //Access granted
  if (isConnected) {
    let url = getUrl(theURL) + '?page=' + pageNr + '&page_size=' + PageSize;
    let response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      },
      muteHttpExceptions: true
    });
    let result = JSON.parse(response.getContentText());
    return result;
  }
}

function apiTSData(theURL, theContent, method) {
  NrOfCalls++;
  let service = getService();
  let isConnected = checkAccess(service);

  //Access granted
  if (isConnected) {
    Logger.log(theURL + " " + method);
    // let tenant = PropertiesService.getUserProperties('CLIENT_TENANT');
    let url = getUrl(theURL);

    let response = UrlFetchApp.fetch(url, {
      method: method,
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      },
      contentType: 'application/json',
      payload: JSON.stringify(theContent),
      muteHttpExceptions: false
    });

    Logger.log(response);
    return response;
  }
  return "is not connected";
}

function checkAccess(theService) {
  //Access granted
  if (theService.hasAccess()) {
    return true;
    //Access not yet granted
  } else {
    Logger.log("OAuth connection error!");
    return resetAccess(theService);
  }
}

function oauthCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

function resetAccess(theService) {
  var authorizationUrl = theService.getAuthorizationUrl();
  var htmlOutput = HtmlService
    .createHtmlOutput('<p style="font-family:verdana"> Before you\'ll be able get date form the platform, you\'ll have to open <a href="' + authorizationUrl + '">this link</a>. <br><br> You should login into the Toolsquare platform and rerun the request.<br><br>The red bar on top should disapear on the next run</p>')
    .setWidth(400) //optional
    .setHeight(200); //optional
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Authorization request');

  //Logger.log('Open the following URL and re-run the script: %s',authorizationUrl);
  return false;
}

function createUserTemplate() {

  doc = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = doc.insertSheet();

  let user_range = "A1:E10"; //Increase if needed
  let rowlength = 10
  let userdatalist = new Array(rowlength);

  for (var i = 0; i < userdatalist.length; i++) {
    userdatalist[i] = new Array(5);
  }

  //Table Headers
  userdatalist[0][0] = "First Name";
  userdatalist[0][1] = "Last Name";
  userdatalist[0][2] = "Email";
  userdatalist[0][3] = "User Group";

  sheet.getRange(user_range).setValues(userdatalist);

  setusergroupvalidation(sheet, rowlength);

}

function setusergroupvalidation(theSheet, theLength) {
  let usergroupsResult = getTSData(USER_GROUPS_URL);
  let usergroups = usergroupsResult.results;

  let grouplist = [];

  for (e in usergroups) {
    grouplist.push(usergroups[e].name);
  }

  let range = "D2:D" + theLength;
  let usergrouprange = theSheet.getRange(range);
  let groupRule = SpreadsheetApp.newDataValidation().requireValueInList(grouplist);
  usergrouprange.setDataValidation(groupRule);
}

/**
 * Logs the redict URI to register.
 */
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}

function setKey_(KEY, Promt) {
  let ui = SpreadsheetApp.getUi();
  var scriptValue = ui.prompt('Please provide your ' + Promt, ui.ButtonSet.OK);
  let value = scriptValue.getResponseText();
  if (value.length > 1) {
    theProps.setProperty(KEY, value);
  }
}

function deleteAll() {
  theProps.deleteAllProperties();
  CacheService.getDocumentCache();
}
