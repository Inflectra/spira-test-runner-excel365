/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 *
 */
export {
  clearAll,
  error,
  getBespoke,
  getCustoms,
  getFromSpiraExcel,
  getFromSheetExcel,
  getProjects,
  getReleases,
  getTemplateFromProjectId,
  operationComplete,
  sendToSpira,
  templateLoader,
  warn
};

import { showPanel, hidePanel } from './taskpane.js';

import { params } from './model.js';

// globals
var API_PROJECT_BASE = '/services/v6_0/RestService.svc/projects/',
  API_URL_BASE = '/services/v6_0/RestService.svc/',
  API_PROJECT_BASE_NO_SLASH = '/services/v6_0/RestService.svc/projects',
  API_TEMPLATE_BASE = '/services/v6_0/RestService.svc/project-templates/',
  ART_ENUMS = {
    requirements: 1,
    testCases: 2,
    incidents: 3,
    releases: 4,
    testRuns: 5,
    tasks: 6,
    testSteps: 7,
    testSets: 8,
    risks: 14,
  },
  INITIAL_HIERARCHY_OUTDENT = -20,
  GET_PAGINATION_SIZE = 100,
  EXCEL_MAX_ROWS = 10000,
  FIELD_MANAGEMENT_ENUMS = {
    all: 1,
    standard: 2,
    subType: 3
  },
  STATUS_ENUM = {
    allSuccess: 1,
    someError: 2,
    allError: 3,
    wrongSheet: 4,
    existingEntries: 5,
    noData: 6,
    preCheckingError: 7,
  },
  SUBTYPE_IDS = ["TestCaseId", "TestStepId"],
  STATUS_MESSAGE_GOOGLE = {
    1: "All done! To send more data over, clear the sheet first.",
    2: "Sorry, but there were some problems (see the cells marked in red). Check any notes on the relevant ID field for explanations.",
    3: "We're really sorry, but we couldn't send anything to SpiraPlan - please check notes on the ID fields for more information.",
    4: "You are not on the correct worksheet. Please go to the sheet that matches the one listed on the Spira taskpane / the selection you made in the sidebar.",
    5: "Some/all of the rows already exist in SpiraPlan. These rows have not been re-added."
  },
  STATUS_MESSAGE_EXCEL = {
    1: "All done! To send more data over, clear the sheet first.",
    2: "Sorry, but there were some problems (see the cells marked in red). Check any notes on the relevant ID field for explanations.",
    3: "We're really sorry, but we couldn't send anything to SpiraPlan - please check notes on the ID fields for more information.",
    4: "You are not on the correct worksheet. Please go to the sheet that matches the one listed on the Spira taskpane / the selection you made in the sidebar.",
    5: "Some/all of the rows already exist in SpiraPlan. These rows have not been re-added.",
    6: "It seems you are not the owner of any Test Case/Test Set in this product or all the artifacts assigned to you have the status 'Passed'. Please select another product or check the current one in Spira and try again.",
    7: "The entered data is missing required fields and/or has invalid execution statuses. Please check notes on the ID fields for more information or refer to the <a href=\"" + params.documentationURL + "\">documentation</a>."
  },
  CUSTOM_PROP_TYPE_ENUM = {
    1: "StringValue",
    2: "IntegerValue",
    3: "DecimalValue",
    4: "BooleanValue",
    5: "DateTimeValue",
    6: "IntegerValue",
    7: "IntegerListValue",
    8: "IntegerValue"
  },
  INLINE_STYLING = "style='font-family: sans-serif'",
  ART_PARENT_IDS = {
    2: 'TestCaseId',
    7: 'TestCaseId'
  },
  TC_ID_COLUMN_INDEX = 1,
  TS_ID_COLUMN_INDEX = 2,
  TX_ID_COLUMN_INDEX = 3;

/*
 * ======================
 * INITIAL LOAD FUNCTIONS
 * ======================
 *
 * These functions are needed for initialization
 * All Google App Script (GAS) files are bundled by the engine
 * at start up so any non-scoped variables declared will be available globally.
 *
 */

// Google App script boilerplate install function
// opens app on install
function onInstall(e) {
  onOpen(e);
}

// App script boilerplate open function
// opens sidebar
// Method `addItem`  is related to the 'Add-on' menu items. Currently just one is listed 'Start' in the dropdown menu
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Start', 'showSidebar').addToUi();
}

// side bar function gets index.html and opens in side window
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('SpiraPlan by Inflectra');

  SpreadsheetApp.getUi().showSidebar(ui);
}



// This function is part of the google template engine and allows for modularization of code
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}









/*
 *
 * ========================
 * TEMPLATE PANEL FUNCTIONS
 * ========================
 *
 */

// copy the first sheet into a new sheet in the same spreadsheet
function save() {
  // pop up telling the user that their data will be saved
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('This will save the current sheet in a new sheet on this spreadsheet. Continue?', ui.ButtonSet.YES_NO);

  // returns with user choice
  if (response == ui.Button.YES) {
    // get first sheet of  active spreadsheet
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadSheet.getActiveSheet();

    // get entire open spreadsheet id
    var id = spreadSheet.getId();

    // set current spreadsheet file as destination
    var destination = SpreadsheetApp.openById(id);

    // copy sheet to current spreadsheet in new sheet
    sheet.copyTo(destination);

    // returns true to queue success popup
    return true;
  } else {
    // returns false to ignore success popup
    return false;
  }
}

//Function to check if a given sheet name exists or not in the workbook
async function checkSheetExists(sheetName) {
  Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
      .then(function () {
        sheets.items.forEach(function (sheet) {
          if (sheet.name == sheetName) {
            return true;
          }
        });
      });
  });
  return false;
}

//clears active sheet in spreadsheet
function clearAll() {
  return Excel.run(context => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    // for excel we do not reset the sheet name because this can cause timing problems on some versions of Excel
    sheet.getRange().clear();

    //check if the database sheet exists
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    var isDatabaseSheet = false;

    return context.sync()
      .then(function () {
        sheets.items.forEach(function (singleSheet) {
          if (singleSheet.name == params.dataSheetName) {
            isDatabaseSheet = true;
          }
        });

        if (!isDatabaseSheet) {
          //if we don't have a database worksheet, create one
          var dbSheet = sheets.add(params.dataSheetName);
          dbSheet.visibility = Excel.SheetVisibility.hidden;
        }
        else {
          //if we have a database worksheet, clear it
          var worksheet = context.workbook.worksheets.getItemOrNullObject(params.dataSheetName);
          worksheet.getRange().clear();
        }
        return context.sync();
      });
  })
}


// handles showing popup messages to user
// @param: message - strng of the raw message to show user
// @param: messageTitle - strng of the message title to use
// @param: isTemplateLoadFail - bool about whether this message means that the template load sequence has failed
function popupShow(message, messageTitle, isTemplateLoadFail) {
  if (!message) return;
  showPanel("confirm");
  document.getElementById("message-confirm").innerHTML = (messageTitle ? "<b>" + messageTitle + ":</b> " : "") + message;
  document.getElementById("btn-confirm-cancel").style.visibility = "hidden";
  document.getElementById("btn-confirm-ok").onclick = function () { popupHide() };
  return !isTemplateLoadFail ? null : {
    isTemplateLoadFail: isTemplateLoadFail,
    message: message
  };
}

function popupHide() {
  hidePanel("confirm");
  document.getElementById("message-confirm").innerHTML = "";
  document.getElementById("btn-confirm-cancel").style.visibility = "visible";
}







/*
 *
 * ====================
 * DATA "GET" FUNCTIONS
 * ====================
 *
 * functions used to retrieve data from Spira - things like projects and users, not specific records
 *
 */

// General fetch function, using Google's built in fetch api
// @param: currentUser - user object storing login data from client
// @param: fetcherUrl - url string passed in to connect with Spira
function fetcher(currentUser, fetcherURL) {
  //use google's Utilities to base64 decode if present, otherwise use standard JS (ie for MS Excel)
  var decoded = typeof Utilities != "undefined" ? Utilities.base64Decode(currentUser.api_key) : atob(currentUser.api_key);
  var APIKEY = typeof Utilities != "undefined" ? Utilities.newBlob(decoded).getDataAsString() : decoded;

  //build URL from args
  var fullUrl = currentUser.url + fetcherURL + "username=" + currentUser.userName + APIKEY;
  //set MIME type
  var params = { "Content-Type": "application/json", "accepts": "application/json" };

  return superagent
    .get(fullUrl)
    .set("Content-Type", "application/json", "accepts", "application/json")

}



// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
// @param: currentUser - object with details about the current user
function getProjects(currentUser) {
  var fetcherURL = API_PROJECT_BASE_NO_SLASH + '?';
  return fetcher(currentUser, fetcherURL);
}



// Gets projects accessible by current logged in user
// This function is called on initial log in and therefore also acts as user validation
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getTemplateFromProjectId(currentUser, projectId) {
  var fetcherURL = API_PROJECT_BASE + projectId + '?';
  return fetcher(currentUser, fetcherURL);
}



// Gets custom fields for selected project and artifact
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
// @param: artifactName - int of the current artifact - API refers to this as the artifactTypeName but the ID is required
function getCustoms(currentUser, templateId, artifactName) {
  var fetcherURL = API_TEMPLATE_BASE + templateId + '/custom-properties/' + artifactName + '?';
  return fetcher(currentUser, fetcherURL);
}



// Gets data for a bespoke specified field (for selected project and artifact)
// @param: currentUser - object with details about the current user
// @param: templateId - int id for current template
// @param: projectId - int id for current project
// @param: artifactName - string name of the current artifact
// @param: field - object of the field from the model
function getBespoke(currentUser, templateId, projectId, artifactName, field) {
  var fetcherURL = "";
  // a couple of dynamic fields are project based - like folders
  if (field.bespoke.isProjectBased) {
    fetcherURL = API_PROJECT_BASE + projectId + field.bespoke.url + '?';
  } else {
    fetcherURL = API_TEMPLATE_BASE + templateId + field.bespoke.url + '?';
  }
  var results = fetcher(currentUser, fetcherURL);
  return results;
}



// Gets releases for selected project
// @param: currentUser - object with details about the current user
// @param: projectId - int id for current project
function getReleases(currentUser, projectId) {
  var fetcherURL = API_PROJECT_BASE + projectId + '/releases?';
  return fetcher(currentUser, fetcherURL);
}


function getArtifacts(user, projectId, artifactTypeId, startRow, numberOfRows, artifactId, parentTypeId) {
  var fullURL = '';
  var body = '';
  var response = null;

  switch (artifactTypeId) {
    case ART_ENUMS.testRuns:
      fullURL += API_URL_BASE + "test-cases?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=TestCaseId&sort_direction=ASC&";
      response = fetcher(user, fullURL);
      break;
    case ART_ENUMS.testSteps:
      if (artifactId) {
        if (parentTypeId == ART_ENUMS.testSets) {
          fullURL += API_PROJECT_BASE + projectId + '/test-runs/create/test_set/' + artifactId + '?';
          body = "[" + artifactId + "]";
          response = poster(body, user, fullURL);
        }
        else if (parentTypeId == ART_ENUMS.testCases) {
          fullURL += API_PROJECT_BASE + projectId + '/test-runs/create' + '?';
          body = "[" + artifactId + "]";
          response = poster(body, user, fullURL);
        }
      }
      break;
    case ART_ENUMS.testSets:
      if (artifactId == null) {
        fullURL += API_URL_BASE + "/test-sets?starting_row=" + startRow + "&number_of_rows=" + numberOfRows + "&sort_field=TestSetId&sort_direction=ASC&";
        response = fetcher(user, fullURL);
      }
      else {
        fullURL += API_PROJECT_BASE + projectId + "/test-sets/" + artifactId + "/test-case-mapping?&";
        response = fetcher(user, fullURL);
      }
      break;
    case ART_ENUMS.testCases:
      fullURL += API_PROJECT_BASE + projectId + "/test-cases/" + artifactId + "?&";
      response = fetcher(user, fullURL);
      break;
  }
  return response;
}




/*
 *
 * =======================
 * CREATE "POST" FUNCTIONS
 * =======================
 *
 * functions to create new records in Spira - eg add new requirements
 *
 */

// General fetch function, using Google's built in fetch api
// @param: body - json object
// @param: currentUser - user object storing login data from client
// @param: postUrl - url string passed in to connect with Spira
function poster(body, currentUser, postUrl) {

  //use google's Utilities to base64 decode if present, otherwise use standard JS (ie for MS Excel)
  var decoded = typeof Utilities != "undefined" ? Utilities.base64Decode(currentUser.api_key) : atob(currentUser.api_key);
  var APIKEY = typeof Utilities != "undefined" ? Utilities.newBlob(decoded).getDataAsString() : decoded;

  //build URL from args
  var fullUrl = currentUser.url + postUrl + "username=" + currentUser.userName + APIKEY;

  //POST headers
  var params = {};
  params.method = 'post';
  params.contentType = 'application/json';
  params.muteHttpExceptions = true;
  if (body) params.payload = body;

  //for MS Excel, use superagent to return a promise to the taskpane
  return superagent
    .post(fullUrl)
    .send(body)
    .set("Content-Type", "application/json", "accepts", "application/json");
}

// General fetch function, using Google's built in fetch api
// @param: body - json object
// @param: currentUser - user object storing login data from client
// @param: PUTUrl - url string passed in to connect with Spira
function putUpdater(body, currentUser, PUTUrl) {
  //use google's Utilities to base64 decode if present, otherwise use standard JS (ie for MS Excel)
  var decoded = typeof Utilities != "undefined" ? Utilities.base64Decode(currentUser.api_key) : atob(currentUser.api_key);
  var APIKEY = typeof Utilities != "undefined" ? Utilities.newBlob(decoded).getDataAsString() : decoded;

  //build URL from args
  var fullUrl = currentUser.url + PUTUrl + "username=" + currentUser.userName + APIKEY;

  //PUT headers
  var params = {};
  params.method = 'put';
  params.contentType = 'application/json';
  params.muteHttpExceptions = true;
  if (body) params.payload = body;

  //for MS Excel, use superagent to return a promise to the taskpane
  var putResult =
    superagent
      .put(fullUrl)
      .send(body)
      .set("Content-Type", "application/json", "accepts", "application/json");
  return putResult;
}


// effectively a switch to manage which artifact we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactId - int of the current artifact
// @param: parentId - optional int of the relevant parent to attach the artifact too
function putArtifactToSpira(entry, user, projectId, artifactTypeId, parentId) {

  var response = "",
    putUrl = "",
    shellUrl = "",
    shellBody = "",
    shellResponse = "";



  //1. Create Test Run Shell, depending on type

  if (entry[params.secondaryShellField]) {
    //Test Set Shell
    shellUrl = API_PROJECT_BASE + projectId + '/test-runs/create/test_set/' + entry[params.specialFields.secondaryShellField] + '?';
    shellBody = "[" + entry[params.specialFields.specialFields.secondaryShellField] + "]";
    shellResponse = poster(shellBody, user, shellUrl);
  }
  else {
    //Test Case Shell
    shellUrl = API_PROJECT_BASE + projectId + '/test-runs/create' + '?';
    shellBody = "[" + entry[params.specialFields.standardShellField] + "]";
    shellResponse = poster(shellBody, user, shellUrl);
  }

  //2. Put the TesRun
  var JSON_body = "[" + JSON.stringify(entry) + "]",
    putUrl = API_PROJECT_BASE + projectId + '/test-runs' + '?';
  var response = putUpdater(JSON_body, user, putUrl);

  return response;
}

// effectively a switch to manage which artifact we have and therefore which API call to use with what data
// returns the response from the specific post service to Spira
// @param: entry - object of single specific entry to send to Spira
// @param: user - user object
// @param: projectId - int of the current project
// @param: artifactId - int of the current artifact
function postArtifactToSpira(entry, user, projectId, artifactTypeId) {
  //stringify
  var JSON_body = JSON.stringify(entry),
    response = "",
    postUrl = "";

  //send JSON object of new item to artifact specific export function
  switch (artifactTypeId) {
    // INCIDENTS
    case ART_ENUMS.incidents:
      postUrl = API_PROJECT_BASE + projectId + '/incidents?';
      break;
  }
  return postUrl ? poster(JSON_body, user, postUrl) : null;
}

/*
 *
 * ==============
 * ERROR MESSAGES
 * ==============
 *
 */

// Error notification function
// Assigns string value and routes error call from client.js.html
// @param: type - string identifying the message to be displayed
// @param: err - the detailed error object (differs between plugin)
function error(type, err) {
  var message = "",
    details = "";
  if (type == 'impExp') {
    message = 'There was an input error. Please check your network connection and make sure your Spira user can view Test Sets, Test Cases, and Test Steps, as well as create Test Runs and Incidents. Also, make sure you have Test Cases and/or Test Sets assigned to you in this product. For further  information, please refer to the <a href=\"' + params.documentationURL + '\">documentation</a>.';
  } else if (type == "network") {
    message = 'Network error. Please check your username, url, and password. If correct make sure you have the correct permissions.';
    details = err ? `<br><br><b>STATUS:</b> ${err.status ? err.status : "unknown"}<br><br><b>MESSAGE:</b> ${err.stack ? err.stack : "unknown"}` : "";
  } else if (type == 'excel') {
    message = 'Excel reported an error!';
    details = err ? `<br><br>Description: ${err.description}` : "";
  } else if (type == 'unknown' || err == 'unknown') {
    message = 'Unkown error. Please try again later or contact your system administrator';
  } else if (type == 'sheet') {
    message = 'There was a problem while retrieving data from the active spreadsheet. Please check the details below and try again. <br><br><b>Details:</b><br>' + err;
  }
  else {
    message = 'Unkown error. Please try again later or contact your system administrator';
  }

  popupShow(message + details, "");
}



// Pop-up notification function
// @param: string - message to be displayed
function success(string) {
  // Show a 2-second popup with the title "Status" and a message passed in as an argument.
  SpreadsheetApp.getActiveSpreadsheet().toast(string, 'Success', 2);
}



// Alert pop up for data clear warning
// @param: string - message to be displayed
function warn(string) {
  var ui = SpreadsheetApp.getUi();
  //alert popup with yes and no button
  var response = ui.alert(string, ui.ButtonSet.YES_NO);

  //returns with user choice
  if (response == ui.Button.YES) {
    return true;
  } else {
    return false;
  }
}



// Alert pop up for export success
// @param: message - string sent from the export function
// @param: isTemplateLoadFail - bool about whether this message means that the template load sequence has failed
function operationComplete(messageEnum, isTemplateLoadFail) {
  var message = STATUS_MESSAGE_EXCEL[messageEnum] || STATUS_MESSAGE_EXCEL['1'];
  return popupShow(message, "", isTemplateLoadFail);
}

// Alert pop up for no template present
function noTemplate() {
  okWarn('Please load a template to continue.');
}



// Google alert popup with OK button
// @param: dialog - message to show
function okWarn(dialog) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(dialog, ui.ButtonSet.OK);
}


/*
 * =================
 * TEMPLATE CREATION
 * =================
 *
 * This function creates a template based on the model template data
 * Takes the entire data model as an argument
 *
 */

// function that manages template creation - creating the header row, formatting cells, setting validation
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldTypeEnums - list of fieldType enums from client params object
function templateLoader(model, fieldTypeEnums) {

  var fields = model.fields;
  var sheet;
  var newSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;

  //if not in advanced mode, ignore the fields only available for that mode

  model.fields = fields.filter(function (item, index) {
    if (!item.isAdvanced) {
      return item;
    }
  })


  return Excel.run(function (context) {
    // store the sheet and worksheet list for use later
    sheet = context.workbook.worksheets.getActiveWorksheet();
    //reset the hidden status of the spreadsheet
    var range = sheet.getRangeByIndexes(0, 0, 1, EXCEL_MAX_ROWS);
    range.columnHidden = false;

    var worksheets = context.workbook.worksheets;
    worksheets.load('items');

    return context.sync()
      .then(function () {
        let onWrongSheet = false;

        // check that no other worksheet has the same name as the one we need to call this sheet
        if (worksheets.items.length > 1) {
          worksheets.items.forEach(x => {
            if (x.name == newSheetName && x.id !== sheet.id) {
              onWrongSheet = true;
            }
          });
        }

        if (onWrongSheet) {
          return operationComplete(STATUS_ENUM.wrongSheet, true);
        } else {
          // otherwise set the sheet name, then create the template
          sheet.name = newSheetName;
          return context.sync()
            .then(function () {
              return sheetSetForTemplate(sheet, model, fieldTypeEnums, context, newSheetName);
            })
        }
      })
      .catch(/*fail quietly*/);
  })
}

// wrapper function to set the header row, validation rules, and any extra formatting
function sheetSetForTemplate(sheet, model, fieldTypeEnums, context, newSheetName) {

  // heading row - sets names and formatting (standard sheet)
  headerSetter(sheet, model.fields, model.colors, context);
  // set validation rules on the columns (standard sheet)
  contentValidationSetter(sheet, model, fieldTypeEnums, context);
  // set any extra formatting options (standard sheet)
  contentFormattingSetter(sheet, model, context);
  //set database fields (database sheet)
  dataBaseValidationSetter(newSheetName, model, fieldTypeEnums, context);
}



// Sets headings for fields
// creates an array of the field names so that changes can be batched to the relevant range in one go for performance reasons
// @param: sheet - the sheet object
// @param: fields - full field data
// @param: colors - global colors used for formatting
function headerSetter(sheet, fields, colors, context) {

  var headerNames = [],
    backgrounds = [],
    fontColors = [],
    fontWeights = [],
    fieldsLength = fields.length;

  for (var i = 0; i < fieldsLength; i++) {
    headerNames.push(fields[i].name);

    // set field text depending on whether is required or not
    var fontColor = (fields[i].required || fields[i].requiredForSubType) ? colors.headerRequired : colors.header;
    var fontWeight = fields[i].required ? 'bold' : 'normal';
    fontColors.push(fontColor);
    fontWeights.push(fontWeight);

    // set background colors based on if it is a subtype only field or not
    var background = fields[i].isSubTypeField ? colors.bgHeaderSubType : colors.bgHeader;
    backgrounds.push(background);
  }

  var range = sheet.getRangeByIndexes(0, 0, 1, fieldsLength);
  range.values = [headerNames];
  for (var i = 0; i < fieldsLength; i++) {
    var cellRange = sheet.getCell(0, i);
    cellRange.set({
      format: {
        fill: { color: backgrounds[i] },
        font: { color: fontColors[i], bold: fontWeights[i] == "bold" }
      }
    });
  }
  return context.sync();
}



// Sets validation on a per column basis, based on the field type passed in by the model
// a switch statement checks for any type requiring validation and carries out necessary action
// @param: sheet - the sheet object
// @param: model - full data to acccess global params as well as all fields
// @param: fieldTypeEnums - enums for field types
function contentValidationSetter(sheet, model, fieldTypeEnums, context) {
  // we can't easily get the max rows for excel so use the number of rows it always seems to have
  var nonHeaderRows = (1048576 - 1);
  for (var index = 0; index < model.fields.length; index++) {
    var columnNumber = index + 1,
      list = [];

    switch (model.fields[index].type) {

      // ID fields: restricted to numbers and protected
      case fieldTypeEnums.id:
      case fieldTypeEnums.subId:
        setPositiveIntValidation(sheet, columnNumber, nonHeaderRows, false);
        protectColumn(
          sheet,
          columnNumber,
          nonHeaderRows,
          model.colors.bgReadOnly,
          "ID field"
        );
        break;

      case fieldTypeEnums.text:
        setTextValidation(sheet, columnNumber, nonHeaderRows, false);
        break;

      // All other types
      default:
        //do nothing
        break;
    }
  }
}

// Sets validation on a per column basis, based on the field type passed in by the model
// a switch statement checks for any type requiring validation and carries out necessary action
// @param: sheet - the sheet object
// @param: model - full data to acccess global params as well as all fields
// @param: fieldTypeEnums - enums for field types
function dataBaseValidationSetter(mainSheetName, model, fieldTypeEnums, context) {
  // we can't easily get the max rows for excel so use the number of rows it always seems to have
  for (var index = 0; index < model.fields.length; index++) {
    var columnNumber = index + 1,
      list = [];

    switch (model.fields[index].type) {
      // DROPDOWNS and MULTIDROPDOWNS are both treated as simple dropdowns (Sheets does not have multi selects)
      case fieldTypeEnums.drop:
      case fieldTypeEnums.multi:
        var fieldList = model.fields[index].values;
        for (var i = 0; i < fieldList.length; i++) {
          list.push(setListItemDisplayName(fieldList[i]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, false, context);
        break;

      // RELEASE fields are dropdowns with the values coming from a project wide set list
      case fieldTypeEnums.release:
        for (var l = 0; l < model.projectReleases.length; l++) {
          list.push(setListItemDisplayName(model.projectReleases[l]));
        }
        setDropdownValidation(mainSheetName, columnNumber, list, false, context);
        break;

      // All other types
      default:
        //do nothing
        break;
    }
  }
  return context.sync();
}


// create dropdown validation on set column based on specified values
// @param: sheet - the sheet object
// @param: dbSheet - the dbSheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: list - array of values to show in a dropdown and use for validation
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
// @param: fieldName - the name of the field to be used as a link between the two worksheets
// @param: context - Excel context to be synced
async function setDropdownValidation(mainSheetName, columnNumber, list, allowInvalid, context) {
  //max rows for Excel
  var nonHeaderRows = 1048576 - 1;
  //first, write the values to the dbSheet
  var values = [];
  list.forEach(function (item) {
    var itemArray = [item];
    values.push(itemArray);
  });

  var dbSheetRange = context.workbook.worksheets.getItem(params.dataSheetName).getRangeByIndexes(0, columnNumber - 1, list.length, 1);
  dbSheetRange.values = values;
  context.sync();
  //Now, point the fields in the mainsheet to the database worksheet (source)
  var range = context.workbook.worksheets.getItem(mainSheetName).getRangeByIndexes(1, columnNumber - 1, nonHeaderRows, 1);
  range.dataValidation.clear();

  range.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: dbSheetRange
    }
  };
  await context.sync();
}


// create positive integer validation on set column based on specified values
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setPositiveIntValidation(sheet, columnNumber, rowLength, allowInvalid) {
  var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
  range.dataValidation.clear();

  var greaterThanZeroRule = {
    wholeNumber: {
      formula1: 0,
      operator: Excel.DataValidationOperator.greaterThan
    }
  };
  range.dataValidation.rule = greaterThanZeroRule;

  range.dataValidation.prompt = {
    message: "Please enter a positive number.",
    showPrompt: true,
    title: "Positive numbers only."
  };
  range.dataValidation.errorAlert = {
    message: "Sorry, only positive numbers are allowed",
    showAlert: true,
    style: "Stop",
    title: "Negative Number Entered"
  };
}


// create text validation on set column base
// @param: sheet - the sheet object
// @param: columnNumber - int of the column to validate
// @param: rowLength - int of the number of rows for range (global param)
// @param: allowInvalid - bool to state whether to restrict any values to those in validation or not
function setTextValidation(sheet, columnNumber, rowLength, allowInvalid) {
  var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);
  range.dataValidation.clear();

  range.numberFormat = '@';

  range.dataValidation.errorAlert = {
    message: "Sorry, only text string is allowed.",
    showAlert: true,
    style: "Stop",
    title: "Invalid entry"
  };
}

// format columns based on a potential range of factors - eg hide unsupported columns
// @param: sheet - the sheet object
// @param: model - full model data set
function contentFormattingSetter(sheet, model) {
  for (var i = 0; i < model.fields.length; i++) {
    var columnNumber = i + 1;
    var nonHeaderRows = (1048576 - 1);

    // protect column
    // read only fields - ie ones you can get from Spira but not create in Spira (as with IDs - eg task component)
    if (model.fields[i].unsupported || model.fields[i].isReadOnly) {
      var warning = "";
      if (model.fields[i].unsupported) {
        warning = model.fields[i].name + "unsupported";
      } else if (model.fields[i].isReadOnly) {
        warning = model.fields[i].name + " is read only";
      }

      protectColumn(
        sheet,
        columnNumber,
        nonHeaderRows,
        model.colors.bgReadOnly,
        warning
      );
    }
    // hide this column if specified in the field model
    if (model.fields[i].isHidden) {
      var warning = "";
      warning = model.fields[i].name + " is hidden";
      hideColumn(
        sheet,
        columnNumber,
        nonHeaderRows,
        model.colors.bgReadOnly
      );

    }

    // green background for test run editable fields
    if (model.fields[i].sendField) {
      var warning = "";
      warning = model.fields[i].name + " is hidden";
      changeColorColumn(
        sheet,
        columnNumber,
        nonHeaderRows,
        model.colors.bgRunField
      );

    }

  }
}



// protects specific column. Edits still allowed - current user not excluded from edit list, but could in future
// @param: sheet - the sheet object
// @param: columnNumber - int of column to hide
// @param: rowLength - int of default number of rows to apply any formattting to
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
// @param: name - string description for the protected range
function protectColumn(sheet, columnNumber, rowLength, bgColor, name) {
  var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);

  // set the background color
  range.set({ format: { fill: { color: bgColor } } });

  // now we can add data validation
  // the easiest hack way to not allow any entry into the cell is to make sure its text length can only be zero
  range.dataValidation.clear();
  var textLengthZero = {
    textLength: {
      formula1: 0,
      operator: Excel.DataValidationOperator.equalTo
    }
  };
  range.dataValidation.rule = textLengthZero;

  range.dataValidation.prompt = {
    message: "This is a protected field and not user editable.",
    showPrompt: true,
    title: "No entry allowed."
  };
  range.dataValidation.errorAlert = {
    message: "Sorry, this is a protected field",
    showAlert: true,
    style: "Stop",
    title: "No entry allowed"
  };

}

// makes a specific cell read-only
// @param: sheet - the sheet object
// @param: col - int of column to protect
// @param: row - int of the row to protect
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
// @param: name - string description for the protected range
function protectCell(sheet, col, row, bgColor, name) {
  var cellRange = sheet.getCell(row + 1, col);

  // set the background color
  cellRange.set({ format: { fill: { color: bgColor } } });

  // now we can add data validation
  // the easiest hack way to not allow any entry into the cell is to make sure its text length can only be zero
  cellRange.dataValidation.clear();
  var textLengthZero = {
    textLength: {
      formula1: 0,
      operator: Excel.DataValidationOperator.equalTo
    }
  };
  cellRange.dataValidation.rule = textLengthZero;

  cellRange.dataValidation.prompt = {
    message: name,
    showPrompt: true,
    title: "Protected Field."
  };
  cellRange.dataValidation.errorAlert = {
    message: name,
    showAlert: true,
    style: "Stop",
    title: "Protected Field."
  };

}

// makes a specific row read-only
// @param: sheet - the sheet object
// @param: row - int of the row to protect
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
// @param: name - string description for the protected range
function protectRow(sheet, colSize, startRow, endRow, bgColor, name) {
  var rowRange = sheet.getRangeByIndexes(startRow + 1, 0, endRow - startRow, colSize);
  // set the background color
  rowRange.set({ format: { fill: { color: bgColor } } });
  // now we can add data validation
  // the easiest hack way to not allow any entry into the cell is to make sure its text length can only be zero
  rowRange.dataValidation.clear();
  var textLengthZero = {
    textLength: {
      formula1: 0,
      operator: Excel.DataValidationOperator.equalTo
    }
  };
  rowRange.dataValidation.rule = textLengthZero;
  rowRange.dataValidation.prompt = {
    message: name,
    showPrompt: true,
    title: "Protected Field."
  };
  rowRange.dataValidation.errorAlert = {
    message: name,
    showAlert: true,
    style: "Stop",
    title: "Protected Field."
  };
}


// hides a specific column range
// @param: sheet - the sheet object
// @param: columnNumber - int of column to hide
// @param: rowLength - int of default number of rows to apply any formattting to
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
function hideColumn(sheet, columnNumber, rowLength, bgColor) {
  // only for google as cannot protect individual cells easily in Excel
  var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);

  // set the background color
  range.set({ format: { fill: { color: bgColor } } });

  //hide the column
  range.columnHidden = true;
}

// change a specific column background color
// @param: sheet - the sheet object
// @param: columnNumber - int of column to hide
// @param: rowLength - int of default number of rows to apply any formattting to
// @param: bgColor - string color to set on background as hex code (eg '#ffffff')
function changeColorColumn(sheet, columnNumber, rowLength, bgColor) {
  // only for google as cannot protect individual cells easily in Excel
  var range = sheet.getRangeByIndexes(1, columnNumber - 1, rowLength, 1);

  // set the background color
  range.set({ format: { fill: { color: bgColor } } });

}

// resets backgroung colors to its original - used before a GET command
// @param: sheetName - current sheet name

function resetSheetColors(model, fieldTypeEnums, sheetRangeOld) {
  Excel.run(function (ctx) {
    var fields = model.fields;
    var columnCount = Object.keys(fields).length;

    //get the previous data number of rows
    var rowCount;
    var sheetOldData = sheetRangeOld.values;

    for (var i = 0; i < sheetOldData.length; i++) {
      // stop at the first row that is fully blank
      if (sheetOldData[i].join("") === "") {
        break;
      }
      else {
        rowCount = i;
      }
    }

    //complete data range from old data
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRangeByIndexes(1, 0, rowCount + 1, columnCount);

    //reset each column color schema (depending on property type)
    for (var j = 0; j < columnCount; j++) {

      var subColumnRange = range.getColumn(j);
      var fieldType = fields[j].type;
      var isReadOnly = fields[j].isReadOnly;
      var isRunField = fields[j].sendField;
      var bgColor;

      if (fieldType == fieldTypeEnums.id || fieldType == fieldTypeEnums.subId || isReadOnly) {
        subColumnRange.format.fill.color = model.colors.bgReadOnly;
      }
      else {
        if (isRunField) {
          subColumnRange.format.fill.color = model.colors.bgRunField;
        }
        else {
          subColumnRange.format.fill.clear();
        }
      }
    }
    range.delete(Excel.DeleteShiftDirection.up);
    return ctx.sync();
  }).catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
    }
  });
}

// resets validations - used before a GET command
// @param: sheetName - current sheet name
function resetSheet(model) {
  Excel.run(function (ctx) {
    var fields = model.fields;
    //complete data range from old data
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS, fields.length);
    range.delete(Excel.DeleteShiftDirection.up);

    ctx.sync();

    //clear database worksheet
    var worksheet = context.workbook.worksheets.getItemOrNullObject(params.dataSheetName);
    worksheet.getRange().clear();

    return ctx.sync();
  }).catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
    }
  });
}

// removes error messages from the 1st field of each row, in case there's any. Only the IDs are preserved
// @param: sheetData - the sheet data object
// returns the filtered sheetData object

function clearErrorMessages(sheetData) {
  for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
    // stop at the first row that is fully blank
    if (sheetData[rowToPrep].join("") === "") {
      break;
    } else {
      //the error messages are placed always at the position 1 of each row
      if (isNaN(sheetData[rowToPrep][1])) {
        try {
          sheetData[rowToPrep][1] = sheetData[rowToPrep][1].split(",")[1];
        }
        catch (err) {
          //do nothing
        }
      }
      //also, add -1 as ID for blank (invalid) new PUT values
      if (!sheetData[rowToPrep][1]) {
        sheetData[rowToPrep][1] = "-1";
      }

    }
  }
  return sheetData;
}

//Function that clear all the comments from the inicitial ID columns in the Spreadsheet
//@param model - contains information related to the artifact we are working with
async function resetComments(model, sheetData, context) {
  try {
    // await Excel.run(async (ctx) => {
    var fields = model.fields;
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
      // stop at the first row that is fully blank
      if (sheetData[rowToPrep].join("") === "") {
        break;
      } else {

        //remove any comments from previous interactions
        for (var col = 0; col < fields.length; col++) {
          if (col == 1 || col == 2) {
            //only from the ID field
            var cellRange = sheet.getCell(rowToPrep + 1, col);
            //reset color to its original value
            cellRange.set({ format: { fill: { color: model.colors.bgReadOnly } } });
          }
          //reset messages from the comments column
          if (fields[col].isComments) {
            var cellRange = sheet.getCell(rowToPrep + 1, col);
            cellRange.clear();
            cellRange.set({ format: { fill: { color: model.colors.bgReadOnly } } });
            await context.sync();
          }
        }
      }
    }
  } catch (err) {
    console.log("There was an error cleaning the spreadsheet: " + err);
  }
}


/*
 * ================
 * SENDING TO SPIRA
 * ================
 *
 * The main function takes the entire data model and the artifact type
 * and calls the child function to set various object values before
 * sending the finished objects to SpiraPlan
 *
 */

// function that manages exporting data from the sheet - creating an array of objects based on entered data, then sending to Spira
// @param: model - full model object from client containing field data for specific artifact, list of project users, components, etc
// @param: fieldTypeEnums - list of fieldType enums from client params object
// @param: isUpdate - boolean that indicates if this is an update operation (true) or create operation (false)
async function sendToSpira(model, fieldTypeEnums) {

  // 0. SETUP FUNCTION LEVEL VARS 
  var entriesLog, extraEntriesLog;
  var fields = model.fields,
    artifact = model.currentArtifact,
    requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;


  // 1. get the active spreadsheet and first sheet

  return await Excel.run({ delayForCellEdit: true }, function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet(),
      sheetRange = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS, fields.length);
    sheet.load("name");
    sheetRange.load("values");

    return new Promise(function (resolve, reject) {
      context.sync()
        .then(function () {
          if (sheet.name == requiredSheetName) {
            var sheetData = sheetRange.values;
            //clear all the comments from possible last executions
            resetComments(model, sheetData, context).then(function () {
              //Clear error messages and comments from the fields, if any
              sheetData = clearErrorMessages(sheetData);
              //First, send the artifact entries for Spira
              var entriesForExport = createExportEntries(sheetData, model, fieldTypeEnums, fields, artifact);
              //Check if the data can actually be sent to Spira
              var preCheckLog = preCheckData(entriesForExport, model);
              //only proceed if there's no data pre-validation errors
              if (!preCheckLog.globalFailureStatus) {
                var extraEntriesForExport = createExtraExportEntries(sheetData, model, fieldTypeEnums, fields, artifact);
                return sendExportEntriesExcel(entriesForExport, '', model, fieldTypeEnums, fields, artifact, '').then(function (response) {
                  entriesLog = response;
                  return sendExportEntriesExcel('', extraEntriesForExport, model, fieldTypeEnums, fields, artifact, entriesLog.associations);
                }).then(function (responseExtra) {
                  extraEntriesLog = responseExtra;
                }).catch(function (err) {
                  reject()
                }).finally(function () {
                  resolve(updateSheetWithExportResults(entriesLog, extraEntriesLog, null, entriesForExport, extraEntriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context))
                })

              }
              else {
                //data pre-validation checks failed
                resolve(updateSheetWithExportResults(null, null, preCheckLog, entriesForExport, null, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context));
              }
            });
          }
          else {
            var log = {
              status: STATUS_ENUM.wrongSheet
            };
            return log;
          }
        })
        .catch();
    })
      .catch();
  })
}



/* Function to verify if the user data does not have any missing fields
* or invalid combinations, depending on specific conditions.
* @param entriesForExport: object containing user data to be sent
* @param model: object containing information about the artifact being sent
* @return log object containing either the success or error status of each TC
*/
function preCheckData(entriesForExport, model) {

  var fields = model.fields;
  var log = {
    globalFailureStatus: false,
    artifacts: []
  };
  var failedIds = [],
    notRunIds = [];

  //populating the failingIds and notRunIds arrays, based on model

  fields.forEach(function (item) {
    if (item.field == params.specialFields.executionStatusField && !item.isHidden) {
      //if we have the key execution statuses... 
      item.values.forEach(function (subItem) {
        //... get just the failing ones and include their Ids in the local array
        if (subItem.isFailedStatus) {
          failedIds.push(subItem.id);
        }
        else if (subItem.isNotRun) {
          //... get the no-Runs as well
          notRunIds.push(subItem.id);
        }
      });
    }
  });

  //checking every Test Case/ Test Step entry

  entriesForExport.forEach(function (item) {
    //checking every Test Case:
    //is there any 'Not Run' Test Step in it?
    var hasNotRunSteps = false;
    //is there any non-success Test Step in it?
    var hasFailedSteps = false;
    //getting every Test Step
    item[params.specialFields.testRunStepsField].forEach(function (subItem) {
      if (failedIds.includes(subItem[params.specialFields.executionStatusField])) {
        //if the Test Step failed and does not have an Actual Result, we must flag it!
        hasFailedSteps = true;
        if (!subItem[params.specialFields.preCheckField1]) {
          // 1. Fist verification: Actual Result
          // every non-success status must have an actual result
          log.globalFailureStatus = true;
          subItem.FailingCondition = params.preCheckEnums.actualResult;
          //if we have a Test Set ID in the parent (TC), also add this to the Test Step object
          if (item[params.specialFields.secondaryShellField]) {
            subItem[params.specialFields.secondaryShellField] = item[params.specialFields.secondaryShellField];
          }
          log.artifacts.push(subItem);
        }
        else {
          //do nothing, that's expected!
        }
      } else if (notRunIds.includes(subItem[params.specialFields.executionStatusField])) {
        //flag if it is a non-run
        hasNotRunSteps = true;
      }

    });
    //if the Test Case is NoRun and non-failure status, we must flag it!
    //2. Second verification: Execution Status
    //  In a Test Case, having a 'Not Run' exec. status is only acceptable
    // when there's at leats one non-success Test Step
    if (hasNotRunSteps && !hasFailedSteps) {
      log.globalFailureStatus = true;
      item.FailingCondition = params.preCheckEnums.executionStatus;
      log.artifacts.push(item);
    }
  });

  return log;
}



//function that verifies if a test Step has a valid TestCase parent
function isValidParent(entriesForExport) {
  for (let index = entriesForExport.length; index > 0; index--) {
    //if this is not a test step, we found the parent index
    if (!entriesForExport[index - 1].isSubType) {
      if (entriesForExport[index - 1].validationMessage) {
        //this is an invalid parent
        return false;
      }
      else {
        //this is a valid parent
        return true;
      }
    }
  }
  return true;
}


// 2. CREATE ARRAY OF ENTRIES
//2.1 Custom and Standard fields - Sent through the artifact API function (POST/PUT)
// loop to create artifact objects from each row taken from the spreadsheet
// vars needed: sheetData, artifact, fields, model, fieldTypeEnums,
function createExportEntries(sheetData, model, fieldTypeEnums, fields, artifact) {
  var lastIndentPosition = null,
    entriesForExport = [];

  for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
    // stop at the first row that is fully blank
    if (sheetData[rowToPrep].join("") === "") {
      break;
    } else {

      var rowChecks = {
        hasSubType: artifact.hasSubType,
        totalFieldsRequired: countRequiredFieldsByType(fields, false),
        totalSubTypeFieldsRequired: artifact.hasSubType ? countRequiredFieldsByType(fields, true) : 0,
        countRequiredFields: rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, false),
        countSubTypeRequiredFields: artifact.hasSubType ? rowCountRequiredFieldsByType(sheetData[rowToPrep], fields, true) : 0,
        subTypeIsBlocked: !artifact.hasSubType ? true : rowBlocksSubType(sheetData[rowToPrep], fields),
        spiraId: rowIdFieldInt(sheetData[rowToPrep], fields, fieldTypeEnums)
      },
        // create entry used to populate all relevant data for this row
        entry = {};

      // first check for errors
      var hasProblems = rowHasProblems(rowChecks);
      if (hasProblems) {
        entry.validationMessage = hasProblems;
        // if error free determine what field filtering is required - needed to choose type/subtype fields if subtype is present
      } else {
        var fieldsToFilter = relevantFields(model.fields);
        entry = createEntryFromRow(rowToPrep, sheetData, model, fieldTypeEnums, lastIndentPosition, fieldsToFilter);
        // FOR SUBTYPE ENTRIES add flag on entry if it is a subtype
        if (entry && fieldsToFilter === FIELD_MANAGEMENT_ENUMS.subType) {
          entry.isSubType = true;
        }
      }

      if (artifact.id == params.artifactEnums.testCases && entry.isSubType) {
        //if this is a testStep, check if the parent is valid
        var validParent = isValidParent(entriesForExport);
        if (!validParent) {
          //if the parent is not valid, mark that as an error
          entry = {};
          entry.validationMessage = 'Invalid TestCase parent. Please check your data.';
        }
      }

      if (entry) { entriesForExport.push(entry); }

    }
  }
  return entriesForExport;
}


// 3. FOR EXCEL ONLY: GET READY TO SEND DATA TO SPIRA + 4. ACTUALLY SEND THE DATA
// DIFFERENT TO GOOGLE: this uses js ES6 a-sync and a-wait for its function and subfunction
// check we have some entries and with no errors
// Create and show a message to tell the user what is going on
async function sendExportEntriesExcel(entriesForExport, extraEntriesForExport, model, fieldTypeEnums, fields, artifact, fieldsAssociations) {
  if (!entriesForExport.length && !extraEntriesForExport.length) {
    popupShow('There are no entries to send to Spira', 'Check Sheet')
    return "nothing to send";
  } else {
    popupShow('Starting to update...', 'Progress');
    // create required variables for managing responses for sending data to spiraplan
    var log = {
      errorCount: 0,
      successCount: 0,
      doNotContinue: false,
      // set var for parent - used to designate eg a test case so it can be sent with the test step post
      parentId: -1,
      entriesLength: entriesForExport.length,
      entries: [],
      associations: []
    };

    // loop through objects to send and update the log
    async function sendSingleEntry(i) {
      //if we not have a parent ID yet, set the correct parentId artifact for subtypes (needed for POST URL)
      if (Number(log.parentId) == -1 && artifact.hasSubType && entriesForExport[i].isSubType) {
        log.parentId = getAssociationParentId(entriesForExport, i, artifact.id);
        //let the child object to hold the parent ID field
        entriesForExport[i][ART_PARENT_IDS[artifact.id]] = log.parentId;
      }
      await manageSendingToSpira(entriesForExport[i], model.user, model.currentProject.id, artifact, fields, fieldTypeEnums, log.parentId)
        .then(function (response) {
          var association = null;
          if (!response.error) {
            //get the association TestRunStepID - TestStepID
            association = getAssociationFromResponse(response.fromSpira, i);
            log.associations = [...log.associations, ...association];
          }
          // update the parent ID for a subtypes based on the successful API call
          if (artifact.hasSubType) {
            log.parentId = response.parentId;
          }
          log = processSendToSpiraResponse(i, response, entriesForExport, artifact, log, false);
        })
      //reset the variable to the next position
      log.parentId = -1;
    }

    // 4. SEND DATA TO SPIRA AND MANAGE RESPONSES
    // KICK OFF THE FOR LOOP (IE THE FUNCTION ABOVE) HERE
    // We use a function rather than a loop so that we can more readily use promises to chain things together and make the calls happen synchronously
    // we need the calls to be synchronous because we need to do the status and ID of the preceding entry for hierarchical artifacts
    //first, send standard and custom artifact properties
    //a) Standard Entries (i.e.: Test Runs)
    for (var i = 0; i < entriesForExport.length; i++) {
      if (!entriesForExport[i].skip) {
        if (!log.doNotContinue) {
          log = checkSingleEntryForErrors(entriesForExport[i], log, artifact);

          if (log.entries.length && log.entries[i] && log.entries[i].error) {
            // do nothing 
          } else {
            await sendSingleEntry(i);
          }
        }
      }
      else {
        var skip = {
          "skip": true
        };
        log.entries.push(skip);
        log.successCount++;

      }
    }
    //b) Extra Entries (i.e.: Incidents associated to Test Run Steps)
    for (var k = 0; k < extraEntriesForExport.length; k++) {
      //make sure we don't have any error and we have the comment populated
      if (Object.keys(extraEntriesForExport[k]).length > 1) {
        await sendSingleExtraEntry(k);
      }
    }

    // loop through association objects to send
    async function sendSingleExtraEntry(k) {
      //First, assign the correct TestRunStepID to each Incident object
      var convertedEntry = await convertExtraEntry(extraEntriesForExport[k], fieldsAssociations);
      //Then, get the Incident object from the model
      let incidentObject = { id: params.artifactEnums.incidents };

      await manageSendingToSpira(convertedEntry, model.user, model.currentProject.id, incidentObject, fields, fieldTypeEnums, 0)
        .then(function (response) {
          // update the parent ID for a subtypes based on the successful API call
          log = processSendToSpiraResponse(k, response, extraEntriesForExport, artifact, log, true);
        })

    }

    // review all activity and set final status
    log.status = setFinalStatus(log);
    // call the final function here - so we know that it is only called after the recursive function above (ie all posting) has ended
    return log;
  }
}


//Function to associate the TestRunStepId to the Test Step ID used to create a new Incident in Spira
//newIncident: the object contaning the new Incident information to be sent to Spira
//fieldsAssociation: the association between the TestRunStepId and TestStepId
//@return the newIncident object containing the correct association 
async function convertExtraEntry(newIncident, fieldsAssociations) {
  fieldsAssociations.forEach(function (association) {
    if ((association[params.specialFields.standardAssociationField] == newIncident[params.specialFields.secondaryAssociationField]) &&
      (newIncident.position == association.position)) {
      //if this is the Test Step we are looking for, associate it with the RunStepID
      newIncident[params.specialFields.secondaryAssociationField + 's'] = [association[params.specialFields.secondaryAssociationField]];
      delete newIncident[params.specialFields.secondaryAssociationField];
    }
  });
  return newIncident;
}




//Function to retrieve the TestRunStepID associated with the TestStepID
//in case the user wants to log an Incident from it
//input: the raw response of PUT Test Run from the server
//output: the association output object
function getAssociationFromResponse(response, i) {
  var associations = [];
  var objResponse = JSON.parse(response);
  //make sure we have the necessary data
  if (objResponse[0].hasOwnProperty('TestRunSteps')) {
    objResponse[0].TestRunSteps.forEach(function (item) {
      var association = {};
      if (item[params.specialFields.standardAssociationField]) {
        association[params.specialFields.standardAssociationField] = item[params.specialFields.standardAssociationField];
        association.position = i;
      }
      if (item[params.specialFields.secondaryAssociationField]) {
        association[params.specialFields.secondaryAssociationField] = item[params.specialFields.secondaryAssociationField];
      }
      if (Object.keys(association).length > 1) {
        associations.push(association);
      }
    });
  }
  return associations;
}







// 5. SET MESSAGES AND FORMATTING ON SHEET
function updateSheetWithExportResults(entriesLog, extraEntriesLog, preCheckingLog, entriesForExport, extraEntriesForExport, sheetData, sheet, sheetRange, model, fieldTypeEnums, fields, artifact, context) {
  var extraFieldCounter = 0;
  var row = 0;
  var entriesCounter = 0;
  var rowCounter = 0;
  // first handle cell formatting
  while (sheetData[row].join("") !== "") {
    rowCounter++;
    var rowBgColors = [],
      rowNotes = [],
      rowValues = [];

    for (var col = 0; col < fields.length; col++) {
      var bgColor,
        note = null,
        extraFieldNote = null,
        associationNote = null,
        value = sheetData[row][col];
      var cellRange = sheet.getCell(row + 1, col);

      async function addCellComment(cellRange, bgColor, checkingFieldNote) {
        bgColor = model.colors.warning;
        cellRange.set({ format: { fill: { color: bgColor } } });
        var subRange = sheet.getCell(row + 1, 0);
        subRange.values = [[checkingFieldNote]];
      }

      //handling the pre-checking log
      if (preCheckingLog != null) {
        //checking if the TC failed
        if (col == 1) {
          preCheckingLog.artifacts.forEach(function (item) {
            //if we had an execution status pre-checking failure, log it to the TC cell
            if (item[params.specialFields.standardShellField] == value && item.FailingCondition == params.preCheckEnums.executionStatus) {
              //since we can have duplicate values in the sheet, we need an extra check
              if ((!item[params.specialFields.secondaryShellField] && sheetData[row][TX_ID_COLUMN_INDEX] == '') ||
                (item[params.specialFields.secondaryShellField] && sheetData[row][TX_ID_COLUMN_INDEX] != '')) {
                var checkingFieldNote = 'Invalid Execution Statuses: This TestCase contains an invalid execution statuses combination. For further information, please refer to the documentation.';
                addCellComment(cellRange, model.colors.warning, checkingFieldNote);
              }
            }
          });
        }
        else if (col == 2) {
          //checking if the TS failed
          preCheckingLog.artifacts.forEach(function (item) {
            var isTestSet = IsTestSet(sheetData, row);
            //if we had an execution status pre-checking failure, log it to the TC cell
            if (item[params.specialFields.standardAssociationField] == value && item.FailingCondition == params.preCheckEnums.actualResult) {
              //since we can have duplicate values in the sheet, we need an extra check
              if ((!item[params.specialFields.secondaryShellField] && !isTestSet) ||
                (item[params.specialFields.secondaryShellField] && isTestSet)) {
                var checkingFieldNote = 'Missing Actual Result: This TestStep needs to have an Actual Result, since it failed.';
                addCellComment(cellRange, model.colors.warning, checkingFieldNote);
              }
            }
          });
        }
        var returnLog = {
          status: STATUS_ENUM.preCheckingError
        };
      }

      // we may have more rows than entries - because the entries can be stopped early (eg when an error is found on a hierarchical artifact)
      if (entriesLog != null) {
        //check for extra fields error
        if (fields[col].extraArtifact && extraEntriesLog != null && sheetData[row][col] != '') {
          if (extraEntriesLog.entries[extraFieldCounter]) {
            if (extraEntriesLog.entries[extraFieldCounter].error) {
              bgColor = model.colors.warning;
              extraFieldNote = 'This incident could not be created in Spira due to an unkown error. ' +
                'Please check your user permissions and the spreadsheet data. If the problem persists, please contact your administrator.';
            }
            else {
              //setting up the results comments for extra files (aka Incidents)
              try {
                var subRange = sheet.getCell(row + 1, 0);
                subRange.values = [['Success! Spira Incident ID ' + extraEntriesLog.entries[extraFieldCounter].details.resultId]];
              } catch (err) {
                console.log("There was a problem when adding a comment: " + err);
              }
            }
          }
          extraFieldCounter++;
        }
        else if (col == 1 && !fields[col].extraArtifact && (sheetData[row][col] != "" && sheetData[row][col] != "-1")) {
          //we only log error info to the head row of each Test Run
          bgColor = setFeedbackBgColor(sheetData[row][col], entriesLog.entries[entriesCounter].error, fields[col], fieldTypeEnums, artifact, model.colors);
          value = setFeedbackValue(sheetData[row][col], entriesLog.entries[entriesCounter].error, fields[col], fieldTypeEnums, entriesLog.entries[entriesCounter].newId || "", null, col);
          note = setFeedbackNote(sheetData[row][col], entriesLog.entries[entriesCounter].error, fields[col], fieldTypeEnums, entriesLog.entries[entriesCounter].message, value);

          //setting up the results comments for Test Runs
          if (!entriesLog.entries[entriesCounter].error) {
            try {
              var subRange = sheet.getCell(row + 1, 0);
              subRange.values = [['Success! Spira TestRun ID ' + entriesLog.entries[entriesCounter].details.resultId]];
            } catch (err) {
              console.log("There was a problem when adding a comment: " + err);
            }
          }
          entriesCounter++;
        }
        else if (col == 1 && sheetData[row][col] == "-1") {
          //clear internal values, so they are not displayed to the user
          value = '';
        }

        if (note) {
          rowNotes.push(note);
          note = null;
        }

        if (bgColor) {
          cellRange.set({ format: { fill: { color: bgColor } } });
          bgColor = null;
        }
        //cellRange.values = [[value]];
        value = null;
        var returnLog = entriesLog;
      }

      var rowFirstCell = sheet.getCell(row + 1, 0);
      if (rowNotes.length) {
        rowFirstCell.set({ format: { fill: { color: model.colors.warning } } });
        rowFirstCell.values = [[rowNotes.join()]];
      }
    }
    row++;
  }
  protectColumn(
    sheet,
    1,
    rowCounter,
    model.colors.bgReadOnly,
    "Spira Log field"
  );
  var subRange = sheet.getRangeByIndexes(0, 0, rowCounter, 1);
  subRange = setRangeBorders(subRange, model.colors.cellBorder);
  return context.sync().then(function () { return returnLog; });
}

function checkSingleEntryForErrors(singleEntry, log, artifact) {
  var response = {};
  // skip if there was an error validating the sheet row
  if (singleEntry.validationMessage) {
    response.error = true;
    response.message = singleEntry.validationMessage;
    log.errorCount++;
    log.entries.push(response);

    // skip if a sub type row does not have a parent to hook to
  } else if (singleEntry.isSubType && !log.parentId) {
    response.error = true;
    response.message = "can't add a child type when there is no corresponding parent type";
    log.errorCount++;
    log.entries.push(response);
  }
  return log;
}

// utility function to set final status of the log
// @param: log - log object
// returns enum for the final status
function setFinalStatus(log) {
  if (log.errorCount) {
    //check if any error message is about data validation etc - if not then all the message are about the entry already existing in Spira (where the message is the INT of the id)
    let logEntriesOnlyAboutIds = true;
    const logMessages = log.entries.filter(x => x.message);
    if (logMessages.length) {
      for (let index = 0; index < logMessages.length; index++) {
        if (!Number.isInteger(parseInt(logMessages[index].message))) {
          logEntriesOnlyAboutIds = false;
          break;
        }
      }
    }

    if (logEntriesOnlyAboutIds) {
      return STATUS_ENUM.existingEntries;
    } else if (log.errorCount == log.entriesLength) {
      return STATUS_ENUM.allError;
    } else {
      return STATUS_ENUM.someError;
    }
  } else {
    return STATUS_ENUM.allSuccess;
  }
}



// function that reviews a specific cell against it's field and errors for providing UI feedback on errors
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldTypeEnums - enum information about field types
// @param: artifact - the currently selected artifact
// @param: colors - object of colors to use based on different conditions
function setFeedbackBgColor(cell, error, field, fieldTypeEnums, artifact, colors) {
  if (error) {
    // if we have a validation error, we can highlight the relevant cells if the art has no sub type
    if (!artifact.hasSubType) {
      if (field.required && cell === "") {
        return colors.warning;
      } else {
        // keep original formatting
        if (field.type == fieldTypeEnums.subId || field.type == fieldTypeEnums.id || field.unsupported) {
          return colors.bgReadOnly;
        } else {
          return null;
        }
      }

      // otherwise highlight the whole row as we don't know the cause of the problem
    } else {
      return colors.warning;
    }

    // no errors
  } else {
    // keep original formatting
    if (field.type == fieldTypeEnums.subId || field.type == fieldTypeEnums.id || field.unsupported) {
      return colors.bgReadOnly;
    } else {
      return null;
    }
  }
}


// function that checks if a given row is part of a Test Set
// @param: sheetData - set of spreadsheet rows
// @param: row - the row number to be analized
// @return: boolean that represents if the giver row is or is not part of a TestSet
function IsTestSet(sheetData, row) {
  //looking back for a Test Case header
  for (var i = row - 1; i >= 0; i--) {
    if (sheetData[i][TC_ID_COLUMN_INDEX] != '') {
      //this is a Test Case header -> we need to check if this has a Test Set ID
      if (sheetData[i][TX_ID_COLUMN_INDEX] != '') {
        return true;
      }
      else {
        return false;
      }
    }
  }
  return false;
}

// function that reviews a specific cell against it's field and sets any notes required
// currently only adds error message as note to ID field
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldTypeEnums - enum information about field types
// @param: message - relevant error message from the entry for this row
// @param: value - original ID of the artifact to be kept in case of an error
function setFeedbackNote(cell, error, field, fieldTypeEnums, message, value) {
  // handle entries with errors - add error notes into ID field
  if (error && field.type == fieldTypeEnums.id) {
    //invalid new rows always have -1 in the ID field
    if (value == "-1") {
      //however, special subtype IDs can be blank, that's not an error
      if (!SUBTYPE_IDS.includes(field.field)) {
        return "Error: you can't send new data to Spira when updating. Please use the 'Send data to Spira' option."
      }
      else {
        return ',Error: ' + message;
      }
    }
    else {
      return value + ',' + message;
    }
  } else {
    return null;
  }
}

// function that updates id fields with new values, otherwise returns existing value
// @param: cell - value contained in specific cell beinq queried
// @param: error - bool flag as to whether the entire row the cell is in contains an error
// @param: field - the field specific to the cell
// @param: fieldTypeEnums - enum information about field types
// @param: newId - int that is the newly created Id for this row
// @param: isSubType - bool if row is subtype or not - false on error as there will be no id to add anyway
function setFeedbackValue(cell, error, field, fieldTypeEnums, newId, isSubType, col) {
  // when there is an error we don't change any of the cell data
  if (error) {
    return cell;

    // handle successful entries - ie add ids into right place
  } else {
    var newIdToEnter = newId || "";
    if (!isSubType && field.type == fieldTypeEnums.id) {
      return newIdToEnter;
    } else if (isSubType && field.type == fieldTypeEnums.subId) {
      return newIdToEnter;
    } else if (isSubType && field.type == fieldTypeEnums.id && col == 0 && cell == -1) {
      //this fix the visual bug for test steps
      return '';
    } else {
      return cell;
    }
  }
}



// on determining that an entry should be sent to Spira, this function handles calling the API function, and parses the data on both success and failure
// @param: entry - object of the specific entry in format ready to attach to body of API request
// @param: parentId - int of the parent id for this specific loop - used for attaching subtype children to the right parent artifact
// @param: artifact - object of the artifact being used here to help manage what specific API call to use
// @param: user - user object for API call authentication
// @param: projectId - int of project id for API call
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldTypeEnums - object of all field types with enums
// @param: isUpdate - flag to indicate if this is an update operation
// @param: isComment - flag to indicate if this is a comment creation operation
function manageSendingToSpira(entry, user, projectId, artifact, fields, fieldTypeEnums, parentId) {
  var data,
    // make sure correct artifact ID is sent to handler (ie type vs subtype)
    artifactTypeIdToSend = entry.isSubType ? artifact.subTypeId : artifact.id,
    // set output parent id here so we know this function will always return a value for this
    output = {
      parentId: parentId,
      entry: entry,
      artifact: {
        artifactId: artifactTypeIdToSend,
        artifactObject: artifact
      }
    };

  // send object to relevant artifact post service

  if (artifact.id == params.artifactEnums.testRuns) {
    return putArtifactToSpira(entry, user, projectId, artifactTypeIdToSend, parentId)
      .then(function (response) {
        var errorStatus = response.error;
        output.fromSpira = response.text;

        if (!errorStatus) {
          // get the id/subType id of the updated artifact
          var artifactIdField = getIdFieldName(fields, fieldTypeEnums, entry.isSubType);
          //just repeat the id - it's the same 
          output.newId = entry[artifactIdField];
          // repeats the output parent ID only if the artifact has a subtype and this entry is NOT a subtype
          if (artifact.hasSubType && !entry.isSubType) {
            output.parentId = parentId;
          }
          //getting the resultId field
          var artifactKeyField = params.specialFields.standardResultField;
          var jsonResponse = JSON.parse(response.text);
          output.resultId = jsonResponse[0][artifactKeyField];
          return output;

        }
      })
      .catch(function (error) {
        //we have an error - so set the flag and the message
        output.error = true;
        if (error) {
          output.errorMessage = error;
        } else {
          output.errorMessage = "update attempt failed";
        }

        // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
        if (artifact.hasSubType && !entry.isSubType) {
          output.parentId = 0;
        }
        return output;
      });
  } else if (artifact.id == params.artifactEnums.incidents) {
    return postArtifactToSpira(entry, user, projectId, artifactTypeIdToSend)
      .then(function (response) {
        var errorStatus = response.error;
        output.fromSpira = response.text;
        if (!errorStatus) {
          var artifactKeyField = params.specialFields.secondaryResultField;
          var jsonResponse = JSON.parse(response.text);
          output.resultId = jsonResponse[artifactKeyField];
          return output;
        }
      })
      .catch(function (error) {
        //we have an error - so set the flag and the message
        output.error = true;
        if (error) {
          output.errorMessage = error;
        } else {
          output.errorMessage = "incident creation attempt failed";
        }

        // reset the parentId if we are not on a subType - to make sure subTypes are not added to the wrong parent
        if (artifact.hasSubType && !entry.isSubType) {
          output.parentId = 0;
        }
        return output;
      });
  }
}



/** Function that create Incident objects (not directly related to the main Test Run object)
 * the incident name is provided by the user, the description is automatically populated 
 * based on other fields
 * vars needed: sheetData, artifact, fields, model, fieldTypeEnums,
 * @return the incident(s) object populated
*/
function createExtraExportEntries(sheetData, model, fieldTypeEnums, fields, artifact) {
  var entriesForExport = [];
  var fields = model.fields;

  //count the future Test Run position - useful to avoid mismatches when there're duplicate TestCases in the spreadsheet
  var trCounter = -1;

  for (var rowToPrep = 0; rowToPrep < sheetData.length; rowToPrep++) {
    // stop at the first row that is fully blank
    if (sheetData[rowToPrep].join("") === "") {
      break;
    } else {

      // create entry used to populate all relevant data for this row
      var entry = {};
      var description = "<p>";
      var parentId = 0;



      //going thru all the columns
      for (var i = 0; i < sheetData[rowToPrep].length; i++) {

        //1. Make sure this is a Test Step.

        if (sheetData[rowToPrep][TC_ID_COLUMN_INDEX] == '-1' || sheetData[rowToPrep][TC_ID_COLUMN_INDEX] == '') {

          if (fields[i].type == fieldTypeEnums.subId) {
            parentId = sheetData[rowToPrep][i];
          }

          //check if this field has to be at the Incident description and is populated
          if (fields[i].extraIncDesc && sheetData[rowToPrep][i] != '') {
            description = description + "<strong><u>" + fields[i].name + ": </u></strong><br />" +
              sheetData[rowToPrep][i] + "<br /><br />";
          }

          //check if this row has an extra field and it is populated
          if (fields[i].extraArtifact && sheetData[rowToPrep][i] != '') {
            //create the object to send
            description = description + "<p>";
            entry.Name = sheetData[rowToPrep][i];
            entry.Description = description;
            //create the association field (to be replaced)
            entry[params.specialFields.secondaryAssociationField] = parentId;
            entry.position = trCounter;
          }
        }
      }

      // If it a Test Case, increase the TR counter
      if (sheetData[rowToPrep][TC_ID_COLUMN_INDEX] != '-1' && sheetData[rowToPrep][TC_ID_COLUMN_INDEX] != '') {
        trCounter++;
      }

      if (Object.keys(entry).length != 0) {

        entriesForExport.push(entry);
      }

    }
  }
  return entriesForExport;
}





// returns the correct parentId for the relevant indent position by looping back through the list of entries
// returns -1 if no match found
// @param: indent - int of the indent position to retrieve the parent for
// @param: previousEntries - object containing all successfully sent entries - with, if a hierarchical artifact, a hierarchy info object
function getHierarchicalParentId(indent, previousEntries) {
  // if there is no indent/ set to initial indent we return out immediately 
  if (indent === 0 || !previousEntries.length) {
    return -1;
  }
  for (var i = previousEntries.length - 1; i >= 0; i--) {
    // when the indent is greater - means we are indenting, so take the last array item and return out
    // check for presence of correct objects in item - should exist, as otherwise error should be thrown
    if (previousEntries[i].details && previousEntries[i].details.hierarchyInfo && previousEntries[i].details.hierarchyInfo.indent < indent) {
      return previousEntries[i].details.hierarchyInfo.id;
    }
  }
  return -1;
}

// returns the correct parentId for the relevant group position by looping back through the list of entries
// returns -1 if no match found
// @param: entriesForExport - the whole list of artifacts in the spreadsheet
// @param: i - target artifact position to look for
// @param: artifactId - artifact type
function getAssociationParentId(entriesForExport, i, artifactId) {
  // we need to know the artifact type to look for specific parent ID fields
  if (artifactId == ART_ENUMS.testCases) {
    //look for parents in the previous rows
    for (i; i > 0; i--) {
      //if we found the specific parentID for the artifact type, return it
      if (entriesForExport[i - 1].hasOwnProperty(ART_PARENT_IDS[artifactId])) {
        return entriesForExport[i - 1][ART_PARENT_IDS[artifactId]];
      }
    }
    return -1;
  }
}

// returns an int of the total number of required fields for the passed in artifact
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
// @param: forSubType - bool to determine whether to check for sub type required fields (true), or not - defaults to false
function countRequiredFieldsByType(fields, forSubType) {
  var count = 0;
  for (var i = 0; i < fields.length; i++) {
    if (forSubType != "undefined" && forSubType) {
      if (fields[i].requiredForSubType) {
        count++;
      }
    } else if (fields[i].required) {
      count++;
    }
  }
  return count;
}



// check to see if a row of data has entries for all required fields
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - the relevant fields for specific artifact, along with all metadata about each
function rowCountRequiredFieldsByType(row, fields, forSubType) {
  var count = 0;
  for (var i = 0; i < row.length; i++) {
    if (forSubType != "undefined" && forSubType) {
      if (fields[i].requiredForSubType && row[i]) {
        count++;
      }
    } else if (fields[i].required && row[i]) {
      count++;
    }

  }
  return count;
}



// check to see if a row for an artifact with a subtype has a field that can't be present if subtype fields are filled in
// this can be useful to make sure that one field - eg Test Case Name would make sure a test step is not created to avoid any confusion
// returns true if all required fields have (any) values, otherwise returns false
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
function rowBlocksSubType(row, fields) {
  var result = false;
  for (var column = 0; column < row.length; column++) {
    if (fields[column].forbidOnSubType && row[column]) {
      result = true;
    }
  }
  return result;
}



// check to see if a row for an artifact has any id field filled in with an int (not a string - a string could mean a different error message was previously added with Excel)
// returns false if id field is not an int, returns the ID in the cell if one is present as an INT (ie send back the Spira ID)
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldTypeEnums - object of all field types with enums
function rowIdFieldInt(row, fields, fieldTypeEnums) {
  var result = false;
  for (var column = 0; column < row.length; column++) {
    const cellIsIdField = fields[column].type === fieldTypeEnums.id || fields[column].type === fieldTypeEnums.subId;
    if (cellIsIdField && Number.isInteger(parseInt(row[column]))) {
      //set the result to the value of the ID so that the row is skipped but in the UI it looks the same - ie has the correct ID etc
      result = row[column];
      break;
    }
  }
  return result;
}



// checks to see if the row is valid - ie required fields present and correct as expected
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: rowChecks - object with different properties for different checks required
function rowHasProblems(rowChecks) {
  var problem = null;
  //carry out the problem analysis
  if (!rowChecks.hasSubType && rowChecks.countRequiredFields < rowChecks.totalFieldsRequired) {
    problem = "Fill in all required fields";
  } else if (rowChecks.hasSubType) {
    if (rowChecks.countSubTypeRequiredFields < rowChecks.totalSubTypeFieldsRequired && !rowChecks.countRequiredFields) {
      problem = "Fill in all required fields";
    } else if (rowChecks.countRequiredFields < rowChecks.totalFieldsRequired && !rowChecks.countSubTypeRequiredFields) {
      problem = "Fill in all required fields";
    }
  }
  return problem;
}

// checks to see if a coment has any problem related to it
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: f - object with different properties for different checks required
function commentHasProblems(artifact) {
  var problem = null;
  // SubTypes can't have comments
  if (artifact.isSubType) {
    problem = "Comment not allowed for this artifact type.";
  }
  return problem;
}

//checks if the text corresponds to a valid string
function isInt(value) {
  return !isNaN(value) && (function (x) { return (x | 0) === x; })(parseFloat(value))
}

//checks if the text we have for associations is a valid number OR a valid comma-separated string
// returns a string - empty if no errors present (to evaluate to false), or an error message object otherwise
// @ param: - row to inspect - object with different properties for different checks required
//          - artifact - the artifact object model
function associationHasProblems(row, artifact) {

  var problem = null;

  var associationText = row.filter(function (item, index) {
    if (artifact[index].association) {
      return row[index];
    }
  })

  if ((associationText + '') != '') {
    //depending on the user's language, the console can misinterpret commas as points
    associationText = (associationText + '').replace('.', ',');
    if (!isInt(associationText)) {
      { //in this case, we must have a comma-separated string
        var associationIds = (associationText + '').split(',');
        var associationCount = associationIds.length;
        if (associationCount == 1) {
          problem = "Artifact Association data wrong format."
        }
        else {
          //check if every chunk is an integer
          associationIds.forEach(function (item) {
            if (!isInt(item)) problem = "Artifact Association data wrong format.";
          });

        }
      }
    }

  }
  return problem;
}

/**
 * function relevantFields:
 * for a given artifact, select fields that are not part of the standard TestRun object, i.e.: fields that need extra API calls
 * @param {*} fields - object with model fields to be filtered
 * @returns : result - selected field(s)
 */
function relevantFields(fields) {

  var extraFields = [];

  fields.forEach(function (item) {
    if (item.extraArtifact || item.isComments) {
      extraFields.push(item.field);
    }

  });
  return extraFields;
}



// function creates a correctly formatted artifact object ready to send to Spira
// it works through each field type to validate and parse the values so object is in correct form
// any field that does not pass validation receives a null value
// @param: row - a 'row' of data that contains a single object representing all fields
// @param: model - full model with info about fields, dropdowns, users, etc
// @param: fieldTypeEnums - object of all field types with enums
// @param: lastIndentPosition - int used for calculating relative indents for hierarchical artifacts
// @param: fieldsToFilter - enum used for selecting fields to not add to object - defaults to using all if omitted
// @param: isUpdate - bool to flag id this is an update operation. If false, it is a creation operation
// @param: isComment - bool to flag if we will return a comment entry (true) or custom/standard (false)
function createEntryFromRow(index, rows, model, fieldTypeEnums, lastIndentPosition, fieldsToFilter) {
  var fields = model.fields;
  var entry = {};
  var parentId = 0;

  //1. Make sure this is a Test Case
  if (rows[index][1] != '-1') {

    //1.1 We need to turn an array of values in the row into a validated object
    for (var i = 0; i < rows[index].length; i++) {

      if (!fieldsToFilter.includes(fields[i].field)) {

        var value = null,
          customType = "",
          idFromName = 0;

        // double check data validation, convert dropdowns to required int values
        // sets both the value, and custom types - so that custom fields are handled correctly
        switch (fields[i].type) {

          // ID fields: restricted to numbers and blank on push, otherwise put
          case fieldTypeEnums.id:

            if (!isNaN(rows[index][i])) {
              value = rows[index][i];
            }
            customType = "IntegerValue";
            //getting the TC id for later
            if (i == 0) {
              parentId = value;
            }
            break;


          case fieldTypeEnums.subId:
            if (!isNaN(rows[index][i])) {
              value = rows[index][i];
            }
            customType = "IntegerValue";
            break;

          // INT fields
          case fieldTypeEnums.int:
            // only set the value if a number has been returned
            if (!isNaN(rows[index][i])) {
              value = rows[index][i];
              customType = "IntegerValue";
            }
            break;

          // DECIMAL fields
          case fieldTypeEnums.num:
            // only set the value if a number has been returned
            if (!isNaN(rows[index][i])) {
              value = rows[index][i];
              customType = "DecimalValue";
            }
            break;

          // BOOL as Sheets has no bool validation, a yes/no dropdown is used
          case fieldTypeEnums.bool:
            // 'True' and 'False' don't work as dropdown choices, so have to convert back
            if (rows[index][i] == "Yes") {
              value = true;
              customType = "BooleanValue";
            } else if (rows[index][i] == "No") {
              value = false;
              customType = "BooleanValue";
            }
            break;

          // DATES - parse the data and add prefix/suffix for WCF
          case fieldTypeEnums.date:
            if (rows[index][i]) {
              // for Excel, dates are returned as days since 1900 - so we need to adjust this for JS date formats
              const DAYS_BETWEEN_1900_1970 = 25567 + 2;
              const dateInMs = (rows[index][i] - DAYS_BETWEEN_1900_1970) * 86400 * 1000;
              value = convertLocalToUTC(new Date(dateInMs), dateInMs);
              customType = "DateTimeValue";
            }
            break;

          // ARRAY fields are for multiselect lists - currently not supported so just push value into an array to make sure server handles it correctly
          case fieldTypeEnums.arr:
            if (rows[index][i]) {
              value = [rows[index][i]];
              customType = ""; // array fields not used for custom properties here
            }
            break;

          // DROPDOWNS - get id from relevant name, if one is present
          case fieldTypeEnums.drop:
            idFromName = getIdFromName(rows[index][i], fields[i].values);
            if (idFromName) {
              value = idFromName;
              customType = "IntegerValue";
            }
            break;

          // MULTIDROPDOWNS - get id from relevant name, if one is present, set customtype to list value
          case fieldTypeEnums.multi:
            idFromName = getIdFromName(rows[index][i], fields[i].values);
            if (idFromName) {
              value = [idFromName];
              customType = "IntegerListValue";
            }
            break;

          // USER fields - get id from relevant name, if one is present
          case fieldTypeEnums.user:
            idFromName = getIdFromName(rows[index][i], model.projectUsers);
            if (idFromName) {
              value = idFromName;
              customType = "IntegerValue";
            }
            break;

          // COMPONENT fields - get id from relevant name, if one is present
          case fieldTypeEnums.component:
            idFromName = getIdFromName(rows[index][i], model.projectComponents);
            if (idFromName) {
              value = idFromName;
              // component is multi select for test cases but not for other artifacts
              customType = fields[i].isMulti ? "IntegerListValue" : "IntegerValue";
            }
            break;

          // RELEASE fields - get id from relevant name, if one is present
          case fieldTypeEnums.release:
            idFromName = getIdFromName(rows[index][i], model.projectReleases);
            if (idFromName) {
              value = idFromName;
              customType = "IntegerValue";
            }
            break;

          // All other types
          default:
            // just assign the value to the cell - used for text
            value = rows[index][i];
            customType = "StringValue";
            break;
        }

        // CUSTOM FIELDS:
        // check whether field is marked as a custom field and as the required property number
        if (fields[i].isCustom && fields[i].propertyNumber) {

          // if field has data create the object
          if (value) {
            var customObject = {};
            customObject.PropertyNumber = fields[i].propertyNumber;
            customObject[customType] = value;

            entry.CustomProperties.push(customObject);
          }

          // STANDARD FIELDS:
          // add standard fields in standard way - only add if field contains data
        } else if (value) {
          // if the standard field is a multi select type as set in the switch above, pass the value through in an array
          entry[fields[i].field] = (customType == "IntegerListValue") ? [value] : value;
        }
      }
    }

    //1.2 Adding extra fields for TestCases
    entry = testRunTCExtraFields(entry);

    //2. Getting the Test Steps
    var localIndex = index + 1; //1st TS position
    var tsObject = {
      TestRunSteps: []
    }
    while (rows[localIndex][TC_ID_COLUMN_INDEX] == '-1') { //while we have TestSteps
      var singleTs = [];
      //2.1 We need to turn an array of TS values in the row into a validated object
      for (var j = 0; j < rows[localIndex].length; j++) {
        value = null,
          customType = "",
          idFromName = 0;

        switch (fields[j].type) {
          case fieldTypeEnums.id:
            //replace TestCaseId for the real one
            if (j == 0 && rows[localIndex][j] == '-1') {
              value = parentId;
            }
            else {
              value = null;
            }
            customType = "IntegerValue";
            break;

          case fieldTypeEnums.subId:
            if (!isNaN(rows[localIndex][j])) {
              value = rows[localIndex][j];
            }
            customType = "IntegerValue";
            break;

          // DROPDOWNS - get id from relevant name, if one is present
          case fieldTypeEnums.drop:
            idFromName = getIdFromName(rows[localIndex][j], fields[j].values);
            if (idFromName) {
              value = idFromName;
              customType = "IntegerValue";
            }
            break;

          // All other types
          default:
            // just assign the value to the cell - used for text
            value = rows[localIndex][j];
            customType = "StringValue";
            break;

        }
        var test = {};
        // if field has data create the object
        if (value) {
          // if the standard field is a multi select type as set in the switch above, pass the value through in an array
          singleTs[fields[j].field] = (customType == "IntegerListValue") ? [value] : value;
        }
      }

      //2.2 Adding extra TS fields
      singleTs = testRunTSExtraFields(singleTs);

      tsObject["TestRunSteps"].push(singleTs);
      localIndex++;
    }

    //add the TestSteps to the TC main object
    entry = { ...entry, ...tsObject }
    return entry;
  }
  else {
    //It's just a Test Step. Already covered above. No need to handle it again.
    return null;
  }
}

//Function to add extra data to the TestRun body (TC related)
// @param: entry - the entry body to be send to Spira
// @return: the same entry body with the extra fields
function testRunTCExtraFields(entry) {
  var newEntry = {};
  newEntry = { ...entry, ...params.extraTcFixedFields }
  return newEntry;
}


//Function to add extra data to the TestRun body (TS related)
// @param: entry - the entry body to be send to Spira
// @return: the same entry body with the extra fields
function testRunTSExtraFields(entry) {
  var newEntry = {};
  newEntry = { ...entry, ...params.extraTsFixedFields }
  return newEntry;
}


//Converts a local time to UTC time
function convertLocalToUTC(convertedDate, originalDate) {
  originalDate = new Date(originalDate).toUTCString();
  var d = new Date();
  var offsetMinutes = d.getTimezoneOffset();
  var utcDate = new Date(convertedDate.getTime() + offsetMinutes * 60000);
  return utcDate.toISOString();
}

//Converts a UTC to local
function convertUTCtoLocal(originalDate) {
  var d = new Date();
  var offsetMinutes = d.getTimezoneOffset();
  var utcDate = new Date(originalDate.getTime() + offsetMinutes * 60000);
  return utcDate.toISOString();
}

// find the corresponding ID for a string value - eg from a dropdown
// dropdowns can only contain one item per row so we have to now get the IDs for sending to Spira
// @param: string - the string of the name value specified
// @param: list - the array of items with keys for id and name values
function getIdFromName(string, list) {
  for (var i = 0; i < list.length; i++) {
    if (setListItemDisplayName(list[i]) == string) {
      return list[i].id;

      // if there's no match with the item, let's try and match on just the name part of the list item  - this is the old way
      // this code is included to accomodate users who create their spreadsheets elsewhere and then dump the data in here without knowing the ids
    } else if (list[i] == unsetListItemDisplayName(string)) {
      return list[i].id;
    }
  }

  // return 0 if there's no match from either method
  return 0;
}

// for dropdown items we need to use the id as well as the name to make sure the entries are unique - so return a standard format here
// @param: item - object of the list item - contains a name and id
// returns the correctly formatted string - so that it is always set consistently
function setListItemDisplayName(item) {

  return item.name + " (#" + item.id + ")";

}

// removes the id from the end of a string to get the initial value, pre setting the display name
// @param: string - of the list item with the id added at the end as in setListItemDisplayName
// returns a new string with the regex match removed
function unsetListItemDisplayName(string) {
  var regex = / \(\#\d+\)$/gi;
  return string.replace(regex, "");
}


// finds and returns the field name for the specific artifiact's ID field
// @param: fields - object of the relevant fields for specific artifact, along with all metadata about each
// @param: fieldTypeEnums - object of all field types with enums
// @param: getSubType - optioanl bool to specify to return the subtype Id field, not the normal field (where two exist)
function getIdFieldName(fields, fieldTypeEnums, getSubType) {
  for (var i = 0; i < fields.length; i++) {
    var fieldToLookup = getSubType ? "subId" : "id";
    if (fields[i].type == fieldTypeEnums[fieldToLookup]) {
      return fields[i].field;
    }
  }
  return null;
}


// returns the correct relative indent position - based on the previous relative indent and other logic (int neg, pos, or zero)
// Currently the API does not support a call to place an artifact at a certain location.
// @param: indentCount - int of the number of indent characters set by user
// @param: lastIndentPosition - int sum of the actual indent positions used for the preceding entries
function setRelativePosition(indentCount, lastIndentPosition) {
  // the first time this is called, last position will be null
  if (lastIndentPosition === null) {
    return 0;
  } else if (indentCount > lastIndentPosition) {
    // only indent one level at a time
    return lastIndentPosition + 1;
  } else {
    // this will manage indents of same level or where outdents are required
    return indentCount;
  }
}

// anaylses the response from posting an item to Spira, and handles updating the log and displaying any messages to the user
function processSendToSpiraResponse(i, sentToSpira, entriesForExport, artifact, log, isExtra) {
  var response = {};
  response.details = sentToSpira;
  var operationString = "";
  if (isExtra) { operationString = " Incident "; }

  // handle success and error cases
  if (sentToSpira.error) {
    log.errorCount++;
    response.error = true;

    //handling different error messages
    if (sentToSpira.errorMessage.status == 409) {
      response.message = "Concurrency Date conflict: please reload your data and try again.";
    }
    else {
      if (sentToSpira.errorMessage.status == 400) {
        response.message = "Error: The Test Run could not be created. Please check your data and user permissions.";
      }
      else {
        response.message = "Error: The Test Run could not be created. Please check your data and user permissions.";
      }
    }
    //Sets error HTML modals
    popupShow('Error sending ' + operationString + (i + 1) + ' of ' + (entriesForExport.length), 'Progress');
  }
  else {
    log.successCount++;
    response.newId = sentToSpira.newId;

    //modal that displays the status of each artifact sent
    popupShow('Sent ' + operationString + (i + 1) + ' of ' + (entriesForExport.length) + '...', 'Progress');
  }

  // finally write out the response to the log and return
  log.entries.push(response);
  return log;
}


// EXCEL SPECIFIC FUNCTION - Verify the current active sheet to be used as the data source (offline TestRun)
// @param: model: full model object from client
// @param: enum of fieldTypeEnums used
function getFromSheetExcel(model, fieldTypeEnums) {
  //Perform a series of verifications to make sure the provided data is in good shape
  var log = {
    loadErrorCounter: 0,
    entries: []
  };

  return Excel.run(function (context) {
    var fields = model.fields;
    var sheet = context.workbook.worksheets.getActiveWorksheet(),
      sheetRange = sheet.getRangeByIndexes(0, 0, 1, 100),
      sheetRangeFull = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS, fields.length);
    var sheets = context.workbook.worksheets;
    sheet.load("name");
    sheetRange.load("values");
    sheetRangeFull.load("values");
    sheets.load("items/name");

    var requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
    var isDatabaseSheet = false;

    return context.sync()
      .then(function () {
        //1. The selected project matches the sheet name

        if (sheet.name != requiredSheetName) {
          log.loadErrorCounter++;
          log.entries.push("The selected product does not match the Spreadsheet data.");
        }

        //2. There is a database sheet (hidden or not)
        sheets.items.forEach(function (singleSheet) {
          if (singleSheet.name == params.dataSheetName) {
            isDatabaseSheet = true;
          }
        });

        if (!isDatabaseSheet) {
          log.loadErrorCounter++;
          log.entries.push("Database sheet is missing.");
        }
        //3. All the columns from the model are present

        var modelFieldsCounter = fields.length;
        var sheetHeader = sheetRange.values[0];
        var sheetFieldsCounter = 0;
        for (var i = 0; i < sheetHeader.length; i++) {

          if (sheetHeader[i] != '') {
            sheetFieldsCounter++;
          }
          else {
            break;
          }
        }
        //there's the same number of columns in the spreadsheet and in the model
        if (modelFieldsCounter != sheetFieldsCounter) {
          log.loadErrorCounter++;
          log.entries.push("There are columns missing in the spreadsheet.");
        }
        sheetRangeFull = sheetRangeFull.values;

        //4. There's no Test Case without TestStep
        for (var i = 0; i < sheetRangeFull.length - 1; i++) {

          var isCurrentTestCase = (sheetRangeFull[i][TC_ID_COLUMN_INDEX] != '' && sheetRangeFull[i][TC_ID_COLUMN_INDEX] != '-1');
          var isNextTestCase = (sheetRangeFull[i + 1][TC_ID_COLUMN_INDEX] != '' && sheetRangeFull[i + 1][TC_ID_COLUMN_INDEX] != '-1');

          //if we have a test case, in the next row we must have a test step!
          if (isCurrentTestCase && isNextTestCase) {
            log.loadErrorCounter++;
            log.entries.push("Invalid Test Case detected: missing Test Step(s).");
          }
        }
        return log;
      })

  })




}

// EXCEL SPECIFIC FUNCTION - handles getting paginated artifacts from Spira and displaying them in the UI
// @param: model: full model object from client
// @param: enum of fieldTypeEnums used
function getFromSpiraExcel(model, fieldTypeEnums) {
  return Excel.run(function (context) {
    var fields = model.fields;
    var sheet = context.workbook.worksheets.getActiveWorksheet(),
      sheetRange = sheet.getRangeByIndexes(1, 0, EXCEL_MAX_ROWS, fields.length);
    sheet.load("name");
    sheetRange.load("values");
    var requiredSheetName = model.currentArtifact.name + ", PR-" + model.currentProject.id;
    return context.sync()
      .then(function () {
        // only get the data if we are on the right sheet - the one with the template loaded on it
        if (sheet.name == requiredSheetName) {
          //first, reset sheet rows (if we had data from the last run)
          resetSheet(model);
          //then, clear the background colors of the spreadsheet (in case we had any errors in the last run)
          resetSheetColors(model, fieldTypeEnums, sheetRange);

          dataBaseValidationSetter(requiredSheetName, model, fieldTypeEnums, context);

          return getDataFromSpiraExcel(model, fieldTypeEnums).then((response) => {
            //error handling
            if (response == 'noData') {
              return operationComplete(STATUS_ENUM.noData, false);
            }
            else {
              return processDataFromSpiraExcel(response, model, fieldTypeEnums)
            }
          });
        } else {
          return operationComplete(STATUS_ENUM.wrongSheet, false);
        }
      })
  })
}

// Get all the data we need from Spira, combining different API operations
// @param: model - full model object from client
// @param: fieldTypeEnums - enum of fieldTypes used
async function getDataFromSpiraExcel(model, fieldTypeEnums) {
  // 1. get Test Cases from spira that are assigned to the logged user 
  var currentPage = 0;
  var artifacts = [];
  var getNextPage = true;

  async function getArtifactsPage(startRow) {
    await getArtifacts(
      model.user,
      model.currentProject.id,
      model.currentArtifact.id,
      startRow,
      GET_PAGINATION_SIZE,
      null
    ).then(function (response) {

      // if we got a non empty array back then we have artifacts to process
      if (response.body && response.body.length) {
        artifacts = artifacts.concat(response.body);
        // if we got less artifacts than the max we asked for, then we reached the end of the list in this request - and should stop
        if (response.body && response.body.length < GET_PAGINATION_SIZE) {
          getNextPage = false;
          // if we got the full page size back then there may be more artifacts to get
        } else {
          currentPage++;
        }
        // if we got no artifacts back, stop now
      } else {
        getNextPage = false;
      }
    })
  }

  while (getNextPage && currentPage < 100) {
    var startRow = (currentPage * GET_PAGINATION_SIZE) + 1;
    await getArtifactsPage(startRow);
  }

  // 2. If there were no artifacts at all break out now
  if (!artifacts.length) return "no artifacts were returned";


  // 3. Filtering the Test Cases to make sure they belong to the current project and also make sure they have steps - if not, they won't be displayed
  artifacts = artifacts.filter(function (item) {
    if ((item.ProjectId == model.currentProject.id) && item[model.currentArtifact.conditionField]) {
      return item;
    }
  });

  //function called below in the foreach call
  async function getArtifactSubs(art) {
    await getArtifacts(
      model.user,
      art.ProjectId,
      model.currentArtifact.subTypeId,
      null,
      null,
      art[params.specialFields.standardShellField],
      params.artifactEnums.testCases
    ).then(function (response) {
      // take action if we got any sub types back - ie if they exist for the specific artifact
      if (response.body && response.body.length && response.body[0][params.specialFields.testRunStepsField]) {
        var subTypeArtifactsWithMeta = response.body[0][params.specialFields.testRunStepsField].map(function (sub) {
          sub.isSubType = true;
          sub.parentId = art[params.specialFields.standardShellField];

          return sub;
        })
        // now add the steps into the original object
        artifactsWithSubTypes = artifactsWithSubTypes.concat(subTypeArtifactsWithMeta);
      }

    });
  };

  // 4. Get the Test Steps of the just-filtered Test Cases (via Shell)
  var artifactsWithSubTypes = [];

  for (var i = 0; i < artifacts.length; i++) {
    artifactsWithSubTypes.push(artifacts[i]);
    await getArtifactSubs(artifacts[i]);
  }

  // update the original array (I know that mutation is bad, but it makes things easy here)
  artifacts = artifactsWithSubTypes;

  /*
  5. At this point, we have TC+TS owned by the user. Now, we need to get those TC that came from TX owned by the user.
  */
  if (model.currentArtifact.hasSecondaryType) {

    //5.1 Retrieve the Test Sets owned by the user
    currentPage = 0;
    getNextPage = true;
    var secondaryArtifact = []; //var to store TX objects

    async function getSecondaryPage(startRow) {
      await getArtifacts(
        model.user,
        model.currentProject.id,
        model.currentArtifact.SecondaryTypeId,
        startRow,
        GET_PAGINATION_SIZE,
        null,
        null
      ).then(function (response) {
        // if we got a non empty array back then we have secondary Artifacts to process
        if (response.body && response.body.length) {
          secondaryArtifact = secondaryArtifact.concat(response.body);
          // if we got less secondary Artifacts than the max we asked for, then we reached the end of the list in this request - and should stop
          if (response.body && response.body.length < GET_PAGINATION_SIZE) {
            getNextPage = false;
            // if we got the full page size back then there may be more artifacts to get
          } else {
            currentPage++;
          }
          // if we got no artifacts back, stop now
        } else {
          getNextPage = false;
        }
      })
    }

    while (getNextPage && currentPage < 100) {
      var startRow = (currentPage * GET_PAGINATION_SIZE) + 1;
      await getSecondaryPage(startRow);
    }

    if (!secondaryArtifact.length) { secondaryArtifact = null; }
    else {
      if (model.currentArtifact.hasSecondaryTarget) {
        //Populate a list of TX owned by the logged user
        var targetIds = []; //List of TX Ids owned by the user
        secondaryArtifact.forEach(function (item, index) {
          //make sure we have only TestSets of this project and also that they are manual type (1)
          if ((item.ProjectId == model.currentProject.id) && (item[model.currentArtifact.secondaryConditionField] == model.currentArtifact.secondaryConditionValue)) {
            targetIds.push(item[model.currentArtifact.SecondaryTypeField]);
          }
        });

        /*
        5.2 Now that we have the filtered IDs of the TestSets, we need to get the Test Cases + Test Steps that are part of them
        */
        currentPage = 0;
        getNextPage = true;
        var destIds = []; //List of Test Case Ids we still need to retrieve

        for (var i = 0; i < targetIds.length; i++) {

          await getArtifacts(
            model.user,
            model.currentProject.id,
            model.currentArtifact.SecondaryTypeId,
            0,
            GET_PAGINATION_SIZE,
            targetIds[i],
            null
          ).then(function (response) {
            // if we got a non empty array back then we have artifacts to process
            if (response.body && response.body.length) {
              destIds = destIds.concat(response.body);

            }
          })
        }
        // 5.3 Finally, getting the Test Cases + Test Steps from the Test Sets

        currentPage = 0;
        getNextPage = true;
        var artifacts2 = [];

        for (var j = 0; j < destIds.length; j++) {
          await getArtifacts(
            model.user,
            model.currentProject.id,
            model.currentArtifact.secondaryTargetId,
            0,
            GET_PAGINATION_SIZE,
            destIds[j][model.currentArtifact.SecondaryTargetFieldName]
          ).then(function (response) {
            // add the response to the object as well as the association field
            if (response.body) {
              //only add this if passes the validation - for TestCases, if it have steps
              if (response.body[model.currentArtifact.conditionField]) {
                response.body = { ...response.body, ...destIds[j] };
                artifacts2 = artifacts2.concat(response.body);
              }
            }
          })

        }
      }

      // if there were no artifacts at all break out now
      if (!artifacts2.length) artifacts2 = null;
      else {
        // if artifact has subtype that needs to be retrieved separately, do so
        if (model.currentArtifact.hasSubType) {
          // find the id field
          var idFieldNameArray = model.fields.filter(function (field) {
            return field.type === fieldTypeEnums.id;
          });

          // 5.4 Retrieving extra information to the Test Cases (from TXs): the TX ReleaseId

          if (model.currentArtifact.hasExtraField) {
            artifacts2.forEach(function (item, index) {
              secondaryArtifact.forEach(function (item2, index2) {
                //If TestCase TestSetId = Test Set TestSetId
                if (item[model.currentArtifact.SecondaryTypeField] == item2[model.currentArtifact.SecondaryTypeField]) {
                  //get the ReleaseId from Test Set object and copy it to the TestCase object
                  artifacts2[index][model.currentArtifact.extraFieldName] = secondaryArtifact[index2][model.currentArtifact.extraFieldName];
                }
              });
            });
          }

          //getting Test Steps (via Shell)
          // if we have an id field, then we can find the id number for each artifact in the array
          if (idFieldNameArray && idFieldNameArray[0].field) {
            //function called below in the foreach call
            async function getArtifactSubs(art) {
              await getArtifacts(
                model.user,
                art.ProjectId,
                model.currentArtifact.subTypeId,
                null,
                null,
                art[params.specialFields.secondaryShellField],
                params.artifactEnums.testSets
              ).then(function (response) {
                // take action if we got any sub types back - ie if they exist for the specific artifact
                if (response.body && response.body.length) {

                  for (var i = 0; i < response.body.length; i++) {

                    if (response.body[i][params.specialFields.testRunStepsField] && response.body[i][params.specialFields.standardTxTsLink]) {
                      //get the correct Tsteps for this TestCaseTestSet
                      if (response.body[i][params.specialFields.standardTxTsLink] == art[params.specialFields.standardTxTsLink]) {
                        var subTypeArtifactsWithMeta = response.body[i][params.specialFields.testRunStepsField].map(function (sub) {
                          sub.isSubType = true;
                          sub.parentId = art[params.specialFields.standardShellField];
                          return sub;
                        })
                        // now add the steps into the original object
                        artifactsWithSubTypes = artifactsWithSubTypes.concat(subTypeArtifactsWithMeta);
                      }
                    }
                  }
                }
              })
            };
            artifactsWithSubTypes = [];

            var idFieldName = idFieldNameArray[0].field;
            //Cases que vem de Sets
            for (var i = 0; i < artifacts2.length; i++) {
              artifactsWithSubTypes.push(artifacts2[i]);
              await getArtifactSubs(artifacts2[i]);
            }

            // update the original array
            artifacts2 = artifactsWithSubTypes;
          }
        }
        //Then, filter the objects again to make sure they belong to the current project
        artifacts = [...artifacts, ...artifacts2];
        if (!artifacts) return 'noData';
      }
      if (!artifacts.length) return 'noData';
    }
  }
  else {
    //if the artifact is the one we want (not a secondary), add it to the list
    //for the TestRun application, we never reach this condition, since it is a sum of objects (TC+TS+TX)
    artifacts = artifacts.concat(secondaryArtifact);
  }
  return artifacts;
}



// EXCEL SPECIFIC to process all the data retrieved from Spira and then display it
// @param: artifacts: array of raw data from Spira (with subtypes already present if needed)
// @param: model: full model object from client
// @param: enum object of the different fieldTypeEnums
function processDataFromSpiraExcel(artifacts, model, fieldTypeEnums) {

  // 5. create 2d array from data to put into sheet
  var artifactsAsCells = matchArtifactsToFields(
    artifacts,
    model.currentArtifact,
    model.fields,
    fieldTypeEnums,
    model.projectUsers,
    model.projectComponents,
    model.projectReleases
  );

  // 6. add data to sheet
  return Excel.run({ delayForCellEdit: true }, function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet(),
      range = sheet.getRangeByIndexes(1, 0, artifacts.length, model.fields.length);
    range.values = artifactsAsCells;

    // 7. Make specific cells read-only 
    artifactsAsCells.forEach(function (row, rowCounter) {

      var isTestCase = (row[TC_ID_COLUMN_INDEX] != '' && row[TC_ID_COLUMN_INDEX] != '-1');

      row.forEach(function (col, colCounter) {
        if (isTestCase) {
          //Make test steps cells read-only for Test Case rows  
          if (model.fields[colCounter].isSubTypeField) {
            protectCell(sheet, colCounter, rowCounter, model.colors.bgReadOnly, "Not available for TestCases.");
          }
        }
        else {
          //release not available for Test Steps
          if (model.fields[colCounter].type == params.fieldType.release) {
            protectCell(sheet, colCounter, rowCounter, model.colors.bgReadOnly, "Not available for TestSteps.");
          }
        }
      });
      rowCounter++;
    });

    // 8.Make other cells protected (the user can't add new data)
    protectRow(sheet, model.fields.length, artifactsAsCells.length, EXCEL_MAX_ROWS, model.colors.bgReadOnly, "Can't add new data.");

    //9.setting cell borders
    var rangeBorder = sheet.getRangeByIndexes(0, 0, EXCEL_MAX_ROWS, model.fields.length);
    rangeBorder = setRangeBorders(rangeBorder, model.colors.cellBorder);


    return context.sync()
      .then(function () {
        return artifactsAsCells;
      })
  })

}

//Function that formats the cell range to have the standard read-only add-in style
//@param rangeBorder - a cell range to apply the changes
//@return - the formatted cell range
function setRangeBorders(rangeBorder, color) {

  rangeBorder.format.borders.getItem('InsideHorizontal').weight = "Thin";
  rangeBorder.format.borders.getItem('InsideHorizontal').color = color;

  rangeBorder.format.borders.getItem('InsideVertical').weight = "Thin";
  rangeBorder.format.borders.getItem('InsideVertical').color = color;

  rangeBorder.format.borders.getItem('EdgeBottom').weight = "Thin";
  rangeBorder.format.borders.getItem('EdgeBottom').color = color;

  rangeBorder.format.borders.getItem('EdgeLeft').weight = "Thin";
  rangeBorder.format.borders.getItem('EdgeLeft').color = color;

  rangeBorder.format.borders.getItem('EdgeRight').weight = "Thin";
  rangeBorder.format.borders.getItem('EdgeRight').color = color;

  rangeBorder.format.borders.getItem('EdgeTop').weight = "Thin";
  rangeBorder.format.borders.getItem('EdgeTop').color = color;

  return rangeBorder;
}


// matches data against the fields to be shown in the spreadsheet - not all data fields are shown
// @param: artifacts - array of the artifact objects we GOT from Spira
// @param: artifactMeta - object of the meta information about the artifact
// @param: fields - array of the fields that make up the sheet display
// @param: fieldTypeEnums - enum object of the different fieldTypes
// @param: users - array of the user objects
// @param: components - array of the component objects
// @param: releases - array of the release objects
function matchArtifactsToFields(artifacts, artifactMeta, fields, fieldTypeEnums, users, components, releases) {
  return artifacts.map(function (art) {
    return fields.map(function (field) {
      var originalFieldValue = "";
      // handle custom fields
      if (field.isCustom && !art.isSubType) {
        // if we have any custom props
        if (art.CustomProperties && art.CustomProperties.length) {
          // look for a match for the current field
          var customProp = art.CustomProperties.filter(function (custom) {
            return custom.Definition.CustomPropertyFieldName == field.field;
          });
          // if the property exists and isn't null - do a null check to handle booleans properly
          if (typeof customProp != "undefined" && customProp.length && customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]] !== null) {
            originalFieldValue = customProp[0][CUSTOM_PROP_TYPE_ENUM[field.type]];
          }
        }

        // handle subtype fields
      } else if (field.isSubTypeField) {
        if (artifactMeta.hasSubType && art.isSubType) {
          // first check to make sure the field exists in the artifact data
          if (typeof art[field.field] != "undefined" && art[field.field]) {
            originalFieldValue = art[field.field]
          }
        }

        // handle standard fields
      } else if (!art.isSubType) {
        // first check to make sure the field exists in the artifact data
        if (typeof art[field.field] != "undefined" && art[field.field]) {
          originalFieldValue = art[field.field];
        }
      }
      // handle list values - turn from the id to the actual string so the string can be displayed
      if (
        field.type == fieldTypeEnums.drop ||
        field.type == fieldTypeEnums.multi ||
        field.type == fieldTypeEnums.user ||
        field.type == fieldTypeEnums.component ||
        field.type == fieldTypeEnums.release
      ) {
        // a field can have display overrides - if one of these overrides is in the artifact field specified, then this is returned instead of the lookup - used specifically to make sure RQ Epics show as Epics 
        if (field.displayOverride && field.displayOverride.field && field.displayOverride.values && field.displayOverride.values.includes(art[field.displayOverride.field])) {
          return art[field.displayOverride.field];
        } else {
          // handle multilist fields (custom props or components for some artifacts) - we can only display one in Excel so pick the first in the array to match
          var fieldValueForLookup = Array.isArray(originalFieldValue) ? originalFieldValue[0] : originalFieldValue;
          var fieldName = getListValueFromId(
            fieldValueForLookup,
            field.type,
            fieldTypeEnums,
            field.values,
            users,
            components,
            releases
          );

          if (fieldName && fieldValueForLookup) {
            fieldName = fieldName + " (#" + fieldValueForLookup + ")";
          }

          return fieldName;

        }

        // handle date fields 
      } else if (field.type == fieldTypeEnums.date) {
        if (originalFieldValue) {
          var jsObj = new Date(originalFieldValue);
          return JSDateToExcelDate(jsObj);
        } else {
          return "";
        }

        // handle booleans - need to make sure null values are ignored ie treated differently to false
      } else if (field.type == fieldTypeEnums.bool) {
        return originalFieldValue ? "Yes" : originalFieldValue === false ? "No" : "";
        // handle hierarchical artifacts
      } else if (field.setsHierarchy) {
        return makeHierarchical(originalFieldValue, art.IndentLevel);
        // handle artifacts that have extra information we can display to the user - ie where there is a linked test step this will add the information about the link at the end of the field
      } else if (field.extraDataField && art[field.extraDataField]) {
        return `${originalFieldValue} ${field.extraDataPrefix ? field.extraDataPrefix + ":" : ""}${art[field.extraDataField]}`;
      } else {
        return originalFieldValue;

      }
    });
  })
}

// takes an id for a lookup field and returns the string to display
// @param: id - int of the id to lookup
// @param: type - enum of the type of filed we need to look up
// @param: fieldTypeEnums - enum object of the different fieldTypes
// @param: fieldValues - array of the value objects for bespoke lookups
// @param: users - array of the user objects
// @param: components - array of the component objects
// @param: releases - array of the release objects
function getListValueFromId(id, type, fieldTypeEnums, fieldValues, users, components, releases) {
  var match = null;
  switch (type) {
    case fieldTypeEnums.drop:
    case fieldTypeEnums.multi:
      match = fieldValues.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.user:
      match = users.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.component:
      match = components.filter(function (val) { return val.id == id; });
      break;
    case fieldTypeEnums.release:
      match = releases.filter(function (val) { return val.id == id; });
      break;
  }
  return typeof match != "undefined" && match && match.length ? match[0].name : "";
}

function makeHierarchical(value, indent) {
  var indentIncrements = Math.floor(indent.length / 3) - 1;
  var indentText = "";
  for (var i = 0; i < indentIncrements; i++) {
    indentText += "> ";
  }
  indentText += value;
  return indentText;
}

//Convert JS date object to excel format - Excel dates start at 1900 not 1970
//@param: inDate - js date object
function JSDateToExcelDate(inDate) {
  var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
  return returnDateTime.toString().substr(0, 20);
}