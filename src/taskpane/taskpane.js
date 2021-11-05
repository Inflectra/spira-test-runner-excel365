/*
 *
 * =============
 * GENERAL SETUP
 * =============
 *
 */

// model becomes a new instance of the data store preserving the immutability of the primary data object.
var model = new Data();
var uiSelection = new tempDataStore();

// if devmode enabled, set the required fields and show the dev button
var devMode = true;
var isGoogle = false;

//global flag for the Get Last TS run status check box
var lastStatusMode = true;

/*
Global Variable to control if advanced options should be enabled to the user
Up to know, the advanced features are :
1."New Comment" field to all the artifacts -> allow creating new comments in Spira
2. Create new Artifacts Association:
  a. TestCase: Requirements, Releases and TestSet
  b. Requirements: Requirents
*/
var advancedMode = false;

//ENUMS

var UI_MODE = {
  newProject: 1,
  newArtifact: 2,
  getData: 3,
  errorMode: 4
};

/*
 *
 * ============================
 * GOOGLE SHEETS SPECIFIC SETUP
 * ============================
 *
 */

// Google Sheets specific code to run at first launch
(function () {
  if (typeof google != "undefined") {
    isGoogle = true;
    // for dev mode only - comment out or set to false to disable any UI dev features
    setDevStuff(devMode);

    // add event listeners to the dom
    setEventListeners();

    // dom specific changes
    document.getElementById("help-connection-excel").style.display = "none";
  }
})();








/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 *
 */
import { params, templateFields, Data, tempDataStore } from './model.js';
import * as msOffice from './server.js';

export { showPanel, hidePanel };


// MS Excel specific code to run at first launch
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // on init make sure to run any required startup functions
    setEventListeners();
    // for dev mode only - comment out or set to false to disable any UI dev features
    setDevStuff(devMode);

    // dom specific changes
    document.body.classList.add('ms-office');
    document.getElementById("help-connection-google").style.display = "none";
  }
});










/*
 *
 * =================================
 * UTILITIES & CROSS PANEL FUNCTIONS
 * =================================
 *
 */

function setDevStuff(devMode) {
  if (devMode) {
    document.getElementById("btn-dev").classList.remove("hidden");
    model.user.url = "";
    model.user.userName = "administrator";
    model.user.api_key = btoa("&api-key=" + encodeURIComponent("{}"));

    loginAttempt();
  }
}


function setEventListeners() {
  document.getElementById("btn-login").onclick = loginAttempt;
  document.getElementById("btn-help-login").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('login');
  };
  document.getElementById("lnk-help-login").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('login');
  };
  document.getElementById("btn-dev").onclick = setAuthDetails;

  document.getElementById("btn-help-main").onclick = function () {
    panelToggle('help');
    showChosenHelpSection('data')
  };

  document.getElementById("btn-logout").onclick = logoutAttempt;
  document.getElementById("btn-main-back").onclick = hideMainPanel;

  // changing of dropdowns
  document.getElementById("select-product").onchange = changeProjectSelect;
  document.getElementById("select-artifact").onchange = changeArtifactSelect;

  document.getElementById("btn-toSpira").onclick = sendToSpiraAttempt;
  document.getElementById("btn-fromSpira").onclick = getFromSpiraAttempt;
  document.getElementById("btn-template").onclick = updateTemplateAttempt;
  document.getElementById("btn-updateToSpira").onclick = updateSpiraAttempt;

  document.getElementById("btn-help-back").onclick = function () { panelToggle("help") };
  document.getElementById("btn-help-section-login").onclick = function () { showChosenHelpSection('login') };
  document.getElementById("btn-help-section-getting").onclick = function () { showChosenHelpSection('getting') };
  document.getElementById("btn-help-section-sending").onclick = function () { showChosenHelpSection('sending') };
  document.getElementById("chkLastStatus").onclick = setLastStatus;
}

// switches the value of the global variable for the Advanced Mode
function setLastStatus() {
  if (document.getElementById('chkLastStatus').checked) {
    lastStatusMode = true;
  } else {
    lastStatusMode = false;
  }
}

// used to show or hide / hide / show a specific panel
// @param: panel - string. suffix for items to act on (eg if id = panel-help, choice = "help")
function panelToggle(panel) {
  var panelElement = document.getElementById("panel-" + panel);
  panelElement.classList.toggle("offscreen");
}


function hidePanel(panel) {
  var panelElement = document.getElementById("panel-" + panel);
  panelElement.classList.add("offscreen");
}



function showPanel(panel) {
  var panelElement = document.getElementById("panel-" + panel);
  panelElement.classList.remove("offscreen");
}



// manage the loading spinner
function showLoadingSpinner() {
  document.getElementById("loader").classList.remove("hidden");
}



function hideLoadingSpinner() {
  document.getElementById("loader").classList.add("hidden");
}



// clear spreadsheet, model
function clearAddonData() {
  model = new Data();
  uiSelection = new tempDataStore();
  setDevStuff(devMode);
}



// clears the first sheet in the book
// @param: shouldClear - optional bool to check
function clearSheet(shouldClear) {
  var shouldClearToUse = typeof shouldClear !== 'undefined' ? shouldClear : true;
  if (shouldClearToUse) {
    if (isGoogle) {
      google.script.run.clearAll();
    } else {
      msOffice.clearAll()
        .then((response) => document.getElementById("panel-confirm").classList.add("offscreen"))
        .catch((error) => errorExcel(error));
    }
  }
}



// resets the sidebar following logout
function resetSidebar() {
  // clear input field values
  document.getElementById("input-url").value = "";
  document.getElementById("input-userName").value = "";
  document.getElementById("input-password").value = "";

  // hide other panels, so login page is visible
  var otherPanels = document.querySelectorAll(".panel:not(#panel-auth)");
  // can't use forEach because that is not supported by Excel
  for (var i = 0; i < otherPanels.length; ++i) {
    otherPanels[i].classList.add("offscreen");
  }

  resetUi();
  // reset anything required if in devmode
  setDevStuff();
}

function resetUi() {
  // disable buttons and dropdowns
  document.getElementById("btn-template").disabled = true;
  document.getElementById("pnl-template").style.display = "none";
  document.getElementById("select-artifact").disabled = true;
  document.getElementById("btn-fromSpira").disabled = true;
  document.getElementById("btn-toSpira").disabled = true;
  document.getElementById("btn-updateToSpira").disabled = true;
  document.getElementById("btn-toSpira").innerHTML = "Prepare Template";

  // reset artifact dropdown to 'Select an Artifact'
  document.getElementById("select-artifact").selectedIndex = "0";
  // hide and clear the template info box
  document.getElementById("template-project").textContent = "";
  document.getElementById("template-artifact").textContent = "";

  // reset action buttons
  document.getElementById("btn-fromSpira").style.display = "";
  document.getElementById("btn-toSpira").style.display = "";
  document.getElementById("btn-updateToSpira").style.display = "";


  // reset guide text on the main pane
  document.getElementById("main-guide-1").classList.remove("pale");
  document.getElementById("main-guide-1-fromSpira").style.display = "";
  document.getElementById("main-guide-1-toSpira").style.display = "";
  document.getElementById("main-heading-fromSpira").style.display = "";
  document.getElementById("main-heading-toSpira").style.display = "";
  document.getElementById("main-guide-2").classList.add("pale");
  document.getElementById("main-guide-3").classList.add("pale");
}



// adds all options to a dropdown
// @param: selectId - is the id of the dom select element
// @param: array - the array of objects (with id, name, and optionally a disabled value, and hidden bool)
// @param: firstMessage - an optional text field to go at the top of the array - the initial choice 
function setDropdown(selectId, array, firstMessage) {
  // first make a deep copy of the array to stop any funny business
  var arrayCopy = JSON.parse(JSON.stringify(array)),
    select = document.getElementById(selectId);
  // if passed in, add default "select" option to top of project array
  if (firstMessage) arrayCopy.unshift({
    id: 0,
    name: firstMessage
  });
  // clear the dropdown
  select.innerHTML = "";
  arrayCopy.forEach(function (item) {
    var option = document.createElement("option");
    option.disabled = item.disabled;
    option.value = item.id;
    option.innerHTML = item.name;

    if (!item.hidden) {
      select.appendChild(option);
    }

  });
}



function isModelDifferentToSelection() {
  if (model.isTemplateLoaded) {
    var projectHasChanged = model.currentProject.id !== getSelectedProject().id;
    var artifactHasChanged = model.currentArtifact.id !== params.artifacts[0].id;
    return projectHasChanged || artifactHasChanged;
  } else {
    return false;
  }
}









/*
*
* ============
* LOGIN SCREEN
* ============
*
*/

// get user data from input fields and store in user data object
// adds the 'api-key' text before the key to make creating the urls simpler
function getAuthDetails() {
  model.user.url = document.getElementById("input-url").value;
  model.user.userName = document.getElementById("input-userName").value;
  var password = document.getElementById("input-password").value;
  model.user.api_key = btoa("&api-key=" + encodeURIComponent(password));
}



// fill in mock values for easy log in development, enable dev button
function setAuthDetails() {
  document.getElementById("input-url").value = model.user.url;
  document.getElementById("input-userName").value = model.user.userName;
  document.getElementById("input-password").value = model.user.password;
}


// handle the click of the login button
function loginAttempt() {
  if (!devMode) getAuthDetails();
  login();
}

// login function that starts the intial data creation
function login() {
  showLoadingSpinner();
  // call server side function to get projects
  // also serves as authentication check, if the user credentials aren't correct it will throw a network error
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(populateProjects)
      .withFailureHandler(errorNetwork)
      .getProjects(model.user);
  } else {
    msOffice.getProjects(model.user)
      .then(response => populateProjects(response.body))
      .catch(err => {
        return errorNetwork(err)
      }
      );
  }
}



// kick off prepping and showing main panel
// @param: projects - passed in projects data returned from the server following successful API call to Spira
function populateProjects(projects) {
  // take projects data from Spira API call, strip out unwanted fields, add to data model
  var pairedDownProjectsData = projects.map(function (project) {
    var result = {
      id: project.ProjectId,
      name: project.Name,
      templateId: project.ProjectTemplateId
    };
    return result;
  });

  // now add paired down project array to data store
  model.projects = pairedDownProjectsData;

  // get UI logic ready for decision panel
  showMainPanel('run');
  hideLoadingSpinner();
}









/*
*
* ===========
* MAIN SCREEN
* ===========
*
*/

// manage the switching of the UI off the login screen on succesful login and retrieval of projects
function showMainPanel(type) {

  setDropdown("select-product", model.projects, "Select a product");
  setDropdown("select-artifact", params.artifacts, "Select an artifact");

  // set the buttons to the correct mode
  if (type == "send") {
    document.getElementById("btn-fromSpira").style.display = "none";
    document.getElementById("main-guide-1-fromSpira").style.display = "none";
    document.getElementById("main-heading-fromSpira").style.display = "none";
    document.getElementById("btn-updateToSpira").style.visibility = "hidden";
    document.getElementById("main-guide-3").style.visibility = "hidden";
  } else {
    document.getElementById("btn-toSpira").style.display = "none";
    document.getElementById("main-guide-1-toSpira").style.display = "none";
    document.getElementById("main-heading-toSpira").style.display = "none";
    document.getElementById("main-guide-3").style.visibility = "visible";
    document.getElementById("btn-updateToSpira").style.visibility = "visible";
  }

  // opens the panel
  showPanel("main");
  hideLoadingSpinner();
  //
  document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'bold';
}


function hideMainPanel() {
  hidePanel("main");
  // reset the buttons and dropdowns
  resetUi();
  uiSelection = new tempDataStore();
  // make sure the system does not think any data is loaded
  model.isTemplateLoaded = false;
}



// run server side code to manage logout
function logoutAttempt() {
  var message = 'All data on the active sheet will be deleted. Continue?'
  //warn user that all data on the first sheet will be lost. Returns true or false
  showPanel("confirm");
  document.getElementById("message-confirm").innerHTML = message;
  document.getElementById("btn-confirm-ok").onclick = () => logout(true);
  document.getElementById("btn-confirm-cancel").onclick = () => hidePanel("confirm");
}



// @param: shouldLogout - a true or false value from google/Excel
function logout(shouldLogout) {
  if (shouldLogout) {
    clearAddonData();
    resetSidebar();
  }
}



function changeProjectSelect(e) {
  //sets the UI to correspond to this mode
  artifactUpdateUI(UI_MODE.newProject);

  // if the project field has not been selected all other selected buttons are disabled
  if (e.target.value == 0) {
    document.getElementById("select-artifact").disabled = true;
    document.getElementById("btn-toSpira").disabled = true;
    document.getElementById("btn-fromSpira").disabled = true;
    document.getElementById("btn-template").disabled = true;
    uiSelection.currentProject = null;
  } else {
    // get the project object and update project information if project has changed
    var chosenProject = getSelectedProject();
    if (chosenProject.id && chosenProject.id !== uiSelection.currentProject.id) {
      //set the temp data store project to the one selected;
      uiSelection.currentProject = chosenProject;

      //manageTemplateBtnState();
      // kick off API calls
      getProjectSpecificInformation(model.user, uiSelection.currentProject.id);

      // for 6.1 the v6 API for get projects does not get the project template IDs so have to do this
      getTemplateFromProjectId(model.user, uiSelection.currentProject.id, uiSelection.currentArtifact);

      // since we already have our artifact selected (test runs), proceed the operations:
      // get the artifact object and update artifact information if artifact has changed
      var chosenArtifact = params.artifacts[0];

      //set the temp date store artifact to the one selected;
      uiSelection.currentArtifact = chosenArtifact;

      // enable template button only when all info is received - otherwise keep it disabled
      manageTemplateBtnState();

      // kick off API calls - if we have a current template and project
      if (uiSelection.currentProject.templateId && uiSelection.currentProject.id) {
        getArtifactSpecificInformation(model.user, uiSelection.currentProject.templateId, uiSelection.currentProject.id, uiSelection.currentArtifact);
      }


      document.getElementById('main-guide-2').style.fontWeight = 'bold';
      document.getElementById('main-guide-1-fromSpira').style.fontWeight = 'normal';
    }
  }
}

function getTemplateFromProjectId(user, projectId, artifact) {
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getArtifactSpecificInformationInterim)
      .withFailureHandler(errorNetwork)
      .getTemplateFromProjectId(user, projectId);
  } else {
    msOffice.getTemplateFromProjectId(user, projectId)
      .then((response) => getArtifactSpecificInformationInterim(response.body))
      .catch((error) => errorNetwork(error));
  }

  function getArtifactSpecificInformationInterim(template) {
    uiSelection.currentProject.templateId = template.ProjectTemplateId;
    if (uiSelection.currentArtifact) {
      getArtifactSpecificInformation(model.user, template.ProjectTemplateId, uiSelection.currentProject.id, uiSelection.currentArtifact)
    }
  }
}

//handles hiding/displaying and changing colors of elements in the UI based on the operation
function artifactUpdateUI(mode) {

  switch (mode) {

    case UI_MODE.newProject:
      //when selecting a new project
      document.getElementById("main-guide-1").classList.remove("pale");
      document.getElementById("main-guide-2").classList.add("pale");
      document.getElementById("main-guide-3").classList.add("pale");
      document.getElementById("btn-fromSpira").disabled = true;
      document.getElementById("btn-updateToSpira").disabled = true;
      break;

    case UI_MODE.newArtifact:
      //when selecting a new artifact
      document.getElementById("main-guide-3").classList.add("pale");
      document.getElementById("btn-updateToSpira").disabled = true;
      break;

    case UI_MODE.getData:
      //when clicking from-Spira button
      document.getElementById("btn-updateToSpira").disabled = false;
      break;

    case UI_MODE.errorMode:
      //in case of any error
      document.getElementById("main-guide-2").classList.remove("pale");
      document.getElementById("btn-fromSpira").disabled = false;
      document.getElementById("main-guide-3").classList.add("pale");
      document.getElementById("btn-updateToSpira").disabled = true;
      break;
  }
}


function changeArtifactSelect(e) {
  //sets the UI to correspond to this mode
  artifactUpdateUI(UI_MODE.newArtifact);
  if (e.target.value == 0) {
    document.getElementById("btn-toSpira").disabled = true;
    document.getElementById("btn-fromSpira").disabled = true;
    document.getElementById("btn-template").disabled = true;
    uiSelection.currentArtifact = null;
  } else {
    // get the artifact object and update artifact information if artifact has changed
    var chosenArtifact = getSelectedArtifact();

    if (chosenArtifact !== uiSelection.currentArtifact) {
      //set the temp date store artifact to the one selected;
      uiSelection.currentArtifact = chosenArtifact;

      // enable template button only when all info is received - otherwise keep it disabled
      manageTemplateBtnState();

      // kick off API calls - if we have a current template and project
      if (uiSelection.currentProject.templateId && uiSelection.currentProject.id) {
        getArtifactSpecificInformation(model.user, uiSelection.currentProject.templateId, uiSelection.currentProject.id, uiSelection.currentArtifact);
      }
    }
  }
}



// disables and enables the main action buttons based on status of required API calls
function manageTemplateBtnState() {
  // initially disable the button, because required API calls not completed
  document.getElementById("btn-toSpira").disabled = true;
  document.getElementById("btn-fromSpira").disabled = true;
  document.getElementById("btn-template").disabled = true;

  // only try to enable the button when both a project and artifact have been chosen
  if (uiSelection.currentProject && uiSelection.currentArtifact) {

    // set a function to run repeatedly until all gets are done
    // then enable the button, and stop the timer loop
    var checkGetsSuccess = setInterval(updateButtonStatus, 500);
    // and show a message while api calls are underway
    document.getElementById("message-fetching-data").style.visibility = "visible";

    function updateButtonStatus() {
      if (allGetsSucceeded()) {
        if (!document.getElementById("btn-updateToSpira").disabled) {
          //sets the UI to allow update
          document.getElementById("btn-fromSpira").disabled = false;
          document.getElementById("main-guide-2").classList.add("pale");
          document.getElementById("main-guide-3").classList.remove("pale");
          document.getElementById("checkboxStatus").classList.add("pale");
          document.getElementById("message-fetching-data").style.visibility = "hidden";
        }
        else {
          document.getElementById("btn-toSpira").disabled = false;
          document.getElementById("btn-fromSpira").disabled = false;

          document.getElementById("message-fetching-data").style.visibility = "hidden";
          document.getElementById("main-guide-1").classList.add("pale");
          document.getElementById("main-guide-2").classList.remove("pale");
          document.getElementById("checkboxStatus").classList.remove("pale");
          document.querySelector(".circleHelp").classList.add("checked");
        }

        clearInterval(checkGetsSuccess);

        // if there is a discrepancy between the dropdown and the currently active template
        // only do this is user will send to spira - determined by whether the send to spira button is visible
        if (document.getElementById("btn-toSpira").style.display != "none") {
          if (isModelDifferentToSelection()) {
            document.getElementById("pnl-template").style.display = "";
            document.getElementById("btn-template").disabled = false;
          } else {
            document.getElementById("pnl-template").style.display = "none";
            document.getElementById("btn-template").disabled = true;
          }
        }
      }
      else {
      }
    }
  }
}



// starts the process to create a template from chosen options
function createTemplateAttempt() {
  var message = 'The active sheet will be replaced. Continue?'
  //warn the user data will be erased
  showPanel("confirm");
  document.getElementById("message-confirm").innerHTML = message;
  document.getElementById("btn-confirm-ok").onclick = () => createTemplate(true);
  document.getElementById("btn-confirm-cancel").onclick = () => hidePanel("confirm");
}



//@param: shouldClearForm - boolean response from confirmation dialog above
function createTemplate(shouldContinue) {
  if (shouldContinue) {
    clearSheet();
    showLoadingSpinner();
    manageTemplateBtnState();

    // all data should already be loaded (as otherwise template button is disabled)
    // but check again that all data is present before kicking off template creation
    // if so, kicks off template creation, otherwise waits and tries again
    if (allGetsSucceeded()) {
      templateLoader();

      // otherwise, run an interval loop (should never get called as template button should be disabled)
    } else {
      var checkGetsSuccess = setInterval(attemptTemplateLoader, 500);
      function attemptTemplateLoader() {
        if (allGetsSucceeded()) {
          templateLoader();
          clearInterval(checkGetsSuccess);
        }
      }
    }
  }
}

function getFromSpiraAttempt() {
  // first update state to reflect user intent
  model.isGettingDataAttempt = true;

  //check that template is loaded and that it matches the UI choices
  if (model.isTemplateLoaded && !isModelDifferentToSelection()) {
    showLoadingSpinner();
    //call export function
    if (isGoogle) {
      google.script.run
        .withFailureHandler(errorImpExp)
        .withSuccessHandler(getFromSpiraComplete)
        .getFromSpiraGoogle(model, params.fieldType);
    } else {
      msOffice.getFromSpiraExcel(model, params.fieldType, lastStatusMode)
        .then((response) => getFromSpiraComplete(response))
        .catch((error) => errorImpExp(error));
    }
  } else {
    //if no template - then get the template
    createTemplateAttempt();
  }
  //sets the UI to correspond to this mode
  artifactUpdateUI(UI_MODE.getData);
}

function getFromSpiraComplete(log) {
  if (devMode) console.log(log);
  //if array (which holds error responses) is present, and errors present
  if (log && log.errorCount) {
    errorMessages = log.entries
      .filter(function (entry) { return entry.error; })
      .map(function (entry) { return entry.message; });
  }
  else {
    manageTemplateBtnState();
  }
  hideLoadingSpinner();

  //runs the export success function, passes a boolean flag, if there are errors the flag is true.
  if (log && log.status) {
    if (isGoogle) {
      google.script.run.operationComplete(log.status);
    } else {
      msOffice.operationComplete(log.status);
    }
  }

  document.getElementById('main-guide-2').style.fontWeight = 'normal';
  document.getElementById('main-guide-3').style.fontWeight = 'bold';

  document.getElementById('btn-fromSpira').classList.remove('ms-Button--primary');
  document.getElementById('btn-fromSpira').classList.add('ms-Button--default');

  document.getElementById('btn-updateToSpira').classList.add('ms-Button--primary');
  document.getElementById('btn-updateToSpira').classList.remove('ms-Button--default');

}



function sendToSpiraAttempt() {
  // first update state to reflect user intent
  model.isGettingDataAttempt = false;
  //check that template is loaded
  if (model.isTemplateLoaded) {
    showLoadingSpinner();

    //call export function
    if (isGoogle) {
      google.script.run
        .withFailureHandler(errorImpExp)
        .withSuccessHandler(sendToSpiraComplete)
        .sendToSpira(model, params.fieldType);
    } else {
      msOffice.sendToSpira(model, params.fieldType)
        .then((response) => sendToSpiraComplete(response))
        .catch((error) => errorImpExp(error));
    }
  } else {
    //if no template - then get the template
    createTemplateAttempt();
  }
}

function updateSpiraAttempt() {

  // first update state to reflect user intent
  model.isGettingDataAttempt = false;
  //check that template is loaded
  if (model.isTemplateLoaded) {
    showLoadingSpinner();

    //call export function
    if (isGoogle) {
      google.script.run
        .withFailureHandler(errorImpExp)
        .withSuccessHandler(sendToSpiraComplete)
        .sendToSpira(model, params.fieldType, true);
    } else {
      msOffice.sendToSpira(model, params.fieldType)
        .then((response) => sendToSpiraComplete(response))
        .catch((error) => errorImpExp(error));
    }
  } else {
    //if no template - throw an error
    errorExcel("The spreadsheet does not match the selected artifact. Please check your data.")
  }

}

function sendToSpiraComplete(log) {
  hideLoadingSpinner();
  if (devMode) console.log(log);

  if (log) {
    //if array (which holds error responses) is present, and errors present
    if (log.errorCount) {
      var errorMessages = log.entries
        .filter(function (entry) { return entry.error; })
        .map(function (entry) { return entry.message; });

    }
    //runs the export success function, passes a boolean flag, if there are errors the flag is true.
    if (log && log.status) {
      if (isGoogle) {
        google.script.run.operationComplete(log.status);
      } else {
        msOffice.operationComplete(log.status);
      }
    }
  }
}



function updateTemplateAttempt() {
  // first update state to reflect user intent
  model.isGettingDataAttempt = false;
  createTemplateAttempt();
}








/*
*
* ===========
* HELP SCREEN
* ===========
*
*/
// manage showing the correct help section to the user
// @param: choice - string. suffix for items to select (eg if id = help-section-fields, choice = "fields")
function showChosenHelpSection(choice) {
  // does not use a dynamic list using queryselectorall and node list because Excel does not support this
  // hide all sections and then only show the one the user wants
  document.getElementById("help-section-login").classList.add("hidden");
  document.getElementById("help-section-getting").classList.add("hidden");
  document.getElementById("help-section-sending").classList.add("hidden");
  document.getElementById("help-section-" + choice).classList.remove("hidden");

  // set all buttons back to normal, then highlight one just clicked
  document.getElementById("btn-help-section-login").classList.remove("create");
  document.getElementById("btn-help-section-getting").classList.remove("create");
  document.getElementById("btn-help-section-sending").classList.remove("create");
  document.getElementById("btn-help-section-" + choice).classList.add("create");
}









/*
*
* =================
* CREATING TEMPLATE
* =================
*
*/

// retrieves the project object that matches the project selected in the dropdown
// returns a project object
function getSelectedProject() {
  // store dropdown value
  var select = document.getElementById("select-product");
  var projectDropdownVal = select.options[select.selectedIndex].value;
  // filter the project lists to those chosen 
  var projectSelected = model.projects.filter(function (project) {
    return project.id == projectDropdownVal;
  })[0];
  return projectSelected;
}



// retrieves the artifact object that matches the artifact selected in the dropdown
// returns an artifact object
function getSelectedArtifact() {
  // store dropdown values
  var select = document.getElementById("select-artifact");
  var artifactDropdownVal = select.options[select.selectedIndex].value;
  // filter the artifact lists to those chosen 
  var artifactSelected = params.artifacts.filter(function (artifact) {
    return artifact.id == artifactDropdownVal;
  })[0];
  return artifactSelected;
}



// kicks off all relevant API calls to get project specific information
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getProjectSpecificInformation(user, projectId) {
  model.projectGetRequestsMade = 0;
  // get project information
  getUsers(user, projectId);
  getComponents(user, projectId);
  getReleases(user, projectId);
}



// kicks off all relevant API calls to get artifact specific information
// @param: user - the user object of the logged in user
// @param: templateId - int of the reqested project template
// @param: artifact - object of the reqested artifact - needed to query on different parts of object
function getArtifactSpecificInformation(user, templateId, projectId, artifact) {
  // first reset get counts
  model.artifactGetRequestsMade = 0;
  model.artifactGetRequestsToMake = model.baselineArtifactGetRequests;
  // increase the count if any bespoke fields are present (eg folders or incident types)
  var bespokeData = fieldsWithBespokeData(templateFields[artifact.field]);
  if (bespokeData) {
    model.artifactGetRequestsToMake += bespokeData.length;
    // get any bespoke field information
    bespokeData.forEach(function (bespokeField) {
      getBespoke(user, templateId, projectId, artifact.field, bespokeField);
    });
  }
  // get standard artifact information - eg custom fields
  getCustoms(user, templateId, artifact.id);
}



// goes through artifact object and returns an array of field objects that have specific rest calls to get their data
// @param: artifact - object of the requested artifact
function fieldsWithBespokeData(artifactFields) {
  if (!artifactFields.length) {
    return;
  }
  var bespokeFields = artifactFields.filter(function (field) {
    return field.bespoke;
  });
  return bespokeFields.length ? bespokeFields : false;
}



// starts GET request to Spira for project components
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getComponents(user, projectId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getComponentsSuccess)
      .withFailureHandler(errorNetwork)
      .getComponents(user, projectId);
  } else {
    msOffice.getComponents(user, projectId)
      .then((response) => getComponentsSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets component data on the model
function getComponentsSuccess(data) {
  // clear old values
  uiSelection.projectComponents = [];
  // add relevant data to the main model store
  uiSelection.projectComponents = data
    .filter(function (item) { return item.IsActive; })
    .map(function (item) {
      return {
        id: item.ComponentId,
        name: item.Name
      };
    });
  model.projectGetRequestsMade++;
}



// starts GET request to Spira for project / artifact custom properties
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
// @param: artifactId - int of the reqested artifact
function getCustoms(user, projectId, artifactId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getCustomsSuccess)
      .withFailureHandler(errorNetwork)
      .getCustoms(user, projectId, artifactId);
  } else {
    msOffice.getCustoms(user, projectId, artifactId)
      .then((response) => getCustomsSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets custom field data on the model - adding to a temp holding area, to allow for changes before template creation
function getCustomsSuccess(data) {
  // clear old values
  uiSelection.artifactCustomFields = [];
  // assign unparsed data to data object
  // these values are parsed later depending on function needs
  uiSelection.artifactCustomFields = data
    .filter(function (item) { return !item.IsDeleted; })
    .map(function (item) {
      var customField = {
        isCustom: true,
        field: item.CustomPropertyFieldName,
        name: item.Name,
        propertyNumber: item.PropertyNumber,
        type: item.CustomPropertyTypeId,
      };

      // mark as required or not - default is that it can be empty
      var allowEmptyOption = item.Options && item.Options.filter(function (option) {
        return option.CustomPropertyOptionId && option.CustomPropertyOptionId === 1;
      });
      if (allowEmptyOption && allowEmptyOption.length && allowEmptyOption[0].Value == "N") {
        customField.required = true;
      }
      // add array of values for dropdowns
      if (item.CustomPropertyTypeId == params.fieldType.drop || item.CustomPropertyTypeId == params.fieldType.multi) {
        customField.values = item.CustomList.Values.map(function (listItem) {
          return {
            id: listItem.CustomPropertyValueId,
            name: listItem.Name
          };
        });
      }
      return customField;
    }
    );
  model.artifactGetRequestsMade++;
}
// starts GET request to Spira for project users properties
// @param: user - the user object of the logged in user
// @param: templateId - int of the reqested template
function getBespoke(user, templateId, projectId, artifactId, field) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getBespokeSuccess)
      .withFailureHandler(errorNetwork)
      .getBespoke(user, templateId, projectId, artifactId, field);
  } else {
    msOffice.getBespoke(user, templateId, projectId, artifactId, field)
      .then((response) => getBespokeSuccess({
        artifactName: artifactId,
        field: field,
        values: response.body
      }))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets user data on the model
function getBespokeSuccess(data) {
  // create and clear old values
  if (typeof uiSelection[data.artifactName] == "undefined") {
    uiSelection[data.artifactName] = {};
  }
  uiSelection[data.artifactName][data.field.field] = [];

  // if there is data take steps to add it to the artifact object
  if (data && data.values && data.values.length) {
    // map through user obj and assign names
    var values = data.values.map(function (item) {
      var obj = {};
      obj.id = item[data.field.bespoke.idField];
      obj.name = item[data.field.bespoke.nameField];
      if (data.field.bespoke.indent) {
        obj.indent = item[data.field.bespoke.indent];
      }
      return obj;
    });

    // indented fields need to specify a field name that contains indent data, so we use this as a check to see if a field is hierarchicel
    if (data.field.bespoke.indent) {
      var hierarchicalValues = values.sort(function (a, b) {
        if (a.indent < b.indent) {
          return -1;
        }
        if (a.indent > b.indent) {
          return 1;
        }
        // names must be equal
        return 0;
      }).map(function (x) {
        var indentAmount = (x.indent.length / 3) - 1;
        var indentString = "...";
        x.name = (indentString.repeat(indentAmount)) + x.name;
        return x;
      });
      uiSelection[data.artifactName][data.field.field].values = hierarchicalValues;

    } else {
      uiSelection[data.artifactName][data.field.field].values = values;
    }
  }
  // in all cases make sure the successful request is recorded
  model.artifactGetRequestsMade++;
}



// starts GET request to Spira for project releases properties
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getReleases(user, projectId, artifactId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getReleasesSuccess)
      .withFailureHandler(errorNetwork)
      .getReleases(user, projectId);
  } else {
    msOffice.getReleases(user, projectId)
      .then((response) => getReleasesSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}
// formats and sets release data on the model
function getReleasesSuccess(data) {
  // clear old values
  uiSelection.projectReleases = [];
  // add relevant data to the main model store
  uiSelection.projectReleases = data.map(function (item) {
    return {
      id: item.ReleaseId,
      name: item.Name
    };
  });
  model.projectGetRequestsMade++;
}



// starts GET request to Spira for project users properties
// @param: user - the user object of the logged in user
// @param: projectId - int of the reqested project
function getUsers(user, projectId) {
  // call server side fetch
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(getUsersSuccess)
      .withFailureHandler(errorNetwork)
      .getUsers(user, projectId);
  } else {
    msOffice.getUsers(user, projectId)
      .then((response) => getUsersSuccess(response.body))
      .catch((error) => errorNetwork(error));
  }
}



// formats and sets user data on the model
function getUsersSuccess(data) {
  // clear old values
  uiSelection.projectUsers = [];
  // map through user obj and assign names
  uiSelection.projectUsers = data.map(function (item) {
    return {
      id: item.UserId,
      username: item.UserName,
      name: item.FullName,
    };
  });
  model.projectGetRequestsMade++;
}



// check to see that all project and artifact requests have been made - ie that successes match required requests
// returns boolean
function allGetsSucceeded() {
  var projectGetsDone = model.projectGetRequestsToMake === model.projectGetRequestsMade,
    //this application uses a 'special' single artifact, so this will always be true
    artifactGetsDone = true;
  return projectGetsDone && artifactGetsDone;
}



// send data to server to manage the creation of the template on the relevant sheet
function templateLoader() {
  // set the model based on data stored based on current dropdown selections
  model.currentProject = uiSelection.currentProject;
  model.currentArtifact = uiSelection.currentArtifact;
  model.projectComponents = [];
  model.projectReleases = [];
  model.projectUsers = [];
  model.projectComponents = uiSelection.projectComponents;
  model.projectReleases = uiSelection.projectReleases;
  model.projectUsers = uiSelection.projectUsers;
  // get variables ready
  var fields = templateFields[model.currentArtifact.field],
    hasBespoke = fieldsWithBespokeData(fields);

  // add bespoke data to relevant fields 
  if (hasBespoke) {
    fields.filter(function (a) {
      var bespokeFieldHasValues = typeof uiSelection[model.currentArtifact.field][a.field] != "undefined" &&
        uiSelection[model.currentArtifact.field][a.field].values;
      return bespokeFieldHasValues;
    }).map(function (field) {
      if (field.bespoke) {
        field.values = uiSelection[model.currentArtifact.field][field.field].values;
      }
      return field;
    });
  }

  model.fields = fields;

  // get rid of any dropdowns that don't have any values attached
  model.fields = model.fields.filter(function (field) {
    var isNotDrop = field.type !== params.fieldType.drop;
    return isNotDrop || field.values.length > 0;
  });

  // call server side template function
  if (isGoogle) {
    google.script.run
      .withSuccessHandler(templateLoaderSuccess)
      .withFailureHandler(errorUnknown)
      .templateLoader(model, params.fieldType, advancedMode);
  } else {
    msOffice.templateLoader(model, params.fieldType, advancedMode)
      .then(response => templateLoaderSuccess(response))
      .catch(error => error.description ? errorExcel(error) : errorNetwork(error));
  }
}



// once template is loaded, enable the "send to Spira" button
function templateLoaderSuccess(response) {
  model.isTemplateLoaded = true;

  //turn off ajax spinner if it's on
  hideLoadingSpinner();

  // if we get a response string back from server then that means the template was not fully loaded 
  if (response && response.isTemplateLoadFail) {
    return;
  }

  // if we are trying to get data from Spira (ie we clicked the button do so that kicked off loading the template before getting the data itself, get it now)
  if (model.isGettingDataAttempt) {
    getFromSpiraAttempt();
  }
  //enable the send to spira button
  document.getElementById("btn-toSpira").innerHTML = "Add To Spira";
  document.getElementById("btn-toSpira").title = "Create entered data in SpiraPlan"

  //show text in the sidebar that tells the user what the template is set to:
  document.getElementById("template-project").textContent = model.currentProject.name;
  document.getElementById("template-artifact").textContent = model.currentArtifact.name;

}









/*
* 
* ==============
* ERROR HANDLING
* ==============
*
* These call a popup using google server side code
* most args are the HTTPResponse objects from the `withFailureHandler` promise
*
*/
function errorPopUp(type, err) {
  if (isGoogle) {
    google.script.run.error(type);
    //sets the UI to correspond to this mode
    artifactUpdateUI(UI_MODE.errorMode);
  } else {
    msOffice.error(type, err);
    //sets the UI to correspond to this mode
    artifactUpdateUI(UI_MODE.errorMode);

    if (err != null) {
      console.error("SpiraPlan Test Runner Tool encountered an error:", err.status ? err.status : "", err.response ? err.response.text : "", err.description ? err.description : "")
    }
    console.info("SpiraPlan Test Runner Tool: full error is... ", err)
  }
  hideLoadingSpinner();
}

function errorNetwork(err) {
  errorPopUp("network", err);
}
function errorImpExp(err) {
  errorPopUp('impExp', err);
}
function errorUnknown(err) {
  errorPopUp('unknown', err);
}
function errorExcel(err) {
  errorPopUp('excel', err);
}