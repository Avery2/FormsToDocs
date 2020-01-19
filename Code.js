var ADDON_TITLE = "Forms to Docs";

var NOTICE = "NOTICE";

/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  FormApp.getUi()
    .createAddonMenu()
    .addItem("Configure", "showSidebar")
    .addItem("About", "showAbout")
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
  // deleteAllTriggers();
  createSpreadsheetOpenTrigger();
}

/**
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile("Sidebar").setTitle(
    "Forms to Docs"
  );
  FormApp.getUi().showSidebar(ui);
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile("About")
    .setWidth(420)
    .setHeight(270);
  FormApp.getUi().showModalDialog(ui, "About Form Notifications");
}

/**
 * Save sidebar settings to this form's Properties, and update the onFormSubmit
 * trigger as needed. Called by Sidebar on user input.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
}

/**
 * Queries the User Properties and adds additional data required to populate
 * the sidebar UI elements.
 *
 * @return {Object} A collection of Property values and
 *     related data used to fill the configuration sidebar.
 */
function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();

  // Use a default email if the creator email hasn't been provided yet.
  if (!settings.creatorEmail) {
    settings.creatorEmail = Session.getEffectiveUser().getEmail();
  }

  // Get text field items in the form and compile a list of their titles and IDs.
  var form = FormApp.getActiveForm();
  var textItems = form.getItems(FormApp.ItemType.TEXT);
  settings.textItems = [];
  for (var i = 0; i < textItems.length; i++) {
    settings.textItems.push({
      title: textItems[i].getTitle(),
      id: textItems[i].getId()
    });
  }

  PropertiesService.getDocumentProperties()
    .getKeys()
    .forEach(function callback(value) {
      Logger.log(value);
    });

  return settings;
}

// usage, settings.getProperty("SettingA") == "true";
/**
 * Responds to a form submission event if an onFormSubmit trigger has been
 * enabled. BROKE SAVED FOR REFERENCE
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
//   var settings = PropertiesService.getDocumentProperties();
//   var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
//   if (
//     authInfo.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.REQUIRED
//   ) {
//     sendReauthorizationRequest();
//   } else {
//     if (settings.getProperty("creatorNotify") == "true") {
//       sendCreatorNotification();
//     }
//     if (
//       settings.getProperty("respondentNotify") == "true" &&
//       MailApp.getRemainingDailyQuota() > 0
//     ) {
//       sendRespondentNotification(e.response);
//     }
//   }
// }

/**
 * Creates a google document representation of each entry into a google forms. Automatically stores google documents in folder.
 */
function exportRecentAsDoc() {
  // SETUP
  var form = FormApp.getActiveForm();
  var d = new Date(); // create date object (for testing purposes)

  // READ FORM ENTRIES
  // parse, to get last row's data and all questions
  var titles = [];
  form.getItems().forEach(function(item, index) {
    titles[index] = item.getTitle();
  });

  if (form.getResponses().length != 0) {
    var latestResponses = form
      .getResponses()
      [form.getResponses().length - 1].getItemResponses();
  } else {
    var latestResponses = [];
  }

  // MAKE DOCUMENT
  // create doc
  var doc = DocumentApp.create(d.toLocaleTimeString());
  // save to folder
  if (!DriveApp.getFoldersByName(form.getTitle() + " (Documents)").hasNext()) {
    // if doesn't exist, create it
    DriveApp.createFolder(form.getTitle() + " (Documents)");
  }
  var folder = DriveApp.getFoldersByName(
    form.getTitle() + " (Documents)"
  ).next();
  folder.addFile(DriveApp.getFileById(doc.getId())); // creates second reference to same file
  // remove from root (default) folder, which was automatically placed there on creation
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

  // FILL DOCUMENT
  var body = doc.getBody();
  // format style
  headerStyle = {};
  //  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
  //  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.UNDERLINE] = true;
  bodyStyle = {};
  bodyStyle[DocumentApp.Attribute.BOLD] = false;
  bodyStyle[DocumentApp.Attribute.UNDERLINE] = false;
  latestResponses.forEach(function(currentValue, index) {
    body.appendParagraph("" + titles[index]).setAttributes(headerStyle);
    body
      .appendParagraph("" + latestResponses[index].getResponse())
      .setAttributes(bodyStyle);
    // TODO handle formatting types: a String or String[] or String[][] of answers to the question item
    body.appendParagraph("");
  });

  // Write log to spreadsheet

  var settings = PropertiesService.getDocumentProperties();
  if (settings.getProperty("exportToSheet")) {
    if (!settings.getProperty("sheetURL")) {
      var ss = SpreadsheetApp.create(form.getTitle() + " (Document Links)");
    } else {
      var ss = SpreadsheetApp.openByUrl(settings.getProperty("sheetURL"));
    }
    ss.appendRow(doc.getUrl());
  }
}

/**
 * Creates a trigger for when form entered.
 */
function createSpreadsheetOpenTrigger() {
  var form = FormApp.getActiveForm();
  try {
    deleteAllTriggers();
    ScriptApp.newTrigger("exportRecentAsDoc")
      .forForm(form)
      .onFormSubmit()
      .create();
  } catch (error) {
    SpreadsheetApp.getUi().alert("Confirmation received.");
  }
}

/**
 * Deletes all triggers.
 */
function deleteAllTriggers() {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // remove triggers
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}
