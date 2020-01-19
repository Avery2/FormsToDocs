var ADDON_TITLE = "Forms to Docs";

var NOTICE = "uh.";

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
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings) {
  PropertiesService.getDocumentProperties().setProperties(settings);
  // adjustFormSubmitTrigger();
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

  // Get text field items in the form and compile a list
  //   of their titles and IDs.
  var form = FormApp.getActiveForm();
  var textItems = form.getItems(FormApp.ItemType.TEXT);
  settings.textItems = [];
  for (var i = 0; i < textItems.length; i++) {
    settings.textItems.push({
      title: textItems[i].getTitle(),
      id: textItems[i].getId()
    });
  }

  PropertiesService.getDocumentProperties().getKeys().forEach(function callback(value) {
    Logger.log(value);
  });

  return settings;
}

/**
 * Adjust the onFormSubmit trigger based on user's requests.
 */
// function adjustFormSubmitTrigger() {
//   var form = FormApp.getActiveForm();
//   var triggers = ScriptApp.getUserTriggers(form);
//   var settings = PropertiesService.getDocumentProperties();
//   var triggerNeeded =
//     settings.getProperty("creatorNotify") == "true" ||
//     settings.getProperty("respondentNotify") == "true";

//   // Create a new trigger if required; delete existing trigger
//   //   if it is not needed.
//   var existingTrigger = null;
//   for (var i = 0; i < triggers.length; i++) {
//     if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
//       existingTrigger = triggers[i];
//       break;
//     }
//   }
//   if (triggerNeeded && !existingTrigger) {
//     var trigger = ScriptApp.newTrigger("respondToFormSubmit")
//       .forForm(form)
//       .onFormSubmit()
//       .create();
//   } else if (!triggerNeeded && existingTrigger) {
//     ScriptApp.deleteTrigger(existingTrigger);
//   }
// }

/**
 * Responds to a form submission event if an onFormSubmit trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
// function respondToFormSubmit(e) {
//   var settings = PropertiesService.getDocumentProperties();
//   var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

//   // Check if the actions of the trigger require authorizations that have not
//   // been supplied yet -- if so, warn the active user via email (if possible).
//   // This check is required when using triggers with add-ons to maintain
//   // functional triggers.
//   if (
//     authInfo.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.REQUIRED
//   ) {
//     // Re-authorization is required. In this case, the user needs to be alerted
//     // that they need to reauthorize; the normal trigger action is not
//     // conducted, since authorization needs to be provided first. Send at
//     // most one 'Authorization Required' email a day, to avoid spamming users
//     // of the add-on.
//     sendReauthorizationRequest();
//   } else {
//     // All required authorizations have been granted, so continue to respond to
//     // the trigger event.

//     // Check if the form creator needs to be notified; if so, construct and
//     // send the notification.
//     if (settings.getProperty("creatorNotify") == "true") {
//       sendCreatorNotification();
//     }

//     // Check if the form respondent needs to be notified; if so, construct and
//     // send the notification. Be sure to respect the remaining email quota.
//     if (
//       settings.getProperty("respondentNotify") == "true" &&
//       MailApp.getRemainingDailyQuota() > 0
//     ) {
//       sendRespondentNotification(e.response);
//     }
//   }
// }

/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
// function sendReauthorizationRequest() {
//   var settings = PropertiesService.getDocumentProperties();
//   var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
//   var lastAuthEmailDate = settings.getProperty("lastAuthEmailDate");
//   var today = new Date().toDateString();
//   if (lastAuthEmailDate != today) {
//     if (MailApp.getRemainingDailyQuota() > 0) {
//       var template = HtmlService.createTemplateFromFile("AuthorizationEmail");
//       template.url = authInfo.getAuthorizationUrl();
//       template.notice = NOTICE;
//       var message = template.evaluate();
//       MailApp.sendEmail(
//         Session.getEffectiveUser().getEmail(),
//         "Authorization Required",
//         message.getContent(),
//         {
//           name: ADDON_TITLE,
//           htmlBody: message.getContent()
//         }
//       );
//     }
//     settings.setProperty("lastAuthEmailDate", today);
//   }
// }

/**
 * Sends out creator notification email(s) if the current number
 * of form responses is an even multiple of the response step
 * setting.
 */
// function sendCreatorNotification() {
//   var form = FormApp.getActiveForm();
//   var settings = PropertiesService.getDocumentProperties();
//   var responseStep = settings.getProperty("responseStep");
//   responseStep = responseStep ? parseInt(responseStep) : 10;

//   // If the total number of form responses is an even multiple of the
//   // response step setting, send a notification email(s) to the form
//   // creator(s). For example, if the response step is 10, notifications
//   // will be sent when there are 10, 20, 30, etc. total form responses
//   // received.
//   if (form.getResponses().length % responseStep == 0) {
//     var addresses = settings.getProperty("creatorEmail").split(",");
//     if (MailApp.getRemainingDailyQuota() > addresses.length) {
//       var template = HtmlService.createTemplateFromFile("CreatorNotification");
//       template.sheet = DriveApp.getFileById(form.getDestinationId()).getUrl();
//       template.summary = form.getSummaryUrl();
//       template.responses = form.getResponses().length;
//       template.title = form.getTitle();
//       template.responseStep = responseStep;
//       template.formUrl = form.getEditUrl();
//       template.notice = NOTICE;
//       var message = template.evaluate();
//       MailApp.sendEmail(
//         settings.getProperty("creatorEmail"),
//         form.getTitle() + ": Form submissions detected",
//         message.getContent(),
//         {
//           name: ADDON_TITLE,
//           htmlBody: message.getContent()
//         }
//       );
//     }
//   }
// }

/**
 * Sends out respondent notification emails.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
// function sendRespondentNotification(response) {
//   var form = FormApp.getActiveForm();
//   var settings = PropertiesService.getDocumentProperties();
//   var emailId = settings.getProperty("respondentEmailItemId");
//   var emailItem = form.getItemById(parseInt(emailId));
//   var respondentEmail = response.getResponseForItem(emailItem).getResponse();
//   if (respondentEmail) {
//     var template = HtmlService.createTemplateFromFile("RespondentNotification");
//     template.paragraphs = settings.getProperty("responseText").split("\n");
//     template.notice = NOTICE;
//     var message = template.evaluate();
//     MailApp.sendEmail(
//       respondentEmail,
//       settings.getProperty("responseSubject"),
//       message.getContent(),
//       {
//         name: form.getTitle(),
//         htmlBody: message.getContent()
//       }
//     );
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
