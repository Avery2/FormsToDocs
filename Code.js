/**
 * Creates a trigger for when form entered.
 */
function createSpreadsheetOpenTrigger() {
  // test message
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
 * Exports most recent entry to spreadsheet from form as a Google Doc.
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
  // TODO error handling if empty (will be empty array)
  Logger.log(form.getResponses());
  var latestResponses = form.getResponses()[form.getResponses().length - 1].getItemResponses();

  // MAKE DOCUMENT
  // create doc
  var doc = DocumentApp.create(d.toLocaleTimeString());
  // save to folder
  // TODO automate this, and create folder if not found
  var folder = DriveApp.getFoldersByName(
    "Contact Information (Documents)"
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
  headerStyle[DocumentApp.Attribute.HEADING] =
    DocumentApp.ParagraphHeading.HEADING1;
  bodyStyle = {};

  Logger.clear();
  titles.forEach(function(currentValue) {
    Logger.log(currentValue);
  });
  latestResponses.forEach(function(currentValue) {
    Logger.log(currentValue);
  });

  // TODO fill doc here
  latestResponses.forEach(function(currentValue, index) {
    body.appendParagraph("" + titles[index]).setAttributes(headerStyle);
    body.appendParagraph("" + latestResponses[index]).setAttributes(bodyStyle);
    body.appendHorizontalRule();
  });
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
