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
  if (!DriveApp.getFoldersByName().hasNext()) {
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
    //body.appendHorizontalRule();
  });
}

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
