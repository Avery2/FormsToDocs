/**
* Creates a trigger for when form entered.
*/
function createSpreadsheetOpenTrigger() {
  var form = FormApp.getActiveForm();
  try {
    ScriptApp.newTrigger('exportRecentAsDoc').forForm(form).onFormSubmit().create();
  } catch (error) {
    SpreadsheetApp.getUi().alert('Confirmation received.');
  }
}

/**
* Exports most recent entry to spreadsheet from form as a Google Doc.
*/
function exportRecentAsDoc() {
  // setup
  var form = FormApp.getActiveForm();
  var d = new Date(); // create date object (for testing purposes)
  
  // parse, to get last row's data
  var items = form.getItems();
  var responses = form.getResponses();
  var titles = items.forEach(function(item, index){
                             titles[index] = item.getTitle();
                             });
  var latestResponse = responses[responses.length-1];
  
  // create doc
  var doc = DocumentApp.create(d.toLocaleTimeString());
  
  // save to folder
  var folder = DriveApp.getFoldersByName("Contact Information (Documents)").next();
  folder.addFile(DriveApp.getFileById(doc.getId())); // creates second reference to same file
  // remove from root (default) folder, which was automatically placed there on creation
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));
  
  // fill doc with parsed data
  var body = doc.getBody();
  // body.appendParagraph("Name: "+name+" Contact Info: "+info);
  
  // format style
  headerStyle = {};
//  headerStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
//  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.HEADING] = DocumentApp.ParagraphHeading.HEADING1;
  bodyStyle = {};
  
  Logger.clear();
  headers.forEach(function(currentValue) {Logger.log(currentValue)});
  lastRow.forEach(function(currentValue) {Logger.log(currentValue)});
  
  // FILL DOC HERE
  lastRow.forEach(function(currentValue, index) {
    //    body.appendParagraph(''+headers[index]).setAttributes(headerStyle);
    body.appendParagraph(''+headers[index]).setAttributes(headerStyle);
    body.appendParagraph(''+lastRow[index]).setAttributes(bodyStyle);
    body.appendHorizontalRule();
  });
}

/**
* Deletes all triggers.
*/
function deleteAllTrigger() {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // remove triggers
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}