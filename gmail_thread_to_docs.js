const spreadsheet2 = SpreadsheetApp.getActiveSpreadsheet();
const sheet2 = spreadsheet2.getActiveSheet();


function myFunction() {
  const unreads = GmailApp.search("label:debate-docs-chains AND is:starred");
  var blake_docs_folder = DriveApp.getFolderById(FOLDER_ID);

  if (unreads.length > 0) {
    for (var i=0; i < unreads.length; i++) {
      var thread = unreads[i];
      if (thread.hasStarredMessages()) {
        var subject = thread.getFirstMessageSubject();
        if (subject.startsWith("[Blake Docs] ")) {
            subject = subject.replace("[Blake Docs] ", "")
          }
        var messages = thread.getMessages();
        var threadIDs = sheet1.getRange('E:E').getValues();
        var threadIDflat = threadIDs.map(function(row) {return row[0];});
        if (threadIDflat.includes(thread.getId())){
          var matching_row = val_search(thread.getId());
          var folderId = sheet1.getRange(matching_row, 4).getValue();
          var round = DriveApp.getFolderById(folderId);
        
        } else {
          var round = DriveApp.createFolder(subject);
          round.moveTo(blake_docs_folder);
          var round_url = round.getUrl();
          sheet1.insertRowBefore(2);
          sheet1.getRange(2, 3).setValue(round_url);
          sheet1.getRange(2, 2).setValue(subject);
          sheet1.getRange(2, 1).setValue(messages[0].getDate());
          sheet1.getRange(2, 4).setValue(round.getId());
          sheet1.getRange(2, 5).setValue(thread.getId());
          sheet1.getRange(2,7).setValue(messages[0].getTo());
        }
  
        for (var z=0; z < messages.length; z++) {
          var message = messages[z];
          if(message.isStarred()) {
            var body = message.getBody();
            var body_blob = Utilities.newBlob("").setDataFromString(body, "UTF-8").setContentType("text/html").setName("Email Body" + z + " - " + subject);
            var attachments = message.getAttachments();
            if (attachments.length > 0) {
              for (b=0; b < attachments.length; b++) {
                var attachment = attachments[b];
                var uploaded = DriveApp.createFile(attachment).moveTo(round);
              }
            }
            var newFileId = Drive.Files.insert({title: body_blob.getName()}, body_blob, {convert: true}).id;
            var recent = DriveApp.getFileById(newFileId).moveTo(round);
            message.markRead();
            message.unstar();
        }
      }
    }
    }
  }
}

function val_search (v) {
  const dataRange = sheet1.getDataRange();
  const values = dataRange.getValues();
  const columnIndex = 4 // INDEX OF COLUMN FOR COMPARISON CELL
  const matchText = v;
  const index = values.findIndex(row => row[columnIndex] === matchText)
  const rowNumber = index + 1
  return rowNumber
}
