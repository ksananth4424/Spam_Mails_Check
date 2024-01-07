function getGmailEmails(){
  var threads = GmailApp.getInboxThreads();
  var startIndex = 0;
  var maxThreads = 500;
  do {
    threads = GmailApp.getInboxThreads(startIndex, maxThreads);
    for(var i=0; i<threads.length; i++){
    var messages = threads[i].getMessages();
    var msgCount = threads[i].getMessageCount();
    for(var j=0; j<messages.length; j++){
      message = messages[j];
      if(message.isInInbox()){
        extractDetails(message, msgCount);
      }
    }
  }
  startIndex += maxThreads;
  } while (threads.length == maxThreads);
}
function extractDetails(message, msgCount){
  var spreadSheetId = '157c9s-QFe31TuvCv_W01bm8OjBXmdDKQBnYZ6fhHstc';
  var sheetname = "Sheet1";
  var ss = SpreadsheetApp.openById(spreadSheetId);
  var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(sheetname);
  const today = new Date();
  var dateTime = Utilities.formatDate(message.getDate(), timezone, "dd-MM-yyyy");
  var subjectText = message.getSubject();
  var fromSemd = message.getFrom();
  var toSend = message.getTo();
  var bodyCount = message.getPlainBody();
  sheet.appendRow([dateTime, msgCount, fromSemd, toSend, subjectText, bodyCount]);
}
function onOpen(e){
  SpreadsheetApp.getUi()
  .createMenu('Click to Fetch Emails')
  .addItem('Get Email', 'getGmailEmails')
  .addToUi();
}