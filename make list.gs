function onOpen() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('List');
  menu.addItem('List All Emails', 'getGmailEmails')
  menu.addItem('List All Files', 'GdriveFiles')
  menu.addToUi();
}

function getGmailEmails(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Penagihan");
  var lastRow = sheet.getLastRow()-1;
  var startRow = 2;
  var dataRange = sheet.getRange(startRow, 1, lastRow, 13);
  var data = dataRange.getValues();
  for (var k = 0; k < data.length; ++k) {
  var row = data[k];
  //var emailAddress = row[4];
  var invoice = row[5];
  var ref = "GMS Invoice # "+ invoice;
  //var name = row[2];
  }
  var threads = GmailApp.search('in:sent subject:"'+ ref +'"');
  
  for(var i = threads.length - 1; i >=0; i--){
    var messages = threads[i].getMessages();
    
    for (var j = 0; j <messages.length; j++){
      var message = messages[j];
      //if (message.isUnread()){
        extractDetails(message);
       // GmailApp.markMessageRead(message);
      }
    }
    //threads[i].removeLabel(label);
  
}
  
function extractDetails(message){
  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var receiver = message.getTo();
  //var bodyContents = message.getPlainBody();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Email List");
  //var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow([dateTime, receiver, subjectText]);
}

function GdriveFiles() {

  const folderId = '1dGuqBGMd8S9fWeHwQ4a4dNmr6LF98NHk'

  const folder = DriveApp.getFolderById(folderId)

  const files = folder.getFiles()

  const source = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = source.getSheetByName('Link');

  const data = [];   

  while (files.hasNext()) {

      const childFile = files.next();

      var info = [ 

        childFile.getName(), 

        childFile.getUrl(),

        childFile.getLastUpdated(),

        Drive.Files.get(childFile.getId()).lastModifyingUser.displayName     

      ];

        data.push(info);

  }

  sheet.getRange(2,1,data.length,data[0].length).setValues(data);

}

