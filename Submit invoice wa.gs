function submitInvoice() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  
  if (sheet.getRange("O2").isBlank()) {
    makePDF();
    makePDForder();
    makeHistory();
    copyRows();
    clearInvoiceFields1();
    SpreadsheetApp.getUi().alert('New Data Saved');
  } else {
    makePDF();
    makePDForder();
    makeHistory1();
    deleteRow();
    copyRows1();
    clearInvoiceFields1();
    SpreadsheetApp.getUi().alert('Data Updated');
  }
}

/* Send Spreadsheet in an email as PDF, automatically */
function makePDF() {
  
  // Get the currently active spreadsheet URL (link)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var token = ScriptApp.getOAuthToken();
  var sheet = ss.getSheetByName("Invoice");

  //Creating an exportable URL
  var url = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());
  var folderID = "1dGuqBGMd8S9fWeHwQ4a4dNmr6LF98NHk"; // Folder id to save in a folder.
  var folder = DriveApp.getFolderById(folderID);
  var invoiceNumber = ss.getRange("'Invoice'!L9").getValue()
  var name = ss.getRange("'Invoice'!D9").getValue()
  var pdfName = "Invoice -"+ name + " # " + invoiceNumber + " - " + Utilities.formatDate(new Date(), "GMT+7", "dd-MMM-yyyy");

  /* Specify PDF export parameters
  From: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579
  */
  var url_ext = 'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&range=A:L'                          // range paper for print
  + '&size=letter'                       // paper size legal / letter / A4
  + '&portrait=true'                    // orientation, false for landscape
  + '&fitw=true&source=labnol'           // fit to page width, false for actual size
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&gid=';                             // the sheet's Id
    
  // Convert individual worksheet to PDF
  var response = UrlFetchApp.fetch(url + url_ext + sheet.getSheetId(), {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  
  //convert the response to a blob
  var blobs = response.getBlob().setName(pdfName + '.pdf');
  
  //saves the file to the specified folder of Google Drive
  var newFile = folder.createFile(blobs);
  
  // Define the scope
  Logger.log("Storage Space used: " + DriveApp.getStorageUsed());

}

function makePDForder() {
  
  // Get the currently active spreadsheet URL (link)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var token = ScriptApp.getOAuthToken();
  var sheet = ss.getSheetByName("Invoice");

  //Creating an exportable URL
  var url = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());
  var folderID = "1dGuqBGMd8S9fWeHwQ4a4dNmr6LF98NHk"; // Folder id to save in a folder.
  var folder = DriveApp.getFolderById(folderID);
  var invoiceNumber = ss.getRange("'Invoice'!L9").getValue()
  var name = ss.getRange("'Invoice'!D9").getValue()
  var pdfName = "Order -"+ name + " # " + invoiceNumber + " - " + Utilities.formatDate(new Date(), "GMT+7", "dd-MMM-yyyy");

  /* Specify PDF export parameters
  From: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579
  */
  var url_ext = 'exportFormat=pdf&format=pdf'        // export as pdf / csv / xls / xlsx
  + '&range=T:AD'                          // range paper for print
  + '&size=letter'                       // paper size legal / letter / A4
  + '&portrait=true'                    // orientation, false for landscape
  + '&fitw=true&source=labnol'           // fit to page width, false for actual size
  + '&sheetnames=false&printtitle=false' // hide optional headers and footers
  + '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&gid=';                             // the sheet's Id
    
  // Convert individual worksheet to PDF
  var response = UrlFetchApp.fetch(url + url_ext + sheet.getSheetId(), {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  
  //convert the response to a blob
  var blobs = response.getBlob().setName(pdfName + '.pdf');
  
  //saves the file to the specified folder of Google Drive
  var newFile = folder.createFile(blobs);
  
  // Define the scope
  Logger.log("Storage Space used: " + DriveApp.getStorageUsed());

}


function makeHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var destSheet = ss.getSheetByName("Tagihan");
  
  //insert new line and copy formatting to it
  
  destSheet.insertRowBefore(6); //Inserts a row before populating it
  //var statusCell = destSheet.getRange("D6").setValue("Unpaid");
  var email = sheet.getRange("O8").getValue();
  //Define all the origin cells, dest cells and transcribe
  var duedate = sheet.getRange("L6").getValue();
  var date = sheet.getRange("L7").getValue();
  var idmember = sheet.getRange("D7").getValue();
  var member = sheet.getRange("D8").getValue();
  var name = sheet.getRange("D9").getValue();
  var contact = sheet.getRange("D10").getValues();
  var address = sheet.getRange("D11").getValues();
  var channel = sheet.getRange("L8").getValues();
  var orderid = sheet.getRange("L9").getValues();
  var ekspedisi = sheet.getRange("L10").getValues();
  var datas = sheet.getRange("R6:R11").getValues(); 
  var newData = [["",duedate,date,"Unpaid","",idmember,member,name,contact, address,email,channel,orderid,ekspedisi,""].concat(datas.flat())]; 
  var rangeToMove = destSheet.getRange(6, 1, newData.length, newData[0].length);
  rangeToMove.setValues(newData); 
}


function copyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var dsheet = ss.getSheetByName("Produk Tagihan");
  var orderid = sheet.getRange("L9").getValues()-1;
  var sheetData = sheet.getDataRange().getValues();
  var stokTerjualData = [];
  var date = sheet.getRange("L7").getValue();
  var channel = sheet.getRange("L8").getValues();
  for (var i = 14; i < sheetData.length; i++) {
    if(sheet.getRange(i,2).isBlank() === false )
    {
    dsheet.insertRowsBefore(2,1);
    var product = sheetData[i][1];
    var size = sheetData[i][4];
    var quantity = sheetData[i][5];
    var rowData = ["","",date,channel,orderid,"",product,size,quantity];
    stokTerjualData.push(rowData);
  }}
  
  var stokTerjualRange = dsheet.getRange(2, 1, stokTerjualData.length, 9);
  stokTerjualRange.setValues(stokTerjualData);
  

}
function makeHistory1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var dSheet = ss.getSheetByName("Tagihan");
  var invoiceId = sheet.getRange("O2").getValue();
  var lastRow = dSheet.getLastRow() - 5;
  var dataRange = dSheet.getRange(6, 1, lastRow, 24);
  var data = dataRange.getValues();
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var invoice = row[12];
  
    if (invoice === invoiceId) {
    var email = sheet.getRange("O8").getValue();
    //Define all the origin cells, dest cells and transcribe
    var duedate = sheet.getRange("L6").getValue();
    var date = sheet.getRange("L7").getValue();
    var idmember = sheet.getRange("D7").getValue();
    var member = sheet.getRange("D8").getValue();
    var name = sheet.getRange("D9").getValue();
    var contact = sheet.getRange("D10").getValues();
    var address = sheet.getRange("D11").getValues();
    var channel = sheet.getRange("L8").getValues();
    var orderid = sheet.getRange("L9").getValues();
    var ekspedisi = sheet.getRange("L10").getValues();
    var datas = sheet.getRange("R6:R11").getValues(); 
    var newData = [["",duedate,date,"Unpaid","",idmember,member,name,contact, address,email,channel,orderid,ekspedisi,""].concat(datas.flat())]; 
    var rangeToMove = dSheet.getRange(i+6, 1, newData.length, newData[0].length);
    rangeToMove.setValues(newData);  
    } 
  }
}


//function deleteRow() {
  
  //var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  //var sheet = ss.getSheetByName("Invoice");
  //var dsheet = ss.getSheetByName("Produk Tagihan");
  //var invoiceid = sheet.getRange("O2").getValue();
  //var lastRowEdit = dsheet.getLastRow();
  
  //for(var i = lastRowEdit; i > 0; i--)
 //{  
 //  if(dsheet.getRange(i,5).getValue() == invoiceid)
 // {
  //  dsheet.deleteRow(i); 
  //}
// }    
//}

function deleteRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var dsheet = ss.getSheetByName("Produk Tagihan");
  var invoiceid = sheet.getRange("O2").getValue();
  var lastRowEdit = dsheet.getLastRow();
  
  var rowsToDelete = [];
  
  for(var i = 1; i <= lastRowEdit; i++) {  
    if(dsheet.getRange(i, 5).getValue() == invoiceid) {
      rowsToDelete.push(i);
    }
  }
  
  for(var j = rowsToDelete.length - 1; j >= 0; j--) {
    dsheet.deleteRow(rowsToDelete[j]);
  }
}




function copyRows1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var dsheet = ss.getSheetByName("Produk Tagihan");
  var orderid = sheet.getRange("L9").getValues();
  var sheetData = sheet.getDataRange().getValues();
  var stokTerjualData = [];
  var date = sheet.getRange("L7").getValue();
  var channel = sheet.getRange("L8").getValues();
  for (var i = 14; i < sheetData.length; i++) {
    if(sheet.getRange(i,2).isBlank() == false )
    {
    dsheet.insertRowsBefore(2,1);
    var product = sheetData[i][1];
    var size = sheetData[i][4];
    var quantity = sheetData[i][5];
    var rowData = ["","",date,channel,orderid,"",product,size,quantity];
    stokTerjualData.push(rowData);
  }}
  
  var stokTerjualRange = dsheet.getRange(2, 1, stokTerjualData.length, 9);
  stokTerjualRange.setValues(stokTerjualData);
  
}



function clearInvoiceFields1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var rangeToClear;

  // Clear client cells
  rangeToClear = sheet.getRange("D8:I13");
  rangeToClear.clearContent();

  // Clear item cells
  rangeToClear = sheet.getRange("B15:B");
  rangeToClear.clearContent();

  // Clear size cells and data validations
  rangeToClear = sheet.getRange("E15:E");
  rangeToClear.clearContent().clearDataValidations();

  // Clear quantity cells
  rangeToClear = sheet.getRange("F15:F");
  rangeToClear.clearContent();

  // Clear via and email cells
  sheet.getRange("L10").clearContent();
  sheet.getRange("O8").clearContent();

  // Clear discCell
  rangeToClear = sheet.getRange("H15:H");
  rangeToClear.clearContent();

  // Clear searchID
  sheet.getRange("O2").clearContent();

  // Clear content on next row after "NOTE:" in column A
  rangeToClear = ss.getRange('A14').getNextDataCell(SpreadsheetApp.Direction.DOWN);
  rangeToClear.offset(0, 1).clearContent();
  // Set default values
  rangeToClear = ss.getRange('L14').getNextDataCell(SpreadsheetApp.Direction.DOWN);
  //rangeToClear.offset(-3, 0, 3).clearContent().setValues([['0%'], ['0%'], ['0']]);
  // rangeToClear.offset(-1, 0, 2).setValue('0');
    
  rangeToClear.offset(-4, 0, 2).setValue('0,00%');
  rangeToClear.offset(-7, 0, 3).setValue('Rp0');
  rangeToClear = ss.getRange('I14').getNextDataCell(SpreadsheetApp.Direction.DOWN);
  rangeToClear.offset(0, 0).setValue('');
  rangeToClear.offset(0, -1).setValue('Weight :');
}



