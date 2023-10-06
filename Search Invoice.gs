function search() {
  searchRow()
  searchInvoiceRows()
}
function searchRow() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName("Invoice");
  var dSheet = ss.getSheetByName("Tagihan");
  var invoiceid = sheet.getRange("O2").getValue();
  var lastRow = dSheet.getLastRow();
  var foundRecord = false;
  
  for(var j = 6; j <  lastRow; j++)
  {
    if(dSheet.getRange(j,13).getValue() ==invoiceid)
    {
      sheet.getRange("D8").setValue(dSheet.getRange(j, 7).getValue()) ;
      sheet.getRange("D9").setValue(dSheet.getRange(j, 8).getValue()) ;
      sheet.getRange("D10").setValue(dSheet.getRange(j, 9).getValue()) ;
      sheet.getRange("D11").setValue(dSheet.getRange(j, 10).getValue()) ;
      sheet.getRange("L10").setValue(dSheet.getRange(j, 14).getValue()) ;
      sheet.getRange("O8").setValue(dSheet.getRange(j, 11).getValue()) ;
      ss.getRange('L14').activate().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
      ss.getCurrentCell().offset(-3, 0).activate();
      ss.getCurrentCell().setValue(dSheet.getRange(j, 21).getValue());
      ss.getCurrentCell().offset(-1, 0).activate();
      ss.getCurrentCell().setValue(dSheet.getRange(j, 20).getValue());
      ss.getCurrentCell().offset(-1, 0).activate();
      ss.getCurrentCell().setValue(dSheet.getRange(j, 19).getValue());
      ss.getCurrentCell().offset(-1, 0).activate();
      ss.getCurrentCell().setValue(dSheet.getRange(j, 18).getValue());
      ss.getCurrentCell().offset(-1, 0).activate();
      ss.getCurrentCell().setValue(dSheet.getRange(j, 17).getValue());
      foundRecord = true; 
    }
  }

  
  if(foundRecord == false)
  {
    sheet.getRange("N3").setValue(['(NO RECORDS FOUND)']); 
  }
}

function searchInvoiceRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Invoice");
  var dsheet = ss.getSheetByName("Produk Tagihan");
  var invoiceid = sheet.getRange("O2").getValue();
  var lastrowStok = dsheet.getLastRow();
  var foundRecord = false;
  //var lastrowInvoice = sheet.getLastRow();

  var z = 15;
  //var subTotal = 0;
   
    for(var y = 2; y <= lastrowStok; y++)
    {
      if(dsheet.getRange(y,5).getValue() ==invoiceid)
      {
      //GET ITEM VALUES FROM STOK SHEET
      var part = dsheet.getRange(y, 7).getValue();
      var size= dsheet.getRange(y, 8).getValue();
      var quantity = dsheet.getRange(y, 9).getValue();
      sheet.getRange(z, 2).setValue(part);
      sheet.getRange(z, 5).setValue(size);
      sheet.getRange(z, 6).setValue(quantity);
      z++;
      foundRecord = true; 
      }
    }
    if(foundRecord == false)
  {
  sheet.getRange("N3").clearContent();
  }
}
