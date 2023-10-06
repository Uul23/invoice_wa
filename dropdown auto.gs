function deleteRowWithEmptyValues() {
  var sheetName = "Produk Tagihan";
  var columnIndex = [7, 8, 9]; // Columns G, H, and I

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, columnIndex[0], lastRow, columnIndex.length);
  var values = range.getValues();

  for (var i = lastRow - 1; i >= 0; i--) {
    var emptyRow = true;
    for (var j = 0; j < columnIndex.length; j++) {
      if (values[i][j] !== "") {
        emptyRow = false;
        break;
      }
    }

    if (emptyRow) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function createTrigger() {
  var sheetName = "Produk Tagihan";
  var triggerName = "onDeleteLastRowWithEmptyValues";

  // Delete existing trigger if it exists
  var existingTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(existingTriggers[i]);
    }
  }

  // Create new trigger
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  ScriptApp.newTrigger(triggerName)
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
}

function onEdit(e) {
  deleteRowWithEmptyValues();
  dropdownlist()
  dropdownreturnlist()
  dropdownlabellist()
  dropdownpackinglist()
}

function dropdownlist1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dsheet = ss.getSheetByName("Produk");
  var sheet = ss.getSheetByName("Invoice");
  var sheetName = ss.getActiveSheet().getName();
  var activeCell = ss.getActiveRange();
  var activeValue = activeCell.getValue();
  var activeColumn = activeCell.getColumn();
  var activeRow = activeCell.getRow();
  var activeSheet = activeCell.getSheet();

  if (activeSheet.getName() == "Invoice" && activeColumn == 2 && activeRow > 14) {
    if (activeCell.isBlank()) {
      activeCell.offset(0, 3).clearContent().clearDataValidations();
    } else {
      var dsheetarray = dsheet.getDataRange().getValues();
      var filtersize = dsheetarray.reduce(function (acc, row) {
        if (row[3] === activeValue) {
          acc.push(row[4]);
        }
        return acc;
      }, []);
      var modelDataValidation = SpreadsheetApp.newDataValidation().requireValueInList(filtersize).build();
      activeCell.offset(0, 3).setDataValidation(modelDataValidation);
    }
  }
}

function dropdownlist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dsheet = ss.getSheetByName("Produk");
  var sheet = ss.getSheetByName("Invoice");
  var sheetName = ss.getActiveSheet().getName();
  var activeCell = ss.getActiveRange();
  var activeValue = activeCell.getValue();
  var activeColumn = activeCell.getColumn();
  var activeRow = activeCell.getRow();
  var activeSheet = activeCell.getSheet();

if( activeSheet.getName()== "Invoice" && activeColumn == 2 && activeRow > 14)
  {
    if( activeCell.isBlank() == true )
  {
    activeCell.offset(0,3).clearContent().clearDataValidations();
  }
 else{
    var dsheetarray = dsheet.getDataRange().getValues();
    var filterproduk = dsheetarray.filter (row=>row[3]== activeValue)
    var filtersize = filterproduk.map(row=>row[4])
    var modelDataValidation = SpreadsheetApp.newDataValidation().requireValueInList(filtersize).build();
    activeCell.offset(0,3).setDataValidation(modelDataValidation)
  }
  }
}

function dropdownreturnlist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dsheet = ss.getSheetByName("Produk");
  var sheet = ss.getSheetByName("Return");
  var sheetName = ss.getActiveSheet().getName();
  var activeCell = ss.getActiveRange();
  var activeValue = activeCell.getValue();
  var activeColumn = activeCell.getColumn();
  var activeRow = activeCell.getRow();
  var activeSheet = activeCell.getSheet();

if( activeSheet.getName()== "Return" && activeColumn == 8 && activeRow > 5)
  {
    if( activeCell.isBlank() == true )
  {
    activeCell.offset(0,1).clearContent().clearDataValidations();
  }
 else{
    var dsheetarray = dsheet.getDataRange().getValues();
    var filterproduk = dsheetarray.filter (row=>row[3]== activeValue)
    var filtersize = filterproduk.map(row=>row[4])
    var modelDataValidation = SpreadsheetApp.newDataValidation().requireValueInList(filtersize).build();
    activeCell.offset(0,1).setDataValidation(modelDataValidation)
  }
  }
}
  
function dropdownlabellist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dsheet = ss.getSheetByName("Produk");
  var sheet = ss.getSheetByName("Pelabelan");
  var sheetName = ss.getActiveSheet().getName();
  var activeCell = ss.getActiveRange();
  var activeValue = activeCell.getValue();
  var activeColumn = activeCell.getColumn();
  var activeRow = activeCell.getRow();
  var activeSheet = activeCell.getSheet();

if( activeSheet.getName()== "Pelabelan" && activeColumn == 4 && activeRow > 1)
  {
    if( activeCell.isBlank() == true )
  {
    activeCell.offset(0,1).clearContent().clearDataValidations();
  }
 else{
    var dsheetarray = dsheet.getDataRange().getValues();
    var filterproduk = dsheetarray.filter (row=>row[3]== activeValue)
    var filtersize = filterproduk.map(row=>row[4])
    var modelDataValidation = SpreadsheetApp.newDataValidation().requireValueInList(filtersize).build();
    activeCell.offset(0,1).setDataValidation(modelDataValidation)
  }
  }
}
function dropdownpackinglist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var dsheet = ss.getSheetByName("Packing");
  var sheet = ss.getSheetByName("Packing");
  var sheetName = ss.getActiveSheet().getName();
  var activeCell = ss.getActiveRange();
  var activeValue = activeCell.getValue();
  var activeColumn = activeCell.getColumn();
  var activeRow = activeCell.getRow();
  var activeSheet = activeCell.getSheet();

if( activeSheet.getName()== "Packing" && activeColumn == 4 && activeRow > 1)
  {
    if( activeCell.isBlank() == true )
  {
    activeCell.offset(0,1).clearContent().clearDataValidations();
  }
 else{
    var sheetarray = sheet.getDataRange().getValues();
    var filterproduk = sheetarray.filter (row=>row[9]== activeValue)
    var filtersize = filterproduk.map(row=>row[10])
    var modelDataValidation = SpreadsheetApp.newDataValidation().requireValueInList(filtersize).build();
    activeCell.offset(0,1).setDataValidation(modelDataValidation)
  }
  }
}
