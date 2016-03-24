
/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  if (e && e.authMode == ScriptApp.AuthMode.LIMITED) {
    menu
      .addItem('Splice by Row', 'SpliceByRow')c
      .addItem('Splice by Column', 'SpliceByColumn')
      .addToUi();
  } else {
    menu
      .addItem('Splice by Row', 'SpliceByRow')
      .addItem('Splice by Column', 'SpliceByColumn')
      .addToUi();
  }
}

function onInstall(e) {
  onOpen(e);
}

function transpose(a)
{
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function removeEmptyColumns(){
  var sh = SpreadsheetApp.getActiveSheet();
  var data = sh.getDataRange().getValues();
  data = transpose(data);
  var targetData = new Array();
  Logger.log(data);
  for(n=0;n<data.length;n++){
    if(data[n].join().replace(/,/g,'')!=''){ targetData.push(data[n])};
  }  
  sh.getDataRange().clear();
  Logger.log(targetData);
  targetData = transpose(targetData);
  sh.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);

}

function removeEmptyRows(){ 
  var sh = SpreadsheetApp.getActiveSheet();
  var data = sh.getDataRange().getValues();
  var targetData = new Array();
  for(n=0;n<data.length;++n) {
    if(data[n].join().replace(/,/g,'')!=''){ targetData.push(data[n])};
    Logger.log(data[n].join().replace(/,/g,''))
  }
  sh.getDataRange().clear();
  sh.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);

}

function SpliceByRow() {
  removeEmptyRows()
  removeEmptyColumns()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  var row = 1;
  var col = 1;
  var row_range = sheet.getRange(row, col);
  var row_counter = 0;
  while (!row_range.isBlank()) {
    row += 1;
    row_range = sheet.getRange(row, col);
    row_counter += 1;
  }
  var col_range = sheet.getRange(1,col);
  var col_counter = 0;
  while (!col_range.isBlank()) {
    col +=1;
    col_range = sheet.getRange(1,col);
    col_counter += 1;
  }
 
  for (var row = 2; row <= row_counter; row++) {   
    var sheetTitle = sheet.getRange(row, 1,1,1).getValues();
    if (ss.getSheetByName(sheetTitle) == null) {
       
      var targetsheet = ss.insertSheet('' + sheetTitle);
      
      var headerrow = sheet.getRange(1, 1, 1, col_counter);
      var pasteheaderrow = targetsheet.getRange(1,1,1,col_counter);    
      pasteheaderrow.setFormula("='" + sheet.getSheetName() +"'!"+ headerrow.getA1Notation().split(":")[0]);
      pasteheaderrow.setNumberFormat("MMM, yy");
      
      var copyrow = sheet.getRange(row, 1, 1, col_counter);
      var pasterow = targetsheet.getRange(2,1,1,col_counter);
      pasterow.setFormula("='" + sheet.getSheetName() +"'!" + copyrow.getA1Notation().split(":")[0]);
      
     } else {
      var existing_sheet = ss.getSheetByName(sheetTitle);
      existing_sheet.getDataRange().clear();
      var existing_headerrow = existing_sheet.getRange(1,1,1,col_counter);
      existing_headerrow.clear();
      
      var originalheader = sheet.getRange(1, 1, 1, col_counter);
      var originalheader_values = originalheader.getValues();
      existing_headerrow.setFormula("='" + sheet.getSheetName() +"'!"+ originalheader.getA1Notation().split(":")[0]); 
      
      var targetrow = existing_sheet.getRange(2, 1, 1, col_counter);
      targetrow.clear();
      var original_target_row = sheet.getRange(row, 1, 1, col_counter);
      targetrow.setFormula("='" + sheet.getSheetName() +"'!"+ original_target_row.getA1Notation().split(":")[0]);  
     }
  } 
}


function SpliceByColumn() {
  removeEmptyRows()
  removeEmptyColumns()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var row = 1;
  var col = 1;
  var row_range = sheet.getRange(row, col);
  var row_counter = 0;
  while (!row_range.isBlank()) {
    row += 1;
    row_range = sheet.getRange(row, col);
    row_counter += 1;
  }
  var col_range = sheet.getRange(1,col);
  var col_counter = 0;
  while (!col_range.isBlank()) {
    col +=1;
    col_range = sheet.getRange(1,col);
    col_counter += 1;
  }
  
  for (var col = 2; col <= col_counter; col++) {   
    var sheetTitle = sheet.getRange(1,col,1,1).getValues();
    if (ss.getSheetByName(sheetTitle) == null) {
      var targetsheet = ss.insertSheet('' + sheetTitle);     
      var headercolumn = sheet.getRange(1, 1, row_counter, 1);
      var pasteheadercolumn = targetsheet.getRange(1,1,row_counter,1);    
      pasteheadercolumn.setFormula("='" + sheet.getSheetName() +"'!"+ headercolumn.getA1Notation().split(":")[0]);
      pasteheadercolumn.setNumberFormat("MMM, yy");
    
      var copycolumn = sheet.getRange(1,col, row_counter, 1)
      var pastecolumn = targetsheet.getRange(1,2,row_counter,1);
      pastecolumn.setFormula("='" + sheet.getSheetName() +"'!" + copycolumn.getA1Notation().split(":")[0]);
      
    } else {
      var existing_sheet = ss.getSheetByName(sheetTitle);
      existing_sheet.getDataRange().clear();
      var existing_headercol = existing_sheet.getRange(1, 1, row_counter, 1);
      existing_headercol.clear();
      
      var originalheader = sheet.getRange(1, 1, row_counter, 1);
      var originalheader_values = originalheader.getValues();
      existing_headercol.setFormula("='" + sheet.getSheetName() +"'!"+ originalheader.getA1Notation().split(":")[0]); 
    
      var targetcol = existing_sheet.getRange(1,2,row_counter,1);
      targetcol.clear()
      var original_target_col = sheet.getRange(1,col, row_counter, 1);
      targetcol.setFormula("='" + sheet.getSheetName() +"'!"+ original_target_col.getA1Notation().split(":")[0]); 
    }
  } 
}


