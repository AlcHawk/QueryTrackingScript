// === Create sidebar in HTML on spreadsheet UI ===
function createSidebarHTML() {
  var ui = HtmlService.createHtmlOutputFromFile("Starter").setTitle("Report Functions");
  SpreadsheetApp.getUi().showSidebar(ui);
}

// === Open sidebar HTML ===
function onOpen(e) {
    SpreadsheetApp.getUi().createAddonMenu()
      .addItem("ReportFuntions", "createSidebarHTML")
      .addToUi();
}

//function onInstall(e) {
//  onOpen(e);
//}

// === Custom function to get sheet name ===
function batchName() {
  var asheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = asheet.getName()
  if ( sheetName.substring(0,3) == "SCF") {
    var bname = "SCF" + sheetName.substring(4,6)
  } else {
    var bname = sheetName.substring(6,8)
  }
  return bname;
}

// === Create sheet Total and assign content ===
function createSheetTotal() {
  var mySS = SpreadsheetApp.getActive();
  Logger.log(mySS.getName());
  
  var sheets = mySS.getSheets();
  var sheetNames = [];
  for (var sheetno in sheets) {
    sheetNames.push(sheets[sheetno].getName());
    Logger.log(sheetno);
  }
  Logger.log(sheetNames);
  
  
  if ( sheetNames.indexOf("Total") === -1 ) {
    var sheettol = mySS.insertSheet("Total", sheets.length);
  } else {
    var rmsheet = mySS.getSheetByName("Total")
    mySS.deleteSheet(rmsheet)
    var sheettol = mySS.insertSheet("Total", sheets.length);
  }
  
  Logger.log(sheettol.getName())
  sheettol.setTabColor("orange");
  
  setSheetHeaderTot(sheettol);
  var cellformula = importSheetDataTot(mySS);
  var frange = sheettol.getRange("A2");
  frange.setFormula(cellformula);
  
//  var lstrowTotal = sheettol.getLastRow();
//  var daterange = sheettol.getRange("A2:F" + lstrowTotal)
//  daterange.setNumberFormats("dd MM ")
}

// === Helper function - create and set format for header ===
function setSheetHeaderTot (sheetname) {
  var colh0 = sheetname.getRange("A1");
  colh0.setValue("Batch");
  var colh1 = sheetname.getRange("B1");
  colh1.setValue("Site");
  var colh2 = sheetname.getRange("C1");
  colh2.setValue("Subject No.");
  var colh3 = sheetname.getRange("D1");
  colh3.setValue("Counts of queries");
  var colh4 = sheetname.getRange("E1");
  colh4.setValue("Date of Requested");
  var colh5 = sheetname.getRange("F1");
  colh5.setValue("Date of Received");
  
  var hrange = sheetname.getRange("A1:F1");
  hrange.setBackgroundRGB(153, 204, 255);
  hrange.setVerticalAlignment("middle");
  hrange.setHorizontalAlignment("center")
  hrange.setFontWeight("bold");
  
  sheetname.setColumnWidth(1, 45);
  sheetname.setColumnWidth(2, 35);
  sheetname.setColumnWidth(3, 80);
  sheetname.setColumnWidth(4, 120);
  sheetname.setColumnWidth(5, 135);
  sheetname.setColumnWidth(6, 135);
  
  sheetname.setFrozenRows(1);
}

// === Helper function - create formula via importrange function, return a string formula ===
function importSheetDataTot (workbook) {
  var sheetlist = workbook.getSheets();
  var sheetID = workbook.getId();
  
  var formula = 'sort(unique(ARRAYFORMULA({';
  for (var i = 0; i < sheetlist.length; i++) {
    var sheet = sheetlist[i]
//    Logger.log(sheet.getName())
    var lstrow = sheet.getMaxRows() - 11
    Logger.log(lstrow)
    if (i == 0) {
      var shrange = 'QUERY(IMPORTRANGE("' + sheetID + '","' + sheet.getName() + '!A6:G' + lstrow + '"), "select Col7, Col1, Col2, Col3, Col4, Col5")';
      Logger.log(shrange);
      formula = formula + shrange;
    } else if (sheet.getName() !== "Total" && sheet.getName() !== "Total_backup" && sheet.getName() != "CountBySubj" && sheet.getName() != "CountByBatch") {
      var shrange = ';QUERY(IMPORTRANGE("' + sheetID + '","' + sheet.getName() + '!A6:G' + lstrow + '"), "select Col7, Col1, Col2, Col3, Col4, Col5")';
      Logger.log(shrange);
      formula = formula + shrange;
    }
  }
  formula = formula + '})),1,TRUE,3,TRUE)';
  
  Logger.log(formula);
  return(formula)
}

function setTotalCol () {
  var mySS = SpreadsheetApp.getActive();
  var sheettol = mySS.getSheetByName("Total");
  var lstrowtot = sheettol.getLastRow();
  Logger.log(lstrowtot);
  
  var totrange = sheettol.getRange("C" + (lstrowtot+3) );
  totrange.setValue("Total")
  totrange.setFontWeight("bold")
  totrange.setHorizontalAlignment("center")
  
  var totrange1 = sheettol.getRange("D" + (lstrowtot+2) );
  totrange1.setValue("Issued");
  var totrange1Val = sheettol.getRange("D" + (lstrowtot+3) );
  totrange1Val.setFormula('sum(D2:D' + lstrowtot + ')');
  
  var totrange2 = sheettol.getRange("E" + (lstrowtot+2) );
  totrange2.setValue("Non-returned");
  var totrange2Val = sheettol.getRange("E" + (lstrowtot+3) );
  totrange2Val.setFormula('sumif(F2:F' + lstrowtot + ', "", D2:D' + lstrowtot + ')');
  
  var totrange3 = sheettol.getRange("F" + (lstrowtot+2) );
  totrange3.setValue("Returned");
  var totrange3Val = sheettol.getRange("F" + (lstrowtot+3) );
  totrange3Val.setFormula('D' + (lstrowtot+3) + '-E' + (lstrowtot+3));

  var totrangeAll = sheettol.getRange( "D" + (lstrowtot+2) + ":F" + (lstrowtot+3) )
  totrangeAll.setFontWeight("bold")
  totrangeAll.setHorizontalAlignment("center")
  
  // set format for border
  var borderrange1 = sheettol.getRange( "A2:F" + (lstrowtot) );
  borderrange1.setBorder(true, true, true, true, true, true);
  
  var borderrange2 = sheettol.getRange( "C" + (lstrowtot+2) + ":F" + (lstrowtot+3) );
  borderrange2.setBorder(null, null, null, null, null, true);
}

// === Create sheet Total Count ===
function createSheetCount () {
  var mySS = SpreadsheetApp.getActive();
  
  var sheets = mySS.getSheets();
  var sheetNames = [];
  for (var sheetno in sheets) {
    sheetNames.push(sheets[sheetno].getName());
    Logger.log(sheetno);
  }
  Logger.log(sheetNames);
  
  
  if ( sheetNames.indexOf("Total") === -1 ) {
    createSheetTotal();
    sheets = mySS.getSheets();
  } else if ( sheetNames.indexOf("CountBySubj") === -1 ) {
    sheets = mySS.getSheets();
    var sheetcount = mySS.insertSheet("CountBySubj", sheets.length);
  } else {
    sheets = mySS.getSheets();
    var rmsheet = mySS.getSheetByName("CountBySubj")
    mySS.deleteSheet(rmsheet);
    var sheetcount = mySS.insertSheet("CountBySubj", sheets.length);
  }
  
  sheetcount.setTabColor("purple");
  sheetcount.setFrozenRows(1);
  
  var colh1 = sheetcount.getRange("A1");
  colh1.setValue("Site");
  var colh2 = sheetcount.getRange("B1");
  colh2.setValue("Subject No.");
  var colh3 = sheetcount.getRange("C1");
  colh3.setValue("Total Queries");
  var colh3 = sheetcount.getRange("D1");
  colh3.setValue("Non-Returned Queries");
  
  var hrange = sheetcount.getRange("A1:D1");
  hrange.setBackgroundRGB(153, 204, 255);
  hrange.setVerticalAlignment("middle");
  hrange.setHorizontalAlignment("center")
  hrange.setFontWeight("bold");
  
  var sphrange = sheetcount.getRange("D1");
  sphrange.setBackground("red");
  
  sheetcount.setColumnWidth(1, 35);
  sheetcount.setColumnWidth(2, 80);
  sheetcount.setColumnWidth(3, 110);
  sheetcount.setColumnWidth(4, 180);
  
  var datarange1 = sheetcount.getRange("A2");
  var dataformula1 = 'sort(unique(Total!B2:C), 1, TRUE, 2, TRUE)';
  datarange1.setFormula(dataformula1);
  
  var datarange2 = sheetcount.getRange("C2");
  datarange2.setFormula('=SUMIF(Total!C$2:C,B2,Total!D$2:D)');
  
  var datarange3 = sheetcount.getRange("D2");
  datarange3.setFormula('SUMIFS(Total!D$2:D,Total!C$2:C,B2,Total!F$2:F,"")')
  
  var lstrowcount = sheetcount.getLastRow();
  var totrange1 = sheetcount.getRange("C" + (lstrowcount - 1) );
  totrange1.setFormula('sum(C2:C' + (lstrowcount - 2) + ')');
  var totrange2 = sheetcount.getRange("D" + (lstrowcount - 1) );
  totrange2.setFormula('sum(D2:D' + (lstrowcount - 2) + ')');
  
  var borderrange = sheetcount.getRange( "A1:D" + (lstrowcount-1) );
  borderrange.setBorder(true, true, true, true, true, true)
  
  var hlight = sheetcount.getRange("B" + (lstrowcount-1) )
  hlight.setHorizontalAlignment("center");
  hlight.setFontWeight("bold");
  hlight.setBackground("yellow");
  
}


// === Create sheet Count by Batch ===
function createSheetBatch () {
  var mySS = SpreadsheetApp.getActive();
  
  var sheets = mySS.getSheets();
  var sheetNames = [];
  for (var sheetno in sheets) {
    sheetNames.push(sheets[sheetno].getName());
    Logger.log(sheetno);
  }
  Logger.log(sheetNames);
  
  if ( sheetNames.indexOf("CountBySubj") === -1 ) {
    createSheetCount();
    sheets = mySS.getSheets();
  } else if ( sheetNames.indexOf("CountByBatch") === -1 ) {
    sheets = mySS.getSheets();
    var sheetbatch = mySS.insertSheet("CountByBatch", sheets.length);
  } else {
    sheets = mySS.getSheets();
    var rmsheet = mySS.getSheetByName("CountByBatch")
    mySS.deleteSheet(rmsheet);
    var sheetbatch = mySS.insertSheet("CountByBatch", sheets.length);
  }
  
  sheetbatch.setTabColor("cyan");
  sheetbatch.setFrozenRows(1);
  
  var colh1 = sheetbatch.getRange("A1");
  colh1.setValue("Batch");
  var colh2 = sheetbatch.getRange("B1");
  colh2.setValue("Total Queries");
  var colh3 = sheetbatch.getRange("C1");
  colh3.setValue("Non-Returned Queries");
  
  var hrange = sheetbatch.getRange("A1:C1");
  hrange.setBackgroundRGB(153, 204, 255);
  hrange.setVerticalAlignment("middle");
  hrange.setHorizontalAlignment("center")
  hrange.setFontWeight("bold");
  
  var sphrange = sheetbatch.getRange("C1");
  sphrange.setBackground("red");
  
  sheetbatch.setColumnWidth(1, 45);
  sheetbatch.setColumnWidth(2, 110);
  sheetbatch.setColumnWidth(3, 180);
  
  var datarange1 = sheetbatch.getRange("A2");
  var dataformula1 = 'sort(unique(filter(Total!A2:A, istext(Total!A2:A))), 1, TRUE, 2, TRUE)';
  datarange1.setFormula(dataformula1);
  
  var lstrowbatch = sheetbatch.getLastRow();
  
  var totcell = sheetbatch.getRange( "A" + (lstrowbatch+1) );
  totcell.setValue("Total");
  totcell.setHorizontalAlignment("center");
  totcell.setFontWeight("bold");
  totcell.setBackground("yellow");
  
  var datarange2 = sheetbatch.getRange("B2");
  datarange2.setFormula('=SUMIF(Total!A$2:A,A2,Total!D$2:D)');
  
  var datarange3 = sheetbatch.getRange("C2");
  datarange3.setFormula('SUMIFS(Total!D$2:D,Total!A$2:A,A2,Total!F$2:F,"")')
  
  var lstrowcount = sheetbatch.getLastRow();
  var totrange1 = sheetbatch.getRange("B" + (lstrowbatch+1) );
  totrange1.setFormula('sum(B2:B' + lstrowbatch + ')');
  var totrange2 = sheetbatch.getRange("C" + (lstrowbatch+1) );
  totrange2.setFormula('sum(C2:C' + lstrowbatch + ')');
  
  var borderrange = sheetbatch.getRange( "A1:C" + (lstrowbatch+1) );
  borderrange.setBorder(true, true, true, true, true, true)
  
}
