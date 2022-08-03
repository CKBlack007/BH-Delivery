function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('BH Menu')
      .addItem('Get Travel Times', 'TravelTime')
      .addToUi();
}

function CTE (){
AuditTimes();  //copies to audit sheet
TravelTime();  //gets new travel estimates
  
}

function CompareTimes (){
//compare row over row, if both cells are not blank. Alarm via sms if true
var ss = SpreadsheetApp.getActiveSpreadsheet();         //cause it's google
var sourceSheet = ss.getSheetByName("rTimeLog");         //take a sheet
var lastRow = sourceSheet.getLastRow();                 //ummm gimme the last row
for (var i = 1; i < lastRow-1; i++) {  

}
}

function TravelTime(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName("Engine");
var lastRow = sourceSheet.getLastRow();
for (var i = 1; i < lastRow-1; i++) {
  var sourceRange = sourceSheet.getRange("AK"+(i+1)).getValue();     //gets api url
  var padTime = sourceSheet.getRange("AD"+(i+1)).getValue();         //gets the delivery time padding
  if (sourceRange != ""){
    var travelJSON = UrlFetchApp.fetch(sourceRange);                   //fetch the json
    var travelObject = JSON.parse(travelJSON.getContentText());        //parse the string into an object that I can access
    var travelTime = travelObject.routes[0].legs[0].duration.value/60; //time to travel in minutes
    var travelTime = Math.round(travelTime);                           //rounds travelTime to a whole number
    var travelTime = travelTime + padTime;
    sourceSheet.getRange("R"+(i+1)).activate().setValue(travelTime);   //sets the time
    //Utilities.sleep(1000);
  } 
  } 
  }
function AuditTimes() {
  /* I need to add an if statement here that only runs this script on dates relative to its period */
var ss = SpreadsheetApp.getActiveSpreadsheet();
var target = SpreadsheetApp.getActiveSpreadsheet();
var source_sheet = ss.getSheetByName("Engine");
var target_sheet = target.getSheetByName("rTimeLog");
 var source_range = source_sheet.getRange("R2:R");
//var target_range = target_sheet.getRange("A1:A1");
var last_col = target_sheet.getLastColumn();
//Logger.log(last_row);
target_sheet.insertColumnAfter(last_col);
var post_time = new Date();  
  var current_time = post_time.toLocaleTimeString(); 
  Logger.log(last_col);
target_sheet.getRange(1,(last_col+1)).activate().setValue(post_time);    

var target_range = target_sheet.getRange(2,(last_col+1));
source_range.copyTo(target_range,{contentsOnly: true});
}

// sheetname created to get date folder   
function BookName() {
    return SpreadsheetApp.getActive().getName();
}

function cleanCalendarCreator() {
  var spreadsheet = SpreadsheetApp.getActive();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('Calendar Creator'))
  spreadsheet.getRange('I2:L61').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('N2:Q61').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function cleanInputSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheetByName('input'))
  spreadsheet.getRange('A3:O31').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('E2').activate();
  spreadsheet.getCurrentCell().setValue('unassigned');
  spreadsheet.getRange('J2:J6').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('J6'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
  spreadsheet.getRange('C2:C5').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('A1').activate();
  //now setting the cop date ahead by one week
   spreadsheet.getRange('A1').activate();
};