/** SO Scoring System for Google Sheets
*   Author: Martin Nguyen
*   
*   This is essentially the SO Scoring System developed by Alan Chalker
*   on Excel moved into Google Sheets so that proctors and graders can 
*   easily input their grades in instead of emailing or importing other workbooks.
*
*   Everything is the exact same, except that the count number of unique ranks
*   formula in the event sheets had to be modified so that it could work on Sheets.
*   Since macros from Excel does not transfer over, it is simply reprogrammed into
*   App Scripts and should work similar to its original purpose. However, only people
*   with editing rights can run scripts, so if they are view only they can not update the sheets.
*
*   Excel Score System from Alan Chalker: https://sourceforge.net/p/soscoring/
*/

//goes to a certain tab in the workbook, used in buttons for navigation
function goTo(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  SpreadsheetApp.setActiveSheet(sheet);
}

function goToStart() {
  goTo("Start Here");
}

function goToSetup() {
  goTo("Setup");
}

function goToMaster() {
  goTo("Master");
}

function goToMedals() {
  goTo("Printable Medals List");
}

function goToPrintBlankScoresheet() {
  goTo("Printable Blank Score Sheet");
}

function goToBlankScoresheet() {
  goTo("Blank Score Sheet");
}

//prompts user if they want to update the events
//if they do, run actualSetUp() which will update events
//used in renaming or adding/removing events
//WILL REMOVE CURRENT SCORE DATA
function setUpEvents() {
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("In order to change the event names, " +
                               "you must first change them in the list on the right hand side of this tab. " +
                               "Note all existing event data will be deleted! " + "Are you sure you want to do this?",
                              ui.ButtonSet.YES_NO);
  if(buttonPressed == ui.Button.YES) {
    actualSetUp();
  }
}

//actual setting up
function actualSetUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //get number of events
  var numEvents = ss.getSheetByName("Setup").getRange('H31').getValue();
  
  //delete spreadsheets 6 and onward (all events after blank score sheet)
  //array starts at 0 so current scoresheets are >6
  var allSheets = ss.getSheets();
  while(allSheets[6] != null) {
    allSheets[6].activate();
    ss.deleteActiveSheet();
    allSheets = ss.getSheets();
  }
  
  //clear all hyperlinks on Start Here spreadsheet
  ss.getSheetByName("Start Here").getRange('B3:B34').clear({formatOnly: false, contentsOnly: true});
  
  //main loop: going backwards, duplicate the blank score sheet,
  //put event name and rename sheet and change tab color,
  //then set up hyperlink on Start Here spreadsheet
  for(var count = numEvents; count > 0; count--) {
    var eventName = ss.getSheetByName("Setup").getRange(2 + count, 6).getValue();
    var scoreSheet = ss.getSheetByName("Blank Score Sheet");
    SpreadsheetApp.setActiveSheet(scoreSheet);
    SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
    var currentSheet = ss.getSheetByName("Copy of Blank Score Sheet");
    currentSheet.setTabColor("green");
    currentSheet.getRange('L3').setValue(eventName);
    currentSheet.setName(eventName);
    var sheetID = currentSheet.getSheetId();
    currentSheet = ss.getSheetByName("Start Here");
    currentSheet.getRange(2 + count, 2).setValue('=HYPERLINK(\"#gid=' + sheetID + '\", \"' + eventName + '\")');
  }
  
  //hide/unhide columns in master sheet if necessary
  var currentSheet = ss.getSheetByName("Master");
  for(var count = 1; count < 33; count++) {
    if(count > numEvents) {
      currentSheet.hideColumns(3 + count);
    } else {
      var range = currentSheet.getRange(1, 3 + count);
      currentSheet.unhideColumn(range);
    }
  }
  
  //WORKS SO FAR UP TO HERE YAY
  
  //do some row hiding/unhiding on medal sheet
  
  //copy F3:F34 to M3:M34
  currentSheet = ss.getSheetByName("Setup");
  currentSheet.getRange('F3:F34').copyTo(currentSheet.getRange('M3:M34'));
}

//changes number of teams, which results in setting max points
//for noshow and participation
//also hiding appropriate rows in scoresheets and master scoresheet
//as well in setup
function changeNumTeams() {
  
}
