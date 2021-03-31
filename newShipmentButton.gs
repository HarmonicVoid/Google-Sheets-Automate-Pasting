var spreadsheet = SpreadsheetApp.getActive();
var displayTackerSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pixel Display Tracker'), true);
var lastRow = displayTackerSheet.getLastRow();

function buttonPressed() {
  showAlert();
}; 

function showAlert() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to add a new shipment?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    ui.alert('Adding shipment!');
    addNewShipment();
  } else {
    ui.alert('No shipment added.');
  }
}

function addNewShipment() {
  var nextRowToPaste = lastRow + 2;
  var firstGroupSelection = nextRowToPaste;
  var secondGroupSelection = firstGroupSelection + 2;
  var beginGrouping = firstGroupSelection.toString() + ":" + secondGroupSelection.toString();

    Logger.log(nextRowToPaste);
    Logger.log("First Group = " + firstGroupSelection);
    Logger.log("Second Group = " + secondGroupSelection);
    Logger.log("beginGrouping = " + beginGrouping);

    displayTackerSheet.getRange('A' + nextRowToPaste).activate();

    spreadsheet.getRange('\'<--Template\'!B2:K6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
    displayTackerSheet.getRange(beginGrouping).activate();
    displayTackerSheet.setCurrentCell(spreadsheet.getRange('A' + nextRowToPaste));
    displayTackerSheet.getActiveRange().shiftRowGroupDepth(1);
    displayTackerSheet.getRowGroup(secondGroupSelection + 1, 1).collapse();

    Logger.log("Updated row to paste:" + nextRowToPaste);
}


