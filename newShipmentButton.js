var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var displayTrackerSheet = spreadsheet.getSheetByName("menues and elements");
// This is the key to finding the last row of your sheet. Use logger to see what your last row is.
var lastRow = displayTrackerSheet.getLastRow();

// **This is the main function linked to the Button. This runs the show**
function buttonPressed() {
  displaysVerification();
}; 

// This is where the magic happens. Due to my template, I had to come up with some workarounds to make it work.
function addNewShipment() {  

  // Since it did not paste on the right row, I made the 'var nextRowToPaste' and adding 2 solved the issue for me.
  var nextRowToPaste = lastRow + 2;
  var firstGroupSelection = nextRowToPaste;
  var secondGroupSelection = firstGroupSelection + 4;
  
  // To group the cells, the 'getRange()' method required a string in this formart --> "20:22"
  var beginGrouping = firstGroupSelection.toString() + ":" + secondGroupSelection.toString();

  // Make sure to use loggers to see what number the method 'lastRow();' gives you.
  // IMPORTANT: Test your code before adding the dialog. Dialogs cannot run in the editor. It runs ONLY when user action is present from sheets.
  Logger.log("First Group = " + firstGroupSelection);
  Logger.log("Second Group = " + secondGroupSelection);
  Logger.log("beginGrouping = " + beginGrouping);

  displayTrackerSheet.getRange('A' + nextRowToPaste).activate();

  spreadsheet.getRange('\'<--Template\'!B2:K8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
  // What goes in these methods depends on how and what you are pasting.
  displayTrackerSheet.getRange(beginGrouping).activate();
  displayTrackerSheet.setCurrentCell(spreadsheet.getRange('A' + nextRowToPaste));
  displayTrackerSheet.getActiveRange().shiftRowGroupDepth(1);
  displayTrackerSheet.getRowGroup(secondGroupSelection + 1, 1).collapse();

  Logger.log("Updated last row number" + nextRowToPaste);
}

/* Created an alert for the user to avoid spams and to confirm the current shipment is done.
  * The user cannot add a new shipment until all display stock has been used and DOAs reported.
*/ 
function displaysVerification() {

  var displayStockValue = displayTrackerSheet.getRange('H' + lastRow).getValue();
  var doaReportChecked = displayTrackerSheet.getRange('J' + lastRow).getDisplayValue();
  var countedFor = displayTrackerSheet.getRange('E' + lastRow).getDisplayValue();

  if (doaReportChecked == "FALSE" || doaReportChecked == "TRUE" && displayStockValue != 0 && countedFor >= 5) {
    cannotAddShipmentAlert();
    displayTrackerSheet.getRange('J' + lastRow).setValue("FALSE")
  } else {
      confirmAddShipmentAlert();
    }
}

// Created an alert for the user to confirm adding a new shipment after 'displaysVerification();' checks current shipment is done.
function confirmAddShipmentAlert() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to add a new shipment?',
      ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    ui.alert('Adding shipment!');
    // if yes is clicked, the function that automates the pasting and grouping is called.
    addNewShipment();
  } else {
    ui.alert('No shipment added.');
  }
}

function cannotAddShipmentAlert() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Cannot add shipment','Please use all the display stock and make sure DOAs are reported.', ui.ButtonSet.OK);
}

