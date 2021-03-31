var spreadsheet = SpreadsheetApp.getActive();
var displayTackerSheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Pixel Display Tracker'), true);

// This is the key to finding the last row of your sheet. Use logger to see what your last row is.
var lastRow = displayTackerSheet.getLastRow();

// This is the main function linked to the Button. This runs the show. 
function buttonPressed() {
  showAlert();
}; 

// Created an alert for the user to avoid accidental clicks or spams.
function showAlert() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to add a new shipment?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    ui.alert('Adding shipment!');
    // if yes is clicked, the function that automates the pasting and grouping is called.
    addNewShipment();
  } else {
    ui.alert('No shipment added.');
  }
}

// This is were the magic happens. Due to my template, I had to come up with some workarounds to make it work.
function addNewShipment() {
  
  // 2 was added because the last row number was 2 rows off. It did not paste on the right row, adding 2 solved the issue for me
  var nextRowToPaste = lastRow + 2;
  var firstGroupSelection = nextRowToPaste;
  var secondGroupSelection = firstGroupSelection + 2;
  
  // To group the cells, the 'getRange()' method required a string in this formart --> "20:22"
  var beginGrouping = firstGroupSelection.toString() + ":" + secondGroupSelection.toString();

  // Make sure to use loggers to see what last row number the method gives you.
  // IMPORTANT: Test your code before adding the dialog. Dialogs cannot run in the editor. It runs ONLY when user action is present from sheets.
    Logger.log(nextRowToPaste);
    Logger.log("First Group = " + firstGroupSelection);
    Logger.log("Second Group = " + secondGroupSelection);
    Logger.log("beginGrouping = " + beginGrouping);

    displayTackerSheet.getRange('A' + nextRowToPaste).activate();

    spreadsheet.getRange('\'<--Template\'!B2:K6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
  // What goes in the methods depends on how and what you are pasting.
    displayTackerSheet.getRange(beginGrouping).activate();
    displayTackerSheet.setCurrentCell(spreadsheet.getRange('A' + nextRowToPaste));
    displayTackerSheet.getActiveRange().shiftRowGroupDepth(1);
    displayTackerSheet.getRowGroup(secondGroupSelection + 1, 1).collapse();

    Logger.log("Updated last row number" + nextRowToPaste);
}


