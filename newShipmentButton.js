/*The key to finding the last row of your sheet --> var lastRow = displayTrackerSheet.getLastRow(). 
  Call it in each individual function that requires the last row becuase 
  we want to grab the last row after checkIfRowsAdded(); is called. 
  Use logger to see what your last row number is to make sure its the correct row you want your template pasted.

  IMPORTANT: lastRow will be a different row number value if the user decides to type into a random cell and row location.
  For example, instead of getting the next row we want to paste the template in,
  the user types something into a cell in row 100, now the lastRow is 100 numbers away the row we want the template to be pasted.
  This issue has been solved by adding extra checks and by saving the correct last row postion after checkIfRowsAdded() has been called.
*/

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const displayTrackerSheet = spreadsheet.getSheetByName("Pixel Display Log");
const savedVariablesSheet = spreadsheet.getSheetByName("Saved Variables");

// **This is the main function linked to the Button. This runs the script**
function buttonClicked() {
  rowsAdded();
  displaysVerification();
}; 

function overrideButton() {
  rowsAdded();
  passwordPrompt();

}

function passwordPrompt() {
  var typeString;
  
  //Creating the prompt.
  var ui = SpreadsheetApp.getUi(); 
  var result = ui.prompt(
      'New Shipment Override',
      'Please enter password:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var userTypedIMEI = result.getResponseText();

  //User clicked "OK".
  if(button == ui.Button.OK) {   
    typeString = verifyPassword(userTypedIMEI);
    
    if(typeString == "WRONG") {
      ui.alert('NOT VALID!');
      passwordPrompt();
    } else if(typeString == "CORRECT"){
      ui.alert('Verified :)');

      confirmAddShipmentAlert();
    } 

  } else if(button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    closedToast();
  } else if(button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    closedToast();
  }
}

function verifyPassword(inputData) {

  if (inputData == "qu!ckfix"){
    return "CORRECT";
  } else {
    return "WRONG";
  }

}

// This will check if any rows got added after the template. If it does, it will delete the extra rows and typed values.
function rowsAdded() {
  /* We use a saved lastRow value becuase we do not want the lastRow from the user if they typed on a cell in row 100. 
    Instead we grab the saved correct last row position that changes until we update it by calling the function saveLastRow().
    Cell 'B3' is where we save the correct last row postion.
  */
  var savedLastRow = savedVariablesSheet.getRange('B2').getValue();
  var getMaxRows = displayTrackerSheet.getMaxRows();
  Logger.log(getMaxRows);

  if(getMaxRows >= savedLastRow) {
    Logger.log("New rows were added!!!! Deleting...");
    displayTrackerSheet.getRange(savedLastRow + ':' + savedLastRow).activate();
    spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  }
};

/* The user cannot add a new shipment until all display stock has been used and DOAs reported.
   There was an issue when a new template was added, the user would be able to bypass "current shipment is done" check.
   This nasty bug has been resolved succesfully by adding new checks to see if the template format is present in a unique way.
   Now the user will get a new alert stating "You just added a new shipment" if the template format is present.
*/ 
function displaysVerification() {
  var lastRow = displayTrackerSheet.getLastRow();
  var lastRowToCheck = lastRow;
  Logger.log(lastRow, lastRowToCheck);
  var displayStockValue = displayTrackerSheet.getRange('H' + lastRowToCheck).getValue();
  var doaReportCheckBox = displayTrackerSheet.getRange('J' + lastRowToCheck).getDisplayValue();
  var acountedFor = displayTrackerSheet.getRange('E' + lastRowToCheck).getDisplayValue();
  var acountedFor2 = displayTrackerSheet.getRange('B' + lastRowToCheck).getDisplayValue();
  
   if(acountedFor2 == "Name/Order#" || acountedFor2 == "" && doaReportCheckBox == "TRUE" || doaReportCheckBox == "FALSE" && acountedFor <= 20  ) {
       cannotAddIfTemplateAlert();
       displayTrackerSheet.getRange('J' + lastRowToCheck).setValue("FALSE")

   } else if(doaReportCheckBox == "TRUE" &&  displayStockValue != 0) {
     cannotAddShipmentAlert2();
     displayTrackerSheet.getRange('J' + lastRowToCheck).setValue("FALSE")

   } else if(doaReportCheckBox == "FALSE" &&  displayStockValue != 0) {
    cannotAddShipmentAlert();
     //displayTrackerSheet.getRange('J' + lastRowToCheck).setValue("FALSE")

  } else if(doaReportCheckBox == "FALSE" &&  displayStockValue == 0) {
     cannotAddShipmentAler3();
     //displayTrackerSheet.getRange('J' + lastRowToCheck).setValue("FALSE")

  } else if(doaReportCheckBox == "TRUE") {
      confirmAddShipmentAlert();
  }
}

// Created an alert for the user to confirm adding a new shipment after 'displaysVerification();' checks current shipment is done.
function confirmAddShipmentAlert() {


  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     'Please confirm',
     'Are you sure you want to add a new shipment?',
      ui.ButtonSet.YES_NO);

  if(result == ui.Button.YES) {
    ui.alert('Adding shipment!');
    // if yes is clicked, the function that automates the pasting and grouping is called.
    addNewShipment();

  
    // Very important function to make sure we always have the correct last row tgat does not update untill we call it.
    saveLastRow();
  } else {
    ui.alert('No shipment added.');
  }
}


// This is where the magic happens. Due to my template, I had to come up with some workarounds to make it work.
function addNewShipment() {  
  var lastRow = displayTrackerSheet.getLastRow();
  // Since it did not paste on the right row, I made the 'var nextRowToPaste' and adding 2 solved the issue for me.
  var nextRowToPaste = lastRow + 2;
  var firstGroupSelection = nextRowToPaste;
  var secondGroupSelection = firstGroupSelection + 4;
  // To group the cells, the 'getRange()' method required a string in this formart --> "20:22"
  var beginGrouping = firstGroupSelection.toString() + ":" + secondGroupSelection.toString();

  addNextRowsToPaste();

  // Make sure to use loggers to see what number the method 'lastRow();' gives you.
  // IMPORTANT: Test your code before adding the dialog. Dialogs cannot run in the editor. It runs ONLY when user action is present from sheets.
  Logger.log("First Group = " + lastRow);
  Logger.log("Second Group = " + secondGroupSelection);
  Logger.log("beginGrouping = " + beginGrouping);

  displayTrackerSheet.getRange('A' + nextRowToPaste).activate();
  spreadsheet.getRange('\'<--Template\'!B2:K8').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    
  // What goes in these methods depends on how and what you are pasting.
  displayTrackerSheet.getRange(beginGrouping).activate();
  displayTrackerSheet.setCurrentCell(spreadsheet.getRange('A' + nextRowToPaste));
  displayTrackerSheet.getActiveRange().shiftRowGroupDepth(1);
  displayTrackerSheet.getRowGroup(secondGroupSelection + 1, 1).collapse();
}

// Adss 7 new rows to make space for the pasting proccess.
function addNextRowsToPaste(){
  displayTrackerSheet.getRange('A19:J19').activate();
  displayTrackerSheet.insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), 7);
};


function saveLastRow() {
//Here we are getting the new updated last row and saving it. 
    // Please note: we know this is the correct last row to save becuase the function checkIfRowsAdded() was alreay called.
    var newLastRow = displayTrackerSheet.getLastRow();
    var currentLastRow = newLastRow + 2;
    savedVariablesSheet.getRange('B2').activate();
    savedVariablesSheet.getCurrentCell().clear().setValue(currentLastRow);
    displayTrackerSheet.getRange('A1').activate();

}

function cannotAddShipmentAlert() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Cannot add shipment','Please use all the display stock and make sure DOAs are reported.', ui.ButtonSet.OK);
}
function cannotAddShipmentAlert2() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Cannot add shipment','Please use all the display stock.', ui.ButtonSet.OK);
}


function cannotAddShipmentAler3() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Cannot add shipment','Please make sure DOAs are reported.', ui.ButtonSet.OK);
}

function cannotAddIfTemplateAlert() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Cannot add shipment','You just added a new shipment.', ui.ButtonSet.OK);
}
