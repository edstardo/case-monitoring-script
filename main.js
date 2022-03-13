var as = SpreadsheetApp.getActiveSpreadsheet();

var ml = as.getSheetByName("Case Master List");
var caseTitleField    = ml.getRange('E3');
var forTypeField      = ml.getRange('E4');
var assignedByField   = ml.getRange('E5');
var dateAssignedField = ml.getRange('E6');
var courtField        = ml.getRange('E7');
var lastTouchField    = ml.getRange('E8');
var statusField       = ml.getRange('E9');

function addNewCase() {
  // get new case info
  var caseTitle    = caseTitleField.getValue();
  var forType      = forTypeField.getValue();
  var assignedBy   = assignedByField.getValue();
  var dateAssigned = dateAssignedField.getValue();
  var court        = courtField.getValue();
  var lastTouch    = lastTouchField.getValue();
  var status       = statusField.getValue();
  
  // validate new case info
  if (!(caseTitle && forType && assignedBy && dateAssigned && court && lastTouch && status)) {
    displayInvalidNewCase();
    return;
  }

  // save new case
  var newCaseRow = getNewCaseRowNumber()
  var caseID = newCaseRow-12+1
  ml.getRange("B".concat(newCaseRow.toString())).setValue(caseID)
  ml.getRange("C".concat(newCaseRow.toString())).setValue(caseTitle)
  ml.getRange("E".concat(newCaseRow.toString())).setValue(forType)
  ml.getRange("F".concat(newCaseRow.toString())).setValue(assignedBy)
  ml.getRange("G".concat(newCaseRow.toString())).setValue(dateAssigned)
  ml.getRange("H".concat(newCaseRow.toString())).setValue(court)
  ml.getRange("I".concat(newCaseRow.toString())).setValue(lastTouch)
  ml.getRange("J".concat(newCaseRow.toString())).setValue(status)
  

  // create new case sheet
  var nc = createNewCaseSheet(caseID)
  nc.getRange("D3").setValue(caseID)
  nc.getRange("D4").setValue(caseTitle)
  nc.getRange("D5").setValue(forType)
  nc.getRange("D6").setValue(assignedBy)
  nc.getRange("D7").setValue(dateAssigned)
  nc.getRange("D8").setValue(court)
  nc.getRange("D9").setValue(lastTouch)
  nc.getRange("D10").setValue(status)

  var ncSheetLink = ml.getRange("K".concat(newCaseRow.toString()));
  var richValue = SpreadsheetApp.newRichTextValue()
    .setText("VIEW")
    .setLinkUrl(getSheetURL(nc))
    .build();
  ncSheetLink.setRichTextValue(richValue);

  // clear new case form
  caseTitleField.setValue("");
  forTypeField.setValue("");
  assignedByField.setValue("");
  dateAssignedField.setValue("");
  courtField.setValue("");
  lastTouchField.setValue("");
  statusField.setValue("");
}

function getSheetURL(sheet) {
  var url = '';
  url += as.getUrl();
  url += '#gid=';
  url += sheet.getSheetId(); 
  return url;
}

function getNewCaseRowNumber() {
  var row = 0;
  var look = 12;
  var prevVal = "";
  while (row == 0) {
    val = ml.getRange("B".concat(look.toString())).getValue()
    if (!val) {
      row = look;
    }
    look++;
  }
  return row;
}

function displayInvalidNewCase() {
  var result = SpreadsheetApp.getUi().alert("Invalid Case Info!");
  // if(result === SpreadsheetApp.getUi().Button.OK) {
  //   //Take some action
  //   SpreadsheetApp.getActive().toast("About to take some action …");
  // }
}

function createNewCaseSheet(caseID) {
  var nc = as.getSheetByName(caseID.toString());

  if (nc == null) {
    var template = as.getSheetByName('Case Template');

    template.copyTo(as).setName(caseID.toString());
  }

  nc = as.getSheetByName(caseID.toString());
  SpreadsheetApp.setActiveSheet(nc);

  return nc
}

function goToCaseMasterList() {
  SpreadsheetApp.setActiveSheet(ml);
}
