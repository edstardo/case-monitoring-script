var as = SpreadsheetApp.getActiveSpreadsheet();

var ml = as.getSheetByName("Case Master List");

var mlEntryRowStart = 16;

var subjectField      = ml.getRange('D3');
var caseTitleField    = ml.getRange('D4')
var caseNumberField   = ml.getRange('D5');
var forTypeField      = ml.getRange('D6');
var assignedByField   = ml.getRange('D7');
var dateAssignedField = ml.getRange('D8');
var courtField        = ml.getRange('D9');
var lastTouchField    = ml.getRange('D10');
var statusField       = ml.getRange('D11');

function getMasterlist() {
  var cases = [];

  var i = mlEntryRowStart;
  var caseID = ml.getRange("B".concat(i.toString())).getValue();

  while (caseID != "") {
    cases.push(
      {
        "caseID":       caseID,
        "subject":      ml.getRange("C".concat(i.toString())).getValue(),
        "caseTitle":    ml.getRange("D".concat(i.toString())).getValue(),
        "caseNumber":   ml.getRange("E".concat(i.toString())).getValue(),
        "for":          ml.getRange("F".concat(i.toString())).getValue(),
        "assignedBy":   ml.getRange("G".concat(i.toString())).getValue(),
        "dateAssigned": ml.getRange("H".concat(i.toString())).getValue(),
        "court":        ml.getRange("I".concat(i.toString())).getValue(),
        "lastTouch":    ml.getRange("J".concat(i.toString())).getValue(),
        "status":       ml.getRange("K".concat(i.toString())).getValue(),
      }
    )
    
    i++;
    caseID = ml.getRange("B".concat(i.toString())).getValue()
  }

  return cases;
}

function goToCaseMasterList() {
  SpreadsheetApp.setActiveSheet(ml);
}
