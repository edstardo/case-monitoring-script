function addNewCase() {
    // get new case info
    var subject      = subjectField.getValue();
    var caseTitle    = caseTitleField.getValue();
    var caseNumber   = caseNumberField.getValue();
    var forType      = forTypeField.getValue();
    var assignedBy   = assignedByField.getValue();
    var dateAssigned = dateAssignedField.getValue();
    var court        = courtField.getValue();
    var lastTouch    = lastTouchField.getValue();
    var status       = statusField.getValue();
    
    // validate new case info
    if (!(subject && caseTitle && caseNumber && forType && assignedBy && dateAssigned && court && lastTouch && status)) {
      displayInvalidNewCase();
      return;
    }
    
    var newCaseRow = getNewCaseRowNumber()
    var caseID = newCaseRow-mlEntryRowStart+1
  
    // create new case sheet
    var nc = createNewCaseSheet(caseID)
  
    // set values
    nc.getRange("D3").setValue(caseID)
    nc.getRange("D4").setValue(subject)
    nc.getRange("D5").setValue(caseTitle)
    nc.getRange("D6").setValue(caseNumber)
    nc.getRange("D7").setValue(forType)
    nc.getRange("D8").setValue(assignedBy)
    nc.getRange("D9").setValue(dateAssigned)
    nc.getRange("D10").setValue(court)
    nc.getRange("D11").setValue(lastTouch)
    nc.getRange("D12").setValue(status)
  
    // save new case
    // create links from master list to case sheets
    var ncSheetLink = ml.getRange("B".concat(newCaseRow.toString()));
    var richValue = SpreadsheetApp.newRichTextValue()
      .setText(caseID.toString())
      .setLinkUrl(getSheetURL(nc))
      .build();
    ncSheetLink.setRichTextValue(richValue);
    ml.getRange("C".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D4"))
    ml.getRange("D".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D5"))
    ml.getRange("E".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D6"))
    ml.getRange("F".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D7"))
    ml.getRange("G".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D8"))
    ml.getRange("H".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D9"))
    ml.getRange("I".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D10"))
    ml.getRange("J".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D11"))
    ml.getRange("K".concat(newCaseRow.toString())).setValue("='".concat(caseID.toString()).concat("'!D12"))
  
    // clear new case form
    subjectField.setValue("");
    caseTitleField.setValue("");
    caseNumberField.setValue("");
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
    var look = mlEntryRowStart;
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
  