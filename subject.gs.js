var subjects = [
    {"subject": "Civil",                "field": ml.getRange("H3"),  "numberOfCases": 0, "sheet": null},
    {"subject": "Criminal",             "field": ml.getRange("H4"),  "numberOfCases": 0, "sheet": null}, 
    {"subject": "Labor",                "field": ml.getRange("H5"),  "numberOfCases": 0, "sheet": null},
    {"subject": "Special Proceeding",   "field": ml.getRange("H6"),  "numberOfCases": 0, "sheet": null},
    {"subject": "Tax",                  "field": ml.getRange("H7"),  "numberOfCases": 0, "sheet": null},
    {"subject": "Administrative",       "field": ml.getRange("H8"),  "numberOfCases": 0, "sheet": null},
    {"subject": "Land Registration",    "field": ml.getRange("H9"),  "numberOfCases": 0, "sheet": null},
    {"subject": "Special Civil Action", "field": ml.getRange("H10"), "numberOfCases": 0, "sheet": null},
    {"subject": "Others",               "field": ml.getRange("H11"), "numberOfCases": 0, "sheet": null}
  ]
  
  var subjectSheetStartRow = 5;
  
  function updateSubjectSheets() {
    var cases = getMasterlist();
  
    clearSubjectsTable();
    deleteSubjectsSheets();
    
    for (let i = 0; i < cases.length; i++) {
      for (let j = 0; j < subjects.length; j++) {
        if (cases[i].subject == subjects[j].subject) {
          subjects[j].numberOfCases++;
          break;
        }
      }
    }
    
    createSubjectSheetsWithCases();
    createSubjectSheetsLinks();
    populateSubjectSheets(cases, subjects);
    // protectSubjectSheets(subjects);
  }
  
  function createSubjectSheetsLinks() {
    for (let i = 0; i < subjects.length; i++) {
      if(subjects[i].sheet != null) {
        var richValue = SpreadsheetApp.newRichTextValue()
          .setText("VIEW")
          .setLinkUrl(getSheetURL(subjects[i].sheet))
          .build();
        subjects[i].field.setRichTextValue(richValue);
      }
    }
  }
  
  function protectSubjectSheets(subjects) {
    for (let i = 0; i < subjects.length; i++) {
      if (subjects[i].sheet != null) {
        
        // var range = subjects[i].sheet.getRange('B5:L1002');
        // var protection = range.protect().setDescription(subjects[i].subject);
        var protection = subjects[i].sheet.protect().setDescription(subjects[i].subject);
        
        var me = Session.getEffectiveUser();
        protection.addEditor(me);
        protection.removeEditors(protection.getEditors());
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
      }
    }
    // reference: https://developers.google.com/apps-script/reference/spreadsheet/protection
  }
  
  function populateSubjectSheets(cases, subjects) {
    for (let i = 0; i < subjects.length; i++) {
      var row = subjectSheetStartRow;
  
      if (subjects[i].sheet == null) {
        continue;
      }
  
      for (let j = 0; j < cases.length; j++) {
        if (cases[j].subject == subjects[i].subject) {
          var caseSheet = as.getSheetByName(cases[j].caseID.toString());
          var richValue = SpreadsheetApp.newRichTextValue()
            .setText(cases[j].caseID.toString())
            .setLinkUrl(getSheetURL(caseSheet))
            .build();
          subjects[i].sheet.getRange("B".concat(row.toString())).setRichTextValue(richValue);
          subjects[i].sheet.getRange("C".concat(row.toString())).setValue(cases[j].subject)
          subjects[i].sheet.getRange("D".concat(row.toString())).setValue(cases[j].caseTitle)
          subjects[i].sheet.getRange("E".concat(row.toString())).setValue(cases[j].caseNumber)
          subjects[i].sheet.getRange("F".concat(row.toString())).setValue(cases[j].for)
          subjects[i].sheet.getRange("G".concat(row.toString())).setValue(cases[j].assignedBy)
          subjects[i].sheet.getRange("H".concat(row.toString())).setValue(cases[j].dateAssigned)
          subjects[i].sheet.getRange("I".concat(row.toString())).setValue(cases[j].court)
          subjects[i].sheet.getRange("J".concat(row.toString())).setValue(cases[j].lastTouch)
          subjects[i].sheet.getRange("K".concat(row.toString())).setValue(cases[j].status)
  
          row++;
        }
      }
    }
  }
  
  function createSubjectSheetsWithCases() {
    for (let i = 0; i < subjects.length; i++) {
      if (subjects[i].numberOfCases != 0) {
        subjects[i].sheet = createNewSubjectSheet(subjects[i].subject);
      }
    }
  }
  
  function createNewSubjectSheet(subject) {
    var template = as.getSheetByName('Subject Template');
    template.copyTo(as).setName(subject.toString());
    return as.getSheetByName(subject.toString());
  }
  
  function deleteSubjectsSheets() {
    for (let i = 0; i < subjects.length; i++) {
      var sc = as.getSheetByName(subjects[i].subject);
      if (sc == null) {
        continue
      }
      as.deleteSheet(sc)
    }
  }
  
  function clearSubjectsTable() {
    for (let i = 0; i < subjects.length; i++) {
      subjects[i].field.setValue("");
    }
  }
  