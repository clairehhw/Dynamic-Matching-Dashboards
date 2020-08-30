//Function: Get information about interview process from all project dashboards
function trackingProgress() {  
  var projectsBook = SpreadsheetApp.openById('1zCmXNJ0ytEcX_YR7f2QIms0gpsjItwKfgYHlBjfHojY').getSheetByName("Project Overview") //FA20 Project Tracking
  var num_projects = projectsBook.getLastRow();
  var col_URL = 8;
  for (var i = 2; i < num_projects; i++) {
    var accepted = 0,
        rejected_ni = 0,
        rejected_i = 0,
        interviewing = 0;
    var url = projectsBook.getRange(i, col_URL).getValue();
    var projectSheet = SpreadsheetApp.openByUrl(url).getSheetByName("StudentInfo");
    var num_applicants = SpreadsheetApp.openByUrl(url).getSheetByName("Codes").getRange("A9").getValue();
    if (num_applicants != 0){
      for (var app = 2; app < num_applicants + 2; app++) {
        var status = projectSheet.getRange("A" + app).getValue();
        if (status == "Accepted"){
        accepted++;
        } else if (status == "Rejected-No interview"){
          rejected_ni++;
        } else if (status == "Rejected-Interviewed"){
          rejected_i++;
        } else {
          interviewing++;
        }
      }
      projectsBook.getRange("I" + i).setValue(num_applicants);
      projectsBook.getRange("J" + i).setValue(accepted);
      projectsBook.getRange("K" + i).setValue(rejected_ni);
      projectsBook.getRange("L" + i).setValue(rejected_i);
      projectsBook.getRange("M" + i).setValue(interviewing);
    } else {
      projectsBook.getRange("I" + i).setValue(0);
      projectsBook.getRange("J" + i).setValue(0);
      projectsBook.getRange("K" + i).setValue(0);
      projectsBook.getRange("L" + i).setValue(0);
      projectsBook.getRange("M" + i).setValue(0);
    }
    Logger.log("row " + i + "updated")
  }
}


//Function: Generate a roster with all student-project matching status 
function trackingRoster(){
  var progressOverview = SpreadsheetApp.openById('1zCmXNJ0ytEcX_YR7f2QIms0gpsjItwKfgYHlBjfHojY').getSheetByName("Project Overview") //FA20 Project Tracking
  var progressRoster = SpreadsheetApp.openById('1zCmXNJ0ytEcX_YR7f2QIms0gpsjItwKfgYHlBjfHojY').getSheetByName("TrackingALL") 
  var num_projects = progressOverview.getLastRow() - 1;
  var projNames = progressOverview.getRange("A2" + ":" + "A" + num_projects).getValues();
  var firstNames = progressOverview.getRange("D2" + ":" + "D" + num_projects).getValues();
  var lastNames = progressOverview.getRange("E2" + ":" + "E" + num_projects).getValues();
  var parEmails = progressOverview.getRange("F2" + ":" + "F" + num_projects).getValues();
  var URLs = progressOverview.getRange("H2" + ":" + "H" + num_projects).getValues();
  var fullNames = [];
  for (var j = 0; j < num_projects - 1; j++) {
    fullNames[j] = firstNames[j] + " " + lastNames[j];
  }
  var projInfo = {},
      studentInfo;
  for (var n = 0; n < num_projects - 1; n++) {
    projInfo[projNames[n]] = [fullNames[n], parEmails[n], URLs[n]]
  }
  var projApp,
      partName,
      partEmail,
      projURL,
      projRoster,
      stuInfo,
      availRow = 2,
      projRow;
  for (var k in projInfo) { 
    Logger.log(availRow)
    projApp = projInfo[k][0]
    partName = projInfo[k][0]
    partEmail = projInfo[k][1]
    projURL = projInfo[k][2]
    projRoster = SpreadsheetApp.openByUrl(projURL);
    projApp = projRoster.getSheetByName("Codes").getRange("A9").getValue();
    var nameCell = progressRoster.getRange("A" + availRow)
    nameCell.setFontStyle("italic")
    nameCell.setFontWeight("bold")
    nameCell.setValue(k);
    progressRoster.getRange("B" + availRow).setValue(projApp);
    progressRoster.getRange("C" + availRow).setValue(partName);
    progressRoster.getRange("D" + availRow).setValue(partEmail);
    Logger.log(projApp)
    if (projApp == 0) {
      Logger.log(k)
      progressRoster.getRange("B" + availRow).setValue("No Applicants");
      availRow += 1
      Logger.log("new starting row:" + availRow)
    } else {
      progressRoster.getRange("B" + availRow).setValue(projApp);
      Logger.log(k)
      stuInfo = trackOneProj(projRoster.getSheetByName("StudentInfo")); //should return A-H info: status, matching score, name, preferred name, email, major, graduation year, hour commitment
      Logger.log("num rows:" + projApp)
      projRow = projApp + availRow - 1
      Logger.log("ending row:" + projRow)
      progressRoster.getRange("E" + availRow + ":L" + projRow).setValues(stuInfo);
      availRow = projRow + 1;
      Logger.log("new starting row:" + availRow)
    }
  }
}

//Function: Get information about interview process from one project dashboard
function trackOneProj(sheet){
  var numAPP = sheet.getLastRow();
  Logger.log(numAPP)
  var projInfo = sheet.getRange("A2" + ":H" + numAPP).getValues()//[[]]
  Logger.log(projInfo)
  return projInfo
}

//Function: make sure Progress & Roster are rendered at the same time
function runThis() {
  trackingProgress()  
  trackingRoster()
}
