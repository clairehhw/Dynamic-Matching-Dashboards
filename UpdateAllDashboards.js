//Function: Make dashboards from template for all projects
function makeProjectSheets() {  
  var projectsBook = SpreadsheetApp.openById('1zCmXNJ0ytEcX_YR7f2QIms0gpsjItwKfgYHlBjfHojY').getSheetByName("Project Overview") //Project Tracking Sheet
  var sheetsFolder = DriveApp.getFolderById('14QSmKHSMVgIaq7H7-GrnOVP8CjHVCeat') //Folder to hold Project Dashboards Sheets
  var templateBook = DriveApp.getFileById('12nmhEyp87787QxMmAk__x9D-4SMS88ZiZd5hImbTWkw') //Template: Partner Dashboard
  var projNameCol = 1,
      projNames,
      numProjects = 51, // UPDATE IF MORE PROJECTS ADDED
      urlCol = 8;
  
  projNames = projectsBook.getRange(2, projNameCol, numProjects).getValues(); //Get project names 
  
  //Creating dashboard templates for each project
  var tempBook,
      tempName;
  for(var i = 0; i<projNames.length; i++){
    tempName = 'Project Applicants - ' + projNames[i][0];
    if (fileExists_(tempName, sheetsFolder) == false){
      tempBook = templateBook.makeCopy(tempName, sheetsFolder);
      projectsBook.getRange(i+2, urlCol).setValue(tempBook.getUrl());
    }
  }
}

//Function: Check if a dashboard has already been made for certain project
function fileExists_(filename, folderId) {
  var folder = DriveApp.getFoldersByName(folderId);
  if (!folder.hasNext()) {
    return false;
  }
  return folder.next().getFilesByName(filename).hasNext();
}

//Function: Make sure project names match with choice names 
function filterName(n){
  var projList = SpreadsheetApp.openById('1zCmXNJ0ytEcX_YR7f2QIms0gpsjItwKfgYHlBjfHojY').getSheetByName("Project Overview").getRange("A2:A52").getValues()
  for (var i = 0; i < projList.length; i++) {
    if (n.indexOf(projList[i]) != -1) {
      return projList[i]
    }
  }
}

//Function: update ALL project dashboards once a submission is made
function updateAll(){
  var projectSheets = getNameID();
  var responseFlagCol = 1; //Initial values set as NULL, change to 2 if mapped to project roster
  var responseSheet = SpreadsheetApp.openById('1zNxeaRAv8HAZG8UmG91oyJTzrc_Z-LHW8eBDoLWWQsQ').getSheetByName('FormResponses');
  var num_applicants = responseSheet.getLastRow();
 
  var first = 33,
      second = 35,
      third = 37,
      flagCheck,
      appMaterial,
      firstChoice,
      secondChoice,
      thirdChoice,
      scholar,
      oneApp,
      oneEmail;
  
  //Search the form responses sheet, update chosen project dashboards 
  if (num_applicants > 1) {
    for(var i = 2; i<num_applicants+1; i++){
      flagCheck = responseSheet.getRange(i, responseFlagCol).getValue();
      if (flagCheck != 2){
        oneApp = responseSheet.getRange("B" + i + ":AN" + i).getValues()[0];
        oneEmail = oneApp[1];
        Logger.log("Updating" + oneApp)
        var firstName = filterName(oneApp[first])
        firstChoice = SpreadsheetApp.openById(projectSheets[firstName]).getSheetByName("StudentInfo_Raw")
        if (checkExist(oneEmail, firstChoice) != true) {
          var firstLast = firstChoice.getLastRow() + 1;
          firstChoice.getRange("AH" + firstLast).setValue(oneApp[first + 1])
          firstChoice.getRange("A" + firstLast + ":AG" + firstLast).setValues([oneApp.slice(0, 33)])
          firstChoice.getRange("AI" + firstLast).setValue(3);
          if (oneApp[28] == "Yes") {
            firstChoice.getRange("AJ" + firstLast).setValue(5);
          } else {
            firstChoice.getRange("AJ" + firstLast).setValue(0);
          }
          generateScores(projectSheets[firstName])
          Logger.log("first:" + oneApp[first] + "updated")
        }
        if (oneApp[second]) {
          var secondName = filterName(oneApp[second])
          secondChoice = SpreadsheetApp.openById(projectSheets[secondName]).getSheetByName("StudentInfo_Raw")
          if (checkExist(oneEmail, secondChoice) != true) {
            var secondLast = secondChoice.getLastRow() + 1;
            secondChoice.getRange("AH" + secondLast).setValue(oneApp[second + 1])
            secondChoice.getRange("A" + secondLast + ":AG" + secondLast).setValues([oneApp.slice(0, 33)])
            secondChoice.getRange("AI" + secondLast).setValue(2); 
            if (oneApp[28] == "Yes") {
              secondChoice.getRange("AJ" + firstLast).setValue(5);
            } else {
              secondChoice.getRange("AJ" + firstLast).setValue(0);
            }
            generateScores(projectSheets[secondName])
            Logger.log("second:" + oneApp[second] + "updated")
          }
        }
        if (oneApp[third]) {
          var thirdName = filterName(oneApp[third])
          thirdChoice = SpreadsheetApp.openById(projectSheets[thirdName]).getSheetByName("StudentInfo_Raw")
          if (checkExist(oneEmail, thirdChoice) != true){
            var thirdLast = thirdChoice.getLastRow() + 1;
            thirdChoice.getRange("AH" + thirdLast).setValue(oneApp[third + 1])
            thirdChoice.getRange("A" + thirdLast + ":AG" + thirdLast).setValues([oneApp.slice(0, 33)])
            thirdChoice.getRange("AI" + thirdLast).setValue(1);
            if (oneApp[28] == "Yes") {
              thirdChoice.getRange("AJ" + firstLast).setValue(5);
            } else {
              thirdChoice.getRange("AJ" + firstLast).setValue(0);
            }
            Logger.log("third:" + thirdName + "updated")
            generateScores(projectSheets[thirdName])
          }
        }
        responseSheet.getRange(i, responseFlagCol).setValue(2);
      } else {
      continue;
      };
    }
  } else {
    return;
  } 
}

//Function: returns dictionary (key-project name, val-project dashboard id)
function getNameID(){
  var sheetsFolder = DriveApp.getFolderById('14QSmKHSMVgIaq7H7-GrnOVP8CjHVCeat') //Folder that holds Project Dashboards
  var projectSheets = {};
  var files = sheetsFolder.getFiles();
  while (files.hasNext()){
    file = files.next();
    var id = file.getId();
    var projectName = file.getName();
    var name = projectName.replace('Project Applicants - ','');
    projectSheets[name] = id;
  }
  return projectSheets;
}

//Function: generate filtered student information sheet with matching scores for one project dashboard
function generateScores(sheetId) {  
  var num_app = SpreadsheetApp.openById(sheetId).getSheetByName("Codes").getRange("A9").getValue();
  var partnerInput = SpreadsheetApp.openById(sheetId).getSheetByName("Codes").getRange("A11:L11").getValues()[0];
  if (num_app > 0) {
    var studentInfo = SpreadsheetApp.openById(sheetId).getSheetByName("StudentInfo_Raw");
    var studentFinal = SpreadsheetApp.openById(sheetId).getSheetByName("StudentInfo");
    var col_Name = 'C',
        col_Preferred = 'D',
        col_Email = 'B',
        col_Major = 'H',
        col_Year = 'J',
        col_Hour = 'K',
        col_Re = 'AD',
        col_ST_1 = 'AE',
        col_ST_2 = 'AF',
        col_ST_3 = 'AG',
        col_ST_4 = 'AH';
    var cols = [col_Name,col_Preferred,col_Email,col_Major,col_Year,col_Hour,
              col_Re, col_ST_1,col_ST_2,col_ST_3,col_ST_4]
    var studentRange = ['C','D','E','F','G','H','I','J','K','L','M']
    var studentScore,
        skillRange,
        scholar,
        choice;
    for (var n = 2; n < num_app + 2; n++) {
      scholar = studentInfo.getRange("AJ" + n).getValue();
      choice = studentInfo.getRange("AI" + n).getValue();
      skillRange = 'N' + n + ':' + 'Y' + n;
      studentScore = calcMatching(partnerInput, studentInfo.getRange(skillRange).getValues()[0]) + choice + scholar;
      for (var i = 0; i < 11;i++) {
        var val = studentInfo.getRange(cols[i] + n).getValue();
        studentFinal.getRange(studentRange[i]+n).setValue(val);
       }
      studentFinal.getRange("B" + n).setValue(studentScore);
     }
  } else {
    return;
  }
}

//Function: calculate matching scores
function calcMatching(exp, stu){
  var numSkills = 12; //12 skills required on the application form
  var matching = 0;
  var stuCode = {"No experience":20, 
                 "Beginner (I can do a few operations":40,
                 "Familiar (I have developed at least one project)":60,
                 "Intermediate (Multiple semesters of experience)": 80,
                 "Advanced (I would feel comfortable teaching the subject)":100}
  var newStu = new Array(numSkills);
  for (var n = 0; n < numSkills; n++) {
    newStu[n] = stuCode[stu[n]]
  }
  for (var i = 0; i < numSkills; i ++){
    if (exp[i] == newStu[i]) {
      matching += 1;
    } else if (exp[i] < newStu[i]) {
      matching += .75;
    } else {
      matching += 0;
    }
  }
  return matching;
}

//Function: check if the student is already mapped to this dashboard
function checkExist(email, roster){
  var exist = false
  var num = roster.getLastRow();
  var emails = roster.getRange("B2:B" + num).getValues();
  for (var i = 0; i < emails.length; i++) {
    if (emails[i][0] == email){
      exist = true
    }
  }
  return exist
}
