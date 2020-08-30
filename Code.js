function makeProjectSheets() {  
  var projectsBook = SpreadsheetApp.openById('1QzY9bsa-MYZuxFB3VgGFGJvX3aF1KcJ7EXGFc7Eyp3M').getSheetByName("Project Overview") //FA20 Project Tracking
  var sheetsFolder = DriveApp.getFolderById('1YRuS28hDuMLhRxhaUR9CKA8Af0rmr-tk') //FA20 Project Matching Sheets
  var templateBook = DriveApp.getFileById('10HO3rjhC_TqUCTUE5UZrtCJuOvNxlP218ZN9732PaGg') //Template: Partner Dashboard
  var projNameCol = 1; 
  var projNames;
  var numProjects = 43 // UPDATE IF MORE PROJECTS ADDED
  var urlCol = 8;
  
  projNames = projectsBook.getRange(2, projNameCol, numProjects).getValues(); //Get project names 
  
  //Creating dashboard templates for each project
  var tempBook;
  var tempName;
  for(var i = 0; i<projNames.length; i++){
    tempName = 'Project Applicants - ' + projNames[i][0];
    if (fileExists_(tempName, sheetsFolder) == false){
      tempBook = templateBook.makeCopy(tempName, sheetsFolder);
      projectsBook.getRange(i+2, 8).setValue(tempBook.getUrl());
    }
  }
}

function fileExists_(filename, folderId) {
  var folder = DriveApp.getFoldersByName(folderId);

  if (!folder.hasNext()) {
    return false;
  }

  return folder.next().getFilesByName(filename).hasNext();
}

function updateAll(){
  var projectSheets = getNameID();
  var responseFlagCol = 1; //Initial values set as 0, change to 1 if mapped to project roster
  var first = 33; 
  var second = 35;
  var third = 37;
  var responseSheet = SpreadsheetApp.openById('14Puy8GQp1VZHawTQycB-dXocUuDNs6PKVNTUo3kD1wk').getSheetByName('FormResponses');
  var projCols = ['B', 'AH', 'A','AG'];
  var num_applicants = SpreadsheetApp.openById('14Puy8GQp1VZHawTQycB-dXocUuDNs6PKVNTUo3kD1wk').getSheetByName('Number').getRange(1,1).getValue();
 
  var flagCheck
  var appMaterial;
  var firstChoice;
  var secondChoice;
  var thirdChoice;
  var scholar
  var oneApp
  
  for(var i = 2; i<num_applicants+2; i++){
    flagCheck = responseSheet.getRange(i, responseFlagCol).getValue();
    if (flagCheck != 2){
      oneApp = responseSheet.getRange("B" + i + ":AN" + i).getValues()[0];
      Logger.log("Updating" + i)
      firstChoice = SpreadsheetApp.openById(projectSheets[oneApp[first]]).getSheetByName("StudentInfo_Raw") 
      var firstLast = firstChoice.getLastRow() + 1;
      firstChoice.getRange("AH" + firstLast).setValue(oneApp[first + 1])
      firstChoice.getRange("A" + firstLast + ":AG" + firstLast).setValues([oneApp.slice(0, 33)])
      firstChoice.getRange("AI" + firstLast).setValue(3);
      Logger.log("first:" + oneApp[first] + "updated")
      secondChoice = SpreadsheetApp.openById(projectSheets[oneApp[second]]).getSheetByName("StudentInfo_Raw")
      var secondLast = secondChoice.getLastRow() + 1;
      secondChoice.getRange("AH" + secondLast).setValue(oneApp[second + 1])
      secondChoice.getRange("A" + secondLast + ":AG" + secondLast).setValues([oneApp.slice(0, 33)])
      secondChoice.getRange("AI" + secondLast).setValue(2); 
      Logger.log("second:" + oneApp[second] + "updated")
      thirdChoice = SpreadsheetApp.openById(projectSheets[oneApp[third]]).getSheetByName("StudentInfo_Raw")
      var thirdLast = thirdChoice.getLastRow() + 1;
      thirdChoice.getRange("AH" + thirdLast).setValue(oneApp[third + 1])
      thirdChoice.getRange("A" + thirdLast + ":AG" + thirdLast).setValues([oneApp.slice(0, 33)])
      thirdChoice.getRange("AI" + thirdLast).setValue(1);
      Logger.log("third:" + oneApp[third] + "updated")
      if (oneApp[29] == "Yes") {
        firstChoice.getRange("AJ" + firstLast).setValue(5);
        secondChoice.getRange("AJ" + secondLast).setValue(5);
        thirdChoice.getRange("AJ" + thirdLast).setValue(5);
      } else {
        firstChoice.getRange("AJ" + firstLast).setValue(0);
        secondChoice.getRange("AJ" + secondLast).setValue(0);
        thirdChoice.getRange("AJ" + thirdLast).setValue(0);
      }
      responseSheet.getRange(i, responseFlagCol).setValue(2);
    } else {
      continue;
    };
  }
  for (var p in projectSheets) {
    generateScores(projectSheets[p])
    Logger.log("Project" + p + "generated")
  }
  
}

function getNameID(){
  var sheetsFolder = DriveApp.getFolderById('1YRuS28hDuMLhRxhaUR9CKA8Af0rmr-tk') //FA20 Project Matching Sheets
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

function generateScores(sheetId) {  
  var num_app = SpreadsheetApp.openById(sheetId).getSheetByName("Codes").getRange("A9").getValue();
  
  var partnerInput = SpreadsheetApp.openById(sheetId).getSheetByName("Codes").getRange("A11:L11").getValues()[0];
  if (num_app > 0) {
    var studentInfo = SpreadsheetApp.openById(sheetId).getSheetByName("StudentInfo_Raw");
    var studentFinal = SpreadsheetApp.openById(sheetId).getSheetByName("StudentInfo");
    var col_Name = 'C';
    var col_Preferred = 'D';
    var col_Email = 'B';
    var col_Major = 'H';
    var col_Year = 'J';
    var col_Hour = 'K';
    var col_Re = 'AD';
    var col_ST_1 = 'AE';
    var col_ST_2 = 'AF';
    var col_ST_3 = 'AG';
    var col_ST_4 = 'AH';
    var cols = [col_Name,col_Preferred,col_Email,col_Major,col_Year,col_Hour,
              col_Re, col_ST_1,col_ST_2,col_ST_3,col_ST_4]
    var studentScore;
    var skillRange;
    var studentRange;
    var scholar;
    var choice;
    for (var n = 2; n < num_app + 2; n++) {
      scholar = studentInfo.getRange("AJ" + n).getValue();
      choice = studentInfo.getRange("AI" + n).getValue();
      skillRange = 'L' + n + ':' + 'W' + n;
      studentScore = calcMatching(partnerInput, studentInfo.getRange(skillRange).getValues()[0]) + choice + scholar;
      studentRange = ['C','D','E','F','G','H','I','J','K','L','M']
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

function calcMatching(exp, input){
  var num_skills = 12; //12 skills required on the application form
  var matching = 0;
  for (var i = 0; i < num_skills; i ++){
    if (exp[i] == input[i]) {
      matching += 1;
    } else if (exp[i] < input[i]) {
      matching += .75;
    } else {
      matching += 0;
    }
  }
  return matching;
}