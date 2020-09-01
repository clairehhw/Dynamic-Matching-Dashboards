//Function: maps student status (available for interview or not) onto sheet 
function mapStatus() {
  var allStatus = SpreadsheetApp.openById('1QzY9bsa-MYZuxFB3VgGFGJvX3aF1KcJ7EXGFc7Eyp3M').getSheetByName("TrackingALL")
  var compactStatus = {}
  var numApp = 1003;
  var stuEmails = allStatus.getRange("I2:I" + numApp).getValues();
  var stuStatus = allStatus.getRange("E2:E" + numApp).getValues();
  var oneEmail,
      oneStatus,
      temp,
      finalDict,
      finalEmail = [],
      finalStatus = [],
      finalRow;
  for (var i = 0; i < stuEmails.length; i++){
    oneEmail = stuEmails[i]
    oneStatus = stuStatus[i][0]
    if (compactStatus[oneEmail]) {
      temp = compactStatus[oneEmail]
      temp.push(oneStatus)
      compactStatus[oneEmail] = temp
    } else {
      compactStatus[oneEmail] = [oneStatus]
    }
  }
  finalDict = determineStatus(compactStatus)
  for (var key in finalDict) {
    finalEmail.push([key])
    finalStatus.push([finalDict[key]])
  }
  var wlStatus = SpreadsheetApp.openById('1AWcAuKMJWpIS80dIA9xhWyTDsYVgJzSvI72W_vdmdlw').getSheetByName("Status")
  finalRow = 2 + finalEmail.length - 1
  wlStatus.getRange("A2:A" + finalRow).setValues(finalEmail)
  wlStatus.getRange("B2:B" + finalRow).setValues(finalStatus)
}

//Function: consolidates student status (available for interview or not) into a dictionary (key-student email; value-status)
function determineStatus(dict) {
  var oneStates,
      newDict = {};
  for (var k in dict){
    oneStates = dict[k]
    if (oneStates.includes("Accepted")){
      newDict[k] = "This student has been accepted to another project."
    } else if (oneStates.includes("Interviewing ")){
      newDict[k] = "This student is currently interviewing for another project."
    } else {
      newDict[k] = "This student is available for interview"
    }
  }
  return newDict
}