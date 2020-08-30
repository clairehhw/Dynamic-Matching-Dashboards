//Function: Get dictionary(key-question name;value-question ID)
function getItemTitleId(formId) {
  var form = FormApp.openById(formId);
  var allItems = form.getItems();
  var thisItem,
      titleId = {},
      thisItemTitle,
      thisItemId;

  for (var i=0;i<allItems.length;i+=1) {
    thisItem = allItems[i];
    thisItemTitle = thisItem.getTitle();
    thisItemId = thisItem.getId();
    titleId[thisItemTitle] = thisItemId
  };
  return titleId;
};

//Function: Get dictionary(key-project name;value-number of applicants)
function getItemResponse(question) {
  var titleId = getItemTitleId('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var form = FormApp.openById('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var projApp = {};
  var choices = form.getItemById(titleId[question]);
  var formResponses = form.getResponses();
  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var itemResponse = formResponse.getResponseForItem(choices);
    if (itemResponse != null) {
      var project = itemResponse.getResponse();
      var projName = filterName(project)
      if (projApp[projName]) {
        projApp[projName] += 1;
      } else {
        projApp[projName] = 1;
      } 
    }
  };
  Logger.log(projApp)
  return projApp;
};

//Function: Update # of applicants for 1st, 2nd, 3rd choices
function runThis(){
  editForm("What is your FIRST choice?", 1)
  editForm("What is your SECOND choice?", 2)
  editForm("What is your THIRD choice?", 3)
}

//Function: Make sure project names match with choice names
function filterName(n){
  var projList = SpreadsheetApp.openById('1QzY9bsa-MYZuxFB3VgGFGJvX3aF1KcJ7EXGFc7Eyp3M').getSheetByName("Project Overview").getRange("A2:A52").getValues()
  for (var i = 0; i < projList.length; i++) {
    if (n.indexOf(projList[i]) != -1) {
      return projList[i]
    }
  }
}

//Function: Update # of applicants for one choice
function editForm(question, pre){
  var titleId = getItemTitleId('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var form = FormApp.openById('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var thisItem = form.getItemById(titleId[question]);
  var myListItem = thisItem.asListItem();
  var choices = myListItem.getChoices();
  var newChoices = new Array(choices.length);
  var dictQuestion = getItemResponse(question);
  Logger.log(dictQuestion)
  var msg;
  for (var v = 0; v < choices.length; v++) {
    var proj = choices[v].getValue()
    var projName = filterName(proj)
    if (dictQuestion[projName]){
      var numApp = dictQuestion[projName]
      if (pre == 1) {
        msg = projName + " - " + numApp + " Applicants list it as 1st Choice"
        newChoices[v] = msg
      } else if (pre == 2) {
        msg = projName + " - " + numApp + " Applicants list it as 2nd Choice"
        newChoices[v] = msg
      } else {
        msg = projName + " - " + numApp + " Applicants list it as 3rd Choice"
        newChoices[v] = msg
      }
    } else {
      newChoices[v] = projName + " - No Applicants"
    }
  }
  Logger.log(newChoices)
  myListItem.setChoiceValues(newChoices);
};
