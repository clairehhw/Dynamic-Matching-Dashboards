function editItemType() {
  var form = FormApp.openById('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var allItems = form.getItems();
  var i=0,
      thisItem,
      thisItemType,
      myCheckBoxItem;

  for (i=0;i<allItems.length;i+=1) {
    thisItem = allItems[i];
    thisItemType = thisItem.getType();
    Logger.log('thisItemType: ' + thisItemType);
  };
};

function getItemTitleId(formId) {
  var form = FormApp.openById(formId);
  var allItems = form.getItems();
  var i=0,
      thisItem,
      titleId = {},
      thisItemTitle,
      thisItemId;

  for (i=0;i<allItems.length;i+=1) {
    thisItem = allItems[i];
    thisItemTitle = thisItem.getTitle();
    thisItemId = thisItem.getId();
    titleId[thisItemTitle] = thisItemId
  };
  return titleId;
};

function getItemResponse(question) {
  var titleId = getItemTitleId('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var form = FormApp.openById('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var projApp = {};
  var choices = form.getItemById(titleId[question]);
  var formResponses = form.getResponses();
  Logger.log(formResponses.length)
  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var itemResponse = formResponse.getResponseForItem(choices);
    if (itemResponse != null) {
      var project = itemResponse.getResponse();
      if (projApp[project]) {
        projApp[project] += 1;
      } else {
        projApp[project] = 1;
      }
    }
  };
  Logger.log(projApp);
  return projApp;
};

function test(){
  editForm("What is your FIRST choice?", 1)
  editForm("What is your SECOND choice?", 2)
  editForm("What is your THIRD choice?", 3)
}

function editForm(question, pre){
  var projList = SpreadsheetApp.openById('1QzY9bsa-MYZuxFB3VgGFGJvX3aF1KcJ7EXGFc7Eyp3M').getSheetByName("Project Overview").getRange("A2:A46").getValues()
  var titleId = getItemTitleId('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var form = FormApp.openById('1utSamTRoZb0uBo26yuSkTgcLH-YP1i42cKHH4zkqO3o');
  var thisItem = form.getItemById(titleId[question]);
  var myListItem = thisItem.asListItem();
  var choices = myListItem.getChoices();
  var newChoices = new Array(choices.length);
  var dictQuestion = getItemResponse(question);
  if (pre == 1) {
    dictQuestion["BEACO2N: Calibration and analyses of sensor networks"] = dictQuestion["BEACO2N: Calibration and analyses of sensor networks"] - 1
    dictQuestion["Wordnik Etymology Search"] = dictQuestion["Wordnik Etymology Search"] - 1
    dictQuestion["Clinical text classification/information extraction to understand real-world treatment effects at a large, academic medical center"] = dictQuestion["Clinical text classification/information extraction to understand real-world treatment effects at a large, academic medical center"] - 1
    dictQuestion["Group Dynamics on Reddit"] = dictQuestion["Group Dynamics on Reddit"] - 2
    dictQuestion["Clinical Natural Language Understanding using transformer models and extensions incorporating tabular data"] = dictQuestion["Clinical Natural Language Understanding using transformer models and extensions incorporating tabular data"] - 2
  } else if (pre == 2) {
    dictQuestion["Exploring spaces and places of violence against young people experiencing homelessness"] = dictQuestion["Exploring spaces and places of violence against young people experiencing homelessness"] - 1 
    dictQuestion["What's Important to the Supreme Court"] = dictQuestion["What's Important to the Supreme Court"] - 1
    dictQuestion["Empirical Examination of Corporate Rebranding and Trademarks"] = dictQuestion["Empirical Examination of Corporate Rebranding and Trademarks"] - 1
    dictQuestion["Wordnik Etymology Search"] = dictQuestion["Wordnik Etymology Search"] - 1
    dictQuestion["Identification and Classification of Intrinsically Disordered Regions in Proteins"] = dictQuestion["Identification and Classification of Intrinsically Disordered Regions in Proteins"] - 1
    dictQuestion["Wordnik Hyphenation Project"] = dictQuestion["Wordnik Hyphenation Project"] - 1
    dictQuestion["System telemetry analysis"] = dictQuestion["System telemetry analysis"] - 1
    dictQuestion["Ancient World Computational Analysis"] = dictQuestion["Ancient World Computational Analysis"] - 1
  } else {
    dictQuestion["Clinical Natural Language Understanding using transformer models and extensions incorporating tabular data"] = dictQuestion["Clinical Natural Language Understanding using transformer models and extensions incorporating tabular data"] - 1
    dictQuestion["NLP for Cannabis Text Data"] = dictQuestion["NLP for Cannabis Text Data"] - 1 
    dictQuestion["Wordnik Hyphenation Project"] = dictQuestion["Wordnik Hyphenation Project"] - 2
    dictQuestion["Wordnik Etymology Search"] = dictQuestion["Wordnik Etymology Search"] - 1
    dictQuestion["Group Dynamics on Reddit"] = dictQuestion["Group Dynamics on Reddit"] - 2
    dictQuestion["Business Case for Investment into MMR vaccination by Developing NAtions"] = dictQuestion["Business Case for Investment into MMR vaccination by Developing NAtions"] - 1
  }
  for (var v = 0; v < choices.length; v++) {
    var proj = choices[v].getValue()
    var projName = projList[v]
    if (dictQuestion[proj]){
      var numApp = dictQuestion[proj]
      if (pre == 1) {
        newChoices[v] = projName + " - " + numApp + " Applicants list it as 1st Choice"
      } else if (pre == 2) {
        newChoices[v] = projName + " - " + numApp + " Applicants list it as 2nd Choice"
      } else {
        newChoices[v] = projName + " - " + numApp + " Applicants list it as 3rd Choice"
      }
    } else {
      newChoices[v] = projName + " - No Applicants"
    }
  }
  Logger.log(newChoices)
  myListItem.setChoiceValues(newChoices);
  //Logger.log('Published URL: ' + form.getPublishedUrl());
  //Logger.log('Editor URL: ' + form.getEditUrl());
};