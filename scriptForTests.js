var sourceSheetName = 'sourseForTest'; // name of sheet with data for test
var uiSheetName = 'uiSheet'; // name of sheet with UI

var textColumnLetter_1 = 'C'; // 'textColumn_1' column in 'sourseForTest'sheet
var hintColumnLetter = 'B'; // 'hint/example' column in 'sourseForTest'sheet
var textColumnLetter_2 = 'D'; // 'textColumn_2' column in 'sourseForTest'sheet
var scoreColumnLetter_1 = 'E'; // score column in 'sourseForTest'sheet for mode of test (flag in B7 cell = 'true')
var scoreColumnLetter_2 = 'F'; // score column in 'sourseForTest'sheet for mode of test (flag in B7 cell = 'false')

var cellFirstModeFlag = 'B8'; // mode selection cell
/* if the flag in B8 cell = 'true':
1) Text from 'textColumn_1' column displayed in cell C1(cellGenerateTestText)
2) Text from 'hint/example' column displayed in cell C2(cellHintText)
3) Text from 'textColumn_2' column displayed in cell C3(cellAnswerText)
if the flag in B8 cell = 'false':
1) Test from 'textColumn_2' column displayed in cell C1(cellAnswerText)
2) Text from 'hint/example' column displayed in cell C2(cellHintText)
3) Text from 'textColumn_1' column displayed in cell C3(cellGenerateTestText)
*/

var testWordColumnLetter;
var answerColumnLetter;
var scoreColumnLetter;

var cellGenerateTestFlag = 'B1';
var cellGenerateTestText = 'C1';
var cellGenerateTestStatus = 'B24';
var generateTestStatus;

var cellHintFlag = 'B2';
var cellHintText = 'C2';

var cellAnswerFlag = 'B3';
var cellAnswerText = 'C3';

var cellKnowFlag = 'B4';
var cellDontKnowFlag = 'B5';
var cellResetFlag = 'B7';

var minCellNumber;
var maxCellNumber;
var cellMinCellNumber = 'B20';
var cellMaxCellNumber = 'B21';
var cellNewMinCellNumber = 'B27';
var cellNewMaxCellNumber = 'B28';
var cellIteration = 'B22';
var cellSelectedLine = 'B23';
var selectedLine;

var cellInserExampleStatus = 'B25';
var insertExampleStatus;
var cellInsertAnswerStatus = 'B26';
var insertAnswerStatus;

var spreadSheetLink;
var sheetLink;
var sheetLinkSource;
var sheetLinkUI;

function editTriger() {
  spreadSheetLink = SpreadsheetApp.getActive();
  sheetLink = spreadSheetLink.getActiveSheet();

  if (sheetLink.getName() == uiSheetName){
    if (sheetLink.getRange(cellResetFlag).getValue() == true) resetTest();
    if (sheetLink.getRange(cellFirstModeFlag).getValue() == true){
      testWordColumnLetter = textColumnLetter_1;
      answerColumnLetter = textColumnLetter_2;
      scoreColumnLetter = scoreColumnLetter_1;
      generateTest();
    }
    if (sheetLink.getRange(cellFirstModeFlag).getValue() == false){
      testWordColumnLetter = textColumnLetter_2;
      answerColumnLetter = textColumnLetter_1;
      scoreColumnLetter = scoreColumnLetter_2;
      generateTest();
    }
  }
}

function generateTest(){
  generateTestStatus = sheetLink.getRange(cellGenerateTestStatus).getValue();
  insertExampleStatus = sheetLink.getRange(cellInserExampleStatus).getValue();
  insertAnswerStatus = sheetLink.getRange(cellInsertAnswerStatus).getValue();

  if(generateTestStatus == true){
    if(insertAnswerStatus == true){
      if((sheetLink.getRange(cellKnowFlag).getValue() == true)){decrementScore();}
      if((sheetLink.getRange(cellDontKnowFlag).getValue() == true)){incrementScore();}
    }
    if((sheetLink.getRange(cellHintFlag).getValue() == true) & ((insertExampleStatus == false))) insertExampleText();
    if((sheetLink.getRange(cellAnswerFlag).getValue() == true) & (insertAnswerStatus == false)) insertAnswerText();
  }
  if((sheetLink.getRange(cellGenerateTestFlag).getValue() == true) & (generateTestStatus == false) & (insertExampleStatus == false) & (insertAnswerStatus == false)) insertTextForTest();
}

function insertTextForTest(){
    sheetLinkSource = spreadSheetLink.getSheetByName(sourceSheetName);
    sheetLinkUI = spreadSheetLink.getSheetByName(uiSheetName);
    minCellNumber = Number(sheetLinkUI.getRange(cellMinCellNumber).getValue());
    maxCellNumber = Number(sheetLinkUI.getRange(cellMaxCellNumber).getValue());

    if(Number(sheetLinkUI.getRange(cellIteration).getValue()) == 0){
      sheetLinkUI.getRange(cellNewMinCellNumber).setValue(minCellNumber);
      sheetLinkUI.getRange(cellNewMaxCellNumber).setValue(maxCellNumber);
    }

    var randomNumber = getRandomNumber(0,1);
    var newMinCellNumber = Number(sheetLinkUI.getRange(cellNewMinCellNumber).getValue());
    var newMaxCellNumber = Number(sheetLinkUI.getRange(cellNewMaxCellNumber).getValue());
    var iteration = Number(sheetLinkUI.getRange(cellIteration).getValue());
    var textForTest = 'test over, click Reset for restart test';

    if((randomNumber == 0) & (iteration <= (maxCellNumber-minCellNumber))){
      textForTest = sheetLinkSource.getRange(testWordColumnLetter + newMinCellNumber).getValue();
      sheetLinkUI.getRange(cellSelectedLine).setValue(newMinCellNumber);

      newMinCellNumber++;
      sheetLinkUI.getRange(cellNewMinCellNumber).setValue(newMinCellNumber);
    }

    if((randomNumber == 1) & (iteration <= (maxCellNumber-minCellNumber))){
      textForTest = sheetLinkSource.getRange(testWordColumnLetter + newMaxCellNumber).getValue();
      sheetLinkUI.getRange(cellSelectedLine).setValue(newMaxCellNumber);
      
      newMaxCellNumber--;
      sheetLinkUI.getRange(cellNewMaxCellNumber).setValue(newMaxCellNumber);
    }


    iteration++;
    sheetLinkUI.getRange(cellIteration).setValue(iteration);
    sheetLinkUI.getRange(cellGenerateTestStatus).setValue(true);
    sheetLinkUI.getRange(cellGenerateTestText).setValue(textForTest);
}

function insertExampleText(){
  sheetLinkSource = spreadSheetLink.getSheetByName(sourceSheetName);
  sheetLinkUI = spreadSheetLink.getSheetByName(uiSheetName);

  selectedLine = sheetLinkUI.getRange(cellSelectedLine).getValue();
  var exampleText = sheetLinkSource.getRange(hintColumnLetter + selectedLine).getValue();

  sheetLinkUI.getRange(cellInserExampleStatus).setValue(true);
  sheetLinkUI.getRange(cellHintText).setValue(exampleText);
}

function insertAnswerText(){
  sheetLinkSource = spreadSheetLink.getSheetByName(sourceSheetName);
  sheetLinkUI = spreadSheetLink.getSheetByName(uiSheetName);

  selectedLine = sheetLinkUI.getRange(cellSelectedLine).getValue();
  var answerText = sheetLinkSource.getRange(answerColumnLetter + selectedLine).getValue();

  sheetLinkUI.getRange(cellInsertAnswerStatus).setValue(true);
  sheetLinkUI.getRange(cellAnswerText).setValue(answerText);
}
    
function getRandomNumber(min, max) {
  return Math.floor(Math.random() * (max - min + 1) + min);
}

function decrementScore(){
  sheetLinkSource = spreadSheetLink.getSheetByName(sourceSheetName);
  sheetLinkUI = spreadSheetLink.getSheetByName(uiSheetName);
  selectedLine = sheetLinkUI.getRange(cellSelectedLine).getValue();

  var score = Number(sheetLinkSource.getRange(scoreColumnLetter + selectedLine).getValue());
  if (score > 0){
    score--;
    sheetLinkSource.getRange(scoreColumnLetter + selectedLine).setValue(score);
  }

  resetUIFields();
  resetHolderFieldsTestUncomplite();
}

function incrementScore(){
  sheetLinkSource = spreadSheetLink.getSheetByName(sourceSheetName);
  sheetLinkUI = spreadSheetLink.getSheetByName(uiSheetName);
  selectedLine = sheetLinkUI.getRange(cellSelectedLine).getValue();

  var score = Number(sheetLinkSource.getRange(scoreColumnLetter + selectedLine).getValue());
  if (score >= 0){
    score++;
    sheetLinkSource.getRange(scoreColumnLetter + selectedLine).setValue(score);
  }

  resetUIFields();
  resetHolderFieldsTestUncomplite();
}

function resetTest(){
  sheetLinkUI = spreadSheetLink.getSheetByName(uiSheetName);
  resetUIFields();
  resetHolderFieldsTestComplite();
  resetHolderFieldsTestUncomplite();
  sheetLinkUI.getRange(cellResetFlag).setValue(false);
}

function resetHolderFieldsTestComplite(){
  sheetLinkUI.getRange(cellIteration).setValue('0');
  sheetLinkUI.getRange(cellNewMinCellNumber).setValue('0');
  sheetLinkUI.getRange(cellNewMaxCellNumber).setValue('0');
}

function resetHolderFieldsTestUncomplite(){
  sheetLinkUI.getRange(cellGenerateTestStatus).setValue(false);
  sheetLinkUI.getRange(cellInserExampleStatus).setValue(false);
  sheetLinkUI.getRange(cellInsertAnswerStatus).setValue(false);
  sheetLinkUI.getRange(cellSelectedLine).setValue('');
}

function resetUIFields(){
  sheetLinkUI.getRange(cellGenerateTestFlag).setValue(false);
  sheetLinkUI.getRange(cellHintFlag).setValue(false);
  sheetLinkUI.getRange(cellAnswerFlag).setValue(false);
  sheetLinkUI.getRange(cellKnowFlag).setValue(false);
  sheetLinkUI.getRange(cellDontKnowFlag).setValue(false);
  sheetLinkUI.getRange(cellGenerateTestText).setValue('');
  sheetLinkUI.getRange(cellHintText).setValue('');
  sheetLinkUI.getRange(cellAnswerText).setValue('');
}