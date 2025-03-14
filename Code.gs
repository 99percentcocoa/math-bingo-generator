const ss = SpreadsheetApp.getActiveSpreadsheet();
const settingsSheet = ss.getSheetByName('Settings');
const questionsSheet = ss.getSheetByName('Questions');
const cardsSheet = ss.getSheetByName('Cards');
const formattingSheet = ss.getSheetByName('Formatting');

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Bingo!')
      .addItem('Generate Bingo Cards!', 'generateBingoCards')
      .addSeparator()
      .addItem('Clear All', 'clearAll') //                    !!PENDING
      .addItem('Clear Questions', 'clearQuestions') //        !!PENDING
      .addItem('Clear Bingo Cards', 'clearBingoCards') //     !!PENDING
      .addToUi();
}

// returns the user settings in the format [numberOfCards, saveLocation]
function getUserSettings() {
  return settingsSheet.getRange(3, 3, 2).getValues().flat();
}

function generateBingoCards() {

  clearBingoCards();

  let userSettings = getUserSettings();

  let answers = questionsSheet.getRange(2, 3, questionsSheet.getLastRow()-1, 1).getValues().flat();
  // let slideQuestions = questionsSheet.getRange(2, 2, questionsSheet.getLastRow()-1, 1).getValues().flat();

  //let gridArr = [];
  //while (questions.length) gridArr.push(questions.splice(0,4));
  Logger.log(answers);
  // Logger.log(shuffle(questions));

  // Logger.log(shuffle(gridArr))

  let numCards = userSettings[0];

  // check if number of questions is less than number of cards
  if (answers.length < 16) {
    throw new Error("Number of questions cannot be less than 16. Please add more questions.")
  }

  for (let i = 0; i < numCards; i++) {
    // cardsSheet.getRange(i*5+1, 1, 4, 4).setValues(returnGrid(shuffle(questions), 4));

    let colNum = (i % 3);
    let lastRow = cardsSheet.getLastRow();
    let rowNum = getRow(i);
    formattingSheet.getRange(1, 1).copyTo(cardsSheet.getRange(rowNum, colNum*5+1));
    cardsSheet.getRange(rowNum, colNum*5+1, 1, 4).merge();

    cardsSheet.getRange(rowNum+1, colNum*5+1, 4, 4).setValues(returnGrid(shuffle(getRandomValues(answers, 16)), 4));
    formattingSheet.getRange(2, 1, 4, 4).copyFormatToRange(cardsSheet, colNum*5+1, colNum*5+5, rowNum+1, rowNum+5);

    cardsSheet.getRange(rowNum, colNum*5+1, 5, 4).setBorder(true, true, true, true, true, true);
  }

  // generatePresentation(shuffle(answers));
}

function generatePresentation (answers) {

  let userSettings = getUserSettings();
  let folderID = userSettings[1];

  let presi = SlidesApp.create(`Bingo ${Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy")}`);
  let file = DriveApp.getFileById(presi.getId());

  file.moveTo(DriveApp.getFolderById(folderID));

  answers.forEach((q) => {
    let newSlide = presi.appendSlide();
    let textBox = newSlide.insertTextBox(q);
    textBox.scaleWidth(3);
    textBox.alignOnPage(SlidesApp.AlignmentPosition.CENTER);
    textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

    textBox.getText().getTextStyle().setBold(true).setFontSize(60);
  })
}

function clearBingoCards() {
  cardsSheet.getRange(1, 1, cardsSheet.getLastRow()+1, cardsSheet.getLastColumn()+1).clear();
}

function clearQuestions() {
  questionsSheet.getRange(2, 1, questionsSheet.getLastRow()-1, 3).clearContent();
}

function shuffle(arr) {
  return arr.sort( () => .5 - Math.random() );
}

function returnGrid(arr, n) {
  let arrCopy = [...arr]
  let out = []
  while (arrCopy.length) out.push(arrCopy.splice(0,n));
  return out;
}

function getRow(input) {
    // Find the group index (0 for the first group, 1 for the second, etc.)
    const groupIndex = Math.floor(input / 3);

    // Calculate the output based on the group index
    const output = 1 + groupIndex * 6;

    return output;
}

function getRandomValues(arr, numValues) {
    // Logger.log(`Original array: ${arr}`);
    return arr.sort(() => 0.5 - Math.random()).slice(0, numValues);
}

function getCurrentDateTime() {
  var now = new Date();
  var day = now.getDate();
  var month = now.getMonth() + 1; // Months are zero-indexed
  var year = now.getFullYear();
  var hours = now.getHours().toString().padStart(2, '0');
  var minutes = now.getMinutes().toString().padStart(2, '0');

  // Format the date and time
  var formattedDateTime = day + "/" + month + "/" + year.toString().substr(-2) + " " + hours + ":" + minutes;

  return formattedDateTime;
}
