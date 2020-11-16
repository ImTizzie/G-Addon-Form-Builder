/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

// global variables: 
var questions = [];
var answers = [];

var questionStyle = {};
questionStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
questionStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
questionStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
questionStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1;
questionStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
  
var infoStyle = {};
infoStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
infoStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
infoStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
infoStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
infoStyle[DocumentApp.Attribute.SPACING_AFTER] = 1;

var creditStyle = {};
creditStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
creditStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
creditStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#666666';
creditStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
creditStyle[DocumentApp.Attribute.SPACING_AFTER] = 1;


/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Text Field', 'showTextFieldSidebar')
  .addItem('Multiple Choice', 'showMultipleChoiceSidebar')
  .addItem('Grading', 'showGradingSidebar')
  .addItem('Header', 'showHeaderSidebar')
  .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**********************************************************
 * These showSidebar-type functions are utilized to open up
 * the sidebar containing the element's user interface.
 **********************************************************/
function showTextFieldSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('textfield')
      .setTitle('Text Field');
  DocumentApp.getUi().showSidebar(ui);
}

function showMultipleChoiceSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('multiplechoice')
      .setTitle('Multiple Choice');
  DocumentApp.getUi().showSidebar(ui);
}

function showGradingSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('grading')
      .setTitle('Grading');
  DocumentApp.getUi().showSidebar(ui);
}

function showHeaderSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('header')
      .setTitle('Header');
  DocumentApp.getUi().showSidebar(ui);
}

/*************************************************
 * This function creates the Text Field question.
 **************************************************/
function addTextField() {
  
  /*************************************************
   * These first three variables create the necessary
   * styles needed to implement the text-field question.
  **************************************************/
  
  var questionStyle = {};
  questionStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  questionStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  questionStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  questionStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1;
  questionStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
  
  var infoStyle = {};
  infoStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  infoStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
  infoStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  infoStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
  infoStyle[DocumentApp.Attribute.SPACING_AFTER] = 1;
  
  var creditStyle = {};
  creditStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
  creditStyle[DocumentApp.Attribute.FONT_SIZE] = 8;
  creditStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#666666';
  creditStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
  creditStyle[DocumentApp.Attribute.SPACING_AFTER] = 1;

  /*************************************************
   * Obtains the document and gets the body section
   * of the given document.
  **************************************************/
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  /*************************************************
   * Stores all the given variables necessary to
   * customize the text-field question.
  **************************************************/
  
  var question = "Question 1: This is a question?";
  var ans = "Answer";
  var points = 10;
  var information = "This is extra detail for the question given. It’s possible that there might be additional instruction to be added for a good response. This is just a bunch of filler text that I have to fill up space.";
  var lines = 5;
  var partialCredit = false;
  
  questions.push(question);
  answers.push(ans);
  
  /*************************************************
   * Creates the initial question, assigns the points,
   * and adds the information text.
  **************************************************/
  
  body.appendParagraph(question).setAttributes(questionStyle).appendText(' (' + points + ' pts)').setBold(true);
  body.appendParagraph(information).setAttributes(infoStyle);
  
  /*************************************************
   * Creates a table based on the variables given.
  **************************************************/
  
  var table = body.appendTable();
  
  for(var i = 0; i < lines; i++){ 
    var tr = table.appendTableRow();
    var td = tr.appendTableCell();
  }
  
  if(partialCredit == true) {
    body.appendParagraph('Partial Credit').setAttributes(creditStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  } else {
    body.appendParagraph('No Partial Credit').setAttributes(creditStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }
  
  table.setAttributes(infoStyle);
 
  /*************************************************
   * Saves and closes the document.
  **************************************************/
  
  doc.saveAndClose();
}


/*************************************************
 * This function creates the Multiple Choice question.
 **************************************************/
function addMultipleChoice() {
   
  /*************************************************
   * Obtains the document and gets the body section
   * of the given document.
  **************************************************/
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  /*************************************************
   * Stores all the given variables necessary to
   * customize the text-field question.
  **************************************************/
  
  var question = "Question 1: This is a question?";
  var points = 10;
  var information = "This is extra detail for the question given. It’s possible that there might be additional instruction to be added for a good response. This is just a bunch of filler text that I have to fill up space.";
  var options = "option1, option2, option3";
  var questionNumStyle = 1;
  var answerNum = 2;
  var partialCredit = false;
  
  /*************************************************
   * Creates the initial question, assigns the points,
   * and adds the information text.
  **************************************************/
  
  body.appendParagraph(question).setAttributes(questionStyle).appendText(' (' + points + ' pts)').setBold(true);
  body.appendParagraph(information).setAttributes(infoStyle);
  
  /************************************************
   * Creates bulleted list based on variable given.
  *************************************************/
  
  questions.push(question);
  
  
  var optionsAr = options.split(",");
  var optionsCnt = 0;
  var listId;
  var item;
  
  for(var idx in optionsAr){
    if(questionNumStyle == 1)
      item = body.appendListItem(optionsAr[idx]).setAttributes(infoStyle).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
    else
      item = body.appendListItem(optionsAr[idx]).setAttributes(infoStyle).setGlyphType(DocumentApp.GlyphType.LATIN_UPPER);
    
    if(optionsCnt == 0){
      item1 = item;
      //Logger.log(item1.getListId());
    } else
      item.setListId(item1);
    
    optionsCnt++;
    if(optionsCnt == answerNum)
      answers.push(optionsAr[idx]);
  }
  
  
  if(partialCredit == true) {
    body.appendParagraph('Partial Credit').setAttributes(creditStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  } else {
    body.appendParagraph('No Partial Credit').setAttributes(creditStyle).setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }

  /*************************************************
   * Saves and closes the document.
  **************************************************/
  
  doc.saveAndClose();
}