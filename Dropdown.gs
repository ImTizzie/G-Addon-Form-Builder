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

var titleStyle = {};
titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
titleStyle[DocumentApp.Attribute.FONT_SIZE] = 18;
titleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';

var headerStyle = {};
headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
headerStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';


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
function addTextField(title,response,answer,points,is_graded) {

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
  
  var question = title;
  var ans = answer;
  var points = points;
  var lines = response;
  
  if(is_graded) {
    questions.push(question);
    answers.push(ans);
    body.appendParagraph(question).setAttributes(questionStyle).appendText(' (' + points + ' pts)').setBold(true);
  }else{
      /*************************************************
   * Creates the initial question, assigns the points,
   * and adds the information text.
  **************************************************/
  
  body.appendParagraph(question).setAttributes(questionStyle).setBold(true);
  }
  
  /*************************************************
   * Creates a table based on the variables given.
  **************************************************/
  
  var table = body.appendTable();
  
  for(var i = 0; i < lines; i++){ 
    var tr = table.appendTableRow();
    var td = tr.appendTableCell();
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
function addMultipleChoice(question, options, questionNumStyle, is_graded, points, answerNum) {
   
  /*************************************************
   * Obtains the document and gets the body section
   * of the given document.
  **************************************************/
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  /*************************************************
   * Creates the initial question, assigns the points,
   * and adds the information text.
  **************************************************/
  
  if(is_graded) body.appendParagraph(question).setAttributes(questionStyle).appendText(' (' + points + ' pts)').setBold(true);
  else body.appendParagraph(question).setAttributes(questionStyle);
  //body.appendParagraph(information).setAttributes(infoStyle);
  
  /************************************************
   * Creates bulleted list based on variable given.
  *************************************************/
  
  if(is_graded) questions.push(question);
  
  //console.log("here");
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
    } else
      item.setListId(item1);
    
    optionsCnt++;
    if(optionsCnt == answerNum && is_graded){
      answers.push(optionsAr[idx]);
      //console.log(optionsAr[idx]);
    }
  }
  
  //body.appendParagraph(questions.length);
  

  /*************************************************
   * Saves and closes the document.
  **************************************************/
  
  doc.saveAndClose();
}


/*************************************************
 * This function creates the Header
 **************************************************/
function addHeader(month, year, day, title, name) {
  /*************************************************
   * Obtains the document and gets the header section
   * of the given document.
  **************************************************/
  
  var doc = DocumentApp.getActiveDocument();
  try{
    var header = doc.addHeader();
  } catch(e) {
    var header = doc.getHeader();
    header.clear();
  }

  /*************************************************
   * Creates the header by adding name, date, and title
  **************************************************/
  
  header.appendParagraph("Name: "+name+"\t\t\t\t\t\t\t\t").setAttributes(headerStyle).appendText("Date: "+month+"/"+day+"/"+year);
  header.appendParagraph(title).setAttributes(titleStyle).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  /*************************************************
   * Saves and closes the document.
  **************************************************/
  
  doc.saveAndClose();   

}

/*************************************************
 * This function generates the answer sheet for the quiz.
 **************************************************/

function addAnswerSheet(){
  
  /*************************************************
   * Obtains the document and gets the body section
   * of the given document.
   *************************************************/
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  
  /*questions.push("q1");
  answers.push("a1");
  
  questions.push("q2");
  answers.push("a2");
  
  questions.push("q3");
  answers.push("a3");*/
  
  /* body.appendParagraph(questions.length);
  /******************************************************************
   * Loops through the questions and answers and adds it to the sheet 
   ******************************************************************/
  
  for(var i = 0; i < questions.length; i++){
    body.appendParagraph(questions[i]).setAttributes(questionStyle);
    body.appendListItem(answers[i]).setAttributes(infoStyle).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);;
  }
  
  doc.saveAndClose();
}