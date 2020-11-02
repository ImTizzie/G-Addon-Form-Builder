/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

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
  .addItem('testing', 'showTesting')
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

/**
 * These showSidebar-type functions are utilized to open up
 * the sidebar containing the element's user interface.
 */
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
