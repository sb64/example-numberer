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
      .addItem('Show Sidebar', 'showSidebar')
      .addSeparator()
      .addItem('Insert Example Number', 'insertReference')
      .addItem('Update Example Numbers', 'refreshReferences')
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
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Example Numberer');
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Inserts a new paragraph with the form of a reference
 */
function insertReference() {
  const doc = DocumentApp.getActiveDocument();
  let cursor = doc.getCursor();  
  
  let insertedText = cursor.insertText("(1) ");
  let newPosition = doc.newPosition(insertedText, 4);
  doc.setCursor(newPosition);
  refreshReferences();
}


/**
 * Refreshes the references
 */
function refreshReferences() {  
  const body = DocumentApp.getActiveDocument().getBody();
    
//  replaceSequentially(body, "^\\(\\d+\\)(.*)$", (count, match) => match.replace(/^\(\d+\)/, `(${count})`));
  replaceSequentially(body, "^\\(\\d+\\)", (count, match) => match.replace(/^\(\d+\)/, `(${count})`));
}

function replaceSequentially(body, regexString, replacerFn) {
  let foundRef = body.findText(regexString);
  let count = 1;
  while (foundRef !== null) {
    const {element, end, start} = unwrapRangeElement(foundRef);
    const text = element.asText();
    const match = text.getText().slice(start, end + 1);
    text.deleteText(start, end);
    text.insertText(start, replacerFn(count, match));
    text.setBold(start, end, false);
    text.setItalic(start, end, false);
    text.setStrikethrough(start, end, false);
    text.setUnderline(start, end, false);
    ++count;
    foundRef = body.findText(regexString, foundRef);
  }
}

function unwrapRangeElement(element) {
  return {
    element: element.getElement(),
    end: element.getEndOffsetInclusive(),
    start: element.getStartOffset(),
  }
}
