/**
Checks if first value
*/

function onlyUnique(value, index, self) { 
    return self.indexOf(value) === index;
}

/**
Find all acronyms (defined as all caps text) and create a table of them.
*/
function findAcronyms(){
  var doc = DocumentApp.getActiveDocument();
  var bodyElement = doc.getBody();
  var bodyText = bodyElement.getText();  // Gets whole body as a text string
  var resultsArray = bodyText.match(/([0-9]*[A-Z]){2,}/g); 
  var cells = [];
  
  // Sort array
  resultsArray.sort();
  
  // De-dupe
  resultsArray = resultsArray.filter( onlyUnique );
  resultsArray.forEach(function(value){
    cells.push([value,'']);
  });
  
  var cursor = DocumentApp.getActiveDocument().getCursor();
  // Need error catching here...
  var element = cursor.getElement();
  var childIndex = bodyElement.getChildIndex(element);  
  
  
  // Build a table from the results array
  bodyElement.insertTable(childIndex,cells);
  bodyElement.insertParagraph(childIndex, 'Acronyms');
}

 /**
 * Create custom menu when document is opened.
 */
function onOpen() {
  DocumentApp.getUi().createMenu('Custom')
      .addItem('Insert table of acronyms', 'findAcronyms')
      .addToUi();
}
