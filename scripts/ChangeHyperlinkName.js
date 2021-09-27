/*
  RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  Changes all your selected hyperlinks text to the one you enter in the prompt.
*/

function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu('Utill')
      .addItem('Change hyperlink text', 'hyperLinkNameChange')
      .addToUi()
}

function hyperLinkNameChange(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const name = ui.prompt('Link name?').getResponseText();
  const dataRange = ss.getActiveRange();
  const data = dataRange.getFormulas();
  
  var array = [];
  
  for (var i=0; i < data.length; i++){
    var values = data[i];
    var url = /"(.*?)"/.exec(values)[1];
    var newUrl = '=HYPERLINK' + '("'+url+'","'+name+'")';
    array.push([newUrl]);
  }

  Logger.log(data.length);
  Logger.log(array);
  dataRange.setFormulas(array);
  
}