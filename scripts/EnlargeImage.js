/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  Select the cell that contains the =IMAGE formula, then run this function to enlage the image.
*/

function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu('Utill')
      .addItem('Enlarge', 'enlargeImage')
      .addToUi()
}

function enlargeImage() {
  const formula = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getFormula();
  const url = /IMAGE\(\"(.*?)\"/g.exec(formula)[1];
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          img {
              max-width: 100%;
              max-height: 100%;
          }
        </style>
        <base target="_top">
      </head>
      <body>
        <div class=img>
          <img src="${url}">
        </div>
      </body>
    </html>
  `
  SpreadsheetApp.getUi().showDialog(HtmlService.createHtmlOutput(html).setHeight(500).setWidth(800));
}