/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
  
  1. Select the urls you want to final url from, then invoke script via menu.
  2. Values will be places in the +1 offset column.
*/

function onOpen(e){
  const menu = SpreadsheetApp.getUi().createMenu('Utill');
  menu.addItem('Find','urlRedirects');
  menu.addToUi();
}

function urlRedirects() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const urls = [].concat(...activeRange.getValues());
  const result = [];

  urls.forEach(url => {
    result.push([findUrl(url)]);
  });

  activeRange.offset(0,1).setValues(result)
}

function findUrl(url) {

  const options = {
    'followRedirects': false,
    'muteHttpExceptions': false
  }
  const response = UrlFetchApp.fetch(url, options);
  const redirectUrl = response.getHeaders()['Location'];

  if (redirectUrl) {
    const nextUrl = findUrl(redirectUrl);
    console.log(nextUrl);
    return nextUrl;
  } else {
    console.log(url);
    return url;

  }
}