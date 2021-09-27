/*
  RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
  https://www.alphavantage.co/documentation/

  The "set column headers from active cell" places the JSON keys to that specific row, 
  from that specific api type.
*/


function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('AlphaVantage')
    .addItem('GetAlphavantage from selected ticker(s)', 'alphavantage')
    .addItem('Get specific value from selected ticker(s)', 'alphavantageSingle')
    .addItem('Set column headers from active cell', 'columnheaders')
    .addItem('Set API key', 'setToken')
    .addToUi();
}

function alphavantage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const tickers = activeRange.getValues().flat();
  const env = PropertiesService.getScriptProperties().getProperty('TOKEN');
  const output = [];

  let functionType;
  try {
    functionType = prompt();
  } catch (err) {
    return;
  }

  tickers.forEach(tic => {
    const url = `https://www.alphavantage.co/query?function=${functionType}&symbol=${tic}&apikey=${env}`;
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());
    const values = Object.values(json)
    values.shift();
    output.push(values);
  })

  sheet.getRange(activeRange.getRow(), activeRange.getColumn() + 1, output.length, output[0].length).setValues(output);

}

function alphavantageSingle(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const activeRange = sheet.getActiveRange();
  const tickers = activeRange.getValues().flat();
  const env = PropertiesService.getScriptProperties().getProperty('TOKEN');
  const output = [];

  let functionType;
  let key;
  try {
    functionType = promptType();
    key = promptKey(functionType);
  } catch (err) {
    return;
  }

  tickers.forEach(tic => {
    const url = `https://www.alphavantage.co/query?function=${functionType}&symbol=${tic}&apikey=${env}`;
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());
    output.push([json[key]]);
  })

  sheet.getRange(activeRange.getRow(), activeRange.getColumn() + 1, output.length, output[0].length).setValues(output);
}

function columnheaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();
  let functionType;
  try {
    functionType = promptType();
  } catch (err) {
    return;
  }
  const url = `https://www.alphavantage.co/query?function=${functionType}&symbol=IBM&apikey=demo`;
  const response = UrlFetchApp.fetch(url);
  const headers = Object.keys(JSON.parse(response.getContentText()));
  const range = sheet.getRange(cell.getRow(), cell.getColumn(), 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold')
}

function promptType() {
  const choices = ['OVERVIEW', 'EARNINGS', 'INCOME_STATEMENT', 'BALANCE_SHEET', 'CASH_FLOW', 'LISTING_STATUS', 'EARNINGS_CALENDAR', 'IPO_CALENDAR'];
  const ui = SpreadsheetApp.getUi();
  const type = ui.prompt(`Choose:  ${choices.join(' | ')}`).getResponseText();
  if (choices.includes(type)) {
    return type;
  } else {
    ui.alert('Input type does not match one of the choises')
    return new Error('No match')
  }
}

function promptKey(type) {
  const url = `https://www.alphavantage.co/query?function=${type}&symbol=IBM&apikey=demo`;
  const response = UrlFetchApp.fetch(url);
  const headers = Object.keys(JSON.parse(response.getContentText()));
  const ui = SpreadsheetApp.getUi()
  const key = ui.prompt(`Choose:  ${headers.join(' | ')}`).getResponseText();
  if (headers.includes(key)) {
    return key;
  } else {
    ui.alert('Input type does not match one of the choises')
    return new Error('No match')
  }
}

function setToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Set token');
  PropertiesService.getScriptProperties().setProperty('TOKEN', response.getResponseText());
}