/*
  Created by RemcoE33
  Docs: https://fmpcloud.io/documentation
  Apps script qoutas: https://developers.google.com/apps-script/guides/services/quotas
*/

function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('fmpcloud')
    .addItem('Set API key', "storeAPIkey")
    .addItem('Run all tickers', 'getAllData')
    .addToUi();
}

function storeAPIkey() {
  const key = SpreadsheetApp.getUi().prompt('Enter API key:').getResponseText();
  ScriptProperties.setProperty('apikey', key);
}

function getAllData() {
  console.time('Timer');
  const errors = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tickersSheet = ss.getSheetByName('Tickers');
  const tickers = tickersSheet.getRange(2, 1, tickersSheet.getLastRow() - 1).getValues().flat();
  const apikey = ScriptProperties.getProperty('apikey');

  const fmpcloud = {
    Balance_Y: `https://fmpcloud.io/api/v3/balance-sheet-statement/###?limit=120&apikey=${apikey}`,
    Balance_Q: `https://fmpcloud.io/api/v3/balance-sheet-statement/###?period=quarter&limit=400&apikey=${apikey}`,
    Income_Q: `https://fmpcloud.io/api/v3/income-statement/###?period=quarter&limit=400&apikey=${apikey}`,
    Income_Y: `https://fmpcloud.io/api/v3/income-statement/###?limit=120&apikey=${apikey}`,
    CashFlow_Y: `https://fmpcloud.io/api/v3/cash-flow-statement/###?limit=120&apikey=${apikey}`,
    CashFlow_Q: `https://fmpcloud.io/api/v3/cash-flow-statement/###?period=quarter&limit=400&apikey=${apikey}`,
    Ratios: `https://fmpcloud.io/api/v3/ratios/###?limit=40&apikey=${apikey}`,
    Metrics: `https://fmpcloud.io/api/v3/key-metrics/###?limit=40&apikey=${apikey}`,
    Press: `https://fmpcloud.io/api/v3/press-releases/###?limit=100&apikey=${apikey}`,
    News: `https://fmpcloud.io/api/v3/stock_news?tickers=###&limit=100&apikey=${apikey}`,
    Surprises: `https://fmpcloud.io/api/v3/earnings-surpises/###?apikey=${apikey}`,
    Transcript: `https://fmpcloud.io/api/v3/earning_call_transcript/###?quarter=3&year=2020&apikey=${apikey}`
  }

  const urlsAndSheetnames = Object.entries(fmpcloud);

  tickers.forEach((tic, index) => {
    urlsAndSheetnames.forEach(endpoint => {
      let [sheetname, url] = endpoint;
      const tickerUrl = url.replace('###', tic);
      try{
        console.log(`${tic} | ${sheetname}`);
        handleAPI(sheetname, tickerUrl, index);
      } catch (err){
        errors.push(err);
        console.log(`${tic} | ${sheetname} | ${err}`);
      }
    })
  })

  console.timeEnd('Timer');
  
  if(errors.length > 0){
    SpreadsheetApp.getUi().alert(errors.join('## '));
  }

}

function handleAPI(sheetname, url, index) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetname);

  const response = UrlFetchApp.fetch(url);
  const dataAll = JSON.parse(response.getContentText());
  const dataRows = dataAll;

  const rowHeaders = Object.keys(dataRows[0]);
  const rows = [rowHeaders];
  for (let i = 0; i < dataRows.length; i++) {
    const rowData = [];
    for (let j = 0; j < rowHeaders.length; j++) {
      rowData.push(dataRows[i][rowHeaders[j]]);
    }
    rows.push(rowData);
  }

  if (index == 0) {
    sheet.getDataRange().clearContent();
  } else {
    rows.shift()
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

}


