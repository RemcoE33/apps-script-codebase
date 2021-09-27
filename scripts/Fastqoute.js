/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  Change sheetname (codeline 19), this function asumes the tickers are in Col A from row 2.
*/

const ss = SpreadsheetApp.getActiveSpreadsheet();

function onOpen(e){
  SpreadsheetApp.getUi().createMenu('Fastqoute')
    .addItem('Update Latest Info','getTickerData')
    .addItem('Update Chart', 'getChartData')
    .addItem('Update Quote', 'getQuote')
    .addToUi();
}

function getTickerData() {
  const sheet = ss.getSheetByName('LatestInfo');
  const tickers = sheet.getRange(2,1,sheet.getLastRow()-1).getValues().flat();
  const output = [];

  tickers.forEach((tic, i) => {
  const url = `https://fastquote.fidelity.com/service/quote/nondisplay/json?productid=research&symbols=${tic}&quotetype=D&callback=jQuery1111012350412486796825_1630778346041&_=1630778346042`

  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText()
    .replace('jQuery1111012350412486796825_1630778346041(','')
    .replace(')','')
    .trim())
    .QUOTE.SYMBOL_RESPONSE.QUOTE_DATA;
  if(i == 0){
    const headers = ['SECTOR', 'INDUSTRY', ...Object.keys(data)]
    output.push(headers);
  }
  const sectorInfo = getSector(tic);
  output.push([...sectorInfo, ...Object.values(data)]);
  });

  sheet.getRange(2,2,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
  sheet.getRange(1,2,output.length, output[0].length).setValues(output);
}

function getChartData(){
  const ticker = ss.getSheetByName('Chart').getRange(1,2).getValue();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const twoYearsAgo = Utilities.formatDate(new Date(new Date().getTime() - 63113852000), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const url = `https://fastquote.fidelity.com/service/historical/chart/lite/json?productid=research&symbols=${ticker}&dateMin=${twoYearsAgo}:00:00:00&dateMax=${today}:00:00:00&intraday=n&granularity=1&incextendedhours=n&dd=y&callback=jQuery1111012350412486796825_1630778346043&_=1630778346044`

  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText()
  .replace('jQuery1111012350412486796825_1630778346043(','')
  .replace(')','')
  .trim());
  const output = [];
  
  data.SYMBOL[0].BARS.CB.forEach(object => {
    output.push(Object.values(object));
  })

  const sheet = ss.getSheetByName('ChartData');
  sheet.getRange(2,1,sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  sheet.getRange(2,1,output.length, output[0].length).setValues(output);
}

function getQuote(){
  const url = 'https://fastquote.fidelity.com/service/quote/json?productid=embeddedquotes&subproductid=eresearch&market_close=1&symbols=.DJI%2C.IXIC%2C.SPX&dojo.preventCache=1630779480852&callback=dojo.io.script.jsonp_dojoIoScript2._jsonpCallback'

  const output = [];
  const response = UrlFetchApp.fetch(url);
  const qoutes = JSON.parse(response.getContentText()
    .replace('dojo.io.script.jsonp_dojoIoScript2._jsonpCallback(','')
    .replace(')','')
    .trim())
    .QUOTES;
  const keys = Object.keys(qoutes);

  keys.forEach((key, i) => {
    const q = qoutes[key];
    if(i == 0){
      const headers = Object.keys(q)
      output.push(['QOUTE', ...headers]);
    }
    output.push([key,...Object.values(q)]);
  })

  ss.getSheetByName('Quotes').getRange(1,1,output.length,output[0].length).setValues(output);
}

function getSector(ticker){
  const url = `https://eresearch.fidelity.com/eresearch/evaluate/snapshot.jhtml?symbols=${ticker}`;
  const html = UrlFetchApp.fetch(url).getContentText();
  let sector;
  let industries;
  html.match(/markets_sectors(.*?)>(.*?)<\/a>/gmi).forEach((element, index) => {
    switch(index){
      case 2:
        sector = />(.*?)</gmi.exec(element)[1]
      break;
      case 3:
        industries = />(.*?)</gmi.exec(element)[1]
      break;
    }
});
return [sector, industries];
}
