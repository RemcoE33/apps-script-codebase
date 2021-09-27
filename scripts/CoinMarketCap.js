/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  1. Change sheetname on codeline 22.
  2. Change API token on codeline 27.

  This script will get the latest info for the top 5000 coins.
*/

function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu('CoinMarketCap')
      .addItem('Update coins', 'coinMarketCap')
      .addToUi()
}

function coinMarketCap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //Change sheetname
  const sheet = ss.getSheetByName("RemcoE33")
  const url = `https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?start=1&limit=5000`
  
  //Change API key
  const headers = {
    "X-CMC_PRO_API_KEY": 'd7e20337-xxxx-xxxx-xxxx-7ed99dd02903'

  };
  
  const options = {
    "method" : "get",
    "headers" : headers 
  };

  const response = UrlFetchApp.fetch(url, options);
  const resonseData = JSON.parse(response.getContentText()).data;
  const newObject = resonseData.map(obj => {
    const coin = obj.symbol
    return {
      id: obj.id,
      name: obj.name,
      symbol: obj.symbol,
      slug: obj.slug,
      cmc_rank: obj.cmc_rank,
      num_market_pairs: obj.num_market_pairs,
      circulating_supply: obj.circulating_supply,
      total_supply: obj.total_supply,
      max_supply: obj.max_supply,
      last_updated: new Date(obj.last_updated),
      date_added: new Date(obj.date_added),
      tags: obj.tags.join(', '),
      platform: obj.platformm,
      USD_price: obj.quote.USD.price,
      USD_volume_24h: obj.quote.USD.volume_24h,
      USD_percent_change_1h: obj.quote.USD.percent_change_1h,
      USD_percent_change_24h: obj.quote.USD.percent_change_24h,
      USD_percent_change_7d: obj.quote.USD.percent_change_7d,
      USD_market_cap: obj.quote.USD.market_cap,
      USD_last_updated: new Date(obj.quote.USD.market_cap)
    }
  })

  const columnNames = Object.keys(newObject[0])
  const rowValues = newObject.map(row => { return Object.values(row) })


  //Entiere sheet get's cleared with every call.
  sheet.getDataRange().clearContent()
  sheet.getRange(1, 1, 1, columnNames.length).setValues([columnNames]);
  sheet.getRange(2, 1, rowValues.length, rowValues[0].length).setValues(rowValues);
}