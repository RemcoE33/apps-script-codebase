/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
  
  Change sheetname to your needs.
*/

const sheetname = 'Coins';

function onOpen(e){
  SpreadsheetApp.getUi()
    .createMenu('WalletCoins')
      .addItem('Update coins', 'getCoins')
      .addToUi()
}

function getCoins(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(sheetname)
 
  const response = UrlFetchApp.fetch('https://wallet-api.staging.celsius.network/util/interest/rates')
  const data = JSON.parse(response.getContentText()).interestRates
  const output = [["Coin","Rate","Id","Name","Short"]]
  const images = []
 
  data.forEach(coin => {
    output.push([coin.coin,coin.rate,coin.currency.id,coin.currency.name,coin.currency.short])
    images.push([`=IMAGE("${coin.currency.image_url}")`])
  })
 
  sheet.getDataRange().clearContent()
  
  sheet.getRange(1,1,output.length,5).setValues(output)
  sheet.getRange(1,6).setValue("Image")
  sheet.getRange(2,6,images.length).setFormulas(images)
  sheet.getRange("1:1").setFontWeight('bold')
  
}