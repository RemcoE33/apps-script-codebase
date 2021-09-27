/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  From codeline 28 are the functions you can invoke via menu "Transcript".
*/

/**
* Returns selected transscript from financialmodelingprep.
*
* @param {"AAPL"} ticker - Input the stock ticker.
* @param {3} quarter - Input quarter number
* @param {2020} year - Input year number
* @param {"demo"} apikey - Input your apikey
* @return {array} transscript
* @customfunction
*/
function GETTRANSCRIPT(ticker = 'AAPL', quarter = 3, year = 2020, apikey = 'demo'){
  const url = `https://financialmodelingprep.com/api/v3/earning_call_transcript/${ticker}?quarter=${quarter}&year=${year}&apikey=${apikey}`
  const res = UrlFetchApp.fetch(url)
  const data = JSON.parse(res.getContentText())
  const split = data[0].content.split(/\n/);
  const output = []
  split.forEach(line => { output.push([line])})
  return output
}


// -------- Functions you can use on other triggers or menu:
function onOpen(e){
  const menu = SpreadsheetApp.getUi().createMenu('Transcript')
    .addItem('Get transcript in current cell', 'transcript')
    .addItem('Set api key', 'setApiKey')
    .addItem('Get stored api key', 'readApiKey')
    .addToUi()
}

function transcript(){
  const paramsRaw = SpreadsheetApp.getUi().prompt(`Enter url parameters '|' separated->: AAPL|3|2020 `).getResponseText()
  const params = paramsRaw.split("|")
  const prop = PropertiesService.getScriptProperties()
  const token = prop.getProperty('TOKEN')
  const response = GETTRANSCRIPT(params[0],params[1],params[2],params[3],(token) ? token : 'demo')

  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const richTextValues = [];

  response.forEach(row =>{
    const text = row[0]
    const colonPositions = text.indexOf(':')
    richTextValues.push([SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setTextStyle(0, colonPositions, SpreadsheetApp.newTextStyle()
      .setBold(true)
      .build())
      .build()])
  })

  ss.getRange(ss.getCurrentCell().getRow(),ss.getCurrentCell().getColumn(),response.length,1).setRichTextValues(richTextValues)

}

function setApiKey(){
  const token = SpreadsheetApp.getUi().prompt('Enter api key:').getResponseText()
  const prop = PropertiesService.getScriptProperties()
  prop.setProperty('TOKEN', token)
}

function readApiKey(){
  const prop = PropertiesService.getScriptProperties()
  const token = prop.getProperty('TOKEN')
  SpreadsheetApp.getUi().alert((token) ? token : "No token is set")
}