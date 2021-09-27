/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  Change the two global variables to your needs.
*/

//Global:
const year = 2021;
const sheetname = 'Inc5000';


function onOpen(e){
  SpreadsheetApp.getUi().createMenu('INC')
    .addItem('Refesh', 'INC')
    .addToUi();
}

function INC(){
  console.time('Timer');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetname);
  //Change year if needed.
  const url = `https://api.inc.com/rest/i5list/${year}`
  const res = UrlFetchApp.fetch(url);
  const data = JSON.parse(res.getContentText());
  const output = [];

  data.companies.map((comp, index) => {
    const object = {
      Rank: comp.rank,
      Company: comp.company,
      Industy: comp.industry,
      Founded: comp.founded,
      Website: comp.website,
      'Years on list': comp.yrs_on_list,
      Growth: comp.growth,
      Revenue: comp.raw_revenue,
      'Previous workers': comp.previous_workers,
      Workers: comp.workers,
      Metro: comp.metrocode,
      State: comp.state_l,
      Zipcode: comp.zipcode,
    };

    if (index == 0){
      output.push(Object.keys(object));
    }

    output.push(Object.values(object));
  })

  sheet.getDataRange().clearContent();
  sheet.getRange(1,1,output.length, output[0].length).setValues(output);
  console.timeEnd('Timer');
}