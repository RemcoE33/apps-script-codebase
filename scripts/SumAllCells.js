/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Return the sum of the same range on all your sheets.
*
* @param {"A1:D10"} range - Your range to sum.
* @param {"sheet1", "Sheet2"} [optional] exclude - Sheetnames to exclude inside an array.
* @return the sum of all the values
* @customfunction
*/
function SUM_ALL_CELLS(range, exclude) {
  if (!Array.isArray(range)) {
    range = [[range]]
  };
 
  const flatRange = range.flat();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet().getName();
 
  if(!exclude){
    exclude = [sheet];
  } else if(!Array.isArray(exclude)) {
    exclude = [sheet, exclude];
  } else {
    exclude.push(sheet);
  }
  console.log(exclude);
  const sheets = ss.getSheets();
  let sum = 0;
 
  sheets.forEach(s => {
    if (!exclude.flat().includes(s.getName())) {
      flatRange.forEach(range => {
        s.getRange(range).getValues().flat().forEach(value => {
          let number = (Number(value)) ? Number(value) : 0 ;
          sum += number
          });
      });
    };
  });
 
  return sum;
 
}