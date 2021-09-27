/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Returns the values from an sheet range left or right from the current sheet.
*
* @param {"A2:B5"} range Enter a cell as a string: "A1".
* @param {"right"} direction Choose left/right: "right".
* @return value from linked sheet.
* @customfunction
*/
function CONNECTED_SHEET_VALUE(range,direction){
  
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const currentIndex = ss.getActiveSheet().getIndex();
  const sheets = ss.getSheets();

  let index = 0;

  if (direction.toLowerCase() == 'right'){
    index = 1
  } else if (direction.toLowerCase() == 'left'){
    index = -1
  }

  const sheetIndex = currentIndex + index;

  if (index == 0){
    return 'Wrong direction'
  } else if (sheetIndex > sheets.length){
    return 'No sheet in direction'
  } else {
    return sheets[sheetIndex - 1].getRange(range).getValues();
  }
}