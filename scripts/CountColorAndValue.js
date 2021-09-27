/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  Sample: =COUNT_COLOR_VALUE(A1:C10,GET_BACKGROUND(A1), "RemcoE33")
*/


/**
* Return count of background in combination of value
*
* @param {A1:D5} range Enter (multicolom) range.
* @param {"#FFFFF"} color Enter colorcode as string.
* @param {"Lifeline"} criteria Enter criteria.
* @return count of cells with given colorbackgrond and criteria.
* @customfunction
*/
function COUNT_COLOR_VALUE(range, color, criteria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const formula = ss.getActiveRange().getFormula()
  const colorRange = /\((.*?),/.exec(formula)[1];
  const backgrounds = ss.getRange(colorRange).getBackgrounds()
  
  let count = 0

  for (i = 0; i < backgrounds.length; i++) {
    for (j = 0; j < backgrounds[0].length; j++) {
      if (backgrounds[i][j] == color && range[i][j] == criteria) {
        count++
      }
    }
  }

  return count
}

/**
* Returns background hex color
*
* @param {A1} cell Enter cell reference to get the background
* @return background hex color
* @customfunction
*/
function GET_BACKGROUND(cell){
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const formula = ss.getActiveRange().getFormula().toLocaleUpperCase();
  const colorRange = /GET_BACKGROUND\((.*?)\)/.exec(formula)[1];
  return ss.getRange(colorRange).getBackground().toString();
}
