/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

function onEdit(e) {
  const cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell()
  const rawValue = cell.getValue().toUpperCase();
  const regExpCurr = new RegExp('[A-Z]');
  const regExpValue = new RegExp('[0-9]?[0-9]?[0-9.]');
  const currValue = regExpCurr.exec(rawValue);
  const value = regExpValue.exec(rawValue);
  
  let output = 0.0;

  switch(currValue[0]){
    case 'K':
      output = Number(value[0]) * 1000;
      break;
    case 'M':
      output = Number(value[0]) * 1000000;
      break;
    case 'B':
      output = Number(value[0]) * 1000000000;
      break;
  }

  cell.setValue(output).setNumberFormat('[<999950]0.0,"K";[<999950000]0.0,,"M";0.0,,,"B"')
  
}