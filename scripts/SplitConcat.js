/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Returns range for each split.
*
* @param {A1:B5} range - Input range.
* @param {3} column - Enter the column number that needs to be split.
* @param {","} delimiter - Enter the delimiter. 
* @return {array} range with values concatenated to the split outcome.
* @customfunction
*/

function SPLIT_CONCAT(range, column, delimiter) {
  
  const output = [];
  range.forEach(row => {
    const split = row[column - 1].split(delimiter);
    row.splice(column - 1,1);
      split.forEach(col => {
       output.push([...row,col.trim()]);
    })
  })

  return output;

}