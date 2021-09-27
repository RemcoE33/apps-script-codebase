/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Returns transposed array split over n rows.
*
* @param {A1:A21} array - Input (single column) range.
* @param {7} skipping - Enter desired n rows for splitting.
* @return {array} transposed dataset split by skipping.
* @customfunction
*/
function NTRANSPOSE(array, skipping){
  const data = array.flat()
  const res = [];
    for (let i = 0; i < data.length; i += skipping) {
        const chunk = data.slice(i, i + skipping);
        res.push(chunk);
    }
  return res;
}