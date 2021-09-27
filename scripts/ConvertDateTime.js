/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Returns converterd timestamps.
*
* @param {A1:B5} range - Input (multiple column) range.
* @param {"UTC"} timezone - Enter desired timezone.
* @param {"yyyy-MM-dd uu:mm:ss"} format - Enter timestamp formatting.
* @return {array} range of converted timezones.
* @customfunction
*/
function CONVERTDATETIME(range, timezone, format) {
  const output = []
  if (Array.isArray(range)) {
    range.forEach(row => {
      const outputRow = []
      row.forEach(col => {
        outputRow.push(Utilities.formatDate(new Date(col.split("+")[0]), timezone, format))
      })
      output.push(outputRow);
    })
    return output
  } else {
    const timestamp = range.toString()
    return Utilities.formatDate(new Date(timestamp.split("+")[0]), timezone, format)
  }
}