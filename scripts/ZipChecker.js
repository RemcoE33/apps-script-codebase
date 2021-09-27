/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
 * Return zipcode information.
 *
 * @param {"99501"} zipcode Enter zipcode.
 * @param {"US"} country Enter country code.
 * @return zipcode information.
 * @customfunction
 */
function ZIPCHECKER(zipcode, country = "US") {
  const options = {
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(
    `http://api.zippopotam.us/${country}/${zipcode}`,
    options
  );
  if (response.getResponseCode()[0] >= 400) {
    return "Zipcode not found";
  }
  const zip = JSON.parse(response.getContentText());

  return [
    [
      zip.country,
      zip.places[0].state,
      zip.places[0]["place name"],
      zip.places[0].latitude,
      zip.places[0].longitude,
    ],
  ];
}
