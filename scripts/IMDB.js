/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

//Global variables
const app = SpreadsheetApp
const ss = app.getActiveSpreadsheet();
const env = PropertiesService.getScriptProperties();
const token = env.getProperty('TOKEN');

function onOpen(e) {
  app.getUi().createMenu('IMDB')
    .addItem('Refresh All', 'refreshAll')
    .addItem('Refresh Top250Movies', 'top250movies')
    .addItem('Refresh Top250Tvs', 'top250tvs')
    .addItem('Refresh MostPopularMovies', 'mostPopularMovies')
    .addItem('Refresh MostPopularTVs', 'mostPopularTVs')
    .addSeparator()
    .addItem('Set API key', 'setAPI')
    .addToUi();
}

function setAPI() {
  const key = app.getUi().prompt('API Key').getResponseText();
  env.setProperty('TOKEN', key);
}

function refreshAll() {
  top250movies()
  top250tvs()
  mostPopularTVs()
  mostPopularMovies()
}

function top250movies() {
  getIMDBtops('Top250Movies', 'Top250Movies')
}

function top250tvs() {
  getIMDBtops('Top250TVs', 'Top250TVs')
}

function mostPopularMovies() {
  getIMDBtops('MostPopularMovies', 'MostPopularMovies')
}

function mostPopularTVs() {
  getIMDBtops('MostPopularTVs', 'MostPopularTVs')
}

function getIMDBtops(endpoint, sheetname) {
  const url = `https://imdb-api.com/en/API/${endpoint}/${token}`
  const response = UrlFetchApp.fetch(url)
  const items = JSON.parse(response.getContentText()).items;
  const headers = Object.keys(items[0]);
  const imageIndex = headers.indexOf('image');
  const data = [];

  items.forEach(movie => {
    data.push(Object.values(movie))
  });

  const formulas = data.map(movie => {
    return [`=IF($B$1 = true,IMAGE("${movie[imageIndex]}"),)`];
  })

  data.unshift(headers);

  const sheet = ss.getSheetByName(sheetname);
  sheet.getRange(3, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  sheet.getRange(3, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(3, sheet.getLastColumn() + 1).setValue('Cover');
  sheet.getRange(4, sheet.getLastColumn(), formulas.length, 1).setFormulas(formulas);
}

function actorsInSameMovie() {
  const sheet = ss.getSheetByName('ActorsInSameMovie');
  const [actorOne, actorTwo] = sheet.getRange(2, 2, 1, 2).getValues().flat();
  const searchActorOne = `https://imdb-api.com/en/API/SearchName/${token}/${actorOne}`
  const searchActorTwo = `https://imdb-api.com/en/API/SearchName/${token}/${actorTwo}`
  const searchResponse = UrlFetchApp.fetchAll([searchActorOne, searchActorTwo]);

  try{
  var actorIds = searchResponse.map(seach => {
    return JSON.parse(seach.getContentText()).results[0].id;
  });
  } catch (err){
    app.getUi().alert('One of the actors is not found')
    return;
  }

  const [actorIdOne, actorIdTwo] = actorIds;

  const actorInformationOne = JSON.parse(UrlFetchApp.fetch(`https://imdb-api.com/en/API/Name/${token}/${actorIdOne}`).getContentText());
  const actorInformationTwo = JSON.parse(UrlFetchApp.fetch(`https://imdb-api.com/en/API/Name/${token}/${actorIdTwo}`).getContentText());

  const moviesOne = actorInformationOne.castMovies.filter(movie => movie.role == 'Actor');
  const moviesTwo = actorInformationTwo.castMovies.filter(movie => movie.role == 'Actor');
  const matches = moviesOne.filter(m1 => moviesTwo.some(m2 => m1.id === m2.id))
    .map(match => {
      return {...match, ...getMovieFromId(match.id)}
    });

  const pictures = [[`=IFERROR(IMAGE("${actorInformationOne.image}"),)`, `=IFERROR(IMAGE("${actorInformationTwo.image}"),)`]]
  const keysNotToProcess = ['id', 'image', 'knownFor', 'castMovies', 'errorMessage'];
  const actorsInfo = [];

  Object.keys(actorInformationOne).forEach(key => {
    if (!keysNotToProcess.includes(key)) {
      actorsInfo.push([key.toUpperCase(), actorInformationOne[key], actorInformationTwo[key]])
    };
  });

  sheet.getRange(4, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();

  if (matches.length > 0) {
    const matchesToArray = matches.map(match => {
      return Object.values(match);
    })
    matchesToArray.unshift(Object.keys(matches[0]).map(key => key.toUpperCase()));
    sheet.getRange(4,2,1,2).setFormulas(pictures);
    sheet.getRange(10, 1, actorsInfo.length, 3).setValues(actorsInfo)
    sheet.getRange(4, 5, matchesToArray.length, matchesToArray[0].length).setValues(matchesToArray);
  } else {
    sheet.getRange(4, 1).setValue('No matches found')
  }

}

function getMovieFromId(id){
  const url = `https://imdb-api.com/en/API/Title/${token}/${id}`
  const data = JSON.parse(UrlFetchApp.fetch(url).getContentText());

  const object = {
    runtime: data.runtimeMins, 
    awards: data.awards, 
    directors: data.directors, 
    genres: data.genres,
    rating: data.imDbRating,
    plot: data.plot
  }

  return object;
}

