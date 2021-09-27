/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Returns definition, example and synonym.
*
* @param {A1:A5} range - Input range.
* @param {true} synonyms - Add synonyms.
* @return {array} range definitions, examples and synonyms.
* @customfunction
*/
function DICTIONARY(words, synonyms){
  if(!Array.isArray(words)){
    words = [[words]]
  }
  const output = [];
  const urls = words.flat().map(word => `https://api.dictionaryapi.dev/api/v2/entries/en/${word}`)

  const response = UrlFetchApp.fetchAll(urls);
  response.forEach(url => {
    const data = JSON.parse(url.getContentText());
    const definition = data[0].meanings[0].definitions[0].definition
    const example = data[0].meanings[0].definitions[0].example
    if (synonyms){
      const synonym = data[0].meanings[0].definitions[0].synonyms.join(', ')
      output.push([definition, example, synonym])
    } else {
    output.push([definition, example])
    }
  })

  return output;
}