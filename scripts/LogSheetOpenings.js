/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase

  1. Run setup once.
  2. Cange your email so this will be excluded from the count.
  3. Read count with function: 'getCurrentNumber'
*/

//Change your email below.
const email = "youremail@gmail.com";

function setup(){
  const counter = PropertiesService.getDocumentProperties();
  counter.setProperty("count", 0);
}

function onOpen(){
  const counter = PropertiesService.getDocumentProperties();
  const current = counter.getProperty("count");
  const number = Number(current)+1;
  if (Session.getActiveUser().getEmail() != email){
  counter.setProperty("count", number);
  }
  
}

//Read the current count
function getCurrentNumber(){
  const counter = PropertiesService.getDocumentProperties();
  console.log(counter.getProperty("count"));
}