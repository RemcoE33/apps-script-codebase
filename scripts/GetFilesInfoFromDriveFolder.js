/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Utill");
  menu.addItem("Run function", "getFileInfo");
  menu.addToUi();
}

function getFileInfo() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = SpreadsheetApp.getUi().prompt("Enter sheetname, creates a new one").getResponseText();
  ss.insertSheet(name);
  const outputSheet = ss.getSheetByName(name);
  const driveId = SpreadsheetApp.getUi().prompt("Enter drive folder ID").getResponseText();
  const driveFolder = DriveApp.getFolderById(driveId);
  const driveFiles = driveFolder.getFiles();
  
  
  const output = [["ID","Name","URL","Create date","Sharing permission"]]

  while(driveFiles.hasNext()){
    let file = driveFiles.next();
    output.push([file.getId, file.getName(), file.getUrl(), file.getDateCreated(), file.getSharingPermission()]);
  }

  outputSheet.clearContents();
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

}