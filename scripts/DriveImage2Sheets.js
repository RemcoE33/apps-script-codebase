/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Drive images")
    .addItem("Setup", "setup")
    .addItem("Run preconfigured", "preconfigured")
    .addItem("Run manual", "manual")
    .addItem(`Download url's`, 'downloadUrls')
    .addToUi();
}

function setup() {
  const ui = SpreadsheetApp.getUi();
  const driveFolder = ui.prompt("Enter google drive folder id").getResponseText().trim()
  const imageType = `image/${ui.prompt("Enter image type: (png / jpeg / gif / svg").getResponseText().toLowerCase().trim()}`
  const mode = Number(ui.prompt("Image mode ( https://support.google.com/docs/answer/3093333?hl=en )").getResponseText().trim());
  const onOff = ui.prompt("If you want a on / off switch enter a cell notation (A1) if not leave blank").getResponseText().trim();
  const propertyService = PropertiesService.getScriptProperties();
  propertyService.setProperties({ 'folder': driveFolder, 'image': imageType, 'mode': mode, 'onOff': onOff });
}

function preconfigured() {
  const propertyService = PropertiesService.getScriptProperties();
  const driveFolder = propertyService.getProperty('folder');
  const imageType = propertyService.getProperty('image');
  const mode = Number(propertyService.getProperty('mode'));
  const onOff = propertyService.getProperty('onOff');
  const images = DriveApp.getFolderById(driveFolder).getFilesByType(imageType);

  _processImages(images, mode, onOff);

}

function manual() {
  const ui = SpreadsheetApp.getUi();
  const driveFolder = ui.prompt("Enter google drive folder id").getResponseText().trim()
  const imageType = `image/${ui.prompt("Enter image type: (png / jpeg / gif / svg").getResponseText().toLowerCase().trim()}`
  const mode = Number(ui.prompt("Image mode ( https://support.google.com/docs/answer/3093333?hl=en )").getResponseText().trim());
  const onOff = ui.prompt("If you want a on / off switch enter a cell notation (A1) if not leave blank").getResponseText().trim();
  const images = DriveApp.getFolderById(driveFolder).getFilesByType(imageType);

  _processImages(images, mode, onOff);

}

function _processImages(images, mode, onOff) {
  const output = [];

  while (images.hasNext()) {
    const file = images.next();
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)
    const downloadUrl = file.getDownloadUrl();
    if (onOff) {
      output.push([`=IF(${onOff} = TRUE,IMAGE("${downloadUrl}",${mode}),)`])
    } else {
      output.push([`=IMAGE("${downloadUrl}",${mode})`])
    }
  }
  if (onOff) {
    SpreadsheetApp.getActiveSheet().getRange(1, 1).insertCheckboxes();
    SpreadsheetApp.getActiveSheet().getRange(2, 1, output.length, 1).setFormulas(output);
  } else {
    SpreadsheetApp.getActiveSheet().getRange(1, 1, output.length, 1).setFormulas(output);
  }
  SpreadsheetApp.getUi().alert(`Processed ${output.length} images`)
}

function downloadUrls(){
  const ui = SpreadsheetApp.getUi();
  const driveFolder = ui.prompt("Enter google drive folder id").getResponseText().trim()
  const imageType = `image/${ui.prompt("Enter image type: (png / jpeg / gif / svg").getResponseText().toLowerCase().trim()}`
  const images = DriveApp.getFolderById(driveFolder).getFilesByType(imageType);

  const output = [['Filename',['Download url']]];

  while (images.hasNext()) {
    const file = images.next();
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)
    const fileName = file.getName();
    const downloadUrl = file.getDownloadUrl();
    output.push([fileName,downloadUrl])
  }

  SpreadsheetApp.getActiveSheet().getRange(1,1,output.length,2).setValues(output);

}