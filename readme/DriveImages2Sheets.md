# apps-script-driveImage2sheets
Get images from a drive file and insert to google sheets with the =IMAGE() formula.


### Be aware:
This script will set the image 'access to everyone with link' as viewer. This may not work on Workspace users. Depends on the admin settings. You have a script runtime limitation, so maybe you need to batch process.

### Installation:
  - Tools -> Script editor.
  - Clear the little code you see and past the code from below.
  - Optional: change the , to ; on codeline 58 / 60 if you have sheets formula's with ;.
  - Execute once, give permission and ignore the error.
  - Close the script editor.
  - Refresh your spreadsheet browser tab.

### Use:
Now you see a new menu: "Drive images" in there there are 4 options:

### Setup:
  - Enter google drive folder id where the images are stored (if you need to batch proces, delete the images that are done and add new ones)
  - Choose image filetype: png / jpeg / gif / svg
  - Choose image mode: 1 / 2 / 3 (4 is not supported in this script)
    1 = resizes the image to fit inside the cell, maintaining aspect ratio.
    2 = stretches or compresses the image to fit inside the cell, ignoring aspect ratio.
    3 = leaves the image at original size, which may cause cropping.
  - On / off switch. If you leave blank then nothing, if you want a on off switch then enter the cell A1Notation like: A1. This wrap the =IMAGE inside a IF         statement. This will make a checkbox in that cell. If it is checked the =IMAGE formula will be used, if it is unchecked then blank.

## Run preconfigured
  - Run the script with the settings above.

## Run manually
  - Run the script manually. So you will get the same questions as Setup 1-4.

## Download url's
  - Creates a list with filenames and drive download url's.

