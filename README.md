# Google Apps Script - Functions and scripts.

### Why:
I wrote quite a lot of (custom) functions to extend Google Sheets with some functionality or exploid some API. Mostly on Reddit ([r/GoogleAppsScript](https://www.reddit.com/r/GoogleAppsScript/) / [r/googlesheets](https://www.reddit.com/r/googlesheets/) / [r/sheets](https://www.reddit.com/r/sheets/)). So i thought to create an code base for functions that could be usefull for others as well.

### Tip:
Almost all functions have a custom menu in there. I you are using multiple, you need merge the .addItem(). There is also the option to invoke the 'not custom functions' via an shortcut:

1. Tools -> Macro -> Import -> Choose function.
2. Tools -> Macro -> Manage -> Assign shortcut.


### Scripts:

**Script**|**Tags**|**ReadMe / Invoke**|**Sample sheet**|**Description**
:--|:--|:--|:--|:--
[AlphaVantage](/scripts/AlphaVantage.js)|API, Finanacial, Stock|-|-|Get information from [AlphaVantage](https://www.alphavantage.co/documentation/)
[ChangeHyperlinkName](/scripts/ChangeHyperlinkName.js)|Sheets, Utility|-|-|Sets the name of the selected =HYPERLINK() formula's to the same name.
[CoinMarketCap](/scripts/CoinMarketCap.js)|API, Finanacial, Stock|-|-|Get your [API key](https://polygon.io/dashboard/signup). Gets the top 5000 coins in your speadsheet.
[ConnectedSheetValue](/scripts/ConnectedSheetValue.js)|Custom function|`=CONNECTED_SHEET_VALUE(range,direction)`|-|Get values from the sheet to the left or to the right.
[ConvertDateTime](/scripts/ConvertDateTime.js)|Custom function|`=CONVERTDATETIME(range, timezone, format)`|-|Converts (multiple) columns to the timezone you enter. You also can give up your formatting.
[CountColorAndValue](/scripts/CountColorAndValue.js)|Custom function|`=COUNT_COLOR_VALUE(F3:G10,GET_BACKGROUND(F3),"RemcoE33")`|[Sheet](https://docs.google.com/spreadsheets/d/1lSdkQpvgQVTa2bQCpSRWqCsWzQfAlNNgDMKiKvBwMJo/edit?usp=sharing)|Count values that meets the value and background color critera.
[Dictionary](/scripts/Dictionary.js)|Sheets, Utility, Custom functions|`=DICTIONARY(words, synonyms)`|-|Get definition and/or synonyms from a word
[DriveImage2Sheets](/scripts/DriveImage2Sheets.js)|Drive, Image|[ReadMe](/readmes/DriveImage2Sheets.md)|-|Get images from specific folder on your drive and insert =IMAGE() formula's. Option: only get the download link. Option: Create a checkbox to activate the =IMAGE() formula.
[EnlargeImage](/scripts/EnlargeImage.js)|Sheets, Utility|-|-|Active a cell that contains an =IMAGE() formula. Invoke the function via Marcos or shortcut (after installing the shortcut via macros -> manage)
[Fastqoute](/scripts/Fastqoute.js)|API, Finanacial, Stock| |-|Get fastqoute data to sheets.
[FinalUrl](/scripts/FinalUrl.js)|URL fetch|-|-|Gets final redirected url from a url.
[FMcloud](/scripts/FMcloud.js)|API, Finanacial, Stock|-|[Sheet](https://docs.google.com/spreadsheets/d/1CMWLlsNE_jbIbmZLyLnu7AEmrEQW0jYcLz-P0uFO8s8/edit?usp=sharing)|Get data from [FMcloud](https://fmpcloud.io/documentation)
[GetFilesInfoFromDriveFolder](/scripts/GetFilesInfoFromDriveFolder)|Drive|-|-|Gets info form all the files inside a specific drive folder.
[GetFinanacialTransscripts](/scripts/GetFinanacialTransscripts.js)|API, Finanacial, Stock, Custom function|`=GETTRANSCRIPT("ticker",periode,year,"apikey")`|-|Needs API key. Get transscript from specific ticker. Can be invoked via custom formula or via trigger inside script editor. [docs](https://financialmodelingprep.com/developer/docs)
[IMDB](/scripts/IMDB.js)|API|-|[Sheet](https://docs.google.com/spreadsheets/d/1V8V02-Vxkry3WbA86iUPEwOfboFszYMe8yku39KzsNw/edit?usp=sharing)|[IMDB docs](https://imdb-api.com/) to get your token. This sheet will load the top250 movies and tv series, plus the option to find movies between two actors.
[INC500](/scripts/INC500.js)|API|-|-|Get the 500 companys from INC to your sheet.
[LogSheetOpenings](/scripts/LogSheetOpenings.js)|onOpen|-|-|Counts the number of times the sheets is opend (excl. yourself)
[NTranspose](/scripts/NTranspose.js)|Sheets, Utility|`=NTRANSPOSE(range, number)`|-|Transposes a single column every n number of rows
[NumberShortHandToRealNumber](/script/NumberShortHandToRealNumber.js)|onEdit(e)|-|-|Type '1K' and the scripts set numberformat to '1K' but the value to 100000 (1k/1m/1b)
[OutOfOffice](/scripts/OutOfOffice.js)|Gmail|-|-|Sets and unset the Gmail out-of-office response based on your ow configured (time) triggers. Example: 17:00 ON --> 09:00 OFF.
[SplitConcat](/scripts/SplitConcat.js)|Custom function|`=SPLIT_CONCAT(range, columnNumberOfValuesToSplit, delimiter)`|[Sheet](https://docs.google.com/spreadsheets/d/1aEIsnmsyiRml5CmwyauXOlUjthf7bTINgEy8qT4340g/edit?usp=sharing)|Splits values in one column and makes a copy of the row for each split value.
[SumAllCells](/scripts/SumAllCells.js)|Custom function|`=SUM_ALL_CELLS(cellToSum, {'SheetToExclude1",SheetToExclude2"})`|-| 
[WalletCoins](/scripts/WalletCoins.js)|API, Finanacial, Stock|-|-|Get coin name, shortname, rate and icon into sheets. [JSON data](https://wallet-api.staging.celsius.network/util/interest/rates)
[Ychart](/scripts/Ycharts.js)|API, Finanacial, Stock, Custom function|`=YCHARTS("AAPL")`|-|Gets raw data from the Ychart chart. [JSON data](https://ycharts.com/charts/fund_data.json?securities=id%3ATSLA%2Cinclude%3Atrue%2C%2C&calcs=id%3Ape_ratio%2Cinclude%3Atrue%2C%2C&correlations=&format=real&recessions=true&zoom=5&startDate=&endDate=&chartView=&splitType=single&scaleType=linear¬e=&title=&source=true&units=true"eLegend=true&partner="es=&legendOnChart=true&securitylistSecurityId=&displayTicker=true&ychartsLogo=&useEstimates=true&maxPoints=862)
[ZipChecker](/scripts/ZipChecker.js)|API|`=ZIPCHECKER(zip, countryISO2)`|-|Get information from a zipcode. [Example](http://api.zippopotam.us/us/99501)