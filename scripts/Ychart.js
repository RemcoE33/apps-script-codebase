/*
  Created by: RemcoE33
  https://github.com/RemcoE33/apps-script-codebase
*/

/**
* Returns YCHART information
*
* @param {"AAPL"} tickers - Input range, can be an array {"AAPL","TSLA"} or A1:B20.
* @return {array} range with security, date and value.
* @customfunction
*/
function YCHARTS(tickers){
  if(!Array.isArray(tickers)){
    tickers = [tickers]
  }
 
  tickers = tickers.flat();
 
  const output = [];
 
   tickers.forEach(tic => {
     const url = `https://ycharts.com/charts/fund_data.json?securities=id%3A${tic}%2Cinclude%3Atrue%2C%2C&calcs=id%3Ape_ratio%2Cinclude%3Atrue%2C%2C&correlations=&format=real&recessions=true&zoom=5&startDate=&endDate=&chartView=&splitType=single&scaleType=linear&note=&title=&source=true&units=true&quoteLegend=true&partner=&quotes=&legendOnChart=true&securitylistSecurityId=&displayTicker=true&ychartsLogo=&useEstimates=true&maxPoints=862`
 
    console.log(url);
 
     try{
       const response = UrlFetchApp.fetch(url);
       const object = JSON.parse(response.getContentText());
       const data = object.chart_data[0][0].raw_data;
 
       data.forEach(row => {
         output.push([tic, new Date(row[0]).toLocaleDateString(), row[1]]);
       })
 
     } catch (err){
       console.error(err);
     }
   })
 
   return output
}