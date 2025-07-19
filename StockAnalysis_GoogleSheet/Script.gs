// SHEET Details
var CONSOLIDATE_VIEW = "ConsolidatedView"
var SECTOR_ANALYSIS_SHEET = "SectorAnalysis";
var PRICE_MOVEMENT_SHEET = "PriceMovements_Core";
var MOVING_AVERAGE_SHEET = "MovingAverages_Core";
var OPEN_HIGH_LOW_SHEET = "OpenHighLow_Core";
var REFERENCE_DATA_SHEET = "ReferenceData";
var FUNDAMENTALS_SHEET = "Fundamentals_Core";

var PriceMovementSheetHeaders = 2;
var MovingAverageSheetHeaders = 3;
var FundamentalsSheetHeaders = 1;
var OpenHighLowSheetHeaders = 2;
var SectorAnalysisSheetHeaders = 2;

/** get current sheet */
function getActiveSpreadsheet(){
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** GETDATA function to be called from sheet to fetch the data */
function GetStocksData() {  

  // Do not run script on saturday and sunday
  // Do not run if weekend
  // Do not run if data already fetched today
  if(IsWeekend() || IsDataFetchedOnDate()){
    return;
  }

  var OPENPRICECOLUMN = 1;
  var HIGHPRICECOLUMN = 2;
  var LOWPRICECOLUMN = 3;
  var CLOSEPRICECOLUMN = 4;
  var VOLUMECOLUMN = 6;

  // Get Symbols
  var stockSymbols = GetSymbols();
  
  // Get Breakout Percentage
  var spreadSheet= getActiveSpreadsheet(); 
  var sheet = spreadSheet.getSheetByName(REFERENCE_DATA_SHEET);
  var breakoutPercentage = sheet.getRange('F2').getValue();

  // Iterate each symbol and get data
  for(var i=0; i<stockSymbols.length; i++){

    // after processing 50 stocks wait for 10 seconds
    if(i%50 == 0){
      Utilities.sleep(10000)
    }
    // get data in csv format
    var data = GetHistoricalData(stockSymbols[i]);      
    Logger.log(stockSymbols[i])
    var openPriceData = new Array();
    var closePriceData = new Array();
    var highPriceData = new Array();
    var lowPriceData = new Array();
    var volumeData = new Array();

    // get data from csv file in increasing order of date. 0th row is header
    for(var j=1; j<data.length; j++){
      
      // check if data is not NULL
      if(!isNaN(data[j][OPENPRICECOLUMN]) && !isNaN(data[j][CLOSEPRICECOLUMN]) && !isNaN(data[j][HIGHPRICECOLUMN]) && !isNaN(data[j][LOWPRICECOLUMN]) && !isNaN(data[j][VOLUMECOLUMN])){
        openPriceData[openPriceData.length] = parseFloat(data[j][OPENPRICECOLUMN]);
        closePriceData[closePriceData.length] = parseFloat(data[j][CLOSEPRICECOLUMN]);
        highPriceData[highPriceData.length] = parseFloat(data[j][HIGHPRICECOLUMN]);
        lowPriceData[lowPriceData.length] = parseFloat(data[j][LOWPRICECOLUMN]);
        volumeData[volumeData.length] = parseFloat(data[j][VOLUMECOLUMN]);
      }
    }

    // Get Today's Data
    var ltp = closePriceData[closePriceData.length - 1];
    var previousClose = closePriceData[closePriceData.length - 2];
    var difference = ltp - previousClose;

    // Get Days Trend Data
    var sevenDaysPriceTrend = AnalyseGeneralTrend(ltp, closePriceData[closePriceData.length - 7]);
    var thirtyDaysPriceTrend = AnalyseGeneralTrend(ltp, closePriceData[closePriceData.length - 30]);
    var hundredDaysPriceTrend = AnalyseGeneralTrend(ltp, closePriceData[closePriceData.length - 100]);
    var twoHundredDaysPriceTrend = AnalyseGeneralTrend(ltp, closePriceData[closePriceData.length - 200]);

    // Analyse 3 Days - Day by Day Trend
    var threeDaysPriceTrend = AnalyseDayByDayTrend(closePriceData, 2);
    var threeDaysVolumeTrend = AnalyseDayByDayTrend(volumeData, 2);
    
    // Calculate SMA
    var tenDaysSMAToday = CalculateSMA(closePriceData, 10);
    var tenDaysSMAYesterday = CalculateSMA(closePriceData, 10, 1);
    var tenDaysSMADayBeforeYesterday = CalculateSMA(closePriceData, 10, 2);
    var thirtyDaysSMAToday = CalculateSMA(closePriceData, 30);
    var thirtyDaysSMAYesterday = CalculateSMA(closePriceData, 30, 1);
    var thirtyDaysSMADayBeforeYesterday = CalculateSMA(closePriceData, 30, 2);
    var fiftyDaysSMAToday = CalculateSMA(closePriceData, 50);
    var fiftyDaysSMAYesterday = CalculateSMA(closePriceData, 50, 1);
    var fiftyDaysSMADayBeforeYesterday = CalculateSMA(closePriceData, 50, 2);
    var hundredDaysSMAToday = CalculateSMA(closePriceData, 100);
    var hundredDaysSMAYesterday = CalculateSMA(closePriceData, 100, 1);
    var hundredDaysSMADayBeforeYesterday = CalculateSMA(closePriceData, 100, 2);
    var twoHundredDaysSMAToday = CalculateSMA(closePriceData, 200);
    var twoHundredDaysSMAYesterday = CalculateSMA(closePriceData, 200, 1);
    var twoHundredDaysSMADayBeforeYesterday = CalculateSMA(closePriceData, 200, 2);

    // Calculate EMA
    var tenDaysEMA = CalculateEMA(closePriceData, 10);
    var thirtyDaysEMA = CalculateEMA(closePriceData, 30);
    var fiftyDaysEMA = CalculateEMA(closePriceData, 50);
    var hundredDaysEMA = CalculateEMA(closePriceData, 100);
    
    // Get Stock Fundamentals
    var stockFundamentals = GetStockFundamentals(stockSymbols[i])
    // Calculate Average True Range - Default 14 Days
    var averageTrueRange = CalculateATR(highPriceData, lowPriceData, closePriceData);
    stockFundamentals = stockFundamentals + "," + averageTrueRange;

    // Identify Range Breakout
    rbResult = IdentifyRangeBreakout(openPriceData, closePriceData, threeDaysPriceTrend)

    // Identify Last Candle
    var candlePatterns = IdentifyCandlestickPatterns(openPriceData, highPriceData, lowPriceData, closePriceData, numberOfDays=2);
 
    // Prepare Data for Sheets
    var yearHigh = Math.max.apply(null, highPriceData);
    var yearLow = Math.min.apply(null, lowPriceData);
    var highToGo = (yearHigh - ltp) / ltp * 100;
    var lowToGo = (ltp - yearLow) / ltp * 100;
    var tenVsFiftyGoldenCross = GoldenOrDeathCrossStrategy(tenDaysEMA[tenDaysEMA.length-1], fiftyDaysEMA[fiftyDaysEMA.length-1], tenDaysEMA[tenDaysEMA.length-2], fiftyDaysEMA[fiftyDaysEMA.length-2], tenDaysEMA[tenDaysEMA.length-3], fiftyDaysEMA[fiftyDaysEMA.length-3]);
    var fiftyVsHundredGoldenCross = GoldenOrDeathCrossStrategy(fiftyDaysEMA[fiftyDaysEMA.length-1], hundredDaysEMA[hundredDaysEMA.length-1], fiftyDaysEMA[fiftyDaysEMA.length-2], hundredDaysEMA[hundredDaysEMA.length-2], fiftyDaysEMA[fiftyDaysEMA.length-3], hundredDaysEMA[hundredDaysEMA.length-3]);
    var ltpVsThirtyEMA = PriceBreakoutDetection(ltp,previousClose,thirtyDaysEMA[thirtyDaysEMA.length-1],thirtyDaysEMA[thirtyDaysEMA.length-2],threeDaysPriceTrend=="Very Bullish",threeDaysPriceTrend=="Very Bearish",threeDaysPriceTrend=="Bullish",threeDaysPriceTrend=="Bearish", breakoutPercentage);
    var ltpVsFiftyEMA = PriceBreakoutDetection(ltp,previousClose,fiftyDaysEMA[fiftyDaysEMA.length-1],fiftyDaysEMA[fiftyDaysEMA.length-2],threeDaysPriceTrend=="Very Bullish",threeDaysPriceTrend=="Very Bearish",threeDaysPriceTrend=="Bullish",threeDaysPriceTrend=="Bearish", breakoutPercentage);
    var signalOnCandlePattern = AnalyseSignalOnCandlePattern(threeDaysPriceTrend, sevenDaysPriceTrend, candlePatterns[0]);
    var finalSignal = CalculateScore(threeDaysVolumeTrend, threeDaysPriceTrend, sevenDaysPriceTrend, thirtyDaysPriceTrend, hundredDaysPriceTrend, twoHundredDaysPriceTrend, tenVsFiftyGoldenCross, fiftyVsHundredGoldenCross, ltpVsFiftyEMA, rbResult, signalOnCandlePattern, highToGo, lowToGo);
   
    var priceMovementData = previousClose + "," + openPriceData[openPriceData.length - 1] + "," + highPriceData[highPriceData.length - 1] + "," + lowPriceData[lowPriceData.length - 1] 
    + "," + ltp + "," + volumeData[volumeData.length - 1] + "," + yearHigh + "," + yearLow + "," + difference + "," + (difference/previousClose)*100 + "," + threeDaysVolumeTrend + "," + threeDaysPriceTrend 
    + "," + sevenDaysPriceTrend + "," + thirtyDaysPriceTrend + "," + hundredDaysPriceTrend + "," + twoHundredDaysPriceTrend 
    + "," + openPriceData[openPriceData.length - 2].toFixed(2) + "," + highPriceData[highPriceData.length - 2].toFixed(2) + "," + lowPriceData[lowPriceData.length - 2].toFixed(2) + "," + rbResult + "," + candlePatterns[0] + "," + candlePatterns[1] + "," + signalOnCandlePattern + "," + tenVsFiftyGoldenCross + "," + fiftyVsHundredGoldenCross + "," + ltpVsThirtyEMA + "," + ltpVsFiftyEMA 
    + "," + finalSignal + "\r\n";
    
    var movingAverageData = tenDaysSMAToday + "," + tenDaysEMA[tenDaysEMA.length-1] + "," + thirtyDaysSMAToday + "," + thirtyDaysEMA[thirtyDaysEMA.length-1] + "," + fiftyDaysSMAToday 
    + "," + fiftyDaysEMA[fiftyDaysEMA.length-1] + "," + hundredDaysSMAToday + "," + hundredDaysEMA[hundredDaysEMA.length-1] + "," +twoHundredDaysSMAToday + "," + tenDaysSMAYesterday 
    + "," + tenDaysEMA[tenDaysEMA.length-2] + "," + thirtyDaysSMAYesterday + "," + thirtyDaysEMA[thirtyDaysEMA.length-2] + "," + fiftyDaysSMAYesterday 
    + "," + fiftyDaysEMA[fiftyDaysEMA.length-2] + "," + hundredDaysSMAYesterday + "," + hundredDaysEMA[hundredDaysEMA.length-2] + "," + twoHundredDaysSMAYesterday 
    + "," + tenDaysSMADayBeforeYesterday + "," + tenDaysEMA[tenDaysEMA.length-3] + "," + thirtyDaysSMADayBeforeYesterday + "," + thirtyDaysEMA[thirtyDaysEMA.length-3] 
    + "," + fiftyDaysSMADayBeforeYesterday + "," + fiftyDaysEMA[fiftyDaysEMA.length-3] + "," + hundredDaysSMADayBeforeYesterday + "," + hundredDaysEMA[hundredDaysEMA.length-3] 
    + ","  + twoHundredDaysSMADayBeforeYesterday + "\r\n";

    // Write to Sheet
    WriteToSheet(Utilities.parseCsv(priceMovementData, ","), PRICE_MOVEMENT_SHEET, PriceMovementSheetHeaders+1+i, 2, true);
    WriteToSheet(Utilities.parseCsv(movingAverageData, ","), MOVING_AVERAGE_SHEET, MovingAverageSheetHeaders+1+i, 2, true);
    WriteToSheet(Utilities.parseCsv(stockFundamentals, ","), FUNDAMENTALS_SHEET, FundamentalsSheetHeaders+1+i, 2, true);

    // Find Historical HighLow Session Details and set in sheet
    SetHistoricalHighLow(openPriceData, closePriceData, OPEN_HIGH_LOW_SHEET, OpenHighLowSheetHeaders+1+i, 2);   
  }

  RefreshLastUpdate();
}

/** Get Symbol Names from sheet */
function GetSymbols(){
  var spreadSheet= getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(REFERENCE_DATA_SHEET);
  var stockSymbolsAll = sheet.getRange('A2:A').getValues();
  
  var stockSymbols = stockSymbolsAll.filter(function (el) {
    return el != null && el != '';
  });
  return stockSymbols;
}

/** Get Stock Fundamentls from Web Scraping */
function GetStockFundamentals(symbol) {
  var url = "https://finance.yahoo.com/quote/"+symbol+".NS";
  var headers = {
     "user-agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
  };
  var options = {
     "headers" : headers
  };
  var response = UrlFetchApp.fetch(url, options); 
  var $ = Cheerio.load(response.getContentText());
  var peRatio = "";
  var epsRatio = "";
  $('span').filter(function (i, el) {
    var eps_pe = $('fin-streamer[data-field="trailingPE"]').text().toString().split(" ")
    // pe ratio
    if (eps_pe[0] != null )
      peRatio = eps_pe[0].replaceAll(",","");
    // eps ratio 
    if (eps_pe[1] != null )
      epsRatio = eps_pe[1].replaceAll(",","");   
  })

  return peRatio + "," + epsRatio;
}

/** Get Historical Data - From Yahoo Finance */
function GetHistoricalData(stockSymbol){
  var today= GetCurrentDate();
  var histDate = new Date(today.getTime()-365*(24*3600*1000));
  today.setHours(23,59,59);
  histDate.setHours(23,59,59);
  
  today= Math.round(today.getTime()/1000);
  histDate= Math.round(histDate.getTime()/1000);
  var url = "https://query1.finance.yahoo.com/v7/finance/download/"+stockSymbol+".NS?period1="+histDate+"&period2="+today+"&interval=1d&events=history&includeAdjustedClose=true";
  try{  
    var response = UrlFetchApp.fetch(url);
    var data = Utilities.parseCsv(response.getContentText(), ',');

    return data;
  }
  catch (e) {
    throw Error('Exception Occurred. URL:' + url + ', Exception Occurred. ExceptionDetails:'+ e);
  }
}

/** Get All Data */
function GetAllData(){

  // Get Latest Stocks Data
  GetStocksData();
}

/** Set High low Session Counts from historical Data */
function SetHistoricalHighLow(openData, closeData, sheetName, rowIndex, columnIndex){
  // get today's date    
  var today = GetCurrentDate();

  // convert complete result set in CSV format to be used in sheet
  var csvFile = "";
  
      
  var openHighCloseHighCount = 0;
  var openHighCloseLowCount = 0;
  var openLowCloseHighCount = 0;
  var openLowCloseLowCount = 0;    
  var sessionCount = 0

  // iterate each row and compare data
  for(var j=1; j<closeData.length; j++)
  {      
    var open = openData[j];
    var close = closeData[j];
    var prevClose = closeData[j-1];     
    
    var openHigh = open > prevClose;
    var closeHigh = close > open;

    if(openHigh && closeHigh){
      openHighCloseHighCount++;
    } 
    else if(openHigh && !closeHigh){
      openHighCloseLowCount++;
    } 
    else if(!openHigh && closeHigh){
      openLowCloseHighCount++;
    }
    else if(!openHigh && !closeHigh){
      openLowCloseLowCount++;
    }

    sessionCount++;   
  }

    //append data in csv file
    csvFile += (openHighCloseHighCount+ "," + openHighCloseLowCount + "," 
    + openLowCloseHighCount + "," + openLowCloseLowCount + "," + sessionCount + "\r\n");

  var spreadSheet= getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(sheetName);

  if(csvFile == ""){
    sheet.getRange('B1:C1').setValue(stockSymbols+"Failed:  " + GetFormattedDate(today));
  }
  else{
    // parse result
    var result = Utilities.parseCsv(csvFile, ",");
    
    sheet.getRange('B1:C1').setValue("Success:  " + GetFormattedDate(today));
    sheet.getRange(rowIndex,columnIndex,result.length,result[0].length).setValues(result);
  }
}

/** Analyse General Price Trend */
function AnalyseGeneralTrend(ltp, historicalPrice){

  var result = "Neutral";
  if(ltp > historicalPrice){
    result = "Bullish";
  }
  else if(ltp < historicalPrice){
    result = "Bearish";
  } 

  return result;
}

/** data should be passed in increasing order of date in array */
function AnalyseDayByDayTrend(data, days, minRequiredDifferencePercentage=0.01){
  
  if(data == null || !Array.isArray(data) || data.length < days){
    return ERROR;
  }

  var resultCount = 0
  var startingIndex = data.length - 1;
  for(var i=startingIndex; i>startingIndex-days;i--){
    if(isNaN(parseFloat(data[i])) || isNaN(parseFloat(data[i-1])))
    {
      return ERROR
    }
    else if(parseFloat(data[i]) > parseFloat(data[i-1])){
      resultCount++;
    }
  }

  var result = "";
  var differenceInPerc = Math.abs(((parseFloat(data[startingIndex]) - parseFloat(data[startingIndex-days])))/parseFloat(data[startingIndex]))
  
  if(resultCount == days && differenceInPerc >= minRequiredDifferencePercentage){
    result = "Very Bullish";
  }
  else if(resultCount == 0 && differenceInPerc >= minRequiredDifferencePercentage){
    result = "Very Bearish";
  }
  else if(parseFloat(data[startingIndex]) > parseFloat(data[startingIndex-days])){
    result = "Bullish";
  }
  else if(parseFloat(data[startingIndex]) <= parseFloat(data[startingIndex-days])){
    result = "Bearish";
  }
  else{
    result = "Neutral"
  }
    
  return result;
}

/** Find Break out for prices provided */
function PriceBreakoutDetection(priceToCompareCurrent, priceToComparePrevious, priceInComparisonToCurrent, priceInComparisonToPrevious, threeDayBullish, threeDayBearish, normalBullish, normalBearish, breakoutPercentage){
  // calulate required percentage of EMA for analysis of breakout
  var val = (priceInComparisonToCurrent * breakoutPercentage)/100;
  var diff = (((priceToCompareCurrent - priceInComparisonToCurrent)/priceInComparisonToCurrent)*100).toFixed(2);
  if(isNaN(priceToCompareCurrent) || isNaN(diff)){
    return ERROR;
  }
  var result = "";
  
  if(priceToComparePrevious <= priceInComparisonToPrevious && priceToCompareCurrent > priceInComparisonToCurrent){
    if(threeDayBullish){
      result = "Very Bullish : BREAKOUT";
    }
    else if(normalBullish){
      result = "Bullish : BREAKOUT";
    }
  }  
  else if(priceToComparePrevious >= priceInComparisonToPrevious && priceToCompareCurrent < priceInComparisonToCurrent){
    if(threeDayBearish){
      result = "Very Bearish : BREAKOUT";
    }
    else if(normalBearish){
      result = "Bearish : BREAKOUT";
    }
  }
  else if(priceToComparePrevious <= priceInComparisonToPrevious && priceToCompareCurrent < priceInComparisonToCurrent && val >= (priceInComparisonToCurrent - priceToCompareCurrent)){
    if(threeDayBullish){      
      result = "Very Bullish : NEAR BREAKOUT";
    }
    else if(normalBullish){
      result = "Bullish : NEAR BREAKOUT";
    }
  } 
  else if(priceToComparePrevious >= priceInComparisonToPrevious && priceToCompareCurrent > priceInComparisonToCurrent && val >= (priceToCompareCurrent - priceInComparisonToCurrent)){
    if(threeDayBearish){      
      result = "Very Bearish : NEAR BREAKOUT";
    }
    else if(normalBearish){
      result = "Bearish : NEAR BREAKOUT";
    }
  }

  // if nothing is concluded from above logics
  if(result == ""){
    if(threeDayBullish || threeDayBearish || normalBullish || normalBearish){
      result = (diff>0 ? (diff+"% above") : ((diff*-1)+"% below")) +" EMA";
    }
    else{
      result = "Neutral";
    }
  } 

  return result;
}

/** Calcualte Simple Moving Average  - data is passed in increasing order of date */
function CalculateSMA(daysClosingPrices, numberOfDays, numberOfDaysBefore = 0 ){
  
  var sum = 0;
  var startingIndex = daysClosingPrices.length - 1 - numberOfDaysBefore;
  for (var i = startingIndex; i > startingIndex - numberOfDays; i--){     
    sum += parseFloat(daysClosingPrices[i]);
  }

  return sum/numberOfDays;
}

/** Calcualte Exponential Moving Average  - data is passed in increasing order of date */
function CalculateEMA(daysClosingPrices, numberOfDays){

  var weightageMultiplier = (2 / (numberOfDays + 1));
  
  var emaData = new Array();
  var yesterdaysEMA = 0;
  // Calculate Simple Moving Average of initial number of Days
  for (var i = 0; i < numberOfDays; i++)
  {     
    yesterdaysEMA = yesterdaysEMA + parseFloat(daysClosingPrices[i]);
  }
  yesterdaysEMA = yesterdaysEMA / numberOfDays;
  emaData[0] = yesterdaysEMA;
  
  for (var i = numberOfDays; i < daysClosingPrices.length; i++)
  {     
    var ema = parseFloat(((daysClosingPrices[i] - yesterdaysEMA) * weightageMultiplier) + yesterdaysEMA); 
    emaData[emaData.length] = ema
    yesterdaysEMA = ema;
  }

  return emaData;
}

/** Add LTP in data based on market open date time, to display realtime results */
function VerifyAndAddLatestData(data, marketOpenedToday, ltp, dataIsInDecreaingOrderOfDate, maintainDataLength = true){  

    if(isNaN(ltp)){
      return ERROR;
    }

    if(IsRealTimeShortTermDataAdded(marketOpenedToday)){
      if(dataIsInDecreaingOrderOfDate){
        var finalData = new Array();
        finalData[0] = ltp;
        if(maintainDataLength){
          data.splice(data.length-1,1);
        }
        data = finalData.concat(data);         
      }
      else{
        if(maintainDataLength){
          data.splice(0,1);
        }
        data.push(ltp);  
      }
    }

    return data;
}

/** check if real time data should be added or not */
function IsRealTimeShortTermDataAdded(marketOpenedToday){

  // current date time
  var today = GetCurrentDate(); 

  // market starts at 9:15 in India    
  var addRealTimeData = !IsWeekend() && marketOpenedToday 
  && (Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(),'HH') > 9 
  || (Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(),'HH') == 9 
  && Utilities.formatDate(today, SpreadsheetApp.getActive().getSpreadsheetTimeZone(),'mm') > 14));
  
  return addRealTimeData;
}

/** check for weekend */
function IsWeekend(date){
  if (date == null){    
    date = GetCurrentDate();
  }
  // 6=Saturday, 0=Sunday
  return (date.getDay() == 6 || date.getDay() == 0)
}

/** Check if data has been updated today */
function IsDataFetchedOnDate(date){
  if (date == null){    
    date = GetCurrentDate();
  }

  var dataFetched = true;
  var spreadSheet = getActiveSpreadsheet();
  var lastUpdate = spreadSheet.getSheetByName(REFERENCE_DATA_SHEET).getRange('F4').getValue();
  // get moment js file
  eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.18.1/moment.min.js').getContentText());
  lastUpdate = moment(lastUpdate.split(" ")[0], "DD/MMM/YYYY").toDate();
  
  if(date.getDate() != lastUpdate.getDate() || 
        date.getMonth() != lastUpdate.getMonth()
          || date.getFullYear() != lastUpdate.getFullYear()){
      dataFetched = false;
  }

  return dataFetched;
}

/** calculate current date time - in Indian TimeZone */
function GetCurrentDate(){
  // current date time - in IST Time Zone
  var dateString = Utilities.formatDate(new Date(), "GMT+5:30", 'MMMM dd, yyyy hh:mm:ss Z');
  var today = new Date(dateString);

  return today;
}

/** calculate formatted date - in Indian TimeZone */
function GetFormattedDate(date_to_format){
  // format date time - in IST Time Zone
  var formatted_date = Utilities.formatDate(date_to_format, SpreadsheetApp.getActive().getSpreadsheetTimeZone(),'dd/MMM/yyyy hh:mm:ss')
  formatted_date += " PM IST"
  
  return formatted_date;
}

/** Calculate score for the shares - for trend buy/sell signal */
function CalculateScore(volumeTwoDaysTrend, priceTwoDaysTrend, sevenDays, oneMonth, sixMonths, twelveMonths, ema10DaysVS50Days, ema50DaysVS100Days, ltpVS50daysEMA, breakout, lastCandleSignal, highToGo, lowToGo){
    
  if(volumeTwoDaysTrend.includes("#ERROR!") || priceTwoDaysTrend.includes("#ERROR!") 
  || ema10DaysVS50Days.includes("#ERROR!")
  || ltpVS50daysEMA.includes("#ERROR!") || ema50DaysVS100Days.includes("#ERROR!") || breakout.includes("#ERROR!")){
    return ERROR;
  }  
  
  // total score should be greater than below minimum scores
  var normalBuySignalScore = 20;
  var strongBuySignalScore = 25;
  var normalSellSignalScore = 20;
  var strongSellSignalScore = 25;
  // LTP to High/Low difference should be greater than 5%
  var minHighLowtoGoPercentage = 5;
  var highToGoValid = highToGo > minHighLowtoGoPercentage 
  var lowToGoValid = lowToGo > minHighLowtoGoPercentage 

  var result = "Neutral";
  // set initial score as 0
  var score = 0;
  
  // Buy score calculation.
  if(highToGoValid && priceTwoDaysTrend.includes("Bullish")){

    // Reference Data
    var ScoreTable = {
      "Volume-Very Bullish" : 10,
      "Volume-Bullish" : 5,
      "7Days-Bullish" : 5,
      "1Month-Bullish" : 3,
      "6Months-Bullish" : 2,
      "12Months-Bullish" : 1,
      "Golden Cross Continued" : 5,
      "Golden Cross" : 3,
      "above" : 1,
      "Very Bullish : BREAKOUT" : 5,
      "Very Bullish : NEAR BREAKOUT" : 5,
      "Bullish : BREAKOUT" : 3,
      "Bullish : NEAR BREAKOUT" : 3,
      "Bullish : Hammer" : 5,
      "Bullish : Inverted Hammer" : 5
    };

    if(volumeTwoDaysTrend.includes("Bullish")){
      volumeTwoDaysTrend = "Volume-" + volumeTwoDaysTrend;
    }
    if(ema10DaysVS50Days.includes("above")){
      ema10DaysVS50Days = "above";
    }
    if(ema50DaysVS100Days.includes("above")){
      ema50DaysVS100Days = "above";
    }
    if(sevenDays.includes("Bullish")){
      sevenDays = "7Days-" + sevenDays;
    }
    if(oneMonth.includes("Bullish")){
      oneMonth = "1Month-" + oneMonth;
    }
    if(sixMonths.includes("Bullish")){
      sixMonths = "6Months-" + sixMonths;
    }
    if(twelveMonths.includes("Bullish")){
      twelveMonths = "12Months-" + twelveMonths;
    }
    if(ltpVS50daysEMA.includes("above")){
      ltpVS50daysEMA = "above";
    }

    score += (isNaN(ScoreTable[volumeTwoDaysTrend]) ? 0 : ScoreTable[volumeTwoDaysTrend]);
    score += (isNaN(ScoreTable[sevenDays]) ? 0 : ScoreTable[sevenDays]);
    score += (isNaN(ScoreTable[oneMonth]) ? 0 : ScoreTable[oneMonth]);
    score += (isNaN(ScoreTable[sixMonths]) ? 0 : ScoreTable[sixMonths]);
    score += (isNaN(ScoreTable[twelveMonths]) ? 0 : ScoreTable[twelveMonths]);
    score += (isNaN(ScoreTable[ema10DaysVS50Days]) ? 0 : ScoreTable[ema10DaysVS50Days]);
    score += (isNaN(ScoreTable[ema50DaysVS100Days]) ? 0 : ScoreTable[ema50DaysVS100Days]);
    score += (isNaN(ScoreTable[ltpVS50daysEMA]) ? 0 : ScoreTable[ltpVS50daysEMA]);
    score += (isNaN(ScoreTable[breakout]) ? 0 : ScoreTable[breakout]);
    
    // verify score
    if(score > normalBuySignalScore){
      result = "BUY";
      // add score of last candle pattern
      score += (isNaN(ScoreTable[lastCandleSignal]) ? 0 : ScoreTable[lastCandleSignal]);
      if(score > strongBuySignalScore){
        result = "STRONG BUY";
      }
    }
  }
  else if(lowToGoValid && priceTwoDaysTrend.includes("Bearish")){
    
    // include 2 points if Volume is Very Bullish
    if(volumeTwoDaysTrend.includes("Very Bullish")){
      score+=2;
    }

    // Reference Data
    var ScoreTable = {
      "Volume-Very Bearish" : 10,
      "Volume-Bearish" : 5,
      "7Days-Bearish" : 5,
      "1Month-Bearish" : 3,
      "6Months-Bearish" : 2,
      "12Months-Bearish" : 1,
      "Death Cross Continued" : 5,
      "Death Cross" : 3,
      "below" : 1,
      "Very Bearish : BREAKOUT" : 5,
      "Very Bearish : NEAR BREAKOUT" : 5,
      "Bearish : BREAKOUT" : 3,
      "Bearish : NEAR BREAKOUT" : 3,
      "Bearish : Hanging Man" : 5,
      "Bearish : Shooting Star" : 5
    };

    if(volumeTwoDaysTrend.includes("Bearish")){
      volumeTwoDaysTrend = "Volume-" + volumeTwoDaysTrend;
    }
    if(ema10DaysVS50Days.includes("below")){
      ema10DaysVS50Days = "below";
    }
    if(ema50DaysVS100Days.includes("below")){
      ema50DaysVS100Days = "below";
    }
    if(sevenDays.includes("Bearish")){
      sevenDays = "7Days-" + sevenDays;
    }
    if(oneMonth.includes("Bearish")){
      oneMonth = "1Month-" + oneMonth;
    }
    if(sixMonths.includes("Bearish")){
      sixMonths = "6Months-" + sixMonths;
    }
    if(twelveMonths.includes("Bearish")){
      twelveMonths = "12Months-" + twelveMonths;
    }
    if(ltpVS50daysEMA.includes("below")){
      ltpVS50daysEMA = "below";
    }

    score += (isNaN(ScoreTable[volumeTwoDaysTrend]) ? 0 : ScoreTable[volumeTwoDaysTrend]);
    score += (isNaN(ScoreTable[sevenDays]) ? 0 : ScoreTable[sevenDays]);
    score += (isNaN(ScoreTable[oneMonth]) ? 0 : ScoreTable[oneMonth]);
    score += (isNaN(ScoreTable[sixMonths]) ? 0 : ScoreTable[sixMonths]);
    score += (isNaN(ScoreTable[twelveMonths]) ? 0 : ScoreTable[twelveMonths]);
    score += (isNaN(ScoreTable[ema10DaysVS50Days]) ? 0 : ScoreTable[ema10DaysVS50Days]);
    score += (isNaN(ScoreTable[ema50DaysVS100Days]) ? 0 : ScoreTable[ema50DaysVS100Days]);
    score += (isNaN(ScoreTable[ltpVS50daysEMA]) ? 0 : ScoreTable[ltpVS50daysEMA]);
    score += (isNaN(ScoreTable[breakout]) ? 0 : ScoreTable[breakout]);
    
    // verify score
    if(score > normalSellSignalScore){
      result = "SELL";
      // add score of last candle pattern
      score += (isNaN(ScoreTable[lastCandleSignal]) ? 0 : ScoreTable[lastCandleSignal]);
      if(score > strongSellSignalScore){
        result = "STRONG SELL";
      }
    }
  }  
  
  return result;
}

/** Identify Range Breakout */
function IdentifyRangeBreakout(openData, closeData, threeDaysPriceTrend, rangeIdentificationDays=20, rangeMovementDays=10, minMovementPercentage=0.0025, nearBreakoutRange=0.01){
  if(!Array.isArray(openData) || !Array.isArray(closeData) || openData.length != closeData.length){
    return ERROR;
  }

  var todayClosePrice = closeData[closeData.length-1]
  
  // IDENTIFY RANGE
  var startingIndex = openData.length - 2 - rangeMovementDays // leave rangeMovementDays for range identification
  var last_days_max_upper_range = openData[startingIndex]
  var last_days_max_lower_range = closeData[startingIndex]

  for(var i=startingIndex; i>startingIndex-rangeIdentificationDays+1;i--){
    // find upper range
    
    if(openData[i-1] > last_days_max_upper_range){
      last_days_max_upper_range = openData[i-1]
    }
    if(closeData[i-1] > last_days_max_upper_range){
      last_days_max_upper_range = closeData[i-1]
    }

    // find lower range
    if(openData[i-1] < last_days_max_lower_range){
      last_days_max_lower_range = openData[i-1]
    }
    if(closeData[i-1] < last_days_max_lower_range){
      last_days_max_lower_range = closeData[i-1]
    }
  }

  // Identify breakout with threshhold
  resultCount=0
  startingIndex = openData.length - 2 
  threshold = Math.abs(((last_days_max_upper_range+last_days_max_lower_range)/2)*minMovementPercentage) // min percentage of average range
  last_days_max_lower_range = last_days_max_lower_range - threshold
  last_days_max_upper_range = last_days_max_upper_range + threshold

  for(var i=startingIndex; i>startingIndex-rangeMovementDays;i--){
    if(openData[i] >= last_days_max_lower_range && openData[i] <= last_days_max_upper_range && closeData[i] >= last_days_max_lower_range && closeData[i] <= last_days_max_upper_range){
      resultCount++
    }
  }

  result = "Neutral"
  if(resultCount == rangeMovementDays){ // Trading in Range
    bullBreakoutRemainingPerc = ((last_days_max_upper_range - todayClosePrice)/last_days_max_upper_range)
    bearBreakoutRemainingPerc = ((todayClosePrice - last_days_max_lower_range)/last_days_max_lower_range)

    result = "Trading in Range"
    if((todayClosePrice > last_days_max_upper_range) || todayClosePrice < last_days_max_lower_range){
      result = threeDaysPriceTrend + " : BREAKOUT"
    }
    else if ((bullBreakoutRemainingPerc >= 0 && bullBreakoutRemainingPerc <= nearBreakoutRange) || (bearBreakoutRemainingPerc >= 0 && bearBreakoutRemainingPerc <= nearBreakoutRange)){
      result = threeDaysPriceTrend + " : NEAR BREAKOUT"
    }
  }
  
  return result;
}

/** Calculate Average True Range */
function CalculateATR(highData, lowData, closeData, numberOfDays = 14, resultOfDaysBeforeToday = 0){
  if(!Array.isArray(highData) || !Array.isArray(lowData) || !Array.isArray(closeData) || highData.length != lowData.length || highData.length != closeData.length){
    return ERROR;
  }

  var trueRange = CalculateTR(highData, lowData, closeData);
  
  var averageTrueRange = new Array();

  // calculate average TR for number of Days
  var averageTR = parseFloat(0);
  for (var i = 0; i < numberOfDays; i++){     
    averageTR += parseFloat(trueRange[i]);
    averageTrueRange[i] = 0;
  }
  averageTR = averageTR / numberOfDays;
  
  var weightageMultiplier = 1/numberOfDays;

  averageTrueRange[averageTrueRange.length] = averageTR;

  // Calculate ATR
  for (var i = averageTrueRange.length; i < trueRange.length; i++){     
    averageTrueRange[averageTrueRange.length] = (weightageMultiplier*trueRange[i]+(1-weightageMultiplier)*averageTrueRange[i-1]);
  }

  // which days of result is required
  var index = (averageTrueRange.length - 1) - resultOfDaysBeforeToday;
  
  return averageTrueRange[index];
}

/** Calculate True Range */
function CalculateTR(highData, lowData, closeData){

  var trueRange = new Array();
  // initialise 0th index with 0
  trueRange[0] = 0;

  for(var i=1;i<highData.length;i++){
    var high = highData[i];
    var low = lowData[i];
    var prevClose = closeData[i-1];
    trueRange[trueRange.length] = Math.max(high-low, high-prevClose, prevClose-low);
  }

  return trueRange;
}

/** Write Data to Sheet */
function WriteToSheet(data, sheetName, rowIndex, columnIndex, IsCSV = false){
  var spreadSheet= getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(sheetName);

  // set data in sheet
  if(IsCSV){
    sheet.getRange(rowIndex,columnIndex,data.length,data[0].length).setValues(data);
  }
  else{
    sheet.getRange(rowIndex, columnIndex).setValue(data);
  }
}
/////////////////// FORMULAE for Consolidated Sheet /////////////////////////

/** Evaluate Golder or Death Crossover */
function GoldenOrDeathCrossStrategy(lessDaysSMAToday, moreDaysSMAToday, lessDaysSMAYest, moreDaysSMAYest,lessDaysSMADayBeforeYest, moreDaysSMADayBeforeYest){
 
  var diff = (((lessDaysSMAToday - moreDaysSMAToday)/moreDaysSMAToday)*100).toFixed(2);
  if(isNaN(diff)){
    return ERROR;
  }

  var todayAbove = lessDaysSMAToday > moreDaysSMAToday;
  var yesterdayAbove = lessDaysSMAYest > moreDaysSMAYest;
  var dayBeforeYesterday = lessDaysSMADayBeforeYest > moreDaysSMADayBeforeYest;

  var result = "Neutral";

  if(todayAbove && !yesterdayAbove && !dayBeforeYesterday){
    result = "Golden Cross";
  }
  else if(!todayAbove && yesterdayAbove && dayBeforeYesterday){
    result = "Death Cross";
  }
  else if(todayAbove && yesterdayAbove && !dayBeforeYesterday){
    result = "Golden Cross Continued";
  }
  else if(!todayAbove && !yesterdayAbove && dayBeforeYesterday){
    result = "Death Cross Continued";
  }
  else if(!todayAbove && yesterdayAbove && !dayBeforeYesterday){
    result = "Golden Cross Reversed";
  }
  else if(todayAbove && !yesterdayAbove && dayBeforeYesterday){
    result = "Death Cross Reversed";
  }
  else if(todayAbove){
    result = diff+"% above";
  }
  else if(!todayAbove){
    result = (diff*-1)+"% below";
  }

  return result;
}

/** Calculate Difference */
function CalculateTargetDifference(lastTwoDaysHighValues, lastTwoDaysLowValues, atrValue){
  var averageMovementOfTwoDays = 0
  for(var i=0;i<lastTwoDaysHighValues.length;i++){
    averageMovementOfTwoDays += parseFloat(lastTwoDaysHighValues[i]) - parseFloat(lastTwoDaysLowValues[i]);
  }
  averageMovementOfTwoDays = averageMovementOfTwoDays/lastTwoDaysHighValues.length;
  
  // if ATR is higher than average of 2 days then take average movement of 2 days as target, otherwise take ATR as target
  //return averageMovementOfTwoDays < atrValue ? averageMovementOfTwoDays : atrValue;
  return averageMovementOfTwoDays;
}

/** Analyse signal based on Candle Pattern */
function AnalyseSignalOnCandlePattern(lastTwoDaysTrend, lastSevenDaysTrend, candleType){
  var result = "Neutral";
  if ((lastTwoDaysTrend.includes("Very Bearish") || lastSevenDaysTrend.includes("Bearish")) && (candleType == "Hammer" || candleType == "Inverted Hammer" || candleType == "Dragonfly Doji")){
    result = "Bullish : " + candleType;
  }
  else if ((lastTwoDaysTrend.includes("Very Bullish") || lastSevenDaysTrend.includes("Bullish")) && (candleType == "Hanging Man" || candleType == "Shooting Star" || candleType == "Gravestone Doji")){
    result = "Bearish : " + candleType;
  }
  else if(candleType == "Indecisive Doji"){
    result = "Indecisive Doji"
  }

  return result;
}

/** Get Data of Consolidated View Sheet and Send in Mail */
function SendDataInMail() {
  
  GetAllData()

  // Wait for 5 seconds
  var waitTimeInSeconds = 5000;
  var start = GetCurrentDate().getTime();
   var end = start;
   while(end < start + waitTimeInSeconds) {
     end = GetCurrentDate().getTime();
  }
  
  // get spread sheet data
  var spreadSheet = getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(CONSOLIDATE_VIEW).getDataRange();
  var lastUpdate = spreadSheet.getSheetByName(REFERENCE_DATA_SHEET).getRange('F4').getValue();
  
  var background = sheet.getBackgrounds();
  var val = sheet.getDisplayValues();
  var fontColor = sheet.getFontColors();
  var fontStyles = sheet.getFontStyles();
  var fontWeight = sheet.getFontWeights();
  var fontSize = sheet.getFontSizes();
  var sheet_data = [val,background,fontColor,fontStyles,fontWeight,fontSize];

  // format data
  var data = sheet_data[0];
  var background = sheet_data[1];
  var fontColor = sheet_data[2];
  var fontStyles = sheet_data[3];
  var fontWeight = sheet_data[4];
  var fontSize = sheet_data[5];
  var html = "";
  var dataNotLoadedCount = 0;
  // check if data is available with Buy/Sell Call
  var htmlRows = ""

  // columns
  var startingColumn = 14
  var endColumn = 20
  
  for (var i = 3; i < data.length; i++) {
    if(data[i].includes("BUY") || data[i].includes("STRONG BUY") || data[i].includes("SELL") || data[i].includes("STRONG SELL")){
        htmlRows += "<tr>"
        // add symbol
        htmlRows += "<td style='height:20px;background:DodgerBlue;color:" + fontColor[i][0] + ";font-style:" + fontStyles[i][0] + ";font-weight:" + fontWeight[i][0] + ";font-size:" + (fontSize[i][0] + 6) + "px;'>" + data[i][0] + "</td>";

        // iterate data for buy sell signal
        for (var j = startingColumn; j <=endColumn; j++) {
            htmlRows += "<td style='width:70%;height:20px;background:" + background[i][j] + ";color:" + fontColor[i][j] + ";font-style:" + fontStyles[i][j] + ";font-weight:" + fontWeight[i][j] + ";font-size:" + (fontSize[i][j] + 6) + "px;'>" + data[i][j] + "</td>";
        }

        //add Upper Range / Lower Range %
        /*var upperRangePercentage = (((parseFloat(data[i][17].replaceAll(",","")) - parseFloat(data[i][16].replaceAll(",","")))/parseFloat(data[i][15].replaceAll(",","")))*100).toFixed(2);
        var lowerRangePercentage = (((parseFloat(data[i][15].replaceAll(",","")) - parseFloat(data[i][19].replaceAll(",","")))/parseFloat(data[i][15].replaceAll(",","")))*100).toFixed(2);
        htmlRows += "<td style='height:20px;background:" + background[i][15] + ";color:" + fontColor[i][15] + ";font-style:" + fontStyles[i][16] + ";font-weight:" + fontWeight[i][15] + ";font-size:" + (fontSize[i][15] + 6) + "px;'>" + upperRangePercentage + "/" + lowerRangePercentage + "</td>";*/
        htmlRows += "</tr>";
    }
    else if(data[i].includes("#ERROR!") || data[i].includes("Loading...")){
      dataNotLoadedCount++;
    }
  }
  
  // if data is not loaded for more than 1 share then send below message
  if(dataNotLoadedCount > 1){
    html = "<br/><h1 style='background-color:Tomato;color:White'>Data Not Loaded Properly.</h1>"
  }
  else if(htmlRows == ""){// if Buy/Sell calls are not generated then send below message
    html = "<br/><h1 style='background-color:Gray;color:White'>No Buy/Sell calls for today.</h1>"
  }
  else{
    html = "<!DOCTYPE html><html><head><style>table { font-family: arial, sans-serif; border-collapse: collapse; width: 40%; } td, th { border: 1px solid; background-color:#1569C7; color:White; text-align: left; padding: 8px; } tr:nth-child(even) { background-color: #dddddd; }</style></head><body><table><tr style='font-size: 15px;'><th>Symbol</th><th>Signal</th><th>Last Price</th><th>PE Ratio</th><th>PIVOT</th><th>S1</th><th>R1</th><th>52 Week High</th></tr>" + htmlRows + "</table>"
  }

  html = html + "<br/><hr/><i>Last Data Refresh - " + lastUpdate + "</i>";
  
  // send mail
  MailApp.sendEmail({
      bcc: "andaaz01@gmail.com,manasiamrale@gmail.com",
      subject: spreadSheet.getName() + " - Consolidated View",
      htmlBody: html
  })

  // throw error to break script if data is not loaded properly.
  if(dataNotLoadedCount > 0){
    throw EvalError("Data not loaded.");
  }
}

/**  Refresh Last Update Header on Consolidated View Sheet to refresh calculations*/
function RefreshLastUpdate() {
  var spreadSheet = getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(REFERENCE_DATA_SHEET);
  sheet.getRange('F4').setValue(GetFormattedDate(GetCurrentDate()));
}

/** price should be multiple of 5 */
function PriceInMultipleOf(price, multipleOf){
  return price-(price%multipleOf);
}

/** Analyse Sector */
function AnalyseSectorData(stockList, closePrice, prevClosePrice){
  
  if(!Array.isArray(stockList) || !Array.isArray(closePrice) || !Array.isArray(prevClosePrice) || stockList.length != closePrice.length || stockList.length != prevClosePrice.length){
    if(stockList.length > 0 && !isNaN(closePrice) && !isNaN(prevClosePrice)){
      stockList = [stockList];
      closePrice = [closePrice];
      prevClosePrice = [prevClosePrice] 
    }
    else{
      return ERROR;
    }
  }
  var result = new Array();
  var totalStocks = stockList.length;
  var upStocks = 0;
  var downStocks = 0;
  var moveList = new Array();

  // Calculate Movements
  for(var i = 0; i < totalStocks; i++){
    move = closePrice[i] - prevClosePrice[i];
    moveList[moveList.length] = move/prevClosePrice[i];
    if (move > 0){
      upStocks++; 
    }
    else if (move < 0){
      downStocks++;
    }
  }

  // Identify Trend
  trend = "Neutral";
  isBullishMove = false;
  isBearishMove = false;
  if(upStocks/totalStocks > 0.9){
    trend = "Very Bullish";
    isBullishMove = true
  }
  else if(downStocks/totalStocks > 0.9){
    trend = "Very Bearish";
    isBearishMove = true;
  }
  else if(upStocks/totalStocks > 0.75){
    trend = "Bullish";
    isBullishMove = true;
  }
  else if(downStocks/totalStocks > 0.75){
    trend = "Bearish";
    isBearishMove = true;
  }
  
  // Stock with Max Move
  stockWithMaxMove = "";
  if (isBullishMove){
    stockWithMaxMove = stockList[moveList.indexOf(Math.max.apply(null, moveList))];
  }
  else if (isBearishMove){
    stockWithMaxMove = stockList[moveList.indexOf(Math.min.apply(null, moveList))];
  }

  // Prepare Result
  result[result.length] = totalStocks;
  result[result.length] = upStocks;
  result[result.length] = downStocks;
  result[result.length] = trend;
  result[result.length] = stockWithMaxMove;

  return result;
}

/** Identify Candlesticks Pattern*/
function IdentifyCandlestickPatterns(openPriceData, highPriceData, lowPriceData, closePriceData, numberOfDays=1){
  var GREEN_CANDLE = "GREEN";
  var RED_CANDLE = "RED";
  var comparisonRatio = 3 // wick should be atleast 3 times of body
  var finalResult = new Array();

  var startingIndex = closePriceData.length - numberOfDays;
  for (var index=startingIndex; index<closePriceData.length; index++){
    var open = openPriceData[index];
    var high = highPriceData[index];
    var low = lowPriceData[index];
    var close = closePriceData[index];

    var previousCandleType = closePriceData[index-1] > openPriceData[index-1] ? GREEN_CANDLE : RED_CANDLE;
    var body = Math.abs(open-close);
    var lowerWick = Math.min(open,close) - low;
    var upperWick = high - Math.max(open,close);
   
    //identify DOJI candle till last day's data
    var totalSize = 0;
    for(var i=0;i<index;i++){
      totalSize+=Math.abs(closePriceData[i]-openPriceData[i]);
    }
    var averageCandleBodySize = totalSize/openPriceData.length;
    var doji = Math.abs(close - open) <= (0.05 * averageCandleBodySize);
    var dragonflyDoji = doji && lowerWick > (upperWick*comparisonRatio); // lower wick is more than 3 times of upper wick
    var gravestoneDoji = doji && upperWick > (lowerWick*comparisonRatio); // upper wick is more than 3 times of lower wick
    var indecisiveDoji = doji && !dragonflyDoji && !gravestoneDoji

    // BULLISH REVERSAL CANDLE INDICATORS    
    // check for hammer
    var hammer = !doji && lowerWick > upperWick && lowerWick > (comparisonRatio*body) && previousCandleType == RED_CANDLE;
    //check for inverted hammer
    var invertedHammer = !doji && upperWick > lowerWick && upperWick > (comparisonRatio*body) && previousCandleType == RED_CANDLE;

    // BEARISH REVERSAL CANDLE INDICATORS
    // check for hanging man
    var hangingMan = !doji && lowerWick > upperWick && lowerWick > (comparisonRatio*body) && previousCandleType == GREEN_CANDLE;   
    // check for shooting star
    var shootingStar = ! doji && upperWick > lowerWick && upperWick > (comparisonRatio*body) && previousCandleType == GREEN_CANDLE;
    
    var result = "Neutral";
    if(indecisiveDoji){
      result = "Indecisive Doji"
    }
    if(dragonflyDoji){
      result = "Dragonfly Doji"
    }
    if(gravestoneDoji){
      result = "Gravestone Doji"
    }
    else if (hammer){
      result = "Hammer"
    }
    else if (invertedHammer){
      result = "Inverted Hammer"
    }
    else if (hangingMan){
      result = "Hanging Man"
    }
    else if (shootingStar){
      result = "Shooting Star"
    }

    finalResult[finalResult.length] = result;
  }
  
  // reverse the list as Today, Yesterday and so on
  return finalResult.reverse();
}