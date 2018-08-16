/**
 * The MIT License (MIT)
 * 
 * Copyright (c) 2018 Marcin P (https://github.com/yu55)
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

var LAST_REFRESH_DATE_CELL = "B1"; // komórka w której przechowujemy datę ostatniej aktualizacji

function onOpen() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{
    name : "Aktualizuj kursy walorów",
    functionName : "updateLastRefreshDate"
  }];
  activeSpreadsheet.addMenu("STOOQ", menuEntries);
  updateLastRefreshDate();
};

function updateLastRefreshDate() {
  SpreadsheetApp.getActiveSpreadsheet().getRange(LAST_REFRESH_DATE_CELL).setValue(new Date());
}

/**
 * Download share price for a gicen ticker from stooq.com web page.
 * Date parameter is optional. If date is not provided or provided date is today
 * then function will download current share price for a given ticker from stooq.com web page.
 * Otherwise function will download historic share price for a given ticker and date.
 *
 * @param {string} ticker share ticker name; e.g. "^SPX".
 * @param {Date=} date given date.
 * @return share price at given date if date is provided or current share price.
 * @customfunction
 */
function STOOQ_GET_PRICE(ticker, date) {
  if (isRequestForCurrentPrice(date)) {
    return stooqGetCurrentPrice(ticker);
  } else {
    return stooqGetHistoricClosingPrice(ticker, date);
  }
}

function isRequestForCurrentPrice(date) {
  var now = initNowDateWithoutHours();
  return (date == null || setZeroHours(date).getTime() == now.getTime());
}

function initNowDateWithoutHours() {
  var now = new Date();
  return setZeroHours(now);
}

function setZeroHours(date) {
  date.setHours(0,0,0,0);
  return date;
}



function stooqGetCurrentPrice(ticker) {
  validateTicker(ticker);
  var stooqWebPageSource = downloadStooqWebPageSource(ticker)
  var priceValue = extractSharePrice(stooqWebPageSource, ticker);
  return priceValue;
}

function validateTicker(ticker) {
  if (!ticker || 0 === ticker.trim().length) {
    throw ("ERROR: function argument \"ticker\" cannot be empty");
  }
}

function downloadStooqWebPageSource(ticker) {
  var url = prepareStooqWebPageUrl(ticker)
  var stooqResponse = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  if (stooqResponse.getResponseCode() != 200) {
    throw ("ERROR: Cannot download \"" + ticker + "\" data. HTTP code: " + stooqResponse.getResponseCode());
  }
  return stooqResponse.getContentText();
}

function prepareStooqWebPageUrl(ticker) {
  var normalizedTicker = normalizeTicker(ticker);
  return 'https://stooq.com/q/?s=' + encodeURIComponent(normalizedTicker);
}

function normalizeTicker(ticker) {
  return ticker.toLowerCase().trim();
}

function extractSharePrice(stooqWebPageSource, ticker) {
  var priceText = findPriceText(stooqWebPageSource, ticker);
  var priceAsNumber = convertTextToNumber(priceText);
  return priceAsNumber;
}

function findPriceText(stooqWebPageSource, ticker) {
  // we're looking for a HTML fragment like this: "<span style="font-weight:bold" id="aq_^spx_c2">2833.28</span>"
  var priceHtmlSpanTagIdLocation = findPriceHtmlSpanTagIdLocation(stooqWebPageSource, ticker);
  var priceTextStartLocation = stooqWebPageSource.indexOf(">", priceHtmlSpanTagIdLocation);
  var priceTextEndLocation = stooqWebPageSource.indexOf("<", priceHtmlSpanTagIdLocation);
  return stooqWebPageSource.substring(priceTextStartLocation + 1, priceTextEndLocation);
}

function findPriceHtmlSpanTagIdLocation(stooqWebPageSource, ticker) {
  var priceHtmlSpanTagId = preparePriceHtmlSpanTagId(ticker);
  var priceHtmlSpanTagIdLocation = stooqWebPageSource.indexOf(priceHtmlSpanTagId);
  if (priceHtmlSpanTagIdLocation == -1) {
    throw ("ERROR: cannot find price of \"" + ticker + "\" ticker on STOOQ web page");
  }
  return priceHtmlSpanTagIdLocation;
}

function preparePriceHtmlSpanTagId(ticker) {
  var normalizedTicker = normalizeTicker(ticker);
  return "aq_" + normalizedTicker + "_c";
}

function convertTextToNumber(text) {
  return Number(text);
}



function stooqGetHistoricClosingPrice(ticker, date) {
  validateTicker(ticker);
  // TODO validateDate(date);
  var historyCsv = downloadStooqHistoryCsv(ticker, date);
  var historicClosingPrice = extractHistoricSharePrice(historyCsv);
  return convertTextToNumber(historicClosingPrice);
}

function downloadStooqHistoryCsv(ticker, date) {
  var url = prepareStooqHistoryCsvUrl(ticker, date)
  var stooqResponse = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  if (stooqResponse.getResponseCode() != 200) {
    throw ("ERROR: Cannot download \"" + ticker + "\" history CSV data. HTTP code: " + stooqResponse.getResponseCode());
  }
  return stooqResponse.getContentText();
}

function prepareStooqHistoryCsvUrl(ticker, date) {
  var normalizedTicker = normalizeTicker(ticker);
  var dateParam = Utilities.formatDate(date, "GMT", "yyyyMMdd");
  var url = "https://stooq.com/q/d/l/?s="+ encodeURIComponent(ticker) + "&d1=" + encodeURIComponent(dateParam) + "&d2=" + encodeURIComponent(dateParam) + "&i=d";
  return url;
}

function extractHistoricSharePrice(historyCsv) {
  var historyCsvWthoutCR = historyCsv.replace(/[\r]/g, "");
  var csvLines = historyCsvWthoutCR.split("\n");
  var data = csvLines.map(function(line) {return line.split(",")});
  if (data.length < 2 || data[1].length < 5) {
    throw ("ERROR: Cannot find price data in Stooq history CSV file. historyCsv=\"" + historyCsv + "\" data=" + data);
  }
  const CSV_DATA_ROW_INDEX = 1;
  const CLOSING_PRICE_DATA_INDEX = 4;
  return data[CSV_DATA_ROW_INDEX][CLOSING_PRICE_DATA_INDEX];
}



////////////////////////////////////////////////////////////////////////////////////////////////////
// Tests

function shouldReturnCurrentPriceForValidExistingTicker() {
  Logger.log(stooqGetCurrentPrice("^spx"));
}

function shouldReturnCurrentPriceForExistingTickerWithCapitalLetters() {
  Logger.log(stooqGetCurrentPrice("^SPX"));
}

function shouldReturnCurrentPriceForUntrimmedExistingTicker() {
  Logger.log(stooqGetCurrentPrice(" ^SPX "));
}

function shouldThrowExceptionForCurrentNonExistingTicker() {
  Logger.log(stooqGetCurrentPrice("castar"));
}

function shouldThrowExceptionForCurrentNullTicker() {
  Logger.log(stooqGetCurrentPrice(null));
}

function shouldThrowExceptionForCurrentEmptyTicker() {
  Logger.log(stooqGetCurrentPrice(""));
}

function shouldThrowExceptionForCurrentBlankTicker() {
  Logger.log(stooqGetCurrentPrice(" "));
}



function shouldReturnHistoricPriceForValidExistingTicker() {
  Logger.log(stooqGetHistoricClosingPrice("^spx", new Date(2017, 7, 11)));
}

function shouldReturnHistoricPriceForExistingTickerWithCapitalLetters() {
  Logger.log(stooqGetHistoricClosingPrice("^SPX", new Date(2017, 7, 11)));
}

function shouldReturnHistoricPriceForUntrimmedExistingTicker() {
  Logger.log(stooqGetHistoricClosingPrice(" ^SPX ", new Date(2017, 7, 11)));
}

function shouldThrowExceptionForHistoricNonExistingTicker() {
  Logger.log(stooqGetHistoricClosingPrice("castar", new Date(2017, 7, 11)));
}

function shouldThrowExceptionForHistoricNullTicker() {
  Logger.log(stooqGetHistoricClosingPrice(null, new Date(2017, 7, 11)));
}

function shouldThrowExceptionForHistoricEmptyTicker() {
  Logger.log(stooqGetHistoricClosingPrice("", new Date(2017, 7, 11)));
}

function shouldThrowExceptionForHistoricBlankTicker() {
  Logger.log(stooqGetHistoricClosingPrice(" ", new Date(2017, 7, 11)));
}

function shouldThrowExceptionForHistoricNullDate() {
  Logger.log(stooqGetHistoricClosingPrice("^spx", null));
}



function shouldInvokeGetCurrentPriceFunctionWhenNoDatePrivided() {
  Logger.log(STOOQ_GET_PRICE("^spx"));
  Logger.log(STOOQ_GET_PRICE("^spx", new Date()));
}

function shouldInvokeGetCurrentPriceFunctionWhenTodayDatePrivided() {
  Logger.log(STOOQ_GET_PRICE("^spx", new Date()));
}

function shouldInvokeGetHistoricPriceFunction() {
  Logger.log(STOOQ_GET_PRICE("^spx", new Date(2017, 7, 11)));
}


