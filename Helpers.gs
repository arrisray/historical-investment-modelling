/**
 * Calculates the amount to contribute per period of some total contribution window
 * 
 * @param {string} startDate The beginning of the contribution window
 * @param {string} endDate The end date of the contribution window
 * @param {number} contributionFrequency The contribution frequency in days
 * @param {number} totalAmount The total amount of money to invest over the contribution window
 * @returns {number}
 */
function calculateContributionPerPeriod(startDate, endDate, contributionFrequency, totalAmount)
{
  startDate = new Date(startDate);
  endDate = new Date(endDate);
  
  var msInDay = 1000 * 60 * 60 * 24;
  contributionFrequency *= msInDay;
  var diff = endDate.getTime() - startDate.getTime();
  var numPeriods = diff / contributionFrequency;
  var contributionPerPeriod = totalAmount / numPeriods;
  
//  Logger.log('Contribution window: ' + startDate.toDateString() + ' - ' + endDate.toDateString());
//  Logger.log('Contribution frequency: ' + contributionFrequency / msInDay);
//  Logger.log('Num periods: ' + numPeriods);
//  Logger.log('Contribution per period: ' + contributionPerPeriod);
  return contributionPerPeriod;
}

/**
 * Given a date, return the next contribution date
 *
 * @param {Date} currentDate
 * @param {number} contributionPeriod A contribution period, in days
 * @returns {Date}
 */
function getNextContributionDate(currentDate, contributionPeriod)
{
  var dayInMs = 1000 * 60 * 60 * 24;
  var contributionPeriodInMs = contributionPeriod * dayInMs;
  var currentTime = currentDate.getTime();
  var d = new Date(currentTime + contributionPeriodInMs);
  return d;
//  currentDate.setMonth(currentDate.getMonth() + 1);
//  currentDate.setDate(1);
//  return currentDate;
}

/**
 * Get the basic metrics of a stock on the given date
 * 
 * @param {string} ticker The stock ticker symbol
 * @param {string} date The date we'd like to examine
 * @returns {(Range|null)}
 *
 * @todo Maybe return null if we don't have fuzzy date available within a week?
 */
function getStockInfoForDate(ticker, date) 
{
  date = new Date(date);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);
  var sqlDate = year + '-' + month + '-' + day;
  
  data = ticker + "!A:G";
  sql = "SELECT * WHERE A >= date '" + sqlDate + "' LIMIT 1";
  
  var result = query(data, sql);
  return result;
}

/**
 * @param {PerformanceInfo} performanceInfo
 */
function calculateNaiveGain(performanceInfo)
{
  var totalShares = performanceInfo.totalFundShares + performanceInfo.totalSalaryShares;
  var startValue = performanceInfo.stockPurchasePrice * totalShares;
  var endValue = performanceInfo.stockSellPrice * totalShares;
  return endValue - startValue;
}

/**
 * @returns {(Date|string)}
 * @see https://weblog.west-wind.com/posts/2014/Jan/06/JavaScript-JSON-Date-Parsing-and-real-Dates
 */
function jsonDateParser(key, value) 
{
  var reISO = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*))(?:Z|(\+|-)([\d|:]*))?$/;
  if (typeof value === 'string') {
    var a = reISO.exec(value);
    if (a) {
      return new Date(value);
    }
  }
  return value;
}

/**
 * @note Omits headers in results by default
 * @param {string} data
 * @param {string} sql
 * @see https://stackoverflow.com/a/28751946
 * @returns {Range}
 */
function query(data, sql)
{
  sql = "=QUERY(" + data + ", \"" + sql + "\", 0)";
  var querySheetName = '__QUERY__';
  var querySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(querySheetName);
  var r = querySheet.getRange(1, 1).setFormula(sql);
  var reply = querySheet.getDataRange();
  return reply;
}

function createQuerySheet()
{
  deleteQuerySheet();
  
  var name = '__QUERY__';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  sheet.setName(name);
}

function deleteQuerySheet()
{
  var name = '__QUERY__';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (sheet) {
    SpreadsheetApp.setActiveSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
  }
}

