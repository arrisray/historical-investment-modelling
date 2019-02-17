/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 *
 * @OnlyCurrentDoc
 */
function onOpen() 
{
  // var spreadsheet = SpreadsheetApp.getActive();
  SpreadsheetApp.getUi()
    .createMenu('Investments')
    .addItem('Create inputs sheet...', '_createInputsSheet')
    .addItem('Calculate performance...', '_calculatePerformance')
    .addToUi();
}

/**
 * A function that creates a sheet to specify investment performance model inputs
 */
function _createInputsSheet() 
{
  var investmentModelSheetName = 'Investment Model';
  var headers = [
    'Stock Ticker',
    'Start Date',
    'End Date',
    'Contribution Frequency',
    'Initial Fund Amount',
    'Salary Contribution Amount'
  ];
  
  // Delete the inputs sheet if it already exists
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(investmentModelSheetName);
  if (sheet != null) {
    SpreadsheetApp.setActiveSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
  }
  
  // Create and configure our inputs sheet
  sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(investmentModelSheetName);
  sheet.setName(investmentModelSheetName);
  sheet.getRange('A1:F1').setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function _calculatePerformance() 
{
  createQuerySheet();
  
  var inputs = getInputs();
  var contributionInfos = generateContributions(
    inputs.ticker.value, 
    inputs.startDate.value,
    inputs.endDate.value,
    inputs.contributionFrequency.value,
    inputs.initialFundAmount.value,
    inputs.salaryContributionAmount.value
  );
  var performanceInfo = calculatePerformanceInfo(inputs, contributionInfos);
  writeResultsToSheet(performanceInfo, inputs);
  
  deleteQuerySheet();
}

//
// Get expected inputs required to run our investment model
//
function getInputs()
{
  var ui = SpreadsheetApp.getUi();
  var investmentModelSheetName = 'Investment Model';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(investmentModelSheetName);
  
  var inputInfo = Object.create(InputInfo);
  for (var key in inputInfo) {
    var property = inputInfo[key];
    var range = sheet.getRange(property.cell);
    if (range.isBlank()) {
      ui.alert('Missing Input', 'Please specify a value for ' + input.label + ' (' + input.cell + ').', ui.ButtonSet.OK);
      return;
    }
    property.value = range.getValue();
  }
  return inputInfo;
}

function generateContributions(ticker, startDate, endDate, contributionFrequency, initialFundAmount, salaryContributionAmount)
{
  var contributionInfos = [];
  var fundContributionAmount = calculateContributionPerPeriod(startDate, endDate, contributionFrequency, initialFundAmount);
  
  var currentDate = startDate;
  while (currentDate <= endDate) {
    // Determine contribution info
    var contributionInfo = calculateContributionInfo(currentDate, ticker, fundContributionAmount, salaryContributionAmount);
    if (contributionInfo == null) {
      break;
    }
    contributionInfos.push(contributionInfo);
    
    // NOTE: This function performs a fuzzy search for the next available date;
    // we need to update the current date if our initial request date had no stock info available
    currentDate = contributionInfo.date;
    
    // Advance date
    currentDate = getNextContributionDate(currentDate, contributionFrequency);
  }
  
  return contributionInfos;
}

/**
 * Calculate contribution amounts for the given ticker on the given date
 *
 * @param {string} currentDate
 * @param {string} ticker
 * @param {number} fundContributionAmount
 * @param {number} salaryContributionAmount
 * @returns {(ContributionInfo|null)}
 */
function calculateContributionInfo(currentDate, ticker, fundContributionAmount, salaryContributionAmount)
{
  // Get stock info
  currentDate = new Date(currentDate);
  var stockInfo = getStockInfoForDate(ticker, currentDate);
  if (stockInfo == null) {
    return null;
  }
  
  stockInfo = stockInfo.getValues();
  var date = stockInfo[0][0];
  var purchasePrice = stockInfo[0][1];
  
  // Calculate contribution info
  var contributionInfo = new ContributionInfo();
  contributionInfo.date = date;
  contributionInfo.ticker = ticker;
  contributionInfo.purchasePrice = purchasePrice;
  contributionInfo.numSalaryShares = salaryContributionAmount / purchasePrice;
  contributionInfo.numFundShares = fundContributionAmount / purchasePrice;
  return contributionInfo;
}

/**
 * @param {ContributionInfo[]} contributionInfos
 * @param {string} ticker
 * @param {Date} endDate
 * @returns {(PerformanceInfo|null)}
 */
function calculatePerformanceInfo(inputs, contributionInfos)
{
  var ticker = inputs.ticker.value;
  var startDate = inputs.startDate.value;
  var endDate = inputs.endDate.value;
  
  // Calculate initial purchase and sell price
  var stockInfo = getStockInfoForDate(ticker, startDate);
  var purchasePrice = stockInfo.getValues()[0][1];
  
  var stockInfo = getStockInfoForDate(ticker, endDate);
  if (stockInfo == null) {
    return null;
  }
  
  stockInfo = stockInfo.getValues();
  var sellPrice = stockInfo[0][1];
  
  // Accumulate totals
  var performanceInfo = new PerformanceInfo();
  var gainPerPeriod = [];
  var totalGain = 0;
  var totalFundShares = 0;
  var totalSalaryShares = 0;
  for (var key in contributionInfos) {
    var contributionInfo = contributionInfos[key];
    var fundGain = (sellPrice * contributionInfo.numFundShares) - (contributionInfo.purchasePrice * contributionInfo.numFundShares);
    var salaryGain = (sellPrice * contributionInfo.numSalaryShares) - (contributionInfo.purchasePrice * contributionInfo.numSalaryShares);
    
    var d = contributionInfo.date.toString();
    gainPerPeriod.push({
      'date': d,
      'numFundShares': contributionInfo.numFundShares,
      'fundGain': fundGain,
      'numSalaryShares': contributionInfo.numSalaryShares,
      'salaryGain': salaryGain
    });
    totalFundShares += contributionInfo.numFundShares;
    totalSalaryShares += contributionInfo.numSalaryShares;
    totalGain += (fundGain + salaryGain);
  }
  
  // Write final performance metrics
  performanceInfo.gainPerPeriod = gainPerPeriod.slice(0);
  performanceInfo.stockPurchasePrice = purchasePrice;
  performanceInfo.stockSellPrice = sellPrice;
  performanceInfo.totalFundShares = totalFundShares;
  performanceInfo.totalSalaryShares = totalSalaryShares;
  performanceInfo.totalFund = inputs.initialFundAmount.value;
  performanceInfo.totalSalary = inputs.salaryContributionAmount.value * performanceInfo.gainPerPeriod.length;
  performanceInfo.totalGain = totalGain;
  performanceInfo.naiveGain = calculateNaiveGain(performanceInfo);
  return performanceInfo;
}

/**
 * @param {PerformanceInfo} performanceInfo
 */
function writeResultsToSheet(performanceInfo)
{
  var sheetName = 'Results';
  var headers = [
    'Date',
    'Num Fund Shares',
    'Fund Gain',
    'Num Salary Shares',
    'Salary Gain',
    'Total Fund Shares',
    'Total Salary Shares',
    'Stock Initial Purchase Price',
    'Stock Sell Price',
    'Total Fund',
    'Total Salary',
    'Total Gain',
    'Total Naive Gain'
  ];
  
  // Delete the results sheet if it already exists
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet != null) {
    SpreadsheetApp.setActiveSheet(sheet);
    SpreadsheetApp.getActiveSpreadsheet().deleteActiveSheet();
  }
  
  // Create and configure our results sheet
  sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  sheet.setName(sheetName);
  sheet.getRange('A1:M1').setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // Populate the periodic metrics
  for (var key in performanceInfo.gainPerPeriod) {
    var row = parseInt(key, 10) + 2; // header row + array off-by-one count
    var gainInfo = performanceInfo.gainPerPeriod[key];
    sheet.getRange('A' + row).setValue(new Date(gainInfo.date));
    sheet.getRange('B' + row).setValue(gainInfo.numFundShares);
    sheet.getRange('C' + row).setValue(gainInfo.fundGain);
    sheet.getRange('D' + row).setValue(gainInfo.numSalaryShares);
    sheet.getRange('E' + row).setValue(gainInfo.salaryGain);
  }
  
  // Write the scalar metrics
  sheet.getRange('F2').setValue(performanceInfo.totalFundShares);
  sheet.getRange('G2').setValue(performanceInfo.totalSalaryShares);
  sheet.getRange('H2').setValue(performanceInfo.stockPurchasePrice);
  sheet.getRange('I2').setValue(performanceInfo.stockSellPrice);
  sheet.getRange('J2').setValue(performanceInfo.totalFund);
  sheet.getRange('K2').setValue(performanceInfo.totalSalary);
  sheet.getRange('L2').setValue(performanceInfo.totalGain);
  sheet.getRange('M2').setValue(performanceInfo.naiveGain);
  
  // Tweak!
  sheet.autoResizeColumns(1, headers.length);
}

