// total loss/gain per contribution period, aggregate total loss/gain
var PerformanceInfo = function()
{
  this.gainPerPeriod; // {<date>:{'fund':<gain-amount>, 'salary':<gain-amount>}}
  this.stockPurchasePrice;
  this.stockSellPrice;
  this.totalFundShares;
  this.totalSalaryShares;
  this.totalFund;
  this.totalSalary;
  this.totalGain;
  this.naiveGain;
};

