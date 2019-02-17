# HISTORICAL INVESTMENT MODELLING

See the effect of dollar cost averaging over some arbitrary time window of a stock's historical performance.

This is a [Google Sheets application](https://docs.google.com/spreadsheets/d/1Ycn_z-d5YjEQqHRzdltxi5kVAHscs_IMMPAvgduiWUQ/edit?usp=sharing).

## ASSUMPTIONS
- Only one stock ticker supported in model
- Generally continguous data is available for all stock tickers over the period being modeled (excluding weekends and holidays)
- Fund contribution window size is the same as the total investment model's window size
- All calculations are performed against the open price of each ticker

## TODO
- Cache stock dates based on contribution window
- Determine how to scale model to permit percentage contributions to multiple ticker symbols
- Use external API for market data: https://www.alphavantage.co/support/#api-key
  - See: https://medium.com/@meetnaren/replicating-a-google-finance-portfolio-on-google-sheets-7a75945a9d9b

## REFERENCES
- Historical data sourced from Yahoo, e.g. https://finance.yahoo.com/quote/IVV/history?period1=958708800&period2=1265778000&interval=1d&filter=history&frequency=1d

