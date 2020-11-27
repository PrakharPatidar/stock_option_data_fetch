# stock_option_data_fetch

## Fetch and populate data in excel sheet at regular interval

API URL = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

Sample data
```json
{
  "records": {},
  "filtered": {
    "data": [
      {
        "strikePrice": 9900,
        "expiryDate": "03-Dec-2020",
        "PE": {
          "strikePrice": 9900,
          "expiryDate": "03-Dec-2020",
          "underlying": "NIFTY",
          "identifier": "OPTIDXNIFTY03-12-2020PE9900.00",
          "openInterest": 1092,
          "changeinOpenInterest": 683,
          "pchangeinOpenInterest": 166.99266503667482,
          "totalTradedVolume": 3034,
          "impliedVolatility": 76.33,
          "lastPrice": 0.5,
          "change": -0.25,
          "pChange": -33.33333333333333,
          "totalBuyQuantity": 129225,
          "totalSellQuantity": 95625,
          "bidQty": 525,
          "bidprice": 0.45,
          "askQty": 10875,
          "askPrice": 0.5,
          "underlyingValue": 12968.95
        },
        "CE": {
          "strikePrice": 9900,
          "expiryDate": "03-Dec-2020",
          "underlying": "NIFTY",
          "identifier": "OPTIDXNIFTY03-12-2020CE9900.00",
          "openInterest": 2,
          "changeinOpenInterest": 0,
          "pchangeinOpenInterest": 0,
          "totalTradedVolume": 0,
          "impliedVolatility": 0,
          "lastPrice": 3108.4,
          "change": 0,
          "pChange": 0,
          "totalBuyQuantity": 7875,
          "totalSellQuantity": 7950,
          "bidQty": 1500,
          "bidprice": 2955.65,
          "askQty": 1500,
          "askPrice": 3211.1,
          "underlyingValue": 12968.95
        }
      },
      {
        "strikePrice": 9950,
        "expiryDate": "03-Dec-2020",
        "CE": {
          "strikePrice": 9950,
          "expiryDate": "03-Dec-2020",
          "underlying": "NIFTY",
          "identifier": "OPTIDXNIFTY03-12-2020CE9950.00",
          "openInterest": 0,
          "changeinOpenInterest": 0,
          "pchangeinOpenInterest": 0,
          "totalTradedVolume": 0,
          "impliedVolatility": 0,
          "lastPrice": 0,
          "change": 0,
          "pChange": 0,
          "totalBuyQuantity": 7875,
          "totalSellQuantity": 7875,
          "bidQty": 1500,
          "bidprice": 2817.95,
          "askQty": 1500,
          "askPrice": 3169.9,
          "underlyingValue": 12968.95
        },
        "PE": {
          "strikePrice": 9950,
          "expiryDate": "03-Dec-2020",
          "underlying": "NIFTY",
          "identifier": "OPTIDXNIFTY03-12-2020PE9950.00",
          "openInterest": 4,
          "changeinOpenInterest": 4,
          "pchangeinOpenInterest": 0,
          "totalTradedVolume": 2727,
          "impliedVolatility": 0,
          "lastPrice": 0.65,
          "change": -95.55,
          "pChange": -99.32432432432432,
          "totalBuyQuantity": 6150,
          "totalSellQuantity": 19800,
          "bidQty": 2400,
          "bidprice": 0.15,
          "askQty": 150,
          "askPrice": 0.7,
          "underlyingValue": 12968.95
        }
      }
    ]
  }
}

```