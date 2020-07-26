# Stock Analysis with VBA

## Overview of Project

### Purpose  

Our core customer, Steve, is conducting analysis on green energy companies in support of his clients.  These clients are his his parents and have invested in DAQO Energy Corps (DQ), a company that makes silicon wafers for solar panels.  Steve wants to assess the stock against other stocks in the industry in order to both identify the health and potential of the stock and also to look for ways to diversify.  Steve has collected [data](https://github.com/MaureenFromuth/Stock-Analysis/blob/master/VBA_Challenge.xlsm) on stocks that he would like to analyze.  This data is for the years of 2017 and 2018 and includes the following:

- Company ticker 
- Date of trading
- Opening stock price for that company
- Highest stock price for the day for that company
- Lowest stock price for the day for that company
- Closing stock price for that company
- An adjusted closing stock price for that company
- Total volume of trades for that day of trading

Using this data, Steve has asked us to review an automated VBA pipeline to analyze this data and identify total daily volume and overall return on investment for each company in 2017 and 2018.  Steve is particularly interested in optimizing the time required to provide the analysis as he wants to add more data in the future to support his client.      

## Results

Using the data that Steve provided and the variable that he asked us to conduct, we can assess positive and negative returns per green energy company for each year.  This will indicate for that year which stocks performed well and which ones did not.  For this analysis we utilized the following forumla:
```
tickerEndingPrices(i) / tickerStartingPrices(i) - 1
*where tickerEndingPrices are the closing prices for the particular stock, tickerStartingPrices are the opening prices for the particular stock, and i is the particular stock*
```

**2017 Analysis**
Using the results depicted in *Figure 1: 2017 All Stock Analysis* the following stocks performed positively for 2017, in descending order of overall return:

- DQ
- SEDG
- ENPH
- FSLR
- JKS
- VSLR
- CSIQ
- HASI
- SPWR
- AY
- RUN

Only one stock did not provide a positive return for 2017: TERP.

Additionally, with reducing performance time an objective for Steve's assignment, our analysis also included run time.  The original code ran in approx .538 to .54 seconds and the refactored code ran in .10 seconds as highlighted in *Figure 1*.

>**Figure 1: 2017 All Stock Analysis**
![Figure 1: 2017 All Stocks Analysis](https://github.com/MaureenFromuth/Stock-Analysis/blob/master/VBA_Challenge_2017.png)

Althought not specifically identified in the original analysis, it is also important to highlight the monetary change for each stock in addition to the percentage change.  We included the following formula into VBA in order to create this column.
```
Within the header row (#2) we added - 
Cells(3, 4).Value = "Dollar Change"

Within the return section (#8) we added -
Cells(4 + i, 4).Value = tickerEndingPrices(i) - tickerStartingPrices(i)

Within the formatting secion (#9) we added - 
Range("D4:D15").NumberFormat = "$0.00"
```

Below outlines the monetary changes for each ticker:
- DQ: $39.95
- FSLR: $33.98
- SEDG: $24.35
- JKS: $8.42
- HASI: $4.94
- CSIQ: $4.19
- AY: $1.74
- SPWR: $1.58
- ENPH: $1.36
- VSLR: $1.35
- RUN: $.31
- TERP: -$.93

Comparing the monetary change in the return on investment to the percentage return on investment will help us to measure exactly which stocks will yield the most money.  For example, if a stock has a starting price of $1 and increases to $2 by the end of the year, it will have a 200% increase but will only yield one dollar of a net profit.  This is the case for ENPH, who had an increase of 129.5% but only a increase in $1.36.  Looking at those stocks with positive returns, DQ is the top runner in both percentage as well as net dollar increase, and FSLR and SEDG remain consistent with high percentage as well as dollar returns.

**2018 Analysis**
As depicted in *Figure 2: 2018 All Stocks Analysis*, there were only two stocks that maintained positive returns in 2018, listed below in descending order of overall return: 
- RUN
- ENPH  

The stock for the remaining companies had a negative return, listed below in descending order of overall return:
- VSLR
- TERP
- AY
- SEDG
- CISQ
- HASI
- FSLR
- SPWR
- JKS
- DQ 

As with analysis conducted for stocks in 2017, reducing compute time is the final metric we looked at.  The original code provided analysis within approx .53-.56 seconds whereas the refactored code ran in .12 as identified in *Figure 2*.

>**Figure 2: 2018 All Stock Analysis**
![Figure 2: 2018 All Stocks Analysis](https://github.com/MaureenFromuth/Stock-Analysis/blob/master/VBA_Challenge_2018.png)

Consistent with 2017, we also looked at the change in actual dollar amounts for the return for 2018.  Below lists those returns for each stock in decending order of overall return amount:
- RUN: $4.97
- ENPH: $2.13
- VSLR: -$.14
- TERP: -$.59
- AY: -$1.54
- CSIQ: -$2.80
- SEDG: -$2.95
- SPWR: -$4.00
- HASI: -$4.96
- JKS: -$15.17
- FSLR: -$27.97
- DQ: -$39.17

While RUN and ENPH have positive returns for percentages, their overall monetary increase from starting price to ending price for 2018 is relatively small.  Additionally, for those stocks that did have a negative return, VSLR lost the least in both percentage and money for the overall annual return.  The rest of the stocks with negative annual return remained fairly consistent in ranking between percentage and monetary changes.

**Year Over Year Trends**
In additional to analyzing each year independently, it is also important to look at trends.  A good metric for this is to identify year over year (YoY) trends.  Without computing additional analysis, we looked at which stocks had positive returns for both 2017 and 2018.  There were only two: ENPH and RUN.  If you consider stocks that yielded positive annual returns in 2017 and then minimized their losses in 2018, SEDG is the most optimal stock.  

If we look at overall changes in monetary value for the starting price in 2017 to the closing price in 2018, we will also be able to identify trends in monetary return.  To do so, we took the dollar change for 2017 and added it to the dollar change for 2018. This results in the following return from 2017 to 2018:
- SEDG: $21.40
- FSLR: $6.01
- RUN: $5.28
- ENPH: $3.49
- CSIQ: $1.39
- VSLR: $1.21
- DQ: $.42
- AY: $0.20
- HASI: -$.02
- TERP: -$1.52
- SPWR: -$2.42
- JKS: -$6.75

Using this information, SEDG appears to be the stock with the highest yield over 2017 to 2018.  

## Summary
