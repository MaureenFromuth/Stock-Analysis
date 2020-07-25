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

Using the data that Steve provided and the variable that he asked us to conduct, we can assess positive and negative returns per green energy company for each year.  This will indicate for that year which stocks performed well and which ones did not.  Using the results depicted in *Figure 1: 2017 All Stock Analysis* the following stocks performed positively for 2017, in descending order of overall return:

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

![Figure 1: 2017 All Stocks Analysis](https://github.com/MaureenFromuth/Stock-Analysis/blob/master/VBA_Challenge_2017.png)

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

As with analysis conducted for stocks in 2017, reducing compute time is the final metric we looked at.  The original code provided analysis within approx .53-.56 seconds whereas the refactored code ran in .10 as identified in *Figure 2*.

![Figure 2: 2018 All Stocks Analysis](https://github.com/MaureenFromuth/Stock-Analysis/blob/master/VBA_Challenge_2018.png)


Based of preliminary analysis, there are two stocks that continued to have positive year or year (YoY) growth: ENPH and RUN.  As you can see from the following graphs

## Summary
