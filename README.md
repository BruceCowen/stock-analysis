# stock-analysis
Module 2 Stock Analysis
# VBA Challenge Data Analysis of Stocks

## Overview of Project
This project used VBA and Excel to analyze 12 different stocks and show the results of the volume traded and the yearly return for each stock. To do this we wrote a VBA script to analyze these 12 different stocks at the same time. This VBA script incorporated most of the tools we learned throughout this modular, which includes for loops, if then statements, formatting, etc. 

### Purpose
The purpose of this project was to put the skills we learned in the modular into practice with a real-world problem. Every lesson was incorporated in some way to this final challenge and allowed us to put our skills to the test.

## Analysis and Challenges
For this challenge we analyzed data given to us that showed the highs and lows as well as the daily volume traded for 12 different stocks with roughly the same underlying theme of green energy. Based on the VBA script written we can see that most of these green stocks did better in 2017 than 2018.

### Analysis of 2017 Performance
The results of the VBA script written showed results of daily volume as well as the yearly return of each stock. Some performed better than others but all of them except one had a positive return. This one was TERP which had a return of -7.2%. without diving into more data about this specific stock its hard to say why this one performed poorly. Four of these stocks showed a return of over 100% leading to the belief that 2017 was a great year for green energy.

### Analysis of 2018 Performance
Out of the 12 stocks analyzed, only 2 had a positive return for the year of 2018. 10 stocks had a negative return although some were very low some were also quite high. Such as our first stock we analyzed which was DQ. This analysis shows us that 2018 was not a good year for green stocks or green companies. 

### Challenges and Difficulties Encountered
There were many challenges to this project but the worst was syntax errors and minor errors such as a minus sign where a place sign should’ve been which led to the starting price and ending price being the same value. This resulted in a return of 0. Below is an example of the code used to calculating the starting price.

If Cells(i-1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
	tickerStartingPrice(tickerIndex) = Cells(i, 6).Value

This code shows how the starting price was calculated. The ending price followed a similar format but replaced the minus sign with a plus sign.


## Results
As discussed earlier the results of our analysis for each year is straightforward. Based on the data our VBA script turned out it is safe to say that the green stocks we were working with performed much better in 2017 than they did in 2018. With only a couple of them being an exception.

## Summary 
For this project we refactored VBA code adding more features. Advantages of refactoring code is to get another set of eyes on the code. This can lead to doing things more efficiently and take parts out that would otherwise take up too much space and time. The original was well purposed in terms of formatting but lacked the core of the VBA script’s purpose which was to analyze the stock data. With the refactored VBA script we could analyze any stocks we choose.
