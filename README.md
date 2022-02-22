# Overview of Project

The project is based on the analysis of price and volume data of a number of stocks in the green energy sector in order to determine which stocks offer a positive return and which do not, so that an investment decision can be made.

As a starting point, an Excel file was received with information on a group of 12 pre-selected stocks.  The Excel file contains 8 columns and 3013 rows.  It was not necessary to work with all the columns, but only with column 1 containing the stock name code, also known as ticker, column 6 containing the daily closing price of the stock and column 8 containing the daily amount of volume traded on the stock exchange for the different stocks.

### Purpose of the analysis

This project has two parts:

#### Part 1
The purpose of the first part of this project is to automate the analysis of stocks by creating a code that reads the values with the stock price at the beginning of the year and at the end of the year and outputs the stock's performance as a percentage.  The code should also report the total volume traded for each stock.  Once the indicated information has been generated, the code should create a table with the results, format the columns of the table and highlight with green color those shares that obtained a profit and with red color the shares that had a negative performance.

The original code works on the basis of two nested loops.  One loop goes through the tickers of the 12 stocks and for each ticker performs another loop that goes through the entire Excel sheet and collects the information about the initial price, final price and volume of each of the tickers.

#### Part 2
In the second part, we seek to improve the execution time of the code and we propose to change (Refactor) the original code.  The change consists of avoiding the use of nested loops and making a loop that runs through the entire Excel sheet once and keeps track of the accumulated total for the shares.  In addition, each time there is a change in the ticker name, it records the initial and final price of the ticker.

The idea is to compare the execution time of the original method described in part 1 against the execution time of the refactored code described in part 2 and see if there is a performance difference when refactoring.

#### Part 3
This part was developed on my own initiative in order to see if the refactored code could be further improved.

In the third method, the data of the Excel sheet is traversed only once and the row numbers in which the change from one ticker to another occurs are stored in an array and the running total volume data for each ticker is extracted.  These rows are called break points.  Once the break points are determined, a call is made to the Excel cells containing the initial and final price data for each stock in the break point lines and the table is constructed with the results, giving it the same green and red color format explained in the previous section.

For better understanding, for example in the year 2017 the change from Ticker AY to Ticker CSIQ occurs in row 253; then row 253 constitutes a break point.  Other break points occur in row 504 (ticker changes from CSIQ to DQ), row 755 (ticker changes from DQ to ENPH) and so on.





## Results

## Conclusions



