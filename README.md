# VBA-Challenge - Stock Market Analysis

This project is a VBA scripting to analyze real stock market data.This analysis used two data-points one is the Test Data which is used while developing the script, and the other is Stock Data which is the main data to run the script. The data is sourced into Excel Microsoft Office, the test data have 7 sheets(A-P) the size of this file is smaller and used for testing. The stock data or the main data has three sheets categorized yearly 2014, 2015, 2016. This file is bigger in size, it may take time when you execute the script into the file. The VBA scripts also found in both data points file directories.

Test Data vbs file & Stock Data vbs file

The script loop through all the stocks data once and the following information displayed.

Solution 1 Ticker Symbol

The script will sort the distinct ticker symbol in one column in column "I" with a column header "Ticker.

Solution 2 Yearly Change

The script will excute Yearly change from opening price at the beginning of a given year to the closing price at the end of that year, and put the value on "J" column. For this task the code added a conditional formatting that highlighted positive change in green and negative change in red.

Solution 3 Precent Change

The script also percent perform a change from opening price at the beginning of a given year to the closing price at the end of that year, and put the value on "K" column.

Solution 4 Total stock Volume

The total stock volume also genereated on "L" column.

Solution 5 Greatest

Also the solution should provide the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
