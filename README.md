# VBA_challenge
Module 2

# Project Overview: Stock Market Data Analysis Using VBA

## Background
This project involves using VBA scripting to analyse stock market data, specifically focusing on quarterly performance metrics. The goal is to develop a script that processes stock data across multiple worksheets, each representing a different quarter, and provides key insights into the performance of each stock.

## Project Objectives
The primary objective of this project is to create a VBA script that loops through all the stocks for each quarter and outputs the following key information:
1. **Ticker Symbol**: The unique identifier for each stock.
2. **Quarterly Change**: The change in stock price from the opening price at the beginning of the quarter to the closing price at the end of the quarter.
3. **Percentage Change**: The percentage change in stock price over the quarter, calculated as the quarterly change divided by the opening price.
4. **Total Stock Volume**: The cumulative volume of shares traded for each stock during the quarter.

## Key Features and Requirements
- **Data Processing Across Multiple Worksheets**: Looping through all worksheets in one workbook, each representing a different quarter, to perform the required analysis.
- **Conditional Formatting**: 
  - Highlight positive price changes in green.
  - Highlight negative price changes in red.
- **Summary Information**: Output the following metrics for each stock:
  - Ticker symbol.
  - Quarterly change in stock price.
  - Percentage change in stock price.
  - Total volume of stock traded.

## Additional Features
To enhance the functionality of the script, the following additional features were implemented:
1. **Greatest % Increase**: Identify and display the stock with the greatest percentage increase over the quarter.
2. **Greatest % Decrease**: Identify and display the stock with the greatest percentage decrease over the quarter.
3. **Greatest Total Volume**: Identify and display the stock with the highest total volume of shares traded during the quarter.
4. **Multi-Sheet Processing**: Modify the script to process data across all worksheets (quarters) in the workbook, in one execution.

## Approach
1. **Looping Through Data**: The script utilises a loop to iterate through each worksheet and another nested loop to process each stock within the worksheet.
2. **Data Calculation**: For each stock:
   - Calculate the quarterly change by subtracting the opening price from the closing price.
   - Calculate the percentage change using the quarterly change and the opening price.
   - Sum the total stock volume.
3. **Conditional Formatting**: Apply conditional formatting to highlight positive and negative changes, aiding in visual analysis.
4. **Logic Implementation**: Additional logic is incorporated to track and display the greatest percentage increase, greatest percentage decrease, and greatest total volume across all stocks.

## Outcome
A robust VBA script capable of analysing and summarising stock performance data across multiple quarters. This script not only provides essential financial metrics but also enhances data readability through visual cues like conditional formatting. Other features add significant value by highlighting exceptional stock performance, enabling quicker insights into market trends.