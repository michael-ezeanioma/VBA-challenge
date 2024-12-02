# Overview

This project leverages VBA (Visual Basic for Applications) scripting to analyze stock market data across multiple quarters. The script automates the tedious task of analyzing large datasets, providing insights into stock performance, including percentage changes, total volume, and identifying key metrics such as the greatest percentage increase, decrease, and total volume. Conditional formatting is applied to enhance data visualization.

# Features

__1. Data Retrieval:__
Reads and processes stock data, including:
Ticker symbols
Opening and closing prices
Total stock volume

__2. Calculated Metrics:__
Quarterly changes in stock prices (absolute and percentage).
Total stock volume for each ticker.
Stocks with:
  Greatest percentage increase.
  Greatest percentage decrease.
  Greatest total volume.

__3. Conditional Formatting:__
Highlights positive changes in green.
Highlights negative changes in red.

__4. Multi-Sheet Analysis:__
Script runs seamlessly across all worksheets in the workbook, automating repetitive tasks.

__5. Results Presentation:__
Outputs results in a structured format with additional calculated columns for:
  Ticker symbol
  Total stock volume
  Quarterly change
  Percentage change
Displays key metrics such as greatest changes and volumes.

# Files

Multiple_year_stock_dataa.xlsm: The Excel file containing multiple sheets of stock market data.

VBA Script File: The VBA script used for automation.

Screenshots: Screenshots of the results for reference.

README.md: This descriptive file.

# Requirements

__1. Data Processing:__
Loop through stock data for each quarter.
Retrieve and calculate relevant metrics.

__2. Column Creation:__
Create new columns for:
Ticker symbol
Quarterly change
Percent change
Total stock volume

__3. Conditional Formatting:__
Apply formatting to:
Highlight positive percentage changes in green.
Highlight negative percentage changes in red.

__4. Calculated Outputs:__
Identify and display:
Stock with the greatest percentage increase.
Stock with the greatest percentage decrease.
Stock with the greatest total volume.

__5. Cross-Sheet Functionality:__
Ensure the script runs effectively across all worksheets in the workbook.

__6. GitHub Submission:__
Include the VBA script, screenshots of results, and a README file.

# Technologies Used

VBA (Visual Basic for Applications): For scripting and automation.

Microsoft Excel: For dataset storage and analysis.

# How to Use
1. Open the Excel file Multiple_year_stock_dataa.xlsm.

2. Access the VBA editor (Alt + F11) and import the VBA script if not already included.
3. Run the script (F5 or through the Excel interface) to process the data.
4. Review the results, including:
  Calculated metrics in new columns.
  Conditional formatting highlighting positive and negative changes.
  Summary of key metrics (greatest percentage increase, decrease, and volume).

<!--Mod 2-->
