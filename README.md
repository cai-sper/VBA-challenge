# VBA Stock Market Analysis

## Overview
This project involves analyzing stock market data using VBA scripting. The goal is to loop through a year's worth of stock data, extract key information, and calculate metrics such as yearly change, percentage change, and total stock volume. Additionally, the script identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Table of Contents
1. [Data Retrieval](#data-retrieval)
2. [Column Creation](#column-creation)
3. [Conditional Formatting](#conditional-formatting)
4. [Calculated Values](#calculated-values)
5. [Looping Across Worksheets](#looping-across-worksheets)

## Data Retrieval
- The VBA script loops through one year of stock data to retrieve the following values for each row:
  - Ticker symbol
  - Volume of stock
  - Opening price
  - Closing price

## Column Creation
- On the same worksheet as the raw data, the script creates columns for the following:
  - Ticker symbol
  - Total stock volume
  - Yearly change ($)
  - Percentage change (%)

## Conditional Formatting
- Conditional formatting is applied to the yearly change and percentage change columns:
  - Positive change is highlighted in green
  - Negative change is highlighted in red

## Calculated Values
- The script calculates and displays the following values:
  - Greatest Percentage Increase
  - Greatest Percentage Decrease
  - Greatest Total Volume

## Looping Across Worksheets
- The VBA script is designed to run on all worksheets, enabling analysis of multiple years of stock data.
