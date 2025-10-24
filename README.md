# Straddle-Master-Stocks

![GitHub stars](https://img.shields.io/github/stars/Connor9994/Straddle-Master-Stocks?style=social) ![GitHub forks](https://img.shields.io/github/forks/Connor9994/Straddle-Master-Stocks?style=social) ![GitHub issues](https://img.shields.io/github/issues/Connor9994/Straddle-Master-Stocks) 

![1](https://github.com/Connor9994/Straddle-Master-Stocks/blob/main/Photos/1.png)

## Overview

This Python script analyzes SPY (SPDR S&P 500 ETF Trust) options data to calculate and export straddle and strangle strategies for various expiration dates. The tool fetches real-time options data from Yahoo Finance and generates a comprehensive Excel workbook with detailed analysis.

## Features

- **Complete Options Chain Data**: Retrieves 100% of available SPY call and put options
- **Current Price Analysis**: Automatically detects the current SPY price and rounds to nearest strike price
- **Strategy Analysis**: Calculates straddle and strangle strategies with break-even points
- **Multi-Sheet Excel Export**: Creates organized Excel workbook with separate sheets for each expiration date
- **Risk Analysis**: Computes break-even ranges and profit/loss spans for various option combinations

## How It Works

### 1. Data Collection
- Fetches all available SPY options expiration dates
- Retrieves both call and put options for each expiration
- Calculates midpoint prices (mark) from bid-ask spreads

### 2. Core Analysis
- **Straddle Analysis**: Analyzes options with the same strike price (±10 points from current price)
- **Strangle Analysis**: Analyzes out-of-the-money call and put combinations
- **Break-even Calculation**: Computes upper and lower break-even points for each strategy

### 3. Output Structure
The generated Excel file contains:
- **SPY_Chain Sheet**: Complete raw options data
- **Date-Specific Sheets**: For each expiration date, includes:
  - Straddle analysis for strikes around current price
  - Strangle analysis for various strike combinations
  - Break-even ranges and span calculations

## Key Metrics Calculated

- **Break-even Up**: Upper break-even price (strike + total premium)
- **Break-even Down**: Lower break-even price (strike - total premium)
- **Span**: Difference between break-even points (profit range width)

## Usage

Simply run the script to:
1. Fetch current SPY options data
2. Analyze straddle and strangle strategies
3. Generate `SPYChain.xlsx` with comprehensive analysis

## Dependencies

- `pandas` - Data manipulation and Excel export
- `numpy` - Numerical operations and filtering
- `yfinance` - Yahoo Finance API for options data
- `openpyxl` - Excel file handling
- `datetime` - Date processing utilities

![2](https://github.com/Connor9994/Straddle-Master-Stocks/blob/main/Photos/2.png)

## Output

The script creates an Excel file with organized sheets showing options strategies, making it easy to identify potential trading opportunities based on break-even analysis and price movement expectations.

## Note

This tool is designed for educational and analytical purposes. Options trading involves significant risk and may not be suitable for all investors. Always conduct thorough research and consider consulting with a financial professional before engaging in options trading.
