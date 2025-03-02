# Cryptocurrency Live Tracker

This project fetches and analyzes live data for the top 50 cryptocurrencies by market capitalization and updates an Excel file with the latest information every 5 minutes.

## Features

- Fetches real-time data for the top 50 cryptocurrencies using the CoinGecko API
- Performs basic analysis including:
  - Identifying top 5 cryptocurrencies by market cap
  - Calculating average price of top 50 cryptocurrencies
  - Finding highest and lowest 24-hour price changes
- Updates an Excel file with live data every 5 minutes
- Generates a text-based analysis report

## Requirements

- Python 3.7 or higher
- Required Python packages (listed in `requirements.txt`):
  - requests
  - pandas
  - openpyxl

## Installation

1. Clone this repository or download the files
2. Install the required packages:

```
pip install -r requirements.txt
```

## Usage

1. Run the script:

```
python crypto_tracker.py
```

2. The script will:
   - Start fetching data from the CoinGecko API
   - Create and update an Excel file named `crypto_data_live.xlsx`
   - Generate an analysis report named `crypto_analysis_report.txt`
   - Continue updating the Excel file every 5 minutes until manually stopped

3. To stop the script, press `Ctrl+C` in the terminal

## Excel File Structure

The generated Excel file contains two sheets:

1. **Live Crypto Data**: Contains the complete dataset for all 50 cryptocurrencies
2. **Analysis**: Contains summary statistics and analysis results, including:
   - Average price, total market cap, and total trading volume
   - Top 5 cryptocurrencies by market cap
   - Highest and lowest 24-hour price changes

## Notes

- The CoinGecko API has rate limits. If you encounter errors, it might be due to exceeding these limits.
- The refresh interval is set to 5 minutes by default. You can modify the `REFRESH_INTERVAL` constant in the script to change this.

## For the Analysis Report Requirement

The script automatically generates a text-based analysis report on the first run. For a more professional report, you can:

1. Open the `crypto_analysis_report.txt` file
2. Copy the content into a Word document
3. Format it as desired and save as PDF or Word format

## Live Excel Sheet Updates

The Excel file is updated automatically by the Python script. To view the live updates:

1. Open the Excel file in Microsoft Excel
2. Enable automatic refresh or periodically save and reload the file to see the updates 